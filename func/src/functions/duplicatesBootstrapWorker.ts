import { app, type InvocationContext } from "@azure/functions";
import {
  BOOTSTRAP_QUEUE_NAME,
  enqueueBootstrapMessage,
  type BootstrapQueueMessage,
} from "../services/duplicatesBootstrapQueue.js";
import {
  getBootstrapJob,
  readBootstrapManifest,
  updateBootstrapJob,
} from "../services/duplicatesBootstrapJobStore.js";
import { getSpoToken, SPO_ORIGIN } from "../services/spoTokenProvider.js";
import {
  overlayVersionHistoryToStore,
  type FileObservation,
  type HistoricalHash,
} from "../services/hashIndexStore.js";
import { hashFileContent, type HashAlgo } from "../services/contentHash.js";
import { extractText, textContentHash, simhash64 } from "../services/textHash.js";
import type { DocumentFileRef } from "../services/spoFilesEnumerator.js";

/**
 * Bootstrap worker — processes one file per queue message.
 *
 * For each file:
 *   1. List its SP version history
 *   2. Download + SHA256 each historical version (capped by the job's
 *      maxVersionsPerFile, oldest-first)
 *   3. Download + hash the current version
 *   4. Overlay into the persistent hash index
 *   5. Update counters on the job row, enqueue next fileIndex
 *
 * Mirrors the finalize pattern of `linkInventoryScanDocsWorker` — if
 * the fileIndex exhausts the manifest, the job is marked completed.
 */

interface SpVersion {
  ID: number;
  Created: string;
  VersionLabel?: string;
  IsCurrentVersion?: boolean;
}

async function fetchVersions(sitePath: string, fileRef: string): Promise<SpVersion[]> {
  const token = await getSpoToken();
  const encoded = fileRef.split("/").map(encodeURIComponent).join("/").replace(/'/g, "''");
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/getfilebyserverrelativeurl('${encoded}')/Versions` +
    `?$select=ID,Created,VersionLabel,IsCurrentVersion`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Versions fetch failed: ${res.status} ${body.slice(0, 200)}`);
  }
  const data = (await res.json()) as { value?: SpVersion[] };
  return data.value ?? [];
}

/**
 * Download a specific historical version via the canonical SP REST
 * endpoint: `/_api/web/GetFileByServerRelativeUrl('<path>')/Versions/
 * getById(<id>)/$value`. Called in the file's own site context — the
 * raw `_vti_history/...` URLs SP returns from the Versions collection
 * aren't directly authorizable with a bearer token on modern tenants.
 */
async function downloadVersionAndHash(
  sitePath: string,
  fileRef: string,
  versionId: number,
  fileType: DocumentFileRef["fileType"],
): Promise<{ hash: string; algo: HashAlgo; textHash?: string; simhash64?: string }> {
  const token = await getSpoToken();
  const encoded = fileRef.split("/").map(encodeURIComponent).join("/").replace(/'/g, "''");
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/getfilebyserverrelativeurl('${encoded}')` +
    `/Versions/getById(${versionId})/$value`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Download failed: ${res.status} ${body.slice(0, 200)}`);
  }
  const ab = await res.arrayBuffer();
  const buf = Buffer.from(ab);
  const content = await hashFileContent(buf, fileType);
  // Text-side hashes for Near-duplicate (current) and Diverged (history)
  // queries. Both empty when the file type has no text path or extraction
  // fails; we leave the fields undefined in that case so the historical
  // entry stays small and downstream readers know "no signal here."
  const text = await extractText(buf, fileType);
  if (text) {
    return {
      hash: content.hash,
      algo: content.algo,
      textHash: textContentHash(text),
      simhash64: simhash64(text),
    };
  }
  return { hash: content.hash, algo: content.algo };
}

async function processFile(
  jobId: string,
  file: DocumentFileRef,
  maxVersionsPerFile: number,
  observedAt: string,
): Promise<{ versionsProcessed: number; errors: string[] }> {
  const errors: string[] = [];
  let versionsProcessed = 0;

  // Historical versions (oldest-first, capped).
  let versions: SpVersion[] = [];
  try {
    versions = await fetchVersions(file.site, file.fileRef);
  } catch (e) {
    errors.push(`${file.fileRef} versions: ${(e as Error).message}`);
  }
  const historical = versions.filter((v) => !v.IsCurrentVersion);
  historical.sort((a, b) => a.Created.localeCompare(b.Created));
  const capped = historical.slice(-maxVersionsPerFile);

  const hashes: HistoricalHash[] = [];
  for (const v of capped) {
    try {
      const { hash, algo, textHash, simhash64: sim } = await downloadVersionAndHash(
        file.site,
        file.fileRef,
        v.ID,
        file.fileType,
      );
      hashes.push({
        hash,
        algo,
        observedAt: v.Created,
        observedByJobId: jobId,
        textHash,
        simhash64: sim,
      });
      versionsProcessed++;
    } catch (e) {
      errors.push(`${file.fileRef} v${v.VersionLabel ?? v.ID}: ${(e as Error).message}`);
    }
  }

  // Current version.
  const token = await getSpoToken();
  const encoded = file.fileRef.split("/").map(encodeURIComponent).join("/").replace(/'/g, "''");
  const currentUrl = `${SPO_ORIGIN}${file.site}/_api/web/getfilebyserverrelativeurl('${encoded}')/$value`;
  try {
    const res = await fetch(currentUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      errors.push(`${file.fileRef} current: ${res.status}`);
      return { versionsProcessed, errors };
    }
    const bytes = Buffer.from(await res.arrayBuffer());
    const { hash: currentSha, algo: currentAlgo } = await hashFileContent(bytes, file.fileType);
    const text = await extractText(bytes, file.fileType);
    const obs: FileObservation = {
      fileRef: file.fileRef,
      fileName: file.fileName,
      size: file.length,
      etag: file.etag ?? "",
      sha256: currentSha,
      algo: currentAlgo,
      textHash: text ? textContentHash(text) : undefined,
      simhash64: text ? simhash64(text) : undefined,
      sitePath: file.site,
      library: file.library,
    };
    await overlayVersionHistoryToStore(obs, hashes, jobId, observedAt);
  } catch (e) {
    errors.push(`${file.fileRef}: ${(e as Error).message}`);
  }

  return { versionsProcessed, errors };
}

async function workerHandler(queueItem: unknown, context: InvocationContext): Promise<void> {
  let msg: BootstrapQueueMessage;
  if (typeof queueItem === "string") {
    msg = JSON.parse(queueItem) as BootstrapQueueMessage;
  } else {
    msg = queueItem as BootstrapQueueMessage;
  }
  const { jobId, fileIndex } = msg;
  if (!jobId || typeof fileIndex !== "number") {
    context.error(`bootstrap-worker: malformed message ${JSON.stringify(msg)}`);
    return;
  }

  const job = await getBootstrapJob(jobId);
  if (!job) {
    context.error(`bootstrap-worker: job ${jobId} not found`);
    return;
  }
  if (job.status === "completed" || job.status === "failed") {
    context.warn(`bootstrap-worker: job ${jobId} already ${job.status}, ignoring`);
    return;
  }

  const manifest = await readBootstrapManifest(jobId);
  if (!manifest) {
    context.error(`bootstrap-worker: job ${jobId} has no manifest`);
    return;
  }

  if (fileIndex >= manifest.length) {
    // Done. Finalize.
    const fresh = await getBootstrapJob(jobId);
    if (!fresh) return;
    fresh.status = "completed";
    fresh.finishedAt = new Date().toISOString();
    fresh.currentFile = undefined;
    await updateBootstrapJob(fresh);
    context.log(
      `[bootstrap ${jobId}] DONE — ${fresh.filesCompleted}/${fresh.filesTotal} files, ${fresh.versionsProcessed} versions, ${fresh.errorCount} errors`,
    );
    return;
  }

  const file = manifest[fileIndex];

  if (job.status === "queued") job.status = "running";
  job.currentFile = file.fileRef;
  await updateBootstrapJob(job);

  context.log(
    `[bootstrap ${jobId}] file ${fileIndex + 1}/${manifest.length}: ${file.fileName}`,
  );

  const observedAt = new Date().toISOString();
  let result;
  try {
    result = await processFile(jobId, file, job.maxVersionsPerFile, observedAt);
  } catch (e) {
    result = { versionsProcessed: 0, errors: [`${file.fileRef}: ${(e as Error).message}`] };
  }

  const fresh = await getBootstrapJob(jobId);
  if (!fresh) return;
  fresh.filesCompleted = fresh.filesCompleted + 1;
  fresh.versionsProcessed = fresh.versionsProcessed + result.versionsProcessed;
  if (result.errors.length > 0) {
    fresh.errorCount = fresh.errorCount + result.errors.length;
    fresh.recentErrors = [...result.errors, ...fresh.recentErrors].slice(0, 10);
  }
  fresh.currentFile = file.fileRef;
  await updateBootstrapJob(fresh);

  await enqueueBootstrapMessage({ jobId, fileIndex: fileIndex + 1 });
}

app.storageQueue("duplicatesBootstrapWorker", {
  connection: "AzureWebJobsStorage",
  queueName: BOOTSTRAP_QUEUE_NAME,
  handler: workerHandler,
});
