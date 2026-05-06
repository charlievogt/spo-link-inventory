import { TableClient, TableServiceClient } from "@azure/data-tables";
import { BlobServiceClient } from "@azure/storage-blob";

/**
 * Persistence for tenant-wide link-inventory scan jobs.
 *
 * - Job header (status + counters) → Azure Table `LinkInventoryJobs`
 * - Per-job results (full link inventory) → blob `link-inventory-results/{jobId}.json`
 *
 * Why this split: Table is great for the small mutable counters that
 * need to be polled by the UI for progress, and bad for large blobs.
 * The full results JSON for a tenant scan can run 5–20MB which is
 * better as a blob fetched once when the user opens the results view.
 *
 * Storage account is the same one used by everything else in the
 * Function — `TABLE_CONNECTION_STRING` is reused.
 */

const JOBS_TABLE = "LinkInventoryJobs";
const RESULTS_CONTAINER = "link-inventory-results";
const PARTITION_KEY = "jobs";

let jobsTableClient: TableClient | null = null;
let blobServiceClient: BlobServiceClient | null = null;
let tableEnsured = false;
let containerEnsured = false;

function getConnString(): string {
  const cs = process.env.TABLE_CONNECTION_STRING;
  if (!cs) throw new Error("TABLE_CONNECTION_STRING not configured");
  return cs;
}

function getJobsTable(): TableClient {
  if (!jobsTableClient) jobsTableClient = TableClient.fromConnectionString(getConnString(), JOBS_TABLE);
  return jobsTableClient;
}

function getBlobService(): BlobServiceClient {
  if (!blobServiceClient) blobServiceClient = BlobServiceClient.fromConnectionString(getConnString());
  return blobServiceClient;
}

async function ensureTable(): Promise<void> {
  if (tableEnsured) return;
  const svc = TableServiceClient.fromConnectionString(getConnString());
  try {
    await svc.createTable(JOBS_TABLE);
  } catch (err) {
    if (!(err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 409)) {
      throw err;
    }
  }
  tableEnsured = true;
}

async function ensureContainer(): Promise<void> {
  if (containerEnsured) return;
  const container = getBlobService().getContainerClient(RESULTS_CONTAINER);
  await container.createIfNotExists();
  containerEnsured = true;
}

export type JobStatus = "queued" | "running" | "completed" | "failed";
export type JobKind = "pages" | "documents" | "orphan-recycle";

/**
 * The header row stored per-job. Counters get bumped as the orchestrator
 * progresses through sites/files; the UI polls this row.
 *
 * The same shape covers both page scans (sitesTotal × pagesPerSite) and
 * document scans (sitesTotal × filesPerSite). For document scans, the
 * "sites" field is the list of sites enumerated, and `filesTotal` /
 * `filesCompleted` track the per-file work. The `kind` discriminator
 * tells the worker which queue to consume from.
 */
export interface LinkInventoryJob {
  jobId: string;
  /** Defaults to `"pages"` for backwards compat with existing rows. */
  kind?: JobKind;
  status: JobStatus;
  startedAt: string; // ISO
  finishedAt?: string;
  /** List of site paths the orchestrator is processing, in order. */
  sites: string[];
  sitesTotal: number;
  sitesCompleted: number;
  /** Site currently being processed, if any. Cleared on completion. */
  currentSite?: string;
  pagesTotal: number;
  linksTotal: number;
  /** Document scans only: total files enumerated across all sites. */
  filesTotal?: number;
  /** Document scans only: files processed so far. */
  filesCompleted?: number;
  /** Document scans only: per-file size cap used during enumeration. */
  maxFileBytes?: number;
  /** Document scans only: ISO date for incremental rescan filter. */
  modifiedAfter?: string;
  /**
   * Document scans only: when true, the worker stops after the
   * enumerate phase and writes a preview-results blob with per-site
   * file counts. The user can then "promote" the preview to a real
   * scan via /scan-docs/promote, optionally narrowing to a subset.
   */
  previewOnly?: boolean;
  /**
   * When true, finalizeJob runs the opt-in SPO file verifier after the
   * aggregate is assembled: every AllItems.aspx?id=<path> link gets a
   * site-scoped HEAD check, and 404 hits are upgraded to
   * `malformed-spo-link`. Set from the `verifyFiles` field on the scan
   * request body. Default `false` — scans stay fast unless the admin
   * opts in.
   */
  verifyFiles?: boolean;
  errorCount: number;
  /** Up to N most recent error messages — older ones drop off. */
  recentErrors: string[];
  /** True once the results blob has been written. */
  resultsAvailable: boolean;
  /** Optional: who/what triggered this scan. */
  caller?: string;
  /**
   * When two jobs are created as a unified scan (one `pages`, one
   * `documents`), they cross-reference each other via this field. UI
   * uses it to group sibling rows into a single "unified scan" view in
   * the job dropdown and a combined progress display.
   */
  siblingJobId?: string;
}

interface JobEntity {
  partitionKey: string;
  rowKey: string;
  Kind?: JobKind;
  Status: JobStatus;
  StartedAt: string;
  FinishedAt?: string;
  SitesJson: string;
  SitesTotal: number;
  SitesCompleted: number;
  CurrentSite?: string;
  PagesTotal: number;
  LinksTotal: number;
  FilesTotal?: number;
  FilesCompleted?: number;
  MaxFileBytes?: number;
  ModifiedAfter?: string;
  PreviewOnly?: boolean;
  VerifyFiles?: boolean;
  ErrorCount: number;
  RecentErrorsJson: string;
  ResultsAvailable: boolean;
  Caller?: string;
  SiblingJobId?: string;
}

function toEntity(job: LinkInventoryJob): JobEntity {
  return {
    partitionKey: PARTITION_KEY,
    rowKey: job.jobId,
    Kind: job.kind,
    Status: job.status,
    StartedAt: job.startedAt,
    FinishedAt: job.finishedAt,
    SitesJson: JSON.stringify(job.sites),
    SitesTotal: job.sitesTotal,
    SitesCompleted: job.sitesCompleted,
    CurrentSite: job.currentSite,
    PagesTotal: job.pagesTotal,
    LinksTotal: job.linksTotal,
    FilesTotal: job.filesTotal,
    FilesCompleted: job.filesCompleted,
    MaxFileBytes: job.maxFileBytes,
    ModifiedAfter: job.modifiedAfter,
    PreviewOnly: job.previewOnly,
    VerifyFiles: job.verifyFiles,
    ErrorCount: job.errorCount,
    RecentErrorsJson: JSON.stringify(job.recentErrors),
    ResultsAvailable: job.resultsAvailable,
    Caller: job.caller,
    SiblingJobId: job.siblingJobId,
  };
}

function fromEntity(e: JobEntity): LinkInventoryJob {
  let sites: string[] = [];
  try { sites = JSON.parse(e.SitesJson) as string[]; } catch { /* ignore */ }
  let recentErrors: string[] = [];
  try { recentErrors = JSON.parse(e.RecentErrorsJson) as string[]; } catch { /* ignore */ }
  return {
    jobId: e.rowKey,
    kind: e.Kind,
    status: e.Status,
    startedAt: e.StartedAt,
    finishedAt: e.FinishedAt,
    sites,
    sitesTotal: e.SitesTotal,
    sitesCompleted: e.SitesCompleted,
    currentSite: e.CurrentSite,
    pagesTotal: e.PagesTotal,
    linksTotal: e.LinksTotal,
    filesTotal: e.FilesTotal,
    filesCompleted: e.FilesCompleted,
    maxFileBytes: e.MaxFileBytes,
    modifiedAfter: e.ModifiedAfter,
    previewOnly: e.PreviewOnly,
    verifyFiles: e.VerifyFiles,
    errorCount: e.ErrorCount,
    recentErrors,
    resultsAvailable: e.ResultsAvailable,
    caller: e.Caller,
    siblingJobId: e.SiblingJobId,
  };
}

export async function createJob(
  jobId: string,
  sites: string[],
  caller?: string,
  kind: JobKind = "pages",
  filesTotal?: number,
  docOptions?: { maxFileBytes?: number; modifiedAfter?: string; previewOnly?: boolean },
  verifyFiles?: boolean,
  siblingJobId?: string,
): Promise<LinkInventoryJob> {
  await ensureTable();
  const job: LinkInventoryJob = {
    jobId,
    kind,
    status: "queued",
    startedAt: new Date().toISOString(),
    sites,
    sitesTotal: sites.length,
    sitesCompleted: 0,
    pagesTotal: 0,
    linksTotal: 0,
    filesTotal,
    filesCompleted: kind === "documents" ? 0 : undefined,
    maxFileBytes: docOptions?.maxFileBytes,
    modifiedAfter: docOptions?.modifiedAfter,
    previewOnly: docOptions?.previewOnly,
    verifyFiles,
    errorCount: 0,
    recentErrors: [],
    resultsAvailable: false,
    caller,
    siblingJobId,
  };
  await getJobsTable().createEntity(toEntity(job));
  return job;
}

export async function updateJob(job: LinkInventoryJob): Promise<void> {
  await ensureTable();
  await getJobsTable().updateEntity(toEntity(job), "Replace");
}

export async function getJob(jobId: string): Promise<LinkInventoryJob | undefined> {
  await ensureTable();
  try {
    const e = await getJobsTable().getEntity<JobEntity>(PARTITION_KEY, jobId);
    return fromEntity(e);
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return undefined;
    }
    throw err;
  }
}

export async function listJobs(limit = 50): Promise<LinkInventoryJob[]> {
  await ensureTable();
  const out: LinkInventoryJob[] = [];
  const iter = getJobsTable().listEntities<JobEntity>({
    queryOptions: { filter: `PartitionKey eq '${PARTITION_KEY}'` },
  });
  for await (const e of iter) {
    out.push(fromEntity(e));
    if (out.length >= limit) break;
  }
  // Newest first
  out.sort((a, b) => b.startedAt.localeCompare(a.startedAt));
  return out;
}

export async function writeResults(jobId: string, payload: unknown): Promise<void> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(`${jobId}.json`);
  const data = JSON.stringify(payload);
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

export async function readResults(jobId: string): Promise<unknown | undefined> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(`${jobId}.json`);
  try {
    const dl = await blob.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    return JSON.parse(Buffer.concat(chunks).toString("utf8")) as unknown;
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return undefined;
    }
    throw err;
  }
}

/**
 * Per-site result storage for the queue-triggered orchestrator.
 *
 * Each site's scan output is written to its own blob at
 *   link-inventory-results/<jobId>/<siteIndex>.json
 *
 * On completion, mergePartialResults reads them all and writes the
 * aggregate to <jobId>.json (the same key the regular readResults
 * endpoint reads from). The intermediate blobs remain for diagnostics.
 *
 * Why per-site instead of one growing blob:
 *   - No read-modify-write contention if we ever go fan-out
 *   - Atomic writes — a worker crash can't corrupt the aggregate
 *   - Easy to retry just one site without re-running the whole job
 */

function partialBlobName(jobId: string, siteIndex: number): string {
  // Pad siteIndex so a directory listing sorts in scan order.
  return `${jobId}/${String(siteIndex).padStart(5, "0")}.json`;
}

export async function writePartialResult(jobId: string, siteIndex: number, payload: unknown): Promise<void> {
  await ensureContainer();
  const blob = getBlobService()
    .getContainerClient(RESULTS_CONTAINER)
    .getBlockBlobClient(partialBlobName(jobId, siteIndex));
  const data = JSON.stringify(payload);
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

/**
 * Persist the document-scan file manifest as a side blob. The doc
 * worker reads this once per message to get the file at fileIndex,
 * avoiding the need to stuff the entire manifest into the Table row
 * (Table cells max out at 64 KiB and a tenant scan can have tens of
 * thousands of files).
 */
export async function writeFileManifest(jobId: string, manifest: unknown): Promise<void> {
  await ensureContainer();
  const blob = getBlobService()
    .getContainerClient(RESULTS_CONTAINER)
    .getBlockBlobClient(`${jobId}/manifest.json`);
  const data = JSON.stringify(manifest);
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

export async function readFileManifest(jobId: string): Promise<unknown | undefined> {
  await ensureContainer();
  const blob = getBlobService()
    .getContainerClient(RESULTS_CONTAINER)
    .getBlockBlobClient(`${jobId}/manifest.json`);
  try {
    const dl = await blob.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    return JSON.parse(Buffer.concat(chunks).toString("utf8")) as unknown;
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return undefined;
    }
    throw err;
  }
}

/**
 * Delete every persisted artifact for a job:
 *   - the job header row in the LinkInventoryJobs table
 *   - the aggregate results blob `<jobId>.json`
 *   - every blob under the `<jobId>/` prefix (per-site partials,
 *     manifest fragments, file manifest, etc.)
 *
 * Idempotent — missing entries are not treated as errors. Returns
 * `{ rowDeleted, blobsDeleted }` for the caller to report. Used by
 * both the manual DELETE endpoint and the daily retention timer.
 */
export async function purgeJob(jobId: string): Promise<{ rowDeleted: boolean; blobsDeleted: number }> {
  await ensureTable();
  await ensureContainer();

  // 1) Delete the table row first. If this fails, the caller can
  //    retry without leaving a half-orphaned state.
  let rowDeleted = false;
  try {
    await getJobsTable().deleteEntity(PARTITION_KEY, jobId);
    rowDeleted = true;
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      // Already gone — fine.
      rowDeleted = false;
    } else {
      throw err;
    }
  }

  // 2) Delete every blob under `<jobId>` — that's the aggregate
  //    `<jobId>.json` AND everything under `<jobId>/` (partials,
  //    manifests, enum fragments). We list with the prefix and
  //    delete one by one.
  const container = getBlobService().getContainerClient(RESULTS_CONTAINER);
  let blobsDeleted = 0;
  // The aggregate blob `<jobId>.json` doesn't share the trailing
  // slash, so we need both prefixes to catch it.
  for (const prefix of [`${jobId}.json`, `${jobId}/`]) {
    for await (const blob of container.listBlobsFlat({ prefix })) {
      try {
        await container.getBlockBlobClient(blob.name).delete();
        blobsDeleted += 1;
      } catch (err) {
        if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
          // Race — someone else deleted it already. Skip.
          continue;
        }
        throw err;
      }
    }
  }

  return { rowDeleted, blobsDeleted };
}

/**
 * Read all per-site partial results for a job and return them in scan
 * order. Used by the queue worker on the last-site step to assemble
 * the final aggregate.
 *
 * Filters to numeric blob names so the manifest blob (`manifest.json`)
 * isn't accidentally included.
 */
export async function readAllPartialResults(jobId: string): Promise<unknown[]> {
  await ensureContainer();
  const container = getBlobService().getContainerClient(RESULTS_CONTAINER);
  const partials: Array<{ name: string; data: unknown }> = [];
  const prefix = `${jobId}/`;
  for await (const blob of container.listBlobsFlat({ prefix })) {
    const tail = blob.name.substring(prefix.length);
    // Only zero-padded numeric `00001.json` partials, not `manifest.json` or `enum/...`
    if (!/^\d+\.json$/.test(tail)) continue;
    const blobClient = container.getBlockBlobClient(blob.name);
    const dl = await blobClient.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    partials.push({ name: blob.name, data: JSON.parse(Buffer.concat(chunks).toString("utf8")) as unknown });
  }
  partials.sort((a, b) => a.name.localeCompare(b.name));
  return partials.map((p) => p.data);
}

/**
 * Per-site manifest fragment storage. The doc-scan enumerate phase
 * walks one site per queue message and writes the enumerated files
 * for that site as a fragment blob. Once enumeration is complete, the
 * worker reads all fragments and consolidates them into the final
 * `manifest.json` blob.
 *
 * Path: `<jobId>/enum/<siteIndex>.json`
 */
function manifestFragmentName(jobId: string, siteIndex: number): string {
  return `${jobId}/enum/${String(siteIndex).padStart(5, "0")}.json`;
}

export async function writeManifestFragment(
  jobId: string,
  siteIndex: number,
  files: unknown,
): Promise<void> {
  await ensureContainer();
  const blob = getBlobService()
    .getContainerClient(RESULTS_CONTAINER)
    .getBlockBlobClient(manifestFragmentName(jobId, siteIndex));
  const data = JSON.stringify(files);
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

/**
 * Read all manifest fragments and return their concatenation in
 * site-index order. Used at the end of the enumerate phase to build
 * the final manifest.
 */
export async function readAllManifestFragments(jobId: string): Promise<unknown[]> {
  await ensureContainer();
  const container = getBlobService().getContainerClient(RESULTS_CONTAINER);
  const fragments: Array<{ name: string; data: unknown[] }> = [];
  const prefix = `${jobId}/enum/`;
  for await (const blob of container.listBlobsFlat({ prefix })) {
    const blobClient = container.getBlockBlobClient(blob.name);
    const dl = await blobClient.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    fragments.push({
      name: blob.name,
      data: JSON.parse(Buffer.concat(chunks).toString("utf8")) as unknown[],
    });
  }
  fragments.sort((a, b) => a.name.localeCompare(b.name));
  // Flatten all fragments into one list
  const out: unknown[] = [];
  for (const f of fragments) for (const item of f.data) out.push(item);
  return out;
}

