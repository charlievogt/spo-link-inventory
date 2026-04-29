import { app, type InvocationContext } from "@azure/functions";
import {
  getJob,
  updateJob,
  writePartialResult,
  readAllPartialResults,
  writeResults,
  readFileManifest,
  writeFileManifest,
  writeManifestFragment,
  readAllManifestFragments,
} from "../services/linkInventoryJobStore.js";
import { scanDocumentFile, type DocumentFileInventory } from "../services/documentLinkScanner.js";
import { enumerateFiles, type DocumentFileRef } from "../services/spoFilesEnumerator.js";
import { verifyMalformedLinks } from "../services/linkVerifier.js";
import {
  enqueueDocsScanMessage,
  SCAN_DOCS_QUEUE_NAME,
  type DocsScanQueueMessage,
} from "../services/linkInventoryDocsQueue.js";
import { mergeScanResultsToStore } from "../services/backlinksIndexStore.js";
import { mergeObservationsToStore, type FileObservation } from "../services/hashIndexStore.js";
import {
  loadSiteAggregate,
  updateSiteAggregate,
  mergeDocScanIntoAggregate,
  carryForwardEntries,
  type AggregateFileEntry,
  type FileScanResult,
} from "../services/aggregateStore.js";
import type { ISiteInventoryShape } from "./linkInventoryReplace.types.js";

/**
 * Document scan worker — two-phase pipeline.
 *
 * Phase 1 — `enumerate`: walk one site per message, list parseable
 * files, write a manifest fragment, enqueue next site OR consolidate
 * fragments and switch to scan phase.
 *
 * Phase 2 — `scan`: open one file per message, extract+classify links,
 * write a partial result, enqueue next file OR finalize aggregate.
 *
 * The split keeps each message under the Function timeout cap and
 * gives the UI live progress through both phases.
 */

async function workerHandler(queueItem: unknown, context: InvocationContext): Promise<void> {
  let msg: DocsScanQueueMessage;
  if (typeof queueItem === "string") {
    msg = JSON.parse(queueItem) as DocsScanQueueMessage;
  } else {
    msg = queueItem as DocsScanQueueMessage;
  }

  const { jobId } = msg;
  // Default phase to "scan" for backwards compat with messages that
  // pre-date the two-phase split.
  const phase = msg.phase ?? "scan";
  if (!jobId) {
    context.error(`docs-worker: malformed message (no jobId) ${JSON.stringify(msg)}`);
    return;
  }
  if (phase === "enumerate") {
    return enumerateHandler(msg, context);
  }
  return scanHandler(msg, context);
}

async function enumerateHandler(msg: DocsScanQueueMessage, context: InvocationContext): Promise<void> {
  const { jobId, siteIndex } = msg;
  if (!jobId || typeof siteIndex !== "number") {
    context.error(`docs-worker enum: malformed message ${JSON.stringify(msg)}`);
    return;
  }

  const job = await getJob(jobId);
  if (!job) {
    context.error(`docs-worker enum: job ${jobId} not found, dropping`);
    return;
  }
  if (job.status === "completed" || job.status === "failed") {
    context.warn(`docs-worker enum: job ${jobId} already ${job.status}, ignoring`);
    return;
  }
  if (siteIndex >= job.sites.length) {
    context.warn(`docs-worker enum: job ${jobId} siteIndex=${siteIndex} out of range`);
    return;
  }

  const sitePath = job.sites[siteIndex];
  const maxFileBytes = job.maxFileBytes ?? 100 * 1024 * 1024;
  const modifiedAfter = job.modifiedAfter ? new Date(job.modifiedAfter) : undefined;

  // Mark running on first message
  if (job.status === "queued") job.status = "running";
  job.currentSite = `Discovering files: ${sitePath}`;
  await updateJob(job);

  context.log(
    `[scan-docs ${jobId}] enumerate ${siteIndex + 1}/${job.sites.length}: ${sitePath}`,
  );

  let files: DocumentFileRef[] = [];
  try {
    files = await enumerateFiles(sitePath, {
      maxBytes: maxFileBytes,
      modifiedAfter,
      includeOther: true,
    });
    await writeManifestFragment(jobId, siteIndex, files);
    context.log(`[scan-docs ${jobId}] ${sitePath}: enumerated ${files.length} files`);
  } catch (err) {
    context.error(`[scan-docs ${jobId}] enumerate ${sitePath} failed: ${(err as Error).message}`);
    // Persist an empty fragment so the consolidation step still finds it
    await writeManifestFragment(jobId, siteIndex, []);
    const fresh = await getJob(jobId);
    if (fresh) {
      fresh.errorCount += 1;
      fresh.recentErrors = [
        `enumerate ${sitePath}: ${(err as Error).message}`,
        ...fresh.recentErrors,
      ].slice(0, 10);
      await updateJob(fresh);
    }
  }

  // Update file count on the job
  const fresh = await getJob(jobId);
  if (!fresh) {
    context.error(`docs-worker enum: job ${jobId} disappeared after enumerate`);
    return;
  }
  fresh.filesTotal = (fresh.filesTotal ?? 0) + files.length;
  fresh.sitesCompleted = siteIndex + 1;
  await updateJob(fresh);

  // Chain to next site or transition to scan phase
  const nextSiteIndex = siteIndex + 1;
  if (nextSiteIndex < job.sites.length) {
    await enqueueDocsScanMessage({ jobId, phase: "enumerate", siteIndex: nextSiteIndex });
    return;
  }

  // Last site enumerated — consolidate manifest fragments
  context.log(`[scan-docs ${jobId}] enumerate complete, consolidating manifest`);
  const allFiles = (await readAllManifestFragments(jobId)) as DocumentFileRef[];
  if (allFiles.length === 0) {
    // Nothing to scan — finalize as completed-empty
    context.warn(`[scan-docs ${jobId}] no files enumerated, finalizing empty`);
    const empty = await getJob(jobId);
    if (empty) {
      empty.currentSite = undefined;
      empty.finishedAt = new Date().toISOString();
      empty.status = "completed";
      empty.resultsAvailable = true;
      await updateJob(empty);
      // Write an empty results blob so the UI can render
      await writeResults(jobId, {
        jobId,
        kind: "documents",
        startedAt: empty.startedAt,
        finishedAt: empty.finishedAt,
        sites: [],
        totals: { sites: 0, pages: 0, links: 0, errors: 0 },
      });
    }
    return;
  }

  // Branch on previewOnly: if set, skip scan phase entirely and write
  // a per-site/per-library count summary as the results blob.
  const ready = await getJob(jobId);
  if (!ready) return;

  if (ready.previewOnly) {
    // Preview keeps the full manifest (supported + other) so the admin
    // can see the whole library and opt extensions in.
    await writeFileManifest(jobId, allFiles);
    context.log(`[scan-docs ${jobId}] preview manifest written: ${allFiles.length} files`);
    ready.filesTotal = allFiles.length;

    context.log(`[scan-docs ${jobId}] preview-only mode — writing summary and stopping`);
    // Build per-site / per-library counts, split into included (supported
    // for link extraction today) and excluded (everything else). Also
    // produce a tenant-wide byExtension breakdown per bucket so the UI
    // can render "opt in to .jpg" style checkboxes.
    interface BucketStats {
      files: number;
      totalBytes: number;
      byExtension: Record<string, { count: number; totalBytes: number }>;
    }
    interface SiteSummary {
      site: string;
      libraries: Record<string, { files: number; totalBytes: number }>;
      totalFiles: number;
      totalBytes: number;
      included: BucketStats;
      excluded: BucketStats;
    }
    const emptyBucket = (): BucketStats => ({ files: 0, totalBytes: 0, byExtension: {} });
    const tenantIncluded: BucketStats = emptyBucket();
    const tenantExcluded: BucketStats = emptyBucket();
    const bump = (b: BucketStats, ext: string, size: number): void => {
      b.files += 1;
      b.totalBytes += size;
      const slot = b.byExtension[ext] ?? { count: 0, totalBytes: 0 };
      slot.count += 1;
      slot.totalBytes += size;
      b.byExtension[ext] = slot;
    };
    const sitesMap = new Map<string, SiteSummary>();
    for (const f of allFiles) {
      let entry = sitesMap.get(f.site);
      if (!entry) {
        entry = {
          site: f.site,
          libraries: {},
          totalFiles: 0,
          totalBytes: 0,
          included: emptyBucket(),
          excluded: emptyBucket(),
        };
        sitesMap.set(f.site, entry);
      }
      const lib = entry.libraries[f.library] ?? { files: 0, totalBytes: 0 };
      lib.files += 1;
      lib.totalBytes += f.length;
      entry.libraries[f.library] = lib;
      entry.totalFiles += 1;
      entry.totalBytes += f.length;
      const bucket = f.fileType === "other" ? entry.excluded : entry.included;
      const tenantBucket = f.fileType === "other" ? tenantExcluded : tenantIncluded;
      bump(bucket, f.extension, f.length);
      bump(tenantBucket, f.extension, f.length);
    }
    const siteSummaries = Array.from(sitesMap.values()).sort((a, b) => b.totalFiles - a.totalFiles);

    await writeResults(jobId, {
      jobId,
      kind: "documents-preview",
      startedAt: ready.startedAt,
      finishedAt: new Date().toISOString(),
      previewOnly: true,
      siteSummaries,
      included: tenantIncluded,
      excluded: tenantExcluded,
      totals: {
        sites: siteSummaries.length,
        files: allFiles.length,
        bytes: allFiles.reduce((s, f) => s + f.length, 0),
      },
    });

    ready.currentSite = undefined;
    ready.finishedAt = new Date().toISOString();
    ready.status = "completed";
    ready.resultsAvailable = true;
    await updateJob(ready);
    context.log(
      `[scan-docs ${jobId}] preview DONE — ${siteSummaries.length} sites, ${allFiles.length} files ` +
      `(included=${tenantIncluded.files}, excluded=${tenantExcluded.files})`,
    );
    return;
  }

  // Normal scan phase: filter the manifest down to extraction-supported
  // types before the scan phase runs. "other" files live only in the
  // preview manifest today; hashing for duplicate detection will pick
  // them up in a later step.
  const scannable = allFiles.filter((f) => f.fileType !== "other");
  await writeFileManifest(jobId, scannable);
  context.log(
    `[scan-docs ${jobId}] manifest written: ${scannable.length} scannable ` +
    `(filtered ${allFiles.length - scannable.length} other)`,
  );
  ready.filesTotal = scannable.length;

  if (scannable.length === 0) {
    context.warn(`[scan-docs ${jobId}] no scannable files after filtering, finalizing empty`);
    ready.currentSite = undefined;
    ready.finishedAt = new Date().toISOString();
    ready.status = "completed";
    ready.resultsAvailable = true;
    await updateJob(ready);
    await writeResults(jobId, {
      jobId,
      kind: "documents",
      startedAt: ready.startedAt,
      finishedAt: ready.finishedAt,
      sites: [],
      totals: { sites: 0, pages: 0, links: 0, errors: 0 },
    });
    return;
  }

  ready.currentSite = "Starting file scan...";
  await updateJob(ready);
  context.log(`[scan-docs ${jobId}] starting scan phase`);
  await enqueueDocsScanMessage({ jobId, phase: "scan", fileIndex: 0 });
}

async function scanHandler(msg: DocsScanQueueMessage, context: InvocationContext): Promise<void> {
  const { jobId, fileIndex } = msg;
  if (!jobId || typeof fileIndex !== "number") {
    context.error(`docs-worker scan: malformed message ${JSON.stringify(msg)}`);
    return;
  }

  const job = await getJob(jobId);
  if (!job) {
    context.error(`docs-worker: job ${jobId} not found, dropping message`);
    return;
  }
  if (job.status === "completed" || job.status === "failed") {
    context.warn(`docs-worker: job ${jobId} already ${job.status}, ignoring fileIndex=${fileIndex}`);
    return;
  }

  // Load the file manifest. Cached per Function instance via the blob
  // service client, so re-reads are cheap.
  const manifest = (await readFileManifest(jobId)) as DocumentFileRef[] | undefined;
  if (!manifest) {
    context.error(`docs-worker: job ${jobId} has no manifest blob`);
    return;
  }
  if (fileIndex >= manifest.length) {
    context.warn(`docs-worker: job ${jobId} fileIndex=${fileIndex} out of range (${manifest.length})`);
    await finalizeJob(jobId, context);
    return;
  }

  const file = manifest[fileIndex];

  // Mark running on first message
  if (job.status === "queued") job.status = "running";
  job.currentSite = file.site;
  await updateJob(job);

  context.log(
    `[scan-docs ${jobId}] file ${fileIndex + 1}/${manifest.length} ` +
    `(${file.site} → ${file.fileName}, ${(file.length / 1024).toFixed(0)}KB)`,
  );

  // Look up the previous aggregate entry for this file — enables the
  // ETag-skip optimization in scanDocumentFile when the file hasn't
  // changed since last scan. Aggregate read failures degrade to "no
  // skip" (full scan path) — never throw out of the worker.
  let previousEntry: AggregateFileEntry | undefined;
  try {
    const siteAggregate = await loadSiteAggregate(file.site);
    previousEntry = siteAggregate.files[file.fileRef];
  } catch (err) {
    context.warn(
      `[scan-docs ${jobId}] aggregate read failed for ${file.site} (${(err as Error).message}); proceeding with full scan`,
    );
  }

  let result: DocumentFileInventory;
  try {
    result = await scanDocumentFile(file, { previousEntry });
  } catch (err) {
    // scanDocumentFile catches its own errors but defensive double-cover
    result = {
      ...file,
      links: [],
      scanMs: 0,
      error: (err as Error).message,
    };
  }

  // Persist per-file result
  await writePartialResult(jobId, fileIndex, result);

  // Update counters (re-read for safety)
  const fresh = await getJob(jobId);
  if (!fresh) {
    context.error(`docs-worker: job ${jobId} disappeared after scan`);
    return;
  }
  fresh.filesCompleted = (fresh.filesCompleted ?? 0) + 1;
  fresh.linksTotal += result.links.length;
  if (result.error) {
    fresh.errorCount += 1;
    fresh.recentErrors = [`${file.fileRef}: ${result.error}`, ...fresh.recentErrors].slice(0, 10);
  } else {
    context.log(`[scan-docs ${jobId}] ${file.fileName} → ${result.links.length}L in ${result.scanMs}ms`);
  }
  fresh.currentSite = file.site;
  await updateJob(fresh);

  // Chain to next file or finalize
  const nextIndex = fileIndex + 1;
  if (nextIndex < manifest.length) {
    await enqueueDocsScanMessage({ jobId, phase: "scan", fileIndex: nextIndex });
  } else {
    await finalizeJob(jobId, context);
  }
}

async function finalizeJob(jobId: string, context: InvocationContext): Promise<void> {
  const job = await getJob(jobId);
  if (!job) return;
  context.log(`[scan-docs ${jobId}] finalizing — assembling aggregate`);

  const partials = (await readAllPartialResults(jobId)) as DocumentFileInventory[];

  // Group per-file results by site so the UI can show per-site rollups
  // matching the page-scan view shape.
  const sitesMap = new Map<string, DocumentFileInventory[]>();
  for (const f of partials) {
    if (!sitesMap.has(f.site)) sitesMap.set(f.site, []);
    sitesMap.get(f.site)!.push(f);
  }

  const sites = Array.from(sitesMap.entries()).map(([site, files]) => {
    const linkCount = files.reduce((s, f) => s + f.links.length, 0);
    const byClass: Record<string, number> = {};
    const bySource: Record<string, number> = {};
    for (const f of files) {
      for (const l of f.links) {
        byClass[l.linkClass] = (byClass[l.linkClass] ?? 0) + 1;
        bySource[l.source] = (bySource[l.source] ?? 0) + 1;
      }
    }
    return {
      site,
      pageCount: files.length, // "files" rendered as "pages" for UI compat
      linkCount,
      byClass,
      bySource,
      pages: files.map((f) => ({
        // Map document file → page-shaped record so the existing UI
        // table can render it without changes
        pageId: 0,
        pageTitle: f.fileName,
        pageUrl: f.fileRef,
        modified: f.modified,
        etag: f.etag,
        links: f.links,
        parseError: f.error,
      })),
      scanMs: 0,
    };
  });

  // Opt-in SPO file verification: HEAD-check every AllItems.aspx `?id=`
  // link and upgrade 404s to `malformed-spo-link`. Mutates in place.
  if (job.verifyFiles) {
    context.log(`[scan-docs ${jobId}] verifyFiles=true — running malformed-link verifier`);
    try {
      const stats = await verifyMalformedLinks(sites, {
        onWarn: (m) => context.warn(`[scan-docs ${jobId}] verifier: ${m}`),
      });
      context.log(
        `[scan-docs ${jobId}] verifier: ${stats.candidates} candidates, ` +
        `${stats.notFound} 404s, ${stats.errors} errors, ${stats.durationMs}ms`,
      );
    } catch (e) {
      context.warn(`[scan-docs ${jobId}] verifier failed (non-fatal): ${(e as Error).message}`);
    }
  }

  const aggregate = {
    jobId,
    kind: "documents",
    startedAt: job.startedAt,
    finishedAt: new Date().toISOString(),
    sites,
    totals: {
      sites: sites.length,
      pages: partials.length,
      links: partials.reduce((s, f) => s + f.links.length, 0),
      errors: partials.filter((f) => f.error).length,
    },
  };

  await writeResults(jobId, aggregate);

  const fresh = (await getJob(jobId)) ?? job;
  fresh.currentSite = undefined;
  fresh.finishedAt = aggregate.finishedAt;
  fresh.status = fresh.errorCount === (fresh.filesTotal ?? 0) && (fresh.filesTotal ?? 0) > 0 ? "failed" : "completed";
  fresh.resultsAvailable = true;
  await updateJob(fresh);
  context.log(
    `[scan-docs ${jobId}] DONE — ${fresh.status}, ${fresh.filesCompleted}/${fresh.filesTotal} files, ${fresh.linksTotal} links`,
  );

  // Merge into the persistent backlinks index. The doc scan's `sites`
  // array is already shaped like ISiteInventoryShape (pageId=0, pageTitle=
  // fileName, pageUrl=fileRef) so the same builder handles both kinds.
  if (fresh.status === "completed") {
    try {
      await mergeScanResultsToStore(
        sites as ISiteInventoryShape[],
        job.sites,
        "document",
        jobId,
        aggregate.finishedAt,
      );
      context.log(`[scan-docs ${jobId}] backlinks index updated`);
    } catch (e) {
      context.warn(`[scan-docs ${jobId}] backlinks index merge failed (non-fatal): ${(e as Error).message}`);
    }

    // Merge into the persistent file-hash index for duplicate detection.
    // Only include files that produced a hash — download-failed entries
    // have no sha256 and shouldn't rotate history on the next clean run.
    try {
      const observations: FileObservation[] = [];
      for (const f of partials) {
        if (!f.sha256) continue;
        observations.push({
          fileRef: f.fileRef,
          fileName: f.fileName,
          size: f.length,
          etag: f.etag ?? "",
          sha256: f.sha256,
          algo: f.hashAlgo,
          textHash: f.textHash,
          simhash64: f.simhash64,
          sitePath: f.site,
          library: f.library,
        });
      }
      if (observations.length > 0) {
        await mergeObservationsToStore(observations, jobId, aggregate.finishedAt);
        context.log(
          `[scan-docs ${jobId}] hash index updated with ${observations.length} observations`,
        );
      }
    } catch (e) {
      context.warn(`[scan-docs ${jobId}] hash index merge failed (non-fatal): ${(e as Error).message}`);
    }

    // Roll forward the per-site aggregate — this is the persistent
    // "current state of all links" that survives the 30-day per-job
    // retention purge and powers ETag-skip on future scans. Group
    // partials by site, then for each site:
    //   - actually-scanned files → mergeDocScanIntoAggregate (replaces entry)
    //   - etag-skipped files → carryForwardEntries (bumps lastConfirmedAt only)
    try {
      const bySiteScanned = new Map<string, FileScanResult[]>();
      const bySiteSkipped = new Map<string, string[]>();
      for (const f of partials) {
        if (f.error) continue; // download or extract failure — don't pollute aggregate
        if (f.reusedFromPreviousScan) {
          if (!bySiteSkipped.has(f.site)) bySiteSkipped.set(f.site, []);
          bySiteSkipped.get(f.site)!.push(f.fileRef);
        } else {
          if (!bySiteScanned.has(f.site)) bySiteScanned.set(f.site, []);
          bySiteScanned.get(f.site)!.push({
            fileRef: f.fileRef,
            fileName: f.fileName,
            fileType: f.fileType,
            modified: f.modified,
            etag: f.etag ?? "",
            size: f.length,
            links: f.links,
            scanError: f.error,
          });
        }
      }
      const allSites = new Set([...bySiteScanned.keys(), ...bySiteSkipped.keys()]);
      let totalScanned = 0;
      let totalSkipped = 0;
      for (const site of allSites) {
        const scanned = bySiteScanned.get(site) ?? [];
        const skipped = bySiteSkipped.get(site) ?? [];
        totalScanned += scanned.length;
        totalSkipped += skipped.length;
        await updateSiteAggregate(site, (current) => {
          let next = current;
          if (scanned.length > 0) {
            next = mergeDocScanIntoAggregate(next, scanned, jobId, aggregate.finishedAt);
          }
          if (skipped.length > 0) {
            next = carryForwardEntries(next, [], skipped, jobId, aggregate.finishedAt);
          }
          return next;
        });
      }
      context.log(
        `[scan-docs ${jobId}] aggregate updated: ${totalScanned} files merged, ${totalSkipped} carried forward across ${allSites.size} sites`,
      );
    } catch (e) {
      context.warn(`[scan-docs ${jobId}] aggregate merge failed (non-fatal): ${(e as Error).message}`);
    }
  }
}

app.storageQueue("linkInventoryScanDocsWorker", {
  connection: "AzureWebJobsStorage",
  queueName: SCAN_DOCS_QUEUE_NAME,
  handler: workerHandler,
});
