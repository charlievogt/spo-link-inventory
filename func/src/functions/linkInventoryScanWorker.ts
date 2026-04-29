import { app, type InvocationContext } from "@azure/functions";
import {
  getJob,
  updateJob,
  writePartialResult,
  readAllPartialResults,
  writeResults,
  readFileManifest,
  type LinkInventoryJob,
} from "../services/linkInventoryJobStore.js";
import { scanSite, type SiteInventory, type PreviousPagesByUrl } from "../services/linkInventoryScanner.js";
import { enqueueScanMessage, SCAN_QUEUE_NAME, type ScanQueueMessage } from "../services/linkInventoryQueue.js";
import { verifyMalformedLinks } from "../services/linkVerifier.js";
import { mergeScanResultsToStore } from "../services/backlinksIndexStore.js";
import {
  loadSiteAggregate,
  updateSiteAggregate,
  mergePageScanIntoAggregate,
  carryForwardEntries,
  type PageScanResult,
} from "../services/aggregateStore.js";

/**
 * Queue worker — processes one site per message, then enqueues the
 * next message in the chain. The scan endpoint enqueues the first
 * message (siteIndex = 0); each invocation here either chains forward
 * by sending a new message, or finalizes the job when the last site
 * is done.
 *
 * This replaces the synchronous orchestrator's per-job loop. The
 * advantages:
 *   - No 230s function timeout cap on a single job
 *   - Survives Function host restarts (queue messages persist)
 *   - Each message has its own context.log for clean diagnostics
 *
 * The trade-off vs synchronous:
 *   - Per-message cold-start cost. For very small site counts (< 5)
 *     synchronous would be faster. We don't optimize for that case
 *     because the user will rarely run 5-site scans in production.
 */

async function workerHandler(queueItem: unknown, context: InvocationContext): Promise<void> {
  // Storage Queue messages come in base64 by default; the runtime
  // decodes the outer base64 layer for us, leaving us with a string
  // body. We JSON.parse it here.
  let msg: ScanQueueMessage;
  if (typeof queueItem === "string") {
    msg = JSON.parse(queueItem) as ScanQueueMessage;
  } else {
    msg = queueItem as ScanQueueMessage;
  }

  const { jobId, siteIndex } = msg;
  if (!jobId || typeof siteIndex !== "number") {
    context.error(`worker: malformed message ${JSON.stringify(msg)}`);
    return;
  }

  const job = await getJob(jobId);
  if (!job) {
    context.error(`worker: job ${jobId} not found, dropping message`);
    return;
  }
  if (job.status === "completed" || job.status === "failed") {
    context.warn(`worker: job ${jobId} already ${job.status}, ignoring siteIndex=${siteIndex}`);
    return;
  }
  if (siteIndex >= job.sites.length) {
    context.warn(`worker: job ${jobId} siteIndex=${siteIndex} out of range (${job.sites.length})`);
    await finalizeJob(job, context);
    return;
  }

  const sitePath = job.sites[siteIndex];

  // Mark the job as running and update currentSite. First message
  // also flips status from queued → running.
  if (job.status === "queued") job.status = "running";
  job.currentSite = sitePath;
  await updateJob(job);

  // Optional targeted page-id filter from the job's manifest blob.
  // For non-targeted scans this read returns undefined and we scan
  // the whole site.
  let pageIdsForSite: number[] | undefined;
  try {
    const manifest = await readFileManifest(jobId) as { pageIdsBySite?: Record<string, number[]> } | undefined;
    if (manifest?.pageIdsBySite && manifest.pageIdsBySite[sitePath]) {
      pageIdsForSite = manifest.pageIdsBySite[sitePath];
    }
  } catch {
    // No manifest blob — full scan
  }

  // Load previous aggregate for this site so unchanged pages can
  // short-circuit the canvas parse via Modified-skip. Aggregate read
  // failure is non-fatal — we just lose the optimization for this site.
  let previousPages: PreviousPagesByUrl | undefined;
  try {
    const siteAggregate = await loadSiteAggregate(sitePath);
    if (Object.keys(siteAggregate.pages).length > 0) {
      previousPages = {};
      for (const [pageUrl, entry] of Object.entries(siteAggregate.pages)) {
        previousPages[pageUrl] = { modified: entry.modified, links: entry.links };
      }
    }
  } catch (err) {
    context.warn(
      `[scan ${jobId}] aggregate read failed for ${sitePath} (${(err as Error).message}); proceeding with full scan`,
    );
  }

  context.log(
    `[scan ${jobId}] site ${siteIndex + 1}/${job.sites.length}: ${sitePath}` +
    (pageIdsForSite ? ` (${pageIdsForSite.length} targeted pages)` : '') +
    (previousPages ? ` (delta-eligible: ${Object.keys(previousPages).length} prior pages)` : ''),
  );

  let result: SiteInventory;
  try {
    result = await scanSite(sitePath, pageIdsForSite, { previousPages });
  } catch (err) {
    // scanSite never throws (it returns SiteInventory with `error` set),
    // but defensive — write a placeholder error result so the chain
    // continues even on unexpected failures.
    result = {
      site: sitePath,
      pageCount: 0,
      linkCount: 0,
      byClass: {},
      bySource: {},
      pages: [],
      scanMs: 0,
      error: (err as Error).message,
    };
  }

  // Persist the per-site result to its own blob.
  await writePartialResult(jobId, siteIndex, result);

  // Update job counters. Re-read first to pick up any concurrent
  // updates (defensive, since per-job processing is sequential there
  // shouldn't be concurrent updates, but Table Storage updateEntity
  // does a Replace which would clobber otherwise).
  const fresh = await getJob(jobId);
  if (!fresh) {
    context.error(`worker: job ${jobId} disappeared after scan`);
    return;
  }
  fresh.sitesCompleted += 1;
  fresh.pagesTotal += result.pageCount;
  fresh.linksTotal += result.linkCount;
  if (result.error) {
    fresh.errorCount += 1;
    fresh.recentErrors = [`${sitePath}: ${result.error}`, ...fresh.recentErrors].slice(0, 10);
    context.error(`[scan ${jobId}] ${sitePath} FAILED: ${result.error}`);
  } else {
    context.log(`[scan ${jobId}] ${sitePath} → ${result.pageCount}p / ${result.linkCount}L in ${result.scanMs}ms`);
  }
  fresh.currentSite = sitePath;
  await updateJob(fresh);

  // Either chain to the next site or finalize.
  const nextIndex = siteIndex + 1;
  if (nextIndex < fresh.sites.length) {
    await enqueueScanMessage({ jobId, siteIndex: nextIndex });
  } else {
    await finalizeJob(fresh, context);
  }
}

/**
 * Read all per-site partial results, write the aggregate blob, and
 * mark the job complete.
 */
async function finalizeJob(job: LinkInventoryJob, context: InvocationContext): Promise<void> {
  context.log(`[scan ${job.jobId}] finalizing — assembling aggregate from ${job.sites.length} partials`);
  const partials = (await readAllPartialResults(job.jobId)) as SiteInventory[];

  // Opt-in SPO file verification: mutates link classifications in place
  // before we write the aggregate. Safe to skip on any error — falls
  // back to the heuristic-only classification.
  if (job.verifyFiles) {
    context.log(`[scan ${job.jobId}] verifyFiles=true — running malformed-link verifier`);
    try {
      const stats = await verifyMalformedLinks(partials, {
        onWarn: (m) => context.warn(`[scan ${job.jobId}] verifier: ${m}`),
      });
      context.log(
        `[scan ${job.jobId}] verifier: ${stats.candidates} candidates, ` +
        `${stats.notFound} 404s, ${stats.errors} errors, ${stats.durationMs}ms`,
      );
    } catch (e) {
      context.warn(`[scan ${job.jobId}] verifier failed (non-fatal): ${(e as Error).message}`);
    }
  }

  const aggregate = {
    jobId: job.jobId,
    startedAt: job.startedAt,
    finishedAt: new Date().toISOString(),
    sites: partials,
    totals: {
      sites: partials.length,
      pages: partials.reduce((s, r) => s + r.pageCount, 0),
      links: partials.reduce((s, r) => s + r.linkCount, 0),
      errors: partials.filter((r) => r.error).length,
    },
  };

  await writeResults(job.jobId, aggregate);

  // Re-read once more to get the latest counters before flipping status.
  const fresh = (await getJob(job.jobId)) ?? job;
  fresh.currentSite = undefined;
  fresh.finishedAt = aggregate.finishedAt;
  fresh.status = fresh.errorCount === fresh.sitesTotal && fresh.sitesTotal > 0 ? "failed" : "completed";
  fresh.resultsAvailable = true;
  await updateJob(fresh);
  context.log(
    `[scan ${job.jobId}] DONE — ${fresh.status}, ${fresh.pagesTotal}p / ${fresh.linksTotal}L across ${fresh.sitesCompleted}/${fresh.sitesTotal} sites`,
  );

  // Merge into the persistent backlinks index (non-fatal — the scan
  // succeeded even if this fails; an admin can trigger a manual
  // rebuild if the index ever drifts).
  if (fresh.status === "completed") {
    try {
      await mergeScanResultsToStore(partials, job.sites, "page", job.jobId, aggregate.finishedAt);
      context.log(`[scan ${job.jobId}] backlinks index updated`);
    } catch (e) {
      context.warn(`[scan ${job.jobId}] backlinks index merge failed (non-fatal): ${(e as Error).message}`);
    }

    // Roll forward the per-site page aggregate. Each site partial maps
    // to one site shard. Pages within: those with `reusedFromPreviousScan`
    // bump lastConfirmedAt only; the rest replace the entry. Failed
    // sites (with `error`) are skipped — don't overwrite a possibly-
    // good aggregate with an empty result from a transient SP error.
    try {
      let totalScanned = 0;
      let totalSkipped = 0;
      for (const sitePartial of partials) {
        if (sitePartial.error) continue;
        const scanned: PageScanResult[] = [];
        const skipped: string[] = [];
        for (const p of sitePartial.pages) {
          if (p.reusedFromPreviousScan) {
            skipped.push(p.pageUrl);
          } else if (!p.parseError) {
            scanned.push({
              pageUrl: p.pageUrl,
              pageTitle: p.pageTitle,
              modified: p.modified ?? "",
              links: p.links,
            });
          }
        }
        totalScanned += scanned.length;
        totalSkipped += skipped.length;
        if (scanned.length === 0 && skipped.length === 0) continue;
        await updateSiteAggregate(sitePartial.site, (current) => {
          let next = current;
          if (scanned.length > 0) {
            next = mergePageScanIntoAggregate(next, scanned, job.jobId, aggregate.finishedAt);
          }
          if (skipped.length > 0) {
            next = carryForwardEntries(next, skipped, [], job.jobId, aggregate.finishedAt);
          }
          return next;
        });
      }
      context.log(
        `[scan ${job.jobId}] aggregate updated: ${totalScanned} pages merged, ${totalSkipped} carried forward across ${partials.length} sites`,
      );
    } catch (e) {
      context.warn(`[scan ${job.jobId}] aggregate merge failed (non-fatal): ${(e as Error).message}`);
    }
  }
}

app.storageQueue("linkInventoryScanWorker", {
  connection: "AzureWebJobsStorage",
  queueName: SCAN_QUEUE_NAME,
  handler: workerHandler,
});
