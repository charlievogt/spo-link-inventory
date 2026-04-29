import { randomUUID } from "node:crypto";
import { createJob, getJob, updateJob, listJobs } from "./linkInventoryJobStore.js";
import { enumerateSites } from "./spoSiteEnumerator.js";
import { enqueueScanMessage } from "./linkInventoryQueue.js";
import { enqueueDocsScanMessage } from "./linkInventoryDocsQueue.js";
import { MAX_FILE_BYTES_DEFAULT } from "./spoFilesEnumerator.js";

/**
 * Shared trigger logic for kicking off a unified (page + doc) scan.
 *
 * Used by:
 *   - `linkInventoryScanUnified` — the admin-triggered HTTP endpoint
 *   - `linkInventoryScheduleTimer` — the daily UI-configured timer
 *
 * Same job shape, same cross-linked sibling pairing, same enqueue
 * pattern — only the inputs (caller, mode, modifiedAfter) differ.
 *
 * The runner is conservative about overlapping scans: when called by
 * the timer, it skips entirely if any scheduled job is still running.
 * The HTTP endpoint can opt out of this guard since admins triggering
 * manually have explicitly chosen to run alongside whatever's in flight.
 */

export interface UnifiedScanRequest {
  /** Sites to scan. If undefined, the runner enumerates the tenant. */
  sites?: string[];
  /** Identifier for `caller` field — UPN for manual, "scheduled" for timer. */
  caller: string;
  /** Optional ISO date — doc scan filters out files modified before this. */
  modifiedAfter?: string;
  /** Doc scan size cap. Defaults to 100 MB. */
  maxFileBytes?: number;
  /** Page scans run their cross-link verifier when true. */
  verifyFiles?: boolean;
  /**
   * When true, skip if any other "scheduled"-caller job is queued/running.
   * The timer always sets this; the HTTP endpoint never does.
   */
  skipIfScheduledActive?: boolean;
}

export interface UnifiedScanResult {
  ok: true;
  pageJobId: string;
  docJobId: string;
  sitesTotal: number;
  enumerated: boolean;
}

export interface UnifiedScanSkipped {
  ok: false;
  skipped: true;
  reason: string;
}

export interface UnifiedScanFailure {
  ok: false;
  skipped?: false;
  error: string;
}

export type UnifiedScanOutcome = UnifiedScanResult | UnifiedScanSkipped | UnifiedScanFailure;

/**
 * Detect overlap with previously-triggered scheduled jobs. We check both
 * page and doc jobs because a stuck enumerate-phase doc scan would still
 * be doing work even if no `pages` job is alive.
 */
async function hasActiveScheduledJob(): Promise<{ active: boolean; jobIds: string[] }> {
  const jobs = await listJobs(200);
  const active = jobs.filter(
    (j) => j.caller === "scheduled" && (j.status === "queued" || j.status === "running"),
  );
  return { active: active.length > 0, jobIds: active.map((j) => j.jobId) };
}

export async function triggerUnifiedScan(req: UnifiedScanRequest): Promise<UnifiedScanOutcome> {
  if (req.skipIfScheduledActive) {
    const { active, jobIds } = await hasActiveScheduledJob();
    if (active) {
      return {
        ok: false,
        skipped: true,
        reason: `${jobIds.length} scheduled job(s) still running: ${jobIds.join(", ")}`,
      };
    }
  }

  let sites = req.sites;
  let enumerated = false;
  if (!sites || sites.length === 0) {
    try {
      const enumeratedSites = await enumerateSites();
      sites = enumeratedSites.map((s) => s.serverRelativeUrl);
      enumerated = true;
    } catch (err) {
      return { ok: false, error: `Site enumeration failed: ${(err as Error).message}` };
    }
  }
  if (sites.length === 0) {
    return { ok: false, error: "No sites to scan" };
  }

  const pageJobId = randomUUID();
  const docJobId = randomUUID();
  const maxFileBytes = req.maxFileBytes ?? MAX_FILE_BYTES_DEFAULT;

  await createJob(pageJobId, sites, req.caller, "pages", undefined, undefined, req.verifyFiles, docJobId);
  await createJob(
    docJobId,
    sites,
    req.caller,
    "documents",
    0,
    {
      maxFileBytes,
      modifiedAfter: req.modifiedAfter,
    },
    req.verifyFiles,
    pageJobId,
  );

  try {
    await Promise.all([
      enqueueScanMessage({ jobId: pageJobId, siteIndex: 0 }),
      enqueueDocsScanMessage({ jobId: docJobId, phase: "enumerate", siteIndex: 0 }),
    ]);
  } catch (err) {
    for (const jid of [pageJobId, docJobId]) {
      try {
        const j = await getJob(jid);
        if (j) {
          j.status = "failed";
          j.recentErrors = [
            `unified scan enqueue failed: ${(err as Error).message}`,
            ...j.recentErrors,
          ].slice(0, 10);
          await updateJob(j);
        }
      } catch {
        // best-effort cleanup
      }
    }
    return { ok: false, error: `Enqueue failed: ${(err as Error).message}` };
  }

  return {
    ok: true,
    pageJobId,
    docJobId,
    sitesTotal: sites.length,
    enumerated,
  };
}
