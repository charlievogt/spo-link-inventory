import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { getJob, listJobs, readResults, type LinkInventoryJob } from "../services/linkInventoryJobStore.js";
import { AuthError, filterReadableSites, parseUserPrincipal } from "../services/linkInventoryAuth.js";
import type { SiteInventory } from "../services/linkInventoryScanner.js";

/**
 * Status + results endpoints for the tenant-wide scan orchestrator.
 *
 *  GET /api/link-inventory/scan                       → list recent jobs
 *  GET /api/link-inventory/scan/{jobId}               → header (counters, status)
 *  GET /api/link-inventory/scan/{jobId}/results       → per-page inventory, filtered to sites the user can read
 *
 * Permission model:
 *   - Job headers (counters, status, sites scanned) are visible to any
 *     authenticated user. They contain no content data — just metadata —
 *     so this is a deliberate transparency choice ("is the inventory
 *     fresh?").
 *   - The /results endpoint filters the per-site inventory down to the
 *     sites the calling user has at least ViewListItems on, via OBO
 *     to SP REST effectiveBasePermissions. Sites the user can't read
 *     are silently dropped from the response (the totals are
 *     recomputed from what's left).
 */

interface AggregateResults {
  jobId: string;
  startedAt: string;
  finishedAt: string;
  sites: SiteInventory[];
  totals: { sites: number; pages: number; links: number; errors: number };
}

async function statusHandler(
  request: HttpRequest,
  _context: InvocationContext,
): Promise<HttpResponseInit> {
  // Job header listing/reading is auth-required but not admin-only.
  // Anyone with a valid bearer token can see what scans have happened.
  try {
    parseUserPrincipal(request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  const jobId = request.params.jobId;
  if (!jobId) {
    const jobs = await listJobs(50);
    return { status: 200, jsonBody: { ok: true, jobs } };
  }

  const job = await getJob(jobId);
  if (!job) {
    return { status: 404, jsonBody: { ok: false, error: `Job not found: ${jobId}` } };
  }
  return { status: 200, jsonBody: { ok: true, job } };
}

/**
 * Recompute aggregate totals from a filtered site list. Used after we
 * drop sites the user can't read so the response totals reflect what
 * the user actually sees, not what was originally scanned.
 */
function recomputeTotals(sites: SiteInventory[]): AggregateResults["totals"] {
  return {
    sites: sites.length,
    pages: sites.reduce((s, r) => s + r.pageCount, 0),
    links: sites.reduce((s, r) => s + r.linkCount, 0),
    errors: sites.filter((r) => r.error).length,
  };
}

async function resultsHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  const jobId = request.params.jobId;
  if (!jobId) return { status: 400, jsonBody: { ok: false, error: "Missing jobId" } };

  const job: LinkInventoryJob | undefined = await getJob(jobId);
  if (!job) return { status: 404, jsonBody: { ok: false, error: `Job not found: ${jobId}` } };
  if (!job.resultsAvailable) {
    return {
      status: 409,
      jsonBody: { ok: false, error: `Job is ${job.status}; results not available yet` },
    };
  }

  // Orphan-recycle audit jobs share the table + writeResults convention with
  // scan jobs but use a different blob shape (no `sites` array; per-file
  // `results`). Refuse them here so callers don't trip the .sites.map below.
  // The audit history surfaces under the Orphans tab via its own endpoint.
  if (job.kind === "orphan-recycle") {
    return {
      status: 400,
      jsonBody: {
        ok: false,
        error:
          `Job ${jobId} is an orphan-recycle audit, not a scan. ` +
          `Audit history lives under the Orphans tab.`,
      },
    };
  }

  const raw = (await readResults(jobId)) as AggregateResults | undefined;
  if (!raw) return { status: 404, jsonBody: { ok: false, error: "Results blob missing" } };

  // Preview-only doc scans produce a different blob shape
  // (`documents-preview` with `siteSummaries` instead of `sites`).
  // These don't need per-site permission filtering — they only
  // contain file counts, not actual link content. Return as-is.
  if ((raw as { kind?: string }).kind === "documents-preview") {
    context.log(`[results ${jobId}] ${user.upn ?? user.userId}: preview blob, returning as-is`);
    return { status: 200, jsonBody: { ok: true, jobId, results: raw } };
  }

  // Filter site results to sites the user can read.
  let readable: string[];
  try {
    readable = await filterReadableSites(user, raw.sites.map((s) => s.site));
  } catch (err) {
    if (err instanceof AuthError) {
      context.error(`results filter failed for ${user.upn}: ${err.message}`);
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  const allowedSet = new Set(readable);
  const filteredSites = raw.sites.filter((s) => allowedSet.has(s.site));
  const droppedSites = raw.sites.filter((s) => !allowedSet.has(s.site)).map((s) => s.site);

  const filtered: AggregateResults & { droppedSites: string[]; userScopedTotals: AggregateResults["totals"] } = {
    ...raw,
    sites: filteredSites,
    droppedSites,
    userScopedTotals: recomputeTotals(filteredSites),
  };

  context.log(
    `[results ${jobId}] ${user.upn ?? user.userId}: ${filteredSites.length}/${raw.sites.length} sites visible`,
  );

  return { status: 200, jsonBody: { ok: true, jobId, results: filtered } };
}

app.http("linkInventoryScanStatus", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "link-inventory/scan/{jobId?}",
  handler: statusHandler,
});

app.http("linkInventoryScanResults", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "link-inventory/scan/{jobId}/results",
  handler: resultsHandler,
});
