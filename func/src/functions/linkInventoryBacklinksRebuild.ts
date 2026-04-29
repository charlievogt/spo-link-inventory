import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { listJobs, readResults } from "../services/linkInventoryJobStore.js";
import { buildPersistentFromAggregates, writeFullIndex } from "../services/backlinksIndexStore.js";
import type { ISiteInventoryShape } from "./linkInventoryReplace.types.js";

/**
 * Admin-only: rebuild the persistent backlinks index from the most
 * recent completed page scan + most recent completed document scan.
 *
 * POST /api/link-inventory/backlinks-index/rebuild
 *
 * Use cases:
 *   - Bootstrap on first deploy (no scan-finalize has run yet)
 *   - Recovery after the persistent blob is manually deleted or gets
 *     out of sync with the latest scans
 *   - Forcing a clean re-aggregation after upgrading the index shape
 */

interface AggregateBlob {
  sites?: unknown[];
  startedAt?: string;
  finishedAt?: string;
  jobId?: string;
}

async function rebuildHandler(
  _request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(_request.headers.get("authorization"));
    await requireAdmin(user);
  } catch (err) {
    if (err instanceof AuthError) return { status: err.status, jsonBody: { ok: false, error: err.message } };
    throw err;
  }

  // Find the most recent completed job for each kind.
  const all = await listJobs(200);
  const latestPages = all.find((j) => (j.kind ?? "pages") === "pages" && j.status === "completed" && j.resultsAvailable);
  const latestDocs = all.find((j) => j.kind === "documents" && j.status === "completed" && j.resultsAvailable);

  if (!latestPages && !latestDocs) {
    return { status: 404, jsonBody: { ok: false, error: "No completed scans to rebuild from" } };
  }

  let pageAgg: ISiteInventoryShape[] | null = null;
  let pageScannedAt: string | null = null;
  if (latestPages) {
    const blob = (await readResults(latestPages.jobId)) as AggregateBlob | undefined;
    if (blob?.sites) {
      pageAgg = blob.sites as ISiteInventoryShape[];
      pageScannedAt = blob.finishedAt ?? latestPages.finishedAt ?? latestPages.startedAt;
    }
  }

  let docAgg: ISiteInventoryShape[] | null = null;
  let docScannedAt: string | null = null;
  if (latestDocs) {
    const blob = (await readResults(latestDocs.jobId)) as AggregateBlob | undefined;
    if (blob?.sites) {
      docAgg = blob.sites as ISiteInventoryShape[];
      docScannedAt = blob.finishedAt ?? latestDocs.finishedAt ?? latestDocs.startedAt;
    }
  }

  const persistent = buildPersistentFromAggregates(
    pageAgg,
    latestPages?.jobId ?? null,
    pageScannedAt,
    docAgg,
    latestDocs?.jobId ?? null,
    docScannedAt,
  );

  try {
    await writeFullIndex(persistent);
  } catch (e) {
    context.error(`backlinks rebuild: write failed: ${(e as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: "Failed to write index" } };
  }

  const siteCount = Object.keys(persistent.bySite).length;
  const pageSiteCount = Object.values(persistent.bySite).filter((s) => s.pages).length;
  const docSiteCount = Object.values(persistent.bySite).filter((s) => s.documents).length;

  return {
    status: 200,
    jsonBody: {
      ok: true,
      builtAt: persistent.builtAt,
      sitesTotal: siteCount,
      sitesWithPageData: pageSiteCount,
      sitesWithDocData: docSiteCount,
      sourcePageJobId: latestPages?.jobId,
      sourceDocJobId: latestDocs?.jobId,
    },
  };
}

app.http("linkInventoryBacklinksRebuild", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/backlinks-index/rebuild",
  handler: rebuildHandler,
});
