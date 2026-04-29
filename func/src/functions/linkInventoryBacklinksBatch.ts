import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, filterReadableSites, isAdmin, parseUserPrincipal } from "../services/linkInventoryAuth.js";
import { getInMemoryIndex } from "../services/backlinksIndexStore.js";
import { applyAclGate, lookupBacklinks, type BacklinkEntry } from "../services/backlinksIndex.js";

/**
 * Batch backlinks lookup — one request, many fileRefs.
 *
 * POST /api/link-inventory/backlinks/batch
 * Body: { fileRefs: string[] }
 *
 * Same auth + ACL model as the per-file endpoint, but the admin check
 * and visible-site set are computed ONCE per request instead of N
 * times. The persistent index is loaded once (with ETag-validated
 * caching) and every fileRef is looked up against the same in-memory
 * map.
 *
 * This exists because a library view with 50 rows was firing 50
 * separate /backlinks requests, each paying its own index-load cost on
 * cold starts. Batching collapses that to one request.
 *
 * Limits: `fileRefs` is capped at 200 per call to protect the function
 * from pathological payloads. SP library views rarely render more than
 * that at once.
 */

const MAX_BATCH_SIZE = 200;

interface BatchBody {
  fileRefs?: string[];
}

interface BatchResult {
  visible: Array<{
    sourceKind: "page" | "document";
    site: string;
    siteUrl: string;
    title: string;
    url: string;
    scannedAt: string;
  }>;
  hiddenCount: number;
}

interface BatchResponse {
  indexBuiltAt?: string;
  isAdmin: boolean;
  results: Record<string, BatchResult>;
}

function projectEntry(e: BacklinkEntry): BatchResult["visible"][number] {
  return {
    sourceKind: e.sourceKind,
    site: e.sourceSite,
    siteUrl: e.sourceSiteUrl,
    title: e.sourceTitle,
    url: e.sourceUrl,
    scannedAt: e.sourceUpdatedAt,
  };
}

async function batchHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof AuthError) return { status: err.status, jsonBody: { ok: false, error: err.message } };
    throw err;
  }

  let body: BatchBody;
  try {
    body = (await request.json()) as BatchBody;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  const fileRefs = Array.isArray(body.fileRefs) ? body.fileRefs.filter((r) => typeof r === "string" && r.length > 0) : [];
  if (fileRefs.length === 0) {
    return { status: 400, jsonBody: { ok: false, error: "fileRefs array is required" } };
  }
  if (fileRefs.length > MAX_BATCH_SIZE) {
    return { status: 400, jsonBody: { ok: false, error: `fileRefs capped at ${MAX_BATCH_SIZE} per call` } };
  }

  let index;
  try {
    index = await getInMemoryIndex();
  } catch (e) {
    context.error(`backlinks/batch: index load failed: ${(e as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: "Backlinks index unavailable" } };
  }

  // No index yet — return empty results for every ref so the customizer
  // renders "no backlinks" rather than erroring.
  if (!index) {
    const empty: BatchResponse = { isAdmin: false, results: {} };
    for (const ref of fileRefs) empty.results[ref] = { visible: [], hiddenCount: 0 };
    return { status: 200, jsonBody: empty };
  }

  // Admin bypass the ACL gate. Otherwise compute the visible-site set
  // ONCE, not per fileRef.
  let admin = false;
  try {
    admin = await isAdmin(user);
  } catch (e) {
    context.warn(`backlinks/batch: isAdmin failed, treating as non-admin: ${(e as Error).message}`);
  }

  // Collect ALL backlinks across the batch first, so we can compute one
  // visible-site set covering every source site that appears anywhere
  // in the results. Then gate each ref against that set.
  const rawByRef = new Map<string, BacklinkEntry[]>();
  const allSourceSites = new Set<string>();
  for (const ref of fileRefs) {
    const entries = lookupBacklinks(index, ref);
    rawByRef.set(ref, entries);
    for (const e of entries) allSourceSites.add(e.sourceSitePath);
  }

  let visibleSet: Set<string> | null = null;
  if (!admin && allSourceSites.size > 0) {
    const allowed = await filterReadableSites(user, [...allSourceSites]);
    visibleSet = new Set(allowed.map((s) => s.toLowerCase()));
  } else if (!admin) {
    visibleSet = new Set();
  }

  const results: Record<string, BatchResult> = {};
  for (const ref of fileRefs) {
    const entries = rawByRef.get(ref) ?? [];
    const { visible, hiddenCount } = applyAclGate(entries, visibleSet);
    results[ref] = {
      visible: visible.map(projectEntry),
      hiddenCount,
    };
  }

  const payload: BatchResponse = {
    indexBuiltAt: index.scannedAt,
    isAdmin: admin,
    results,
  };
  return { status: 200, jsonBody: payload };
}

app.http("linkInventoryBacklinksBatch", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/backlinks/batch",
  handler: batchHandler,
});
