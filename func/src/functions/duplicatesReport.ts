import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, filterReadableSites, isAdmin, parseUserPrincipal } from "../services/linkInventoryAuth.js";
import { getInMemoryHashIndex } from "../services/hashIndexStore.js";
import { getAllowlist } from "../services/duplicatesAllowlistStore.js";
import { buildReport, filterReportByVisibility } from "../services/duplicatesQuery.js";

/**
 * Tenant-wide duplicates report. Non-admins see an ACL-filtered slice.
 *
 * GET /api/duplicates/report
 *
 * Query params:
 *   maxExactGroups       — cap exact-dup groups returned (default 200)
 *   maxStalePairs        — cap stale pairs returned (default 500)
 *   maxSameNamePairs     — cap same-name pairs returned (default 500)
 *   maxNearDuplicatePairs — cap near-duplicate pairs returned (default 500)
 *   maxDivergedPairs     — cap diverged pairs returned (default 500)
 *
 * Counts in the response reflect the full tenant report BEFORE the
 * caps were applied so the UI can tell the user "showing 200 of 1873".
 */

const DEFAULT_MAX_EXACT = 200;
const DEFAULT_MAX_STALE = 500;
const DEFAULT_MAX_SAMENAME = 500;
const DEFAULT_MAX_NEARDUP = 500;
const DEFAULT_MAX_DIVERGED = 500;

function readCap(value: string | null, fallback: number): number {
  if (!value) return fallback;
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n) || n <= 0) return fallback;
  return Math.min(n, 5000);
}

async function reportHandler(
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

  let index;
  try {
    index = await getInMemoryHashIndex();
  } catch (e) {
    context.error(`duplicates report: index load failed: ${(e as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: "Hash index unavailable" } };
  }
  if (!index) {
    return {
      status: 200,
      jsonBody: {
        indexBuiltAt: null,
        isAdmin: false,
        totals: {
          exactGroups: 0,
          staleFiles: 0,
          sameNamePairs: 0,
          nearDuplicatePairs: 0,
          divergedPairs: 0,
        },
        exactGroups: [],
        stalePairs: [],
        sameNamePairs: [],
        nearDuplicatePairs: [],
        divergedPairs: [],
      },
    };
  }

  const allowlist = await getAllowlist();
  const full = buildReport(index, allowlist);

  let admin = false;
  try {
    admin = await isAdmin(user);
  } catch (e) {
    context.warn(`duplicates report: isAdmin failed: ${(e as Error).message}`);
  }

  let scoped = full;
  if (!admin) {
    const sites = new Set<string>();
    for (const g of full.exactGroups) for (const f of g.files) sites.add(f.sitePath);
    for (const p of full.stalePairs) {
      sites.add(p.staleSitePath);
      sites.add(p.authoritativeSitePath);
    }
    for (const p of full.sameNamePairs) {
      sites.add(p.aSitePath);
      sites.add(p.bSitePath);
    }
    for (const p of full.nearDuplicatePairs) {
      sites.add(p.aSitePath);
      sites.add(p.bSitePath);
    }
    for (const p of full.divergedPairs) {
      sites.add(p.aSitePath);
      sites.add(p.bSitePath);
    }
    const allowed = sites.size > 0 ? await filterReadableSites(user, [...sites]) : [];
    scoped = filterReportByVisibility(full, new Set(allowed.map((s) => s.toLowerCase())));
  }

  const maxExact = readCap(request.query.get("maxExactGroups"), DEFAULT_MAX_EXACT);
  const maxStale = readCap(request.query.get("maxStalePairs"), DEFAULT_MAX_STALE);
  const maxSameName = readCap(request.query.get("maxSameNamePairs"), DEFAULT_MAX_SAMENAME);
  const maxNearDup = readCap(request.query.get("maxNearDuplicatePairs"), DEFAULT_MAX_NEARDUP);
  const maxDiverged = readCap(request.query.get("maxDivergedPairs"), DEFAULT_MAX_DIVERGED);

  return {
    status: 200,
    jsonBody: {
      indexBuiltAt: index.builtAt,
      isAdmin: admin,
      totals: scoped.totals,
      exactGroups: scoped.exactGroups.slice(0, maxExact),
      stalePairs: scoped.stalePairs.slice(0, maxStale),
      sameNamePairs: scoped.sameNamePairs.slice(0, maxSameName),
      nearDuplicatePairs: scoped.nearDuplicatePairs.slice(0, maxNearDup),
      divergedPairs: scoped.divergedPairs.slice(0, maxDiverged),
      truncated: {
        exactGroups: scoped.exactGroups.length > maxExact,
        stalePairs: scoped.stalePairs.length > maxStale,
        sameNamePairs: scoped.sameNamePairs.length > maxSameName,
        nearDuplicatePairs: scoped.nearDuplicatePairs.length > maxNearDup,
        divergedPairs: scoped.divergedPairs.length > maxDiverged,
      },
    },
  };
}

app.http("duplicatesReport", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "duplicates/report",
  handler: reportHandler,
});
