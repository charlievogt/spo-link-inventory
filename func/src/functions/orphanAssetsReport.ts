import { app, type HttpRequest, type HttpResponseInit, type InvocationContext } from "@azure/functions";
import {
  AuthError,
  filterReadableSites,
  parseUserPrincipal,
  requireAdmin,
} from "../services/linkInventoryAuth.js";
import { getInMemoryIndex } from "../services/backlinksIndexStore.js";
import {
  enumerateSitePagesAssets,
  type SitePagesAssetFile,
} from "../services/spoSitePagesAssetsEnumerator.js";
import { findOrphans, type OrphanFile } from "../services/orphanQuery.js";
import type { BacklinksIndex } from "../services/backlinksIndex.js";

/**
 * Orphan-asset report endpoint.
 *
 * POST /api/orphan-assets/report
 *
 * Body:
 *   {
 *     sites?: string[],                  // optional restriction; default: every site in the index
 *     maxConcurrency?: number,           // 1..16, default 6
 *     acknowledgedStaleIndex?: boolean,  // override the > BLOCK-age refusal
 *   }
 *
 * Response:
 *   {
 *     ok: true,
 *     generatedAt: ISO,
 *     indexAge: { scannedAt, ageMinutes, staleWarn },
 *     warning?: string,                  // user-facing note for stale-but-acknowledged runs
 *     sites: OrphansReportSite[],
 *     totals: { sitesScanned, filesScanned, orphans },
 *   }
 *
 * Behavior:
 *   - Admin-only (LINK_INVENTORY_ADMIN_GROUP_ID gate).
 *   - Filters target sites to those the calling user can read via SP search
 *     ACL trim, mirroring the existing inventory read paths.
 *   - Refuses (409) when the persistent backlinks index is missing, or when
 *     it's older than ORPHAN_INDEX_BLOCK_AGE_MIN unless the body sets
 *     acknowledgedStaleIndex=true. Warns (in the response body, 200) when
 *     older than ORPHAN_INDEX_WARN_AGE_MIN.
 *
 * Performance:
 *   - One blob read for the index (cached after first call within a Function
 *     instance), one HTTP call per requested site for the SitePages walk
 *     (recursive into subfolders inside that one call). Total runtime is
 *     dominated by the SitePages walks; capped via maxConcurrency.
 */

interface ReportRequest {
  sites?: string[];
  maxConcurrency?: number;
  acknowledgedStaleIndex?: boolean;
}

interface OrphansReportSite {
  sitePath: string;
  siteUrl: string;
  filesScanned: number;
  orphans: OrphanFile[];
  error?: string;
}

const ORPHAN_INDEX_WARN_AGE_MIN = parseInt(
  process.env.ORPHAN_INDEX_WARN_AGE_MIN ?? "60",
  10,
);
const ORPHAN_INDEX_BLOCK_AGE_MIN = parseInt(
  process.env.ORPHAN_INDEX_BLOCK_AGE_MIN ?? "360",
  10,
);

const DEFAULT_CONCURRENCY = 6;
const MAX_CONCURRENCY = 16;

/**
 * Walk every backlink entry in the inflated index and collect the unique
 * set of `sourceSitePath` values for `page`-kind sources. That set is the
 * "sites we have page-scan coverage for" — orphan detection is only
 * reliable for these.
 *
 * Edge case: a site that's been page-scanned but contributed zero outbound
 * links won't appear here (the index has no entries pointing back to it).
 * Such sites are dropped from this v1 report — at typical content density
 * the case is rare. A future revision could read the persistent index's
 * `bySite` keys directly to cover that gap, at the cost of one extra blob
 * read per request.
 */
function pageScannedSites(index: BacklinksIndex): Map<string, string> {
  const out = new Map<string, string>();
  for (const entries of index.byCanonicalKey.values()) {
    for (const e of entries) {
      if (e.sourceKind !== "page") continue;
      if (!out.has(e.sourceSitePath)) {
        out.set(e.sourceSitePath, e.sourceSiteUrl);
      }
    }
  }
  return out;
}

async function reportHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // 1. Auth
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
    await requireAdmin(user);
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  // 2. Body
  let body: ReportRequest = {};
  const bodyText = await request.text();
  if (bodyText) {
    try {
      body = JSON.parse(bodyText) as ReportRequest;
    } catch {
      return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
    }
  }

  const concurrency = Math.min(
    Math.max(1, body.maxConcurrency ?? DEFAULT_CONCURRENCY),
    MAX_CONCURRENCY,
  );

  // 3. Load index
  const index = await getInMemoryIndex();
  if (!index) {
    return {
      status: 409,
      jsonBody: {
        ok: false,
        error:
          "No persistent backlinks index — run a tenant scan first. " +
          "Orphan detection requires the page-scan index to know what's referenced.",
      },
    };
  }

  // 4. Staleness gate
  const scannedAt = new Date(index.scannedAt);
  const nowMs = Date.now();
  const ageMs = nowMs - scannedAt.getTime();
  const ageMinutes = Math.max(0, Math.floor(ageMs / 60_000));
  const staleWarn = ageMinutes > ORPHAN_INDEX_WARN_AGE_MIN;
  const staleBlock = ageMinutes > ORPHAN_INDEX_BLOCK_AGE_MIN;

  if (staleBlock && !body.acknowledgedStaleIndex) {
    return {
      status: 409,
      jsonBody: {
        ok: false,
        error:
          `Backlinks index is ${ageMinutes} min old, exceeding the ` +
          `${ORPHAN_INDEX_BLOCK_AGE_MIN} min freshness limit. Run a fresh ` +
          `tenant scan, or pass "acknowledgedStaleIndex": true to override.`,
        indexAge: { scannedAt: index.scannedAt, ageMinutes, staleWarn: true, staleBlock: true },
      },
    };
  }

  // 5. Determine target sites
  const coveredSiteMap = pageScannedSites(index);
  const allCoveredSites = [...coveredSiteMap.keys()];

  let requestedSites: string[];
  if (body.sites && body.sites.length > 0) {
    const requestedSet = new Set(body.sites.map((s) => s.toLowerCase()));
    requestedSites = allCoveredSites.filter((s) => requestedSet.has(s.toLowerCase()));
  } else {
    requestedSites = allCoveredSites;
  }

  if (requestedSites.length === 0) {
    return {
      status: 200,
      jsonBody: {
        ok: true,
        generatedAt: new Date().toISOString(),
        indexAge: { scannedAt: index.scannedAt, ageMinutes, staleWarn },
        sites: [],
        totals: { sitesScanned: 0, filesScanned: 0, orphans: 0 },
      },
    };
  }

  // 6. Filter to user-readable sites (SP search ACL trim, mirrors the existing inventory paths)
  const readableSites = await filterReadableSites(user, requestedSites);

  // 7. Walk SitePages for each site, find orphans (worker pool)
  // Bind a non-null local so the closure-captured reference is narrowed.
  const indexRef: BacklinksIndex = index;
  const reportSites: OrphansReportSite[] = [];
  let cursor = 0;
  async function worker(): Promise<void> {
    while (true) {
      const idx = cursor++;
      if (idx >= readableSites.length) return;
      const sitePath = readableSites[idx];
      const siteUrl = coveredSiteMap.get(sitePath) ?? "";
      try {
        const files: SitePagesAssetFile[] = await enumerateSitePagesAssets(sitePath);
        const orphans = findOrphans(files, indexRef);
        reportSites.push({ sitePath, siteUrl, filesScanned: files.length, orphans });
      } catch (err) {
        const msg = (err as Error).message;
        context.warn(`[orphan-report] ${sitePath}: ${msg}`);
        reportSites.push({ sitePath, siteUrl, filesScanned: 0, orphans: [], error: msg });
      }
    }
  }
  const workerCount = Math.min(concurrency, readableSites.length);
  const workers: Promise<void>[] = [];
  for (let i = 0; i < workerCount; i++) workers.push(worker());
  await Promise.all(workers);

  // 8. Totals + response
  const totals = {
    sitesScanned: reportSites.length,
    filesScanned: reportSites.reduce((s, r) => s + r.filesScanned, 0),
    orphans: reportSites.reduce((s, r) => s + r.orphans.length, 0),
  };

  context.log(
    `[orphan-report] ${user.upn ?? user.userId} ` +
      `${totals.sitesScanned} sites, ${totals.filesScanned} files, ${totals.orphans} orphans`,
  );

  const response: Record<string, unknown> = {
    ok: true,
    generatedAt: new Date().toISOString(),
    indexAge: { scannedAt: index.scannedAt, ageMinutes, staleWarn },
    sites: reportSites,
    totals,
  };
  if (staleBlock && body.acknowledgedStaleIndex) {
    response.warning =
      `Backlinks index is ${ageMinutes} min old (threshold: ${ORPHAN_INDEX_BLOCK_AGE_MIN}). ` +
      `Override accepted; results may flag actively-referenced files added since the last scan.`;
  } else if (staleWarn) {
    response.warning =
      `Backlinks index is ${ageMinutes} min old (warn threshold: ${ORPHAN_INDEX_WARN_AGE_MIN}). ` +
      `Pages added or edited since then are not yet reflected.`;
  }

  return { status: 200, jsonBody: response };
}

app.http("orphanAssetsReport", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "orphan-assets/report",
  handler: reportHandler,
});
