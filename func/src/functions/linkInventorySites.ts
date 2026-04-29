import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, parseUserPrincipal } from "../services/linkInventoryAuth.js";
import { enumerateSites, deriveHubs, type SiteSummary, type HubSiteSummary } from "../services/spoSiteEnumerator.js";

/**
 * Site list endpoint for the SPFx site picker.
 *
 * GET /api/link-inventory/sites
 *
 * Returns the full enumerated tenant site list (with hub associations)
 * plus the derived hub list. The SPFx UI uses this to render a
 * checklist grouped by hub, replacing the free-text site list textarea
 * we had on the doc-scan dialog.
 *
 * Auth: any authenticated user — same bar as listJobs/getResults. Site
 * metadata (titles, paths, hub IDs) is not sensitive.
 *
 * Caching: results are cached in module-local state for 5 minutes
 * since the underlying SP search call is the slowest part of the page
 * load and the site list rarely changes between page opens.
 */

interface CachedResult {
  expiresAt: number;
  sites: SiteSummary[];
  hubs: HubSiteSummary[];
}

const CACHE_TTL_MS = 5 * 60 * 1000;
let cache: CachedResult | undefined;

async function sitesHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // Auth gate: must be authenticated, no admin requirement.
  try {
    parseUserPrincipal(request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  const now = Date.now();
  if (cache && cache.expiresAt > now) {
    return {
      status: 200,
      jsonBody: {
        ok: true,
        cached: true,
        sites: cache.sites,
        hubs: cache.hubs,
      },
    };
  }

  try {
    const sites = await enumerateSites();
    const hubs = deriveHubs(sites);
    cache = { expiresAt: now + CACHE_TTL_MS, sites, hubs };
    context.log(`[sites] enumerated ${sites.length} sites, ${hubs.length} hubs`);
    return {
      status: 200,
      jsonBody: {
        ok: true,
        cached: false,
        sites,
        hubs,
      },
    };
  } catch (err) {
    context.error(`[sites] enumeration failed: ${(err as Error).message}`);
    return {
      status: 500,
      jsonBody: { ok: false, error: `Site enumeration failed: ${(err as Error).message}` },
    };
  }
}

app.http("linkInventorySites", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "link-inventory/sites",
  handler: sitesHandler,
});
