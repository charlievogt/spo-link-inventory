import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { scanSite } from "../services/linkInventoryScanner.js";
import { AuthError, getSitePermissions, parseUserPrincipal } from "../services/linkInventoryAuth.js";

/**
 * Single-site link inventory scan endpoint.
 *
 * GET /api/link-inventory/scan-site?site=/sites/<sitename>[&compact=1]
 *
 * Synchronous: returns when the scan finishes. For tenant-wide scans
 * use the orchestrator at /api/link-inventory/scan instead, which
 * persists progress and returns a job id immediately.
 *
 * Requires Sites.Selected `Read` granted on the target site.
 */

async function scanSiteHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // Per-site read enforcement: caller must have ViewListItems on the
  // requested site. This endpoint runs an ad-hoc scan against live data
  // (not the cached job results), so we enforce as the user before
  // touching SP.
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  const sitePathRaw = request.query.get("site");
  if (!sitePathRaw) {
    return {
      status: 400,
      jsonBody: { ok: false, error: "Missing required query param: site (e.g. /sites/charlie-test-site)" },
    };
  }
  const sitePath = sitePathRaw.startsWith("/") ? sitePathRaw : `/${sitePathRaw}`;
  const compact = request.query.get("compact") === "1";

  try {
    const perms = await getSitePermissions(user, sitePath);
    if (!perms.canRead) {
      return { status: 403, jsonBody: { ok: false, error: `User ${user.upn} cannot read ${sitePath}` } };
    }
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  context.log(`linkInventoryScanSite: starting scan of ${sitePath} for ${user.upn ?? user.userId}`);
  const result = await scanSite(sitePath);
  context.log(
    `linkInventoryScanSite: ${sitePath} → ${result.pageCount} pages, ${result.linkCount} links, ${result.scanMs}ms${result.error ? ` (ERROR: ${result.error})` : ""}`,
  );

  return {
    status: result.error ? 500 : 200,
    jsonBody: {
      ok: !result.error,
      ...result,
      pages: compact ? [] : result.pages,
    },
  };
}

app.http("linkInventoryScanSite", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "link-inventory/scan-site",
  handler: scanSiteHandler,
});
