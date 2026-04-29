import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import { createJob } from "../services/linkInventoryJobStore.js";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { enumerateSites } from "../services/spoSiteEnumerator.js";
import { enqueueScanMessage } from "../services/linkInventoryQueue.js";

/**
 * Tenant-wide link-inventory scan orchestrator (queue-triggered).
 *
 * POST /api/link-inventory/scan
 *   body (optional):
 *     {
 *       sites?: string[],         // explicit list, overrides everything
 *       enumerateAll?: boolean    // force enumeration even if app setting is set
 *     }
 *   - When neither `sites` nor `LINK_INVENTORY_SCAN_SITES` is set,
 *     enumerates the tenant via SP REST search.
 *
 * Response:
 *   {
 *     ok: true,
 *     jobId,
 *     status: "queued",
 *     sitesTotal,
 *     sites,
 *     enumerated  // true when SP REST enumeration was used
 *   }
 *
 * Behavior:
 *   - Creates a job row in the LinkInventoryJobs table immediately
 *   - Enqueues a single Storage Queue message `{jobId, siteIndex: 0}`
 *   - Returns the jobId; the queue worker takes over and processes
 *     one site per message, chaining to the next message after each
 *   - No sites-per-scan cap — the worker processes durably across
 *     Function host restarts and isn't bound by the 230s HTTP timeout
 */

const SITE_PATH_RE = /^\/sites\/[a-z0-9-]+$|^\/$/i;

interface ScanRequestBody {
  sites?: string[];
  /** When `true`, enumerate the tenant via SP REST search and scan
   *  every site. Defaults to enumerating when neither `sites` nor the
   *  LINK_INVENTORY_SCAN_SITES app setting is set. */
  enumerateAll?: boolean;
  /**
   * Targeted page rescan: a map from site path to page ids to scan
   * within that site. When set, only the listed pages are scanned in
   * each site (rather than the whole library). The site list is
   * derived from the keys of this map; explicit `sites` is ignored.
   * Use `Refresh affected pages` after find-and-replace previews
   * detect stale ETags.
   */
  pageIdsBySite?: Record<string, number[]>;
  /**
   * Opt-in: after scanning completes, HEAD-check every AllItems.aspx
   * `?id=` path against SP REST. 404s get upgraded from `spo-internal`
   * to `malformed-spo-link` (reason `file-not-found`). Adds one REST
   * call per unique AllItems link, so it's behind a checkbox in the
   * admin UI.
   */
  verifyFiles?: boolean;
}

function getDefaultSites(): string[] {
  const raw = process.env.LINK_INVENTORY_SCAN_SITES ?? "";
  return raw
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

function validateSites(sites: string[]): { valid: string[]; invalid: string[] } {
  const valid: string[] = [];
  const invalid: string[] = [];
  for (const s of sites) {
    const path = s.startsWith("/") ? s : `/${s}`;
    if (SITE_PATH_RE.test(path)) valid.push(path);
    else invalid.push(s);
  }
  return { valid, invalid };
}

async function scanHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // Triggering a tenant scan is admin-only — gated by Entra group membership.
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

  let body: ScanRequestBody = {};
  try {
    const text = await request.text();
    if (text) body = JSON.parse(text) as ScanRequestBody;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  // Site selection precedence:
  //   1. Explicit `pageIdsBySite` (targeted rescan) — sites are the keys
  //   2. Explicit `sites` in the request body
  //   3. `LINK_INVENTORY_SCAN_SITES` app setting (operational override
  //      for testing — comma-separated)
  //   4. Tenant-wide enumeration via SP REST search
  let requested: string[] = [];
  let enumerated = false;
  if (body.pageIdsBySite && Object.keys(body.pageIdsBySite).length > 0) {
    requested = Object.keys(body.pageIdsBySite);
  } else if (body.sites && body.sites.length > 0) {
    requested = body.sites;
  } else if (!body.enumerateAll) {
    requested = getDefaultSites();
  }
  if (requested.length === 0) {
    try {
      context.log("[scan] no explicit sites provided — enumerating tenant via SP REST search");
      const sites = await enumerateSites();
      requested = sites.map((s) => s.serverRelativeUrl);
      enumerated = true;
      context.log(`[scan] enumerator found ${requested.length} sites`);
    } catch (err) {
      return {
        status: 500,
        jsonBody: {
          ok: false,
          error: `Site enumeration failed: ${(err as Error).message}`,
        },
      };
    }
  }
  if (requested.length === 0) {
    return {
      status: 400,
      jsonBody: { ok: false, error: "No sites to scan (enumeration returned 0)." },
    };
  }

  const { valid, invalid } = validateSites(requested);
  if (valid.length === 0) {
    return {
      status: 400,
      jsonBody: { ok: false, error: "No valid site paths in request", invalid },
    };
  }

  // No more sites-per-scan cap — the queue worker processes one site
  // per message and chains forward, so the only limit is wall-clock
  // patience and SP throttling.
  const jobId = randomUUID();
  await createJob(jobId, valid, user.upn ?? user.userId, "pages", undefined, undefined, body.verifyFiles);
  // Persist the targeted page-id filter as a side blob so the worker
  // can read it on each message without inflating the job row.
  if (body.pageIdsBySite) {
    const { writeFileManifest } = await import("../services/linkInventoryJobStore.js");
    await writeFileManifest(jobId, { pageIdsBySite: body.pageIdsBySite });
  }
  context.log(`[scan ${jobId}] enqueuing first message for ${valid.length} sites${body.pageIdsBySite ? ' (targeted)' : ''}`);

  // Enqueue the first message. The worker will chain to siteIndex+1
  // after each site, eventually finalizing the job. If enqueue itself
  // fails (rare), bubble back as a 500 — caller can retry.
  try {
    await enqueueScanMessage({ jobId, siteIndex: 0 });
  } catch (err) {
    context.error(`[scan ${jobId}] enqueue failed: ${(err as Error).message}`);
    return {
      status: 500,
      jsonBody: { ok: false, error: `Failed to enqueue scan: ${(err as Error).message}` },
    };
  }

  return {
    status: 202,
    jsonBody: {
      ok: true,
      jobId,
      status: "queued",
      sitesTotal: valid.length,
      sites: valid,
      invalid: invalid.length > 0 ? invalid : undefined,
      enumerated,
    },
  };
}

app.http("linkInventoryScan", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/scan",
  handler: scanHandler,
});
