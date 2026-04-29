import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { getJob, readResults } from "../services/linkInventoryJobStore.js";
import { AuthError, getSitePermissions, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { applyReplacementsToPage, type LinkReplacement, type PageReplaceResult } from "../services/linkInventoryWriter.js";
import type { ISiteInventoryShape } from "./linkInventoryReplace.types.js";
// linkInventoryReplace.types.ts may not include canonicalKey on the
// link shape — declare a local refinement so this file compiles cleanly.
// (The actual aggregate blob does carry canonicalKey since the scanner
// emits it.)

/**
 * Find-and-replace endpoint for the link inventory.
 *
 * POST /api/link-inventory/replace
 *
 * Body:
 *   {
 *     jobId: string,        // a completed scan job id
 *     find: string,         // substring to look for in link.url
 *     replace: string,      // replacement substring
 *     sites?: string[],     // restrict to these site paths (default: all)
 *     dryRun?: boolean,     // default true — preview only
 *     pageIds?: number[]    // optional: restrict to a subset of pages
 *   }
 *
 * Permission model:
 *   - Caller must be a member of the SP Redirect Manager Admins group
 *     (this is a destructive operation; admin-only)
 *   - For each affected page, we ALSO check the user's per-site
 *     EditListItems via OBO before touching that site. The user must
 *     have edit on the site or its results are silently dropped.
 *
 * Concurrency:
 *   - Per-page ETag check via the linkInventoryWriter service
 *   - Stale pages return status: "stale" with no write attempted
 *
 * Response:
 *   {
 *     ok: true,
 *     dryRun: bool,
 *     summary: { totalPages, applied, preview, stale, conflict, noMatch, error },
 *     results: PageReplaceResult[]
 *   }
 */

interface ReplaceRequest {
  jobId: string;
  find: string;
  replace: string;
  sites?: string[];
  dryRun?: boolean;
  pageIds?: number[];
  /**
   * Match mode:
   *   - "substring" (default): the previous behavior — find URLs whose
   *     `link.url` contains `find`, replace that substring with `replace`.
   *   - "canonical": find URLs whose `link.canonicalKey === find`, replace
   *     each one's full url with `replace`. This is the "align all forms
   *     of a sharing link to the canonical AllItems URL" path. Each
   *     match becomes a per-link replacement of the entire URL, regardless
   *     of how the form is encoded.
   */
  mode?: "substring" | "canonical";
}

interface AggregateResults {
  jobId: string;
  startedAt: string;
  finishedAt: string;
  sites: ISiteInventoryShape[];
}

async function replaceHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // Auth: admin-only
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

  // Parse body
  let body: ReplaceRequest;
  try {
    const text = await request.text();
    if (!text) return { status: 400, jsonBody: { ok: false, error: "Empty body" } };
    body = JSON.parse(text) as ReplaceRequest;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  if (!body.jobId || !body.find) {
    return { status: 400, jsonBody: { ok: false, error: "jobId and find are required" } };
  }
  if (typeof body.find !== "string" || body.find.length === 0) {
    return { status: 400, jsonBody: { ok: false, error: "find must be a non-empty string" } };
  }
  if (typeof body.replace !== "string") {
    return { status: 400, jsonBody: { ok: false, error: "replace must be a string (use empty string to delete)" } };
  }
  const dryRun = body.dryRun !== false; // default true
  const mode = body.mode === "canonical" ? "canonical" : "substring";

  // Load job + results
  const job = await getJob(body.jobId);
  if (!job) return { status: 404, jsonBody: { ok: false, error: `Job ${body.jobId} not found` } };
  if (!job.resultsAvailable) {
    return { status: 409, jsonBody: { ok: false, error: `Job ${body.jobId} has no results yet` } };
  }
  const aggregate = (await readResults(body.jobId)) as AggregateResults | undefined;
  if (!aggregate) return { status: 404, jsonBody: { ok: false, error: "Results blob missing" } };

  // Build per-page replacement plans:
  //  - Walk every site/page/link in scan results
  //  - Filter to sites in body.sites (if provided)
  //  - Filter to pageIds in body.pageIds (if provided)
  //  - For each link whose `url` contains `find`, compute the new url
  //    by substring substitution and add a LinkReplacement
  // Then for each (site, page) pair, run applyReplacementsToPage.

  const sitesFilter = body.sites && body.sites.length > 0 ? new Set(body.sites) : undefined;
  const pageIdsFilter = body.pageIds && body.pageIds.length > 0 ? new Set(body.pageIds) : undefined;

  // Group replacements by site → page id → list. Two match modes:
  //
  //   - substring: traditional find/replace — match URLs containing
  //     `find`, do a string substitution to produce the new URL.
  //   - canonical: match URLs whose `canonicalKey` equals `find`, and
  //     replace each one's full URL with `replace`. This collapses all
  //     forms (sharing wrappers, direct paths, AllItems URLs) of a
  //     given target file to the same canonical replacement.
  const plansBySite = new Map<string, Map<number, LinkReplacement[]>>();
  for (const site of aggregate.sites) {
    if (sitesFilter && !sitesFilter.has(site.site)) continue;
    if (site.error) continue;
    for (const page of site.pages) {
      if (pageIdsFilter && !pageIdsFilter.has(page.pageId)) continue;
      const matches: LinkReplacement[] = [];
      for (const link of page.links) {
        let newUrl: string;
        if (mode === "canonical") {
          if (link.canonicalKey !== body.find) continue;
          newUrl = body.replace;
          if (newUrl === link.url) continue;
        } else {
          if (link.url.indexOf(body.find) === -1) continue;
          newUrl = link.url.split(body.find).join(body.replace);
          if (newUrl === link.url) continue;
        }
        matches.push({
          pageId: page.pageId,
          scanEtag: page.etag,
          rawUrl: link.rawUrl,
          oldUrl: link.url,
          newUrl,
        });
      }
      if (matches.length === 0) continue;
      let pageMap = plansBySite.get(site.site);
      if (!pageMap) {
        pageMap = new Map<number, LinkReplacement[]>();
        plansBySite.set(site.site, pageMap);
      }
      pageMap.set(page.pageId, matches);
    }
  }

  if (plansBySite.size === 0) {
    return {
      status: 200,
      jsonBody: {
        ok: true,
        dryRun,
        summary: { totalPages: 0, applied: 0, preview: 0, stale: 0, conflict: 0, noMatch: 0, error: 0 },
        results: [],
      },
    };
  }

  // Per-site permission check (user EditListItems via OBO). Sites the
  // user can't write to are dropped from the plan and recorded.
  const droppedSites: string[] = [];
  const allowedSites: string[] = [];
  for (const sitePath of plansBySite.keys()) {
    let perms: { canRead: boolean; canWrite: boolean };
    try {
      // Pass context so getSitePermissions logs the actual SP response —
      // this is the most common source of "I have access but the tool
      // says I don't" reports, and we need the trace to diagnose.
      // Use the WRITE scope: SP masks write bits out of the response
      // when the OBO token has only read scope, so a read-purpose
      // check returns canWrite=false even when the user actually
      // has EditListItems.
      perms = await getSitePermissions(user, sitePath, context, 'write');
    } catch (err) {
      context.error(`replace: permission check failed for ${sitePath}: ${(err as Error).message}`);
      droppedSites.push(sitePath);
      continue;
    }
    if (!perms.canWrite) {
      context.warn(`replace: ${user.upn ?? user.userId} skipped on ${sitePath} — canWrite=false`);
      droppedSites.push(sitePath);
    } else {
      allowedSites.push(sitePath);
    }
  }

  // Execute the plan, page-by-page, sequentially per site (to stay
  // under SP throttling on writes). Pass the user principal so SP REST
  // calls run as the user via delegated OBO — the audit trail reflects
  // the actual change author and SP enforces the user's permissions.
  const results: Array<PageReplaceResult & { site: string }> = [];
  for (const sitePath of allowedSites) {
    const pageMap = plansBySite.get(sitePath)!;
    for (const [pageId, replacements] of pageMap) {
      try {
        const r = await applyReplacementsToPage(sitePath, replacements, { dryRun, user });
        results.push({ ...r, site: sitePath });
      } catch (err) {
        results.push({
          site: sitePath,
          pageId,
          status: "error",
          replacementCount: 0,
          unmatchedCount: replacements.length,
          details: replacements.map((r) => ({ oldUrl: r.oldUrl, newUrl: r.newUrl, matched: false })),
          error: (err as Error).message,
        });
      }
    }
  }

  const summary = {
    totalPages: results.length,
    applied: results.filter((r) => r.status === "applied").length,
    preview: results.filter((r) => r.status === "preview").length,
    stale: results.filter((r) => r.status === "stale").length,
    conflict: results.filter((r) => r.status === "conflict").length,
    noMatch: results.filter((r) => r.status === "no-match").length,
    error: results.filter((r) => r.status === "error" || r.status === "not-found").length,
  };
  context.log(
    `[replace ${body.jobId}] ${user.upn ?? user.userId} ${dryRun ? "DRY-RUN" : "APPLY"}: ${JSON.stringify(summary)}`,
  );

  return {
    status: 200,
    jsonBody: {
      ok: true,
      dryRun,
      summary,
      droppedSites: droppedSites.length > 0 ? droppedSites : undefined,
      results,
    },
  };
}

app.http("linkInventoryReplace", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/replace",
  handler: replaceHandler,
});
