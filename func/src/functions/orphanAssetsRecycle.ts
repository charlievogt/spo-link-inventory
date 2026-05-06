import { app, type HttpRequest, type HttpResponseInit, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import {
  AuthError,
  getSitePermissions,
  getSpoUserWriteToken,
  parseUserPrincipal,
  requireAdmin,
} from "../services/linkInventoryAuth.js";
import {
  createJob,
  getJob,
  updateJob,
  writeResults,
} from "../services/linkInventoryJobStore.js";
import {
  recycleSitePagesAsset,
  type RecycleResult,
} from "../services/siteAssetsWriter.js";

/**
 * Recycle endpoint for orphan-asset cleanup.
 *
 * POST /api/orphan-assets/recycle
 *
 * Body:
 *   {
 *     files: Array<{ sitePath: string; serverRelativeUrl: string }>,
 *     dryRun?: boolean,                // default true
 *     confirmReportGeneratedAt?: string,  // optional staleness guard
 *   }
 *
 * Response:
 *   {
 *     ok: true,
 *     dryRun: boolean,
 *     runId: string,                  // synthetic job id; queryable via existing job APIs
 *     summary: { totalFiles, recycled, preview, notFound, forbidden, error, droppedSites: string[] },
 *     results: Array<{ sitePath, serverRelativeUrl, status, recycleBinItemId?, error? }>,
 *   }
 *
 * Behavior:
 *   - Admin-only (LINK_INVENTORY_ADMIN_GROUP_ID gate).
 *   - Per-site OBO write check via `getSitePermissions(... 'write')`. Sites where the
 *     calling user lacks EditListItems are dropped from the plan and recorded in
 *     `summary.droppedSites`. SP enforces actual DeleteListItems at the file level
 *     during the recycle call (returning 403 → result.status: 'forbidden').
 *   - All recycle calls execute as the user via delegated AllSites.Write OBO. The
 *     site recycle bin records the user's identity as the deleter.
 *   - Audit: every run creates a synthetic 'orphan-recycle' job header in
 *     LinkInventoryJobs (counters reused — sitesTotal=unique sites,
 *     linksTotal=files-recycled, errorCount=failures). The full per-file results
 *     are persisted to the same `<jobId>.json` blob the scan jobs use, so the
 *     existing list/show/delete UI works without modification, and the existing
 *     30-day retention timer reaps these audit rows alongside scan rows.
 *
 * Dry-run path:
 *   - No SP REST calls. No job header is written. The response shape mirrors
 *     a real run with status='preview' on every row, so the UI can display the
 *     same table either way.
 */

interface RecycleRequest {
  files?: Array<{ sitePath?: string; serverRelativeUrl?: string }>;
  dryRun?: boolean;
  confirmReportGeneratedAt?: string;
}

interface PerFileResult {
  sitePath: string;
  serverRelativeUrl: string;
  status: "recycled" | "preview" | "not-found" | "forbidden" | "error" | "site-not-writable";
  recycleBinItemId?: string;
  error?: string;
}

async function recycleHandler(
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
  let body: RecycleRequest = {};
  const bodyText = await request.text();
  if (!bodyText) {
    return { status: 400, jsonBody: { ok: false, error: "Empty body" } };
  }
  try {
    body = JSON.parse(bodyText) as RecycleRequest;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  if (!Array.isArray(body.files) || body.files.length === 0) {
    return { status: 400, jsonBody: { ok: false, error: "files (non-empty array) is required" } };
  }

  // Validate each file row up front; reject the whole request if any are malformed
  // (rather than partially executing some and confusing the caller).
  const fileRows: Array<{ sitePath: string; serverRelativeUrl: string }> = [];
  for (const f of body.files) {
    if (!f || typeof f.sitePath !== "string" || typeof f.serverRelativeUrl !== "string") {
      return {
        status: 400,
        jsonBody: { ok: false, error: "every entry in files must have sitePath and serverRelativeUrl strings" },
      };
    }
    fileRows.push({ sitePath: f.sitePath, serverRelativeUrl: f.serverRelativeUrl });
  }

  const dryRun = body.dryRun !== false; // default true

  // Group files by site for the per-site write check + per-site execution loop.
  const bySite = new Map<string, string[]>();
  for (const r of fileRows) {
    if (!bySite.has(r.sitePath)) bySite.set(r.sitePath, []);
    bySite.get(r.sitePath)!.push(r.serverRelativeUrl);
  }

  const runId = randomUUID();
  const startedAt = new Date().toISOString();

  // 3. Dry-run short-circuit — no SP calls, no job header, just emit preview rows.
  if (dryRun) {
    const results: PerFileResult[] = fileRows.map((r) => ({
      sitePath: r.sitePath,
      serverRelativeUrl: r.serverRelativeUrl,
      status: "preview",
    }));
    const summary = {
      totalFiles: results.length,
      recycled: 0,
      preview: results.length,
      notFound: 0,
      forbidden: 0,
      error: 0,
      droppedSites: [] as string[],
    };
    return {
      status: 200,
      jsonBody: {
        ok: true,
        dryRun: true,
        runId,
        summary,
        results,
      },
    };
  }

  // 4. Real run — admin-only gate already passed; now per-site write check.
  const droppedSites: string[] = [];
  const allowedSites: string[] = [];
  for (const sitePath of bySite.keys()) {
    let perms;
    try {
      perms = await getSitePermissions(user, sitePath, context, "write");
    } catch (err) {
      context.warn(`[recycle] permission check failed for ${sitePath}: ${(err as Error).message}`);
      droppedSites.push(sitePath);
      continue;
    }
    if (!perms.canWrite) {
      context.warn(
        `[recycle] ${user.upn ?? user.userId} lacks write on ${sitePath} — dropping ${bySite.get(sitePath)?.length ?? 0} file(s)`,
      );
      droppedSites.push(sitePath);
    } else {
      allowedSites.push(sitePath);
    }
  }

  // 5. Acquire user OBO write token (single token, used for every recycle call).
  let userToken: string;
  try {
    userToken = await getSpoUserWriteToken(user);
  } catch (err) {
    return {
      status: 401,
      jsonBody: {
        ok: false,
        error: `Failed to get user OBO write token: ${(err as Error).message}`,
      },
    };
  }

  // 6. Create the synthetic job header NOW so the run is visible in the job UI
  // even if it crashes mid-flight. status flips to 'completed'/'failed' at the
  // end; until then it shows 'running'.
  await createJob(
    runId,
    allowedSites,
    user.upn ?? user.userId,
    "orphan-recycle",
  );
  // createJob set status='queued'; flip to 'running' immediately.
  const job = await getJob(runId);
  if (job) {
    job.status = "running";
    await updateJob(job);
  }

  // 7. Execute. Sequential per site (to stay under SP write throttling), parallel
  // within a site? No — files within a site are also serial for v1 to keep
  // throttling boring and predictable. Total runtime scales linearly with file
  // count; admins typically recycle dozens to a few hundred files at a time, so
  // this is fine.
  const results: PerFileResult[] = [];
  let recycledCount = 0;
  let errorCount = 0;
  for (const sitePath of allowedSites) {
    const urls = bySite.get(sitePath) ?? [];
    for (const url of urls) {
      const r: RecycleResult = await recycleSitePagesAsset(url, userToken);
      const out: PerFileResult = { sitePath, ...r };
      results.push(out);
      if (r.status === "recycled") {
        recycledCount += 1;
      } else {
        errorCount += 1;
      }
    }
  }

  // Add explicit rows for files in dropped sites so the response makes the loss visible.
  for (const sitePath of droppedSites) {
    for (const url of bySite.get(sitePath) ?? []) {
      results.push({
        sitePath,
        serverRelativeUrl: url,
        status: "site-not-writable",
        error: `User lacks write permission on ${sitePath}; file skipped.`,
      });
    }
  }

  const summary = {
    totalFiles: results.length,
    recycled: recycledCount,
    preview: 0,
    notFound: results.filter((r) => r.status === "not-found").length,
    forbidden: results.filter((r) => r.status === "forbidden").length,
    error: results.filter((r) => r.status === "error" || r.status === "site-not-writable").length,
    droppedSites,
  };

  // 8. Persist audit detail to the same blob shape the scan jobs use
  const auditDetail = {
    runId,
    runAt: startedAt,
    finishedAt: new Date().toISOString(),
    runBy: user.upn ?? user.userId,
    runByOid: user.userId,
    dryRun: false,
    confirmReportGeneratedAt: body.confirmReportGeneratedAt,
    summary,
    results,
  };
  try {
    await writeResults(runId, auditDetail);
  } catch (err) {
    context.warn(
      `[recycle ${runId}] audit blob write failed (non-fatal): ${(err as Error).message}`,
    );
  }

  // 9. Update the job header with final state.
  if (job) {
    job.status = errorCount > 0 && recycledCount === 0 ? "failed" : "completed";
    job.finishedAt = auditDetail.finishedAt;
    // Reuse linksTotal as "files recycled" — counter rendering in the existing
    // UI will show this as a recognizable count even without orphan-specific labels.
    job.linksTotal = recycledCount;
    job.errorCount = errorCount;
    job.recentErrors = results
      .filter((r) => r.status === "error" || r.status === "forbidden")
      .slice(0, 10)
      .map((r) => `${r.serverRelativeUrl}: ${r.error ?? r.status}`);
    job.resultsAvailable = true;
    job.sitesCompleted = allowedSites.length;
    try {
      await updateJob(job);
    } catch (err) {
      context.warn(
        `[recycle ${runId}] job header update failed (non-fatal): ${(err as Error).message}`,
      );
    }
  }

  context.log(
    `[recycle ${runId}] ${user.upn ?? user.userId} ${dryRun ? "DRY-RUN" : "APPLY"}: ` +
      `recycled=${recycledCount} error=${errorCount} droppedSites=${droppedSites.length}`,
  );

  return {
    status: 200,
    jsonBody: {
      ok: true,
      dryRun: false,
      runId,
      summary,
      results,
    },
  };
}

app.http("orphanAssetsRecycle", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "orphan-assets/recycle",
  handler: recycleHandler,
});
