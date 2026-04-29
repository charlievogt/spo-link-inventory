import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { getJob, purgeJob } from "../services/linkInventoryJobStore.js";

/**
 * Manual job purge endpoint.
 *
 * DELETE /api/link-inventory/scan/{jobId}
 *
 * Removes a single scan job and every artifact it produced — the
 * Table row, the aggregate result blob, all per-site partials, the
 * file manifest (for doc scans), and the manifest fragments. Used by
 * admins to clean up failed runs, test scans, and superseded results
 * without waiting for the 30-day retention sweep.
 *
 * Permission model: admin-only. The retention timer uses the same
 * underlying `purgeJob` helper but bypasses auth (it runs as the MI).
 */

async function deleteScanHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // Admin gate
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

  const jobId = request.params.jobId;
  if (!jobId) {
    return { status: 400, jsonBody: { ok: false, error: "Missing jobId in path" } };
  }

  // Confirm the job exists before purging — gives the caller a clean
  // 404 instead of a silent no-op if they typo the id.
  const job = await getJob(jobId);
  if (!job) {
    return { status: 404, jsonBody: { ok: false, error: `Job not found: ${jobId}` } };
  }

  try {
    const result = await purgeJob(jobId);
    context.log(
      `[scan delete] ${jobId} purged by ${user.upn ?? user.userId}: ` +
      `rowDeleted=${result.rowDeleted}, blobsDeleted=${result.blobsDeleted}`,
    );
    return {
      status: 200,
      jsonBody: {
        ok: true,
        jobId,
        rowDeleted: result.rowDeleted,
        blobsDeleted: result.blobsDeleted,
      },
    };
  } catch (err) {
    context.error(`[scan delete] ${jobId} failed: ${(err as Error).message}`);
    return {
      status: 500,
      jsonBody: { ok: false, error: `Purge failed: ${(err as Error).message}` },
    };
  }
}

app.http("linkInventoryScanDelete", {
  methods: ["DELETE"],
  authLevel: "anonymous",
  route: "link-inventory/scan/{jobId}",
  handler: deleteScanHandler,
});
