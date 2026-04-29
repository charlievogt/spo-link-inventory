import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { getBootstrapJob, updateBootstrapJob } from "../services/duplicatesBootstrapJobStore.js";

/**
 * POST /api/duplicates/bootstrap/{jobId}/cancel
 *
 * Marks the bootstrap job as failed so the worker skips any remaining
 * queue messages on next dequeue. Admin-only. Idempotent — running on
 * an already-completed/failed job returns 200 with the current row.
 *
 * Note: queued messages aren't drained from the queue here; they just
 * dequeue, see the status, and exit without enqueueing the next file.
 */

async function handler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
    await requireAdmin(user);
  } catch (err) {
    if (err instanceof AuthError) return { status: err.status, jsonBody: { ok: false, error: err.message } };
    throw err;
  }

  const jobId = request.params.jobId;
  if (!jobId) return { status: 400, jsonBody: { ok: false, error: "Missing jobId route parameter" } };

  const job = await getBootstrapJob(jobId);
  if (!job) return { status: 404, jsonBody: { ok: false, error: "Job not found" } };

  if (job.status === "completed" || job.status === "failed") {
    return { status: 200, jsonBody: { ok: true, job, info: `Already ${job.status}` } };
  }

  job.status = "failed";
  job.finishedAt = new Date().toISOString();
  job.currentFile = undefined;
  job.recentErrors = [`cancelled by ${user.upn ?? user.userId ?? "(unknown)"}`, ...job.recentErrors].slice(0, 10);
  await updateBootstrapJob(job);

  context.log(`[bootstrap ${jobId}] cancelled by ${user.upn ?? user.userId}`);

  return { status: 200, jsonBody: { ok: true, job } };
}

app.http("duplicatesBootstrapCancel", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "duplicates/bootstrap/{jobId}/cancel",
  handler,
});
