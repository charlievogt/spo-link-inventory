import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { getBootstrapJob } from "../services/duplicatesBootstrapJobStore.js";

/**
 * GET /api/duplicates/bootstrap/{jobId}
 *
 * Admin-only poll endpoint for a queued/running bootstrap job.
 */

async function handler(
  request: HttpRequest,
  _context: InvocationContext,
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
  return { status: 200, jsonBody: { ok: true, job } };
}

app.http("duplicatesBootstrapStatus", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "duplicates/bootstrap/{jobId}",
  handler,
});
