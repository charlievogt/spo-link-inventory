import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import {
  AuthError,
  parseUserPrincipal,
  requireAdmin,
} from "../services/linkInventoryAuth.js";
import {
  getScheduleConfig,
  updateScheduleConfig,
  ScheduleConfigValidationError,
  type ScheduleConfigInput,
} from "../services/scheduleConfigStore.js";

/**
 * Read/write the daily-scan schedule config.
 *
 *   GET  /api/link-inventory/schedule  → current config (admin-only)
 *   PUT  /api/link-inventory/schedule  → update (admin-only)
 *
 * GET is admin-gated even though the data isn't sensitive — the schedule
 * UI lives entirely behind the admin panel, no value in opening it up.
 *
 * The request body for PUT is partial: any field not present is left
 * unchanged. validation is delegated to the store layer.
 */

async function scheduleHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
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

  if (request.method === "GET") {
    try {
      const config = await getScheduleConfig();
      return { status: 200, jsonBody: { ok: true, config } };
    } catch (err) {
      context.error(`schedule GET failed: ${(err as Error).message}`);
      return { status: 500, jsonBody: { ok: false, error: (err as Error).message } };
    }
  }

  if (request.method === "PUT") {
    let body: ScheduleConfigInput;
    try {
      body = (await request.json()) as ScheduleConfigInput;
      if (!body || typeof body !== "object") {
        return { status: 400, jsonBody: { ok: false, error: "Request body must be a JSON object" } };
      }
    } catch {
      return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
    }
    try {
      const config = await updateScheduleConfig(body, user.upn ?? user.userId ?? "unknown");
      return { status: 200, jsonBody: { ok: true, config } };
    } catch (err) {
      if (err instanceof ScheduleConfigValidationError) {
        return { status: 400, jsonBody: { ok: false, error: err.message } };
      }
      context.error(`schedule PUT failed: ${(err as Error).message}`);
      return { status: 500, jsonBody: { ok: false, error: (err as Error).message } };
    }
  }

  return { status: 405, jsonBody: { ok: false, error: `Method ${request.method} not allowed` } };
}

app.http("linkInventorySchedule", {
  methods: ["GET", "PUT"],
  authLevel: "anonymous",
  route: "link-inventory/schedule",
  handler: scheduleHandler,
});
