import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, isAdmin, parseUserPrincipal } from "../services/linkInventoryAuth.js";

/**
 * Identity / role probe for the SPFx web part.
 *
 * GET /api/link-inventory/whoami
 *
 * Returns the calling user's principal info plus whether they're a
 * member of the configured admin group (LINK_INVENTORY_ADMIN_GROUP_ID).
 * The web part uses this to hide the "Run scan" button for non-admins
 * (the backend still enforces — this is just UX).
 */

async function whoamiHandler(
  request: HttpRequest,
  _context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof AuthError) {
      return { status: err.status, jsonBody: { ok: false, error: err.message } };
    }
    throw err;
  }

  let admin = false;
  try {
    admin = await isAdmin(user);
  } catch {
    admin = false;
  }

  return {
    status: 200,
    jsonBody: {
      ok: true,
      userId: user.userId,
      upn: user.upn,
      isAdmin: admin,
    },
  };
}

app.http("linkInventoryWhoami", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "link-inventory/whoami",
  handler: whoamiHandler,
});
