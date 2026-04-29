import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, isAdmin, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import {
  addAllowlistEntry,
  getAllowlist,
  removeAllowlistEntry,
  type AllowlistKind,
} from "../services/duplicatesAllowlistStore.js";

/**
 * Admin-managed duplicates allowlist.
 *
 * GET    /api/duplicates/allowlist        — read (any authenticated user;
 *                                           non-admins still need it to
 *                                           understand why some pairs
 *                                           don't show up for them)
 * POST   /api/duplicates/allowlist        — add entry (admin-only)
 * DELETE /api/duplicates/allowlist        — remove entry (admin-only)
 *
 * POST/DELETE body shape:
 *   { kind: 'hash' | 'path' | 'name',
 *     sha256?: string,           // hash kind
 *     pattern?: string,          // path / name kind
 *     note?: string              // POST only
 *   }
 */

function parseKind(raw: unknown): AllowlistKind | undefined {
  if (raw === "hash" || raw === "path" || raw === "name") return raw;
  return undefined;
}

async function handler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
  } catch (err) {
    if (err instanceof AuthError) return { status: err.status, jsonBody: { ok: false, error: err.message } };
    throw err;
  }

  if (request.method === "GET") {
    const allowlist = await getAllowlist();
    let admin = false;
    try {
      admin = await isAdmin(user);
    } catch (e) {
      context.warn(`duplicates allowlist GET: isAdmin failed: ${(e as Error).message}`);
    }
    return { status: 200, jsonBody: { ok: true, isAdmin: admin, allowlist } };
  }

  // Mutations require admin.
  try {
    await requireAdmin(user);
  } catch (err) {
    if (err instanceof AuthError) return { status: err.status, jsonBody: { ok: false, error: err.message } };
    throw err;
  }

  let body: { kind?: string; sha256?: string; pattern?: string; note?: string } = {};
  try {
    const text = await request.text();
    if (text) body = JSON.parse(text) as typeof body;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  const kind = parseKind(body.kind);
  if (!kind) return { status: 400, jsonBody: { ok: false, error: "kind must be hash | path | name" } };

  if (kind === "hash" && !body.sha256) {
    return { status: 400, jsonBody: { ok: false, error: "sha256 is required for hash kind" } };
  }
  if ((kind === "path" || kind === "name") && !body.pattern) {
    return { status: 400, jsonBody: { ok: false, error: "pattern is required for path/name kinds" } };
  }

  const addedBy = user.upn ?? user.userId ?? "(unknown)";

  try {
    if (request.method === "POST") {
      const note = (body.note ?? "").trim();
      if (!note) return { status: 400, jsonBody: { ok: false, error: "note is required" } };
      const next = await addAllowlistEntry({
        kind,
        sha256: body.sha256,
        pattern: body.pattern,
        note,
        addedBy,
      });
      return { status: 200, jsonBody: { ok: true, allowlist: next } };
    }
    if (request.method === "DELETE") {
      const next = await removeAllowlistEntry({
        kind,
        sha256: body.sha256,
        pattern: body.pattern,
      });
      return { status: 200, jsonBody: { ok: true, allowlist: next } };
    }
    return { status: 405, jsonBody: { ok: false, error: "Method not allowed" } };
  } catch (e) {
    context.error(`duplicates allowlist mutation failed: ${(e as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: (e as Error).message } };
  }
}

app.http("duplicatesAllowlist", {
  methods: ["GET", "POST", "DELETE"],
  authLevel: "anonymous",
  route: "duplicates/allowlist",
  handler,
});
