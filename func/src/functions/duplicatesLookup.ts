import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, filterReadableSites, isAdmin, parseUserPrincipal } from "../services/linkInventoryAuth.js";
import { getInMemoryHashIndex } from "../services/hashIndexStore.js";
import { getAllowlist } from "../services/duplicatesAllowlistStore.js";
import { signalsForFile } from "../services/duplicatesQuery.js";

/**
 * Per-file duplicate-signals lookup.
 *
 * GET /api/duplicates/lookup?fileRef=<serverRelativeUrl>
 *
 * Auth: OBO bearer token required. Admin bypasses ACL; non-admins only
 * see counterpart files from sites they can Read. Counterparts on
 * forbidden sites collapse into a `hiddenCount` per bucket.
 */

async function lookupHandler(
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

  const fileRef = request.query.get("fileRef");
  if (!fileRef) return { status: 400, jsonBody: { ok: false, error: "Missing fileRef query parameter" } };

  let index;
  try {
    index = await getInMemoryHashIndex();
  } catch (e) {
    context.error(`duplicates lookup: index load failed: ${(e as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: "Hash index unavailable" } };
  }
  if (!index) {
    return {
      status: 200,
      jsonBody: {
        fileRef,
        known: false,
        isAdmin: false,
        exact: [],
        stale: [],
        sameName: [],
        hiddenCount: { exact: 0, stale: 0, sameName: 0 },
      },
    };
  }

  const allowlist = await getAllowlist();
  const signals = signalsForFile(index, fileRef, allowlist);
  if (!signals) {
    return {
      status: 200,
      jsonBody: {
        fileRef,
        known: false,
        isAdmin: false,
        exact: [],
        stale: [],
        sameName: [],
        hiddenCount: { exact: 0, stale: 0, sameName: 0 },
      },
    };
  }

  let admin = false;
  try {
    admin = await isAdmin(user);
  } catch (e) {
    context.warn(`duplicates lookup: isAdmin failed: ${(e as Error).message}`);
  }

  let visibleSet: Set<string> | null = null;
  if (!admin) {
    const sites = new Set<string>();
    for (const e of signals.exact) sites.add(e.sitePath);
    for (const s of signals.stale) sites.add(s.authoritativeSitePath);
    for (const n of signals.sameName) sites.add(n.sitePath);
    if (sites.size === 0) {
      visibleSet = new Set();
    } else {
      const allowed = await filterReadableSites(user, [...sites]);
      visibleSet = new Set(allowed.map((s) => s.toLowerCase()));
    }
  }

  const canSee = (sitePath: string): boolean => admin || !!visibleSet?.has(sitePath.toLowerCase());
  const visibleExact = signals.exact.filter((e) => canSee(e.sitePath));
  const visibleStale = signals.stale.filter((s) => canSee(s.authoritativeSitePath));
  const visibleSameName = signals.sameName.filter((n) => canSee(n.sitePath));

  return {
    status: 200,
    jsonBody: {
      fileRef: signals.fileRef,
      sha256: signals.sha256,
      known: true,
      isAdmin: admin,
      indexBuiltAt: index.builtAt,
      exact: visibleExact,
      stale: visibleStale,
      sameName: visibleSameName,
      hiddenCount: {
        exact: signals.exact.length - visibleExact.length,
        stale: signals.stale.length - visibleStale.length,
        sameName: signals.sameName.length - visibleSameName.length,
      },
    },
  };
}

app.http("duplicatesLookup", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "duplicates/lookup",
  handler: lookupHandler,
});
