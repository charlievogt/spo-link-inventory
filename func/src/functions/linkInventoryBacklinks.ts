import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, filterReadableSites, isAdmin, parseUserPrincipal } from "../services/linkInventoryAuth.js";
import { getInMemoryIndex } from "../services/backlinksIndexStore.js";
import { applyAclGate, lookupBacklinks, type BacklinkEntry } from "../services/backlinksIndex.js";

/**
 * Per-file backlinks lookup for the field customizer.
 *
 * GET /api/link-inventory/backlinks?fileRef=<serverRelativeUrl>
 *
 * Auth: OBO bearer token required. Admin membership is checked to
 * decide whether to bypass the per-site ACL filter. Non-admins only see
 * backlinks from sites they can Read; the rest collapse to a count.
 */

interface BacklinksResponse {
  fileRef: string;
  /** When the persistent index was last updated (top-level). */
  indexBuiltAt?: string;
  /** True when the caller is a Redirect Manager Admin. */
  isAdmin: boolean;
  /** Source pages/docs the caller can see. */
  visible: Array<{
    sourceKind: "page" | "document";
    site: string;
    siteUrl: string;
    title: string;
    url: string;
    scannedAt: string;
  }>;
  /** Count of backlinks hidden by the ACL gate. */
  hiddenCount: number;
}

function projectEntry(e: BacklinkEntry): BacklinksResponse["visible"][number] {
  return {
    sourceKind: e.sourceKind,
    site: e.sourceSite,
    siteUrl: e.sourceSiteUrl,
    title: e.sourceTitle,
    url: e.sourceUrl,
    scannedAt: e.sourceUpdatedAt,
  };
}

async function backlinksHandler(
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
  if (!fileRef) {
    return { status: 400, jsonBody: { ok: false, error: "Missing fileRef query parameter" } };
  }

  // Load persistent index (cached 5 min). If there's no index blob yet,
  // the feature hasn't been bootstrapped — return an empty payload so
  // the customizer renders "no backlinks" rather than erroring.
  let index;
  try {
    index = await getInMemoryIndex();
  } catch (e) {
    context.error(`backlinks: index load failed: ${(e as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: "Backlinks index unavailable" } };
  }
  if (!index) {
    return {
      status: 200,
      jsonBody: {
        fileRef,
        isAdmin: false,
        visible: [],
        hiddenCount: 0,
      } satisfies BacklinksResponse,
    };
  }

  const all = lookupBacklinks(index, fileRef);

  // Admin bypass the ACL gate. Otherwise compute the visible-site set
  // via the same SP search ACL trim used by the scan results endpoint.
  let admin = false;
  try {
    admin = await isAdmin(user);
  } catch (e) {
    context.warn(`backlinks: isAdmin failed, treating as non-admin: ${(e as Error).message}`);
  }

  let visibleSet: Set<string> | null = null;
  if (!admin) {
    const sourceSites = [...new Set(all.map((e) => e.sourceSitePath))];
    if (sourceSites.length === 0) {
      visibleSet = new Set();
    } else {
      const allowed = await filterReadableSites(user, sourceSites);
      visibleSet = new Set(allowed.map((s) => s.toLowerCase()));
    }
  }

  const { visible, hiddenCount } = applyAclGate(all, visibleSet);

  const payload: BacklinksResponse = {
    fileRef,
    indexBuiltAt: index.scannedAt,
    isAdmin: admin,
    visible: visible.map(projectEntry),
    hiddenCount,
  };

  return { status: 200, jsonBody: payload };
}

app.http("linkInventoryBacklinks", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "link-inventory/backlinks",
  handler: backlinksHandler,
});
