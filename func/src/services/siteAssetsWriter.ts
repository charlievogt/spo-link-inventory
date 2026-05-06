import { SPO_ORIGIN } from "./spoTokenProvider.js";
import { encodeServerRelativePath } from "./spoSitePagesAssetsEnumerator.js";

/**
 * Recycle helper for the orphan-asset cleanup feature.
 *
 * Lives in its own file (separate from `linkInventoryWriter.ts` which
 * does page canvas PATCH + Publish) because the responsibilities are
 * different shapes:
 *   - linkInventoryWriter operates on a SitePages list item (PATCH,
 *     If-Match ETag, MERGE verb, optional Publish).
 *   - siteAssetsWriter operates on a single file via the file's REST
 *     endpoint (POST recycle()).
 * Conflating them would mean carrying ETag/Publish concerns into the
 * recycle path and file-shape concerns into the page-edit path.
 *
 * The recycle call uses the user's OBO write token (delegated
 * AllSites.Write), so:
 *   - SP enforces the user's actual per-list/per-item permissions
 *     (DeleteListItems specifically; the report endpoint's coarse
 *     EditListItems gate is a fast pre-check, not the authoritative
 *     decision).
 *   - The recycle bin records the user's identity as the deleter, so
 *     restoration audit trail is correct.
 *   - No app-level "delete everywhere" capability is exposed — the
 *     blast radius is bounded by what the calling user can already
 *     delete in the SP UI.
 */

/**
 * Outcome of recycling one file.
 *
 *   - "recycled":  SP accepted the recycle and returned 200. The file
 *                  is in the first-stage recycle bin for ~93 days
 *                  total (30d first-stage + 63d second-stage).
 *   - "not-found": SP returned 404. Either the file was already deleted
 *                  before our call, or the URL didn't resolve.
 *   - "forbidden": SP returned 403. The calling user lacks
 *                  DeleteListItems on the item (or its parent list).
 *   - "error":     Any other failure — network, throttling, malformed
 *                  response, etc. The `error` field carries the detail.
 */
export interface RecycleResult {
  serverRelativeUrl: string;
  status: "recycled" | "not-found" | "forbidden" | "error";
  /** Recycle bin item ID (GUID) when SP returned one; useful for restore audit. */
  recycleBinItemId?: string;
  error?: string;
}

interface RecycleOptions {
  /** Override fetch — for tests. Defaults to globalThis.fetch. */
  fetchImpl?: typeof fetch;
}

/**
 * Derive the SP site collection context from a server-relative URL,
 * matching the pattern in `spoFilesEnumerator.fetchFileBytes`.
 *
 *   /sites/<name>/...   → /sites/<name>
 *   /teams/<name>/...   → /teams/<name>
 *   anything else       → ""  (root web)
 *
 * The context matters: `getfilebyserverrelativeurl` MUST be called at
 * the file's own site collection — calling at the root web for a file
 * in a subsite returns 404 even when the URL is otherwise valid.
 */
export function deriveSiteContextFromFile(serverRelativeUrl: string): string {
  const m = /^(\/(?:sites|teams)\/[^/]+)\//i.exec(serverRelativeUrl);
  return m ? m[1] : "";
}

/**
 * Pull the recycle-bin item ID out of an SP REST `/recycle()` success
 * response. SP returns:
 *   nometadata:    { "value": "<guid>" }
 *   minimalmeta:   { "Recycle": "<guid>" }     (rare here — we ask for nometadata)
 *   verbose:       { "d": { "Recycle": "<guid>" } }
 *
 * Tolerates all three. Returns undefined when none of them parse — the
 * caller still considers the recycle successful (the 200 status is
 * authoritative) but loses the ID for audit. Defensive: not worth
 * failing the recycle just because we couldn't pull the ID out.
 */
export function parseRecycleResponse(body: unknown): string | undefined {
  if (!body || typeof body !== "object") return undefined;
  const o = body as Record<string, unknown>;
  if (typeof o.value === "string") return o.value;
  if (typeof o.Recycle === "string") return o.Recycle;
  const d = o.d;
  if (d && typeof d === "object" && typeof (d as Record<string, unknown>).Recycle === "string") {
    return (d as Record<string, unknown>).Recycle as string;
  }
  return undefined;
}

/**
 * Recycle a single file as the calling user.
 *
 * The file is moved to the site's first-stage recycle bin. After 30
 * days it ages into the second-stage recycle bin (admin-only). Total
 * recoverable window is 93 days from the recycle date, configurable
 * per tenant.
 *
 * Always returns a typed `RecycleResult`. Doesn't throw on SP-level
 * errors (404/403/500); does throw on completely-broken inputs (e.g.
 * a malformed token producing a non-Response from fetch).
 */
export async function recycleSitePagesAsset(
  serverRelativeUrl: string,
  userToken: string,
  options: RecycleOptions = {},
): Promise<RecycleResult> {
  const fetchImpl = options.fetchImpl ?? fetch;

  const sitePath = deriveSiteContextFromFile(serverRelativeUrl);
  const encodedRef = encodeServerRelativePath(serverRelativeUrl);
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/getfilebyserverrelativeurl('${encodedRef}')/recycle()`;

  let res: Response;
  try {
    res = await fetchImpl(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${userToken}`,
        Accept: "application/json;odata=nometadata",
      },
    });
  } catch (err) {
    return {
      serverRelativeUrl,
      status: "error",
      error: `fetch threw: ${(err as Error).message}`,
    };
  }

  if (res.status === 200) {
    let recycleBinItemId: string | undefined;
    try {
      const json = (await res.json()) as unknown;
      recycleBinItemId = parseRecycleResponse(json);
    } catch {
      // body unparseable — not fatal, recycle still succeeded
    }
    return { serverRelativeUrl, status: "recycled", recycleBinItemId };
  }

  // Non-200: pull the body for diagnostics.
  let body = "";
  try {
    body = (await res.text()).slice(0, 400);
  } catch {
    // ignore — we still know the status code
  }

  if (res.status === 404) {
    return {
      serverRelativeUrl,
      status: "not-found",
      error: `File not found (already recycled?): ${body}`,
    };
  }
  if (res.status === 403) {
    return {
      serverRelativeUrl,
      status: "forbidden",
      error: `User lacks delete permission on this file: ${body}`,
    };
  }
  return {
    serverRelativeUrl,
    status: "error",
    error: `SP returned ${res.status}: ${body}`,
  };
}
