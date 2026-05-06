import { getSpoToken, SPO_ORIGIN } from "./spoTokenProvider.js";

/**
 * Thin SP REST client for the Site Pages library. Pages over `nextLink`
 * automatically and returns the union of items.
 *
 * Why no @pnp/sp dependency: bundle size, node-vs-browser awkwardness in
 * Functions, and we only need three calls. Hand-rolling keeps the
 * Function lean and the auth integration is just `Authorization: Bearer`.
 */

export interface SitePageItem {
  Id: number;
  Title?: string;
  FileLeafRef?: string;
  FileRef?: string;
  Modified?: string;
  CanvasContent1?: string | null;
  LayoutWebpartsContent?: string | null;
  /**
   * Page-level banner image URL. Returned by SP REST as either a plain
   * string or a URL-field object (`{ Url, Description }`) depending on
   * site configuration; the scanner unwraps via `extractBannerImageUrl`.
   * Required for orphan-asset detection: banner-only references don't
   * appear in CanvasContent1 or LayoutWebpartsContent.
   */
  BannerImageUrl?: string | { Url?: string; Description?: string } | null;
  /**
   * SharePoint list item ETag, captured at scan time. Used by the
   * find-and-replace write path to detect lost-update races: if the
   * page has been edited since the scan, the write is refused with a
   * 412 and the user is asked to rescan. Format: `"<version>"` (with
   * the surrounding quotes — that's how SP returns it and how
   * If-Match expects to receive it).
   */
  etag?: string;
}

/**
 * Pull a string URL out of a SitePageItem.BannerImageUrl regardless of
 * which on-the-wire shape SP returned (string vs. URL-field object).
 * Returns undefined when the field is missing/empty.
 */
export function extractBannerImageUrl(
  field: SitePageItem["BannerImageUrl"],
): string | undefined {
  if (!field) return undefined;
  if (typeof field === "string") return field.length > 0 ? field : undefined;
  if (typeof field === "object" && typeof field.Url === "string" && field.Url.length > 0) {
    return field.Url;
  }
  return undefined;
}

// We deliberately do NOT request `Title` here. Some site collections —
// notably Project Web App (`/sites/pwa`) — provision a Site Pages list
// without the standard `Title` column, and SP returns a 400 SPException
// "The field or property 'Title' does not exist." for the entire query.
// FileLeafRef is always present and our `pageTitle` builder already
// falls back to it cleanly, so dropping Title makes the SELECT survive
// non-standard list schemas without losing anything in the common case.
const SELECT_FIELDS = [
  "Id",
  "FileLeafRef",
  "FileRef",
  "Modified",
  "CanvasContent1",
  "LayoutWebpartsContent",
  "BannerImageUrl",
].join(",");

interface RawListItem extends SitePageItem {
  // Minimal-metadata responses include the ETag as a per-item property
  // named `@odata.etag` (yes, with the @). We hoist it onto the
  // SitePageItem.etag field after fetching.
  "@odata.etag"?: string;
}

interface ListItemsResponse {
  value: RawListItem[];
  "odata.nextLink"?: string;
}

/**
 * Fetch a single Site Pages list item by id, returning the current ETag.
 * Used by the find-and-replace write path for pre-write concurrency
 * checks (compare current ETag to the one captured at scan time).
 *
 * Pass `userToken` to make the request as a specific user (delegated)
 * instead of as the service identity (app-only). The find-and-replace
 * path uses the user token so the audit trail reflects who actually
 * triggered the change.
 */
export async function fetchPageById(
  sitePath: string,
  pageId: number,
  userToken?: string,
): Promise<SitePageItem | undefined> {
  const token = userToken ?? (await getSpoToken());
  const url = `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('Site Pages')/items(${pageId})?$select=${SELECT_FIELDS}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=minimalmetadata",
    },
  });
  if (res.status === 404) return undefined;
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`fetchPageById ${sitePath}/${pageId} failed: ${res.status} ${body.slice(0, 300)}`);
  }
  const raw = (await res.json()) as RawListItem;
  const { ["@odata.etag"]: rawEtag, ...rest } = raw;
  return { ...rest, etag: rawEtag };
}

/**
 * PATCH a single Site Pages item with a new CanvasContent1 (and
 * optionally LayoutWebpartsContent), enforced via If-Match against the
 * caller-supplied ETag for optimistic concurrency.
 *
 * Returns one of:
 *   - { ok: true, newEtag } on success
 *   - { ok: false, conflict: true } when SP returns 412 (page changed)
 *   - { ok: false, error } for any other failure
 *
 * Pass `userToken` to PATCH as the calling user (delegated AllSites.Write)
 * instead of the app-only service identity. The find-and-replace path
 * uses the user token so:
 *   - SP audit logs show the actual user who made the change
 *   - SP enforces the user's per-list/per-item permissions
 *   - No Sites.Selected per-site grant is needed (the user's normal
 *     SP edit permissions apply)
 *
 * The caller is still responsible for verifying the user has Edit on
 * the site via the OBO `effectiveBasePermissions` check (defense in
 * depth + early/clean error).
 */
export async function patchPageCanvas(
  sitePath: string,
  pageId: number,
  ifMatchEtag: string,
  canvasContent1: string,
  layoutWebpartsContent?: string,
  userToken?: string,
): Promise<{ ok: true; newEtag?: string } | { ok: false; conflict?: boolean; status?: number; error: string }> {
  const token = userToken ?? (await getSpoToken());
  const url = `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('Site Pages')/items(${pageId})`;
  const body: Record<string, unknown> = {
    __metadata: { type: "SP.Data.SitePagesItem" },
    CanvasContent1: canvasContent1,
  };
  if (layoutWebpartsContent !== undefined) {
    body.LayoutWebpartsContent = layoutWebpartsContent;
  }

  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=verbose",
      "X-HTTP-Method": "MERGE",
      "If-Match": ifMatchEtag,
    },
    body: JSON.stringify(body),
  });

  if (res.status === 412) {
    return { ok: false, conflict: true, status: 412, error: "Page modified since scan" };
  }
  if (!res.ok) {
    const text = await res.text();
    return { ok: false, status: res.status, error: text.slice(0, 400) };
  }
  // SP MERGE doesn't return a body; the new ETag is in the response header
  const newEtag = res.headers.get("ETag") ?? undefined;
  return { ok: true, newEtag };
}

/**
 * Publish a Site Pages item so it goes from draft to live. Called
 * after a successful PATCH to ensure the user doesn't have to
 * manually publish every page that was modified by Find & Replace.
 *
 * Uses the File/Publish endpoint. Fails silently if the page doesn't
 * support publishing (e.g. no minor versioning enabled) — the PATCH
 * still succeeded, it just stays as a draft.
 */
export async function publishPage(
  sitePath: string,
  pageId: number,
  userToken?: string,
): Promise<{ ok: boolean; error?: string }> {
  const token = userToken ?? (await getSpoToken());
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('Site Pages')/items(${pageId})/File/Publish('')`;
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
      },
    });
    if (!res.ok) {
      const body = await res.text();
      return { ok: false, error: `Publish failed: ${res.status} ${body.slice(0, 200)}` };
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: `Publish threw: ${(err as Error).message}` };
  }
}

/**
 * Fetch Site Pages list items for the given site path.
 *
 * Pass `pageIds` to fetch only those specific items instead of the
 * entire list — used by the targeted-page rescan path. When `pageIds`
 * is omitted, fetches every item via paged $top=100 requests.
 *
 * Site path is the server-relative URL like `/sites/charlie-test-site`
 * (no host, no trailing slash).
 */
export async function fetchSitePages(sitePath: string, pageIds?: number[]): Promise<SitePageItem[]> {
  const token = await getSpoToken();
  const out: SitePageItem[] = [];

  // Build $filter clause when targeting specific page ids. SP REST
  // accepts up to ~50 OR clauses cleanly; for the targeted rescan
  // case the caller will rarely pass more than a handful so this is
  // fine. Larger lists would need batching.
  const filterClause = pageIds && pageIds.length > 0
    ? "&$filter=" + encodeURIComponent(pageIds.map((id) => `Id eq ${id}`).join(" or "))
    : "";

  let url: string | undefined =
    `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('Site Pages')/items` +
    `?$select=${SELECT_FIELDS}&$top=100${filterClause}`;

  while (url) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        // minimalmetadata (not nometadata) so that each item carries
        // its `@odata.etag`, which we need for optimistic concurrency
        // on the find-and-replace write path.
        Accept: "application/json;odata=minimalmetadata",
      },
    });
    if (!res.ok) {
      const body = await res.text();
      // Sites that don't have a modern Site Pages library at all —
      // Project Web App (PWA), classic team sites, some templates from
      // SP 2010 / 2013 — return either a 404 (the list doesn't exist)
      // or a 400 with "field 'CanvasContent1' does not exist" (the list
      // exists but is the classic Wiki Pages library without the
      // modern canvas column). Treat both as "no modern pages here"
      // and return an empty list — there's nothing for us to scan,
      // and it's not an error from the user's perspective.
      const isMissingList = res.status === 404;
      const isMissingCanvasField =
        res.status === 400 &&
        /CanvasContent1.{0,40}does not exist/i.test(body);
      const isMissingListAlt =
        res.status === 400 &&
        /(List 'Site Pages' does not exist|getbytitle.*does not exist)/i.test(body);
      if (isMissingList || isMissingCanvasField || isMissingListAlt) {
        // eslint-disable-next-line no-console
        console.warn(
          `[fetchSitePages] ${sitePath}: no modern Site Pages library — skipping (${res.status})`,
        );
        return out;
      }
      throw new Error(`Site Pages fetch failed for ${sitePath}: ${res.status} ${body.slice(0, 300)}`);
    }
    const data = (await res.json()) as ListItemsResponse;
    for (const raw of data.value) {
      const { ["@odata.etag"]: rawEtag, ...rest } = raw;
      out.push({ ...rest, etag: rawEtag });
    }
    url = data["odata.nextLink"];
  }

  return out;
}
