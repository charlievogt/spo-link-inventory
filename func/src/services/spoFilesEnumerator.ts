import { getSpoToken, SPO_ORIGIN } from "./spoTokenProvider.js";
import { detectFileType } from "./documentLinkExtractor.js";

/**
 * SP REST file enumeration for the document link inventory.
 *
 * Walks every document library in a given site and returns the files
 * that the document extractor knows how to parse (.docx/.xlsx/.pptx/.pdf).
 * Other file types are silently skipped.
 *
 * Why we don't try to be cute about library structure:
 *   - Lots of sites have multiple custom document libraries (e.g.
 *     "Forms", "Templates", "Shared Documents") — we want all of them
 *   - List BaseTemplate 101 is the canonical document library template;
 *     filtering on that gives us libraries while excluding announcement
 *     lists, calendars, custom lists with attachments, etc.
 *   - We don't follow folders explicitly — `CamlQuery` with
 *     `RecursiveAll` walks the whole tree in one query per library.
 *
 * Pagination is via `$skiptoken` (the cursor SP returns when we hit
 * `$top`). We page until exhausted per library. Each `enumerateFiles`
 * call returns the full list — pagination is internal.
 *
 * Filtering at the SP layer:
 *   - Limit to files (FSObjType eq 0)
 *   - Skip checked-out files (CheckOutType ne 2 implies checked in)
 *   - We do NOT filter by extension at the SP layer because CAML
 *     filters are awkward and the extension check is cheap client-side.
 *
 * Throttling:
 *   - SP throttles aggressively on bulk file metadata reads. We page
 *     at 500 items per request and rely on the orchestrator's per-file
 *     queue chaining for the actual download phase to stay tame.
 */

export interface DocumentFileRef {
  /** Owning site path, e.g. `/sites/hub` */
  site: string;
  /** Library title, e.g. `Shared Documents` */
  library: string;
  /** Server-relative URL of the file, e.g. `/sites/hub/Shared Documents/foo.pdf` */
  fileRef: string;
  /** Just the leaf name with extension */
  fileName: string;
  /** Last modified timestamp (ISO 8601) */
  modified: string;
  /** SP item ETag — for change detection between scans */
  etag?: string;
  /** File size in bytes */
  length: number;
  /**
   * Detected file type from the extension. `"other"` means we don't
   * have a link-extraction parser for it; the enumerator still yields
   * such files when `includeOther: true` so downstream code (duplicate
   * detection, preview) can reason about every file in the library.
   */
  fileType: "docx" | "xlsx" | "pptx" | "pdf" | "other";
  /**
   * Lowercased file extension with no leading dot (`"pdf"`, `"jpg"`,
   * `""` when the name has no extension). Useful for the preview's
   * per-extension breakdown and for admin opt-in to scan additional
   * extensions beyond the default supported set.
   */
  extension: string;
}

interface RawListItem {
  Id: number;
  FileLeafRef?: string;
  FileRef?: string;
  Modified?: string;
  File?: { Length?: string };
  "@odata.etag"?: string;
}

interface ListItemsResponse {
  value: RawListItem[];
  "odata.nextLink"?: string;
}

interface RawList {
  Id: string;
  Title: string;
  BaseTemplate: number;
  Hidden: boolean;
  RootFolder?: { ServerRelativeUrl: string };
}

interface ListsResponse {
  value: RawList[];
}

const DOCUMENT_LIBRARY_TEMPLATE = 101;

/**
 * Default file-size cap. The user can override this on a per-scan basis
 * via the trigger endpoint body. The default is conservative because
 * the function downloads the entire file before parsing — pulling a
 * 500 MB binary into a Function instance is the kind of thing that
 * gets you a 503.
 *
 * In practice the user should pick a cap that fits their actual content:
 * The upstream has plenty of multi-MB Word and Excel files that need scanning,
 * so the default is intentionally generous.
 */
export const MAX_FILE_BYTES_DEFAULT = 100 * 1024 * 1024; // 100 MB

/**
 * Hard ceiling we will never exceed regardless of what the caller asks
 * for. Sized to comfortably fit in a Function instance's memory budget
 * including JSZip's working buffers.
 */
export const MAX_FILE_BYTES_HARD_LIMIT = 500 * 1024 * 1024; // 500 MB

interface EnumerateOptions {
  /**
   * Files larger than this are dropped from the result. Defaults to
   * MAX_FILE_BYTES_DEFAULT (100 MB). Capped to MAX_FILE_BYTES_HARD_LIMIT
   * (500 MB) — values above are silently clamped.
   */
  maxBytes?: number;
  /**
   * If set, only return files modified strictly after this timestamp.
   * Used for incremental scans.
   */
  modifiedAfter?: Date;
  /**
   * When true, yield every file regardless of extension — unsupported
   * types come back tagged `fileType: "other"`. When false (default),
   * retain current behavior and drop unsupported types silently.
   *
   * The duplicate-detection pipeline sets this to true so the preview
   * can show admins what their libraries actually contain. The normal
   * link-extraction scan filters the manifest back down to supported
   * types before handing it to the scan phase.
   */
  includeOther?: boolean;
}

/**
 * Enumerate every document-library file in `sitePath` whose extension
 * is one we know how to parse. Returns a flat array.
 *
 * The orchestrator should call this once per site at scan-trigger time,
 * persist the file list, and then process files via the queue worker.
 */
export async function enumerateFiles(
  sitePath: string,
  options: EnumerateOptions = {},
): Promise<DocumentFileRef[]> {
  const requested = options.maxBytes ?? MAX_FILE_BYTES_DEFAULT;
  const maxBytes = Math.min(Math.max(1, requested), MAX_FILE_BYTES_HARD_LIMIT);
  const token = await getSpoToken();
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json;odata=minimalmetadata",
  };

  // 1) List all visible document libraries on the site.
  const listsUrl =
    `${SPO_ORIGIN}${sitePath}/_api/web/lists` +
    `?$select=Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl` +
    `&$expand=RootFolder` +
    `&$filter=BaseTemplate eq ${DOCUMENT_LIBRARY_TEMPLATE} and Hidden eq false`;
  const listsRes = await fetch(listsUrl, { headers });
  if (!listsRes.ok) {
    const body = await listsRes.text();
    throw new Error(`Lists fetch failed for ${sitePath}: ${listsRes.status} ${body.slice(0, 300)}`);
  }
  const listsJson = (await listsRes.json()) as ListsResponse;
  const libraries = listsJson.value.filter((l) => l.RootFolder?.ServerRelativeUrl);
  // eslint-disable-next-line no-console
  console.log(
    `[enumerateFiles] ${sitePath}: ${libraries.length} document librar${libraries.length === 1 ? 'y' : 'ies'} found: ${libraries.map((l) => l.Title).join(', ') || '(none)'}`,
  );

  const out: DocumentFileRef[] = [];
  for (const lib of libraries) {
    // Skip standard noise libraries that exist on every site:
    //   - "Site Pages" — page library, scanned by the page inventory
    //   - "Form Templates" — XSN form templates, not user content
    //   - "Style Library" — branding assets
    if (
      lib.Title === "Site Pages" ||
      lib.Title === "Form Templates" ||
      lib.Title === "Style Library"
    ) {
      continue;
    }

    let libItemsSeen = 0;
    let libParseable = 0;
    // 2) Page through the library's items. We use $select to keep
    // payload small, $top=500 for batch size, and $orderby=Id so the
    // query uses the always-indexed Id column. Without orderby, libraries
    // over the 5000-item list view threshold throw SPQueryThrottledException
    // even with a small $top, because $filter forces a full table scan.
    let url: string | undefined =
      `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('${encodeURIComponent(lib.Title)}')/items` +
      `?$select=Id,FileLeafRef,FileRef,Modified,File/Length` +
      `&$expand=File` +
      `&$filter=FSObjType eq 0` +
      `&$orderby=Id` +
      `&$top=500`;

    while (url) {
      const res: Response = await fetch(url, { headers });
      if (!res.ok) {
        const body = await res.text();
        // Library-level failures are non-fatal — we still want the rest
        // of the site's libraries scanned. The most common failure is
        // SPQueryThrottledException on libraries that have grown past
        // the list view threshold without indexed columns; nothing we
        // can do at runtime besides skip them.
        // eslint-disable-next-line no-console
        console.warn(
          `[enumerateFiles] skipping ${sitePath} library "${lib.Title}": ${res.status} ${body.slice(0, 200)}`,
        );
        url = undefined;
        break;
      }
      const data = (await res.json()) as ListItemsResponse;
      for (const item of data.value) {
        libItemsSeen++;
        const fileName = item.FileLeafRef;
        const fileRef = item.FileRef;
        const modified = item.Modified;
        if (!fileName || !fileRef || !modified) continue;
        const detected = detectFileType(fileName);
        const isSupported = detected !== "unsupported";
        if (!isSupported && !options.includeOther) continue;
        const length = parseInt(item.File?.Length ?? "0", 10);
        if (Number.isNaN(length) || length === 0 || length > maxBytes) continue;
        if (options.modifiedAfter && new Date(modified) <= options.modifiedAfter) continue;
        if (isSupported) libParseable++;
        const dot = fileName.lastIndexOf(".");
        const extension = dot >= 0 ? fileName.slice(dot + 1).toLowerCase() : "";
        out.push({
          site: sitePath,
          library: lib.Title,
          fileRef,
          fileName,
          modified,
          etag: item["@odata.etag"],
          length,
          fileType: isSupported ? detected : "other",
          extension,
        });
      }
      url = data["odata.nextLink"];
    }
    // eslint-disable-next-line no-console
    console.log(
      `[enumerateFiles]   ${sitePath}/${lib.Title}: ${libItemsSeen} items seen, ${libParseable} parseable`,
    );
  }

  return out;
}

/**
 * Fetch metadata for a single file by server-relative URL. Used by
 * the targeted-file scan path: when the user passes `fileRefs` to the
 * trigger, we don't run library enumeration and instead build the
 * manifest one file at a time via this helper.
 *
 * Returns `undefined` for files that don't exist or aren't a parseable
 * type. Throws on auth/network failures.
 */
export async function fetchFileMetadata(
  fileRef: string,
  options: { includeOther?: boolean } = {},
): Promise<DocumentFileRef | undefined> {
  const token = await getSpoToken();
  // Same site-context + encoding rules as fetchFileBytes — calling at
  // the root web returns 404 for files in subsites.
  const siteMatch = /^(\/(?:sites|teams)\/[^/]+)\//i.exec(fileRef);
  const sitePath = siteMatch ? siteMatch[1] : "";
  const encodedRef = fileRef
    .split("/")
    .map((seg) => encodeURIComponent(seg))
    .join("/")
    .replace(/'/g, "''");
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/getfilebyserverrelativeurl('${encodedRef}')` +
    `?$select=Length,TimeLastModified,ServerRelativeUrl,Name`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
    },
  });
  if (res.status === 404) return undefined;
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`fetchFileMetadata ${fileRef} failed: ${res.status} ${body.slice(0, 200)}`);
  }
  const data = (await res.json()) as {
    Length?: string;
    TimeLastModified?: string;
    ServerRelativeUrl?: string;
    Name?: string;
  };

  const fileName = data.Name;
  if (!fileName) return undefined;
  const detected = detectFileType(fileName);
  const isSupported = detected !== "unsupported";
  if (!isSupported && !options.includeOther) return undefined;

  const length = parseInt(data.Length ?? "0", 10);
  // Derive site from the server-relative path: `/sites/<site>/...` → `/sites/<site>`
  const fullRef = data.ServerRelativeUrl ?? fileRef;
  const fullSiteMatch = /^(\/(?:sites|teams)\/[^/]+)\//i.exec(fullRef);
  const fullSitePath = fullSiteMatch ? fullSiteMatch[1] : "/";
  // Library is the first folder under the site root
  const afterSite = fullSiteMatch ? fullRef.slice(fullSiteMatch[1].length + 1) : fullRef.replace(/^\//, "");
  const library = afterSite.split("/", 1)[0] ?? "(unknown)";
  const dot = fileName.lastIndexOf(".");
  const extension = dot >= 0 ? fileName.slice(dot + 1).toLowerCase() : "";

  return {
    site: fullSitePath,
    library,
    fileRef: fullRef,
    fileName,
    modified: data.TimeLastModified ?? new Date().toISOString(),
    length,
    fileType: isSupported ? detected : "other",
    extension,
  };
}

/**
 * Stream a single file's bytes from SP REST. Used by the document
 * scan worker right before calling `extractLinksFromDocument`.
 *
 * Returns a Buffer to keep the byte-in / link-out contract simple.
 *
 * IMPORTANT: `GetFileByServerRelativeUrl` MUST be called in the
 * context of the file's own site collection — calling it at the root
 * web (`https://contoso.sharepoint.com/_api/web/...`) returns a 404
 * for files in subsites because cross-site resolution isn't supported
 * from the root context. We derive the site path from the fileRef and
 * invoke the API at that site.
 */
export async function fetchFileBytes(fileRef: string): Promise<Buffer> {
  const token = await getSpoToken();
  // Derive site collection path from the server-relative URL:
  //   /sites/<name>/...   → /sites/<name>
  //   /teams/<name>/...   → /teams/<name>
  //   anything else       → "" (root web)
  const siteMatch = /^(\/(?:sites|teams)\/[^/]+)\//i.exec(fileRef);
  const sitePath = siteMatch ? siteMatch[1] : "";

  // SharePoint REST `getfilebyserverrelativeurl` encoding:
  //   1. Per-segment URL-encode (spaces → %20) but PRESERVE slashes.
  //      `encodeURIComponent` on the whole path encodes "/" as %2F → 404.
  //   2. Single quotes inside the path must be doubled (SP OData escape),
  //      otherwise they terminate the OData string literal early.
  const encodedRef = fileRef
    .split("/")
    .map((seg) => encodeURIComponent(seg))
    .join("/")
    .replace(/'/g, "''");
  const url = `${SPO_ORIGIN}${sitePath}/_api/web/getfilebyserverrelativeurl('${encodedRef}')/$value`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      // No Accept header — we want raw bytes, not an OData wrapper
    },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`File download failed for ${fileRef}: ${res.status} ${body.slice(0, 200)}`);
  }
  const ab = await res.arrayBuffer();
  return Buffer.from(ab);
}
