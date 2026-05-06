import { getSpoToken, SPO_ORIGIN } from "./spoTokenProvider.js";

/**
 * SP REST file enumeration for `SiteAssets/SitePages/` — the asset
 * subtree where modern pages store their per-page images and uploads.
 *
 * Walks the folder tree recursively and returns every file underneath.
 * Used by the orphan-asset feature to compare what's on disk against
 * what the persistent backlinks index says is referenced.
 *
 * Why a separate walker from `spoFilesEnumerator`:
 *   - That one targets visible document libraries via `web/lists`
 *     (BaseTemplate=101 + Hidden=false) and explicitly skips
 *     "Site Pages", "Form Templates", "Style Library". Site Assets is
 *     typically hidden on classic sites and frequently missing
 *     entirely on freshly-provisioned Communication sites.
 *   - We don't need library-level metadata, just file metadata under a
 *     known folder path. `getfolderbyserverrelativeurl` is the simpler
 *     fit and works regardless of the list's hidden flag.
 *
 * Failure modes handled inline:
 *   - 404 at the root SitePages folder → site has never produced page
 *     assets; return empty array, no error.
 *   - 404 / failure on a subfolder during recursion → warn and skip
 *     just that subfolder; the rest of the walk continues.
 *
 * Authentication:
 *   - Uses app-only token (`getSpoToken`). Read-only enumeration; the
 *     caller filters the resulting site list to user-readable sites
 *     via `filterReadableSites` before any results leave the function.
 */

export interface SitePagesAssetFile {
  /** Owning site path, e.g. `/sites/hub` */
  sitePath: string;
  /** Full server-relative URL of the file */
  serverRelativeUrl: string;
  /** Just the leaf name with extension */
  name: string;
  /** File size in bytes */
  size: number;
  /** Last modified ISO timestamp */
  timeLastModified: string;
}

/**
 * Per-segment URL-encode the given server-relative path while preserving
 * slashes, then escape single quotes the SP OData way (`''`). Matches
 * the encoding in `spoFilesEnumerator.fetchFileBytes`.
 */
export function encodeServerRelativePath(path: string): string {
  return path
    .split("/")
    .map((seg) => encodeURIComponent(seg))
    .join("/")
    .replace(/'/g, "''");
}

interface RawFolderListing {
  Folders?: { results?: unknown[] } | unknown[];
  Files?: { results?: unknown[] } | unknown[];
}

/**
 * Pure helper: extract typed file + subfolder records from a SP REST
 * `getfolderbyserverrelativeurl?$expand=Folders,Files` response.
 *
 * Tolerates both shapes SP can return depending on metadata mode:
 *   nometadata    → { Folders: [...], Files: [...] }
 *   minimalmetadata / verbose → { Folders: { results: [...] }, Files: { results: [...] } }
 */
export function parseFolderListing(
  data: unknown,
  sitePath: string,
): { files: SitePagesAssetFile[]; subfolders: string[] } {
  if (!data || typeof data !== "object") return { files: [], subfolders: [] };

  const raw = data as RawFolderListing;

  const filesArray = Array.isArray(raw.Files)
    ? raw.Files
    : Array.isArray((raw.Files as { results?: unknown[] } | undefined)?.results)
      ? (raw.Files as { results: unknown[] }).results
      : [];

  const foldersArray = Array.isArray(raw.Folders)
    ? raw.Folders
    : Array.isArray((raw.Folders as { results?: unknown[] } | undefined)?.results)
      ? (raw.Folders as { results: unknown[] }).results
      : [];

  const files: SitePagesAssetFile[] = [];
  for (const f of filesArray) {
    if (!f || typeof f !== "object") continue;
    const o = f as {
      ServerRelativeUrl?: unknown;
      Name?: unknown;
      Length?: unknown;
      TimeLastModified?: unknown;
    };
    if (typeof o.ServerRelativeUrl !== "string" || typeof o.Name !== "string") continue;
    const sizeRaw = o.Length;
    const size =
      typeof sizeRaw === "number"
        ? sizeRaw
        : typeof sizeRaw === "string"
          ? parseInt(sizeRaw, 10) || 0
          : 0;
    files.push({
      sitePath,
      serverRelativeUrl: o.ServerRelativeUrl,
      name: o.Name,
      size,
      timeLastModified: typeof o.TimeLastModified === "string" ? o.TimeLastModified : "",
    });
  }

  const subfolders: string[] = [];
  for (const sub of foldersArray) {
    if (!sub || typeof sub !== "object") continue;
    const o = sub as { ServerRelativeUrl?: unknown; Name?: unknown };
    if (typeof o.ServerRelativeUrl !== "string") continue;
    // Defensive: skip SP system folders. SitePages doesn't have a Forms
    // subfolder in practice, but if SP ever includes one or any other
    // underscore-prefixed system folder, don't recurse into it.
    if (typeof o.Name === "string" && o.Name.startsWith("_")) continue;
    subfolders.push(o.ServerRelativeUrl);
  }

  return { files, subfolders };
}

interface EnumerateOptions {
  /** Override the fetch implementation. For tests. */
  fetchImpl?: typeof fetch;
  /** Override the token provider. For tests. */
  tokenProvider?: () => Promise<string>;
  /** Logger; defaults to console. */
  log?: (msg: string) => void;
  warn?: (msg: string) => void;
}

/**
 * Enumerate every file under `<sitePath>/SiteAssets/SitePages/` recursively.
 *
 * Returns an empty array if the folder doesn't exist (sites that have
 * never had a page with assets — common on fresh Communication sites).
 *
 * Per-subfolder failures during recursion are logged and skipped; only a
 * failure at the root that isn't 404 throws.
 */
export async function enumerateSitePagesAssets(
  sitePath: string,
  options: EnumerateOptions = {},
): Promise<SitePagesAssetFile[]> {
  const fetchImpl = options.fetchImpl ?? fetch;
  const tokenProvider = options.tokenProvider ?? getSpoToken;
  const warn = options.warn ?? ((m) => {
    // eslint-disable-next-line no-console
    console.warn(m);
  });

  const token = await tokenProvider();
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json;odata=nometadata",
  };

  const rootFolder = `${sitePath}/SiteAssets/SitePages`;
  const out: SitePagesAssetFile[] = [];
  const queue: string[] = [rootFolder];

  while (queue.length > 0) {
    const folder = queue.shift()!;
    const encoded = encodeServerRelativePath(folder);
    const url =
      `${SPO_ORIGIN}${sitePath}/_api/web/getfolderbyserverrelativeurl('${encoded}')` +
      `?$expand=Folders,Files`;

    const res = await fetchImpl(url, { headers });
    if (!res.ok) {
      const body = await res.text();
      // Root folder missing → site has never produced page assets.
      // Both 404 (folder absent) and 400 + "does not exist" (some SP
      // tenants return 400 for missing folders) are normal here.
      const isMissing =
        res.status === 404 ||
        (res.status === 400 && /does not exist/i.test(body));
      if (folder === rootFolder && isMissing) {
        return [];
      }
      // Per-subfolder failure: log + continue. Don't fail the whole walk
      // on a single bad subfolder.
      warn(
        `[enumerateSitePagesAssets] ${folder}: ${res.status} ${body.slice(0, 200)}`,
      );
      continue;
    }

    const json = (await res.json()) as unknown;
    const { files, subfolders } = parseFolderListing(json, sitePath);
    for (const f of files) out.push(f);
    for (const sf of subfolders) queue.push(sf);
  }

  return out;
}
