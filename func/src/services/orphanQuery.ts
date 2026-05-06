import { canonicalKeyForFile, type BacklinksIndex } from "./backlinksIndex.js";
import type { SitePagesAssetFile } from "./spoSitePagesAssetsEnumerator.js";

/**
 * Pure orphan-query: given a file list and a backlinks index, return
 * the subset of files no current page references.
 *
 * The premise — already validated by the standalone PowerShell module
 * against a real tenant — is that the existing persistent backlinks
 * index already answers "what does any current page reference?" for
 * every URL in the tenant. Computing each file's canonical key and
 * looking it up in `index.byCanonicalKey` is the whole detection.
 *
 * Why this is pure: no SP calls, no I/O. Both inputs are produced by
 * other services (`enumerateSitePagesAssets` for the file list,
 * `getInMemoryIndex` for the index). Unit tests can construct both
 * by hand and exhaustively cover the rename / custom-template /
 * cross-folder / deleted-page / stray-upload scenarios.
 *
 * Reliability cases handled implicitly because the backlinks index
 * already handles them:
 *
 *   - Renamed pages: SharePoint does NOT rename the asset folder when
 *     a page is renamed. The page's CanvasContent1 still references
 *     the original folder URL. That URL is in the index, so files in
 *     the original folder are correctly retained.
 *   - Custom-template GUID folders: pages from custom templates name
 *     their SitePages folder with a GUID instead of the page title.
 *     Same mechanism — the page references the GUID folder URL, the
 *     index has it, files are correctly retained.
 *   - Cross-folder references: page A references a file that lives
 *     in page B's folder. The reference is in the index keyed by the
 *     file's canonical URL (not by which folder it "belongs to"), so
 *     the file is retained as long as anything points at it.
 *   - Banner image references: handled by the BannerImageUrl scanner
 *     extension (see linkInventoryScanner.buildPageInventory). Without
 *     that extension, banner-only assets would false-flag as orphan.
 */

export interface OrphanFile {
  /** Owning site path, e.g. `/sites/hub` */
  sitePath: string;
  /** Server-relative URL of the file */
  serverRelativeUrl: string;
  /** Just the leaf name with extension */
  fileName: string;
  /**
   * First path segment under `SiteAssets/SitePages/` — usually a page
   * name, sometimes a GUID for custom-template-derived pages, sometimes
   * a stray-upload folder. Empty string if the URL doesn't match the
   * expected shape.
   */
  pageFolder: string;
  /** File size in bytes */
  size: number;
  /** Last modified ISO timestamp */
  modified: string;
  /** Path-only canonical key (the same key used for index lookup) */
  canonicalKey: string;
}

/**
 * Pull the first path segment under `/SiteAssets/SitePages/` from a
 * server-relative URL. Defensive against odd shapes — returns "" when
 * the URL doesn't have the expected prefix or has no segment after it.
 */
export function extractPageFolder(serverRelativeUrl: string): string {
  const m = /\/SiteAssets\/SitePages\/([^/]+)/i.exec(serverRelativeUrl);
  return m ? m[1] : "";
}

/**
 * Return the subset of `files` that no entry in `index` references.
 *
 * Files whose canonical key can't be computed (defensive: shouldn't
 * happen on real SP server-relative URLs, but possible for malformed
 * inputs) are SKIPPED — neither flagged as orphan nor counted as
 * referenced. Logging/surfacing those is the caller's responsibility.
 */
export function findOrphans(
  files: SitePagesAssetFile[],
  index: BacklinksIndex,
): OrphanFile[] {
  const out: OrphanFile[] = [];
  for (const f of files) {
    const canonicalKey = canonicalKeyForFile(f.serverRelativeUrl);
    if (!canonicalKey) continue;
    const refs = index.byCanonicalKey.get(canonicalKey);
    if (refs && refs.length > 0) continue; // referenced → not orphan
    out.push({
      sitePath: f.sitePath,
      serverRelativeUrl: f.serverRelativeUrl,
      fileName: f.name,
      pageFolder: extractPageFolder(f.serverRelativeUrl),
      size: f.size,
      modified: f.timeLastModified,
      canonicalKey,
    });
  }
  return out;
}
