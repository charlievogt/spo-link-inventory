import { fetchSitePages, extractBannerImageUrl, type SitePageItem } from "./spoPagesClient.js";
import { extractLinksFromCanvas, type ExtractedLink } from "./canvasLinkExtractor.js";
import { normalizeUrl } from "./urlNormalizer.js";
import { getTenantHost, getTenantOrigin } from "./config.js";

/**
 * Core scanning logic. The HTTP endpoint and the tenant-wide orchestrator
 * both call into this so the extraction + classification rules stay in
 * one place.
 */

export type LinkClass =
  | "spo-internal"
  | "onprem"
  | "external"
  | "mailto"
  | "tel"
  | "anchor-only"
  | "javascript"
  | "sharing-link"
  | "office-online"
  | "doc-aspx"
  | "relative"
  | "malformed-spo-link"
  | "unknown";

/**
 * Why a link was classified as `malformed-spo-link`.
 *
 * - `search-fragment` — AllItems.aspx URL carries a leftover `#search=`
 *   fragment from the library-search UI. Usually indicates the link
 *   was captured while someone was searching inside a library, or a
 *   find-and-replace swapped the id= filename and left the fragment
 *   in place (smoking gun for a bad replace).
 * - `file-not-found` — the file path referenced in `?id=` returned
 *   404 when verified against SP REST. Only set when the opt-in
 *   verification phase runs.
 * - `both` — both of the above.
 */
export type MalformedReason = "search-fragment" | "file-not-found" | "both";

export interface ClassifiedLink extends ExtractedLink {
  linkClass: LinkClass;
  /**
   * Populated only when `linkClass === "malformed-spo-link"`. Tells the
   * UI why the link is flagged so it can render a precise hint.
   */
  malformedReason?: MalformedReason;
  /** Normalized lookup key (from urlNormalizer) when applicable. */
  normalizedKey: string;
  /**
   * Suggested replacement URL for sharing links and Office Online
   * viewer wrappers — both of which encode the canonical file path
   * inline. The UI uses this as the smart default when the user clicks
   * "Fix this link" on a row. Empty when we can't recover a target
   * (e.g. opaque `/:u:/g/<token>` shares).
   */
  suggestion?: string;
  /**
   * Canonical equivalence key — collapses every URL form that points
   * at the same target item to the same string. Sharing wrappers,
   * Office Online viewer wrappers, AllItems forms, direct file paths,
   * and absolute SPO URLs all resolve to the same canonical key when
   * they reference the same file. The UI uses this for "this link is
   * also at N other places" grouping in the detail panel.
   *
   * Empty when neither the URL nor its suggestion can be normalized
   * (true opaque sharing tokens, mailto, anchor, javascript, etc.).
   */
  canonicalKey: string;
}

export interface PageInventory {
  pageId: number;
  pageTitle: string;
  pageUrl: string;
  modified?: string;
  /** ETag captured at scan time. Used by find-and-replace for If-Match. */
  etag?: string;
  links: ClassifiedLink[];
  parseError?: string;
  /**
   * True when this page's links were reused from the previous aggregate
   * because SP-reported `Modified` matches the prior scan. The page
   * fetch + canvas parse was skipped — `links` here is a copy of the
   * aggregate entry's links. Used at finalize to decide merge-or-
   * carry-forward semantics in the rollforward aggregate.
   */
  reusedFromPreviousScan?: boolean;
}

export interface SiteInventory {
  site: string;
  pageCount: number;
  linkCount: number;
  byClass: Record<string, number>;
  bySource: Record<string, number>;
  pages: PageInventory[];
  scanMs: number;
  error?: string;
}

const SPO_HOST = getTenantHost();
/**
 * Comma-separated list of legacy/on-prem SharePoint hostnames classified
 * as "onprem" (e.g. an SP 2016 farm being decommissioned). Optional —
 * empty by default. Set LEGACY_SP_HOSTS="sp-prod.example.com,sp-test.example.com"
 * in App Settings to enable.
 */
const ONPREM_HOSTS: readonly string[] = (process.env.LEGACY_SP_HOSTS ?? "")
  .split(",")
  .map((s) => s.trim().toLowerCase())
  .filter(Boolean);

/**
 * Heuristic: does this URL look like an AllItems.aspx library-view link
 * that's carrying a leftover `#search=` fragment? That fragment is URL
 * state SharePoint adds when a user clicks a search result inside a
 * library; it should never appear in an authored link. A link that has
 * it is almost always the result of someone pasting a URL from their
 * search session — and the `id=` file inside may not even match.
 *
 * Runs purely on the URL string, no network. Used at classify time for
 * always-on detection. The opt-in verification phase sets the same
 * `malformed-spo-link` class on `id=` paths that return 404.
 */
function hasAllItemsSearchFragment(url: string): boolean {
  if (!/\/forms\/allitems\.aspx\?/i.test(url)) return false;
  const hashIdx = url.indexOf("#");
  if (hashIdx < 0) return false;
  const fragment = url.slice(hashIdx + 1);
  return /(^|&)search=/i.test(fragment);
}

export function classifyLink(url: string): LinkClass {
  if (!url) return "unknown";
  const trimmed = url.trim();
  if (trimmed.startsWith("#")) return "anchor-only";
  if (/^mailto:/i.test(trimmed)) return "mailto";
  if (/^tel:/i.test(trimmed)) return "tel";
  if (/^javascript:/i.test(trimmed)) return "javascript";
  // SP sharing wrappers: `/:LETTER:/MODE/...` where LETTER is one of
  // u (opaque user share), b (browser/PDF viewer), w (Word viewer),
  // x (Excel viewer), p (PowerPoint viewer), o (OneNote viewer),
  // f (folder share), and MODE is one of r (read), s (share),
  // g (global/anonymous), w (write — rare). Use [a-z] for both
  // positions so we catch any single-letter combination SP throws
  // at us — the surrounding `:LETTER:/MODE/` shape is rigid enough
  // that false positives aren't a concern.
  if (/^(https?:\/\/[^/]+)?\/:[a-z]:\/[a-z]\//i.test(trimmed)) {
    // Office Online viewer letters: w (Word), x (Excel),
    // p (PowerPoint), o (OneNote). Everything else in the sharing
    // family is bucketed as `sharing-link` — including b (the
    // browser PDF viewer), u (opaque user share), and f (folder).
    if (/\/:[wxpo]:/i.test(trimmed)) return "office-online";
    return "sharing-link";
  }
  if (/\/_layouts\/15\/doc\.aspx/i.test(trimmed)) return "doc-aspx";
  if (/^https?:\/\//i.test(trimmed)) {
    try {
      const u = new URL(trimmed);
      const host = u.hostname.toLowerCase();
      if (host === SPO_HOST) {
        return hasAllItemsSearchFragment(trimmed) ? "malformed-spo-link" : "spo-internal";
      }
      if (ONPREM_HOSTS.includes(host)) return "onprem";
      return "external";
    } catch {
      return "unknown";
    }
  }
  // Server-relative AllItems.aspx URLs carry the same malformed pattern.
  if (trimmed.startsWith("/")) {
    if (hasAllItemsSearchFragment(trimmed)) return "malformed-spo-link";
    return "relative";
  }
  return "unknown";
}

export function classifyAll(links: ExtractedLink[]): ClassifiedLink[] {
  return links.map((link) => {
    const linkClass = classifyLink(link.url);
    let normalizedKey = "";
    if (
      linkClass === "spo-internal" ||
      linkClass === "onprem" ||
      linkClass === "relative" ||
      linkClass === "office-online" ||
      linkClass === "malformed-spo-link"
    ) {
      try { normalizedKey = normalizeUrl(link.url); } catch { /* ignore */ }
    }
    const suggestion = suggestReplacement(link.url, linkClass);
    let canonicalKey = normalizedKey;
    if (!canonicalKey && suggestion) {
      try { canonicalKey = normalizeUrl(suggestion); } catch { /* ignore */ }
    }
    const malformedReason: MalformedReason | undefined =
      linkClass === "malformed-spo-link" ? "search-fragment" : undefined;
    return { ...link, linkClass, normalizedKey, suggestion, canonicalKey, malformedReason };
  });
}

/**
 * Build a canonical AllItems URL from a server-relative file path.
 * Mirrors the logic the page admins manually apply when fixing sharing
 * links by hand: take the file's directory, append `Forms/AllItems.aspx`,
 * and use `?id=<full-path>&parent=<dir>` to land in the library view
 * with the file selected.
 *
 * Returns an empty string if the path doesn't look like a file we can
 * resolve (e.g. it's already a Forms URL, or it's a folder).
 */
function buildAllItemsUrl(serverRelativePath: string): string {
  // Strip query/fragment and leading slash
  const path = serverRelativePath.split(/[?#]/, 1)[0];
  if (!path || !path.startsWith("/")) return "";
  // Already an AllItems URL — nothing to suggest
  if (/\/forms\/(allitems|dispform)\.aspx$/i.test(path)) return "";
  // It's a folder if there's no extension on the last segment
  const segments = path.split("/").filter((s) => s.length > 0);
  if (segments.length < 3) return ""; // need at least /sites/<site>/<lib>
  // Reject SP system folders — `_layouts`, `_api`, `_vti_bin`, `_catalogs`
  // etc. are NOT document libraries and shouldn't be canonicalized.
  // Their canonical "library" segment would start with `_`.
  if (segments.some((s) => s.startsWith("_"))) return "";
  const last = segments[segments.length - 1];
  if (!/\.[a-z0-9]{1,8}$/i.test(last)) return ""; // no extension → folder
  // Reject web page extensions — `.aspx`, `.html`, etc. are pages, not
  // documents, and the AllItems library viewer won't open them anyway.
  if (/\.(aspx|asmx|asp|html|htm|php|jsp|master|mht|mhtml)$/i.test(last)) return "";
  // For a file inside a subfolder, library = first 3 segments (/sites/<site>/<lib>)
  // Encode using SP-style: spaces → %20, & → %26, etc. encodeURIComponent
  // per segment then rejoin with literal slashes
  const encPath = "/" + segments.map((s) => encodeURIComponent(s)).join("/");
  const encLib = "/" + segments.slice(0, 3).map((s) => encodeURIComponent(s)).join("/");
  return `${getTenantOrigin()}${encLib}/Forms/AllItems.aspx?id=${encPath}&parent=${encLib}`;
}

/**
 * Suggest a canonical replacement URL for the given link, when its
 * shape allows. Sharing-link shims and Office Online wrappers both
 * encode the original server-relative file path inline; we extract it
 * and rebuild it as a stable AllItems URL.
 *
 * Returns an empty string when no suggestion is possible (opaque
 * sharing-link tokens, external links, mailto, etc.).
 */
function suggestReplacement(url: string, linkClass: LinkClass): string {
  if (!url) return "";

  // Office Online viewer wrappers: /:w:/r/, /:x:/r/, /:p:/r/, /:o:/r/
  // Path comes after the wrapper.
  if (linkClass === "office-online") {
    const trimmed = url.replace(/^https?:\/\/[^/]+/i, "");
    const m = /^\/:[wxpo]:\/[rs]\/(.+)$/i.exec(trimmed);
    if (m) {
      // Decode any %xx escapes the wrapper added, then rebuild
      let path: string;
      try { path = decodeURIComponent("/" + m[1]); } catch { return ""; }
      return buildAllItemsUrl(path);
    }
    return "";
  }

  // /:b:/r/ — PDF viewer wrapper. Treat the same as Office Online.
  if (linkClass === "sharing-link") {
    const trimmed = url.replace(/^https?:\/\/[^/]+/i, "");
    const m = /^\/:b:\/[rs]\/(.+)$/i.exec(trimmed);
    if (m) {
      let path: string;
      try { path = decodeURIComponent("/" + m[1]); } catch { return ""; }
      return buildAllItemsUrl(path);
    }
    // /:u:/g/ — opaque sharing token, no path inside, can't suggest.
    return "";
  }

  // malformed-spo-link: at classify time this only fires for the
  // `search-fragment` reason (AllItems URL with a leftover `#search=`
  // tail). The safe canonical form is the same URL minus the fragment
  // — the `?id=` path is already an AllItems address. The opt-in
  // verifier clears the suggestion if the id= path later proves 404,
  // since fragment-stripping doesn't fix a link pointing at a non-
  // existent file.
  if (linkClass === "malformed-spo-link") {
    const hashIdx = url.indexOf("#");
    if (hashIdx > 0) return url.slice(0, hashIdx);
    return "";
  }

  // doc-aspx: /_layouts/15/Doc.aspx?sourcedoc={GUID} — needs SP API to
  // resolve the GUID to a path. Skipped for v1.

  return "";
}

export function buildPageInventory(item: SitePageItem): PageInventory {
  const pageUrl = item.FileRef ?? "";
  const pageTitle = item.Title ?? item.FileLeafRef ?? `#${item.Id}`;
  const canvas = item.CanvasContent1 ?? "";
  const layout = item.LayoutWebpartsContent ?? "";

  let links: ExtractedLink[] = [];
  let parseError: string | undefined;
  try {
    links = [
      ...extractLinksFromCanvas(canvas),
      ...extractLinksFromCanvas(layout),
    ];
  } catch (e) {
    parseError = (e as Error).message;
  }

  // Banner image URL lives on the page-level BannerImageUrl field, not
  // inside CanvasContent1 or LayoutWebpartsContent. The title-region
  // web part stores `imageSourceType: 4` ("URL") with `imageSources: {}`
  // — the actual URL is on the list item field. Without this synthesis,
  // a banner-only asset (no other web-part reference) is invisible to
  // the backlinks index and would false-flag as orphan.
  const bannerUrl = extractBannerImageUrl(item.BannerImageUrl);
  if (bannerUrl) {
    links.push({
      rawUrl: bannerUrl,
      url: bannerUrl,
      source: "banner",
      // No webPartInstanceId — banner is page-level, not a web part.
    });
  }

  return {
    pageId: item.Id,
    pageTitle,
    pageUrl,
    modified: item.Modified,
    etag: item.etag,
    links: classifyAll(links),
    parseError,
  };
}

export function tallyPages(pages: PageInventory[]): Pick<SiteInventory, "linkCount" | "byClass" | "bySource"> {
  let linkCount = 0;
  const byClass: Record<string, number> = {};
  const bySource: Record<string, number> = {};
  for (const p of pages) {
    for (const l of p.links) {
      linkCount++;
      byClass[l.linkClass] = (byClass[l.linkClass] ?? 0) + 1;
      bySource[l.source] = (bySource[l.source] ?? 0) + 1;
    }
  }
  return { linkCount, byClass, bySource };
}

/**
 * Optional reuse hint for the Modified-based delta-skip path. The
 * worker passes this from the previous per-site aggregate so unchanged
 * pages can short-circuit the canvas extraction.
 *
 * Map shape: pageUrl → { modified, links }. The keys are the pageUrl
 * strings emitted by `buildPageInventory` (SP `FileRef`).
 */
export interface PreviousPagesByUrl {
  [pageUrl: string]: { modified: string; links: ClassifiedLink[] };
}

/**
 * Scan one site end-to-end. Pass `pageIds` to scan only those specific
 * pages within the site (used by the targeted-page rescan path).
 *
 * `options.previousPages` enables the **Modified-skip optimization**: a
 * page whose SP-reported `Modified` matches the previous aggregate
 * entry's `modified` reuses the previously-classified links. The canvas
 * is still fetched (SP doesn't separate page metadata fetch from canvas
 * fetch), but the parse + classify pass is skipped. Net win is small
 * per-page but adds up across thousands of unchanged pages on daily
 * scheduled scans.
 *
 * Returns a populated SiteInventory or one with `error` set if the SP
 * fetch failed (e.g. permission missing). Never throws — the
 * orchestrator should be able to keep going past a single failed site.
 */
export async function scanSite(
  sitePath: string,
  pageIds?: number[],
  options: { previousPages?: PreviousPagesByUrl } = {},
): Promise<SiteInventory> {
  const start = Date.now();
  const previousPages = options.previousPages;
  try {
    const items = await fetchSitePages(sitePath, pageIds);
    const pages = items.map((it) => {
      if (previousPages) {
        const prev = previousPages[it.FileRef ?? ""];
        if (prev && it.Modified && prev.modified === it.Modified) {
          return {
            pageId: it.Id,
            pageTitle: it.Title ?? it.FileLeafRef ?? `#${it.Id}`,
            pageUrl: it.FileRef ?? "",
            modified: it.Modified,
            etag: it.etag,
            links: prev.links,
            reusedFromPreviousScan: true,
          } as PageInventory;
        }
      }
      return buildPageInventory(it);
    });
    const totals = tallyPages(pages);
    return {
      site: sitePath,
      pageCount: pages.length,
      pages,
      scanMs: Date.now() - start,
      ...totals,
    };
  } catch (err) {
    return {
      site: sitePath,
      pageCount: 0,
      linkCount: 0,
      byClass: {},
      bySource: {},
      pages: [],
      scanMs: Date.now() - start,
      error: (err as Error).message,
    };
  }
}
