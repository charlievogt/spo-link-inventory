/**
 * Client-side fallback for `suggestion` and `canonicalKey` on link
 * inventory rows.
 *
 * Older scan jobs (anything from before the server-side suggestion
 * feature shipped) don't have these fields populated in the results
 * blob. To keep the "Fix this link", "Align all to canonical", and
 * sibling-grouping features working without forcing a fresh tenant
 * scan, we compute them here on the SPFx side at results-load time.
 *
 * Newer scans already populate these fields server-side via
 * `linkInventoryScanner.ts`'s `suggestReplacement` and the canonical
 * key computation. When present in the blob, the server values are
 * preferred over the client fallback.
 *
 * Both functions are pure string manipulation — no network calls,
 * no DOM, no Fluent UI dependency — so they can run cheaply on
 * thousands of links during a results render.
 */

/**
 * Tenant origin used when building canonical SPO URLs from server-relative
 * paths. In SPFx the page is always served from the SPO origin, so
 * `window.location.origin` resolves to e.g. `https://contoso.sharepoint.com`
 * regardless of which site the web part is running on.
 */
export const TENANT_ORIGIN: string =
  typeof window !== 'undefined' && window.location ? window.location.origin : '';

/**
 * Convert a server-relative file path into the canonical AllItems URL.
 * Mirrors the server-side `buildAllItemsUrl` exactly so client and
 * server fallbacks produce the same string.
 */
function buildAllItemsUrl(serverRelativePath: string): string {
  const path = serverRelativePath.split(/[?#]/, 1)[0];
  if (!path || !path.startsWith('/')) return '';
  if (/\/forms\/(allitems|dispform)\.aspx$/i.test(path)) return '';
  const segments = path.split('/').filter((s) => s.length > 0);
  if (segments.length < 3) return '';
  // Reject SP system folders — `_layouts`, `_api`, `_vti_bin`,
  // `_catalogs` etc. are NOT document libraries. Without this check,
  // a URL like `/sites/<x>/_layouts/15/viewer.aspx` produced a fake
  // canonical AllItems URL pointing at `_layouts` which 404s.
  if (segments.some((s) => s.startsWith('_'))) return '';
  const last = segments[segments.length - 1];
  if (!/\.[a-z0-9]{1,8}$/i.test(last)) return ''; // no extension → folder
  // Reject web page extensions — `.aspx`, `.html`, etc. are pages,
  // not documents, and the AllItems library viewer won't open them.
  if (/\.(aspx|asmx|asp|html|htm|php|jsp|master|mht|mhtml)$/i.test(last)) return '';
  const encPath = '/' + segments.map((s) => encodeURIComponent(s)).join('/');
  const encLib = '/' + segments.slice(0, 3).map((s) => encodeURIComponent(s)).join('/');
  return `${TENANT_ORIGIN}${encLib}/Forms/AllItems.aspx?id=${encPath}&parent=${encLib}`;
}

/**
 * Suggest a canonical replacement URL for a sharing-link / Office
 * Online wrapper / direct SPO document URL. Returns empty string when
 * no suggestion is possible (already canonical, points at a folder
 * not a file, opaque token, etc.).
 *
 * Handles three input shapes:
 *   1. `office-online` — `/:w:/r/`, `/:x:/r/`, `/:p:/r/`, `/:o:/r/`
 *      viewer wrappers. Strip the wrapper, decode the path, build
 *      the AllItems URL.
 *   2. `sharing-link` — `/:b:/r/...` "browser-viewer" wrappers,
 *      same approach.
 *   3. `spo-internal` — direct SPO file URLs like
 *      `https://contoso.sharepoint.com/sites/<site>/<lib>/<file>.<ext>`
 *      that aren't yet in AllItems form. Decode the path and build
 *      the canonical AllItems URL. URLs already in `Forms/AllItems.aspx`
 *      form are detected and return empty (already canonical).
 */
export function clientSuggestReplacement(url: string, linkClass: string): string {
  if (!url) return '';
  if (linkClass === 'office-online') {
    const trimmed = url.replace(/^https?:\/\/[^/]+/i, '');
    const m = /^\/:[wxpo]:\/[rs]\/(.+)$/i.exec(trimmed);
    if (m) {
      let path: string;
      try { path = decodeURIComponent('/' + m[1]); } catch { return ''; }
      return buildAllItemsUrl(path);
    }
    return '';
  }
  if (linkClass === 'sharing-link') {
    const trimmed = url.replace(/^https?:\/\/[^/]+/i, '');
    const m = /^\/:b:\/[rs]\/(.+)$/i.exec(trimmed);
    if (m) {
      let path: string;
      try { path = decodeURIComponent('/' + m[1]); } catch { return ''; }
      return buildAllItemsUrl(path);
    }
    return '';
  }
  if (linkClass === 'malformed-spo-link') {
    // Heuristic-only detection (search-fragment) suggests the URL minus
    // its fragment. If the verifier later proves the id= path is 404,
    // the server clears the suggestion; we never produce one for the
    // 404 case client-side because we can't HEAD-check from here.
    const hashIdx = url.indexOf('#');
    if (hashIdx > 0) return url.slice(0, hashIdx);
    return '';
  }

  if (linkClass === 'spo-internal') {
    // Already in AllItems / DispForm format → no suggestion needed.
    if (/\/forms\/(allitems|dispform)\.aspx[?#]/i.test(url)) return '';
    // Strip host + query/fragment, decode the path
    const trimmed = url.replace(/^https?:\/\/[^/]+/i, '').split(/[?#]/, 1)[0];
    if (!trimmed.startsWith('/')) return '';
    let path: string;
    try { path = decodeURIComponent(trimmed); } catch { return ''; }
    return buildAllItemsUrl(path);
  }
  return '';
}

/**
 * Compact URL normalizer matching the subset of `shared/urlNormalizer`
 * we actually use for canonical-key grouping. Lowercases, decodes
 * percent-encoding, strips Office Online wrappers, extracts `?id=`
 * from AllItems, and returns a `host/path`-style key for absolute
 * URLs (or `/path` for relative ones).
 *
 * This isn't a 100% port of the shared normalizer — it's the minimum
 * needed for sibling grouping in the inventory UI. When the server
 * already populated `canonicalKey` we use that instead.
 */
function clientNormalizeUrl(url: string): string {
  if (!url || !url.trim()) return '';
  const trimmed = url.trim();
  // Skip opaque sharing tokens
  if (/^(https?:\/\/[^/]+)?\/:u:\/g\//i.test(trimmed)) return '';
  // Skip doc.aspx GUIDs
  if (/\/_layouts\/15\/doc\.aspx/i.test(trimmed)) return '';

  let parsed: URL;
  let hasHost = false;
  try {
    if (/^https?:\/\//i.test(trimmed)) {
      parsed = new URL(trimmed);
      hasHost = true;
    } else {
      const pathPart = trimmed.startsWith('/') ? trimmed : `/${trimmed}`;
      parsed = new URL(pathPart, 'https://placeholder.local');
    }
  } catch {
    return '';
  }

  let pathname = parsed.pathname.toLowerCase();
  const host = hasHost ? parsed.hostname.toLowerCase() : '';

  try { pathname = decodeURIComponent(pathname); } catch { /* ignore */ }

  // Strip Office Online viewer wrappers
  pathname = pathname.replace(/^\/:[wxpoub]:(\/[rs])?\//i, '/');

  // Extract ?id= from AllItems / DispForm
  if (/\/forms\/(allitems|dispform)\.aspx$/i.test(pathname)) {
    const idParam = parsed.searchParams.get('id');
    if (idParam) {
      const innerKey = clientNormalizeUrl(idParam);
      if (!innerKey) return '';
      if (host && !innerKey.includes('/')) return host + '/' + innerKey;
      if (host && innerKey.startsWith('/')) return host + innerKey;
      return innerKey;
    }
    return '';
  }

  if (pathname.length > 1 && pathname.endsWith('/')) {
    pathname = pathname.slice(0, -1);
  }

  return host ? host + pathname : pathname;
}

/**
 * Compute the canonical equivalence key for a link given its url,
 * link class, and (optional) suggestion. Same precedence rules as
 * the server.
 */
/**
 * Heuristic-only malformed-link detection, mirroring the server-side
 * `hasAllItemsSearchFragment` check. Returns the upgraded class/reason
 * when the URL is an AllItems.aspx link carrying a leftover `#search=`
 * fragment; returns undefined otherwise.
 *
 * Older scan blobs predate the server-side heuristic — applying this
 * at results-load time means the malformed class surfaces in the UI
 * for them too, without requiring a rescan. Does NOT handle the
 * `file-not-found` reason (that requires a live SP REST check — only
 * the server's opt-in verifier produces it).
 */
export function clientMalformedCheck(url: string): { linkClass: 'malformed-spo-link'; malformedReason: 'search-fragment' } | undefined {
  if (!url) return undefined;
  if (!/\/forms\/allitems\.aspx\?/i.test(url)) return undefined;
  const hashIdx = url.indexOf('#');
  if (hashIdx < 0) return undefined;
  const fragment = url.slice(hashIdx + 1);
  if (!/(^|&)search=/i.test(fragment)) return undefined;
  return { linkClass: 'malformed-spo-link', malformedReason: 'search-fragment' };
}

export function clientCanonicalKey(url: string, linkClass: string, suggestion: string): string {
  // Try the URL itself first — covers spo-internal/relative/onprem/office-online
  if (
    linkClass === 'spo-internal' ||
    linkClass === 'onprem' ||
    linkClass === 'relative' ||
    linkClass === 'office-online'
  ) {
    const k = clientNormalizeUrl(url);
    if (k) return k;
  }
  // Fall back to normalizing the suggestion (sharing-link case)
  if (suggestion) {
    const k = clientNormalizeUrl(suggestion);
    if (k) return k;
  }
  return '';
}
