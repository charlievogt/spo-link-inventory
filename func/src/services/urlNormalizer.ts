const DUMMY_ORIGIN = "https://placeholder.local";

/** Office Online path prefixes: /:w:/r/, /:x:/r/, /:p:/r/, /:o:/r/ */
const OFFICE_ONLINE_RE = /^\/:([wxpo]):(\/[rs])?\//i;

/** _layouts/15/Doc.aspx uses sourcedoc GUID — not path-matchable */
const DOC_ASPX_RE = /\/_layouts\/15\/doc\.aspx$/i;

/** AllItems.aspx or DispForm.aspx — extract ?id= param */
const FORMS_ASPX_RE = /\/forms\/(allitems|dispform)\.aspx$/i;

/** Sharing links with opaque tokens — not matchable */
const SHARING_LINK_RE = /^\/:u:\/g\//i;

/**
 * Normalize a SharePoint URL to a canonical lookup key for comparison.
 *
 * Output format:
 *   - When input has a hostname: "host/path" (e.g. "contoso.sharepoint.com/sites/hr/handbook.docx")
 *   - When input is a relative path: "/path" (no host prefix)
 *
 * Steps:
 * 1. Parse URL (prepend dummy origin for relative paths)
 * 2. Lowercase pathname and hostname
 * 3. Decode URI components
 * 4. Strip Office Online prefixes (/:w:/r/, /:x:/r/, etc.)
 * 5. Extract ?id= from AllItems.aspx / DispForm.aspx
 * 6. Skip _layouts/15/Doc.aspx (GUID-based, not path-matchable)
 * 7. Remove trailing slash (except root)
 * 8. Return "host/path" when host is present, "/path" otherwise
 *
 * Tenant-agnostic: the patterns handled here (Office Online wrappers, SP forms,
 * sharing link shapes) are stable across all SharePoint Online tenants.
 */
export function normalizeUrl(url: string): string {
  if (!url || !url.trim()) return "";

  const trimmed = url.trim();

  // Skip sharing links — opaque tokens, not matchable
  if (SHARING_LINK_RE.test(trimmed)) return "";

  let parsed: URL;
  let hasExplicitHost = false;
  try {
    if (/^https?:\/\//i.test(trimmed)) {
      parsed = new URL(trimmed);
      hasExplicitHost = true;
    } else {
      const pathPart = trimmed.startsWith("/") ? trimmed : `/${trimmed}`;
      parsed = new URL(pathPart, DUMMY_ORIGIN);
    }
  } catch {
    return "";
  }

  let pathname = parsed.pathname.toLowerCase();
  const host = hasExplicitHost ? parsed.hostname.toLowerCase() : "";

  try {
    pathname = decodeURIComponent(pathname);
  } catch {
    // Already decoded or malformed — use as-is
  }

  if (DOC_ASPX_RE.test(pathname)) return "";

  pathname = pathname.replace(OFFICE_ONLINE_RE, "/");

  if (FORMS_ASPX_RE.test(pathname)) {
    const idParam = parsed.searchParams.get("id");
    if (idParam) {
      const innerKey = normalizeUrl(idParam);
      if (!innerKey) return "";
      if (host && !innerKey.includes("/")) return host + "/" + innerKey;
      if (host && innerKey.startsWith("/")) return host + innerKey;
      return innerKey;
    }
    return "";
  }

  if (pathname.length > 1 && pathname.endsWith("/")) {
    pathname = pathname.slice(0, -1);
  }

  if (host) {
    return host + pathname;
  }
  return pathname;
}
