import { fetchPageById, patchPageCanvas, publishPage } from "./spoPagesClient.js";
import { getSpoUserWriteToken, type UserPrincipal } from "./linkInventoryAuth.js";

/**
 * Find-and-replace write path for the link inventory.
 *
 * Workflow per page:
 *   1. Re-fetch the page item with current ETag (NOT trusting the
 *      scan-time ETag — we need fresh state)
 *   2. Compare current ETag against the scan-time ETag. If they differ,
 *      the page has been edited since the scan and we refuse with
 *      `staleScan: true`. The user must rescan and retry.
 *   3. Apply the replacement(s) to CanvasContent1 (and optionally
 *      LayoutWebpartsContent) by literal string substitution of the
 *      scan-time `rawUrl` with the freshly re-encoded new URL.
 *   4. PATCH back with `If-Match: <currentEtag>` for belt-and-suspenders
 *      protection against TOCTOU between our re-read and our write. If
 *      SP returns 412, surface as a conflict.
 *   5. On dryRun, skip step 4 entirely and just return the diff.
 *
 * URL encoding for canvas substitution:
 *   - Anchor href values use light HTML entity encoding: `&` `"` `<` `>`
 *   - Web part data JSON inside `data-sp-webpartdata` uses heavy
 *     numeric-entity encoding for `:` `{` `}` in addition to the light
 *     entities. We detect which form by sniffing for `&#58;` in the
 *     scan-time rawUrl and apply the same encoding to the replacement.
 */

export interface LinkReplacement {
  /** Page id within the site (matches scan output) */
  pageId: number;
  /** ETag captured at scan time — used for staleness check */
  scanEtag?: string;
  /** Substring on the canvas to replace (the raw, encoded form from scan) */
  rawUrl: string;
  /** Decoded form of the original URL — for the response/preview */
  oldUrl: string;
  /** Decoded form of the replacement URL */
  newUrl: string;
}

export interface PageReplaceResult {
  pageId: number;
  pageTitle?: string;
  /**
   * Server-relative URL of the page (e.g. `/sites/foo/SitePages/Page.aspx`).
   * The UI uses this to link directly to the patched page in its result
   * row. Constructed from the SitePages list item's `FileLeafRef`.
   */
  pageUrl?: string;
  status: "applied" | "preview" | "stale" | "conflict" | "no-match" | "not-found" | "error";
  /** Number of substring replacements that succeeded on the canvas */
  replacementCount: number;
  /** Number of replacements that didn't find their target rawUrl */
  unmatchedCount: number;
  /** Per-link details */
  details: Array<{
    oldUrl: string;
    newUrl: string;
    matched: boolean;
  }>;
  error?: string;
  /** New ETag after a successful PATCH (real-run only) */
  newEtag?: string;
}

/**
 * Re-encode a decoded URL to match the encoding form of the original
 * `rawUrl` it's replacing. We sniff for `&#58;` to detect heavy form
 * (web part JSON in HTML attribute) vs light form (anchor href).
 */
function encodeForCanvas(decodedUrl: string, originalRaw: string): string {
  const heavy = originalRaw.indexOf("&#58;") !== -1;
  let encoded = decodedUrl
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  if (heavy) {
    encoded = encoded
      .replace(/:/g, "&#58;")
      .replace(/\{/g, "&#123;")
      .replace(/\}/g, "&#125;");
  }
  return encoded;
}

/**
 * Apply a list of replacements to a canvas string. Returns the new
 * canvas plus per-replacement match counts so the caller can report
 * which links were actually found.
 */
export function applyReplacementsToCanvas(
  canvas: string,
  replacements: LinkReplacement[],
): { canvas: string; matchCounts: Map<string, number> } {
  let working = canvas;
  const matchCounts = new Map<string, number>();
  for (const rep of replacements) {
    const before = working;
    // Determine which form of the rawUrl actually appears in the
    // canvas. Links sourced from `<a href>` anchors have rawUrl in
    // the same encoding as the canvas HTML (entity-encoded `&amp;`).
    // But links sourced from JSON inside `data-sp-webpartdata`
    // attributes have rawUrl as a decoded JS string (`&` not
    // `&amp;`), because the extractor parsed the JSON after entity-
    // decoding the attribute. The canvas itself has the JSON double-
    // encoded, so the literal `&` in rawUrl won't be found.
    //
    // Strategy: try the raw form first. If that doesn't match, try
    // the HTML-entity-encoded form (`& → &amp;`). Use whichever
    // one the canvas actually contains. If neither matches, count=0.
    let findStr = rep.rawUrl;
    if (working.indexOf(findStr) === -1) {
      const entityEncoded = rep.rawUrl.replace(/&(?!amp;|lt;|gt;|quot;|apos;|#)/g, "&amp;");
      if (entityEncoded !== rep.rawUrl && working.indexOf(entityEncoded) !== -1) {
        findStr = entityEncoded;
      }
    }

    const replacementEncoded = encodeForCanvas(rep.newUrl, findStr);
    // Skip if the encoded form is identical (no-op replacement)
    if (replacementEncoded === findStr) {
      matchCounts.set(rep.rawUrl, 0);
      continue;
    }
    working = working.split(findStr).join(replacementEncoded);
    // Count occurrences by length delta
    const occurrences = (before.length - working.length) /
      Math.max(1, findStr.length - replacementEncoded.length);
    matchCounts.set(rep.rawUrl, before === working ? 0 : Math.max(1, Math.round(occurrences)));
  }
  return { canvas: working, matchCounts };
}

interface ApplyOptions {
  dryRun: boolean;
  /**
   * Calling user. When provided, all SP REST calls run as the user via
   * delegated `AllSites.Write` OBO. Required for real-run mode; optional
   * for dry-run (we still want to use the user token for the read so the
   * preview reflects the user's actual permissions).
   */
  user?: UserPrincipal;
}

/**
 * Apply a set of replacements to a single page, with concurrency
 * checking. The replacements all apply to the same pageId.
 */
export async function applyReplacementsToPage(
  sitePath: string,
  replacements: LinkReplacement[],
  options: ApplyOptions,
): Promise<PageReplaceResult> {
  if (replacements.length === 0) {
    return {
      pageId: -1,
      status: "no-match",
      replacementCount: 0,
      unmatchedCount: 0,
      details: [],
    };
  }
  const pageId = replacements[0].pageId;
  const scanEtag = replacements[0].scanEtag;

  // Acquire user write token if a user is supplied. Same token is used
  // for both the read-back and the PATCH so SP sees a single identity.
  let userToken: string | undefined;
  if (options.user) {
    try {
      userToken = await getSpoUserWriteToken(options.user);
    } catch (e) {
      return {
        pageId,
        status: "error",
        replacementCount: 0,
        unmatchedCount: replacements.length,
        details: replacements.map((r) => ({ oldUrl: r.oldUrl, newUrl: r.newUrl, matched: false })),
        error: `Failed to get user OBO token for SP write: ${(e as Error).message}`,
      };
    }
  }

  let item;
  try {
    item = await fetchPageById(sitePath, pageId, userToken);
  } catch (e) {
    return {
      pageId,
      status: "error",
      replacementCount: 0,
      unmatchedCount: replacements.length,
      details: replacements.map((r) => ({ oldUrl: r.oldUrl, newUrl: r.newUrl, matched: false })),
      error: (e as Error).message,
    };
  }
  if (!item) {
    return {
      pageId,
      status: "not-found",
      replacementCount: 0,
      unmatchedCount: replacements.length,
      details: replacements.map((r) => ({ oldUrl: r.oldUrl, newUrl: r.newUrl, matched: false })),
      error: `Page ${pageId} not found in ${sitePath}`,
    };
  }

  const currentEtag = item.etag;
  // Stale scan check: if scan-time ETag and current ETag differ, the
  // page was edited after the scan, so the link inventory may be wrong.
  // Refuse the write and ask the user to rescan.
  if (scanEtag && currentEtag && currentEtag !== scanEtag) {
    return {
      pageId,
      pageTitle: item.Title ?? item.FileLeafRef,
      pageUrl: item.FileLeafRef ? `${sitePath}/SitePages/${item.FileLeafRef}` : undefined,
      status: "stale",
      replacementCount: 0,
      unmatchedCount: replacements.length,
      details: replacements.map((r) => ({ oldUrl: r.oldUrl, newUrl: r.newUrl, matched: false })),
      error: `Page modified since scan (scan etag ${scanEtag}, current ${currentEtag}). Please rescan and retry.`,
    };
  }

  // Apply replacements in-memory.
  let canvasResult = applyReplacementsToCanvas(item.CanvasContent1 ?? "", replacements);
  let layoutResult = applyReplacementsToCanvas(item.LayoutWebpartsContent ?? "", replacements);

  // Sum match counts across both fields per replacement key.
  const totalMatchCounts = new Map<string, number>();
  for (const [k, v] of canvasResult.matchCounts) totalMatchCounts.set(k, (totalMatchCounts.get(k) ?? 0) + v);
  for (const [k, v] of layoutResult.matchCounts) totalMatchCounts.set(k, (totalMatchCounts.get(k) ?? 0) + v);

  const details = replacements.map((r) => ({
    oldUrl: r.oldUrl,
    newUrl: r.newUrl,
    matched: (totalMatchCounts.get(r.rawUrl) ?? 0) > 0,
  }));
  const replacementCount = details.filter((d) => d.matched).length;
  const unmatchedCount = details.length - replacementCount;

  if (replacementCount === 0) {
    return {
      pageId,
      pageTitle: item.Title ?? item.FileLeafRef,
      pageUrl: item.FileLeafRef ? `${sitePath}/SitePages/${item.FileLeafRef}` : undefined,
      status: "no-match",
      replacementCount,
      unmatchedCount,
      details,
      error: "Could not find the scan-time URLs in the current page content",
    };
  }

  if (options.dryRun) {
    return {
      pageId,
      pageTitle: item.Title ?? item.FileLeafRef,
      pageUrl: item.FileLeafRef ? `${sitePath}/SitePages/${item.FileLeafRef}` : undefined,
      status: "preview",
      replacementCount,
      unmatchedCount,
      details,
    };
  }

  // Real run — PATCH the page. Use If-Match when we have an ETag for
  // optimistic concurrency; proceed without it when the ETag is missing
  // (some SP responses omit @odata.etag for reasons that aren't fully
  // documented). Blocking the write entirely on a missing ETag is too
  // aggressive — the user already previewed and decided to apply.
  const patchResult = await patchPageCanvas(
    sitePath,
    pageId,
    currentEtag ?? "*",
    canvasResult.canvas,
    layoutResult.canvas !== (item.LayoutWebpartsContent ?? "") ? layoutResult.canvas : undefined,
    userToken,
  );
  if (!patchResult.ok) {
    return {
      pageId,
      pageTitle: item.Title ?? item.FileLeafRef,
      pageUrl: item.FileLeafRef ? `${sitePath}/SitePages/${item.FileLeafRef}` : undefined,
      status: patchResult.conflict ? "conflict" : "error",
      replacementCount: 0,
      unmatchedCount: replacements.length,
      details,
      error: patchResult.error,
    };
  }

  // Auto-publish so the user doesn't have to manually publish every
  // page that was modified. If publishing fails (e.g. the library
  // doesn't have minor versioning, or the user lacks Publish rights),
  // we still report the PATCH as successful — the content was saved,
  // just as a draft.
  const publishResult = await publishPage(sitePath, pageId, userToken);
  // eslint-disable-next-line no-console
  if (!publishResult.ok) console.warn(`[writer] publish ${sitePath}/${pageId} failed: ${publishResult.error}`);

  return {
    pageId,
    pageTitle: item.Title ?? item.FileLeafRef,
    pageUrl: item.FileLeafRef ? `${sitePath}/SitePages/${item.FileLeafRef}` : undefined,
    status: "applied",
    replacementCount,
    unmatchedCount,
    details,
    newEtag: patchResult.newEtag,
  };
}
