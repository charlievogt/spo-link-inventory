/**
 * Canvas link extractor.
 *
 * Modern SharePoint pages store their content in `CanvasContent1` as a
 * sequence of HTML elements with embedded JSON. Each web part is a
 * `<div data-sp-controldata="<json>">` whose JSON describes the part type
 * and its serverProcessedContent (links, images, text references). Text
 * web parts use `<div data-sp-rte="">` containing arbitrary HTML.
 *
 * This module pulls every URL out of a CanvasContent1 string, classified
 * by source so the find-and-replace UI knows which property to write back
 * to.
 *
 * It is intentionally a hand-rolled parser (no DOM dependency) so it can
 * run inside an Azure Function without bringing in jsdom or cheerio. The
 * regex-based approach is good enough because:
 *   - SP emits `data-sp-controldata` and `data-sp-rte` attributes in a
 *     stable, machine-generated format (not author-edited HTML)
 *   - We only care about extracting URLs, not validating structure
 *   - Anything we miss with a regex shows up in the test output and we
 *     extend the extractor to cover it
 *
 * If we ever need to *write back* changes (find-and-replace), the same
 * sites where we extract a URL are the sites we patch — see notes inline
 * about which JSON path each URL came from.
 */

export type LinkSource =
  | "anchor" // <a href> inside a Text web part
  | "quickLinks" // QuickLinks web part item.url
  | "hero" // Hero web part item.url
  | "image" // Image web part imageSource / linkUrl
  | "button" // Button web part url
  | "fileViewer" // File Viewer web part fileUrl
  | "embed" // Embed web part embedCode src
  | "imageGallery" // Image Gallery web part item.url
  | "callToAction" // CallToAction web part url
  | "unknown"; // Found a URL we couldn't attribute to a known source

export interface ExtractedLink {
  /** The raw URL exactly as it appeared in CanvasContent1 (still HTML-encoded if it was). */
  rawUrl: string;
  /** Decoded URL — &amp; → &, etc. Suitable for normalization. */
  url: string;
  /** Where on the page the URL was found (which web part type / which attribute). */
  source: LinkSource;
  /** Web part instanceId (controldata.id) when known — lets us locate the part for write-back. */
  webPartInstanceId?: string;
  /** Display text (link label) when we can recover it from the JSON. */
  text?: string;
}

/**
 * Decode HTML entities found in CanvasContent1. SharePoint emits the
 * embedded JSON heavily entity-encoded — every `:`, `{`, `}` and most
 * punctuation comes through as numeric character references like `&#58;`,
 * `&#123;`, `&#125;`. We need to handle the full numeric form (decimal
 * and hex) plus the few common named entities.
 */
function decodeEntities(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ");
}

/**
 * Pull every `<a href="...">` out of an HTML chunk (the contents of a
 * Text web part). We capture the href and the inner text so the UI can
 * show "this link labeled 'Click here' on this page points at X".
 */
function extractAnchors(html: string, instanceId: string | undefined): ExtractedLink[] {
  const out: ExtractedLink[] = [];
  // Non-greedy capture of href + inner text. Anchors in SP text web
  // parts are simple (no nested tags inside the anchor in the common case),
  // so this is robust enough.
  const anchorRe = /<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m: RegExpExecArray | null;
  while ((m = anchorRe.exec(html)) !== null) {
    const rawUrl = m[1];
    const innerText = decodeEntities(m[2].replace(/<[^>]+>/g, "")).trim();
    out.push({
      rawUrl,
      url: decodeEntities(rawUrl),
      source: "anchor",
      webPartInstanceId: instanceId,
      text: innerText || undefined,
    });
  }
  return out;
}

/**
 * Walk a parsed JSON object and pull out URL-shaped string values from
 * keys we recognize. We don't blindly grab every string that looks like a
 * URL — too many false positives from analytics IDs, image dimensions,
 * etc. Instead we look at known property names per web part type.
 *
 * Returns links with their `source` set to `unknown` — the caller refines
 * it based on which web part the JSON came from.
 */
const URL_KEYS = new Set([
  "url",
  "href",
  "linkUrl",
  "fileUrl",
  "imageSource",
  "imageUrl",
  "siteUrl",
  "webUrl",
  "navigateUri",
  "src",
]);

/**
 * SP `serverProcessedContent` exposes flat path-keyed dictionaries where
 * the KEYS are like `"images[0].imageSource"` and the VALUES are URLs.
 * The walker can't tell those values are URLs from key name alone, so we
 * special-case the dictionaries we know hold URLs.
 */
const URL_DICT_KEYS = new Set([
  "imageSources",
  "links",
  "audienceTargetLinks",
]);

function walkForUrls(node: unknown, instanceId: string | undefined, out: ExtractedLink[], textHint?: string): void {
  if (node === null || node === undefined) return;
  if (typeof node === "string") return; // strings are only interesting in context
  if (Array.isArray(node)) {
    for (const item of node) walkForUrls(item, instanceId, out);
    return;
  }
  if (typeof node !== "object") return;

  const obj = node as Record<string, unknown>;
  // Surface a `title`, `text`, or `displayText` from the same object as a
  // text hint for any URL keys we find on this object.
  const localText =
    (typeof obj.title === "string" && obj.title) ||
    (typeof obj.text === "string" && obj.text) ||
    (typeof obj.displayText === "string" && obj.displayText) ||
    textHint;

  for (const [key, value] of Object.entries(obj)) {
    if (typeof value === "string" && URL_KEYS.has(key) && value.length > 0 && !value.startsWith("data:")) {
      out.push({
        rawUrl: value,
        url: value, // already decoded — JSON.parse handled it
        source: "unknown",
        webPartInstanceId: instanceId,
        text: typeof localText === "string" ? localText : undefined,
      });
    } else if (URL_DICT_KEYS.has(key) && value && typeof value === "object" && !Array.isArray(value)) {
      // Flat path-keyed dictionary — every string value is a URL.
      for (const [innerKey, innerVal] of Object.entries(value as Record<string, unknown>)) {
        if (typeof innerVal === "string" && innerVal.length > 0 && !innerVal.startsWith("data:")) {
          out.push({
            rawUrl: innerVal,
            url: innerVal,
            source: "unknown",
            webPartInstanceId: instanceId,
            text: innerKey, // path expression like "items[0].url" — useful for write-back
          });
        }
      }
    } else if (typeof value === "object") {
      walkForUrls(value, instanceId, out, typeof localText === "string" ? localText : undefined);
    }
  }
}

/**
 * Map a web part type id (from controldata.webPartId) to a LinkSource.
 * GUIDs are stable across tenants — these are Microsoft's first-party
 * web part IDs.
 */
const WEBPART_ID_TO_SOURCE: Record<string, LinkSource> = {
  "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd": "quickLinks",
  "daf0b71c-6de8-4ef7-b511-faae7c388708": "hero",
  "d1d91016-032f-456d-98a4-721247c305e8": "image",
  "8c88f208-6c77-4bdb-86a0-0c47b4316588": "fileViewer",
  "490d7c76-1824-45b2-9de3-676421c997fa": "embed",
  "af8be689-990e-492a-81f7-ba3e4cd3ed9c": "imageGallery",
  "df8e44e7-edd5-46d5-90da-aca1539313b8": "callToAction",
  "8654b779-4886-46d4-8ffb-b5ed960ee986": "button",
  // Text web part — not a "URL source" itself, but we extract anchors
  // from its inner HTML separately
};

interface CanvasControl {
  controldataJson: string; // data-sp-controldata attr value (position + webPartId)
  webpartdataJson?: string; // data-sp-webpartdata attr value (the real payload)
  rteHtml?: string; // inner HTML of data-sp-rte (Text web parts)
}

/**
 * Pull every attribute value out of the canvas string for a given attr name.
 * SP emits attributes as `name="..."` with HTML-entity-encoded contents,
 * so simple regex extraction is reliable here.
 */
function extractAttr(canvas: string, attr: string): string[] {
  const re = new RegExp(`\\b${attr}=["']([^"']+)["']`, "gi");
  const out: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(canvas)) !== null) out.push(m[1]);
  return out;
}

/**
 * Split a CanvasContent1 string into controls. SP's structure is roughly:
 *   <div data-sp-canvascontrol data-sp-controldata="<position+webPartId>">
 *     <div data-sp-webpart data-sp-webpartdata="<full payload>">
 *       <div data-sp-rte>...inner HTML for Text web parts...</div>
 *     </div>
 *   </div>
 *
 * The interesting JSON for link extraction is in `data-sp-webpartdata`
 * (it has serverProcessedContent + the per-webpart property bag). The
 * controldata is mostly position info — but we still parse it to grab
 * webPartId so we can map URLs back to web part type for find-and-replace.
 *
 * This implementation avoids depending on a balanced-div parse: we walk
 * the canvas and group controldata + the *next* webpartdata together,
 * since SP always emits them in that order with one webpartdata per
 * controldata. Text web parts also have a sibling div with `data-sp-rte`.
 */
function splitControls(canvas: string): CanvasControl[] {
  const controldatas = extractAttr(canvas, "data-sp-controldata");
  const webpartdatas = extractAttr(canvas, "data-sp-webpartdata");
  // RTE inner HTML matched separately (not an attribute value)
  const rteRe = /<div[^>]*\bdata-sp-rte=["'][^"']*["'][^>]*>([\s\S]*?)<\/div>/gi;
  const rtes: string[] = [];
  let rm: RegExpExecArray | null;
  while ((rm = rteRe.exec(canvas)) !== null) rtes.push(rm[1]);

  // Pair controldata to webpartdata by index. Some controls (text RTE)
  // have controldata but no webpartdata; others have both. We loop the
  // longer of the two so nothing gets dropped.
  const max = Math.max(controldatas.length, webpartdatas.length);
  const out: CanvasControl[] = [];
  for (let i = 0; i < max; i++) {
    out.push({
      controldataJson: controldatas[i] ?? "",
      webpartdataJson: webpartdatas[i],
      rteHtml: rtes[i],
    });
  }
  return out;
}

interface ParsedControlData {
  id?: string;
  instanceId?: string;
  webPartId?: string;
  webPartData?: unknown;
  serverProcessedContent?: unknown;
}

/**
 * Extract every URL from a CanvasContent1 HTML string. Order is roughly
 * top-to-bottom of the page but not guaranteed — callers should not
 * depend on it.
 */
export function extractLinksFromCanvas(canvas: string | null | undefined): ExtractedLink[] {
  if (!canvas) return [];
  const out: ExtractedLink[] = [];

  for (const ctrl of splitControls(canvas)) {
    // Parse controldata (position info + webPartId for type lookup)
    let controlMeta: ParsedControlData = {};
    if (ctrl.controldataJson) {
      try {
        controlMeta = JSON.parse(decodeEntities(ctrl.controldataJson)) as ParsedControlData;
      } catch {
        // Malformed — keep going, the webpartdata may still be parseable
      }
    }

    // Parse webpartdata (the real payload — serverProcessedContent etc.)
    let webPartPayload: ParsedControlData = {};
    if (ctrl.webpartdataJson) {
      try {
        webPartPayload = JSON.parse(decodeEntities(ctrl.webpartdataJson)) as ParsedControlData;
      } catch {
        // ignore — fall through to RTE-only extraction
      }
    }

    const instanceId = controlMeta.id ?? webPartPayload.instanceId ?? webPartPayload.id;
    const webPartId = (controlMeta.webPartId ?? webPartPayload.id)?.toLowerCase();
    const knownSource = (webPartId && WEBPART_ID_TO_SOURCE[webPartId]) || "unknown";

    // 1) Anchors inside a Text web part RTE region
    if (ctrl.rteHtml) {
      out.push(...extractAnchors(ctrl.rteHtml, instanceId));
    }

    // 2) URLs inside the web part data payload. Walk the entire payload
    //    so we catch URLs at every nesting depth (Hero items, Quick Links
    //    items, etc. — they all live under serverProcessedContent or
    //    similar sub-objects).
    const collected: ExtractedLink[] = [];
    walkForUrls(webPartPayload, instanceId, collected);

    // Refine source from "unknown" to the web part's known source
    for (const link of collected) {
      if (knownSource !== "unknown") link.source = knownSource;
      out.push(link);
    }
  }

  return out;
}
