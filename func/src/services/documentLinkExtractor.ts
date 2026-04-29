/**
 * Document link extractor.
 *
 * Mirrors the philosophy of `canvasLinkExtractor.ts` — pull every URL out
 * of a binary artifact that lives in a SharePoint document library, with
 * enough source attribution that the orchestrator and future find-and-
 * replace tooling can locate where each link came from.
 *
 * Two file families are supported in v1:
 *
 *   1. Office Open XML (`.docx`, `.xlsx`, `.pptx`)
 *      OOXML files are zip archives. Every clickable external hyperlink
 *      the user inserted ends up in a `<Relationship Type="…/hyperlink"
 *      Target="…"/>` entry in one of the per-part rels files (paths
 *      that match `<dir>/_rels/<name>.xml.rels`). We do NOT need to
 *      parse `document.xml` or the sheet XML to find the URLs — the
 *      rels files are the authoritative list. Parsing just the rels
 *      also sidesteps the headache of OOXML namespaces.
 *
 *      For spreadsheets we *additionally* open the worksheet XML so we
 *      can pair a hyperlink rels `rId` with its cell anchor (`ref="A1"`),
 *      which is invaluable context for the UI ("this link lives in cell
 *      Sheet1!B3"). Slides encode their slide number in the rels path
 *      itself (`ppt/slides/_rels/slide3.xml.rels`), so the rels path
 *      alone is good enough for pptx.
 *
 *   2. PDF
 *      For PDFs we walk page annotation arrays looking for `/Subtype
 *      /Link` dictionaries with a `/A` action of `/S /URI`. That catches
 *      every hyperlink a PDF author inserted via Word → Export, Acrobat,
 *      or any toolchain that follows the PDF spec. We deliberately do
 *      NOT do text extraction in v1 — a URL that's only visible as
 *      body text (typed literally by the author, not a hyperlink) isn't
 *      clickable, so it's not part of the "dead link" risk the redirect
 *      manager cares about. If that turns out to be wrong we'll add text
 *      extraction in a follow-up.
 *
 * All other file types (`.doc`/`.xls`/`.ppt` binary formats, `.one`,
 * `.txt`, `.html`, etc.) return `{ fileType: 'unsupported', links: [] }`
 * without erroring so the orchestrator can just skip them.
 *
 * Input is always a `Buffer` — the caller is expected to have already
 * streamed the file from SharePoint REST (via `getSpoToken()` + the
 * `/_api/web/GetFileByServerRelativeUrl(...)/$value` endpoint). Keeping
 * this module byte-in / link-out means it has no dependency on storage,
 * auth, or file-system layout and stays trivially unit-testable.
 *
 * Errors inside either parser are caught and returned via the
 * `parseError` field rather than thrown — the scanner must be able to
 * keep going past a single corrupt or password-protected document.
 */

import JSZip from "jszip";
import { PDFDocument, PDFArray, PDFDict, PDFName, PDFString, PDFHexString, PDFRef } from "pdf-lib";

/**
 * Supported file families. `unsupported` is the catch-all for anything
 * we don't parse (legacy binary Office formats, text/HTML, OneNote, etc.)
 * and is returned without an error so the orchestrator can cleanly skip
 * the file.
 */
export type FileType = "docx" | "xlsx" | "pptx" | "pdf" | "unsupported";

/**
 * One URL extracted from a document, tagged with enough context for
 * write-back and debugging.
 */
export interface ExtractedDocLink {
  /**
   * The URL exactly as it appeared in the source artifact (still XML-
   * entity-encoded for OOXML, still in its PDFString form for PDFs).
   * Useful when the find-and-replace code eventually needs to patch the
   * byte-for-byte match back out.
   */
  rawUrl: string;
  /**
   * The decoded URL — `&amp;` → `&`, numeric character references
   * resolved, PDF strings converted to JS strings. This is the form you
   * should feed to `normalizeUrl` and the classifier.
   */
  url: string;
  /**
   * Where the URL came from, at a high level. One of:
   *   - `ooxml-relationship` — an external-target hyperlink in a
   *     `.xml.rels` part of an OOXML archive (the common case)
   *   - `ooxml-text` — reserved for future use (e.g. body text URL
   *     scraping). Not produced by v1.
   *   - `pdf-annotation` — a `/Link` annotation with a `/URI` action on
   *     a PDF page
   *   - `pdf-text` — reserved for future use (PDF text stream URL
   *     scraping). Not produced by v1.
   */
  source: "ooxml-relationship" | "ooxml-text" | "pdf-annotation" | "pdf-text";
  /**
   * Free-form context hint about where in the document the link lives.
   * Shape depends on source:
   *   - OOXML docx: `word/document.xml` (or the rels file path if we
   *     couldn't normalize it)
   *   - OOXML xlsx: `xl/worksheets/sheet1.xml!A1` when we were able to
   *     pair the rels entry to a cell anchor; otherwise the rels file
   *     path alone
   *   - OOXML pptx: the slide XML path, e.g. `ppt/slides/slide3.xml`
   *   - PDF: `page=3` (1-based page number where the annotation lives)
   */
  context?: string;
  /**
   * Display text shown to the reader for this hyperlink, when known.
   *
   * Source-specific availability:
   *   - OOXML docx: extracted from `<w:t>` elements inside the
   *     `<w:hyperlink r:id="...">` block in `word/document.xml`
   *   - OOXML pptx: extracted from `<a:t>` siblings of the
   *     `<a:hlinkClick r:id="..."/>` element in slide XML
   *   - OOXML xlsx: read from the `display` attribute on the
   *     worksheet `<hyperlink>` element when present
   *   - PDF: not available — `/Link` annotations don't carry display
   *     text (the visible text is just regular page content beneath
   *     the link's bounding box)
   */
  text?: string;
}

/**
 * Pick a file type from a filename. Case-insensitive, tolerates
 * query strings that SharePoint sometimes appends (e.g. `?web=1`).
 *
 * Recognizes the full Open XML format family for Word / Excel /
 * PowerPoint — workbook, macro-enabled workbook, template, and
 * macro-enabled template variants are all OOXML zip archives with
 * identical hyperlink rels structure, so they map to the same
 * internal type and share the same extraction path.
 *
 *   docx / docm / dotx / dotm           → docx
 *   xlsx / xlsm / xltx / xltm           → xlsx
 *   pptx / pptm / potx / potm / ppsx /  → pptx
 *     ppsm
 *   pdf                                  → pdf
 *
 * Pre-2007 binary formats (`.doc` / `.xls` / `.ppt`) are still
 * `unsupported` — they require a totally different parser.
 */
export function detectFileType(filename: string): FileType {
  if (!filename) return "unsupported";
  // Strip query/fragment in case the caller passed a full URL instead
  // of a bare filename — harmless to be lenient.
  const noQuery = filename.split(/[?#]/, 1)[0];
  const dot = noQuery.lastIndexOf(".");
  if (dot < 0) return "unsupported";
  const ext = noQuery.slice(dot + 1).toLowerCase();
  switch (ext) {
    // Word — workbook, macro-enabled workbook, template, macro-enabled template
    case "docx":
    case "docm":
    case "dotx":
    case "dotm":
      return "docx";
    // Excel — same set of variants
    case "xlsx":
    case "xlsm":
    case "xltx":
    case "xltm":
      return "xlsx";
    // PowerPoint — same plus the slide-show variants
    case "pptx":
    case "pptm":
    case "potx":
    case "potm":
    case "ppsx":
    case "ppsm":
      return "pptx";
    case "pdf":
      return "pdf";
    default:
      return "unsupported";
  }
}

/**
 * OOXML part path containing hyperlink relationships. Hyperlink rels
 * always use this well-known URI as their `Type` attribute — it's part
 * of the OOXML spec (ISO/IEC 29500) and is not a tenant-specific thing.
 */
const HYPERLINK_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/**
 * Decode the small set of XML entities we actually see in rels files.
 * Rels `Target` attributes are written with standard XML escaping —
 * `&amp;` for `&`, numeric refs for anything outside ASCII — and are
 * NOT subjected to the heavy SharePoint canvas encoding we see in
 * `canvasLinkExtractor.ts`. Keep this narrow so we don't accidentally
 * "fix up" URL content that wasn't escaped in the first place.
 */
function decodeXmlEntities(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&"); // MUST be last — otherwise `&amp;lt;` double-decodes
}

/**
 * Parse a single XML attribute out of an opening-tag attribute run. We
 * can't use a full XML parser without pulling in a dependency, but rels
 * files are machine-generated with rigid formatting, so regex is fine.
 *
 * Handles both `attr="value"` and `attr='value'` quoting. Returns
 * `undefined` when the attribute is absent.
 */
function readAttr(attrsChunk: string, name: string): string | undefined {
  // Word boundary on the name so `Target` doesn't match inside
  // `TargetMode`.
  const re = new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`);
  const m = re.exec(attrsChunk);
  if (!m) return undefined;
  return m[1] ?? m[2];
}

/**
 * One entry extracted from a rels file. We keep the rId so we can pair
 * it with worksheet `<hyperlink>` cell-anchor entries during xlsx
 * enrichment.
 */
interface RelsHyperlink {
  /** Relationship id (e.g. `rId4`). Non-empty in well-formed rels. */
  id: string;
  /** Target URL as it appeared in the rels file (still XML-encoded). */
  rawTarget: string;
  /** Target URL after XML entity decoding. */
  target: string;
}

/**
 * Walk a rels XML body and return every external-target hyperlink
 * relationship. Non-hyperlink relationships (images, styles, themes,
 * etc.) are ignored — we only care about things the user clicks.
 *
 * We're defensive about `TargetMode`:
 *   - `External` — treat as URL (this is the normal case)
 *   - missing — treat as URL too (some older Office versions omit the
 *     attribute for hyperlinks; the rels spec says default is `Internal`
 *     but the hyperlink relationship type is meaningless for internals,
 *     so we still take it)
 *   - `Internal` — skip; this would be a relative ref inside the zip
 *     and is never a real hyperlink in practice
 */
function parseRelsHyperlinks(relsXml: string): RelsHyperlink[] {
  const out: RelsHyperlink[] = [];
  // Match every `<Relationship ... />` (self-closing) or
  // `<Relationship ...>...</Relationship>`. Capture the attributes chunk.
  const relRe = /<Relationship\s+([^>]*?)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = relRe.exec(relsXml)) !== null) {
    const attrs = m[1];
    const type = readAttr(attrs, "Type");
    if (type !== HYPERLINK_REL_TYPE) continue;
    const target = readAttr(attrs, "Target");
    if (!target) continue;
    const mode = readAttr(attrs, "TargetMode");
    if (mode && mode.toLowerCase() === "internal") continue;
    const id = readAttr(attrs, "Id") ?? "";
    out.push({
      id,
      rawTarget: target,
      target: decodeXmlEntities(target),
    });
  }
  return out;
}

/**
 * Given a rels path like `xl/worksheets/_rels/sheet1.xml.rels`, derive
 * the sibling worksheet XML path (`xl/worksheets/sheet1.xml`). Returns
 * `undefined` if the path doesn't fit the `_rels/<name>.xml.rels`
 * pattern.
 */
function siblingPartPath(relsPath: string): string | undefined {
  // Matches: <dir>/_rels/<basename>.rels → <dir>/<basename>
  const m = /^(.*?)_rels\/([^/]+)\.rels$/.exec(relsPath);
  if (!m) return undefined;
  return m[1] + m[2];
}

/**
 * Scan an OOXML worksheet XML body for `<hyperlink r:id="rIdN"
 * ref="A1"/>` elements and return a map of rId → cell anchor. Used to
 * enrich xlsx hyperlinks with their cell coordinates.
 *
 * The Excel OOXML schema puts these inside `<hyperlinks>` under the
 * worksheet root. They reference rels via the `r:id` attribute (the
 * `r:` prefix is a namespace alias for the relationships namespace),
 * with the actual cell coordinate in `ref`.
 */
function parseWorksheetHyperlinkRefs(sheetXml: string): Map<string, string> {
  const map = new Map<string, string>();
  // Match `<hyperlink ... />` regardless of attribute ordering. We look
  // for both `r:id` and `ref` on each one.
  const hRe = /<hyperlink\s+([^>]*?)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = hRe.exec(sheetXml)) !== null) {
    const attrs = m[1];
    // Excel sometimes uses the un-prefixed `id` depending on the
    // namespace declarations; handle both.
    const rid = readAttr(attrs, "r:id") ?? readAttr(attrs, "id");
    const ref = readAttr(attrs, "ref");
    if (rid && ref) map.set(rid, ref);
  }
  return map;
}

/**
 * Extract the optional `display` attribute from each `<hyperlink>`
 * element in a worksheet — that's the human-readable label Excel shows
 * for the link, when the author set it explicitly. Pairs by rId so
 * the caller can match against the rels entries.
 *
 * If the author didn't set `display`, the cell content itself is the
 * label; we don't bother resolving that here because it would require
 * loading sharedStrings.xml and walking row/column maps.
 */
function parseWorksheetHyperlinkDisplays(sheetXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const hRe = /<hyperlink\s+([^>]*?)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = hRe.exec(sheetXml)) !== null) {
    const attrs = m[1];
    const rid = readAttr(attrs, "r:id") ?? readAttr(attrs, "id");
    const display = readAttr(attrs, "display");
    if (rid && display) map.set(rid, decodeXmlEntities(display));
  }
  return map;
}

/**
 * Per-rId metadata extracted from a Word body/header/footer/etc. part.
 *
 * Word stores hyperlinks as `<w:hyperlink r:id="rIdN" [w:anchor="..."]>
 * ... </w:hyperlink>` blocks. We walk every `<w:t>` inside each block
 * to reconstruct the visible display text (runs get split across
 * multiple `<w:t>` elements when there's mid-link formatting), and we
 * also capture the `w:anchor` attribute when present.
 *
 * The anchor matters because Word composes the click-through URL as
 * `<relsTarget>#<w:anchor>`. When the anchor value is URL-fragment
 * junk like `search=mrm` (pattern seen in the upstream's OTM docs — likely
 * an old SPO UI bug that baked library-search state into the element
 * instead of the URL), the effective URL differs from what's stored
 * in the rels file. SharePoint's library view then reinterprets the
 * `#search=` fragment client-side and breaks the link at click time.
 * Scanning rels alone misses this entirely, so the inventory must
 * stitch anchor back onto the URL before classification.
 */
interface DocxHyperlinkMetadata {
  /** Display text concatenated from all `<w:t>` runs inside the link. */
  texts: Map<string, string>;
  /** `w:anchor` attribute value (XML-decoded) when present. */
  anchors: Map<string, string>;
}

function parseDocxHyperlinkMetadata(xml: string): DocxHyperlinkMetadata {
  const texts = new Map<string, string>();
  const anchors = new Map<string, string>();
  // Match <w:hyperlink ... [r:id="..."] [w:anchor="..."]> ... </w:hyperlink>
  const re = /<w:hyperlink\b([^>]*)>([\s\S]*?)<\/w:hyperlink>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const attrs = m[1];
    const inner = m[2];
    const rid = readAttr(attrs, "r:id") ?? readAttr(attrs, "id");
    if (!rid) continue;

    const runs: string[] = [];
    const tRe = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRe.exec(inner)) !== null) runs.push(decodeXmlEntities(tm[1]));
    const joined = runs.join("").trim();
    if (joined) texts.set(rid, joined);

    const anchor = readAttr(attrs, "w:anchor");
    if (anchor) anchors.set(rid, decodeXmlEntities(anchor));
  }
  return { texts, anchors };
}

/**
 * Extract hyperlink display text from a pptx slide. PowerPoint stores
 * hyperlinks differently from Word: the link marker
 * `<a:hlinkClick r:id="rIdN"/>` lives INSIDE the run properties
 * `<a:rPr>`, and the visible text is a sibling `<a:t>` element of the
 * same `<a:r>` parent. So we walk every `<a:r>` block, look for an
 * `hlinkClick` inside, then extract any `<a:t>` content within the
 * same block.
 */
function parsePptxHyperlinkTexts(slideXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const runRe = /<a:r\b[^>]*>([\s\S]*?)<\/a:r>/g;
  let rm: RegExpExecArray | null;
  while ((rm = runRe.exec(slideXml)) !== null) {
    const inner = rm[1];
    const linkM = /<a:hlinkClick\s+([^>]*?)\/?>/.exec(inner);
    if (!linkM) continue;
    const rid = readAttr(linkM[1], "r:id") ?? readAttr(linkM[1], "id");
    if (!rid) continue;
    const texts: string[] = [];
    const tRe = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRe.exec(inner)) !== null) {
      texts.push(decodeXmlEntities(tm[1]));
    }
    const joined = texts.join("").trim();
    if (joined) map.set(rid, joined);
  }
  return map;
}

/**
 * Extract every clickable hyperlink from an OOXML archive (docx, xlsx,
 * pptx). Walks every per-part rels file (paths matching
 * `<dir>/_rels/<name>.xml.rels`), filters for the hyperlink relationship
 * type, and returns the decoded targets.
 *
 * For xlsx we also load each `xl/worksheets/sheetN.xml` sibling so we
 * can pair rels entries with their cell anchors; this makes the
 * `context` field substantially more useful ("Sheet1!B3" instead of
 * just "a sheet somewhere").
 */
async function extractOoxmlLinks(buffer: Buffer): Promise<ExtractedDocLink[]> {
  const zip = await JSZip.loadAsync(buffer);
  const out: ExtractedDocLink[] = [];

  // First pass: discover every rels file in the archive. We walk
  // `zip.files` directly rather than using `file(RegExp)` so we can keep
  // track of which rels belong to which part for the xlsx enrichment
  // step below.
  const relsPaths: string[] = [];
  for (const path of Object.keys(zip.files)) {
    if (/(^|\/)_rels\/[^/]+\.rels$/.test(path) && !zip.files[path].dir) {
      relsPaths.push(path);
    }
  }

  // Cache of sheet rId → cell ref. Keyed by sheet XML path. We build
  // these lazily: only if a rels file turns out to reference a sheet and
  // have at least one hyperlink do we pay the cost of reading+parsing
  // the sheet XML.
  const sheetRidCache = new Map<string, Map<string, string>>();
  // Lazy cache of owner path → rId → display text. One entry per
  // body part (docx document.xml / pptx slide / xlsx sheet) we've
  // already parsed. Empty string is a sentinel for "we tried but
  // there's nothing" so we don't reparse.
  const ridToTextCache = new Map<string, Map<string, string>>();
  // Lazy cache of owner path → rId → w:anchor value. Only populated
  // for Word parts (docx body/header/footer/footnotes/etc.) since the
  // `<relsTarget>#<anchor>` URL composition is a WordprocessingML-
  // specific behavior. PowerPoint and Excel hyperlinks don't compose
  // this way.
  const ridToAnchorCache = new Map<string, Map<string, string>>();

  for (const relsPath of relsPaths) {
    const file = zip.file(relsPath);
    if (!file) continue;
    let xml: string;
    try {
      xml = await file.async("text");
    } catch {
      // Rels file unreadable — skip it, keep going with the rest.
      continue;
    }
    const hyperlinks = parseRelsHyperlinks(xml);
    if (hyperlinks.length === 0) continue;

    // Derive the owning part path (the thing this rels file describes).
    // For `word/_rels/document.xml.rels` the owning part is
    // `word/document.xml`. This is what we want to report as context
    // for docx and pptx, and what we use to load the sheet XML for
    // xlsx.
    const ownerPath = siblingPartPath(relsPath);
    const isSheetRels =
      ownerPath !== undefined && /^xl\/worksheets\/sheet\d+\.xml$/.test(ownerPath);
    // Any Word part that can host <w:hyperlink> elements: the body,
    // any header/footer per section, footnotes, endnotes, comments,
    // and the glossary. Matching them all lets us pick up `w:anchor`
    // values from links that live anywhere in the doc, not just the
    // body — which matters for the fragment-composition fix below.
    const isWordPart =
      ownerPath !== undefined &&
      /^word\/(?:document\.xml|header\d+\.xml|footer\d+\.xml|footnotes\.xml|endnotes\.xml|comments\.xml|glossary\/document\.xml)$/.test(
        ownerPath,
      );
    const isDocxBodyRels = ownerPath === "word/document.xml";
    const isPptxSlideRels =
      ownerPath !== undefined && /^ppt\/slides\/slide\d+\.xml$/.test(ownerPath);

    // For sheet rels, pre-build the rId → cellRef map (lazy, cached).
    let ridToCell: Map<string, string> | undefined;
    if (isSheetRels && ownerPath) {
      if (sheetRidCache.has(ownerPath)) {
        ridToCell = sheetRidCache.get(ownerPath);
      } else {
        const sheetFile = zip.file(ownerPath);
        if (sheetFile) {
          try {
            const sheetXml = await sheetFile.async("text");
            const map = parseWorksheetHyperlinkRefs(sheetXml);
            sheetRidCache.set(ownerPath, map);
            ridToCell = map;
            // Also pull the optional `display` attribute as link text.
            if (!ridToTextCache.has(ownerPath)) {
              ridToTextCache.set(ownerPath, parseWorksheetHyperlinkDisplays(sheetXml));
            }
          } catch {
            // Parsing failed — fall back to no cell enrichment, but
            // still emit the raw hyperlinks. Better partial context than
            // dropped links.
            ridToCell = undefined;
          }
        }
      }
    }

    // For any Word part or pptx slide, lazy-load and parse the body XML
    // to get rId → display text mappings. Word parts also get rId →
    // w:anchor so we can compose the effective click-through URL.
    // Cached so multiple rels files for the same body don't reparse.
    if ((isWordPart || isPptxSlideRels) && ownerPath && !ridToTextCache.has(ownerPath)) {
      const bodyFile = zip.file(ownerPath);
      if (bodyFile) {
        try {
          const bodyXml = await bodyFile.async("text");
          if (isWordPart) {
            const { texts, anchors } = parseDocxHyperlinkMetadata(bodyXml);
            ridToTextCache.set(ownerPath, texts);
            ridToAnchorCache.set(ownerPath, anchors);
          } else {
            ridToTextCache.set(ownerPath, parsePptxHyperlinkTexts(bodyXml));
          }
        } catch {
          // Parse failed — set empty maps so we don't retry.
          ridToTextCache.set(ownerPath, new Map());
          if (isWordPart) ridToAnchorCache.set(ownerPath, new Map());
        }
      }
    }
    const ridToText = ownerPath ? ridToTextCache.get(ownerPath) : undefined;
    const ridToAnchor = ownerPath ? ridToAnchorCache.get(ownerPath) : undefined;

    for (const h of hyperlinks) {
      let context: string | undefined;
      if (isSheetRels && ownerPath && ridToCell) {
        const cell = ridToCell.get(h.id);
        context = cell ? `${ownerPath}!${cell}` : ownerPath;
      } else if (ownerPath) {
        context = ownerPath;
      } else {
        context = relsPath;
      }

      // Word-only: if the <w:hyperlink> that referenced this rId carried
      // a `w:anchor` attribute, Word composes the click-through URL as
      // `<relsTarget>#<anchor>`. Stitch that back on so the classifier
      // sees the URL the user actually navigates to — crucial for
      // picking up the `#search=` malformed pattern. `rawUrl` stays as
      // the pure rels attribute for find-and-replace byte matching.
      const anchor = ridToAnchor?.get(h.id);
      const url = anchor ? `${h.target}#${anchor}` : h.target;

      out.push({
        rawUrl: h.rawTarget,
        url,
        source: "ooxml-relationship",
        context,
        text: ridToText?.get(h.id),
      });
    }
  }

  return out;
}

/**
 * PDF `/Link` annotation action subtype name. Per PDF spec §12.6.4.7, a
 * URI action has `/S /URI` and a `/URI` string entry with the target.
 */
const PDF_NAME_LINK = "Link";
const PDF_NAME_URI = "URI";

/**
 * Turn a PDFString / PDFHexString from the `/URI` field into a plain JS
 * string. Both types expose `decodeText()` which handles PDF's two
 * string encodings (PDFDocEncoding vs UTF-16BE with BOM).
 *
 * Returns the empty string if the value isn't one of those two types.
 */
function pdfStringToJs(obj: unknown): string {
  if (obj instanceof PDFString) return obj.decodeText();
  if (obj instanceof PDFHexString) return obj.decodeText();
  return "";
}

/**
 * Resolve a PDF object reference through the context. pdf-lib's page
 * `Annots()` returns a `PDFArray` whose entries are either direct
 * objects or `PDFRef`s that need to be looked up. This helper unifies
 * both cases and returns a `PDFDict` or `undefined`.
 */
function resolveDict(pdfDoc: PDFDocument, obj: unknown): PDFDict | undefined {
  if (obj instanceof PDFDict) return obj;
  if (obj instanceof PDFRef) {
    const resolved = pdfDoc.context.lookup(obj);
    if (resolved instanceof PDFDict) return resolved;
  }
  return undefined;
}

/**
 * Extract every `/Link` annotation URL from a PDF buffer, paired with
 * the visible page text underneath each link's bounding rectangle.
 *
 * We use pdfjs-dist (Mozilla's PDF.js) instead of pdf-lib because PDF
 * link annotations don't carry display text — the visible label is
 * just regular text in the page content stream that happens to lie
 * within the link's `/Rect` bounding box. pdfjs-dist exposes both:
 *   1. Page annotations with `subtype: 'Link'` and a parsed `url`
 *   2. Page text content as positioned items
 * So we can intersect each link's rect with the text items and
 * concatenate the matches as the link's display text.
 *
 * pdf-lib stays in dependencies for now because the rest of the
 * codebase uses it, but the link extraction path is fully on pdfjs.
 *
 * Iteration order: pages in order, annotations within each page
 * roughly top-down, so the emitted list reflects reading order.
 */
async function extractPdfLinks(buffer: Buffer): Promise<ExtractedDocLink[]> {
  // pdfjs-dist legacy ESM build runs in Node. Lazy-import so the
  // module is only loaded for PDF files (skipping the load on docx
  // batches keeps cold-start fast).
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");

  const data = new Uint8Array(buffer);
  const loadingTask = (pdfjs as { getDocument: (opts: unknown) => { promise: Promise<unknown> } }).getDocument({
    data,
    // Disable worker — we're in Node, no off-main-thread benefit and
    // the worker setup is one more thing that can fail.
    disableWorker: true,
    // Defensive: don't throw on minor structural issues.
    stopAtErrors: false,
    // Suppress XFA / font / image fetching — we only want annotations
    // and text content, no rendering.
    isEvalSupported: false,
    useSystemFonts: false,
  });

  let pdfDoc: PdfjsDocument;
  try {
    pdfDoc = (await loadingTask.promise) as PdfjsDocument;
  } catch (err) {
    const msg = (err as Error).message ?? "";
    const name = (err as Error & { name?: string }).name ?? "";
    // Encrypted / password-protected PDFs are common in HR / legal
    // libraries and can't be read without the password. Treat them
    // like a clean "no readable content" result rather than failing
    // the file — we still want the rest of the doc scan to proceed
    // and we don't want them filling up the errors panel.
    if (
      name === "PasswordException" ||
      /password|encrypted/i.test(msg)
    ) {
      // eslint-disable-next-line no-console
      console.warn(`[extractPdfLinks] skipping encrypted PDF: ${msg.slice(0, 120)}`);
      return [];
    }
    throw new Error(`PDF parse failed: ${msg}`);
  }

  const out: ExtractedDocLink[] = [];
  const numPages = pdfDoc.numPages;
  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    let page: PdfjsPage;
    try {
      page = await pdfDoc.getPage(pageNum);
    } catch {
      // Skip pages we can't load — keep going so we don't lose the rest.
      continue;
    }
    const annots: PdfjsAnnotation[] = await page.getAnnotations();
    const linkAnnots = annots.filter(
      (a) => a && a.subtype === "Link" && typeof a.url === "string" && a.url.length > 0,
    );
    if (linkAnnots.length === 0) {
      // pdfjs caches text content per page. Free it explicitly so we
      // don't accumulate memory across hundreds of pages.
      page.cleanup?.();
      continue;
    }

    // Pull text content for the page. Each item carries a 6-element
    // transform matrix `[a, b, c, d, e, f]` — for our purposes only
    // the translation `(e, f)` (the lower-left of the glyph baseline
    // in page user-space) and `width`/`height` matter.
    let textItems: PdfjsTextItem[] = [];
    try {
      const textContent = await page.getTextContent();
      textItems = (textContent.items as PdfjsTextItem[]).filter((it) => typeof it.str === "string");
    } catch {
      // Text extraction failed — emit links with empty text rather
      // than dropping them entirely.
      textItems = [];
    }

    for (const annot of linkAnnots) {
      const rect = annot.rect;
      let text = "";
      if (Array.isArray(rect) && rect.length === 4 && textItems.length > 0) {
        text = collectTextInRect(textItems, rect);
      }
      out.push({
        rawUrl: annot.url!,
        url: annot.url!,
        source: "pdf-annotation",
        context: `page=${pageNum}`,
        text: text || undefined,
      });
    }
    page.cleanup?.();
  }
  await pdfDoc.cleanup?.();
  await pdfDoc.destroy?.();

  // Silence unused-import warnings — pdf-lib symbols are kept for
  // future PDF write operations even though the read path is now on pdfjs.
  void PDFArray;
  void PDFDocument;
  void PDFDict;
  void PDFName;
  void PDFRef;
  void PDFString;
  void PDFHexString;
  void resolveDict;
  void pdfStringToJs;
  void PDF_NAME_LINK;
  void PDF_NAME_URI;

  return out;
}

/** Minimal type stubs for the pdfjs-dist ESM API we use. */
interface PdfjsDocument {
  numPages: number;
  getPage(n: number): Promise<PdfjsPage>;
  cleanup?: () => Promise<void>;
  destroy?: () => Promise<void>;
}
interface PdfjsPage {
  getAnnotations(): Promise<PdfjsAnnotation[]>;
  getTextContent(): Promise<{ items: unknown[] }>;
  cleanup?: () => void;
}
interface PdfjsAnnotation {
  subtype?: string;
  url?: string;
  /** Lower-left and upper-right corners in page user-space units. */
  rect?: number[];
}
interface PdfjsTextItem {
  str: string;
  /** `[a, b, c, d, e, f]` transform — `(e, f)` is the baseline origin. */
  transform: number[];
  width: number;
  height: number;
}

/**
 * Find every text item whose center lies inside the given rect and
 * return the concatenated text. PDF link annotation rects are in
 * page user-space units (lower-left origin, points = 1/72 inch). Text
 * item transforms put the glyph baseline at `(e, f)`, so the visual
 * center is approximately `(e + width/2, f + height/2)`.
 *
 * Using the center point (rather than full bounding-box intersection)
 * is robust against minor PDF rendering quirks where a link rect is
 * slightly tighter or looser than the text it covers.
 */
function collectTextInRect(items: PdfjsTextItem[], rect: number[]): string {
  const [x1, y1, x2, y2] = rect;
  const xMin = Math.min(x1, x2);
  const xMax = Math.max(x1, x2);
  const yMin = Math.min(y1, y2);
  const yMax = Math.max(y1, y2);
  // Add a small fudge to capture text that just brushes the rect.
  const PAD = 1;
  const matched: string[] = [];
  for (const item of items) {
    const [, , , , e, f] = item.transform;
    const cx = e + (item.width || 0) / 2;
    const cy = f + (item.height || 0) / 2;
    if (cx >= xMin - PAD && cx <= xMax + PAD && cy >= yMin - PAD && cy <= yMax + PAD) {
      matched.push(item.str);
    }
  }
  // Some PDFs split a single word across multiple text items with no
  // space; others use a leading space. Join with empty and then
  // collapse runs of whitespace to keep the output readable.
  const joined = matched.join("").replace(/\s+/g, " ").trim();
  return joined;
}

/**
 * Top-level dispatch. Pick a parser based on filename extension, run
 * it, and return a uniform result. Errors inside either parser are
 * caught and reported via `parseError` — the orchestrator must be able
 * to continue after a single bad document.
 *
 * `unsupported` file types return immediately with an empty list and
 * no error. This lets the orchestrator call us unconditionally without
 * pre-filtering.
 */
export async function extractLinksFromDocument(
  buffer: Buffer,
  filename: string,
): Promise<{ fileType: FileType; links: ExtractedDocLink[]; parseError?: string }> {
  const fileType = detectFileType(filename);

  if (fileType === "unsupported") {
    return { fileType, links: [] };
  }

  try {
    if (fileType === "pdf") {
      const links = await extractPdfLinks(buffer);
      return { fileType, links };
    }
    // docx, xlsx, pptx — all OOXML, all handled the same way
    const links = await extractOoxmlLinks(buffer);
    return { fileType, links };
  } catch (err) {
    return {
      fileType,
      links: [],
      parseError: (err as Error).message,
    };
  }
}
