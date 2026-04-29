import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import JSZip from "jszip";
import { PDFDocument, PDFName, PDFString } from "pdf-lib";
import {
  detectFileType,
  extractLinksFromDocument,
  type ExtractedDocLink,
} from "./documentLinkExtractor.js";

/**
 * Tests for `documentLinkExtractor`.
 *
 * These build their fixtures in-memory rather than loading real Office
 * files from disk, for three reasons:
 *
 *   1. Repository hygiene — no binary fixtures in git
 *   2. Determinism — JSZip/pdf-lib produce byte-stable output for
 *      identical inputs, so these tests don't flake
 *   3. Clarity — the fixture *is* the assertion; seeing the test build
 *      a rels file with a specific hyperlink and then seeing the
 *      extractor pull that exact URL out proves the round trip
 *
 * The OOXML fixture intentionally includes non-hyperlink relationships
 * (an image rel, for instance) so we assert the filter actually
 * discriminates on relationship type. Similarly the xlsx fixture
 * includes the worksheet XML with a `<hyperlink>` cell anchor so we
 * can assert the cell-address enrichment path works end-to-end.
 */

describe("documentLinkExtractor.detectFileType", () => {
  it("recognizes docx", () => {
    assert.equal(detectFileType("handbook.docx"), "docx");
    assert.equal(detectFileType("HANDBOOK.DOCX"), "docx");
    assert.equal(detectFileType("path/to/handbook.docx"), "docx");
  });

  it("recognizes xlsx", () => {
    assert.equal(detectFileType("budget.xlsx"), "xlsx");
  });

  it("recognizes pptx", () => {
    assert.equal(detectFileType("deck.pptx"), "pptx");
  });

  it("recognizes pdf", () => {
    assert.equal(detectFileType("policy.pdf"), "pdf");
  });

  it("recognizes Word OOXML variants (docm/dotx/dotm)", () => {
    // Macro-enabled and template variants are structurally identical
    // to .docx — same OOXML zip, same hyperlink rels — so they map to
    // the same internal type and reuse the same extraction path.
    assert.equal(detectFileType("memo.docm"), "docx");
    assert.equal(detectFileType("contract-template.dotx"), "docx");
    assert.equal(detectFileType("macro-template.dotm"), "docx");
  });

  it("recognizes Excel OOXML variants (xlsm/xltx/xltm)", () => {
    assert.equal(detectFileType("budget.xlsm"), "xlsx");
    assert.equal(detectFileType("monthly-template.xltx"), "xlsx");
    assert.equal(detectFileType("forecast-template.xltm"), "xlsx");
  });

  it("recognizes PowerPoint OOXML variants (pptm/potx/potm/ppsx/ppsm)", () => {
    assert.equal(detectFileType("deck.pptm"), "pptx");
    assert.equal(detectFileType("brand-template.potx"), "pptx");
    assert.equal(detectFileType("macro-template.potm"), "pptx");
    assert.equal(detectFileType("show.ppsx"), "pptx");
    assert.equal(detectFileType("macro-show.ppsm"), "pptx");
  });

  it("returns unsupported for .txt", () => {
    assert.equal(detectFileType("notes.txt"), "unsupported");
  });

  it("returns unsupported for .aspx", () => {
    assert.equal(detectFileType("Home.aspx"), "unsupported");
  });

  it("returns unsupported for legacy binary office formats", () => {
    // .doc/.xls/.ppt are deliberately not handled — they need a
    // totally different parser and v1 explicitly scopes them out.
    assert.equal(detectFileType("oldHandbook.doc"), "unsupported");
    assert.equal(detectFileType("oldBudget.xls"), "unsupported");
    assert.equal(detectFileType("oldDeck.ppt"), "unsupported");
  });

  it("returns unsupported for a filename with no extension", () => {
    assert.equal(detectFileType("README"), "unsupported");
  });

  it("strips query strings when detecting", () => {
    // SharePoint sometimes appends `?web=1` when building download URLs;
    // the detector should tolerate that if a caller passes the full URL
    // shape.
    assert.equal(detectFileType("handbook.docx?web=1"), "docx");
  });
});

describe("documentLinkExtractor.extractLinksFromDocument — unsupported", () => {
  it("returns fileType=unsupported and empty links for .txt", async () => {
    const result = await extractLinksFromDocument(Buffer.from("hello"), "notes.txt");
    assert.equal(result.fileType, "unsupported");
    assert.deepEqual(result.links, []);
    assert.equal(result.parseError, undefined);
  });

  it("returns empty cleanly for a filename with no extension", async () => {
    const result = await extractLinksFromDocument(Buffer.from("hello"), "README");
    assert.equal(result.fileType, "unsupported");
    assert.deepEqual(result.links, []);
  });
});

/**
 * Build a minimal docx-shaped OOXML archive with a single external
 * hyperlink relationship in `word/_rels/document.xml.rels`. The archive
 * is NOT a fully-valid Word document — we don't populate the real body
 * content or content-type defaults — but it *is* a structurally valid
 * OOXML rels file which is all the extractor cares about.
 *
 * The function also adds a non-hyperlink relationship (an image rel)
 * so we can assert that the extractor correctly filters by Type.
 */
async function buildMinimalDocx(opts: {
  hyperlinkTarget: string;
  /** Raw `Target` attribute value as it appears in the XML (allows
   *  the test to exercise XML entity decoding). Defaults to
   *  `hyperlinkTarget` verbatim. */
  rawHyperlinkTarget?: string;
  /** When set, the body will contain a `<w:hyperlink r:id="rId1"
   *  w:anchor="<value>">` element. Word composes the click-through
   *  URL as `<Target>#<anchor>`, so this lets tests exercise that
   *  composition path. */
  hyperlinkAnchor?: string;
  /** Display text inside the `<w:hyperlink>` block, if any. */
  hyperlinkText?: string;
}): Promise<Buffer> {
  const rawTarget = opts.rawHyperlinkTarget ?? opts.hyperlinkTarget;
  const zip = new JSZip();

  // Root content types — present so the archive looks roughly like a
  // real docx; not strictly required for extraction.
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );

  // Root rels — points at the main document part.
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );

  // Minimal document body. The extractor opens this file to pair
  // rIds with display text and `w:anchor` values; when neither is
  // needed the trivial paragraph below is enough.
  const anchorAttr = opts.hyperlinkAnchor ? ` w:anchor="${opts.hyperlinkAnchor}"` : "";
  const hyperlinkBlock = (opts.hyperlinkAnchor || opts.hyperlinkText)
    ? `<w:hyperlink r:id="rId1"${anchorAttr} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>${opts.hyperlinkText ?? "link"}</w:t></w:r></w:hyperlink>`
    : "";
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>placeholder</w:t></w:r>${hyperlinkBlock}</w:p></w:body>
</w:document>`,
  );

  // The important file. Contains:
  //   - one external hyperlink (should be extracted)
  //   - one image relationship (should be IGNORED — wrong type)
  //   - one internal hyperlink rel with TargetMode=Internal (should be
  //     IGNORED — pathological but we want to prove the filter)
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${rawTarget}" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="should-be-ignored.internal" TargetMode="Internal"/>
</Relationships>`,
  );

  const uint8 = await zip.generateAsync({ type: "uint8array" });
  return Buffer.from(uint8);
}

describe("documentLinkExtractor.extractLinksFromDocument — docx", () => {
  it("extracts a single hyperlink from a minimal docx", async () => {
    const buffer = await buildMinimalDocx({
      hyperlinkTarget: "https://contoso.sharepoint.com/sites/hr/handbook.docx",
    });
    const result = await extractLinksFromDocument(buffer, "handbook.docx");
    assert.equal(result.fileType, "docx");
    assert.equal(result.parseError, undefined);
    assert.equal(result.links.length, 1, "expected exactly one hyperlink");
    const link = result.links[0];
    assert.equal(link.source, "ooxml-relationship");
    assert.equal(link.url, "https://contoso.sharepoint.com/sites/hr/handbook.docx");
    assert.equal(link.rawUrl, "https://contoso.sharepoint.com/sites/hr/handbook.docx");
    // Context should be the owning part, not the rels file itself.
    assert.equal(link.context, "word/document.xml");
  });

  it("decodes XML entities in the hyperlink target", async () => {
    // Real docx files encode `&` as `&amp;` inside rels Target. The
    // extractor must XML-decode or the URL comparison later will miss.
    const buffer = await buildMinimalDocx({
      hyperlinkTarget: "https://example.com/a?x=1&y=2",
      rawHyperlinkTarget: "https://example.com/a?x=1&amp;y=2",
    });
    const result = await extractLinksFromDocument(buffer, "handbook.docx");
    assert.equal(result.links.length, 1);
    assert.equal(result.links[0].url, "https://example.com/a?x=1&y=2");
    // rawUrl keeps the encoded form for byte-accurate write-back later.
    assert.equal(result.links[0].rawUrl, "https://example.com/a?x=1&amp;y=2");
  });

  it("ignores non-hyperlink relationships and internal-mode hyperlinks", async () => {
    // buildMinimalDocx adds one image rel and one Internal-mode
    // hyperlink alongside the real hyperlink. Exactly one link should
    // come back.
    const buffer = await buildMinimalDocx({ hyperlinkTarget: "https://example.com/" });
    const result = await extractLinksFromDocument(buffer, "handbook.docx");
    assert.equal(result.links.length, 1);
  });

  it("combines w:anchor into the URL as a fragment when the hyperlink carries one", async () => {
    // Real-world pattern from the upstream's OTM docs: the URL stored in rels
    // is already a clean AllItems URL, but the <w:hyperlink> element
    // in the body carries `w:anchor="search=mrm"`. Word composes the
    // click-through URL as `<Target>#<anchor>`, and SharePoint then
    // reinterprets the `#search=` fragment and breaks the link. The
    // extractor must stitch the anchor back on so the classifier sees
    // the effective malformed URL instead of the clean Target.
    const buffer = await buildMinimalDocx({
      hyperlinkTarget:
        "https://contoso.sharepoint.com/sites/online-training-manual/Shared%20Documents/Forms/AllItems.aspx?id=/sites/online-training-manual/Shared%20Documents/MRM.pdf&parent=/sites/online-training-manual/Shared%20Documents",
      rawHyperlinkTarget:
        "https://contoso.sharepoint.com/sites/online-training-manual/Shared%20Documents/Forms/AllItems.aspx?id=/sites/online-training-manual/Shared%20Documents/MRM.pdf&amp;parent=/sites/online-training-manual/Shared%20Documents",
      hyperlinkAnchor: "search=mrm",
      hyperlinkText: "MRM",
    });
    const result = await extractLinksFromDocument(buffer, "handbook.docx");
    assert.equal(result.links.length, 1);
    const link = result.links[0];
    // url carries the composed fragment so the classifier sees it as malformed.
    assert.ok(
      link.url.endsWith("#search=mrm"),
      `expected url to end with #search=mrm, got ${link.url}`,
    );
    // rawUrl stays as the pure rels Target for byte-accurate write-back.
    assert.equal(
      link.rawUrl,
      "https://contoso.sharepoint.com/sites/online-training-manual/Shared%20Documents/Forms/AllItems.aspx?id=/sites/online-training-manual/Shared%20Documents/MRM.pdf&amp;parent=/sites/online-training-manual/Shared%20Documents",
    );
    // Display text comes through from the body.
    assert.equal(link.text, "MRM");
  });

  it("emits the rels Target unchanged when no w:anchor is present", async () => {
    // Regression guard — anchor composition must only kick in when the
    // body actually carried a `w:anchor`. A hyperlink with display text
    // but no anchor stays at the rels Target.
    const buffer = await buildMinimalDocx({
      hyperlinkTarget: "https://example.com/doc.pdf",
      hyperlinkText: "see the doc",
    });
    const result = await extractLinksFromDocument(buffer, "handbook.docx");
    assert.equal(result.links.length, 1);
    assert.equal(result.links[0].url, "https://example.com/doc.pdf");
    assert.equal(result.links[0].text, "see the doc");
  });

  it("returns parseError (not throw) for a corrupt docx", async () => {
    // Byte-level garbage — not a zip at all.
    const result = await extractLinksFromDocument(Buffer.from("not a zip"), "handbook.docx");
    assert.equal(result.fileType, "docx");
    assert.deepEqual(result.links, []);
    assert.ok(result.parseError, "expected parseError to be set");
  });
});

/**
 * Build a minimal xlsx-shaped OOXML archive with a hyperlink rel and a
 * matching `<hyperlink>` element in the worksheet XML. This exercises
 * the rId → cell address pairing logic that enriches the xlsx context
 * hint.
 */
async function buildMinimalXlsx(): Promise<Buffer> {
  const zip = new JSZip();

  // Sheet XML with a hyperlink anchored to cell B3 pointing at rId1.
  // Extractor reads this to build the rId→cell map.
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <hyperlinks>
    <hyperlink ref="B3" r:id="rId1"/>
  </hyperlinks>
</worksheet>`,
  );

  // The sheet's rels — contains the actual URL target.
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://contoso.sharepoint.com/sites/finance/budget.xlsx" TargetMode="External"/>
</Relationships>`,
  );

  const uint8 = await zip.generateAsync({ type: "uint8array" });
  return Buffer.from(uint8);
}

describe("documentLinkExtractor.extractLinksFromDocument — xlsx", () => {
  it("extracts a hyperlink and enriches context with cell address", async () => {
    const buffer = await buildMinimalXlsx();
    const result = await extractLinksFromDocument(buffer, "budget.xlsx");
    assert.equal(result.fileType, "xlsx");
    assert.equal(result.parseError, undefined);
    assert.equal(result.links.length, 1);
    const link = result.links[0];
    assert.equal(link.source, "ooxml-relationship");
    assert.equal(
      link.url,
      "https://contoso.sharepoint.com/sites/finance/budget.xlsx",
    );
    // Context should include the sheet path AND the cell ref.
    assert.equal(link.context, "xl/worksheets/sheet1.xml!B3");
  });
});

/**
 * Build a minimal in-memory PDF with a single `/Link` annotation whose
 * action is a `/URI` pointing at a SharePoint doc. We hand-construct the
 * annotation dict via `pdfDoc.context.obj(...)` rather than any higher-
 * level wrapper because pdf-lib does not expose a typed "add link
 * annotation" helper — and the lower-level approach is exactly what
 * the extractor code needs to read back out anyway.
 */
async function buildMinimalPdf(uri: string): Promise<Buffer> {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([300, 300]);

  // Build the URI action dict: { Type: /Action, S: /URI, URI: (string) }
  const actionDict = pdfDoc.context.obj({
    Type: "Action",
    S: "URI",
  });
  // `context.obj` can't take a literal string without treating it as a
  // Name (e.g. `/URI`), so we set the URI field explicitly using a
  // PDFString.
  actionDict.set(PDFName.of("URI"), PDFString.of(uri));

  // Build the link annotation dict: { Type: /Annot, Subtype: /Link,
  //   Rect: [x1 y1 x2 y2], Border: [0 0 0], A: <actionDict> }
  const annotDict = pdfDoc.context.obj({
    Type: "Annot",
    Subtype: "Link",
    Rect: [50, 50, 250, 100],
    Border: [0, 0, 0],
  });
  annotDict.set(PDFName.of("A"), actionDict);

  // Register the annotation as an indirect object so it gets a ref,
  // then attach to the page's /Annots array via the page leaf helper.
  const annotRef = pdfDoc.context.register(annotDict);
  page.node.addAnnot(annotRef);

  const bytes = await pdfDoc.save();
  return Buffer.from(bytes);
}

describe("documentLinkExtractor.extractLinksFromDocument — pdf", () => {
  it("extracts a /Link annotation's URI", async () => {
    const target = "https://contoso.sharepoint.com/sites/hr/policy.pdf";
    const buffer = await buildMinimalPdf(target);
    const result = await extractLinksFromDocument(buffer, "policy.pdf");
    assert.equal(result.fileType, "pdf");
    assert.equal(result.parseError, undefined);
    assert.equal(result.links.length, 1, "expected exactly one PDF link");
    const link = result.links[0] as ExtractedDocLink;
    assert.equal(link.source, "pdf-annotation");
    assert.equal(link.url, target);
    assert.equal(link.rawUrl, target);
    assert.equal(link.context, "page=1");
  });

  it("returns parseError (not throw) for a corrupt pdf", async () => {
    // Clearly not a PDF — pdf-lib's loader should throw and we should
    // capture it rather than propagate.
    const result = await extractLinksFromDocument(Buffer.from("definitely not a pdf"), "policy.pdf");
    assert.equal(result.fileType, "pdf");
    assert.deepEqual(result.links, []);
    assert.ok(result.parseError, "expected parseError to be set");
  });
});
