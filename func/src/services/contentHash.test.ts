import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import { createHash } from "node:crypto";
import JSZip from "jszip";
import { hashFileContent } from "./contentHash.js";

/**
 * Build a minimal `.docx`-shaped zip in memory. Only the structural parts
 * the OOXML hasher actually consumes; SharePoint and Word both produce
 * far more parts in the real world, but the hasher only cares about
 * which paths are kept vs. dropped.
 *
 * `propertyOverrides` lets a test customize the property-bag parts to
 * simulate SharePoint's property-promotion rewrites between uploads.
 */
async function makeDocx(opts: {
  documentXml?: string;
  styles?: string;
  propertyOverrides?: { coreXml?: string; itemPropsXml?: string };
}): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
  );
  zip.folder("_rels")!.file(
    ".rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  zip
    .folder("word")!
    .file(
      "document.xml",
      opts.documentXml ??
        `<?xml version="1.0"?><w:document xmlns:w="..."><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>`,
    );
  zip.folder("word")!.file("styles.xml", opts.styles ?? `<?xml version="1.0"?><w:styles/>`);
  zip
    .folder("docProps")!
    .file(
      "core.xml",
      opts.propertyOverrides?.coreXml ??
        `<?xml version="1.0"?><cp:coreProperties><dc:creator>alice</dc:creator></cp:coreProperties>`,
    );
  zip
    .folder("customXml")!
    .file(
      "itemProps1.xml",
      opts.propertyOverrides?.itemPropsXml ??
        `<?xml version="1.0"?><ds:datastoreItem ds:itemID="{id}"/>`,
    );
  return zip.generateAsync({ type: "nodebuffer" });
}

describe("contentHash.hashFileContent — OOXML mode", () => {
  it("returns ooxml-content-v1 algo for .docx/.xlsx/.pptx", async () => {
    const docx = await makeDocx({});
    const { algo } = await hashFileContent(docx, "docx");
    assert.equal(algo, "ooxml-content-v1");

    const { algo: xlsxAlgo } = await hashFileContent(docx, "xlsx");
    assert.equal(xlsxAlgo, "ooxml-content-v1");

    const { algo: pptxAlgo } = await hashFileContent(docx, "pptx");
    assert.equal(pptxAlgo, "ooxml-content-v1");
  });

  it("produces identical hashes when only docProps differ", async () => {
    const a = await makeDocx({
      propertyOverrides: { coreXml: `<core>creator=alice,modified=2024</core>` },
    });
    const b = await makeDocx({
      propertyOverrides: { coreXml: `<core>creator=bob,modified=2026</core>` },
    });
    const ha = await hashFileContent(a, "docx");
    const hb = await hashFileContent(b, "docx");
    assert.equal(ha.hash, hb.hash, "docProps changes must not affect the content hash");
  });

  it("produces identical hashes when only customXml differs", async () => {
    const a = await makeDocx({
      propertyOverrides: { itemPropsXml: `<item>library-column-A=foo</item>` },
    });
    const b = await makeDocx({
      propertyOverrides: { itemPropsXml: `<item>library-column-B=bar</item>` },
    });
    const ha = await hashFileContent(a, "docx");
    const hb = await hashFileContent(b, "docx");
    assert.equal(ha.hash, hb.hash, "customXml differences must not affect the content hash");
  });

  it("produces different hashes when document.xml differs", async () => {
    const a = await makeDocx({ documentXml: `<doc>Hello</doc>` });
    const b = await makeDocx({ documentXml: `<doc>Goodbye</doc>` });
    const ha = await hashFileContent(a, "docx");
    const hb = await hashFileContent(b, "docx");
    assert.notEqual(ha.hash, hb.hash, "document.xml differences must produce different hashes");
  });

  it("produces different hashes when a content part is added", async () => {
    const a = await makeDocx({});
    const zipB = await JSZip.loadAsync(await makeDocx({}));
    zipB.folder("word")!.file("footer1.xml", `<?xml version="1.0"?><w:ftr/>`);
    const b = await zipB.generateAsync({ type: "nodebuffer" });
    const ha = await hashFileContent(a, "docx");
    const hb = await hashFileContent(b, "docx");
    assert.notEqual(ha.hash, hb.hash, "adding a content part must change the hash");
  });

  it("falls back to full-sha256 when the OOXML buffer is malformed", async () => {
    const garbage = Buffer.from("not a zip file at all");
    const { algo, hash } = await hashFileContent(garbage, "docx");
    assert.equal(algo, "full-sha256");
    assert.equal(hash, createHash("sha256").update(garbage).digest("hex"));
  });
});

describe("contentHash.hashFileContent — full-sha256 mode", () => {
  it("uses full-sha256 for PDFs", async () => {
    const bytes = Buffer.from([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34]); // "%PDF-1.4"
    const { hash, algo } = await hashFileContent(bytes, "pdf");
    assert.equal(algo, "full-sha256");
    assert.equal(hash, createHash("sha256").update(bytes).digest("hex"));
  });

  it("uses full-sha256 for unknown types", async () => {
    const bytes = Buffer.from("anything");
    const { hash, algo } = await hashFileContent(bytes, "other");
    assert.equal(algo, "full-sha256");
    assert.equal(hash, createHash("sha256").update(bytes).digest("hex"));
  });

  it("two identical PDF buffers hash identically", async () => {
    const bytes = Buffer.from("%PDF-1.4\nidentical content here");
    const a = await hashFileContent(bytes, "pdf");
    const b = await hashFileContent(Buffer.from(bytes), "pdf");
    assert.equal(a.hash, b.hash);
    assert.equal(a.algo, b.algo);
  });
});
