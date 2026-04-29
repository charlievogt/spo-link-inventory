import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import JSZip from "jszip";
import {
  extractText,
  normalizeText,
  textContentHash,
  simhash64,
  simhashDistance,
} from "./textHash.js";

async function makeDocxWithText(documentXml: string): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
  );
  zip.folder("_rels")!.file(
    ".rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  zip.folder("word")!.file("document.xml", documentXml);
  return zip.generateAsync({ type: "nodebuffer" });
}

async function makeXlsxWithSharedStrings(strings: string[], cells: string[]): Promise<Buffer> {
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
    .folder("xl")!
    .file(
      "sharedStrings.xml",
      `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${strings
        .map((s) => `<si><t>${s}</t></si>`)
        .join("")}</sst>`,
    );
  zip
    .folder("xl")!
    .folder("worksheets")!
    .file(
      "sheet1.xml",
      `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${cells
        .map((c, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${c}</v></c></row>`)
        .join("")}</sheetData></worksheet>`,
    );
  return zip.generateAsync({ type: "nodebuffer" });
}

async function makePptxWithSlides(slides: string[]): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
  );
  zip.folder("_rels")!.file(
    ".rels",
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  const folder = zip.folder("ppt")!.folder("slides")!;
  slides.forEach((text, i) => {
    folder.file(
      `slide${i + 1}.xml`,
      `<?xml version="1.0"?><p:sld xmlns:p="..." xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    );
  });
  return zip.generateAsync({ type: "nodebuffer" });
}

describe("textHash.normalizeText", () => {
  it("lowercases and collapses whitespace", () => {
    assert.equal(normalizeText("  Hello   WORLD\n\nfoo "), "hello world foo");
  });
  it("preserves punctuation", () => {
    assert.equal(normalizeText("Hello, world!"), "hello, world!");
  });
  it("returns empty string for empty/whitespace input", () => {
    assert.equal(normalizeText(""), "");
    assert.equal(normalizeText("   \n\t  "), "");
  });
});

describe("textHash.extractText — docx", () => {
  it("pulls the visible <w:t> text", async () => {
    const buf = await makeDocxWithText(
      `<?xml version="1.0"?><w:document xmlns:w="..."><w:body><w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t xml:space="preserve"> world</w:t></w:r></w:p></w:body></w:document>`,
    );
    const text = await extractText(buf, "docx");
    assert.equal(text, "hello world");
  });

  it("decodes XML entities", async () => {
    const buf = await makeDocxWithText(
      `<?xml version="1.0"?><w:document><w:body><w:p><w:r><w:t>caf&#233; &amp; tea</w:t></w:r></w:p></w:body></w:document>`,
    );
    const text = await extractText(buf, "docx");
    assert.equal(text, "café & tea");
  });

  it("returns empty string for malformed/encrypted docx", async () => {
    const text = await extractText(Buffer.from("not a zip"), "docx");
    assert.equal(text, "");
  });
});

describe("textHash.extractText — xlsx", () => {
  it("pulls shared strings", async () => {
    const buf = await makeXlsxWithSharedStrings(["Apple", "Banana", "Cherry"], []);
    const text = await extractText(buf, "xlsx");
    // Cells reference shared strings by index, but the shared strings themselves
    // also contribute their text — we want both paths covered.
    assert.ok(text.includes("apple"));
    assert.ok(text.includes("banana"));
    assert.ok(text.includes("cherry"));
  });

  it("pulls inline cell <v> values", async () => {
    const buf = await makeXlsxWithSharedStrings([], ["100", "200", "300"]);
    const text = await extractText(buf, "xlsx");
    assert.ok(text.includes("100"));
    assert.ok(text.includes("300"));
  });
});

describe("textHash.extractText — pptx", () => {
  it("pulls slide text in order", async () => {
    const buf = await makePptxWithSlides(["Intro slide", "Conclusion slide"]);
    const text = await extractText(buf, "pptx");
    assert.ok(text.includes("intro slide"));
    assert.ok(text.includes("conclusion slide"));
    assert.ok(text.indexOf("intro") < text.indexOf("conclusion"));
  });
});

describe("textHash.textContentHash", () => {
  it("returns empty string for empty input", () => {
    assert.equal(textContentHash(""), "");
  });

  it("is deterministic for identical input", () => {
    const a = textContentHash("hello world");
    const b = textContentHash("hello world");
    assert.equal(a, b);
  });

  it("produces different hashes for different input", () => {
    const a = textContentHash("hello world");
    const b = textContentHash("hello earth");
    assert.notEqual(a, b);
  });
});

describe("textHash.simhash64", () => {
  it("returns empty string for empty input", () => {
    assert.equal(simhash64(""), "");
  });

  it("produces 16-character hex strings for non-empty input", () => {
    const h = simhash64("the quick brown fox jumps over the lazy dog");
    assert.equal(h.length, 16);
    assert.match(h, /^[0-9a-f]{16}$/);
  });

  it("is deterministic", () => {
    const a = simhash64("the quick brown fox jumps over the lazy dog");
    const b = simhash64("the quick brown fox jumps over the lazy dog");
    assert.equal(a, b);
  });

  it("produces near-zero Hamming distance for tiny edits", () => {
    // Same body, one word swapped — should still be very close in SimHash space.
    const original = simhash64("policy section one introduction this document outlines our standard operating procedures for vendor onboarding and compliance review across all departments");
    const tweaked = simhash64("policy section one introduction this document outlines our standard operating procedures for vendor onboarding and compliance review across departments");
    const d = simhashDistance(original, tweaked);
    // A two-word deletion in a ~25-word document should land in the single digits.
    assert.ok(d < 16, `Hamming distance should be small for a tiny edit, got ${d}`);
  });

  it("produces large Hamming distance for unrelated text", () => {
    const a = simhash64("the quick brown fox jumps over the lazy dog");
    const b = simhash64("entirely different content with nothing in common at all between them whatsoever");
    const d = simhashDistance(a, b);
    assert.ok(d >= 16, `Hamming distance should be large for unrelated text, got ${d}`);
  });
});

describe("textHash.simhashDistance", () => {
  it("returns 0 for identical hashes", () => {
    assert.equal(simhashDistance("abcd1234abcd1234", "abcd1234abcd1234"), 0);
  });

  it("returns the bit count of the XOR for distinct hashes", () => {
    // 0x0000...0000 ^ 0xffff...ffff == 64 set bits.
    assert.equal(simhashDistance("0000000000000000", "ffffffffffffffff"), 64);
    // 0xff ^ 0x00 in only the low byte → 8 bits.
    assert.equal(simhashDistance("00000000000000ff", "0000000000000000"), 8);
  });

  it("returns 64 (max) when either side is empty", () => {
    assert.equal(simhashDistance("", "abcd1234abcd1234"), 64);
    assert.equal(simhashDistance("abcd1234abcd1234", ""), 64);
    assert.equal(simhashDistance("", ""), 64);
  });
});
