import { createHash } from "node:crypto";
import JSZip from "jszip";
import type { DocumentFileRef } from "./spoFilesEnumerator.js";

/**
 * Text-content extraction + similarity hashing for "fuzzy" duplicate
 * detection: rename-chain forks, copy-then-edit divergences, polish-only
 * variants. Lives next to `contentHash.ts`; the two modules answer
 * different questions about the same buffer.
 *
 * Three exports:
 *
 *   - `extractText(bytes, fileType)` — walk the document, return its
 *     visible text as a single normalized utf-8 string. Whitespace runs
 *     are collapsed to a single space and the whole thing is lowercased.
 *     Formatting, page breaks, comments, and per-cell positioning are
 *     deliberately discarded — we want "what does this document say,"
 *     not "how does it look." Identical wording with different fonts or
 *     reviewer comments → identical text.
 *
 *   - `textContentHash(text)` — SHA-256 hex of the normalized text.
 *     Used for the **Diverged-pair** query: two files share an ancestor
 *     when one of A's historical text-hashes matches one of B's.
 *
 *   - `simhash64(text)` — 64-bit Charikar SimHash over 3-shingle word
 *     tokens. Returned as a 16-char hex string. Used for the
 *     **Near-duplicate** query via Hamming distance.
 *
 * All three are pure (byte-in / value-out) so they unit-test cleanly
 * and can run in any of the existing scan paths.
 */

export const TEXT_HASH_VERSION = "text-content-v1" as const;
export const SIMHASH_VERSION = "simhash-v1" as const;

/** OOXML zip parts that hold the visible text we want to hash. */
const DOCX_TEXT_PART = "word/document.xml";
const XLSX_SHARED_STRINGS = "xl/sharedStrings.xml";
const XLSX_SHEET_RE = /^xl\/worksheets\/sheet\d+\.xml$/;
const PPTX_SLIDE_RE = /^ppt\/slides\/slide\d+\.xml$/;

/** Pull the inner text of every `<w:t>` / `<a:t>` / `<t>` element. */
function pullElementText(xml: string, tag: "w:t" | "a:t" | "t"): string {
  // Permissive: matches `<tag>X</tag>` and `<tag attr="...">X</tag>`.
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, "g");
  const out: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) out.push(m[1]);
  return out.join(" ");
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, n) => String.fromCharCode(parseInt(n, 16)));
}

/** Lowercase, collapse whitespace, trim. Punctuation is preserved. */
export function normalizeText(text: string): string {
  return text
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

async function extractDocxText(bytes: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const part = zip.file(DOCX_TEXT_PART);
  if (!part) return "";
  const xml = await part.async("string");
  return decodeXmlEntities(pullElementText(xml, "w:t"));
}

async function extractXlsxText(bytes: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const out: string[] = [];
  // Shared strings table — text in any cell that uses an inline string
  // index lives here. Pulling `<t>` directly catches both the
  // <si><t>foo</t></si> and <si><r><t>foo</t></r></si> shapes.
  const ss = zip.file(XLSX_SHARED_STRINGS);
  if (ss) {
    const xml = await ss.async("string");
    out.push(decodeXmlEntities(pullElementText(xml, "t")));
  }
  // Inline values on sheets — anything not using shared strings shows
  // up directly inside <c><v>...</v></c>. Cell coordinates are dropped
  // intentionally: same words in different positions hash the same.
  const sheetNames = Object.keys(zip.files).filter((n) => XLSX_SHEET_RE.test(n)).sort();
  for (const name of sheetNames) {
    const xml = await zip.files[name].async("string");
    // Pull <v>…</v> values — Excel writes inline strings as <is><t>…</t></is>
    // too (caught by the shared-strings extractor above). <v> covers the
    // numeric and inline-string-index forms.
    const vRe = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/g;
    let m: RegExpExecArray | null;
    while ((m = vRe.exec(xml))) out.push(decodeXmlEntities(m[1]));
  }
  return out.join(" ");
}

async function extractPptxText(bytes: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const slideNames = Object.keys(zip.files).filter((n) => PPTX_SLIDE_RE.test(n)).sort();
  const out: string[] = [];
  for (const name of slideNames) {
    const xml = await zip.files[name].async("string");
    out.push(decodeXmlEntities(pullElementText(xml, "a:t")));
  }
  return out.join(" ");
}

async function extractPdfText(bytes: Buffer): Promise<string> {
  // pdfjs-dist legacy build is what the link extractor uses, so the
  // worker config is already settled in the runtime.
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  // The legacy build is happy without a worker for Node usage.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const loadingTask = (pdfjs as any).getDocument({
    data: new Uint8Array(bytes),
    isEvalSupported: false,
    useSystemFonts: false,
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const doc: any = await loadingTask.promise;
  try {
    const out: string[] = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const items: any[] = tc.items;
      for (const it of items) {
        if (typeof it.str === "string") out.push(it.str);
      }
      page.cleanup();
    }
    return out.join(" ");
  } finally {
    await doc.destroy();
  }
}

/**
 * Extract normalized visible text from a document. Empty string for
 * unsupported types or extraction failures (caller decides whether
 * absence of a text hash is a no-op or a flag).
 */
export async function extractText(
  bytes: Buffer,
  fileType: DocumentFileRef["fileType"],
): Promise<string> {
  try {
    let raw = "";
    switch (fileType) {
      case "docx":
        raw = await extractDocxText(bytes);
        break;
      case "xlsx":
        raw = await extractXlsxText(bytes);
        break;
      case "pptx":
        raw = await extractPptxText(bytes);
        break;
      case "pdf":
        raw = await extractPdfText(bytes);
        break;
      default:
        return "";
    }
    return normalizeText(raw);
  } catch {
    return "";
  }
}

/**
 * SHA-256 of normalized text. Empty input returns empty string so
 * callers can cheaply detect "no text was extracted" instead of treating
 * the well-known SHA-256 of the empty string as a real ancestor.
 */
export function textContentHash(normalizedText: string): string {
  if (!normalizedText) return "";
  return createHash("sha256").update(normalizedText).digest("hex");
}

/**
 * 3-shingle tokenization. Split on word boundaries, drop empty tokens,
 * then group into overlapping triples. "the quick brown fox" → ["the
 * quick brown", "quick brown fox"]. Two-word texts produce one shingle;
 * one-word texts produce one shingle (their single token).
 */
function shingles(text: string): string[] {
  const tokens = text.split(/[^\p{L}\p{N}_]+/u).filter((t) => t.length > 0);
  if (tokens.length === 0) return [];
  if (tokens.length <= 3) return [tokens.join(" ")];
  const out: string[] = [];
  for (let i = 0; i + 2 < tokens.length; i++) {
    out.push(`${tokens[i]} ${tokens[i + 1]} ${tokens[i + 2]}`);
  }
  return out;
}

/**
 * Charikar SimHash, 64-bit. For each shingle, hash to 64 bits and vote
 * per bit position (+1 for set, -1 for clear). Final bit is 1 when the
 * vote is positive.
 *
 * Returns a 16-char zero-padded hex string. Empty input returns empty
 * string (matches `textContentHash`'s convention so a caller's
 * "did the file have text?" check is one comparison).
 */
export function simhash64(normalizedText: string): string {
  const shings = shingles(normalizedText);
  if (shings.length === 0) return "";
  const votes = new Array(64).fill(0);
  for (const sh of shings) {
    // SHA-256 → first 8 bytes → 64-bit BigInt.
    const digest = createHash("sha256").update(sh).digest();
    let h = 0n;
    for (let i = 0; i < 8; i++) {
      h = (h << 8n) | BigInt(digest[i]);
    }
    for (let bit = 0; bit < 64; bit++) {
      if ((h >> BigInt(bit)) & 1n) votes[bit] += 1;
      else votes[bit] -= 1;
    }
  }
  let result = 0n;
  for (let bit = 0; bit < 64; bit++) {
    if (votes[bit] > 0) result |= (1n << BigInt(bit));
  }
  return result.toString(16).padStart(16, "0");
}

/**
 * Hamming distance between two 64-bit SimHash hex strings. Empty
 * strings (no text was extracted on at least one side) return 64 — the
 * maximum, so they never sneak under any reasonable Near-duplicate
 * threshold.
 */
export function simhashDistance(a: string, b: string): number {
  if (!a || !b) return 64;
  let xor = BigInt("0x" + a) ^ BigInt("0x" + b);
  let count = 0;
  while (xor > 0n) {
    if (xor & 1n) count++;
    xor >>= 1n;
  }
  return count;
}

/** Default Near-duplicate threshold over 64-bit SimHash. */
export const NEAR_DUPLICATE_THRESHOLD = 3;
