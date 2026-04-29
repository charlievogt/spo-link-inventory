import { extractLinksFromDocument, type ExtractedDocLink } from "./documentLinkExtractor.js";
import { fetchFileBytes, type DocumentFileRef } from "./spoFilesEnumerator.js";
import { classifyLink, type ClassifiedLink, type LinkClass, type MalformedReason } from "./linkInventoryScanner.js";
import { hashFileContent, type HashAlgo } from "./contentHash.js";
import { extractText, textContentHash, simhash64 } from "./textHash.js";
import type { AggregateFileEntry } from "./aggregateStore.js";

/**
 * Per-file document scan: download bytes → extract links → classify.
 *
 * Mirrors the shape of the page scanner so per-file results can sit
 * alongside per-page results in the UI eventually. The classifier
 * (`classifyLink`) is reused verbatim — every URL the document
 * extractor produces flows through the same on-prem / SPO / external /
 * sharing-link / etc. taxonomy as page links.
 *
 * The scanner never throws — every per-file failure is captured as
 * `error` on the result so the queue worker can keep going past
 * corrupted documents and DRM-protected files.
 */

export interface ClassifiedDocLink extends ExtractedDocLink {
  linkClass: LinkClass;
  /** Normalized lookup key (empty when the URL doesn't normalize). */
  normalizedKey: string;
  /** Set only when linkClass === "malformed-spo-link". */
  malformedReason?: MalformedReason;
}

export interface DocumentFileInventory {
  site: string;
  library: string;
  fileRef: string;
  fileName: string;
  modified: string;
  etag?: string;
  length: number;
  fileType: "docx" | "xlsx" | "pptx" | "pdf" | "other";
  links: ClassifiedDocLink[];
  /** Set when extraction failed — file metadata is still returned. */
  error?: string;
  /** Wall-clock ms spent downloading + parsing this file. */
  scanMs: number;
  /**
   * Hex digest of the file content. Computed from the same buffer used
   * for link extraction — near-zero marginal cost beyond the OOXML
   * unzip-and-hash-parts step. Used downstream by the duplicate-detection
   * hash index. Unset when the download failed (error is populated instead).
   */
  sha256?: string;
  /**
   * Algorithm that produced `sha256`. Pairs with the digest so the hash
   * index can keep `full-sha256` and `ooxml-content-v1` digests in
   * separate buckets — they are not comparable across modes.
   */
  hashAlgo?: HashAlgo;
  /**
   * SHA-256 of the document's normalized visible text. Empty string when
   * the file type has no text path or extraction failed. Used for
   * Diverged-pair detection (history-history textHash matches).
   */
  textHash?: string;
  /**
   * 64-bit Charikar SimHash of the same normalized text, 16-char hex.
   * Used by the Near-duplicate query.
   */
  simhash64?: string;
  /**
   * True when the scan short-circuited because the SP-reported ETag
   * matched the previous aggregate entry's ETag — bytes haven't changed
   * since last scan, links were reused from the aggregate, no download.
   * The file's hash isn't recomputed in this path (stays unset on this
   * record) — duplicate detection still has the previous hash via the
   * persistent hash index.
   */
  reusedFromPreviousScan?: boolean;
}

/**
 * Convert raw extractor output into a classified, normalized link.
 * The scanner intentionally does NOT call `normalizeUrl` for now —
 * the page-side classifier already special-cases the URL types we
 * care about, and adding normalized keys per document link can wait
 * for find-and-replace v2.
 */
function classifyDocLinks(links: ExtractedDocLink[]): ClassifiedDocLink[] {
  return links.map((l) => {
    const linkClass = classifyLink(l.url);
    const malformedReason: MalformedReason | undefined =
      linkClass === "malformed-spo-link" ? "search-fragment" : undefined;
    return {
      ...l,
      linkClass,
      normalizedKey: "",
      malformedReason,
    };
  });
}

/**
 * Scan one document file end-to-end. Always returns a populated
 * `DocumentFileInventory`; failures are captured in `error`.
 *
 * `options.previousEntry` enables the **ETag-skip optimization**: when
 * the previous aggregate entry's `etag` matches what SP reports for
 * this file today, the bytes haven't changed since last scan and we
 * reuse the previously-extracted links — skipping download, hash, and
 * link extraction entirely. This is the bandwidth-saver that makes
 * daily scheduled scans cheap on stable libraries.
 *
 * The skip is silent on the wire — caller checks `reusedFromPreviousScan`
 * on the result to know whether actual work was done.
 */
export async function scanDocumentFile(
  file: DocumentFileRef,
  options: { previousEntry?: AggregateFileEntry } = {},
): Promise<DocumentFileInventory> {
  const start = Date.now();
  const out: DocumentFileInventory = {
    site: file.site,
    library: file.library,
    fileRef: file.fileRef,
    fileName: file.fileName,
    modified: file.modified,
    etag: file.etag,
    length: file.length,
    fileType: file.fileType,
    links: [],
    scanMs: 0,
  };

  // ETag-skip path: same SP-reported etag as the previous aggregate
  // entry → file bytes are unchanged → reuse previously-extracted links.
  // Empty etag from SP (rare, but seen for some doc types) defeats the
  // skip — fall through to full scan in that case.
  const prev = options.previousEntry;
  if (prev && file.etag && prev.etag && prev.etag === file.etag) {
    out.links = prev.links;
    out.reusedFromPreviousScan = true;
    out.scanMs = Date.now() - start;
    return out;
  }

  let buf: Buffer;
  try {
    buf = await fetchFileBytes(file.fileRef);
  } catch (e) {
    out.error = `download failed: ${(e as Error).message}`;
    out.scanMs = Date.now() - start;
    return out;
  }

  // Hash the bytes we just downloaded. For OOXML formats this opens the
  // OPC zip and hashes only the content parts (skipping the property-bag
  // parts SharePoint rewrites on upload), so the same source file uploaded
  // into two libraries with different column schemas dedups correctly.
  // PDFs and unknown types take a full-bytes SHA-256.
  // Done before link extraction so a parser throw doesn't lose the hash.
  const { hash, algo } = await hashFileContent(buf, file.fileType);
  out.sha256 = hash;
  out.hashAlgo = algo;

  // Text hashes for Near-duplicate (SimHash) and Diverged-pair (textHash)
  // queries. Both return empty strings on extraction failure or
  // unsupported file types — downstream readers treat empty/undefined
  // identically and just skip the file from those queries.
  const text = await extractText(buf, file.fileType);
  if (text) {
    out.textHash = textContentHash(text);
    out.simhash64 = simhash64(text);
  }

  try {
    const result = await extractLinksFromDocument(buf, file.fileName);
    if (result.parseError) {
      out.error = result.parseError;
    }
    out.links = classifyDocLinks(result.links);
  } catch (e) {
    // Defensive — extractor wraps its own errors but we double-cover
    // here in case a future code path throws.
    out.error = `extract failed: ${(e as Error).message}`;
  }

  out.scanMs = Date.now() - start;
  return out;
}
