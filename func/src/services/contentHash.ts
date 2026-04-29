import { createHash } from "node:crypto";
import JSZip from "jszip";
import type { DocumentFileRef } from "./spoFilesEnumerator.js";

/**
 * Content-aware file hashing for the duplicate-detection index.
 *
 * Two modes:
 *
 *   - `full-sha256` — SHA-256 of the entire file bytes. Used for PDFs and
 *     anything else; PDFs round-trip through SharePoint upload without
 *     mutation, so the byte-level hash is stable across copies.
 *
 *   - `ooxml-content-v1` — open the OPC zip, drop the property-bag parts
 *     SharePoint rewrites on upload (`docProps/`, `customXml/`,
 *     `[Content_Types].xml`, `_rels/.rels`), hash each remaining entry
 *     with its name as a salt, then SHA-256 the concatenation. Stable
 *     across SharePoint's property-promotion mutations: the same source
 *     `.docx` uploaded into two libraries with different column schemas
 *     produces the same hash.
 *
 * The two modes are NEVER comparable to each other. The hash index keys
 * groups by `(algo, hash)` — see `hashIndexStore`.
 */

export type HashAlgo = "full-sha256" | "ooxml-content-v1";

/**
 * OPC parts SharePoint rewrites on upload as part of property promotion.
 * Anything matching this prefix-set is excluded from the OOXML content
 * hash. The set is conservative: omitting more than necessary just makes
 * the hash more permissive (more files group together), which is the
 * intended direction for duplicate detection.
 *
 *   - `docProps/`        core, app, custom property XMLs
 *   - `customXml/`       library-column-driven property item XMLs
 *   - `[Content_Types].xml`  package manifest, mutates when customXml
 *                       parts are added or removed
 *   - `_rels/.rels`      package-level relationships, gains entries when
 *                       new parts appear
 */
const OOXML_SKIP = /^(docProps\/|customXml\/|\[Content_Types\]\.xml$|_rels\/\.rels$)/;

async function ooxmlContentHash(bytes: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files)
    .filter((n) => !zip.files[n].dir && !OOXML_SKIP.test(n))
    .sort();
  const acc = createHash("sha256");
  for (const name of names) {
    const partBytes = await zip.files[name].async("nodebuffer");
    acc
      .update(name)
      .update(":")
      .update(createHash("sha256").update(partBytes).digest())
      .update("\n");
  }
  return acc.digest("hex");
}

function fullSha256(bytes: Buffer): string {
  return createHash("sha256").update(bytes).digest("hex");
}

/**
 * Hash a file's content for duplicate-detection purposes. OOXML formats
 * (`.docx`, `.xlsx`, `.pptx`) get the content-aware path; everything else
 * falls through to a full-bytes SHA-256.
 *
 * Encrypted or malformed OOXML files (where JSZip cannot open the
 * archive) silently fall back to `full-sha256`. They will not group with
 * other OOXML entries of the same logical content, but they will still
 * dedup against bytewise-identical copies of themselves.
 */
export async function hashFileContent(
  bytes: Buffer,
  fileType: DocumentFileRef["fileType"],
): Promise<{ hash: string; algo: HashAlgo }> {
  if (fileType === "docx" || fileType === "xlsx" || fileType === "pptx") {
    try {
      return { hash: await ooxmlContentHash(bytes), algo: "ooxml-content-v1" };
    } catch {
      // Fall through to full-bytes — better to contribute SOMETHING to
      // the index than skip the file.
    }
  }
  return { hash: fullSha256(bytes), algo: "full-sha256" };
}
