import { BlobServiceClient, type BlobDownloadResponseParsed } from "@azure/storage-blob";
import type { ClassifiedLink } from "./linkInventoryScanner.js";
import type { ClassifiedDocLink } from "./documentLinkScanner.js";

/**
 * Rollforward aggregate of "current link state per site".
 *
 * Stored as one blob per site at:
 *   `link-inventory-results/aggregate/<sitePathSlug>.json`
 *
 * Why per-site sharding rather than one big blob: each scan worker only
 * touches its own site, so per-site shards mean independent ETag-
 * optimistic writes with no contention. A tenant-wide blob would force
 * serial writes across the whole scan.
 *
 * Why this exists separate from the per-job result blobs: the per-job
 * blobs get purged after 30 days (link-inventory retention). The
 * aggregate is the long-lived "current state" — daily delta scans roll
 * forward into it. When a scheduled delta scan runs:
 *   1. Read this site's aggregate
 *   2. Skip files where SP-reported ETag matches the aggregate entry
 *   3. Skip pages where SP-reported Modified <= aggregate.lastPageScanAt
 *   4. Scan only what's new/changed
 *   5. Merge new results into the aggregate via `mergeScanIntoAggregate`
 *
 * The aggregate is what the UI eventually displays as "current state of
 * all links" — independent of which scan touched each file last.
 *
 * NOTE: The merge is pure (exported for unit tests). Blob I/O is
 * isolated to `loadSiteAggregate` / `saveSiteAggregate` so tests don't
 * need Azurite. ETag-optimistic concurrency mirrors `hashIndexStore.ts`.
 */

const RESULTS_CONTAINER = "link-inventory-results";
const AGGREGATE_PREFIX = "aggregate/";

let blobServiceClient: BlobServiceClient | null = null;
let containerEnsured = false;

function getConnString(): string {
  const cs = process.env.TABLE_CONNECTION_STRING;
  if (!cs) throw new Error("TABLE_CONNECTION_STRING not configured");
  return cs;
}

function getBlobService(): BlobServiceClient {
  if (!blobServiceClient) blobServiceClient = BlobServiceClient.fromConnectionString(getConnString());
  return blobServiceClient;
}

async function ensureContainer(): Promise<void> {
  if (containerEnsured) return;
  const container = getBlobService().getContainerClient(RESULTS_CONTAINER);
  await container.createIfNotExists();
  containerEnsured = true;
}

/**
 * Convert a server-relative site path like "/sites/charlie-test-site"
 * into a safe blob name suffix. Only [a-z0-9-] passes through; everything
 * else collapses to "-". Avoids path-traversal and Azure blob-name
 * restrictions in one pass.
 */
export function siteSlug(sitePath: string): string {
  return sitePath
    .toLowerCase()
    .replace(/^\/sites\//, "")
    .replace(/^\//, "")
    .replace(/[^a-z0-9-]/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 200) || "root";
}

function blobName(sitePath: string): string {
  return `${AGGREGATE_PREFIX}${siteSlug(sitePath)}.json`;
}

// --- Persistent shape ---

export interface AggregatePageEntry {
  pageUrl: string;
  pageTitle: string;
  /** SP-reported Modified ISO. Used for delta skip. */
  modified: string;
  links: ClassifiedLink[];
  /** ISO of last scan that confirmed this entry. */
  lastConfirmedAt: string;
  /** Job ID that last touched this entry. */
  lastConfirmedByJobId: string;
}

export interface AggregateFileEntry {
  fileRef: string;
  fileName: string;
  fileType: "docx" | "xlsx" | "pptx" | "pdf" | "other";
  /** SP-reported Modified ISO. */
  modified: string;
  /** SP-reported ETag. Primary delta-skip key for doc scans. */
  etag: string;
  size: number;
  links: ClassifiedDocLink[];
  /** When extraction failed; most recent error. Cleared on successful rescan. */
  scanError?: string;
  lastConfirmedAt: string;
  lastConfirmedByJobId: string;
}

export interface SiteAggregate {
  version: 1;
  sitePath: string;
  builtAt: string;
  lastPageScanAt: string | null;
  lastDocScanAt: string | null;
  pages: Record<string, AggregatePageEntry>;
  files: Record<string, AggregateFileEntry>;
}

export function emptySiteAggregate(sitePath: string): SiteAggregate {
  return {
    version: 1,
    sitePath,
    builtAt: new Date().toISOString(),
    lastPageScanAt: null,
    lastDocScanAt: null,
    pages: {},
    files: {},
  };
}

// --- Merge inputs ---

/** Per-page scan result shape for merge — mirrors what the page worker emits. */
export interface PageScanResult {
  pageUrl: string;
  pageTitle: string;
  modified: string;
  links: ClassifiedLink[];
}

/** Per-file scan result shape for merge — mirrors `DocumentFileInventory`. */
export interface FileScanResult {
  fileRef: string;
  fileName: string;
  fileType: "docx" | "xlsx" | "pptx" | "pdf" | "other";
  modified: string;
  etag: string;
  size: number;
  links: ClassifiedDocLink[];
  scanError?: string;
}

// --- Merge (pure) ---

/**
 * Merge a single site's page scan into its aggregate. Pages absent from
 * `results` are left untouched (they may belong to a different scan
 * scope — leave their previous links intact).
 *
 * `confirmedAt` is the timestamp written into every confirmed entry.
 * Caller should use the scan job's start time so all pages confirmed by
 * one scan share the same value.
 */
export function mergePageScanIntoAggregate(
  current: SiteAggregate,
  results: PageScanResult[],
  jobId: string,
  confirmedAt: string,
): SiteAggregate {
  const pages = { ...current.pages };
  for (const r of results) {
    if (!r.pageUrl) continue;
    pages[r.pageUrl] = {
      pageUrl: r.pageUrl,
      pageTitle: r.pageTitle,
      modified: r.modified,
      links: r.links,
      lastConfirmedAt: confirmedAt,
      lastConfirmedByJobId: jobId,
    };
  }
  return {
    ...current,
    builtAt: new Date().toISOString(),
    lastPageScanAt: confirmedAt,
    pages,
  };
}

/**
 * Merge a single site's doc scan into its aggregate. Files absent from
 * `results` are left untouched. The merge replaces the entry wholesale
 * — extracted links are NOT diffed since the bytes have changed (or
 * else the file would have been delta-skipped before reaching here).
 *
 * Pre-existing `scanError` is cleared on a successful re-scan; preserved
 * if the new scan also failed.
 */
export function mergeDocScanIntoAggregate(
  current: SiteAggregate,
  results: FileScanResult[],
  jobId: string,
  confirmedAt: string,
): SiteAggregate {
  const files = { ...current.files };
  for (const r of results) {
    if (!r.fileRef) continue;
    files[r.fileRef] = {
      fileRef: r.fileRef,
      fileName: r.fileName,
      fileType: r.fileType,
      modified: r.modified,
      etag: r.etag,
      size: r.size,
      links: r.links,
      scanError: r.scanError,
      lastConfirmedAt: confirmedAt,
      lastConfirmedByJobId: jobId,
    };
  }
  return {
    ...current,
    builtAt: new Date().toISOString(),
    lastDocScanAt: confirmedAt,
    files,
  };
}

/**
 * Bump the lastConfirmedAt + jobId on a list of entries WITHOUT touching
 * their content. Called for delta-skipped items so we know they were
 * still present and unchanged at confirmedAt — preserves the "latest
 * verified" timestamp for the UI without re-extracting links.
 *
 * The pageUrls/fileRefs that aren't in the aggregate are silently
 * ignored — the caller may have just enumerated something the previous
 * scan never reached.
 */
export function carryForwardEntries(
  current: SiteAggregate,
  pageUrls: string[],
  fileRefs: string[],
  jobId: string,
  confirmedAt: string,
): SiteAggregate {
  const pages = { ...current.pages };
  const files = { ...current.files };
  for (const u of pageUrls) {
    const e = pages[u];
    if (e) pages[u] = { ...e, lastConfirmedAt: confirmedAt, lastConfirmedByJobId: jobId };
  }
  for (const r of fileRefs) {
    const e = files[r];
    if (e) files[r] = { ...e, lastConfirmedAt: confirmedAt, lastConfirmedByJobId: jobId };
  }
  return {
    ...current,
    builtAt: new Date().toISOString(),
    pages,
    files,
  };
}

// --- Blob I/O with ETag optimistic concurrency ---

async function readBlobWithEtag(sitePath: string): Promise<{ body: SiteAggregate | null; etag: string | null }> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(blobName(sitePath));
  try {
    const dl: BlobDownloadResponseParsed = await blob.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    const body = JSON.parse(Buffer.concat(chunks).toString("utf8")) as SiteAggregate;
    return { body, etag: dl.etag ?? null };
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return { body: null, etag: null };
    }
    throw err;
  }
}

async function writeBlobWithEtag(
  sitePath: string,
  body: SiteAggregate,
  ifMatchEtag: string | null,
): Promise<void> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(blobName(sitePath));
  const data = JSON.stringify(body);
  const conditions = ifMatchEtag ? { ifMatch: ifMatchEtag } : { ifNoneMatch: "*" };
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
    conditions,
  });
}

/**
 * Read the current aggregate for a site, returning an empty aggregate
 * if no blob exists yet. The caller doesn't need to distinguish "first
 * scan" from "fully populated" — the empty case lets the merge functions
 * produce a freshly-built aggregate on first run.
 */
export async function loadSiteAggregate(sitePath: string): Promise<SiteAggregate> {
  const { body } = await readBlobWithEtag(sitePath);
  return body ?? emptySiteAggregate(sitePath);
}

/**
 * Read-modify-write pattern with ETag-optimistic concurrency. Retries
 * on 412/409 (the typical race) up to three times. The mutation
 * function receives the current aggregate (or empty) and returns the
 * new aggregate; it MUST be pure / cheap since it can run multiple
 * times under contention.
 */
export async function updateSiteAggregate(
  sitePath: string,
  mutate: (current: SiteAggregate) => SiteAggregate,
): Promise<SiteAggregate> {
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const { body, etag } = await readBlobWithEtag(sitePath);
    const current = body ?? emptySiteAggregate(sitePath);
    const next = mutate(current);
    try {
      await writeBlobWithEtag(sitePath, next, etag);
      return next;
    } catch (err) {
      const status = err instanceof Error && "statusCode" in err ? (err as { statusCode: number }).statusCode : 0;
      if ((status === 412 || status === 409) && attempt < maxAttempts) continue;
      throw err;
    }
  }
  // Unreachable — loop either returns or throws.
  throw new Error("updateSiteAggregate: exhausted retries");
}
