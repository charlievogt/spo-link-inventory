import { BlobServiceClient, type BlobDownloadResponseParsed } from "@azure/storage-blob";
import type { ISiteInventoryShape } from "../functions/linkInventoryReplace.types.js";
import {
  type BacklinkEntry,
  type BacklinksIndex,
  type SourceKind,
  deriveSiteLabel,
  toAbsoluteSpoUrl,
} from "./backlinksIndex.js";
import { normalizeUrl } from "./urlNormalizer.js";
import { getTenantHost } from "./config.js";

/**
 * Persistent backlinks index, stored as its own blob outside any jobId
 * prefix so it survives the 30-day scan-retention purge.
 *
 * Why not read per-request from an individual scan blob: scans get
 * purged after 30 days, and partial scans (subset of sites) would blank
 * out sites they didn't cover. The persistent index decouples the
 * library-column feature from scan retention and preserves per-site
 * data across partial rescans via merge-update semantics.
 *
 * Layout (`link-inventory-results/backlinks-index.json`):
 *
 *   {
 *     version: 1,
 *     builtAt: ISO,
 *     bySite: {
 *       "/sites/hr": {
 *         siteLabel: "hr",
 *         siteUrl: "https://contoso.sharepoint.com/sites/hr",
 *         pages?: { updatedAt, fromJobId, sources: [...] },
 *         documents?: { updatedAt, fromJobId, sources: [...] }
 *       },
 *       ...
 *     }
 *   }
 *
 * Pages and documents are stored separately per site so a page scan
 * only replaces the `pages` slice, leaving `documents` intact (and vice
 * versa). Sites the scan didn't cover stay untouched.
 *
 * Concurrency: writes use ETag optimistic concurrency with 3 retries.
 * Two scans completing at once is rare but possible (the worker
 * finalize path runs on queue-triggered invocations that aren't
 * serialized).
 *
 * In-memory cache: the read path memoizes the inverted
 * `Map<canonicalKey, BacklinkEntry[]>` for 5 minutes per Function
 * instance, mirroring redirectService.
 */

const RESULTS_CONTAINER = "link-inventory-results";
const INDEX_BLOB_NAME = "backlinks-index.json";
/**
 * Fast-path TTL — if we looked up < this long ago, return the cached
 * index without any I/O. Short because the next tier (HEAD validation)
 * is also cheap.
 */
const CACHE_FAST_TTL_MS = 30 * 1000;
/**
 * Validation-path TTL — after the fast TTL expires but within this
 * window, HEAD the blob to compare ETag. If the ETag is unchanged,
 * refresh the fast timer and return the cached index (no parse). This
 * is the usual steady-state path: the index rarely changes between
 * reads, so HEAD-validate and skip the multi-MB download+parse.
 */
const CACHE_HEAD_VALIDATE_TTL_MS = 30 * 60 * 1000;

const TENANT_HOST = getTenantHost();
const TENANT_HOST_PREFIX = TENANT_HOST + "/";

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

// --- On-disk persistent shape ---

interface SourceRecord {
  /** Page title or file name. */
  title: string;
  /** Absolute source URL for display. */
  url: string;
  /** Path-only canonicalKeys this source links to. */
  linksTo: string[];
}

interface SectionSlice {
  updatedAt: string;
  fromJobId: string;
  sources: SourceRecord[];
}

interface SiteSection {
  siteLabel: string;
  siteUrl: string;
  pages?: SectionSlice;
  documents?: SectionSlice;
}

export interface PersistentBacklinksIndex {
  version: 1;
  builtAt: string;
  bySite: Record<string, SiteSection>;
}

function emptyIndex(): PersistentBacklinksIndex {
  return {
    version: 1,
    builtAt: new Date().toISOString(),
    bySite: {},
  };
}

function toPathOnlyKey(key: string): string {
  if (!key) return "";
  if (key.startsWith(TENANT_HOST_PREFIX)) {
    return key.slice(TENANT_HOST.length);
  }
  return key;
}

// --- Transform: scan aggregate → site sections ---

/**
 * Pure function: convert a scan aggregate into a map of sitePath →
 * SectionSlice. Exported so the rebuild endpoint and unit tests can
 * exercise it without touching storage.
 */
export function buildSectionSlicesFromScan(
  aggregate: ISiteInventoryShape[],
  jobId: string,
  scannedAt: string,
): Map<string, SectionSlice & { siteLabel: string; siteUrl: string }> {
  const out = new Map<string, SectionSlice & { siteLabel: string; siteUrl: string }>();
  for (const site of aggregate ?? []) {
    const rawSite = site.site ?? "";
    const sitePath = rawSite.toLowerCase();
    if (!sitePath) continue;

    const sources: SourceRecord[] = [];
    for (const page of site.pages ?? []) {
      const url = toAbsoluteSpoUrl(page.pageUrl ?? "");
      const title = page.pageTitle ?? "(untitled)";
      const linksTo = new Set<string>();
      for (const link of page.links ?? []) {
        if (!link.canonicalKey) continue;
        const k = toPathOnlyKey(link.canonicalKey);
        if (!k) continue;
        linksTo.add(k);
      }
      if (linksTo.size === 0) continue;
      sources.push({ title, url, linksTo: [...linksTo].sort() });
    }

    out.set(sitePath, {
      siteLabel: deriveSiteLabel(rawSite),
      siteUrl: toAbsoluteSpoUrl(rawSite),
      updatedAt: scannedAt,
      fromJobId: jobId,
      sources,
    });
  }
  return out;
}

// --- Merge semantics ---

/**
 * Merge a kind-specific scan result into the current persistent index.
 * Replaces the `pages` or `documents` slice (per `kind`) for every site
 * covered by the scan; other sites and the other kind are untouched.
 *
 * Pure function (no I/O) so tests can cover the merge rules
 * exhaustively.
 *
 * Sites present in the scan with zero sources (no outbound links
 * captured) still trigger a slice write — an empty slice with a fresh
 * `updatedAt` tells the read path "yes, we scanned this, there were no
 * backlinks." That's different from a missing slice, which means "we
 * never scanned this."
 */
export function mergeScanIntoIndex(
  current: PersistentBacklinksIndex,
  slices: Map<string, SectionSlice & { siteLabel: string; siteUrl: string }>,
  kind: SourceKind,
  scannedSitePaths: string[],
): PersistentBacklinksIndex {
  const out: PersistentBacklinksIndex = {
    version: 1,
    builtAt: new Date().toISOString(),
    bySite: { ...current.bySite },
  };

  // Apply every scanned site, even if the scan produced an empty slice
  // for it. The "scan ran, zero sources" signal matters for staleness.
  for (const sitePath of scannedSitePaths) {
    const key = sitePath.toLowerCase();
    const slice = slices.get(key);
    const existing = out.bySite[key];

    const siteLabel = slice?.siteLabel ?? existing?.siteLabel ?? deriveSiteLabel(key);
    const siteUrl = slice?.siteUrl ?? existing?.siteUrl ?? toAbsoluteSpoUrl(key);

    const newSlice: SectionSlice | undefined = slice
      ? { updatedAt: slice.updatedAt, fromJobId: slice.fromJobId, sources: slice.sources }
      : existing && kind === "page" && existing.pages
        ? existing.pages
        : existing && kind === "document" && existing.documents
          ? existing.documents
          : undefined;

    out.bySite[key] = {
      siteLabel,
      siteUrl,
      pages: kind === "page" ? newSlice : existing?.pages,
      documents: kind === "document" ? newSlice : existing?.documents,
    };
  }

  return out;
}

// --- Blob I/O with ETag optimistic concurrency ---

async function readBlobWithEtag(): Promise<{ body: PersistentBacklinksIndex | null; etag: string | null }> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(INDEX_BLOB_NAME);
  try {
    const dl: BlobDownloadResponseParsed = await blob.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    const text = Buffer.concat(chunks).toString("utf8");
    const body = JSON.parse(text) as PersistentBacklinksIndex;
    return { body, etag: dl.etag ?? null };
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return { body: null, etag: null };
    }
    throw err;
  }
}

/**
 * Cheap metadata-only probe — returns the current blob ETag without
 * downloading the body. Used to validate an in-memory cached index
 * without a full re-parse. Costs ~50ms vs hundreds of ms + JSON.parse
 * for a full download.
 */
async function headBlobEtag(): Promise<string | null> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(INDEX_BLOB_NAME);
  try {
    const props = await blob.getProperties();
    return props.etag ?? null;
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return null;
    }
    throw err;
  }
}

async function writeBlobWithEtag(
  body: PersistentBacklinksIndex,
  ifMatchEtag: string | null,
): Promise<void> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(INDEX_BLOB_NAME);
  const data = JSON.stringify(body);
  const conditions = ifMatchEtag
    ? { ifMatch: ifMatchEtag }
    : { ifNoneMatch: "*" }; // first write — blob must not exist yet
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
    conditions,
  });
}

/**
 * Merge a scan's results into the persistent index with ETag optimistic
 * concurrency. Reads current → merges in memory → writes with ifMatch.
 * Retries up to 3 times on 412 (precondition failed), which means
 * another finalize raced us.
 */
export async function mergeScanResultsToStore(
  aggregate: ISiteInventoryShape[],
  scannedSitePaths: string[],
  kind: SourceKind,
  jobId: string,
  scannedAt: string,
): Promise<void> {
  const slices = buildSectionSlicesFromScan(aggregate, jobId, scannedAt);

  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const { body, etag } = await readBlobWithEtag();
    const current = body ?? emptyIndex();
    const next = mergeScanIntoIndex(current, slices, kind, scannedSitePaths);
    try {
      await writeBlobWithEtag(next, etag);
      // Bust the in-memory cache on successful write.
      cachedInMemory = null;
      cachedEtag = null;
      cachedAt = 0;
      return;
    } catch (err) {
      const status = err instanceof Error && "statusCode" in err ? (err as { statusCode: number }).statusCode : 0;
      // 412 = ETag mismatch (concurrent writer). 409 = blob-exists on
      // first-write race. Both are retryable.
      if ((status === 412 || status === 409) && attempt < maxAttempts) continue;
      throw err;
    }
  }
}

/**
 * Full rebuild: replace the entire persistent index with data from a
 * fresh in-memory build (caller typically assembles it from the latest
 * page scan + latest doc scan aggregates).
 *
 * Unlike merge, this does NOT preserve prior data for sites not covered
 * — it's intended for "blow away and rebuild from scratch." Callers
 * that want partial updates should use `mergeScanResultsToStore`.
 */
export async function writeFullIndex(index: PersistentBacklinksIndex): Promise<void> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(INDEX_BLOB_NAME);
  const data = JSON.stringify(index);
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
  cachedInMemory = null;
  cachedEtag = null;
  cachedAt = 0;
}

/**
 * Assemble a PersistentBacklinksIndex from one page-scan aggregate plus
 * one doc-scan aggregate. Used by the rebuild endpoint.
 */
export function buildPersistentFromAggregates(
  pageAggregate: ISiteInventoryShape[] | null,
  pageJobId: string | null,
  pageScannedAt: string | null,
  docAggregate: ISiteInventoryShape[] | null,
  docJobId: string | null,
  docScannedAt: string | null,
): PersistentBacklinksIndex {
  const out = emptyIndex();
  if (pageAggregate && pageJobId && pageScannedAt) {
    const slices = buildSectionSlicesFromScan(pageAggregate, pageJobId, pageScannedAt);
    for (const [sitePath, slice] of slices) {
      const s: SiteSection = out.bySite[sitePath] ?? {
        siteLabel: slice.siteLabel,
        siteUrl: slice.siteUrl,
      };
      s.pages = { updatedAt: slice.updatedAt, fromJobId: slice.fromJobId, sources: slice.sources };
      out.bySite[sitePath] = s;
    }
  }
  if (docAggregate && docJobId && docScannedAt) {
    const slices = buildSectionSlicesFromScan(docAggregate, docJobId, docScannedAt);
    for (const [sitePath, slice] of slices) {
      const s: SiteSection = out.bySite[sitePath] ?? {
        siteLabel: slice.siteLabel,
        siteUrl: slice.siteUrl,
      };
      s.documents = { updatedAt: slice.updatedAt, fromJobId: slice.fromJobId, sources: slice.sources };
      out.bySite[sitePath] = s;
    }
  }
  return out;
}

// --- Read path: persistent → in-memory inverted index ---

let cachedInMemory: BacklinksIndex | null = null;
let cachedEtag: string | null = null;
let cachedAt = 0;
/**
 * Promise guard — when one request is already fetching + parsing the
 * index, concurrent requests on the same instance await the same
 * promise instead of each doing their own download. Collapses a burst
 * of N cold-start requests into 1 actual fetch.
 */
let inflightLoad: Promise<BacklinksIndex | null> | null = null;

/**
 * Expand the persistent index into the in-memory inverted form used by
 * the backlinks endpoint and CSV export. Pure — exported for tests.
 */
export function inflateToMemory(persistent: PersistentBacklinksIndex): BacklinksIndex {
  const byCanonicalKey = new Map<string, BacklinkEntry[]>();
  let latestAt = "";

  for (const [sitePath, section] of Object.entries(persistent.bySite)) {
    const push = (slice: SectionSlice, kind: SourceKind): void => {
      if (slice.updatedAt > latestAt) latestAt = slice.updatedAt;
      for (const source of slice.sources) {
        const entry: BacklinkEntry = {
          sourceKind: kind,
          sourceSite: section.siteLabel,
          sourceSitePath: sitePath,
          sourceSiteUrl: section.siteUrl,
          sourceTitle: source.title,
          sourceUrl: source.url,
          sourceUpdatedAt: slice.updatedAt,
        };
        for (const key of source.linksTo) {
          let list = byCanonicalKey.get(key);
          if (!list) {
            list = [];
            byCanonicalKey.set(key, list);
          }
          list.push(entry);
        }
      }
    };
    if (section.pages) push(section.pages, "page");
    if (section.documents) push(section.documents, "document");
  }

  return {
    jobId: "(persistent)",
    scannedAt: latestAt || persistent.builtAt,
    byCanonicalKey,
  };
}

/**
 * Fetch the in-memory inverted index with tiered caching:
 *
 *   1. Fast path (< 30s since last validation) — return cached, no I/O.
 *   2. Validation path (30s – 30min since last validation) — HEAD the
 *      blob, compare ETag. Match → refresh fast timer, return cached.
 *      Mismatch → fall through to full reload.
 *   3. Reload path (> 30min OR ETag changed OR forceRefresh) — download
 *      + parse + cache.
 *
 * Concurrent requests on a cold instance collapse onto a single
 * in-flight promise via `inflightLoad`, so a burst of 50 parallel
 * backlinks lookups does ONE blob download, not 50.
 *
 * Returns null if the persistent blob doesn't exist yet (no scan has
 * run and populated it).
 */
export async function getInMemoryIndex(forceRefresh = false): Promise<BacklinksIndex | null> {
  const now = Date.now();

  // Fast path: recently validated, assume fresh.
  if (!forceRefresh && cachedInMemory && now - cachedAt < CACHE_FAST_TTL_MS) {
    return cachedInMemory;
  }

  // Validation path: cached but stale — HEAD-check the ETag. Only
  // within the validation TTL; older than that, force a full reload
  // to guard against drift if HEAD is unreliable.
  if (!forceRefresh && cachedInMemory && cachedEtag && now - cachedAt < CACHE_HEAD_VALIDATE_TTL_MS) {
    try {
      const freshEtag = await headBlobEtag();
      if (freshEtag && freshEtag === cachedEtag) {
        cachedAt = now;
        return cachedInMemory;
      }
    } catch {
      // HEAD failed — fall through to reload. Safer to pay the reload
      // cost than to return possibly-stale cached data indefinitely.
    }
  }

  // Reload path: coalesce concurrent callers onto one in-flight fetch.
  if (inflightLoad) return inflightLoad;
  inflightLoad = (async (): Promise<BacklinksIndex | null> => {
    try {
      const { body, etag } = await readBlobWithEtag();
      if (!body) {
        // Blob missing — clear cache so subsequent reads retry.
        cachedInMemory = null;
        cachedEtag = null;
        cachedAt = 0;
        return null;
      }
      const inflated = inflateToMemory(body);
      cachedInMemory = inflated;
      cachedEtag = etag;
      cachedAt = Date.now();
      return inflated;
    } finally {
      inflightLoad = null;
    }
  })();
  return inflightLoad;
}

/**
 * Load the raw persistent index (no inversion). Used by the rebuild
 * endpoint to display current state.
 */
export async function loadPersistentIndex(): Promise<PersistentBacklinksIndex | null> {
  const { body } = await readBlobWithEtag();
  return body;
}

// --- Test-only reset ---

/** Reset the in-memory cache. Used by unit tests to avoid cross-test bleed. */
export function _resetBacklinksIndexCache(): void {
  cachedInMemory = null;
  cachedEtag = null;
  cachedAt = 0;
  inflightLoad = null;
}
