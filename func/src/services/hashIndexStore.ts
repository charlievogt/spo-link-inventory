import { BlobServiceClient, type BlobDownloadResponseParsed } from "@azure/storage-blob";
import type { HashAlgo } from "./contentHash.js";

/**
 * Compose the cross-file group key used by `byHash`. Two entries are
 * considered exact duplicates only when both their hash AND their algo
 * match — `full-sha256` and `ooxml-content-v1` produce non-comparable
 * digests, so they must never collide in the same bucket.
 */
function groupKey(hash: string, algo: HashAlgo): string {
  return `${algo}:${hash}`;
}

/** Default algo for legacy persisted entries written before the algo field existed. */
const LEGACY_ALGO: HashAlgo = "full-sha256";

/**
 * Persistent file-hash index for duplicate detection.
 *
 * Stored outside any jobId prefix so it survives the 30-day scan-retention
 * purge. Mirrors the pattern of backlinksIndexStore — single blob,
 * ETag-optimistic writes, tiered in-memory cache on reads.
 *
 * Layout (`link-inventory-results/file-hash-index.json`):
 *
 *   {
 *     version: 1,
 *     builtAt: ISO,
 *     byFileRef: {
 *       "/sites/hr/Shared Documents/Handbook.pdf": {
 *         currentHash: "abc123...",
 *         fileName: "Handbook.pdf",
 *         size: 48291,
 *         etag: "\"{...},N\"",
 *         sitePath: "/sites/hr",
 *         library: "Shared Documents",
 *         lastConfirmedAt: "2026-04-22T...",
 *         lastConfirmedByJobId: "uuid",
 *         currentHashObservedAt: "2026-04-22T...",
 *         previousHashes: [
 *           { hash: "def456...", observedAt: "2026-04-01T...", observedByJobId: "uuid" },
 *           ...
 *         ]
 *       }
 *     }
 *   }
 *
 * The cross-file `byHash` view is rebuilt in memory from `byFileRef` on
 * read — keeps the blob smaller and avoids write-side consistency
 * issues with a second index structure.
 */

const RESULTS_CONTAINER = "link-inventory-results";
const INDEX_BLOB_NAME = "file-hash-index.json";

const CACHE_FAST_TTL_MS = 30 * 1000;
const CACHE_HEAD_VALIDATE_TTL_MS = 30 * 60 * 1000;

/**
 * History caps applied on merge. Whichever is tighter wins — the PICKUP
 * spec says "last 20 entries or 2 years, whichever is shorter". A chatty
 * file (daily edits) will fill 20 slots in under a month; a slow-moving
 * one may never hit 20 but will still prune entries older than 2 years.
 */
const HISTORY_MAX_ENTRIES = 20;
const HISTORY_MAX_AGE_MS = 2 * 365 * 24 * 60 * 60 * 1000;

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

// --- Persistent shape ---

export interface HistoricalHash {
  hash: string;
  /**
   * Hash algorithm used to compute `hash`. Defaults to `"full-sha256"`
   * when missing on legacy persisted entries — see `inflateHashIndex`.
   */
  algo: HashAlgo;
  observedAt: string;
  observedByJobId: string;
  /**
   * SHA-256 of the document's normalized visible text at this version.
   * Empty string when the file type has no text-extraction path or the
   * extractor failed. Used for Diverged-pair detection: two files share
   * an ancestor when one of A's previous textHashes equals one of B's.
   */
  textHash?: string;
  /**
   * 64-bit Charikar SimHash of the same normalized text, as a 16-char
   * hex string. Empty/absent when no text was extracted. Used by the
   * Near-duplicate query against current values, and surfaced on
   * historical entries for future use.
   */
  simhash64?: string;
}

export interface FileEntry {
  currentHash: string;
  /**
   * Algorithm used for `currentHash` (and historically — when the algo
   * changes on rescan, the prior currentHash rotates into history with
   * its OWN algo, so the field is per-hash, not per-file).
   */
  algo: HashAlgo;
  fileName: string;
  size: number;
  etag: string;
  sitePath: string;
  library: string;
  lastConfirmedAt: string;
  lastConfirmedByJobId: string;
  currentHashObservedAt: string;
  /** Same semantics as the per-history field, but for the current version. */
  textHash?: string;
  /** Same semantics as the per-history field, but for the current version. */
  simhash64?: string;
  previousHashes: HistoricalHash[];
}

export interface PersistentHashIndex {
  version: 1;
  builtAt: string;
  byFileRef: Record<string, FileEntry>;
}

/**
 * One file's contribution to a scan — what the merge needs to decide
 * whether the hash changed. Builders at the scan-worker layer shape
 * their per-file results into this.
 */
export interface FileObservation {
  fileRef: string;
  fileName: string;
  size: number;
  etag: string;
  sha256: string;
  /**
   * Algorithm that produced `sha256`. Optional for backwards compatibility
   * with callers that haven't been updated yet — when absent, the merge
   * treats it as `"full-sha256"`.
   */
  algo?: HashAlgo;
  /** SHA-256 of normalized visible text. Optional for callers/file types that don't extract text. */
  textHash?: string;
  /** 64-bit SimHash of the same normalized text, 16-char hex. Optional. */
  simhash64?: string;
  sitePath: string;
  library: string;
}

export function emptyHashIndex(): PersistentHashIndex {
  return {
    version: 1,
    builtAt: new Date().toISOString(),
    byFileRef: {},
  };
}

// --- Merge (pure) ---

function pruneHistory(history: HistoricalHash[], nowMs: number): HistoricalHash[] {
  // Drop entries older than HISTORY_MAX_AGE_MS, then cap at the tail.
  // Oldest-first ordering in `previousHashes` means the tail is the
  // most-recently-rotated hashes — what we want to keep.
  const cutoff = nowMs - HISTORY_MAX_AGE_MS;
  const fresh = history.filter((h) => {
    const t = Date.parse(h.observedAt);
    return Number.isFinite(t) ? t >= cutoff : true;
  });
  if (fresh.length <= HISTORY_MAX_ENTRIES) return fresh;
  return fresh.slice(fresh.length - HISTORY_MAX_ENTRIES);
}

/**
 * Merge one scan's observations into the current index. Pure — exported
 * for unit tests.
 *
 * Semantics:
 *   - File seen, hash unchanged → update `lastConfirmedAt/ByJobId` and
 *     mutable metadata (size/etag/fileName/sitePath/library). History
 *     untouched.
 *   - File seen, hash changed → push previous current hash onto
 *     `previousHashes`, set new `currentHash` + `currentHashObservedAt`.
 *   - New file → create entry, empty history.
 *   - File present in the index but NOT in this scan's observations →
 *     untouched (the caller decides scan scope; absence here just means
 *     "this scan didn't cover it").
 */
export function mergeObservationsIntoIndex(
  current: PersistentHashIndex,
  observations: FileObservation[],
  jobId: string,
  observedAt: string,
): PersistentHashIndex {
  const out: PersistentHashIndex = {
    version: 1,
    builtAt: new Date().toISOString(),
    byFileRef: { ...current.byFileRef },
  };
  const nowMs = Date.parse(observedAt);
  const nowValid = Number.isFinite(nowMs) ? nowMs : Date.now();

  for (const obs of observations) {
    if (!obs.fileRef || !obs.sha256) continue;
    const obsAlgo: HashAlgo = obs.algo ?? LEGACY_ALGO;
    const existing = out.byFileRef[obs.fileRef];
    if (!existing) {
      out.byFileRef[obs.fileRef] = {
        currentHash: obs.sha256,
        algo: obsAlgo,
        fileName: obs.fileName,
        size: obs.size,
        etag: obs.etag,
        sitePath: obs.sitePath,
        library: obs.library,
        lastConfirmedAt: observedAt,
        lastConfirmedByJobId: jobId,
        currentHashObservedAt: observedAt,
        textHash: obs.textHash,
        simhash64: obs.simhash64,
        previousHashes: [],
      };
      continue;
    }
    const existingAlgo: HashAlgo = existing.algo ?? LEGACY_ALGO;
    // Algo upgrade (e.g., a .docx that was previously hashed full-sha256
    // is now being hashed ooxml-content-v1): replace silently. Rotating
    // the old hash into history would generate spurious "stale" matches
    // against any other file that happens to share the legacy hash.
    if (existingAlgo !== obsAlgo) {
      out.byFileRef[obs.fileRef] = {
        ...existing,
        currentHash: obs.sha256,
        algo: obsAlgo,
        fileName: obs.fileName,
        size: obs.size,
        etag: obs.etag,
        sitePath: obs.sitePath,
        library: obs.library,
        lastConfirmedAt: observedAt,
        lastConfirmedByJobId: jobId,
        currentHashObservedAt: observedAt,
        textHash: obs.textHash ?? existing.textHash,
        simhash64: obs.simhash64 ?? existing.simhash64,
      };
      continue;
    }
    if (existing.currentHash === obs.sha256) {
      out.byFileRef[obs.fileRef] = {
        ...existing,
        algo: existingAlgo,
        fileName: obs.fileName,
        size: obs.size,
        etag: obs.etag,
        sitePath: obs.sitePath,
        library: obs.library,
        lastConfirmedAt: observedAt,
        lastConfirmedByJobId: jobId,
        // Refresh text/simhash if the observation provides them; otherwise
        // keep what we had. Same-content rescans should produce the same
        // textHash, but text extraction can be lossy so don't overwrite
        // a known-good value with an absent one.
        textHash: obs.textHash ?? existing.textHash,
        simhash64: obs.simhash64 ?? existing.simhash64,
      };
      continue;
    }
    const rotated: HistoricalHash[] = [
      ...existing.previousHashes,
      {
        hash: existing.currentHash,
        algo: existingAlgo,
        observedAt: existing.currentHashObservedAt,
        observedByJobId: existing.lastConfirmedByJobId,
        textHash: existing.textHash,
        simhash64: existing.simhash64,
      },
    ];
    out.byFileRef[obs.fileRef] = {
      currentHash: obs.sha256,
      algo: obsAlgo,
      fileName: obs.fileName,
      size: obs.size,
      etag: obs.etag,
      sitePath: obs.sitePath,
      library: obs.library,
      lastConfirmedAt: observedAt,
      lastConfirmedByJobId: jobId,
      currentHashObservedAt: observedAt,
      textHash: obs.textHash,
      simhash64: obs.simhash64,
      previousHashes: pruneHistory(rotated, nowValid),
    };
  }

  return out;
}

// --- Cross-file projection (pure) ---

/** fileRef with a subset of metadata used by duplicate queries. */
export interface FileRefSummary {
  fileRef: string;
  fileName: string;
  size: number;
  sitePath: string;
  library: string;
  currentHash: string;
  algo: HashAlgo;
  textHash?: string;
  simhash64?: string;
  currentHashObservedAt: string;
  lastConfirmedAt: string;
}

export interface InMemoryHashIndex {
  builtAt: string;
  byFileRef: Map<string, FileEntry>;
  /**
   * `${algo}:${hash}` → file summaries currently at that hash. The
   * composite key prevents `full-sha256` and `ooxml-content-v1` digests
   * from ever colliding into the same group.
   */
  byHash: Map<string, FileRefSummary[]>;
}

export { groupKey as hashGroupKey };

export function inflateHashIndex(persistent: PersistentHashIndex): InMemoryHashIndex {
  const byFileRef = new Map<string, FileEntry>();
  const byHash = new Map<string, FileRefSummary[]>();
  for (const [fileRef, raw] of Object.entries(persistent.byFileRef)) {
    // Backfill missing algo on legacy persisted entries (anything written
    // before contentHash existed). They are full-sha256 by definition.
    const entry: FileEntry = {
      ...raw,
      algo: raw.algo ?? LEGACY_ALGO,
      previousHashes: raw.previousHashes.map((h) => ({ ...h, algo: h.algo ?? LEGACY_ALGO })),
    };
    byFileRef.set(fileRef, entry);
    const summary: FileRefSummary = {
      fileRef,
      fileName: entry.fileName,
      size: entry.size,
      sitePath: entry.sitePath,
      library: entry.library,
      currentHash: entry.currentHash,
      algo: entry.algo,
      textHash: entry.textHash,
      simhash64: entry.simhash64,
      currentHashObservedAt: entry.currentHashObservedAt,
      lastConfirmedAt: entry.lastConfirmedAt,
    };
    const key = groupKey(entry.currentHash, entry.algo);
    let list = byHash.get(key);
    if (!list) {
      list = [];
      byHash.set(key, list);
    }
    list.push(summary);
  }
  return { builtAt: persistent.builtAt, byFileRef, byHash };
}

// --- Blob I/O with ETag optimistic concurrency ---

async function readBlobWithEtag(): Promise<{ body: PersistentHashIndex | null; etag: string | null }> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(INDEX_BLOB_NAME);
  try {
    const dl: BlobDownloadResponseParsed = await blob.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    const body = JSON.parse(Buffer.concat(chunks).toString("utf8")) as PersistentHashIndex;
    return { body, etag: dl.etag ?? null };
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return { body: null, etag: null };
    }
    throw err;
  }
}

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

async function writeBlobWithEtag(body: PersistentHashIndex, ifMatchEtag: string | null): Promise<void> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(INDEX_BLOB_NAME);
  const data = JSON.stringify(body);
  const conditions = ifMatchEtag ? { ifMatch: ifMatchEtag } : { ifNoneMatch: "*" };
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
    conditions,
  });
}

/**
 * Merge a scan's observations into the persistent index with ETag
 * optimistic concurrency. Retries up to 3 times on 412/409.
 */
export async function mergeObservationsToStore(
  observations: FileObservation[],
  jobId: string,
  observedAt: string,
): Promise<void> {
  if (observations.length === 0) return;
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const { body, etag } = await readBlobWithEtag();
    const current = body ?? emptyHashIndex();
    const next = mergeObservationsIntoIndex(current, observations, jobId, observedAt);
    try {
      await writeBlobWithEtag(next, etag);
      cachedInMemory = null;
      cachedEtag = null;
      cachedAt = 0;
      return;
    } catch (err) {
      const status = err instanceof Error && "statusCode" in err ? (err as { statusCode: number }).statusCode : 0;
      if ((status === 412 || status === 409) && attempt < maxAttempts) continue;
      throw err;
    }
  }
}

// --- Tiered in-memory read cache ---

let cachedInMemory: InMemoryHashIndex | null = null;
let cachedEtag: string | null = null;
let cachedAt = 0;
let inflightLoad: Promise<InMemoryHashIndex | null> | null = null;

export async function getInMemoryHashIndex(forceRefresh = false): Promise<InMemoryHashIndex | null> {
  const now = Date.now();
  if (!forceRefresh && cachedInMemory && now - cachedAt < CACHE_FAST_TTL_MS) {
    return cachedInMemory;
  }
  if (!forceRefresh && cachedInMemory && cachedEtag && now - cachedAt < CACHE_HEAD_VALIDATE_TTL_MS) {
    try {
      const freshEtag = await headBlobEtag();
      if (freshEtag && freshEtag === cachedEtag) {
        cachedAt = now;
        return cachedInMemory;
      }
    } catch {
      // fall through to reload
    }
  }
  if (inflightLoad) return inflightLoad;
  inflightLoad = (async (): Promise<InMemoryHashIndex | null> => {
    try {
      const { body, etag } = await readBlobWithEtag();
      if (!body) {
        cachedInMemory = null;
        cachedEtag = null;
        cachedAt = 0;
        return null;
      }
      const inflated = inflateHashIndex(body);
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

export async function loadPersistentHashIndex(): Promise<PersistentHashIndex | null> {
  const { body } = await readBlobWithEtag();
  return body;
}

/**
 * Bootstrap helper: overlay a per-file version history onto the
 * persistent index. Used by the one-shot bootstrap endpoint that reads
 * SP version history and backfills `previousHashes` for a single file.
 *
 * `historicalHashes` must be oldest-first. Merged with any existing
 * history, de-duped on hash, pruned + capped. The current observation
 * refreshes metadata the same way a normal scan would.
 */
export async function overlayVersionHistoryToStore(
  observation: FileObservation,
  historicalHashes: HistoricalHash[],
  jobId: string,
  observedAt: string,
): Promise<void> {
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const { body, etag } = await readBlobWithEtag();
    const current = body ?? emptyHashIndex();
    // Run the normal merge first so current gets created/refreshed.
    const merged = mergeObservationsIntoIndex(current, [observation], jobId, observedAt);
    const entry = merged.byFileRef[observation.fileRef];
    if (!entry) throw new Error(`overlayVersionHistoryToStore: entry missing after merge for ${observation.fileRef}`);

    // Dedup historical hashes by (algo, hash). Two algos can coincidentally
    // produce the same hex string for unrelated digests, so the algo MUST
    // be part of the dedup key.
    const seen = new Set(
      entry.previousHashes.map((h) => `${h.algo}:${h.hash.toLowerCase()}`),
    );
    const currentKey = `${entry.algo}:${entry.currentHash.toLowerCase()}`;
    const combined: HistoricalHash[] = [...entry.previousHashes];
    for (const h of historicalHashes) {
      const key = `${h.algo}:${h.hash.toLowerCase()}`;
      if (key === currentKey) continue;
      if (seen.has(key)) continue;
      seen.add(key);
      combined.push(h);
    }
    combined.sort((a, b) => a.observedAt.localeCompare(b.observedAt));
    const nowMs = Date.parse(observedAt);
    const trimmed = pruneHistory(combined, Number.isFinite(nowMs) ? nowMs : Date.now());

    const next: PersistentHashIndex = {
      version: 1,
      builtAt: new Date().toISOString(),
      byFileRef: {
        ...merged.byFileRef,
        [observation.fileRef]: { ...entry, previousHashes: trimmed },
      },
    };

    try {
      await writeBlobWithEtag(next, etag);
      cachedInMemory = null;
      cachedEtag = null;
      cachedAt = 0;
      return;
    } catch (err) {
      const status = err instanceof Error && "statusCode" in err ? (err as { statusCode: number }).statusCode : 0;
      if ((status === 412 || status === 409) && attempt < maxAttempts) continue;
      throw err;
    }
  }
}

/** Test-only reset. */
export function _resetHashIndexCache(): void {
  cachedInMemory = null;
  cachedEtag = null;
  cachedAt = 0;
  inflightLoad = null;
}
