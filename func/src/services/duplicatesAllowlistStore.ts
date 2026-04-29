import { BlobServiceClient } from "@azure/storage-blob";
import {
  emptyAllowlist,
  type DuplicatesAllowlist,
} from "./duplicatesQuery.js";

/**
 * Admin-managed allowlist persistence. Small blob, read-through cache,
 * ETag-optimistic writes. Follows the same shape as hashIndexStore but
 * trimmed — payload is kilobytes at most and reads happen per-request,
 * so the cache is simple wall-clock rather than HEAD-validated.
 */

const RESULTS_CONTAINER = "link-inventory-results";
const BLOB_NAME = "duplicates-allowlist.json";
const CACHE_TTL_MS = 60 * 1000;

let blobServiceClient: BlobServiceClient | null = null;
let containerEnsured = false;
let cached: DuplicatesAllowlist | null = null;
let cachedAt = 0;

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

async function readBlobWithEtag(): Promise<{ body: DuplicatesAllowlist | null; etag: string | null }> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(BLOB_NAME);
  try {
    const dl = await blob.download();
    const chunks: Buffer[] = [];
    for await (const chunk of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Buffer));
    }
    const body = JSON.parse(Buffer.concat(chunks).toString("utf8")) as DuplicatesAllowlist;
    return { body, etag: dl.etag ?? null };
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return { body: null, etag: null };
    }
    throw err;
  }
}

async function writeBlobWithEtag(body: DuplicatesAllowlist, ifMatchEtag: string | null): Promise<void> {
  await ensureContainer();
  const blob = getBlobService().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(BLOB_NAME);
  const data = JSON.stringify(body);
  const conditions = ifMatchEtag ? { ifMatch: ifMatchEtag } : { ifNoneMatch: "*" };
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
    conditions,
  });
}

export async function getAllowlist(forceRefresh = false): Promise<DuplicatesAllowlist> {
  const now = Date.now();
  if (!forceRefresh && cached && now - cachedAt < CACHE_TTL_MS) return cached;
  const { body } = await readBlobWithEtag();
  const next = body ?? emptyAllowlist();
  cached = next;
  cachedAt = now;
  return next;
}

export type AllowlistKind = "hash" | "path" | "name";

export interface AllowlistAddInput {
  kind: AllowlistKind;
  /** For `hash` kind. */
  sha256?: string;
  /** For `path` and `name` kinds. */
  pattern?: string;
  note: string;
  addedBy: string;
}

export async function addAllowlistEntry(input: AllowlistAddInput): Promise<DuplicatesAllowlist> {
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const { body, etag } = await readBlobWithEtag();
    const current = body ?? emptyAllowlist();
    const addedAt = new Date().toISOString();
    const next: DuplicatesAllowlist = {
      hashAllowlist: [...current.hashAllowlist],
      pathAllowlist: [...current.pathAllowlist],
      nameAllowlist: [...current.nameAllowlist],
    };
    if (input.kind === "hash") {
      if (!input.sha256) throw new Error("sha256 required for hash allowlist");
      const lower = input.sha256.toLowerCase();
      if (!next.hashAllowlist.some((h) => h.sha256.toLowerCase() === lower)) {
        next.hashAllowlist.push({ sha256: lower, note: input.note, addedBy: input.addedBy, addedAt });
      }
    } else if (input.kind === "path") {
      if (!input.pattern) throw new Error("pattern required for path allowlist");
      if (!next.pathAllowlist.some((p) => p.pattern === input.pattern)) {
        next.pathAllowlist.push({ pattern: input.pattern, note: input.note, addedBy: input.addedBy, addedAt });
      }
    } else {
      if (!input.pattern) throw new Error("pattern required for name allowlist");
      if (!next.nameAllowlist.some((p) => p.pattern === input.pattern)) {
        next.nameAllowlist.push({ pattern: input.pattern, note: input.note, addedBy: input.addedBy, addedAt });
      }
    }
    try {
      await writeBlobWithEtag(next, etag);
      cached = next;
      cachedAt = Date.now();
      return next;
    } catch (err) {
      const status = err instanceof Error && "statusCode" in err ? (err as { statusCode: number }).statusCode : 0;
      if ((status === 412 || status === 409) && attempt < maxAttempts) continue;
      throw err;
    }
  }
  throw new Error("addAllowlistEntry exhausted retries");
}

export interface AllowlistRemoveInput {
  kind: AllowlistKind;
  /** For `hash`. */
  sha256?: string;
  /** For `path`/`name`. */
  pattern?: string;
}

export async function removeAllowlistEntry(input: AllowlistRemoveInput): Promise<DuplicatesAllowlist> {
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const { body, etag } = await readBlobWithEtag();
    const current = body ?? emptyAllowlist();
    const next: DuplicatesAllowlist = {
      hashAllowlist: [...current.hashAllowlist],
      pathAllowlist: [...current.pathAllowlist],
      nameAllowlist: [...current.nameAllowlist],
    };
    if (input.kind === "hash") {
      if (!input.sha256) throw new Error("sha256 required for hash allowlist");
      const lower = input.sha256.toLowerCase();
      next.hashAllowlist = next.hashAllowlist.filter((h) => h.sha256.toLowerCase() !== lower);
    } else if (input.kind === "path") {
      if (!input.pattern) throw new Error("pattern required for path allowlist");
      next.pathAllowlist = next.pathAllowlist.filter((p) => p.pattern !== input.pattern);
    } else {
      if (!input.pattern) throw new Error("pattern required for name allowlist");
      next.nameAllowlist = next.nameAllowlist.filter((p) => p.pattern !== input.pattern);
    }
    try {
      await writeBlobWithEtag(next, etag);
      cached = next;
      cachedAt = Date.now();
      return next;
    } catch (err) {
      const status = err instanceof Error && "statusCode" in err ? (err as { statusCode: number }).statusCode : 0;
      if ((status === 412 || status === 409) && attempt < maxAttempts) continue;
      throw err;
    }
  }
  throw new Error("removeAllowlistEntry exhausted retries");
}

/** Test-only reset. */
export function _resetAllowlistCache(): void {
  cached = null;
  cachedAt = 0;
}
