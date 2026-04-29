import { TableClient, TableServiceClient } from "@azure/data-tables";
import { BlobServiceClient } from "@azure/storage-blob";
import type { DocumentFileRef } from "./spoFilesEnumerator.js";

/**
 * Persistence for version-history bootstrap jobs.
 *
 * Kept deliberately separate from the link-inventory job table because
 * the shapes diverge (bootstrap has no notion of sites/pages, just a
 * single target library and a per-file manifest). Uses the same storage
 * account and shares the `link-inventory-results` blob container for
 * the file manifest.
 */

const JOBS_TABLE = "DuplicatesBootstrapJobs";
const RESULTS_CONTAINER = "link-inventory-results";
const PARTITION_KEY = "bootstrap";

let jobsTableClient: TableClient | null = null;
let blobServiceClient: BlobServiceClient | null = null;
let tableEnsured = false;
let containerEnsured = false;

function getConnString(): string {
  const cs = process.env.TABLE_CONNECTION_STRING;
  if (!cs) throw new Error("TABLE_CONNECTION_STRING not configured");
  return cs;
}

function getTable(): TableClient {
  if (!jobsTableClient) jobsTableClient = TableClient.fromConnectionString(getConnString(), JOBS_TABLE);
  return jobsTableClient;
}

function getBlob(): BlobServiceClient {
  if (!blobServiceClient) blobServiceClient = BlobServiceClient.fromConnectionString(getConnString());
  return blobServiceClient;
}

async function ensureTable(): Promise<void> {
  if (tableEnsured) return;
  const svc = TableServiceClient.fromConnectionString(getConnString());
  try {
    await svc.createTable(JOBS_TABLE);
  } catch (err) {
    if (!(err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 409)) {
      throw err;
    }
  }
  tableEnsured = true;
}

async function ensureContainer(): Promise<void> {
  if (containerEnsured) return;
  await getBlob().getContainerClient(RESULTS_CONTAINER).createIfNotExists();
  containerEnsured = true;
}

export type BootstrapJobStatus = "queued" | "running" | "completed" | "failed";

export interface BootstrapJob {
  jobId: string;
  status: BootstrapJobStatus;
  startedAt: string;
  finishedAt?: string;
  sitePath: string;
  libraryTitle: string;
  maxVersionsPerFile: number;
  includeOther: boolean;
  filesTotal: number;
  filesCompleted: number;
  versionsProcessed: number;
  errorCount: number;
  recentErrors: string[];
  caller?: string;
  currentFile?: string;
}

interface JobEntity {
  partitionKey: string;
  rowKey: string;
  Status: BootstrapJobStatus;
  StartedAt: string;
  FinishedAt?: string;
  SitePath: string;
  LibraryTitle: string;
  MaxVersionsPerFile: number;
  IncludeOther: boolean;
  FilesTotal: number;
  FilesCompleted: number;
  VersionsProcessed: number;
  ErrorCount: number;
  RecentErrorsJson: string;
  Caller?: string;
  CurrentFile?: string;
}

function toEntity(j: BootstrapJob): JobEntity {
  return {
    partitionKey: PARTITION_KEY,
    rowKey: j.jobId,
    Status: j.status,
    StartedAt: j.startedAt,
    FinishedAt: j.finishedAt,
    SitePath: j.sitePath,
    LibraryTitle: j.libraryTitle,
    MaxVersionsPerFile: j.maxVersionsPerFile,
    IncludeOther: j.includeOther,
    FilesTotal: j.filesTotal,
    FilesCompleted: j.filesCompleted,
    VersionsProcessed: j.versionsProcessed,
    ErrorCount: j.errorCount,
    RecentErrorsJson: JSON.stringify(j.recentErrors),
    Caller: j.caller,
    CurrentFile: j.currentFile,
  };
}

function fromEntity(e: JobEntity): BootstrapJob {
  let recentErrors: string[] = [];
  try { recentErrors = JSON.parse(e.RecentErrorsJson) as string[]; } catch { /* ignore */ }
  return {
    jobId: e.rowKey,
    status: e.Status,
    startedAt: e.StartedAt,
    finishedAt: e.FinishedAt,
    sitePath: e.SitePath,
    libraryTitle: e.LibraryTitle,
    maxVersionsPerFile: e.MaxVersionsPerFile,
    includeOther: e.IncludeOther,
    filesTotal: e.FilesTotal,
    filesCompleted: e.FilesCompleted,
    versionsProcessed: e.VersionsProcessed,
    errorCount: e.ErrorCount,
    recentErrors,
    caller: e.Caller,
    currentFile: e.CurrentFile,
  };
}

export async function createBootstrapJob(init: {
  jobId: string;
  sitePath: string;
  libraryTitle: string;
  maxVersionsPerFile: number;
  includeOther: boolean;
  filesTotal: number;
  caller?: string;
}): Promise<BootstrapJob> {
  await ensureTable();
  const job: BootstrapJob = {
    jobId: init.jobId,
    status: "queued",
    startedAt: new Date().toISOString(),
    sitePath: init.sitePath,
    libraryTitle: init.libraryTitle,
    maxVersionsPerFile: init.maxVersionsPerFile,
    includeOther: init.includeOther,
    filesTotal: init.filesTotal,
    filesCompleted: 0,
    versionsProcessed: 0,
    errorCount: 0,
    recentErrors: [],
    caller: init.caller,
  };
  await getTable().createEntity(toEntity(job));
  return job;
}

export async function updateBootstrapJob(job: BootstrapJob): Promise<void> {
  await ensureTable();
  await getTable().updateEntity(toEntity(job), "Replace");
}

export async function getBootstrapJob(jobId: string): Promise<BootstrapJob | undefined> {
  await ensureTable();
  try {
    const e = await getTable().getEntity<JobEntity>(PARTITION_KEY, jobId);
    return fromEntity(e);
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return undefined;
    }
    throw err;
  }
}

/**
 * Persist the per-file manifest as a side blob under
 * `link-inventory-results/<jobId>/bootstrap-manifest.json`. The worker
 * reads this by fileIndex each message.
 */
export async function writeBootstrapManifest(jobId: string, files: DocumentFileRef[]): Promise<void> {
  await ensureContainer();
  const blob = getBlob().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(`${jobId}/bootstrap-manifest.json`);
  const data = JSON.stringify(files);
  await blob.upload(data, Buffer.byteLength(data), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

export async function readBootstrapManifest(jobId: string): Promise<DocumentFileRef[] | undefined> {
  await ensureContainer();
  const blob = getBlob().getContainerClient(RESULTS_CONTAINER).getBlockBlobClient(`${jobId}/bootstrap-manifest.json`);
  try {
    const dl = await blob.download();
    const chunks: Buffer[] = [];
    for await (const c of dl.readableStreamBody as NodeJS.ReadableStream) {
      chunks.push(typeof c === "string" ? Buffer.from(c) : (c as Buffer));
    }
    return JSON.parse(Buffer.concat(chunks).toString("utf8")) as DocumentFileRef[];
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return undefined;
    }
    throw err;
  }
}
