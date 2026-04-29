import { QueueServiceClient } from "@azure/storage-queue";

/**
 * Storage Queue producer for the document scan worker. Mirrors
 * `linkInventoryQueue.ts` but on a separate queue so the page worker
 * and document worker process independently and don't head-of-line
 * block each other.
 */

export const SCAN_DOCS_QUEUE_NAME = "link-inventory-scan-docs-queue";

/**
 * Two-phase pipeline:
 *   1. `enumerate` — walk one site per message, list parseable files
 *      under the size cap, write a manifest fragment, repeat. When the
 *      last site is done, consolidate fragments into `manifest.json`
 *      and enqueue the first `scan` message.
 *   2. `scan` — open one file per message, extract+classify links,
 *      write a partial result, enqueue next message OR finalize.
 *
 * Splitting like this keeps each message under the Function timeout
 * (a tenant-wide enumeration alone can take minutes for big tenants),
 * and gives the UI live progress through both phases.
 */
export type DocsScanPhase = "enumerate" | "scan";

export interface DocsScanQueueMessage {
  jobId: string;
  /** Phase discriminator. Defaults to "scan" for backwards compat. */
  phase?: DocsScanPhase;
  /** Index into the job's site list. Used in `enumerate` phase. */
  siteIndex?: number;
  /** Index into the file manifest. Used in `scan` phase. */
  fileIndex?: number;
}

let queueClient: ReturnType<QueueServiceClient["getQueueClient"]> | undefined;
let ensured = false;

function getConnString(): string {
  const cs = process.env.AzureWebJobsStorage ?? process.env.TABLE_CONNECTION_STRING;
  if (!cs) throw new Error("AzureWebJobsStorage / TABLE_CONNECTION_STRING not configured");
  return cs;
}

async function getClient(): Promise<ReturnType<QueueServiceClient["getQueueClient"]>> {
  if (!queueClient) {
    const svc = QueueServiceClient.fromConnectionString(getConnString());
    queueClient = svc.getQueueClient(SCAN_DOCS_QUEUE_NAME);
  }
  if (!ensured) {
    await queueClient.createIfNotExists();
    ensured = true;
  }
  return queueClient;
}

export async function enqueueDocsScanMessage(msg: DocsScanQueueMessage): Promise<void> {
  const client = await getClient();
  const json = JSON.stringify(msg);
  const base64 = Buffer.from(json, "utf8").toString("base64");
  await client.sendMessage(base64);
}
