import { QueueServiceClient } from "@azure/storage-queue";

/**
 * Storage Queue producer for the link-inventory scan worker.
 *
 * The queue is processed by the storageQueue trigger function in
 * `linkInventoryScanWorker.ts`. Each message represents one site to
 * scan within a job; the worker enqueues the next message after
 * processing one site, so the chain runs sequentially per job.
 *
 * Why a single queue + sequential chaining instead of fan-out:
 *   - Sequential keeps SP throttling tame (one site at a time, per job)
 *   - Job state updates (counters, currentSite) don't race
 *   - The queue still gives us durability + the ability to resume after
 *     a Function host restart, which the synchronous orchestrator
 *     couldn't
 *
 * If we ever need parallelism, we can switch to fan-out (one message
 * per site enqueued at job-create time) without changing the message
 * shape — just make sure the worker uses transactional update on the
 * job counters.
 */

export const SCAN_QUEUE_NAME = "link-inventory-scan-queue";

export interface ScanQueueMessage {
  jobId: string;
  siteIndex: number;
}

let queueClient: ReturnType<QueueServiceClient["getQueueClient"]> | undefined;
let ensured = false;

function getConnString(): string {
  // Use the function's default storage. Azure Functions always sets
  // AzureWebJobsStorage; locally we fall back to TABLE_CONNECTION_STRING
  // which points at the same account.
  const cs = process.env.AzureWebJobsStorage ?? process.env.TABLE_CONNECTION_STRING;
  if (!cs) throw new Error("AzureWebJobsStorage / TABLE_CONNECTION_STRING not configured");
  return cs;
}

async function getClient(): Promise<ReturnType<QueueServiceClient["getQueueClient"]>> {
  if (!queueClient) {
    const svc = QueueServiceClient.fromConnectionString(getConnString());
    queueClient = svc.getQueueClient(SCAN_QUEUE_NAME);
  }
  if (!ensured) {
    await queueClient.createIfNotExists();
    ensured = true;
  }
  return queueClient;
}

/**
 * Enqueue a scan message. Storage Queue messages are base64-encoded
 * by default for the Functions queue trigger; we encode here on send
 * so the worker's `messageEncoding: 'base64'` (default) decodes it
 * cleanly into the original JSON shape.
 */
export async function enqueueScanMessage(msg: ScanQueueMessage): Promise<void> {
  const client = await getClient();
  const json = JSON.stringify(msg);
  const base64 = Buffer.from(json, "utf8").toString("base64");
  await client.sendMessage(base64);
}
