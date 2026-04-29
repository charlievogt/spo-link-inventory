import { QueueServiceClient } from "@azure/storage-queue";

/**
 * Storage Queue producer for the duplicates version-history bootstrap
 * worker. One message per file; the worker chains to the next file
 * after hashing all its versions.
 */

export const BOOTSTRAP_QUEUE_NAME = "duplicates-bootstrap-queue";

export interface BootstrapQueueMessage {
  jobId: string;
  /** Index into the manifest blob. */
  fileIndex: number;
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
    queueClient = QueueServiceClient.fromConnectionString(getConnString()).getQueueClient(BOOTSTRAP_QUEUE_NAME);
  }
  if (!ensured) {
    await queueClient.createIfNotExists();
    ensured = true;
  }
  return queueClient;
}

export async function enqueueBootstrapMessage(msg: BootstrapQueueMessage): Promise<void> {
  const client = await getClient();
  const base64 = Buffer.from(JSON.stringify(msg), "utf8").toString("base64");
  await client.sendMessage(base64);
}
