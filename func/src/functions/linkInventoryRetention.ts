import { app, type InvocationContext, type Timer } from "@azure/functions";
import { listJobs, purgeJob } from "../services/linkInventoryJobStore.js";

/**
 * Daily retention sweep for link-inventory scan jobs.
 *
 * Walks the job table and purges any job whose `startedAt` is older
 * than `RETENTION_DAYS` (default 30). Each purge removes the table
 * row, the aggregate result blob, and every per-site partial / file
 * manifest / enum fragment associated with the job.
 *
 * Schedule: daily at 03:00 UTC. The CRON expression is six fields
 * (Functions extension v4 format): `{second} {minute} {hour} {day}
 * {month} {dayOfWeek}`. We pick 03:00 UTC because that's outside the
 * typical scan-trigger window (admins kick scans during business
 * hours), so the sweep doesn't compete with running jobs for table
 * throughput.
 *
 * Soft caps:
 *   - We never purge a job whose status is `running` or `queued`,
 *     even if the row is mysteriously older than the retention
 *     window. Better to leak a row than to half-delete a live job.
 *   - We log a summary at the end so the App Insights trace history
 *     can confirm the sweep ran.
 */

const RETENTION_DAYS = 30;
const RETENTION_MS = RETENTION_DAYS * 24 * 60 * 60 * 1000;

async function retentionHandler(_timer: Timer, context: InvocationContext): Promise<void> {
  const cutoff = Date.now() - RETENTION_MS;
  context.log(
    `[retention] sweep start — purging jobs with startedAt < ${new Date(cutoff).toISOString()}`,
  );

  // listJobs caps at 50 by default. Use a high limit so a long-lived
  // tenant doesn't accumulate untouched rows past the cap.
  const jobs = await listJobs(1000);

  let considered = 0;
  let skippedActive = 0;
  let purged = 0;
  let purgeErrors = 0;
  let blobsDeletedTotal = 0;

  for (const job of jobs) {
    considered += 1;
    const startedMs = Date.parse(job.startedAt);
    if (Number.isNaN(startedMs) || startedMs >= cutoff) continue;

    if (job.status === "running" || job.status === "queued") {
      skippedActive += 1;
      context.warn(
        `[retention] skipping ${job.jobId} — status=${job.status} but startedAt is older than retention window`,
      );
      continue;
    }

    try {
      const result = await purgeJob(job.jobId);
      purged += 1;
      blobsDeletedTotal += result.blobsDeleted;
      context.log(
        `[retention] purged ${job.jobId} (${job.kind ?? "pages"}, started ${job.startedAt}) ` +
        `— rowDeleted=${result.rowDeleted}, blobsDeleted=${result.blobsDeleted}`,
      );
    } catch (err) {
      purgeErrors += 1;
      context.error(
        `[retention] purge failed for ${job.jobId}: ${(err as Error).message}`,
      );
    }
  }

  context.log(
    `[retention] sweep done — considered=${considered}, purged=${purged}, ` +
    `blobsDeleted=${blobsDeletedTotal}, skippedActive=${skippedActive}, errors=${purgeErrors}`,
  );
}

app.timer("linkInventoryRetention", {
  // Six-field CRON: second minute hour day month dayOfWeek
  // 0 0 3 * * * → daily at 03:00:00 UTC
  schedule: "0 0 3 * * *",
  handler: retentionHandler,
});
