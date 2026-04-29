import { app, type InvocationContext, type Timer } from "@azure/functions";
import {
  dateInZone,
  getScheduleConfig,
  isScheduleDue,
  markScheduleFired,
} from "../services/scheduleConfigStore.js";
import { triggerUnifiedScan } from "../services/scheduledScanRunner.js";

/**
 * Hourly polling timer that owns the UI-configured daily delta scan.
 *
 * Cron: `0 0 * * * *` — top of every hour. The cron itself is fixed
 * (must be a build-time constant for Azure Functions); the actual
 * fire-or-skip decision happens at runtime by reading the schedule
 * config row written from the redirect-admin web part.
 *
 * Why hourly and not :00 of the scheduled hour: lets admins change
 * `timeOfDay` from the UI without redeploying or waiting until the
 * old time rolls around. Trade-off is up to ~59 minutes of latency
 * between scheduled time and actual fire — acceptable for a daily
 * background scan.
 *
 * Per-day idempotency: `lastFiredOnDate` is the YYYY-MM-DD in the
 * schedule's local zone. Once fired for today, subsequent hourly ticks
 * see the equality and bail. Survives function restarts since it's
 * persisted in `LinkInventoryConfig`.
 *
 * Replaces the old hardcoded weekly schedules
 * (`linkInventoryScheduledDocScan` / `linkInventoryScheduledPageScan`)
 * — those used `SCHEDULED_*_CRON` env vars and required a redeploy to
 * adjust.
 */

async function handler(_timer: Timer, context: InvocationContext): Promise<void> {
  let config;
  try {
    config = await getScheduleConfig();
  } catch (err) {
    context.error(`[schedule-timer] failed to read config: ${(err as Error).message}`);
    return;
  }

  const now = new Date();

  if (!config.enabled) {
    // Quiet — disabled is the default and runs every hour. Verbose log
    // would drown out signal in App Insights.
    return;
  }

  if (!isScheduleDue(config, now)) {
    return;
  }

  const todayLocal = dateInZone(now, config.timeZone);
  context.log(
    `[schedule-timer] due — enabled=true date=${todayLocal} timeOfDay=${config.timeOfDay} zone=${config.timeZone}`,
  );

  const outcome = await triggerUnifiedScan({
    caller: "scheduled",
    skipIfScheduledActive: true,
    // modifiedAfter wiring for delta will plug in once aggregateStore
    // has been built; for now this fires a full scan, which is the
    // safer pre-delta default.
  });

  if ("skipped" in outcome && outcome.skipped) {
    context.warn(`[schedule-timer] skipped: ${outcome.reason}`);
    // Don't mark fired — let the next hourly tick try again. (If a
    // scheduled job is genuinely stuck, retention will purge it after
    // 30 days; admins can also manually delete it.)
    return;
  }

  if (!outcome.ok) {
    context.error(`[schedule-timer] trigger failed: ${outcome.error}`);
    return;
  }

  try {
    await markScheduleFired(todayLocal, outcome.pageJobId, now.toISOString());
  } catch (err) {
    context.error(
      `[schedule-timer] enqueued page=${outcome.pageJobId} doc=${outcome.docJobId} but markFired failed: ${(err as Error).message}`,
    );
    // Best-effort — duplicate fire on next hour is acceptable; far
    // worse to silently never fire again.
    return;
  }

  context.log(
    `[schedule-timer] fired page=${outcome.pageJobId} doc=${outcome.docJobId} sites=${outcome.sitesTotal}${outcome.enumerated ? " (enumerated)" : ""}`,
  );
}

app.timer("linkInventoryScheduleTimer", {
  schedule: "0 0 * * * *",
  handler,
});
