import { TableClient, TableServiceClient } from "@azure/data-tables";

/**
 * Persistence for the link-inventory daily-scan schedule.
 *
 * Single-row config in Azure Table `LinkInventoryConfig`:
 *   - PartitionKey: "config"
 *   - RowKey: "schedule"
 *
 * The hourly schedule timer (`linkInventoryScheduleTimer`) reads this
 * row to decide whether to enqueue a delta scan. The redirect-admin web
 * part reads/writes it via /api/link-inventory/schedule.
 *
 * Time zone is intentionally stored alongside the time string so the
 * "next due" calculation can be done in the user's local zone (default
 * America/Chicago) regardless of which Azure region the Function runs
 * in. We never store the cron expression itself — the timer is fixed at
 * hourly and the row controls what hour-of-day the scan should fire.
 */

const TABLE = "LinkInventoryConfig";
const PARTITION_KEY = "config";
const ROW_KEY = "schedule";

let tableClient: TableClient | null = null;
let tableEnsured = false;

function getConnString(): string {
  const cs = process.env.TABLE_CONNECTION_STRING;
  if (!cs) throw new Error("TABLE_CONNECTION_STRING not configured");
  return cs;
}

function getTable(): TableClient {
  if (!tableClient) tableClient = TableClient.fromConnectionString(getConnString(), TABLE);
  return tableClient;
}

async function ensureTable(): Promise<void> {
  if (tableEnsured) return;
  const svc = TableServiceClient.fromConnectionString(getConnString());
  try {
    await svc.createTable(TABLE);
  } catch (err) {
    if (!(err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 409)) {
      throw err;
    }
  }
  tableEnsured = true;
}

export interface ScheduleConfig {
  /** Master switch — when false the hourly timer is a no-op. */
  enabled: boolean;
  /** "HH:MM" 24-hour string in `timeZone`. */
  timeOfDay: string;
  /** IANA TZ name. Defaults to America/Chicago. */
  timeZone: string;
  /**
   * Date (YYYY-MM-DD) of the last day the schedule successfully fired,
   * computed in `timeZone`. Used to enforce "fire at most once per local
   * calendar day" semantics — survives function restarts. Empty string
   * before the first run.
   */
  lastFiredOnDate: string;
  /**
   * Job ID of the last scheduled run, for the UI to link back to.
   * Empty string before the first run.
   */
  lastJobId: string;
  /** ISO timestamp the last fire was enqueued at. Empty string initially. */
  lastFiredAt: string;
  /** ISO timestamp of last config edit (UI-visible audit). */
  updatedAt: string;
  /** UPN of the last admin who edited (UI-visible audit). */
  updatedBy: string;
}

const DEFAULT_CONFIG: ScheduleConfig = Object.freeze({
  enabled: false,
  timeOfDay: "04:00",
  timeZone: "America/Chicago",
  lastFiredOnDate: "",
  lastJobId: "",
  lastFiredAt: "",
  updatedAt: "",
  updatedBy: "",
});

/**
 * Validation — keep the surface tight. The endpoints layer relies on
 * this throwing on bad input rather than letting malformed config
 * propagate into the schedule timer.
 */
const TIME_OF_DAY_RE = /^([01]\d|2[0-3]):[0-5]\d$/;

export class ScheduleConfigValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ScheduleConfigValidationError";
  }
}

export interface ScheduleConfigInput {
  enabled?: boolean;
  timeOfDay?: string;
  timeZone?: string;
}

export function validateScheduleInput(input: ScheduleConfigInput): void {
  if (input.timeOfDay !== undefined) {
    if (typeof input.timeOfDay !== "string" || !TIME_OF_DAY_RE.test(input.timeOfDay)) {
      throw new ScheduleConfigValidationError("timeOfDay must be in HH:MM 24-hour format");
    }
  }
  if (input.timeZone !== undefined) {
    if (typeof input.timeZone !== "string" || input.timeZone.trim() === "") {
      throw new ScheduleConfigValidationError("timeZone must be a non-empty string");
    }
    // Validate via Intl — throws RangeError for an unknown zone.
    try {
      new Intl.DateTimeFormat("en-US", { timeZone: input.timeZone }).format(0);
    } catch {
      throw new ScheduleConfigValidationError(`timeZone "${input.timeZone}" is not a valid IANA name`);
    }
  }
  if (input.enabled !== undefined && typeof input.enabled !== "boolean") {
    throw new ScheduleConfigValidationError("enabled must be a boolean");
  }
}

/**
 * Read the current schedule config. Returns the default config (disabled,
 * 04:00 Central) when no row exists yet — the timer treats default
 * `enabled: false` as "do nothing", so a missing row is safe.
 */
export async function getScheduleConfig(): Promise<ScheduleConfig> {
  await ensureTable();
  try {
    const entity = await getTable().getEntity<{
      enabled: boolean;
      timeOfDay: string;
      timeZone: string;
      lastFiredOnDate: string;
      lastJobId: string;
      lastFiredAt: string;
      updatedAt: string;
      updatedBy: string;
    }>(PARTITION_KEY, ROW_KEY);
    return {
      enabled: !!entity.enabled,
      timeOfDay: entity.timeOfDay || DEFAULT_CONFIG.timeOfDay,
      timeZone: entity.timeZone || DEFAULT_CONFIG.timeZone,
      lastFiredOnDate: entity.lastFiredOnDate || "",
      lastJobId: entity.lastJobId || "",
      lastFiredAt: entity.lastFiredAt || "",
      updatedAt: entity.updatedAt || "",
      updatedBy: entity.updatedBy || "",
    };
  } catch (err) {
    if (err instanceof Error && "statusCode" in err && (err as { statusCode: number }).statusCode === 404) {
      return { ...DEFAULT_CONFIG };
    }
    throw err;
  }
}

/**
 * Persist a partial config update. Merges with current values — fields
 * not in `input` keep their existing value. Validation happens before
 * any I/O.
 */
export async function updateScheduleConfig(
  input: ScheduleConfigInput,
  updatedBy: string,
): Promise<ScheduleConfig> {
  validateScheduleInput(input);
  const current = await getScheduleConfig();
  const next: ScheduleConfig = {
    ...current,
    ...(input.enabled !== undefined && { enabled: input.enabled }),
    ...(input.timeOfDay !== undefined && { timeOfDay: input.timeOfDay }),
    ...(input.timeZone !== undefined && { timeZone: input.timeZone }),
    updatedAt: new Date().toISOString(),
    updatedBy,
  };
  await getTable().upsertEntity(
    {
      partitionKey: PARTITION_KEY,
      rowKey: ROW_KEY,
      ...next,
    },
    "Replace",
  );
  return next;
}

/**
 * Mark a scheduled run as fired. Called by the timer after it
 * successfully enqueues the delta scan. Updates lastFiredOnDate +
 * lastJobId + lastFiredAt without touching enabled/timeOfDay/timeZone.
 *
 * `firedOnDate` is the YYYY-MM-DD string in the schedule's `timeZone`,
 * not UTC — the timer computes it before calling this.
 */
export async function markScheduleFired(
  firedOnDate: string,
  jobId: string,
  firedAt: string,
): Promise<void> {
  const current = await getScheduleConfig();
  const next: ScheduleConfig = {
    ...current,
    lastFiredOnDate: firedOnDate,
    lastJobId: jobId,
    lastFiredAt: firedAt,
  };
  await getTable().upsertEntity(
    {
      partitionKey: PARTITION_KEY,
      rowKey: ROW_KEY,
      ...next,
    },
    "Replace",
  );
}

// --- Pure due-check (exported for unit tests) ---

/**
 * Format a Date as YYYY-MM-DD in the given IANA time zone. Uses
 * `Intl.DateTimeFormat` so DST is handled correctly without pulling in
 * a date library.
 */
export function dateInZone(d: Date, timeZone: string): string {
  const fmt = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  return fmt.format(d); // en-CA → "YYYY-MM-DD"
}

/**
 * Format a Date as HH:MM in the given IANA time zone (24-hour).
 */
export function timeInZone(d: Date, timeZone: string): string {
  const fmt = new Intl.DateTimeFormat("en-GB", {
    timeZone,
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  return fmt.format(d); // en-GB → "HH:MM"
}

/**
 * Decide whether the schedule is due to fire NOW.
 *
 * Logic: enabled AND (today's local-zone date != lastFiredOnDate) AND
 * (now's local-zone time-of-day >= timeOfDay). The hourly timer calls
 * this on every tick; the lastFiredOnDate gate ensures at most one fire
 * per local calendar day.
 *
 * Pure — exported for unit tests.
 */
export function isScheduleDue(config: ScheduleConfig, now: Date): boolean {
  if (!config.enabled) return false;
  if (!TIME_OF_DAY_RE.test(config.timeOfDay)) return false;
  const todayLocal = dateInZone(now, config.timeZone);
  if (todayLocal === config.lastFiredOnDate) return false;
  const nowTimeLocal = timeInZone(now, config.timeZone);
  return nowTimeLocal >= config.timeOfDay;
}
