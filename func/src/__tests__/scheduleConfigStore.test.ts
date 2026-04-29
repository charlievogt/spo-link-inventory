import { describe, it } from "node:test";
import { strict as assert } from "node:assert";
import {
  isScheduleDue,
  dateInZone,
  timeInZone,
  validateScheduleInput,
  ScheduleConfigValidationError,
  type ScheduleConfig,
} from "../services/scheduleConfigStore.js";

function mkConfig(overrides: Partial<ScheduleConfig> = {}): ScheduleConfig {
  return {
    enabled: true,
    timeOfDay: "04:00",
    timeZone: "America/Chicago",
    lastFiredOnDate: "",
    lastJobId: "",
    lastFiredAt: "",
    updatedAt: "",
    updatedBy: "",
    ...overrides,
  };
}

describe("scheduleConfigStore.dateInZone / timeInZone", () => {
  it("formats UTC dates in Central time correctly (DST-aware)", () => {
    // 2026-07-15 12:00 UTC is 2026-07-15 07:00 Central (CDT, UTC-5)
    const summer = new Date("2026-07-15T12:00:00Z");
    assert.equal(dateInZone(summer, "America/Chicago"), "2026-07-15");
    assert.equal(timeInZone(summer, "America/Chicago"), "07:00");

    // 2026-01-15 12:00 UTC is 2026-01-15 06:00 Central (CST, UTC-6)
    const winter = new Date("2026-01-15T12:00:00Z");
    assert.equal(dateInZone(winter, "America/Chicago"), "2026-01-15");
    assert.equal(timeInZone(winter, "America/Chicago"), "06:00");
  });

  it("crosses local-midnight correctly when UTC is the next day", () => {
    // 2026-07-15 03:00 UTC is 2026-07-14 22:00 Central (CDT)
    const lateNight = new Date("2026-07-15T03:00:00Z");
    assert.equal(dateInZone(lateNight, "America/Chicago"), "2026-07-14");
    assert.equal(timeInZone(lateNight, "America/Chicago"), "22:00");
  });
});

describe("scheduleConfigStore.isScheduleDue", () => {
  it("returns false when disabled", () => {
    const cfg = mkConfig({ enabled: false });
    const now = new Date("2026-07-15T12:00:00Z"); // 07:00 Central — past 04:00
    assert.equal(isScheduleDue(cfg, now), false);
  });

  it("returns true at 04:00 Central on a fresh day", () => {
    const cfg = mkConfig({ timeOfDay: "04:00" });
    const now = new Date("2026-07-15T09:00:00Z"); // 04:00 CDT
    assert.equal(isScheduleDue(cfg, now), true);
  });

  it("returns true any time after the scheduled hour on a fresh day", () => {
    const cfg = mkConfig({ timeOfDay: "04:00" });
    const now = new Date("2026-07-15T18:00:00Z"); // 13:00 CDT — well past 04:00
    assert.equal(isScheduleDue(cfg, now), true);
  });

  it("returns false before the scheduled hour", () => {
    const cfg = mkConfig({ timeOfDay: "04:00" });
    const now = new Date("2026-07-15T08:30:00Z"); // 03:30 CDT — before 04:00
    assert.equal(isScheduleDue(cfg, now), false);
  });

  it("returns false once today's date matches lastFiredOnDate", () => {
    const cfg = mkConfig({ timeOfDay: "04:00", lastFiredOnDate: "2026-07-15" });
    const now = new Date("2026-07-15T18:00:00Z"); // 13:00 CDT
    assert.equal(isScheduleDue(cfg, now), false);
  });

  it("re-fires the next day after a previous fire", () => {
    const cfg = mkConfig({ timeOfDay: "04:00", lastFiredOnDate: "2026-07-14" });
    const now = new Date("2026-07-15T09:00:00Z"); // 04:00 CDT next day
    assert.equal(isScheduleDue(cfg, now), true);
  });

  it("respects DST transition — same wall-clock fire works in winter and summer", () => {
    const cfg = mkConfig({ timeOfDay: "04:00" });
    // Summer: 04:00 Central = 09:00 UTC
    assert.equal(isScheduleDue(cfg, new Date("2026-07-15T09:00:00Z")), true);
    // Winter: 04:00 Central = 10:00 UTC
    assert.equal(isScheduleDue(cfg, new Date("2026-01-15T10:00:00Z")), true);
    // Winter at 09:00 UTC = 03:00 Central — too early
    assert.equal(isScheduleDue(cfg, new Date("2026-01-15T09:00:00Z")), false);
  });

  it("returns false on malformed timeOfDay (defensive — bad data shouldn't fire)", () => {
    const cfg = mkConfig({ timeOfDay: "garbage" });
    const now = new Date("2026-07-15T18:00:00Z");
    assert.equal(isScheduleDue(cfg, now), false);
  });
});

describe("scheduleConfigStore.validateScheduleInput", () => {
  it("accepts valid input", () => {
    validateScheduleInput({ enabled: true, timeOfDay: "04:00", timeZone: "America/Chicago" });
    validateScheduleInput({ enabled: false });
    validateScheduleInput({ timeOfDay: "23:59" });
    validateScheduleInput({ timeOfDay: "00:00" });
  });

  it("rejects malformed timeOfDay", () => {
    assert.throws(() => validateScheduleInput({ timeOfDay: "4:00" }), ScheduleConfigValidationError);
    assert.throws(() => validateScheduleInput({ timeOfDay: "24:00" }), ScheduleConfigValidationError);
    assert.throws(() => validateScheduleInput({ timeOfDay: "04:60" }), ScheduleConfigValidationError);
    assert.throws(() => validateScheduleInput({ timeOfDay: "garbage" }), ScheduleConfigValidationError);
  });

  it("rejects unknown IANA time zone", () => {
    assert.throws(
      () => validateScheduleInput({ timeZone: "Mars/Olympus_Mons" }),
      ScheduleConfigValidationError,
    );
  });

  it("rejects empty time zone", () => {
    assert.throws(() => validateScheduleInput({ timeZone: "" }), ScheduleConfigValidationError);
  });

  it("rejects non-boolean enabled", () => {
    assert.throws(
      () => validateScheduleInput({ enabled: "true" as unknown as boolean }),
      ScheduleConfigValidationError,
    );
  });
});
