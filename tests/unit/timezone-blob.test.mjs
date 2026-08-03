// TimeZoneBlob (the 172-byte packed VTIMEZONE blob EAS <=14.x carries as
// <Calendar:Timezone>) and isAllZero. No WebExtension host dependency
// at all here - pure DataView/atob/btoa - so no shim needed.

import { test } from "vitest";
import assert from "node:assert/strict";
import { TimeZoneBlob, isAllZero } from "../../src/modules/eas/timezone-blob.mjs";

test("round-trip: every field written survives a set/get(easTimeZone64)/reload cycle", () => {
  const tz = new TimeZoneBlob();
  tz.utcOffset = -60;
  tz.standardName = "Central Europe Standard Time";
  tz.standardBias = 0;
  tz.daylightName = "Central Europe Daylight Time";
  tz.daylightBias = -60;
  tz.standardDate.wMonth = 10;
  tz.standardDate.wDayOfWeek = 0;
  tz.standardDate.wDay = 5; // "last" occurrence
  tz.standardDate.wHour = 3;
  tz.daylightDate.wMonth = 3;
  tz.daylightDate.wDayOfWeek = 0;
  tz.daylightDate.wDay = 5;
  tz.daylightDate.wHour = 2;

  const reloaded = new TimeZoneBlob();
  reloaded.easTimeZone64 = tz.easTimeZone64;

  assert.equal(reloaded.utcOffset, -60);
  assert.equal(reloaded.standardName, "Central Europe Standard Time");
  assert.equal(reloaded.standardBias, 0);
  assert.equal(reloaded.daylightName, "Central Europe Daylight Time");
  assert.equal(reloaded.daylightBias, -60);
  assert.equal(reloaded.standardDate.wMonth, 10);
  assert.equal(reloaded.standardDate.wDay, 5);
  assert.equal(reloaded.standardDate.wHour, 3);
  assert.equal(reloaded.daylightDate.wMonth, 3);
  assert.equal(reloaded.daylightDate.wHour, 2);
});

test("setting easTimeZone64 clears any previously-set bytes first (no stale leftovers)", () => {
  const tz = new TimeZoneBlob();
  tz.standardName = "Some Long Standard Name That Fills Space";
  tz.utcOffset = 120;

  tz.easTimeZone64 = null; // per the setter, falsy input zeroes the buffer
  assert.equal(tz.utcOffset, 0);
  assert.equal(tz.standardName, "");
});

test("a name is truncated to 32 UTF-16 code units, not overflowed into the next field", () => {
  const tz = new TimeZoneBlob();
  const tooLong = "x".repeat(40);
  tz.standardName = tooLong;
  assert.equal(tz.standardName, "x".repeat(32));
  // daylightBias (next field after standardName+standardDate+standardBias)
  // must not have been clobbered by the overflow.
  assert.equal(tz.standardBias, 0);
});

test("real captured blob (kovacik@dgtfactory.com, live Exchange, Central Europe zone) decodes correctly", () => {
  // Captured verbatim from a TbSync debug log, percent-encoded exactly as
  // it appears in the wire <Calendar:Timezone> text content.
  const wireText =
    "xP%2F%2F%2F0MAZQBuAHQAcgBhAGwAIABFAHUAcgBvAHAAZQAgAFMAdABhAG4AZABhAHIAZAAgAFQAaQBtAGUAAAAAAAAAAAAAAAoAAAAFAAMAAAAAAAAAAAAAACgAVQBUAEMAKwAwADEAOgAwADAAKQAgAEIAZQBsAGcAcgBhAGQAZQAsACAAQgByAGEAdABpAHMAbABhAHYAYQAAAAMAAAAFAAIAAAAAAAAAxP%2F%2F%2Fw%3D%3D";
  const tz = new TimeZoneBlob();
  tz.easTimeZone64 = decodeURIComponent(wireText);

  assert.equal(tz.utcOffset, -60);
  assert.equal(tz.standardName, "Central Europe Standard Time");
  assert.equal(tz.standardBias, 0);
  assert.equal(tz.daylightName, "(UTC+01:00) Belgrade, Bratislava");
  assert.equal(tz.daylightBias, -60);
});

test("isAllZero: true for null/undefined/empty", () => {
  assert.equal(isAllZero(null), true);
  assert.equal(isAllZero(undefined), true);
  assert.equal(isAllZero(""), true);
});

test("isAllZero: true for a genuinely all-zero blob, false once any byte is set", () => {
  const zero = new TimeZoneBlob();
  assert.equal(isAllZero(zero.easTimeZone64), true);

  const nonZero = new TimeZoneBlob();
  nonZero.utcOffset = 1;
  assert.equal(isAllZero(nonZero.easTimeZone64), false);
});

test("isAllZero: malformed base64 is treated as all-zero (safe fallback), not a throw", () => {
  assert.equal(isAllZero("not valid base64 !!!"), true);
});
