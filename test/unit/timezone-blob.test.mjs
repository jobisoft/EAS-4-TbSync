/**
 * TimeZoneBlob - the 172-byte packed TIME_ZONE_INFORMATION structure EAS
 * ≤14.x carries base64-coded in <Calendar:TimeZone> - and isAllZero, the
 * discriminator between "no real zone" (the Z-Push family's all-day
 * convention) and a described zone. Pure DataView/atob/btoa, no host
 * environment involved.
 *
 * Cases ported from PR #345 (tomaskovacik), including a blob captured
 * verbatim from a live Exchange debug log - the kind of fixture that
 * catches byte-layout drift no synthetic blob would.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  TimeZoneBlob,
  isAllZero,
} from "../../src/modules/eas/timezone-blob.mjs";

test("round-trip: every field survives set → easTimeZone64 → reload", () => {
  const tz = new TimeZoneBlob();
  tz.utcOffset = -60;
  tz.standardName = "Central Europe Standard Time";
  tz.standardBias = 0;
  tz.daylightName = "Central Europe Daylight Time";
  tz.daylightBias = -60;
  tz.standardDate.wMonth = 10;
  tz.standardDate.wDayOfWeek = 0;
  tz.standardDate.wDay = 5; // "last occurrence in the month"
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

test("assigning a falsy easTimeZone64 zeroes the whole buffer", () => {
  const tz = new TimeZoneBlob();
  tz.standardName = "Some Long Standard Name That Fills Space";
  tz.utcOffset = 120;

  tz.easTimeZone64 = null;
  assert.equal(tz.utcOffset, 0);
  assert.equal(tz.standardName, "");
});

test("a name is truncated to its 32-WCHAR field, never overflowing the neighbour", () => {
  const tz = new TimeZoneBlob();
  tz.standardName = "x".repeat(40);
  assert.equal(tz.standardName, "x".repeat(32));
  // standardBias sits right behind the name block and must be untouched.
  assert.equal(tz.standardBias, 0);
});

test("a blob captured from a live Exchange log decodes field-correct", () => {
  // Verbatim from a TbSync debug log, percent-encoded exactly as it
  // appears in the wire <Calendar:TimeZone> text content (capture by
  // tomaskovacik, PR #345).
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

test("isAllZero: nullish and empty inputs count as all-zero", () => {
  assert.equal(isAllZero(null), true);
  assert.equal(isAllZero(undefined), true);
  assert.equal(isAllZero(""), true);
});

test("isAllZero: a fresh blob is all-zero; one set byte flips it", () => {
  const zero = new TimeZoneBlob();
  assert.equal(isAllZero(zero.easTimeZone64), true);

  const nonZero = new TimeZoneBlob();
  nonZero.utcOffset = 1;
  assert.equal(isAllZero(nonZero.easTimeZone64), false);
});

test("isAllZero: malformed base64 falls back to all-zero rather than throwing", () => {
  // The safe direction: an unreadable blob is treated like "no real
  // zone", which the all-day reader then resolves by value shape.
  assert.equal(isAllZero("not valid base64 !!!"), true);
});
