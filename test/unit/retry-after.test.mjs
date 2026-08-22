/**
 * Unit tests for reading a 503's Retry-After header.
 *
 * [MS-ASCMD] 2.2.2: a pre-14.0 server answers HTTP 503 where 14.0+
 * answers Status 111 (ServerErrorRetryLater), so the transport code is a
 * protocol statement, not noise - and RFC 9110 lets the server say how
 * long, as delay-seconds or an HTTP-date. The parse is clamped: a pause
 * under a minute is not worth recording, and a confused server must not
 * silence an account for a day on one header nobody can see.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const { parseRetryAfterMs } = await import("../../src/modules/network.mjs");

test("delay-seconds is read as milliseconds", () => {
  assert.equal(parseRetryAfterMs("120"), 120_000);
  assert.equal(parseRetryAfterMs(" 300 "), 300_000, "whitespace tolerated");
});

test("an HTTP-date is read relative to now", () => {
  const inTenMin = new Date(Date.now() + 10 * 60_000).toUTCString();
  const ms = parseRetryAfterMs(inTenMin);
  // toUTCString drops sub-second precision, so allow a little drift.
  assert.ok(Math.abs(ms - 10 * 60_000) < 2_000, `got ${ms}`);
});

test("the clamp holds at both ends", () => {
  assert.equal(parseRetryAfterMs("5"), 60_000, "below a minute rounds up");
  assert.equal(
    parseRetryAfterMs(String(48 * 60 * 60)),
    4 * 60 * 60 * 1000,
    "two days is capped at four hours",
  );
  const past = new Date(Date.now() - 60_000).toUTCString();
  assert.equal(parseRetryAfterMs(past), 60_000, "a date in the past too");
});

test("absent or unreadable answers null, never a guess", () => {
  assert.equal(parseRetryAfterMs(null), null);
  assert.equal(parseRetryAfterMs(""), null);
  assert.equal(parseRetryAfterMs("soon"), null);
});
