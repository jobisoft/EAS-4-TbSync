/**
 * A per-instance change for an event that is not a series.
 *
 * A mailbox invited to one occurrence of somebody else's series holds that
 * occurrence as an event of its own, and the server still answers a change
 * to it with an <InstanceId>. The pull built an override on it, Thunderbird
 * refused the result ("Exceptions were supplied to a non-recurring item"),
 * and the refusal failed the whole folder on every sync (#372).
 *
 * The decision is whether the stored item recurs; these pin what counts as
 * one that does, and that the pull asks before it builds an override.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import { storesRecurrence } from "../../src/modules/eas/calendar-codec.mjs";

const cal = (...events) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    ...events.flat(),
    "END:VCALENDAR",
    "",
  ].join("\r\n");

const vevent = (...lines) => [
  "BEGIN:VEVENT",
  "UID:single@eas-test.invalid",
  "DTSTAMP:20261121T120000Z",
  "DTSTART:20261224T160000Z",
  "DTEND:20261224T163000Z",
  "SUMMARY:one occurrence",
  ...lines,
  "END:VEVENT",
];

test("an event that is one occurrence is not a series", () => {
  assert.equal(storesRecurrence(cal(vevent())), false);
});

test("an event with a rule is a series", () => {
  assert.equal(
    storesRecurrence(cal(vevent("RRULE:FREQ=WEEKLY;COUNT=4"))),
    true,
  );
});

test("an event with a list of dates is a series", () => {
  assert.equal(
    storesRecurrence(cal(vevent("RDATE:20261231T160000Z"))),
    true,
  );
});

test("the master is the event without RECURRENCE-ID, wherever it stands", () => {
  const override = vevent("RECURRENCE-ID:20261231T160000Z");
  const master = vevent("RRULE:FREQ=WEEKLY;COUNT=4");
  assert.equal(storesRecurrence(cal(override, master)), true);
});

test("a lone override states no series", () => {
  assert.equal(
    storesRecurrence(cal(vevent("RECURRENCE-ID:20261231T160000Z"))),
    false,
  );
});

test("a blob that cannot be read is no answer, not a no", () => {
  assert.equal(storesRecurrence("not a calendar"), null);
  assert.equal(storesRecurrence(""), null);
  assert.equal(storesRecurrence(undefined), null);
});

test("the pull asks whether the item recurs before it builds an override", async () => {
  // The refusal this prevents happens after the override is built, so the
  // question has to come first - read out of the runner rather than
  // restated here.
  const { readFileSync } = await import("node:fs");
  const src = readFileSync(
    new URL("../../src/modules/eas/sync-runner.mjs", import.meta.url),
    "utf8",
  );
  const change = src.slice(src.indexOf("async function applyChange("));
  const asked = change.indexOf("storesRecurrence");
  const built = change.indexOf("applyExceptionChange(");
  assert.ok(asked > 0, "the pull no longer asks whether the item recurs");
  assert.ok(
    asked < built,
    "the override is built before the pull asks whether there is a series",
  );
});
