/**
 * Unit tests for noticing that the server has thrown a series' exceptions
 * away.
 *
 * [MS-ASCAL] 2.2.2.22 and 2.2.2.42: at 16.0 and 16.1, changing a series'
 * recurrence pattern or its start or end times deletes every exception on
 * the item. The server then restates the series with no `<Exceptions>`, and
 * the inbound merge keeps the ones we hold - it only replaces the set when
 * the ApplicationData mentions it, because a partial echo that says nothing
 * about exceptions must not wipe them. So the two copies silently disagree.
 *
 * What is asserted here is the reading: a restated `<Recurrence>` with no
 * `<Exceptions>` beside it means the server holds the series bare, and the
 * item is queued with an empty baseline so the next push re-asserts the
 * lot. Everything that is NOT that shape must queue nothing - a partial
 * echo above all, which arrives on every ordinary edit and would otherwise
 * re-send every exception of every series forever.
 *
 * Measured live on Exchange 16.1 (suite 6.5): one hour added to a master's
 * end left the server holding neither the three moved occurrences nor the
 * cancellation, and a clean pull returned the bare series.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();

const { noteExceptionsDroppedByServer } = await import(
  new URL("../../src/modules/eas/sync-runner.mjs", import.meta.url)
);

/** A series with three overrides and one cancellation, as stored. */
const BLOB = [
  "BEGIN:VCALENDAR",
  "BEGIN:VEVENT",
  "UID:s@eas-test.invalid",
  "DTSTART:20260907T080000Z",
  "DTEND:20260907T100000Z",
  "RRULE:FREQ=WEEKLY;COUNT=5",
  "EXDATE:20261005T080000Z",
  "END:VEVENT",
  "BEGIN:VEVENT",
  "UID:s@eas-test.invalid",
  "RECURRENCE-ID:20260914T080000Z",
  "DTSTART:20260914T120000Z",
  "END:VEVENT",
  "END:VCALENDAR",
].join("\r\n");

const BARE = BLOB.replace(/BEGIN:VEVENT[\s\S]*?RECURRENCE-ID[\s\S]*?END:VEVENT\r\n/, "")
  .replace("EXDATE:20261005T080000Z\r\n", "");

/** The server restating the whole series - the shape that means "bare". */
const RESTATED = parseAdNode(`
  <ApplicationData>
    <Subject>PROBE digest</Subject>
    <StartTime>20260907T080000Z</StartTime>
    <EndTime>20260907T100000Z</EndTime>
    <Recurrence><Type>1</Type><Interval>1</Interval><Occurrences>5</Occurrences></Recurrence>
  </ApplicationData>`);

/** The same, still carrying its exceptions. */
const WITH_EXCEPTIONS = parseAdNode(`
  <ApplicationData>
    <StartTime>20260907T080000Z</StartTime>
    <Recurrence><Type>1</Type><Interval>1</Interval></Recurrence>
    <Exceptions><Exception><ExceptionStartTime>20260914T080000Z</ExceptionStartTime></Exception></Exceptions>
  </ApplicationData>`);

/** What Exchange echoes after an ordinary edit: a few fields, no rule. */
const PARTIAL_ECHO = parseAdNode(`
  <ApplicationData>
    <DtStamp>20260830T102854Z</DtStamp>
    <Subject>PROBE digest</Subject>
  </ApplicationData>`);

function contextFor({ asVersion = "16.1", syncRecurrence = true, kind = "event" } = {}) {
  const recorded = [];
  return {
    recorded,
    ctx: {
      asVersion,
      syncRecurrence,
      itemKind: { changelogKind: kind },
      targetID: "cal-1",
      accountId: "1",
      folderId: "f-3",
      queue: {
        async record(entry) {
          recorded.push(entry);
        },
      },
      provider: { reportEventLog() {} },
    },
  };
}

const run = async (ad, blob, options) => {
  const { ctx, recorded } = contextFor(options);
  await noteExceptionsDroppedByServer(ctx, ad, { itemId: "s@eas-test.invalid" }, blob);
  return recorded;
};

test("a restated series with no exceptions queues the item to re-assert them", async () => {
  const recorded = await run(RESTATED, BLOB);
  assert.equal(recorded.length, 1, "the divergence was not noticed");
  assert.equal(recorded[0].itemId, "s@eas-test.invalid");
  assert.equal(recorded[0].op, "updated");
  // The baseline is what the server now holds, which is nothing. Anything
  // else and the push would compare against a set that no longer exists
  // and send only the difference - which is none of it.
  assert.deepEqual(recorded[0].detail, { exceptions: { exdates: [], overrides: [] } });
});

test("a series that still carries its exceptions queues nothing", async () => {
  assert.deepEqual(await run(WITH_EXCEPTIONS, BLOB), []);
});

test("a partial echo queues nothing - it says nothing about exceptions", async () => {
  // The one that matters for cost: this arrives on every ordinary edit.
  assert.deepEqual(await run(PARTIAL_ECHO, BLOB), []);
});

test("a series we hold no exceptions for queues nothing", async () => {
  assert.deepEqual(await run(RESTATED, BARE), []);
});

test("below 16.1 queues nothing - exceptions travel embedded there", async () => {
  assert.deepEqual(await run(RESTATED, BLOB, { asVersion: "14.1" }), []);
});

test("an account that does not sync recurrence queues nothing", async () => {
  assert.deepEqual(await run(RESTATED, BLOB, { syncRecurrence: false }), []);
});

test("a task queues nothing - only a calendar has exceptions", async () => {
  assert.deepEqual(await run(RESTATED, BLOB, { kind: "task" }), []);
});
