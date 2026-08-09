/**
 * Unit tests for the calendar codec's all-day exception form.
 *
 * The rule under test (RFC 5545 §3.8.4.4): RECURRENCE-ID and EXDATE bind
 * to an occurrence by VALUE TYPE as well as by value, so an all-day
 * master - whose DTSTART is a DATE - needs DATE-valued exceptions or
 * Thunderbird binds nothing: the override floats beside the series and a
 * cancelled day keeps rendering. Timed masters keep DATE-TIME, and blobs
 * written before the DATE form (DATE-TIME rows on all-day masters) must
 * still match, or every re-delivery would duplicate them.
 *
 * Run with `npm run test:unit` (node --test). Fixtures stay on
 * UTC-shaped values (`YYYYMMDDT000000Z` fake-local instants) because the
 * zone-conversion branch of `allDayDateFromUtcInstant` needs the
 * extension's timezone service; that branch is shared with DTSTART
 * handling (long since live-tested) and the live suite covers it against
 * a real server.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  applicationDataToIcal,
  applyInstanceChange,
  applyInstanceDelete,
  exceptionFingerprint,
  listInstanceCommands,
} from "../../src/modules/eas/calendar-codec.mjs";

// ── Fabricated ApplicationData nodes ─────────────────────────────────────
// The shape the WBXML decoder hands the codec: {tagName, textContent,
// children}. `el(tag, "text")` is a leaf, `el(tag, [children])` a wrapper.

function el(tagName, value) {
  return Array.isArray(value)
    ? { tagName, children: value }
    : { tagName, textContent: value, children: [] };
}

function allDayAdNode({ exceptions = [] } = {}) {
  return el("ApplicationData", [
    el("UID", "unit-allday@eas-test.invalid"),
    el("Subject", "unit allday master"),
    el("StartTime", "20261012T000000Z"),
    el("EndTime", "20261013T000000Z"),
    el("AllDayEvent", "1"),
    el("Recurrence", [
      el("Type", "0"),
      el("Interval", "1"),
      el("Occurrences", "3"),
    ]),
    ...(exceptions.length ? [el("Exceptions", exceptions)] : []),
  ]);
}

function readerArgs(adNode) {
  return {
    adNode,
    existingIcal: null,
    serverID: "srv-1",
    asVersion: "14.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: null,
    userEmail: null,
    eventLog: null,
  };
}

const propLines = (ical, name) =>
  ical.split(/\r?\n/).filter((l) => l.startsWith(name));

test("14.x embedded exceptions on an all-day master come out as DATEs", () => {
  const ical = applicationDataToIcal(
    readerArgs(
      allDayAdNode({
        exceptions: [
          el("Exception", [
            el("ExceptionStartTime", "20261014T000000Z"),
            el("Deleted", "1"),
          ]),
          el("Exception", [
            el("ExceptionStartTime", "20261013T000000Z"),
            el("Subject", "unit allday OVERRIDDEN"),
            el("StartTime", "20261013T000000Z"),
            el("EndTime", "20261014T000000Z"),
            el("AllDayEvent", "1"),
          ]),
        ],
      }),
    ),
  );
  assert.match(ical, /DTSTART;VALUE=DATE:20261012/, "master is a DATE");
  assert.deepEqual(
    propLines(ical, "EXDATE"),
    ["EXDATE;VALUE=DATE:20261014"],
    "the cancelled day is a DATE on the occurrence grid",
  );
  assert.deepEqual(
    propLines(ical, "RECURRENCE-ID"),
    ["RECURRENCE-ID;VALUE=DATE:20261013"],
    "the override anchors as a DATE on the occurrence grid",
  );
});

test("a timed master keeps DATE-TIME exceptions", () => {
  const node = allDayAdNode({
    exceptions: [
      el("Exception", [
        el("ExceptionStartTime", "20261013T090000Z"),
        el("Deleted", "1"),
      ]),
    ],
  });
  // Same node, timed: flip the flag and give the boundaries a time of day.
  for (const c of node.children) {
    if (c.tagName === "AllDayEvent") c.textContent = "0";
    if (c.tagName === "StartTime") c.textContent = "20261012T090000Z";
    if (c.tagName === "EndTime") c.textContent = "20261012T100000Z";
  }
  const ical = applicationDataToIcal(readerArgs(node));
  assert.deepEqual(
    propLines(ical, "EXDATE"),
    ["EXDATE:20261013T090000Z"],
    "timed events are untouched by the all-day form",
  );
});

// ── The 16.1 per-instance path ───────────────────────────────────────────

const ALLDAY_MASTER = [
  "BEGIN:VCALENDAR",
  "VERSION:2.0",
  "PRODID:-//eas-test//EN",
  "BEGIN:VEVENT",
  "UID:unit-allday@eas-test.invalid",
  "DTSTAMP:20260801T120000Z",
  "SUMMARY:unit allday master",
  "DTSTART;VALUE=DATE:20261012",
  "DTEND;VALUE=DATE:20261013",
  "RRULE:FREQ=DAILY;COUNT=3",
  "END:VEVENT",
  "END:VCALENDAR",
  "",
].join("\r\n");

test("an instance delete on an all-day master writes a DATE EXDATE", () => {
  const out = applyInstanceDelete({
    ical: ALLDAY_MASTER,
    instanceUtc: new Date("2026-10-14T00:00:00Z"),
  });
  assert.deepEqual(propLines(out, "EXDATE"), ["EXDATE;VALUE=DATE:20261014"]);
});

test("a legacy DATE-TIME row still matches - no duplicate, old override replaced", () => {
  // A blob written before the DATE form: all-day master carrying a
  // DATE-TIME EXDATE and a DATE-TIME-anchored override for the same days.
  const legacy = ALLDAY_MASTER.replace(
    "RRULE:FREQ=DAILY;COUNT=3",
    "RRULE:FREQ=DAILY;COUNT=3\r\nEXDATE:20261014T000000Z",
  ).replace(
    "END:VCALENDAR",
    [
      "BEGIN:VEVENT",
      "UID:unit-allday@eas-test.invalid",
      "RECURRENCE-ID:20261013T000000Z",
      "DTSTAMP:20260801T120000Z",
      "SUMMARY:old-form override",
      "DTSTART;VALUE=DATE:20261013",
      "DTEND;VALUE=DATE:20261014",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\r\n"),
  );

  // Re-delivered delete for the day the DATE-TIME EXDATE already names.
  const afterDelete = applyInstanceDelete({
    ical: legacy,
    instanceUtc: new Date("2026-10-14T00:00:00Z"),
  });
  assert.equal(
    propLines(afterDelete, "EXDATE").length,
    1,
    "the legacy row is recognised - a duplicate would push its own <Delete>",
  );

  // Re-delivered change for the occurrence the old-form override anchors:
  // it must replace that override, not sit beside it.
  const afterChange = applyInstanceChange({
    ical: legacy,
    adNode: el("ApplicationData", [
      el("Subject", "new-form override"),
      el("StartTime", "20261013T000000Z"),
      el("EndTime", "20261014T000000Z"),
      el("AllDayEvent", "1"),
    ]),
    instanceUtc: new Date("2026-10-13T00:00:00Z"),
    asVersion: "16.1",
    defaultTimezone: "UTC",
    userEmail: null,
  });
  assert.deepEqual(
    propLines(afterChange, "RECURRENCE-ID"),
    ["RECURRENCE-ID;VALUE=DATE:20261013"],
    "one override, DATE-anchored",
  );
});

test("fingerprint and instance commands agree on DATE rows", () => {
  const withExc = applyInstanceDelete({
    ical: ALLDAY_MASTER,
    instanceUtc: new Date("2026-10-14T00:00:00Z"),
  });
  const fp = exceptionFingerprint(withExc);
  assert.deepEqual(
    fp.exdates,
    ["20261014T000000Z"],
    "a DATE row keys as the 16.1 fake-local form - the wire InstanceId",
  );
  const commands = listInstanceCommands({
    blob: withExc,
    serverID: "srv-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    previous: fp,
  });
  assert.deepEqual(
    commands,
    [],
    "a baseline taken from the same blob reports nothing to send",
  );
});

test("an unknown DATE exdate becomes a 16.1 delete with a fake-local InstanceId", () => {
  const withExc = applyInstanceDelete({
    ical: ALLDAY_MASTER,
    instanceUtc: new Date("2026-10-14T00:00:00Z"),
  });
  const commands = listInstanceCommands({
    blob: withExc,
    serverID: "srv-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    previous: { exdates: [], overrides: [] },
  });
  assert.equal(commands.length, 1);
  assert.equal(commands[0].kind, "delete");
  assert.equal(commands[0].instanceId, "20261014T000000Z");
});
