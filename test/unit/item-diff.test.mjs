/**
 * Unit tests for naming what an incoming write changes.
 *
 * The line this feeds is the only trace of a write that came from outside
 * this add-on, and it is filed at info - so it reaches a bug report made
 * with default settings, and must name properties without ever quoting
 * their values. Written for a report (marco, #351-adjacent) where the same
 * event was rewritten eight times in 34 seconds with nothing to say by
 * whom, on a server whose HTTP 500 carries an empty body.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const { differingPropertyNames } = await import(
  "../../src/modules/eas/calendar-codec.mjs"
);

const ics = (...lines) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VEVENT",
    "UID:diff-uid",
    "DTSTAMP:20260801T090000Z",
    "DTSTART:20260810T140000Z",
    "DTEND:20260810T150000Z",
    "SUMMARY:a meeting nobody should read in a log",
    ...lines,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");

test("an identical pair differs in nothing", () => {
  assert.deepEqual(differingPropertyNames(ics(), ics()), []);
});

test("the changed property is named, and its value never appears", () => {
  const before = ics("X-MOZ-LASTACK:20260810T120000Z");
  const after = ics("X-MOZ-LASTACK:20260810T130000Z");
  const names = differingPropertyNames(before, after);
  assert.deepEqual(names, ["x-moz-lastack"]);
  assert.ok(
    !names.join(" ").includes("20260810"),
    "a value must never reach the log",
  );
});

test("added and removed properties count as differences", () => {
  assert.deepEqual(differingPropertyNames(ics(), ics("LOCATION:room 1")), [
    "location",
  ]);
  assert.deepEqual(differingPropertyNames(ics("LOCATION:room 1"), ics()), [
    "location",
  ]);
});

test("our own stamps are excluded - the caller already knows about those", () => {
  // The whole reason the line is printed is that these differ.
  const before = ics("X-EAS-SERVERID:abc");
  const after = ics();
  assert.deepEqual(differingPropertyNames(before, after), []);
});

test("reordered parameters and multi-value order are not changes", () => {
  const a = ics("ATTENDEE;PARTSTAT=ACCEPTED;CN=A:mailto:a@example.invalid",
                "ATTENDEE;CN=B;PARTSTAT=DECLINED:mailto:b@example.invalid");
  const b = ics("ATTENDEE;CN=B;PARTSTAT=DECLINED:mailto:b@example.invalid",
                "ATTENDEE;CN=A;PARTSTAT=ACCEPTED:mailto:a@example.invalid");
  assert.deepEqual(
    differingPropertyNames(a, b),
    [],
    "iCal permits both orders; neither is an edit",
  );
});

test("a change inside an alarm is seen", () => {
  // An acknowledgement can land on the VALARM rather than the event, and
  // a diff that walked only the event's own properties would miss it.
  const withAlarm = (trigger) =>
    ics("BEGIN:VALARM", "ACTION:DISPLAY", `TRIGGER:${trigger}`, "END:VALARM");
  assert.deepEqual(differingPropertyNames(withAlarm("-PT15M"), withAlarm("-PT30M")), [
    "trigger",
  ]);
});

test("a change to one occurrence is not read as a change to the series", () => {
  // Keyed per component: an override moving must not cancel out against
  // the master, nor be reported as the master changing.
  const series = (overrideLocation) =>
    [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//test//EN",
      "BEGIN:VEVENT",
      "UID:diff-uid",
      "DTSTART:20260810T140000Z",
      "RRULE:FREQ=DAILY;COUNT=3",
      "END:VEVENT",
      "BEGIN:VEVENT",
      "UID:diff-uid",
      "RECURRENCE-ID:20260811T140000Z",
      "DTSTART:20260811T160000Z",
      `LOCATION:${overrideLocation}`,
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\r\n");
  assert.deepEqual(differingPropertyNames(series("room 1"), series("room 2")), [
    "location",
  ]);
  assert.deepEqual(differingPropertyNames(series("room 1"), series("room 1")), []);
});

test("an unparseable side is unknown, not 'nothing'", () => {
  // The caller must not report "nothing else differs" - the loudest
  // answer it can give - because a blob failed to parse.
  assert.equal(differingPropertyNames("not an ical at all", ics()), null);
  assert.equal(differingPropertyNames(ics(), ""), null);
});
