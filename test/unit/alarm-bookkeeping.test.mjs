/**
 * A dismissed reminder is not an edit.
 *
 * Thunderbird keeps the answer to "has this reminder been dealt with" inside
 * the item - `X-MOZ-SNOOZE-TIME` when you snooze, `X-MOZ-LASTACK` when you
 * dismiss - and saves it through the ordinary write path, on the parent, so
 * acknowledging one occurrence's alarm rewrites the whole series. Queued as
 * a user edit it leaves as a `<Change>`, and a server that does the
 * scheduling then mails every attendee to say the meeting changed. It
 * carries nothing: no outbound codec maps either property.
 *
 * These fix the shape of that judgement. The dangerous direction is the
 * false positive - a real edit swallowed here is gone with no trace - so the
 * cases that must stay queued outnumber the ones that must not.
 *
 * Reported as TbSync #816.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const { onlyAlarmBookkeeping } = await import(
  "../../src/modules/calendar-provider.mjs"
);

/** A meeting with a reminder, as the calendar holds it before the alarm
 *  fires. `extra` lands inside the VEVENT, `stamp` is what every save
 *  rewrites. */
const event = ({ extra = [], stamp = "20260801T090000Z", summary = "standup" } = {}) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VEVENT",
    "UID:alarm-uid",
    `DTSTAMP:${stamp}`,
    `LAST-MODIFIED:${stamp}`,
    "DTSTART:20260810T140000Z",
    "DTEND:20260810T150000Z",
    `SUMMARY:${summary}`,
    "BEGIN:VALARM",
    "ACTION:DISPLAY",
    "TRIGGER:-PT15M",
    "DESCRIPTION:Reminder",
    "END:VALARM",
    ...extra,
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");

const task = ({ extra = [], stamp = "20260801T090000Z", summary = "file it" } = {}) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VTODO",
    "UID:alarm-todo",
    `DTSTAMP:${stamp}`,
    `LAST-MODIFIED:${stamp}`,
    "DUE:20260810T140000Z",
    `SUMMARY:${summary}`,
    ...extra,
    "END:VTODO",
    "END:VCALENDAR",
    "",
  ].join("\r\n");

/** Every save moves these, so a realistic "after" carries a later stamp. */
const LATER = "20260810T134500Z";

test("a dismissed reminder is not queued", () => {
  const before = event();
  const after = event({ extra: [`X-MOZ-LASTACK:${LATER}`], stamp: LATER });
  assert.equal(onlyAlarmBookkeeping(before, after), true);
});

test("a snoozed reminder is not queued", () => {
  const before = event();
  const after = event({
    extra: ["X-MOZ-SNOOZE-TIME:20260810T140000Z"],
    stamp: LATER,
  });
  assert.equal(onlyAlarmBookkeeping(before, after), true);
});

test("snoozing one occurrence is not queued - the name carries its id", () => {
  const before = event();
  const after = event({
    extra: ["X-MOZ-SNOOZE-TIME-1793558400000000:20260810T140000Z"],
    stamp: LATER,
  });
  assert.equal(onlyAlarmBookkeeping(before, after), true);
});

test("dismissing a snoozed reminder is not queued, though it moves both", () => {
  // What `dismissAlarm` actually writes: the acknowledgement set and the
  // snooze it supersedes removed, in one save.
  const before = event({ extra: ["X-MOZ-SNOOZE-TIME:20260810T140000Z"] });
  const after = event({ extra: [`X-MOZ-LASTACK:${LATER}`], stamp: LATER });
  assert.equal(onlyAlarmBookkeeping(before, after), true);
});

test("the same judgement on a task", () => {
  const before = task();
  const after = task({ extra: [`X-MOZ-LASTACK:${LATER}`], stamp: LATER });
  assert.equal(onlyAlarmBookkeeping(before, after), true);
});

/* ── and everything that must still reach the server ──────────────────── */

test("a retitled event is queued, dismissed reminder or not", () => {
  const before = event();
  const after = event({
    extra: [`X-MOZ-LASTACK:${LATER}`],
    stamp: LATER,
    summary: "standup, moved room",
  });
  assert.equal(onlyAlarmBookkeeping(before, after), false);
});

test("a retitled event on its own is queued", () => {
  assert.equal(
    onlyAlarmBookkeeping(event(), event({ summary: "renamed", stamp: LATER })),
    false,
  );
});

test("changing the reminder itself is queued - the server keeps that", () => {
  // The alarm's trigger is content: it round-trips as <Reminder>. Only the
  // acknowledgement is local.
  const before = event();
  const after = event({ stamp: LATER }).replace("TRIGGER:-PT15M", "TRIGGER:-PT5M");
  assert.equal(onlyAlarmBookkeeping(before, after), false);
});

test("a write that changed nothing at all is left alone", () => {
  assert.equal(onlyAlarmBookkeeping(event(), event()), false);
});

test("an item that will not parse is queued rather than swallowed", () => {
  assert.equal(onlyAlarmBookkeeping(event(), "not an icalendar document"), false);
  assert.equal(onlyAlarmBookkeeping("", event()), false);
});

test("a create or a delete is never this", () => {
  assert.equal(onlyAlarmBookkeeping(null, event()), false);
  assert.equal(onlyAlarmBookkeeping(event(), null), false);
});
