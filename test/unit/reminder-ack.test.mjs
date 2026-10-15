/**
 * A dismissed reminder stays dismissed when the server changes something else.
 *
 * `<Reminder>` in the ApplicationData is the server's authority over the
 * alarm, so the merge clears the VALARMs before re-adding it - otherwise a
 * partial Change stacks a second alarm. The user's answer to that alarm was
 * cleared with them, and Exchange re-sends the element whenever anything
 * else about the meeting changes, so an organiser moving a room resurrected
 * every attendee's dismissed reminder.
 *
 * `X-MOZ-LASTACK` and `X-MOZ-SNOOZE-TIME` are one answer in two properties:
 * Thunderbird writes the first for a snooze as well as a dismiss, and the
 * second only says when to come back. They are carried together or not at
 * all, because an ack without its snooze fires at once instead of later.
 *
 * The other direction has its own file - `alarm-bookkeeping.test.mjs`,
 * which keeps such a write from being pushed. Both come from TbSync #816.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import {
  applicationDataToIcal,
  preserveAlarmAck,
} from "../../src/modules/eas/calendar-codec.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";

before(() => ensureLoaded());

const ACK = "20260916T084500Z";
const SNOOZE = "20260916T090000Z";

/** An event with one reminder, plus whatever answer the user gave it. */
const event = ({ trigger = "-PT15M", marks = [], rid = null } = {}) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VEVENT",
    "UID:ack-uid",
    "DTSTAMP:20260901T090000Z",
    "DTSTART:20260916T090000Z",
    "DTEND:20260916T093000Z",
    "SUMMARY:planning",
    ...(rid ? [`RECURRENCE-ID:${rid}`] : []),
    ...marks,
    "BEGIN:VALARM",
    "ACTION:DISPLAY",
    `TRIGGER:${trigger}`,
    "DESCRIPTION:Reminder",
    "END:VALARM",
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");

/** The same event with no alarm at all. */
const noAlarm = (marks = []) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VEVENT",
    "UID:ack-uid",
    "DTSTAMP:20260901T090000Z",
    "DTSTART:20260916T090000Z",
    "DTEND:20260916T093000Z",
    "SUMMARY:planning",
    ...marks,
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");

const has = (ical, name) =>
  ical.split(/\r?\n/).some((l) => l.toUpperCase().startsWith(`${name}:`));

test("an unchanged reminder keeps the dismissal", () => {
  const out = preserveAlarmAck({
    builtIcal: event(),
    priorIcal: event({ marks: [`X-MOZ-LASTACK:${ACK}`] }),
  });
  assert.ok(has(out, "X-MOZ-LASTACK"), "the dismissal was dropped");
  assert.match(out, new RegExp(ACK));
});

test("a snooze keeps both halves of its answer, together", () => {
  const out = preserveAlarmAck({
    builtIcal: event(),
    priorIcal: event({
      marks: [`X-MOZ-LASTACK:${ACK}`, `X-MOZ-SNOOZE-TIME:${SNOOZE}`],
    }),
  });
  assert.ok(has(out, "X-MOZ-LASTACK"), "the suppression was dropped");
  assert.ok(has(out, "X-MOZ-SNOOZE-TIME"), "the re-arm was dropped");
});

test("a per-occurrence snooze is carried by its own name", () => {
  const out = preserveAlarmAck({
    builtIcal: event(),
    priorIcal: event({
      marks: [
        `X-MOZ-LASTACK:${ACK}`,
        `X-MOZ-SNOOZE-TIME-1793558400000000:${SNOOZE}`,
      ],
    }),
  });
  assert.ok(has(out, "X-MOZ-SNOOZE-TIME-1793558400000000"));
});

test("a reminder the organiser moved rings again", () => {
  const out = preserveAlarmAck({
    builtIcal: event({ trigger: "-PT60M" }),
    priorIcal: event({ trigger: "-PT15M", marks: [`X-MOZ-LASTACK:${ACK}`] }),
  });
  assert.equal(has(out, "X-MOZ-LASTACK"), false, "a new alarm was suppressed");
});

test("a reminder the server removed carries nothing over", () => {
  const out = preserveAlarmAck({
    builtIcal: noAlarm(),
    priorIcal: event({ marks: [`X-MOZ-LASTACK:${ACK}`] }),
  });
  assert.equal(has(out, "X-MOZ-LASTACK"), false);
});

test("a second alarm is a different alarm", () => {
  const twoAlarms = event().replace(
    "END:VEVENT",
    "BEGIN:VALARM\r\nACTION:DISPLAY\r\nTRIGGER:-PT5M\r\nEND:VALARM\r\nEND:VEVENT",
  );
  const out = preserveAlarmAck({
    builtIcal: twoAlarms,
    priorIcal: event({ marks: [`X-MOZ-LASTACK:${ACK}`] }),
  });
  assert.equal(has(out, "X-MOZ-LASTACK"), false);
});

test("an item nobody answered is left exactly as built", () => {
  const built = event();
  assert.equal(preserveAlarmAck({ builtIcal: built, priorIcal: event() }), built);
});

/** A series and one override, each with its own reminder. `marks` go on
 *  whichever component the caller names. */
const RID = "20260923T090000Z";
const vevent = (extra) =>
  [
    "BEGIN:VEVENT",
    "UID:ack-uid",
    "DTSTAMP:20260901T090000Z",
    "DTSTART:20260916T090000Z",
    "DTEND:20260916T093000Z",
    "SUMMARY:planning",
    ...extra,
    "BEGIN:VALARM",
    "ACTION:DISPLAY",
    "TRIGGER:-PT15M",
    "DESCRIPTION:Reminder",
    "END:VALARM",
    "END:VEVENT",
  ].join("\r\n");

const seriesWithOverride = ({ onMaster = [], onOverride = [] } = {}) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    vevent(onMaster),
    vevent([`RECURRENCE-ID:${RID}`, ...onOverride]),
    "END:VCALENDAR",
    "",
  ].join("\r\n");

test("each occurrence is judged on its own", () => {
  // Answered on the override and nowhere else, which is how Thunderbird
  // records a reminder dealt with on a single occurrence.
  const out = preserveAlarmAck({
    builtIcal: seriesWithOverride(),
    priorIcal: seriesWithOverride({ onOverride: [`X-MOZ-LASTACK:${ACK}`] }),
  });
  const bodies = out.split("BEGIN:VEVENT").slice(1);
  assert.equal(bodies.length, 2, "the fixture lost a component");
  const marked = bodies.filter((b) =>
    b.toUpperCase().includes("X-MOZ-LASTACK"),
  );
  assert.equal(marked.length, 1, "the answer spread beyond its occurrence");
  assert.ok(
    marked[0].includes(`RECURRENCE-ID:${RID}`),
    "the answer landed on the series instead of the occurrence",
  );
});

test("an item that will not parse is returned untouched", () => {
  const built = event();
  assert.equal(preserveAlarmAck({ builtIcal: built, priorIcal: "rubbish" }), built);
  assert.equal(preserveAlarmAck({ builtIcal: "rubbish", priorIcal: event({ marks: [`X-MOZ-LASTACK:${ACK}`] }) }), "rubbish");
});

/* ── through the real merge, the way a pull reaches it ─────────────────── */

const CHANGE_SAME_REMINDER = `<ApplicationData>
  <Subject xmlns="Calendar">planning, moved room</Subject>
  <StartTime xmlns="Calendar">20260916T090000Z</StartTime>
  <EndTime xmlns="Calendar">20260916T093000Z</EndTime>
  <Reminder xmlns="Calendar">15</Reminder>
</ApplicationData>`;

test("a server change to something else does not ring the reminder again", async () => {
  const prior = event({ marks: [`X-MOZ-LASTACK:${ACK}`] });
  const merged = await applicationDataToIcal({
    adNode: parseAdNode(CHANGE_SAME_REMINDER),
    existingIcal: prior,
    serverID: "srv-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: "ack-uid",
  });
  assert.equal(
    has(merged, "X-MOZ-LASTACK"),
    false,
    "the merge no longer clears it - this test would then prove nothing",
  );
  const out = preserveAlarmAck({ builtIcal: merged, priorIcal: prior });
  assert.ok(has(out, "X-MOZ-LASTACK"), "the dismissal did not survive the pull");
  // iCalendar escapes the comma, hence the backslash.
  assert.match(out, /SUMMARY:planning\\, moved room/, "the server's edit was lost");
});
