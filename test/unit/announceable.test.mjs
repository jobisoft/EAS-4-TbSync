/**
 * What counts as a change worth telling the attendees about.
 *
 * `announceableOf` is the whole definition, and both sides of the decision
 * use it: the item hook records one of these as the meeting stood before
 * the user's edit, and the phase that sends the message builds another once
 * the sync has settled. Equal means nobody is mailed.
 *
 * So every case here is a message somebody either receives or does not.
 * A false positive is a mail about nothing, sent to everyone invited; a
 * false negative leaves attendees holding the old time. The normalising is
 * what keeps the first from happening on every round trip, because the
 * server hands back its own spelling of what we sent it.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  announceableOf,
  droppedAttendees,
  sameAnnounceable,
} from "../../src/modules/eas/calendar-codec.mjs";

const BASE = [
  "DTSTART;TZID=Europe/Berlin:20260901T110000",
  "DTEND;TZID=Europe/Berlin:20260901T120000",
  "SUMMARY:Standup",
  "LOCATION:Room 1",
  "ORGANIZER;CN=Me:mailto:me@x.de",
  "ATTENDEE;PARTSTAT=NEEDS-ACTION;CN=Bob:mailto:BOB@x.de",
  "ATTENDEE;PARTSTAT=ACCEPTED:mailto:ann@x.de",
];

/** One VEVENT built from `lines`, wrapped as a calendar. */
function ics(lines = BASE, extra = []) {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    "UID:u1",
    ...lines,
    ...extra,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

/** `BASE` with one line swapped, so each test names only what it changed. */
function swap(find, replaceWith) {
  return BASE.map((l) => (l.startsWith(find) ? replaceWith : l));
}

const same = (a, b) => assert.deepEqual(announceableOf(a), announceableOf(b));
const differs = (a, b) =>
  assert.notDeepEqual(announceableOf(a), announceableOf(b));

test("the bag holds the announceable fields and nothing else", () => {
  assert.deepEqual(announceableOf(ics()), {
    start: { at: "2026-09-01T09:00:00.000Z" },
    end: { at: "2026-09-01T10:00:00.000Z" },
    allDay: false,
    location: "Room 1",
    summary: "Standup",
    status: "",
    rrule: null,
    attendees: ["ann@x.de", "bob@x.de"],
  });
});

test("one instant written two ways is not a reschedule", () => {
  // The round trip is the reason: we send a zoned DTSTART and the server
  // may hand back the same moment as UTC. Comparing the text would mail
  // every attendee after every sync.
  same(
    ics(),
    ics(
      swap("DTSTART", "DTSTART:20260901T090000Z").map((l) =>
        l.startsWith("DTEND") ? "DTEND:20260901T100000Z" : l,
      ),
    ),
  );
});

test("PARTSTAT, RSVP and CN churn is not a change to who is invited", () => {
  // Replies arriving rewrite PARTSTAT on the organiser's own copy. Nobody
  // needs telling that the meeting they accepted still exists.
  same(
    ics(),
    ics(swap("ATTENDEE;PARTSTAT=NEEDS-ACTION", "ATTENDEE;PARTSTAT=DECLINED;CN=Robert;RSVP=TRUE:mailto:bob@x.de")),
  );
});

test("attendee order and case are not a change either", () => {
  same(ics(), ics([...BASE.slice(0, 5), BASE[6], BASE[5]]));
});

test("an alarm, a category or a colour says nothing", () => {
  same(ics(), ics(BASE, ["BEGIN:VALARM", "ACTION:DISPLAY", "TRIGGER:-PT15M", "END:VALARM"]));
  same(ics(), ics(BASE, ["CATEGORIES:Work"]));
  same(ics(), ics(BASE, ["X-APPLE-CALENDAR-COLOR:#FF0000", "TRANSP:TRANSPARENT"]));
  same(ics(), ics(BASE, ["DESCRIPTION:some notes the user typed"]));
});

test("each of the announceable fields, on its own, is a change", () => {
  differs(ics(), ics(swap("DTSTART", "DTSTART;TZID=Europe/Berlin:20260901T120000")));
  differs(ics(), ics(swap("DTEND", "DTEND;TZID=Europe/Berlin:20260901T130000")));
  differs(ics(), ics(swap("SUMMARY", "SUMMARY:Standup (moved)")));
  differs(ics(), ics(swap("LOCATION", "LOCATION:Room 2")));
  differs(ics(), ics(BASE, ["STATUS:CANCELLED"]));
  differs(ics(), ics(BASE.slice(0, 6)));
});

test("cancelling and un-cancelling are both visible", () => {
  const live = ics();
  const cancelled = ics(BASE, ["STATUS:CANCELLED"]);
  assert.equal(announceableOf(cancelled).status, "CANCELLED");
  assert.equal(announceableOf(live).status, "");
  differs(live, cancelled);
});

test("a dropped attendee is recoverable from the two bags", () => {
  // The CANCEL to somebody removed goes to an address only the earlier bag
  // still knows, which is why these hold values rather than a digest.
  const before = announceableOf(ics());
  const after = announceableOf(ics(BASE.slice(0, 6)));
  assert.deepEqual(after.attendees, ["bob@x.de"], "Ann was the one removed");
  assert.deepEqual(
    before.attendees.filter((a) => !after.attendees.includes(a)),
    ["ann@x.de"],
  );
});

test("an all-day boundary is a date, never a midnight in some zone", () => {
  // Compared as an instant it would move by a day for anyone west of UTC,
  // and mail everybody every time the item crossed the wire.
  const allDay = announceableOf(
    ics(
      swap("DTSTART", "DTSTART;VALUE=DATE:20260901").map((l) =>
        l.startsWith("DTEND") ? "DTEND;VALUE=DATE:20260902" : l,
      ),
    ),
  );
  assert.deepEqual(allDay.start, { date: "2026-09-01" });
  assert.equal(allDay.allDay, true);
  assert.equal(announceableOf(ics()).allDay, false);
});

test("an occurrence edit does not read as a change to the series", () => {
  // The first pass announces the series only. Walking every component here
  // would turn one occurrence's edit into a message to everyone about the
  // whole series.
  const series = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    "UID:u1",
    ...BASE,
    "RRULE:FREQ=WEEKLY;COUNT=4",
    "END:VEVENT",
    "BEGIN:VEVENT",
    "UID:u1",
    "RECURRENCE-ID;TZID=Europe/Berlin:20260908T110000",
    "DTSTART;TZID=Europe/Berlin:20260908T150000",
    "DTEND;TZID=Europe/Berlin:20260908T160000",
    "SUMMARY:Standup (this one moved)",
    "LOCATION:Room 9",
    "ORGANIZER;CN=Me:mailto:me@x.de",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
  const master = announceableOf(series);
  assert.equal(master.summary, "Standup");
  assert.equal(master.location, "Room 1");
  assert.deepEqual(master.start, { at: "2026-09-01T09:00:00.000Z" });
});

test("something that is not a calendar answers null, and mails nobody", () => {
  assert.equal(announceableOf(""), null);
  assert.equal(announceableOf("not a calendar"), null);
  assert.equal(announceableOf(undefined), null);
  assert.equal(
    announceableOf("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR"),
    null,
  );
});

// ── The shapes a round trip changes ───────────────────────────────────────
//
// Every one of these is the same failure: the user saves a meeting one way,
// the server hands it back spelled another, and a text comparison calls
// that a change. The difference from the cases above is that these were
// found by review rather than by thinking about them - each would have
// mailed every attendee of every meeting, once per sync, forever.

test("DURATION and DTEND are the same fact", () => {
  same(
    ics(),
    ics(BASE.filter((l) => !l.startsWith("DTEND")), ["DURATION:PT1H"]),
  );
});

test("changing only the length of a meeting IS announceable", () => {
  differs(
    ics(BASE.filter((l) => !l.startsWith("DTEND")), ["DURATION:PT1H"]),
    ics(BASE.filter((l) => !l.startsWith("DTEND")), ["DURATION:PT2H"]),
  );
});

test("an all-day event's end is a day, however it is written", () => {
  // RFC 5545 §3.6.1 gives a missing DTEND one day; Outlook writes DTEND
  // equal to DTSTART and means the same thing.
  const day = (dtend) =>
    ics([
      "DTSTART;VALUE=DATE:20260901",
      ...(dtend ? [dtend] : []),
      ...BASE.slice(2),
    ]);
  same(day("DTEND;VALUE=DATE:20260902"), day(null));
  same(day("DTEND;VALUE=DATE:20260902"), day("DTEND;VALUE=DATE:20260901"));
});

test("a meeting the server has confirmed has not changed", () => {
  // The decode stamps STATUS:CONFIRMED on anything the server calls a
  // meeting. A locally-authored one carries no STATUS at all, so without
  // this every meeting mails everybody after its first round trip.
  same(ics(), ics(BASE, ["STATUS:CONFIRMED"]));
  differs(ics(), ics(BASE, ["STATUS:CANCELLED"]));
  differs(ics(), ics(BASE, ["STATUS:TENTATIVE"]));
});

test("the organiser is not one of the people we invite", () => {
  // Exchange returns the organiser inside its own Attendees list, so the
  // copy that comes back can carry an ATTENDEE the saved one never had.
  same(ics(), ics(BASE, ["ATTENDEE;CN=Me:mailto:ME@x.de"]));
});

test("changing the recurrence of a series is announceable", () => {
  const weekly = ics(BASE, ["RRULE:FREQ=WEEKLY;COUNT=4"]);
  differs(ics(), weekly);
  differs(weekly, ics(BASE, ["RRULE:FREQ=DAILY;COUNT=9"]));
  same(weekly, ics(BASE, ["RRULE:FREQ=WEEKLY;COUNT=4"]));
});

test("but deleting one occurrence is not", () => {
  // An EXDATE on the master is one occurrence removed, and the rule is
  // that a deletion is never announced - not even this one.
  const weekly = ics(BASE, ["RRULE:FREQ=WEEKLY;COUNT=4"]);
  same(
    weekly,
    ics(BASE, [
      "RRULE:FREQ=WEEKLY;COUNT=4",
      "EXDATE;TZID=Europe/Berlin:20260908T110000",
    ]),
  );
});

// ── The comparison that decides whether anybody is mailed ────────────────

test("a moved meeting is not the same meeting", () => {
  // Found by checking rather than by a test failing: the first version
  // compared with a JSON.stringify replacer array, which drops every
  // NESTED value - so start and end always compared equal and a moved
  // meeting told nobody. The false negative is the dangerous direction.
  const a = announceableOf(ics());
  const b = announceableOf(
    ics(swap("DTSTART", "DTSTART;TZID=Europe/Berlin:20260901T150000")),
  );
  assert.equal(sameAnnounceable(a, b), false);
  assert.equal(sameAnnounceable(a, announceableOf(ics())), true);
});

test("every announceable field is seen through the comparison", () => {
  const base = announceableOf(ics());
  const changed = [
    swap("DTEND", "DTEND;TZID=Europe/Berlin:20260901T133000"),
    swap("SUMMARY", "SUMMARY:Elsewhere"),
    swap("LOCATION", "LOCATION:Room 2"),
    BASE.slice(0, 6),
  ];
  for (const lines of changed) {
    assert.equal(sameAnnounceable(base, announceableOf(ics(lines))), false);
  }
  assert.equal(
    sameAnnounceable(base, announceableOf(ics(BASE, ["STATUS:CANCELLED"]))),
    false,
  );
  assert.equal(
    sameAnnounceable(base, announceableOf(ics(BASE, ["RRULE:FREQ=DAILY"]))),
    false,
  );
});

test("a reminder-only edit tells nobody", () => {
  assert.equal(
    sameAnnounceable(
      announceableOf(ics()),
      announceableOf(ics(BASE, ["BEGIN:VALARM", "ACTION:DISPLAY", "TRIGGER:-PT5M", "END:VALARM"])),
    ),
    true,
  );
});

test("a bag from another build is not compared, it is ignored", () => {
  // If a later version defines the fields differently, every stored bag
  // would otherwise read as changed and mail every attendee of every
  // pending meeting on the first sync after the update.
  const now = announceableOf(ics());
  const older = { ...now };
  delete older.rrule;
  assert.equal(sameAnnounceable(older, now), true, "unreadable means silent");
  assert.equal(sameAnnounceable(null, now), false);
  assert.equal(sameAnnounceable(now, null), false);
});

test("who was dropped comes out of the earlier bag", () => {
  const before = announceableOf(ics());
  const after = announceableOf(ics(BASE.slice(0, 6)));
  assert.deepEqual(droppedAttendees(before, after), ["ann@x.de"]);
  assert.deepEqual(droppedAttendees(before, before), []);
  assert.deepEqual(droppedAttendees(null, after), []);
  assert.deepEqual(
    droppedAttendees(before, announceableOf(ics(BASE.slice(0, 5)))),
    ["ann@x.de", "bob@x.de"],
    "removing everyone still owes everyone a cancellation",
  );
});
