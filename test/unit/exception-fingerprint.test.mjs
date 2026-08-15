/**
 * The fingerprint that decides whether an occurrence still matches what the
 * server was last told.
 *
 * Both failure directions cost something, and they are not symmetric. Too
 * sensitive and we re-send exceptions the server already holds - noise, and
 * the cause of 3.3's intermittent failure. Too blunt and a real edit to an
 * occurrence is never pushed at all, which is data loss. So these tests
 * come in pairs: what must NOT change the digest, and what must.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { exceptionFingerprint } from "../../src/modules/eas/calendar-codec.mjs";

// A real blob carries the zone it references; without one ical.js cannot
// resolve the TZID and treats the time as floating, so the conversion this
// tests could never happen.
const NY = [
  "BEGIN:VTIMEZONE", "TZID:America/New_York",
  "BEGIN:DAYLIGHT", "DTSTART:19700308T020000",
  "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU",
  "TZOFFSETFROM:-0500", "TZOFFSETTO:-0400", "TZNAME:EDT", "END:DAYLIGHT",
  "BEGIN:STANDARD", "DTSTART:19701101T020000",
  "RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU",
  "TZOFFSETFROM:-0400", "TZOFFSETTO:-0500", "TZNAME:EST", "END:STANDARD",
  "END:VTIMEZONE",
];

const series = (overrideLines) =>
  [
    "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//t//EN", ...NY,
    "BEGIN:VEVENT",
    "UID:s1", "DTSTAMP:20260801T120000Z",
    "SUMMARY:weekly", "DTSTART:20260902T110000Z", "DTEND:20260902T120000Z",
    "RRULE:FREQ=WEEKLY;COUNT=4",
    "EXDATE:20260916T110000Z",
    "END:VEVENT",
    "BEGIN:VEVENT", "UID:s1", ...overrideLines, "END:VEVENT",
    "END:VCALENDAR", "",
  ].join("\r\n");

const digest = (ical) => exceptionFingerprint(ical).overrides[0].digest;

test("the same instant in another zone is not a change", () => {
  // What a round trip through the server does: the occurrence comes back
  // rendered in the default zone. Same moment, different text - and taken
  // over the raw text it re-asserted the exception on every later edit.
  const ny = digest(series([
    "RECURRENCE-ID:20260909T110000Z",
    "DTSTART;TZID=America/New_York:20260909T090000",
    "DTEND;TZID=America/New_York:20260909T100000",
    "SUMMARY:moved",
  ]));
  const utc = digest(series([
    "RECURRENCE-ID:20260909T110000Z",
    "DTSTART:20260909T130000Z",
    "DTEND:20260909T140000Z",
    "SUMMARY:moved",
  ]));
  assert.equal(ny, utc, "13:00 New York and 13:00Z are the same instant");
});

test("bookkeeping that moves on every write is not a change", () => {
  const a = digest(series([
    "RECURRENCE-ID:20260909T110000Z", "DTSTART:20260909T130000Z",
    "DTEND:20260909T140000Z", "SUMMARY:moved",
    "DTSTAMP:20260801T120000Z", "SEQUENCE:0", "X-EAS-SERVERID:abc",
  ]));
  const b = digest(series([
    "RECURRENCE-ID:20260909T110000Z", "DTSTART:20260909T130000Z",
    "DTEND:20260909T140000Z", "SUMMARY:moved",
    "DTSTAMP:20260815T090000Z", "SEQUENCE:7", "X-EAS-SERVERID:zzz",
  ]));
  assert.equal(a, b);
});

test("but moving the occurrence IS a change", () => {
  const at13 = digest(series([
    "RECURRENCE-ID:20260909T110000Z", "DTSTART:20260909T130000Z",
    "DTEND:20260909T140000Z", "SUMMARY:moved",
  ]));
  const at14 = digest(series([
    "RECURRENCE-ID:20260909T110000Z", "DTSTART:20260909T140000Z",
    "DTEND:20260909T150000Z", "SUMMARY:moved",
  ]));
  assert.notEqual(at13, at14, "an hour later must still be pushed");
});

test("and so is anything the user can see", () => {
  const base = [
    "RECURRENCE-ID:20260909T110000Z", "DTSTART:20260909T130000Z",
    "DTEND:20260909T140000Z",
  ];
  const d = (extra) => digest(series([...base, ...extra]));
  const plain = d(["SUMMARY:moved"]);
  assert.notEqual(plain, d(["SUMMARY:moved again"]), "summary");
  assert.notEqual(plain, d(["SUMMARY:moved", "LOCATION:Room 2"]), "location");
  assert.notEqual(plain, d(["SUMMARY:moved", "DESCRIPTION:why"]), "description");
  assert.notEqual(plain, d(["SUMMARY:moved", "STATUS:CANCELLED"]), "status");
});

test("an all-day override keeps its date rather than being shifted to UTC", () => {
  // Forcing a DATE to UTC would move the day, and a DATE carries no zone to
  // disagree about in the first place.
  const fp = exceptionFingerprint([
    "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//t//EN",
    "BEGIN:VEVENT", "UID:a1", "DTSTAMP:20260801T120000Z",
    "SUMMARY:daily", "DTSTART;VALUE=DATE:20261012", "DTEND;VALUE=DATE:20261013",
    "RRULE:FREQ=DAILY;COUNT=3", "END:VEVENT",
    "BEGIN:VEVENT", "UID:a1",
    "RECURRENCE-ID;VALUE=DATE:20261013",
    "DTSTART;VALUE=DATE:20261013", "DTEND;VALUE=DATE:20261014",
    "SUMMARY:that day", "END:VEVENT",
    "END:VCALENDAR", "",
  ].join("\r\n"));
  assert.equal(fp.overrides.length, 1);
  // A DATE row keeps its day. `instanceKey` renders it as midnight-as-UTC,
  // which is the encoding the wire uses - what matters is that the day did
  // not move, which forcing a real zone conversion would have done.
  assert.equal(fp.overrides[0].rid, "20261013T000000Z");
});
