/**
 * The invitation, the update and the cancellation an organiser owes.
 *
 * Every assertion here is about a message a real person receives, so the
 * failures are of two kinds and both matter: something the recipient needs
 * is missing, or something from inside this profile is in it.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { buildMeetingRequestMime } from "../../src/modules/eas/meeting-request-mail.mjs";

const NOW = new Date("2026-08-14T10:00:00Z");
const ME = "me@x.de";

function blob({ extra = [], attendees = ["bob@x.de", "ann@x.de"] } = {}) {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VTIMEZONE",
    "TZID:Europe/Berlin",
    "BEGIN:STANDARD",
    "DTSTART:19701025T030000",
    "TZOFFSETFROM:+0200",
    "TZOFFSETTO:+0100",
    "TZNAME:CET",
    "END:STANDARD",
    "END:VTIMEZONE",
    "BEGIN:VEVENT",
    "UID:meeting-1",
    "DTSTART;TZID=Europe/Berlin:20260901T110000",
    "DTEND;TZID=Europe/Berlin:20260901T120000",
    "SUMMARY:Standup",
    "LOCATION:Room 1",
    `ORGANIZER;CN=Me:mailto:${ME}`,
    ...attendees.map((a) => `ATTENDEE;PARTSTAT=NEEDS-ACTION:mailto:${a}`),
    ...extra,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

const build = (over = {}) =>
  buildMeetingRequestMime({
    blob: blob(),
    method: "REQUEST",
    recipients: ["bob@x.de", "ann@x.de"],
    userEmail: ME,
    userName: "Me",
    now: NOW,
    ...over,
  });

const lines = (mime) => mime.split("\r\n");
const has = (mime, re) => lines(mime).some((l) => re.test(l));

test("an invitation carries the item's own UID", () => {
  // [MS-ASCMD]: this is what lets the organiser's mailbox reconcile the
  // replies that come back against the meeting they belong to. If it
  // diverges, every answer is orphaned.
  assert.ok(has(build(), /^UID:meeting-1$/));
});

test("it is addressed to the attendees and announces its method", () => {
  const mime = build();
  assert.ok(has(mime, /^To: bob@x\.de, ann@x\.de$/));
  assert.ok(has(mime, /^Subject: Standup$/));
  assert.ok(has(mime, /^METHOD:REQUEST$/));
  assert.ok(has(mime, /^Content-Type: text\/calendar;.*method=REQUEST/));
});

test("nothing from inside this profile travels", () => {
  // Our stamps say what the server calls the item; Thunderbird's say which
  // alarms this user has dismissed. Neither is the recipient's business,
  // and both used to survive because ical.js hands back a live array and
  // removing while iterating skips the next element.
  const mime = build({
    blob: blob({
      extra: [
        "X-EAS-SERVERID:abc",
        "X-MOZ-LASTACK:20260101T000000Z",
        "X-EAS-RESPONSETYPE:3",
        "X-MOZ-SEND-INVITATIONS:TRUE",
        "BEGIN:VALARM",
        "ACTION:DISPLAY",
        "TRIGGER:-PT15M",
        "END:VALARM",
        "BEGIN:VALARM",
        "ACTION:DISPLAY",
        "TRIGGER:-PT5M",
        "END:VALARM",
      ],
    }),
  });
  assert.equal(lines(mime).filter((l) => /^X-(EAS|MOZ)-/i.test(l)).length, 0);
  assert.ok(!has(mime, /VALARM/));
});

test("the timezone definition is carried, not assumed", () => {
  // A receiver cannot resolve a TZID it was never given.
  const mime = build();
  assert.ok(has(mime, /^BEGIN:VTIMEZONE$/));
  assert.ok(has(mime, /^TZID:Europe\/Berlin$/));
  assert.ok(has(mime, /^DTSTART;TZID=Europe\/Berlin:20260901T110000$/));
});

test("DTSTAMP is the caller's clock, and SEQUENCE is zero", () => {
  const mime = build();
  assert.ok(has(mime, /^DTSTAMP:20260814T100000Z$/));
  assert.ok(has(mime, /^SEQUENCE:0$/));
});

test("a cancellation says so, and only to the people it is for", () => {
  // Somebody dropped from a meeting that is going ahead gets their own
  // CANCEL; naming the others in it would tell them they were uninvited.
  const mime = build({ method: "CANCEL", recipients: ["ann@x.de"] });
  assert.ok(has(mime, /^To: ann@x\.de$/));
  assert.ok(has(mime, /^Subject: Canceled: Standup$/));
  assert.ok(has(mime, /^METHOD:CANCEL$/));
  assert.ok(has(mime, /^STATUS:CANCELLED$/));
  assert.ok(has(mime, /^ATTENDEE.*ann@x\.de$/));
  assert.ok(!has(mime, /bob@x\.de/), "the attendee who is still invited");
});

test("a series carries its rule", () => {
  const mime = build({ blob: blob({ extra: ["RRULE:FREQ=WEEKLY;COUNT=3"] }) });
  assert.ok(has(mime, /^RRULE:FREQ=WEEKLY;COUNT=3$/));
});

test("only the master travels, never an override", () => {
  // An occurrence edit is not announced in this pass, and shipping the
  // override would describe a change the recipient has no context for.
  const series = blob({ extra: ["RRULE:FREQ=WEEKLY;COUNT=3"] }).replace(
    "END:VCALENDAR",
    [
      "BEGIN:VEVENT",
      "UID:meeting-1",
      "RECURRENCE-ID;TZID=Europe/Berlin:20260908T110000",
      "DTSTART;TZID=Europe/Berlin:20260908T150000",
      "SUMMARY:Standup (moved)",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\r\n"),
  );
  const mime = build({ blob: series });
  assert.ok(!has(mime, /RECURRENCE-ID/));
  assert.ok(!has(mime, /Standup \(moved\)/));
  assert.equal(lines(mime).filter((l) => l === "BEGIN:VEVENT").length, 1);
});

test("nothing is built when there is nobody to tell, or nothing to tell", () => {
  assert.equal(build({ recipients: [] }), null);
  assert.equal(build({ recipients: undefined }), null);
  assert.equal(build({ blob: "" }), null);
  assert.equal(build({ blob: "not a calendar" }), null);
  assert.equal(build({ userEmail: "" }), null);
  assert.equal(build({ method: "REPLY" }), null, "not ours to send");
  assert.equal(
    build({ blob: blob().replace("UID:meeting-1\r\n", "") }),
    null,
    "without a UID no reply could ever be reconciled",
  );
});

test("a subject that spans lines cannot forge a header", () => {
  const mime = build({
    blob: blob({ extra: [] }).replace(
      "SUMMARY:Standup",
      "SUMMARY:Standup\\nBcc: someone@elsewhere.invalid",
    ),
  });
  assert.ok(!has(mime, /^Bcc:/i));
});

test("a cancellation reads as one in the calendar, not only the subject", () => {
  // A client renders an iTIP message from the calendar item rather than
  // from the MIME headers, so a prefix that lived only in the subject
  // arrived looking exactly like the invitation it was cancelling -
  // measured on Microsoft, which showed the bare title.
  const mime = build({ method: "CANCEL", recipients: ["ann@x.de"] });
  assert.ok(has(mime, /^SUMMARY:Canceled: Standup$/));
  assert.ok(has(mime, /^Subject: Canceled: Standup$/));
  // and an invitation is untouched
  assert.ok(has(build(), /^SUMMARY:Standup$/));
});
