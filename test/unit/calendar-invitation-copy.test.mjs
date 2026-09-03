/**
 * Unit tests for an invitation the mailbox holds no copy of.
 *
 * The case: an invitation addressed to somebody else is delivered here - a
 * redirected mail - and the user accepts it. Office 365 created nothing,
 * correctly: there was nothing for this mailbox to create. So no ServerId
 * ever arrives, `invitationPhase` can address no MeetingResponse, and the
 * queued answer is held on every sync forever (TbSync #811).
 *
 * It cannot be sent as an item either, not while it still names a guest
 * list: on 16.x the client may state neither the organizer nor an
 * attendee's status, so the server would make the user the organizer and
 * invite all of them in the user's name.
 *
 * After long enough the answer is given up on and the appointment is kept
 * as the user's own. The two things that must be true of the result are
 * asserted here: it no longer reads as somebody else's meeting, so the push
 * stops diverting it, and it carries no answer anybody could try to send.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { installWebextEnv } from "./support/webext-env.mjs";

installWebextEnv();

const { isReceivedMeeting, plainCopyOfInvitation, selfUserResponses } =
  await import("../../src/modules/eas/calendar-codec.mjs");

const MAILBOX = "user@example.invalid";
const ORG = "organizer@example.invalid";
const OTHER = "redirected@example.invalid";

/** An invitation as Thunderbird files it from an emailed iTIP: no X-EAS-*
 *  of any kind, the organizer somebody else, and the marker naming the
 *  identity the user answered as. */
function filed({
  attendees = [OTHER],
  marker = OTHER,
  organizer = ORG,
  summary = "Quarterly review",
  description = "Bring the numbers",
} = {}) {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    "UID:uid-redirected-1",
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20260901T100000Z",
    "DTEND:20260901T110000Z",
    `SUMMARY:${summary}`,
    ...(description ? [`DESCRIPTION:${description}`] : []),
    ...(organizer ? [`ORGANIZER;CN=Big Boss:mailto:${organizer}`] : []),
    ...attendees.map(
      (a) => `ATTENDEE;PARTSTAT=ACCEPTED;ROLE=REQ-PARTICIPANT:mailto:${a}`,
    ),
    ...(marker ? [`X-MOZ-INVITED-ATTENDEE:mailto:${marker}`] : []),
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

/** The same, recurring, with one modified occurrence carrying its own
 *  attendee block - which `selfUserResponses` reads on its own. */
function filedSeries() {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    "UID:uid-series-1",
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20260901T100000Z",
    "RRULE:FREQ=WEEKLY;COUNT=4",
    "SUMMARY:Standup",
    `ORGANIZER:mailto:${ORG}`,
    `ATTENDEE;PARTSTAT=ACCEPTED:mailto:${OTHER}`,
    `X-MOZ-INVITED-ATTENDEE:mailto:${OTHER}`,
    "END:VEVENT",
    "BEGIN:VEVENT",
    "UID:uid-series-1",
    "RECURRENCE-ID:20260908T100000Z",
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20260908T113000Z",
    "SUMMARY:Standup",
    `ORGANIZER:mailto:${ORG}`,
    `ATTENDEE;PARTSTAT=TENTATIVE:mailto:${OTHER}`,
    `X-MOZ-INVITED-ATTENDEE:mailto:${OTHER}`,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

const copied = (ical) =>
  plainCopyOfInvitation(ical, { summaryPrefix: "[Copy]", organizer: MAILBOX });


/* ── Making the copy ──────────────────────────────────────────────────── */
test("the copy is nobody's meeting but the user's", () => {
  // The pair this whole change exists for: the push no longer diverts it,
  // and nothing is left for the answer phase to try to send.
  const out = copied(filed());
  assert.equal(isReceivedMeeting(out, MAILBOX), false);
  assert.deepEqual(selfUserResponses(out, MAILBOX), []);
});

test("the guest list and the marker are gone, the organizer is ours", () => {
  const out = copied(filed({ attendees: [OTHER, "third@example.invalid"] }));
  assert.equal(/^ATTENDEE/m.test(out), false);
  assert.equal(/X-MOZ-INVITED-ATTENDEE/.test(out), false);
  assert.equal(out.includes(`ORGANIZER:mailto:${MAILBOX}`), true);
  assert.equal(out.includes(ORG), false);
  assert.equal(out.includes("third@example.invalid"), false);
});

test("every occurrence is stripped, not just the master", () => {
  // An override carries its own attendee block, and the answer phase reads
  // each one on its own - so a master-only strip leaves the item answerable.
  const out = copied(filedSeries());
  assert.equal(/^ATTENDEE/m.test(out), false);
  assert.equal(/X-MOZ-INVITED-ATTENDEE/.test(out), false);
  assert.equal(out.includes(ORG), false);
  assert.deepEqual(selfUserResponses(out, MAILBOX), []);
  // Still one series: the override stays bound to its master.
  assert.equal(out.match(/UID:uid-series-1/g).length, 2);
  assert.equal(out.includes("RECURRENCE-ID:20260908T100000Z"), true);
  assert.equal(out.includes("RRULE:FREQ=WEEKLY;COUNT=4"), true);
});

test("the title says [Copy], once, however often this runs", () => {
  const once = copied(filed());
  assert.equal(/SUMMARY:\[Copy\] Quarterly review/.test(once), true);
  assert.equal(copied(once), once);
});

test("the note is left exactly as it came", () => {
  // Deliberate: nothing is rewritten into the description, so a note the
  // user wrote - or an ALTREP the server will read as the HTML body -
  // survives untouched.
  const out = copied(filed());
  assert.equal(out.includes("DESCRIPTION:Bring the numbers"), true);
  const none = copied(filed({ description: "" }));
  assert.equal(/DESCRIPTION/.test(none), false);
});

test("what the item is stays what it was", () => {
  const out = copied(filed());
  assert.equal(out.includes("UID:uid-redirected-1"), true);
  assert.equal(out.includes("DTSTART:20260901T100000Z"), true);
  assert.equal(out.includes("DTEND:20260901T110000Z"), true);
});

test("a blob we cannot read is handed back untouched", () => {
  // The hook is holding the user's save; an item we cannot rewrite is
  // stored as it came, which is exactly what happens today.
  for (const junk of ["", "not an ical at all", "BEGIN:VCALENDAR"]) {
    assert.equal(copied(junk), junk);
  }
});

test("without a prefix or an owner, only the guest list goes", () => {
  const out = plainCopyOfInvitation(filed(), {});
  assert.equal(/^ATTENDEE/m.test(out), false);
  assert.equal(/ORGANIZER/.test(out), false);
  assert.equal(out.includes("SUMMARY:Quarterly review"), true);
});
