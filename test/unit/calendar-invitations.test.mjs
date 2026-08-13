/**
 * Unit tests for recognising an invitation and reading the user's answer.
 *
 * The rule under test: a meeting somebody else organised is never sent as
 * an Add or a Change. On 16.0/16.1 the client may state neither the
 * organizer ([MS-ASCAL] 2.2.2.35) nor an attendee's status (2.2.2.5), and
 * the server substitutes the current user for both - so what arrives is not
 * their meeting changed but ours, re-invited to everyone on it. The only
 * thing we may say is the user's answer, as a MeetingResponse.
 *
 * Two signals decide it, because neither sees every item:
 *   - X-EAS-MEETINGSTATUS, recorded inbound on anything we pull. Blind to
 *     an item Thunderbird filed from an emailed invitation, which carries
 *     no X-EAS-* at all.
 *   - X-MOZ-INVITED-ATTENDEE against ORGANIZER, which is what the platform
 *     itself uses (calProviderBase.isInvitation) and the only thing that
 *     sees that item.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  isReceivedMeeting,
  preserveSelfPartstat,
  selfUserResponse,
} from "../../src/modules/eas/calendar-codec.mjs";

const ME = "john.bieling@cvjmbonn.de";
const ORG = "john.bieling@outlook.de";

/** An item as it looks after a pull: the server's own MeetingStatus. */
function pulled(meetingStatus, { partstat = "ACCEPTED", serverId = "sid-1" } = {}) {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    "UID:040000008200E00074C5B7101A82E008",
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20260901T100000Z",
    "SUMMARY:probe",
    `ORGANIZER;CN=Someone:mailto:${ORG}`,
    `ATTENDEE;PARTSTAT=${partstat};CN=${ME}:mailto:${ME}`,
    `X-EAS-SERVERID:${serverId}`,
    `X-EAS-MEETINGSTATUS:${meetingStatus}`,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

/** An invitation as Thunderbird files one from an emailed iTIP: no
 *  X-EAS-* anywhere, because it never came through our codec. */
function itip({ organizer = ORG, invited = ME, partstat = "ACCEPTED" } = {}) {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Microsoft Exchange//EN",
    "BEGIN:VEVENT",
    "UID:040000008200E00074C5B7101A82E008",
    "DTSTAMP:20260811T212553Z",
    "DTSTART:20260812T160000Z",
    "SUMMARY:Berlin 2",
    ...(organizer ? [`ORGANIZER;CN=John Bieling:mailto:${organizer}`] : []),
    ...(invited
      ? [
          `ATTENDEE;PARTSTAT=${partstat};CN=${invited}:mailto:${invited}`,
          `X-MOZ-INVITED-ATTENDEE:mailto:${invited}`,
        ]
      : []),
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

test("the server's own R bit is the answer, both ways", () => {
  assert.equal(isReceivedMeeting(pulled("3"), ME), true, "meeting + received");
  assert.equal(isReceivedMeeting(pulled("7"), ME), true, "the same, cancelled");
  // A clear R bit is a statement, not silence - and it outranks the address
  // comparison, which is what keeps an account whose organizer address is an
  // alias of its own from being locked out of its own meetings.
  assert.equal(isReceivedMeeting(pulled("1"), ME), false, "we organise it");
  assert.equal(isReceivedMeeting(pulled("0"), ME), false, "not a meeting");
});

test("an emailed invitation is caught with no X-EAS-* at all", () => {
  // The shape that matters: this is what actually sits in a calendar after
  // accepting an invitation, and the server's marker cannot see it.
  assert.equal(isReceivedMeeting(itip(), ME), true);
  // And without knowing our own address, because the item names it.
  assert.equal(isReceivedMeeting(itip(), null), true);
});

test("our own meetings stay ours", () => {
  assert.equal(
    isReceivedMeeting(itip({ organizer: ME, invited: ME }), ME),
    false,
    "we organised and invited ourselves",
  );
  assert.equal(
    isReceivedMeeting(itip({ organizer: ME.toUpperCase(), invited: ME }), ME),
    false,
    "case is not identity",
  );
  assert.equal(
    isReceivedMeeting(itip({ organizer: null }), ME),
    false,
    "no organizer names nobody, which is not somebody else",
  );
});

test("an item nobody has said anything about is ours", () => {
  assert.equal(isReceivedMeeting("BEGIN:VCALENDAR\r\nEND:VCALENDAR", ME), false);
  assert.equal(isReceivedMeeting("", ME), false);
  assert.equal(isReceivedMeeting("not a calendar", ME), false);
});

test("the answer maps to a UserResponse", () => {
  assert.equal(selfUserResponse(itip({ partstat: "ACCEPTED" }), ME), 1);
  assert.equal(selfUserResponse(itip({ partstat: "TENTATIVE" }), ME), 2);
  assert.equal(selfUserResponse(itip({ partstat: "DECLINED" }), ME), 3);
});

test("no answer is not an answer", () => {
  // The trap PR #339 was held up on: an unanswered invitation must not
  // produce a response, or every unrelated edit sends an Accept the user
  // never made. There is no UserResponse for "no reply", so this is
  // structural rather than a guard that could be forgotten.
  assert.equal(selfUserResponse(itip({ partstat: "NEEDS-ACTION" }), ME), null);
  assert.equal(selfUserResponse(itip({ invited: null }), ME), null);
  assert.equal(selfUserResponse("", ME), null);
});

test("the local answer survives adopting the server's copy", () => {
  // The pull runs before the response is sent, so the server's copy - which
  // does not know about the answer yet - would otherwise overwrite it, and
  // we would read that back and send it. Fails silently: the user accepts,
  // the calendar looks right, the organizer never hears.
  const adopted = preserveSelfPartstat({
    builtIcal: pulled("3", { partstat: "NEEDS-ACTION" }),
    priorIcal: itip({ partstat: "ACCEPTED" }),
    userEmail: ME,
  });
  assert.match(adopted, /PARTSTAT=ACCEPTED/);
  assert.equal(selfUserResponse(adopted, ME), 1);
});

test("but an answer we do not have does not overwrite one we do", () => {
  // Answered on another device: the server knows, we do not. Its copy wins,
  // because NEEDS-ACTION is the absence of an answer rather than one.
  const adopted = preserveSelfPartstat({
    builtIcal: pulled("3", { partstat: "DECLINED" }),
    priorIcal: itip({ partstat: "NEEDS-ACTION" }),
    userEmail: ME,
  });
  assert.match(adopted, /PARTSTAT=DECLINED/);
});
