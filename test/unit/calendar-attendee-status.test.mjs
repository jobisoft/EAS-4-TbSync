/**
 * Unit tests for the two EAS enums that both mean "has this person
 * answered?" and are not the same enum.
 *
 * The rule under test: `AttendeeStatus` ([MS-ASCAL] 2.2.2.5) and
 * `ResponseType` (2.2.2.40) agree on 2, 3 and 4 and part company either
 * side of them. `5` is "Not responded" in both; `1` is "Organizer - no
 * reply required" and exists only in ResponseType.
 *
 * One table used to serve both, mapping 5 to ACCEPTED - so an invitation
 * whose per-attendee status the server omits, falling back to a
 * ResponseType of 5, was written into the calendar as "you accepted".
 * That is a lie on its own, and once answers are sent it becomes an Accept
 * the user never made: the defect recorded as the first merge condition on
 * PR #339.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { applicationDataToIcal } from "../../src/modules/eas/calendar-codec.mjs";
import { el } from "./support/ad-node.mjs";

const ME = "john.bieling@ekir.de";

/** An invitation as the server sends one: the organizer is somebody else,
 *  and the user is in the attendee list. Omitting `attendeeStatus` is the
 *  case that falls back to the event-level ResponseType. */
function invitation({ attendeeStatus = null, responseType = null } = {}) {
  return el("ApplicationData", [
    el("UID", "040000008200E00074C5B7101A82E008"),
    el("Subject", "Hamburg"),
    el("StartTime", "20260901T100000Z"),
    el("EndTime", "20260901T110000Z"),
    el("OrganizerEmail", "john.bieling@outlook.de"),
    el("OrganizerName", "John Bieling"),
    el("MeetingStatus", "3"),
    ...(responseType === null ? [] : [el("ResponseType", String(responseType))]),
    el("Attendees", [
      el("Attendee", [
        el("Email", ME),
        el("Name", "John Bieling"),
        ...(attendeeStatus === null
          ? []
          : [el("AttendeeStatus", String(attendeeStatus))]),
      ]),
    ]),
  ]);
}

async function partstatOf(ad) {
  const ical = await applicationDataToIcal({
    adNode: ad,
    serverID: "sid-1",
    asVersion: "14.1",
    uid: "040000008200E00074C5B7101A82E008",
    userEmail: ME,
  });
  // Unfold first: iCalendar wraps at 75 octets, and an ATTENDEE line
  // carrying an address and parameters is routinely longer than that, so
  // matching raw lines finds the property without its value.
  const lines = [];
  for (const raw of ical.split(/\r?\n/)) {
    if (/^[ \t]/.test(raw) && lines.length) lines[lines.length - 1] += raw.slice(1);
    else lines.push(raw);
  }
  const line = lines.find((l) => l.startsWith("ATTENDEE") && l.includes(ME));
  return /PARTSTAT=([A-Z-]+)/.exec(line ?? "")?.[1] ?? null;
}

test("an unanswered invitation is not an accepted one", async () => {
  // The case that was wrong: no per-attendee status, ResponseType 5, which
  // means "the user has not yet responded to the meeting request".
  assert.equal(await partstatOf(invitation({ responseType: 5 })), "NEEDS-ACTION");
  // And the same value in the other vocabulary.
  assert.equal(await partstatOf(invitation({ attendeeStatus: 5 })), "NEEDS-ACTION");
});

test("a real answer is carried through, from either element", async () => {
  assert.equal(await partstatOf(invitation({ attendeeStatus: 2 })), "TENTATIVE");
  assert.equal(await partstatOf(invitation({ attendeeStatus: 3 })), "ACCEPTED");
  assert.equal(await partstatOf(invitation({ attendeeStatus: 4 })), "DECLINED");
  assert.equal(await partstatOf(invitation({ responseType: 2 })), "TENTATIVE");
  assert.equal(await partstatOf(invitation({ responseType: 3 })), "ACCEPTED");
  assert.equal(await partstatOf(invitation({ responseType: 4 })), "DECLINED");
});

test("no answer at all is no answer", async () => {
  assert.equal(await partstatOf(invitation({ responseType: 0 })), "NEEDS-ACTION");
  assert.equal(await partstatOf(invitation({ attendeeStatus: 0 })), "NEEDS-ACTION");
  assert.equal(await partstatOf(invitation()), "NEEDS-ACTION");
});

test("the organizer is not waiting on themselves", async () => {
  // ResponseType 1 only: "the current user is the organizer of the meeting
  // and, therefore, no reply is required". AttendeeStatus has no 1 at all,
  // which is why these cannot share a table.
  assert.equal(await partstatOf(invitation({ responseType: 1 })), "ACCEPTED");
});
