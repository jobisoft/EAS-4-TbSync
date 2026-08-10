/**
 * Ported from PR #345 (tomaskovacik) to the node:test layer; fixtures
 * kept verbatim (several are live-server captures), expectations
 * re-verified against current master.
 */

// Round-trip tests for applicationDataToIcal against a real, captured
// EAS <Add> payload (verbatim from a live Exchange Online test run) -
// proves the harness works end to end before we use it to pin down bugs.

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import ICAL from "../../src/vendor/ical.min.js";
import { applicationDataToIcal } from "../../src/modules/eas/calendar-codec.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";

before(() => ensureLoaded());

// Captured live: kovo1@dgtfactory.com invites kovacik@dgtfactory.com to a
// weekly recurring meeting, EAS 16.1, unresponded.
const ADD_TEST24 = `<ApplicationData>
  <AllDayEvent xmlns='Calendar'>0</AllDayEvent>
  <TimeZone xmlns='Calendar'>xP%2F%2F%2F0MAZQBuAHQAcgBhAGwAIABFAHUAcgBvAHAAZQAgAFMAdABhAG4AZABhAHIAZAAgAFQAaQBtAGUAAAAAAAAAAAAAAAoAAAAFAAMAAAAAAAAAAAAAACgAVQBUAEMAKwAwADEAOgAwADAAKQAgAEIAZQBsAGcAcgBhAGQAZQAsACAAQgByAGEAdABpAHMAbABhAHYAYQAAAAMAAAAFAAIAAAAAAAAAxP%2F%2F%2Fw%3D%3D</TimeZone>
  <DtStamp xmlns='Calendar'>20260731T163448Z</DtStamp>
  <StartTime xmlns='Calendar'>20260731T133000Z</StartTime>
  <Subject xmlns='Calendar'>test24</Subject>
  <UID xmlns='Calendar'>040000008200E00074C5B7101A82E00800000000E9ABA67E0A21DD01000000000000000010000000977A7C48535D7E4BB962D5FAB0E784A9</UID>
  <OrganizerName xmlns='Calendar'>test%20kovo1</OrganizerName>
  <OrganizerEmail xmlns='Calendar'>kovo1%40dgtfactory.com</OrganizerEmail>
  <Attendees xmlns='Calendar'>
    <Attendee xmlns='Calendar'>
      <Email xmlns='Calendar'>kovacik%40dgtfactory.com</Email>
      <Name xmlns='Calendar'>Tom%C3%83%C2%A1%C3%85%C2%A1%20Kov%C3%83%C2%A1%C3%84%C2%8Dik</Name>
      <AttendeeType xmlns='Calendar'>1</AttendeeType>
    </Attendee>
  </Attendees>
  <Location xmlns='AirSyncBase'/>
  <EndTime xmlns='Calendar'>20260731T140000Z</EndTime>
  <Recurrence xmlns='Calendar'>
    <Type xmlns='Calendar'>1</Type>
    <Interval xmlns='Calendar'>1</Interval>
    <Until xmlns='Calendar'>20270122T143000Z</Until>
    <DayOfWeek xmlns='Calendar'>32</DayOfWeek>
    <FirstDayOfWeek xmlns='Calendar'>0</FirstDayOfWeek>
  </Recurrence>
  <Sensitivity xmlns='Calendar'>0</Sensitivity>
  <BusyStatus xmlns='Calendar'>1</BusyStatus>
  <Reminder xmlns='Calendar'>15</Reminder>
  <MeetingStatus xmlns='Calendar'>3</MeetingStatus>
  <NativeBodyType xmlns='AirSyncBase'>2</NativeBodyType>
  <ResponseRequested xmlns='Calendar'>1</ResponseRequested>
  <ResponseType xmlns='Calendar'>5</ResponseType>
</ApplicationData>`;

function masterVevent(icalString) {
  const vcal = new ICAL.Component(ICAL.parse(icalString));
  return vcal
    .getAllSubcomponents("vevent")
    .find((v) => !v.getFirstProperty("recurrence-id"));
}

test("applicationDataToIcal: fresh Add maps Subject/Organizer/Attendee/times", async () => {
  const adNode = parseAdNode(ADD_TEST24);
  const icalString = await applicationDataToIcal({
    adNode,
    existingIcal: null,
    serverID: "server-id-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: null,
    userEmail: "kovacik@dgtfactory.com",
  });

  const vevent = masterVevent(icalString);
  assert.ok(vevent, "expected a master VEVENT with no RECURRENCE-ID");
  assert.equal(vevent.getFirstPropertyValue("summary"), "test24");

  const organizer = vevent.getFirstProperty("organizer");
  assert.ok(organizer, "expected an ORGANIZER property");
  assert.equal(organizer.getFirstValue(), "mailto:kovo1@dgtfactory.com");
  assert.equal(organizer.getParameter("cn"), "test kovo1");

  const attendees = vevent.getAllProperties("attendee");
  assert.equal(attendees.length, 1);
  assert.equal(attendees[0].getFirstValue(), "mailto:kovacik@dgtfactory.com");

  const dtstart = vevent.getFirstPropertyValue("dtstart");
  assert.equal(dtstart.toString(), "2026-07-31T13:30:00Z");

  assert.ok(
    vevent.getFirstPropertyValue("rrule"),
    "expected an RRULE from the Recurrence block",
  );
});

test("applicationDataToIcal: a delta-only Change (DtStamp only) leaves other fields untouched", async () => {
  const first = await applicationDataToIcal({
    adNode: parseAdNode(ADD_TEST24),
    existingIcal: null,
    serverID: "server-id-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: null,
    userEmail: "kovacik@dgtfactory.com",
  });

  const deltaOnly = parseAdNode(
    `<ApplicationData><DtStamp xmlns='Calendar'>20260731T170000Z</DtStamp></ApplicationData>`,
  );
  const second = await applicationDataToIcal({
    adNode: deltaOnly,
    existingIcal: first,
    serverID: "server-id-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: null,
    userEmail: "kovacik@dgtfactory.com",
  });

  const vevent = masterVevent(second);
  assert.equal(
    vevent.getFirstPropertyValue("summary"),
    "test24",
    "Subject should survive a delta that doesn't mention it",
  );
  const organizer = vevent.getFirstProperty("organizer");
  assert.equal(
    organizer?.getFirstValue(),
    "mailto:kovo1@dgtfactory.com",
    "Organizer should survive a delta that doesn't mention it",
  );
});
