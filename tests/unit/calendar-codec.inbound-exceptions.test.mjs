// Characterization tests for jobisoft#342: a recurrence exception that
// arrives as a status-only delta (Exchange doesn't repeat a field that
// hasn't changed since the exception's first sighting) loses that field
// locally instead of inheriting it, because `appendInboundExceptions`
// rebuilds each override VEVENT from an empty component every time an
// `<Exceptions>` block appears, with no fallback to the master's or the
// prior override's values.
//
// These tests assert TODAY'S (buggy) behavior on purpose - see
// TEST-PLAN.md. They exist so a fix for #342 is forced to touch this
// file (the assertions below will start failing) rather than silently
// landing unverified. Flip `todo: true` back off and update the
// assertions to the correct/expected values once the fix lands.

import { test, beforeAll } from "vitest";
import assert from "node:assert/strict";
import "../support/webext-shim.mjs";
import ICAL from "../../src/vendor/ical.min.js";
import { applicationDataToIcal } from "../../src/modules/eas/calendar-codec.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "../support/xml-node.mjs";

beforeAll(() => ensureLoaded());

const ADD_TEST35 = `<ApplicationData>
  <AllDayEvent xmlns='Calendar'>0</AllDayEvent>
  <DtStamp xmlns='Calendar'>20260731T182049Z</DtStamp>
  <StartTime xmlns='Calendar'>20260807T060000Z</StartTime>
  <Subject xmlns='Calendar'>test35</Subject>
  <UID xmlns='Calendar'>0400-test35-uid</UID>
  <OrganizerName xmlns='Calendar'>Tomas%20Kovacik</OrganizerName>
  <OrganizerEmail xmlns='Calendar'>tomas.kovacik%40dxc-tech.sk</OrganizerEmail>
  <Attendees xmlns='Calendar'>
    <Attendee xmlns='Calendar'>
      <Email xmlns='Calendar'>kovacik%40dgtfactory.com</Email>
      <AttendeeType xmlns='Calendar'>1</AttendeeType>
    </Attendee>
  </Attendees>
  <EndTime xmlns='Calendar'>20260807T063000Z</EndTime>
  <Recurrence xmlns='Calendar'>
    <Type xmlns='Calendar'>1</Type>
    <Interval xmlns='Calendar'>1</Interval>
    <Until xmlns='Calendar'>20270122T200000Z</Until>
    <DayOfWeek xmlns='Calendar'>32</DayOfWeek>
    <FirstDayOfWeek xmlns='Calendar'>0</FirstDayOfWeek>
  </Recurrence>
  <Sensitivity xmlns='Calendar'>0</Sensitivity>
  <BusyStatus xmlns='Calendar'>2</BusyStatus>
  <Reminder xmlns='Calendar'>15</Reminder>
  <MeetingStatus xmlns='Calendar'>3</MeetingStatus>
  <ResponseType xmlns='Calendar'>3</ResponseType>
</ApplicationData>`;

// First sighting of the 2026-08-07 occurrence's exception: full fields,
// including Subject - exactly what Exchange sends the first time.
const CHANGE_FIRST_EXCEPTION = `<ApplicationData>
  <Exceptions xmlns='Calendar'>
    <Exception xmlns='Calendar'>
      <InstanceId xmlns='AirSyncBase'>20260807T060000Z</InstanceId>
      <StartTime xmlns='Calendar'>20260807T060000Z</StartTime>
      <Subject xmlns='Calendar'>Canceled%3A%20test35</Subject>
      <EndTime xmlns='Calendar'>20260807T063000Z</EndTime>
      <BusyStatus xmlns='Calendar'>0</BusyStatus>
      <MeetingStatus xmlns='Calendar'>7</MeetingStatus>
    </Exception>
  </Exceptions>
</ApplicationData>`;

// Second touch of the SAME instance (e.g. the whole series then also
// gets cancelled): Exchange only resends what changed. Subject is
// unchanged since the first sighting, so it's omitted here - this is
// the real, captured shape of the delta that triggers #342.
const CHANGE_SECOND_TOUCH_DROPS_SUBJECT = `<ApplicationData>
  <Exceptions xmlns='Calendar'>
    <Exception xmlns='Calendar'>
      <InstanceId xmlns='AirSyncBase'>20260807T060000Z</InstanceId>
      <StartTime xmlns='Calendar'>20260807T060000Z</StartTime>
      <EndTime xmlns='Calendar'>20260807T063000Z</EndTime>
      <MeetingStatus xmlns='Calendar'>7</MeetingStatus>
    </Exception>
  </Exceptions>
</ApplicationData>`;

function overrideVevent(icalString) {
  const vcal = new ICAL.Component(ICAL.parse(icalString));
  return vcal
    .getAllSubcomponents("vevent")
    .find((v) => v.getFirstProperty("recurrence-id"));
}

test(
  "jobisoft#342: a second, status-only touch of an exception drops its Subject instead of inheriting it",
  () => {
    const commonArgs = {
      serverID: "server-id-test35",
      asVersion: "16.1",
      defaultTimezone: "UTC",
      syncRecurrence: true,
      uid: null,
      userEmail: "kovacik@dgtfactory.com",
    };

    const afterAdd = applicationDataToIcal({
      adNode: parseAdNode(ADD_TEST35),
      existingIcal: null,
      ...commonArgs,
    });

    const afterFirstException = applicationDataToIcal({
      adNode: parseAdNode(CHANGE_FIRST_EXCEPTION),
      existingIcal: afterAdd,
      ...commonArgs,
    });

    // Sanity check: the exception's first sighting DOES carry the title.
    const firstOverride = overrideVevent(afterFirstException);
    assert.equal(
      firstOverride.getFirstPropertyValue("summary"),
      "Canceled: test35",
      "first sighting of the exception should carry its own Subject",
    );

    const afterSecondTouch = applicationDataToIcal({
      adNode: parseAdNode(CHANGE_SECOND_TOUCH_DROPS_SUBJECT),
      existingIcal: afterFirstException,
      ...commonArgs,
    });

    const secondOverride = overrideVevent(afterSecondTouch);
    assert.ok(secondOverride, "expected the override VEVENT to still exist");

    // BUG (jobisoft#342): this should still read "Canceled: test35"
    // (inherited from the prior override / master), but today it comes
    // back null because appendInboundExceptions rebuilds the override
    // from an empty component and populateVeventFromAd only sets SUMMARY
    // when the update's XML actually mentions <Subject>.
    assert.equal(
      secondOverride.getFirstPropertyValue("summary"),
      null,
      "documents jobisoft#342: title is lost on the second, status-only touch",
    );
  },
);

// The sparsest real-world variant, captured live: an occurrence with no
// prior per-instance touch at all gets a bare status-only exception
// (e.g. the whole series is cancelled, and this occurrence's own
// exception has never been separately established before) - InstanceId
// and MeetingStatus only, no StartTime/EndTime/Subject/DtStamp
// whatsoever. Same root cause as above, but the resulting override has
// no DTSTART/DTEND either, not just no title.
const CHANGE_BARE_EXCEPTION = `<ApplicationData>
  <Exceptions xmlns='Calendar'>
    <Exception xmlns='Calendar'>
      <InstanceId xmlns='AirSyncBase'>20260814T060000Z</InstanceId>
      <Location xmlns='AirSyncBase'/>
      <MeetingStatus xmlns='Calendar'>7</MeetingStatus>
    </Exception>
  </Exceptions>
</ApplicationData>`;

test(
  "jobisoft#342 (sparsest variant): a bare status-only exception with no prior touch has no DTSTART/DTEND/SUMMARY at all",
  () => {
    const commonArgs = {
      serverID: "server-id-test35",
      asVersion: "16.1",
      defaultTimezone: "UTC",
      syncRecurrence: true,
      uid: null,
      userEmail: "kovacik@dgtfactory.com",
    };

    const afterAdd = applicationDataToIcal({
      adNode: parseAdNode(ADD_TEST35),
      existingIcal: null,
      ...commonArgs,
    });

    const afterBareException = applicationDataToIcal({
      adNode: parseAdNode(CHANGE_BARE_EXCEPTION),
      existingIcal: afterAdd,
      ...commonArgs,
    });

    const override = overrideVevent(afterBareException);
    assert.ok(override, "expected an override VEVENT for the instance");

    // BUG (jobisoft#342): all three should inherit from the master
    // (SUMMARY: "test35", DTSTART/DTEND: the occurrence's own computed
    // time from the RRULE), but today all three are simply absent -
    // appendInboundExceptions seeds the override from an empty
    // component and nothing in this delta mentions any of them.
    assert.equal(override.getFirstPropertyValue("summary"), null);
    assert.equal(override.getFirstPropertyValue("dtstart"), null);
    assert.equal(override.getFirstPropertyValue("dtend"), null);
  },
);
