// The 2.5/14.x embedded <Exceptions> path - appendInboundExceptions'
// <Deleted>1</Deleted> handling and appendOutboundExceptions - as
// opposed to the 16.1 standalone-<Change> path already covered by
// calendar-codec.inbound-edits.test.mjs / outbound-edits.test.mjs
// (applyInstanceChange/applyInstanceDelete/appendInstanceChanges).
// Genuinely untested before this file (see TEST-PLAN.md's coverage
// matrix). Also what Exchange sends for occurrence deletes even
// against a 16.1 account when the delete happens via OWA, per live
// cross-checks during the jobisoft#334 investigation - so despite the
// "2.5/14.x" label, this shape shows up inbound on modern accounts too.

import { test, beforeAll } from "vitest";
import assert from "node:assert/strict";
import "../support/webext-shim.mjs";
import ICAL from "../../src/vendor/ical.min.js";
import { applicationDataToIcal, appendApplicationDataFromIcal } from "../../src/modules/eas/calendar-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "../support/xml-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

beforeAll(() => ensureLoaded());

const COMMON = {
  serverID: "server-id-embedded",
  defaultTimezone: "UTC",
  syncRecurrence: true,
  uid: null,
  userEmail: "kovacik@dgtfactory.com",
};

const ADD_RECURRING_SERIES = `<ApplicationData>
  <AllDayEvent xmlns='Calendar'>0</AllDayEvent>
  <DtStamp xmlns='Calendar'>20260801T090000Z</DtStamp>
  <StartTime xmlns='Calendar'>20260801T100000Z</StartTime>
  <Subject xmlns='Calendar'>test-embedded-series</Subject>
  <UID xmlns='Calendar'>0400-embedded-series-uid</UID>
  <EndTime xmlns='Calendar'>20260801T103000Z</EndTime>
  <Recurrence xmlns='Calendar'>
    <Type xmlns='Calendar'>1</Type>
    <Interval xmlns='Calendar'>1</Interval>
    <Until xmlns='Calendar'>20261231T100000Z</Until>
    <DayOfWeek xmlns='Calendar'>32</DayOfWeek>
    <FirstDayOfWeek xmlns='Calendar'>0</FirstDayOfWeek>
  </Recurrence>
  <Sensitivity xmlns='Calendar'>0</Sensitivity>
  <BusyStatus xmlns='Calendar'>2</BusyStatus>
  <MeetingStatus xmlns='Calendar'>0</MeetingStatus>
</ApplicationData>`;

function masterVevent(icalString) {
  const vcal = new ICAL.Component(ICAL.parse(icalString));
  return vcal
    .getAllSubcomponents("vevent")
    .find((v) => !v.getFirstProperty("recurrence-id"));
}
function overrideByRecurrenceId(icalString, isoDate) {
  const vcal = new ICAL.Component(ICAL.parse(icalString));
  return vcal
    .getAllSubcomponents("vevent")
    .find((v) => v.getFirstPropertyValue("recurrence-id")?.toString() === isoDate);
}

/* ── Inbound ──────────────────────────────────────────────────────── */

test("inbound: an embedded Exception identified by ExceptionStartTime (2.5/14.x, not InstanceId) with Deleted=1 adds an EXDATE", () => {
  const afterAdd = applicationDataToIcal({
    adNode: parseAdNode(ADD_RECURRING_SERIES),
    existingIcal: null,
    asVersion: "14.1",
    ...COMMON,
  });

  const deleteException = parseAdNode(`<ApplicationData>
    <Exceptions xmlns='Calendar'>
      <Exception xmlns='Calendar'>
        <ExceptionStartTime xmlns='Calendar'>20260808T100000Z</ExceptionStartTime>
        <Deleted xmlns='Calendar'>1</Deleted>
      </Exception>
    </Exceptions>
  </ApplicationData>`);

  const afterDelete = applicationDataToIcal({
    adNode: deleteException,
    existingIcal: afterAdd,
    asVersion: "14.1",
    ...COMMON,
  });

  const master = masterVevent(afterDelete);
  const exdates = master
    .getAllProperties("exdate")
    .map((p) => p.getFirstValue().toString());
  assert.deepEqual(exdates, ["2026-08-08T10:00:00Z"]);
  assert.equal(
    overrideByRecurrenceId(afterDelete, "2026-08-08T10:00:00Z"),
    undefined,
    "a deleted instance must not also leave an override behind",
  );
});

test("inbound: an embedded Exception without Deleted creates an override, keyed the same way as the InstanceId form", () => {
  const afterAdd = applicationDataToIcal({
    adNode: parseAdNode(ADD_RECURRING_SERIES),
    existingIcal: null,
    asVersion: "14.1",
    ...COMMON,
  });

  const editException = parseAdNode(`<ApplicationData>
    <Exceptions xmlns='Calendar'>
      <Exception xmlns='Calendar'>
        <ExceptionStartTime xmlns='Calendar'>20260808T100000Z</ExceptionStartTime>
        <StartTime xmlns='Calendar'>20260808T140000Z</StartTime>
        <EndTime xmlns='Calendar'>20260808T143000Z</EndTime>
        <Subject xmlns='Calendar'>test-embedded-series (one occurrence moved)</Subject>
      </Exception>
    </Exceptions>
  </ApplicationData>`);

  const afterEdit = applicationDataToIcal({
    adNode: editException,
    existingIcal: afterAdd,
    asVersion: "14.1",
    ...COMMON,
  });

  const override = overrideByRecurrenceId(afterEdit, "2026-08-08T10:00:00Z");
  assert.ok(override, "expected an override VEVENT for the edited occurrence");
  assert.equal(
    override.getFirstPropertyValue("summary"),
    "test-embedded-series (one occurrence moved)",
  );
  assert.equal(
    override.getFirstPropertyValue("dtstart").toString(),
    "2026-08-08T14:00:00Z",
  );
});

/* ── Outbound ─────────────────────────────────────────────────────── */

function roundTripApplicationData(icalString, asVersion) {
  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendApplicationDataFromIcal({
    builder: w,
    ical: icalString,
    asVersion,
    defaultTimezone: "UTC",
    syncRecurrence: true,
    userEmail: "kovacik@dgtfactory.com",
  });
  w.switchpage("AirSync");
  w.ctag();
  return parseAdNode(decodeWBXML(w.getBytes()));
}

test("outbound (14.1): a master with an EXDATE emits an EMBEDDED <Exceptions><Exception> inside the same ApplicationData, not a separate command", () => {
  const seriesWithExdate = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-embedded-outbound-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T100000Z
DTEND:20260801T103000Z
SUMMARY:test-embedded-outbound
RRULE:FREQ=WEEKLY;INTERVAL=1;UNTIL=20261231T100000Z;BYDAY=SA
EXDATE:20260808T100000Z
END:VEVENT
END:VCALENDAR
`;

  const node = roundTripApplicationData(seriesWithExdate, "14.1");
  assert.equal(node.tagName, "ApplicationData");
  const exceptions = node.children.find((c) => c.tagName === "Exceptions");
  assert.ok(exceptions, "expected an embedded Exceptions wrapper");
  const exception = exceptions.children.find((c) => c.tagName === "Exception");
  assert.equal(
    readPathFrom(exception, ["ExceptionStartTime"]),
    "20260808T100000Z",
  );
  assert.equal(readPathFrom(exception, ["Deleted"]), "1");
});

test("outbound (14.1): a master with a RECURRENCE-ID override emits an embedded Exception carrying the override's own fields", () => {
  const seriesWithOverride = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-embedded-outbound-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T100000Z
DTEND:20260801T103000Z
SUMMARY:test-embedded-outbound
RRULE:FREQ=WEEKLY;INTERVAL=1;UNTIL=20261231T100000Z;BYDAY=SA
END:VEVENT
BEGIN:VEVENT
UID:0400-embedded-outbound-uid
RECURRENCE-ID:20260808T100000Z
DTSTAMP:20260801T090000Z
DTSTART:20260808T140000Z
DTEND:20260808T143000Z
SUMMARY:test-embedded-outbound (moved)
END:VEVENT
END:VCALENDAR
`;

  const node = roundTripApplicationData(seriesWithOverride, "14.1");
  const exceptions = node.children.find((c) => c.tagName === "Exceptions");
  assert.ok(exceptions, "expected an embedded Exceptions wrapper");
  const exception = exceptions.children.find((c) => c.tagName === "Exception");
  assert.equal(
    readPathFrom(exception, ["ExceptionStartTime"]),
    "20260808T100000Z",
  );
  assert.equal(
    readPathFrom(exception, ["Subject"]),
    "test-embedded-outbound (moved)",
  );
  assert.equal(readPathFrom(exception, ["StartTime"]), "20260808T140000Z");
});

test("outbound (16.1): the same EXDATE/override data does NOT get an embedded Exceptions wrapper - 16.1 handles exceptions as separate <Change> commands at the orchestrator level (appendInstanceChanges), not inside ApplicationData", () => {
  const seriesWithExdate = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-embedded-outbound-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T100000Z
DTEND:20260801T103000Z
SUMMARY:test-embedded-outbound
RRULE:FREQ=WEEKLY;INTERVAL=1;UNTIL=20261231T100000Z;BYDAY=SA
EXDATE:20260808T100000Z
END:VEVENT
END:VCALENDAR
`;

  const node = roundTripApplicationData(seriesWithExdate, "16.1");
  assert.equal(
    node.children.find((c) => c.tagName === "Exceptions"),
    undefined,
  );
});
