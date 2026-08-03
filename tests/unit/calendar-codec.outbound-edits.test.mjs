// Outbound iCal -> EAS: edit/reschedule a single event, edit a whole
// recurring series' RRULE, and the 16.1 single-occurrence shape
// (appendInstanceChanges) - the exact wire shape at the center of
// jobisoft#334 (Exchange rejects it; see TEST-PLAN.md / project memory).
//
// These round-trip through the REAL WBXML encoder (wbxml.mjs's
// createWBXML) and the REAL decoder (decodeWBXML) - no mocking, this
// exercises the actual wire format, not just the codec's internal
// intent.

import { test, beforeAll } from "vitest";
import assert from "node:assert/strict";
import "../support/webext-shim.mjs";
import {
  appendApplicationDataFromIcal,
  appendInstanceChanges,
} from "../../src/modules/eas/calendar-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "../support/xml-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

beforeAll(() => ensureLoaded());

const SINGLE_EVENT_ICS = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-outbound-single-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T140000Z
DTEND:20260801T143000Z
SUMMARY:test-outbound-single (moved)
END:VEVENT
END:VCALENDAR
`;

/** Wraps a codec append call the same way sync-runner.mjs's
 *  appendCommands does for a <Change>: <ApplicationData> ... </>,
 *  switch back to AirSync, close. Returns the decoded node for the
 *  <ApplicationData> element so callers can readPathFrom it directly. */
function roundTripApplicationData(appendFn) {
  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendFn(w);
  w.switchpage("AirSync");
  w.ctag();
  const xml = decodeWBXML(w.getBytes());
  return parseAdNode(xml);
}

/** Parses the codec's own YYYYMMDDTHHMMSSZ ("basic UTC") format. */
function parseBasicUtc(s) {
  const m = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$/.exec(s);
  if (!m) throw new Error(`not a basic-UTC timestamp: ${s}`);
  const [, y, mo, d, h, mi, se] = m.map(Number);
  return new Date(Date.UTC(y, mo - 1, d, h, mi, se));
}

test(
  "a VEVENT with no DTSTART/DTEND (the #342 corrupted-override shape) " +
    "gets pushed with StartTime==EndTime==now(), not omitted",
  () => {
    // No DTSTART/DTEND at all - exactly what #342's sparsest variant
    // leaves an override with (see calendar-codec.inbound-exceptions
    // .test.mjs). This is the malformed shape that got rejected with a
    // top-level Status 4 in the live log tonight.
    const noStartEndIcs = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-outbound-no-start-end-uid
SUMMARY:
END:VEVENT
END:VCALENDAR
`;

    const before = new Date();
    const node = roundTripApplicationData((w) =>
      appendApplicationDataFromIcal({
        builder: w,
        ical: noStartEndIcs,
        asVersion: "16.1",
        defaultTimezone: "UTC",
        syncRecurrence: false,
        userEmail: "kovacik@dgtfactory.com",
      }),
    );
    const after = new Date();

    const startTime = parseBasicUtc(readPathFrom(node, ["StartTime"]));
    const endTime = parseBasicUtc(readPathFrom(node, ["EndTime"]));

    assert.equal(startTime.getTime(), endTime.getTime());
    assert.ok(
      startTime.getTime() >= before.getTime() - 1000 &&
        startTime.getTime() <= after.getTime() + 1000,
      `expected StartTime (${startTime.toISOString()}) to be close to ` +
        `now (test ran ${before.toISOString()} - ${after.toISOString()})`,
    );
  },
);

test("single event: appendApplicationDataFromIcal emits the rescheduled Start/End/Subject", () => {
  const node = roundTripApplicationData((w) =>
    appendApplicationDataFromIcal({
      builder: w,
      ical: SINGLE_EVENT_ICS,
      asVersion: "16.1",
      defaultTimezone: "UTC",
      syncRecurrence: false,
      userEmail: "kovacik@dgtfactory.com",
    }),
  );

  assert.equal(
    readPathFrom(node, ["Subject"]),
    "test-outbound-single (moved)",
  );
  assert.equal(readPathFrom(node, ["StartTime"]), "20260801T140000Z");
  assert.equal(readPathFrom(node, ["EndTime"]), "20260801T143000Z");
});

test(
  "16.1 Organizer-emission gate: OrganizerName/OrganizerEmail are never " +
    "emitted outbound on 16.1, even though a real 16.1 <Add> DOES carry " +
    "them inbound (see calendar-codec.basic.test.mjs)",
  () => {
    const withOrganizerIcs = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-outbound-organizer-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T140000Z
DTEND:20260801T143000Z
SUMMARY:test-outbound-organizer
ORGANIZER;CN=test kovo1:mailto:kovo1@dgtfactory.com
END:VEVENT
END:VCALENDAR
`;

    const node16_1 = roundTripApplicationData((w) =>
      appendApplicationDataFromIcal({
        builder: w,
        ical: withOrganizerIcs,
        asVersion: "16.1",
        defaultTimezone: "UTC",
        syncRecurrence: false,
        userEmail: "kovacik@dgtfactory.com",
      }),
    );
    assert.equal(readPathFrom(node16_1, ["OrganizerEmail"]), null);
    assert.equal(readPathFrom(node16_1, ["OrganizerName"]), null);

    // Contrast: on <=14.x the same local ORGANIZER IS emitted - confirms
    // the gate is asVersion-specific, not a blanket "never emit" bug.
    const node14_1 = roundTripApplicationData((w) =>
      appendApplicationDataFromIcal({
        builder: w,
        ical: withOrganizerIcs,
        asVersion: "14.1",
        defaultTimezone: "UTC",
        syncRecurrence: false,
        userEmail: "kovacik@dgtfactory.com",
      }),
    );
    assert.equal(readPathFrom(node14_1, ["OrganizerEmail"]), "kovo1@dgtfactory.com");
    assert.equal(readPathFrom(node14_1, ["OrganizerName"]), "test kovo1");
  },
);

const RECURRING_SERIES_ICS = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-outbound-series-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T100000Z
DTEND:20260801T103000Z
SUMMARY:test-outbound-series
RRULE:FREQ=WEEKLY;INTERVAL=1;UNTIL=20260901T100000Z;BYDAY=SA
END:VEVENT
END:VCALENDAR
`;

test("whole series: appendApplicationDataFromIcal emits a Recurrence block for the RRULE", () => {
  const node = roundTripApplicationData((w) =>
    appendApplicationDataFromIcal({
      builder: w,
      ical: RECURRING_SERIES_ICS,
      asVersion: "16.1",
      defaultTimezone: "UTC",
      syncRecurrence: true,
      userEmail: "kovacik@dgtfactory.com",
    }),
  );

  assert.equal(readPathFrom(node, ["Recurrence", "Type"]), "1");
  assert.equal(readPathFrom(node, ["Recurrence", "Interval"]), "1");
  assert.equal(readPathFrom(node, ["Recurrence", "Until"]), "20260901T100000Z");
});

test(
  "jobisoft#334: 16.1 single-occurrence delete is emitted as a standalone " +
    "<Change> sharing the master's ServerId - the exact shape Exchange " +
    "rejects (Status 6), confirmed live; not yet fixed",
  () => {
    const seriesWithExdate = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-outbound-series-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T100000Z
DTEND:20260801T103000Z
SUMMARY:test-outbound-series
RRULE:FREQ=WEEKLY;INTERVAL=1;UNTIL=20260901T100000Z;BYDAY=SA
EXDATE:20260808T100000Z
END:VEVENT
END:VCALENDAR
`;

    const w = createWBXML("AirSync");
    appendInstanceChanges({
      builder: w,
      blob: seriesWithExdate,
      serverID: "server-id-series",
      asVersion: "16.1",
      defaultTimezone: "UTC",
      syncRecurrence: true,
      userEmail: "kovacik@dgtfactory.com",
    });
    const xml = decodeWBXML(w.getBytes());
    const changeNode = parseAdNode(xml);

    assert.equal(changeNode.tagName, "Change");
    // Same ServerId as the master's own <Change> - this is exactly the
    // "shares the master's ServerId" shape #334's root-cause writeup
    // describes, kept as-is here (not split into a separate Sync
    // request) because appendInstanceChanges only emits the <Change>
    // command itself; request-level bundling is sync-runner.mjs's job.
    assert.equal(readPathFrom(changeNode, ["ServerId"]), "server-id-series");
    const ad = changeNode.children.find((c) => c.tagName === "ApplicationData");
    assert.equal(readPathFrom(ad, ["InstanceId"]), "20260808T100000Z");
    assert.equal(readPathFrom(ad, ["Deleted"]), "1");
  },
);

test(
  "jobisoft#334: 16.1 single-occurrence EDIT (not delete) is also a " +
    "standalone <Change> sharing the master's ServerId - same shape, " +
    "same unresolved rejection risk",
  () => {
    // Master + one RECURRENCE-ID override representing an edited (not
    // deleted) occurrence - appendInstanceChanges' "overrides" branch,
    // untested until now (only its "exdates" branch was covered above).
    const seriesWithOverride = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-outbound-series-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T100000Z
DTEND:20260801T103000Z
SUMMARY:test-outbound-series
RRULE:FREQ=WEEKLY;INTERVAL=1;UNTIL=20260901T100000Z;BYDAY=SA
END:VEVENT
BEGIN:VEVENT
UID:0400-outbound-series-uid
RECURRENCE-ID:20260808T100000Z
DTSTAMP:20260801T090000Z
DTSTART:20260808T140000Z
DTEND:20260808T143000Z
SUMMARY:test-outbound-series (one occurrence moved)
END:VEVENT
END:VCALENDAR
`;

    const w = createWBXML("AirSync");
    appendInstanceChanges({
      builder: w,
      blob: seriesWithOverride,
      serverID: "server-id-series",
      asVersion: "16.1",
      defaultTimezone: "UTC",
      syncRecurrence: true,
      userEmail: "kovacik@dgtfactory.com",
    });
    const xml = decodeWBXML(w.getBytes());
    const changeNode = parseAdNode(xml);

    assert.equal(changeNode.tagName, "Change");
    assert.equal(readPathFrom(changeNode, ["ServerId"]), "server-id-series");
    const ad = changeNode.children.find((c) => c.tagName === "ApplicationData");
    assert.equal(readPathFrom(ad, ["InstanceId"]), "20260808T100000Z");
    assert.equal(
      readPathFrom(ad, ["Subject"]),
      "test-outbound-series (one occurrence moved)",
    );
    assert.equal(readPathFrom(ad, ["StartTime"]), "20260808T140000Z");
    // isException:true suppresses Organizer/Recurrence re-emission on the
    // override body - confirm that suppression is actually happening
    // rather than silently duplicating the master's own fields here.
    assert.equal(readPathFrom(ad, ["Recurrence", "Type"]), null);
  },
);
