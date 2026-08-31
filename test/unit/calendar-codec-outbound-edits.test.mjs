/**
 * Ported from PR #345 (tomaskovacik) to the node:test layer; fixtures
 * kept verbatim (several are live-server captures), expectations
 * re-verified against current master.
 */

// Outbound iCal -> EAS: edit/reschedule a single event, edit a whole
// recurring series' RRULE, and the 16.1 single-occurrence shape
// (listInstanceCommands) - the exact wire shape at the center of
//
// These round-trip through the REAL WBXML encoder (wbxml.mjs's
// createWBXML) and the REAL decoder (decodeWBXML) - no mocking, this
// exercises the actual wire format, not just the codec's internal
// intent.

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import {
  appendApplicationDataFromIcal,
  listInstanceCommands,
} from "../../src/modules/eas/calendar-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

before(() => ensureLoaded());

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
    // top-level Status 4 in a live log (9 Aug 2026).
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
        userEmail: "user@example.invalid",
      }),
    );
    const after = new Date();

    const startTime = parseBasicUtc(readPathFrom(node, ["StartTime"]));
    const endTime = parseBasicUtc(readPathFrom(node, ["EndTime"]));
    // Both fall inside the test's own time window. NOT asserted equal to
    // each other: startTimeFor and endTimeFor each read the clock, so a
    // second boundary between the two calls is a designed-in possibility.
    for (const [name, value] of [
      ["StartTime", startTime],
      ["EndTime", endTime],
    ]) {
      assert.ok(
        value.getTime() >= before.getTime() - 1000 &&
          value.getTime() <= after.getTime() + 1000,
        `expected ${name} (${value.toISOString()}) to be close to ` +
          `now (test ran ${before.toISOString()} - ${after.toISOString()})`,
      );
    }
  },
);

test("single event: appendApplicationDataFromIcal emits the rescheduled Start/End/Subject", () => {
  const node = roundTripApplicationData((w) =>
    appendApplicationDataFromIcal({
      builder: w,
      ical: SINGLE_EVENT_ICS,
      asVersion: "16.1",
      defaultTimezone: "UTC",
      userEmail: "user@example.invalid",
    }),
  );

  assert.equal(readPathFrom(node, ["Subject"]), "test-outbound-single (moved)");
  assert.equal(readPathFrom(node, ["StartTime"]), "20260801T140000Z");
  assert.equal(readPathFrom(node, ["EndTime"]), "20260801T143000Z");
});

test(
  "16.1 Organizer-emission gate: OrganizerName/OrganizerEmail are never " +
    "emitted outbound on 16.1, even though a real 16.1 <Add> DOES carry " +
    "them inbound (see calendar-codec-basic.test.mjs)",
  () => {
    const withOrganizerIcs = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-outbound-organizer-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T140000Z
DTEND:20260801T143000Z
SUMMARY:test-outbound-organizer
ORGANIZER;CN=test organizer:mailto:organizer@example.invalid
END:VEVENT
END:VCALENDAR
`;

    const node16_1 = roundTripApplicationData((w) =>
      appendApplicationDataFromIcal({
        builder: w,
        ical: withOrganizerIcs,
        asVersion: "16.1",
        defaultTimezone: "UTC",
        userEmail: "user@example.invalid",
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
        userEmail: "user@example.invalid",
      }),
    );
    assert.equal(
      readPathFrom(node14_1, ["OrganizerEmail"]),
      "organizer@example.invalid",
    );
    assert.equal(readPathFrom(node14_1, ["OrganizerName"]), "test organizer");
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
      userEmail: "user@example.invalid",
    }),
  );

  assert.equal(readPathFrom(node, ["Recurrence", "Type"]), "1");
  assert.equal(readPathFrom(node, ["Recurrence", "Interval"]), "1");
  assert.equal(readPathFrom(node, ["Recurrence", "Until"]), "20260901T100000Z");
});

const SERIES_WITH_EXDATE = `BEGIN:VCALENDAR
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

const SERIES_WITH_OVERRIDE = SERIES_WITH_EXDATE.replace(
  "END:VCALENDAR",
  `BEGIN:VEVENT
UID:0400-outbound-series-uid
RECURRENCE-ID:20260815T100000Z
DTSTAMP:20260801T090000Z
DTSTART:20260815T140000Z
DTEND:20260815T143000Z
SUMMARY:test-outbound-series (one occurrence moved)
END:VEVENT
END:VCALENDAR`,
);

// The original PR's two cases here were characterization tests of the
// #334-era shape: InstanceId INSIDE ApplicationData, master <Change> and
// occurrence <Change> bundled under one ServerId in one request - the
// shape Exchange rejects with Status 6. That code is gone; the instance
// phase now emits ONE command per occurrence with InstanceId as a
// SIBLING of ServerId ([MS-ASCMD]), which these rewritten cases pin.

function emitted(command) {
  const w = createWBXML("AirSync");
  command.emit(w);
  return parseAdNode(decodeWBXML(w.getBytes()));
}

test("16.1: a cancelled occurrence becomes one <Delete> with InstanceId as a SIBLING of ServerId", () => {
  const commands = listInstanceCommands({
    blob: SERIES_WITH_EXDATE,
    serverID: "server-id-series",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    previous: { exdates: [], overrides: [] },
  });
  assert.equal(commands.length, 1);
  assert.equal(commands[0].kind, "delete");

  const node = emitted(commands[0]);
  assert.equal(node.tagName, "Delete");
  assert.equal(readPathFrom(node, ["ServerId"]), "server-id-series");
  assert.equal(readPathFrom(node, ["InstanceId"]), "20260808T100000Z");
  assert.equal(
    node.children.find((c) => c.tagName === "ApplicationData"),
    undefined,
    "a 16.x instance Delete carries no ApplicationData at all",
  );
});

test("16.1: a moved occurrence becomes one <Change> whose ApplicationData does NOT contain the InstanceId", () => {
  const commands = listInstanceCommands({
    blob: SERIES_WITH_OVERRIDE,
    serverID: "server-id-series",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    previous: { exdates: ["20260808T100000Z"], overrides: [] },
  });
  assert.equal(commands.length, 1, "the known EXDATE must not re-send");
  assert.equal(commands[0].kind, "change");

  const node = emitted(commands[0]);
  assert.equal(node.tagName, "Change");
  assert.equal(readPathFrom(node, ["ServerId"]), "server-id-series");
  // InstanceId is a sibling of ServerId - inside ApplicationData Exchange
  // rejects the command with Status 6, which was the #334-era shape.
  assert.equal(readPathFrom(node, ["InstanceId"]), "20260815T100000Z");
  const ad = node.children.find((c) => c.tagName === "ApplicationData");
  assert.equal(readPathFrom(ad, ["InstanceId"]), null);
  assert.equal(
    readPathFrom(ad, ["Subject"]),
    "test-outbound-series (one occurrence moved)",
  );
  assert.equal(readPathFrom(ad, ["StartTime"]), "20260815T140000Z");
  assert.equal(
    ad.children.find((c) => c.tagName === "Recurrence"),
    undefined,
    "an exception body re-emitting the RRULE would rewrite the series",
  );
  assert.equal(
    ad.children.find((c) => c.tagName === "TimeZone"),
    undefined,
    "a TimeZone element on an instance Change is rejected with Status 6",
  );
});
