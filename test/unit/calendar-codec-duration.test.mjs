/**
 * Item 30 — an event may express its end as DURATION instead of DTEND
 * (RFC 5545 allows either, never both). EAS always needs an EndTime, so
 * the writer derives end = DTSTART + DURATION; and an event whose end
 * cannot be derived at all - or does not come after its start - is held
 * by `clientRejectReason` instead of being pushed with invented data.
 * Only externally-authored items carry DURATION (our inbound writer
 * always stores DTEND), so fixtures here are import-shaped, not server
 * captures.
 *
 * The derivation shape follows thomcuddihy's PR #322 `eventTimingFor`;
 * the reject-instead-of-default policy is this repo's.
 */

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import ICAL from "../../src/vendor/ical.min.js";
import {
  appendApplicationDataFromIcal,
  applicationDataToIcal,
  clientRejectReason,
} from "../../src/modules/eas/calendar-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

before(() => ensureLoaded());

const vcal = (veventLines) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    "UID:duration-uid",
    "DTSTAMP:20260801T090000Z",
    ...veventLines,
    "SUMMARY:duration-event",
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");

/** Same wrapper as calendar-codec-outbound-edits.test.mjs — the real
 *  WBXML encoder and decoder, no mocks. */
function roundTripApplicationData(appendFn) {
  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendFn(w);
  w.switchpage("AirSync");
  w.ctag();
  const xml = decodeWBXML(w.getBytes());
  return parseAdNode(xml);
}

const emit = (ics, asVersion = "16.1") =>
  roundTripApplicationData((w) =>
    appendApplicationDataFromIcal({
      builder: w,
      ical: ics,
      asVersion,
      defaultTimezone: "UTC",
      syncRecurrence: true,
      userEmail: null,
    }),
  );

test("timed DTSTART + DURATION goes out with the derived EndTime", () => {
  const node = emit(vcal(["DTSTART:20260810T140000Z", "DURATION:PT1H30M"]));
  assert.equal(readPathFrom(node, ["StartTime"]), "20260810T140000Z");
  assert.equal(readPathFrom(node, ["EndTime"]), "20260810T153000Z");
  assert.equal(readPathFrom(node, ["AllDayEvent"]), "0");
});

test("DATE DTSTART + DURATION:P1D is an all-day event with date-shaped boundaries", () => {
  const node = emit(vcal(["DTSTART;VALUE=DATE:20260810", "DURATION:P1D"]));
  // Without the DURATION branch this was misclassified as timed
  // (isAllDayProp(null) is false), flipping the whole all-day encoding.
  assert.equal(readPathFrom(node, ["AllDayEvent"]), "1");
  assert.equal(readPathFrom(node, ["StartTime"]), "20260810T000000Z");
  assert.equal(readPathFrom(node, ["EndTime"]), "20260811T000000Z");
});

test("an inbound exception delta on a DURATION master seeds the override with a derived DTEND", async () => {
  // The stored local item expresses its length as DURATION (an import);
  // the server then sends an embedded Exception that changes only the
  // Subject. The override must come out with rid as DTSTART and rid plus
  // the master's DURATION as an explicit DTEND - and without the
  // master's DURATION beside it, which iCal forbids and which would
  // describe a length the DTEND already states.
  const existing = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    "UID:duration-series-uid",
    "DTSTAMP:20260801T090000Z",
    "DTSTART:20260810T140000Z",
    "DURATION:PT45M",
    "RRULE:FREQ=DAILY;COUNT=5",
    "SUMMARY:duration-series",
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");
  const after = await applicationDataToIcal({
    adNode: parseAdNode(`<ApplicationData>
      <Exceptions xmlns='Calendar'><Exception>
        <ExceptionStartTime>20260812T140000Z</ExceptionStartTime>
        <Subject>duration-series (renamed occurrence)</Subject>
      </Exception></Exceptions>
    </ApplicationData>`),
    existingIcal: existing,
    serverID: "srv-duration-series",
    asVersion: "14.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: null,
    userEmail: null,
  });
  const override = new ICAL.Component(ICAL.parse(after))
    .getAllSubcomponents("vevent")
    .find((v) => v.getFirstProperty("recurrence-id"));
  assert.ok(override, "expected an override VEVENT");
  assert.equal(
    override.getFirstPropertyValue("dtstart")?.toString(),
    "2026-08-12T14:00:00Z",
  );
  assert.equal(
    override.getFirstPropertyValue("dtend")?.toString(),
    "2026-08-12T14:45:00Z",
    "the override's end is rid plus the master's DURATION",
  );
  assert.equal(
    override.getFirstProperty("duration"),
    null,
    "the cloned master DURATION must not survive beside the new DTEND",
  );
});

test("an override carrying its own DURATION emits a derived EndTime in the embedded exception", () => {
  // Import-shaped: both master and override express length as DURATION.
  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    "UID:duration-series-uid",
    "DTSTAMP:20260801T090000Z",
    "DTSTART:20260810T140000Z",
    "DURATION:PT45M",
    "RRULE:FREQ=DAILY;COUNT=5",
    "SUMMARY:duration-series",
    "END:VEVENT",
    "BEGIN:VEVENT",
    "UID:duration-series-uid",
    "DTSTAMP:20260801T090000Z",
    "RECURRENCE-ID:20260812T140000Z",
    "DTSTART:20260812T160000Z",
    "DURATION:PT20M",
    "SUMMARY:duration-series (moved occurrence)",
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");
  const node = emit(ics, "14.1");
  assert.equal(readPathFrom(node, ["EndTime"]), "20260810T144500Z");
  const exceptions = node.children.find((c) => c.tagName === "Exceptions");
  assert.ok(exceptions, "expected an embedded <Exceptions> wrapper");
  const exception = exceptions.children.find((c) => c.tagName === "Exception");
  assert.equal(readPathFrom(exception, ["StartTime"]), "20260812T160000Z");
  assert.equal(
    readPathFrom(exception, ["EndTime"]),
    "20260812T162000Z",
    "the override's own DURATION derives its EndTime",
  );
});

test("clientRejectReason holds endless and non-positive events, passes representable ones", () => {
  const reject = (lines) =>
    clientRejectReason({ blob: vcal(lines), syncRecurrence: true });

  assert.match(
    String(reject(["DTSTART:20260810T140000Z"])),
    /neither DTEND nor DURATION/,
    "a start with no expressed end is held",
  );
  assert.match(
    String(reject(["DTSTART:20260810T140000Z", "DTEND:20260810T140000Z"])),
    /does not end after it starts/,
    "zero length is held",
  );
  assert.match(
    String(reject(["DTSTART:20260810T140000Z", "DTEND:20260810T130000Z"])),
    /does not end after it starts/,
    "a negative length is held",
  );
  assert.equal(
    reject(["DTSTART:20260810T140000Z", "DURATION:PT1H"]),
    null,
    "DTSTART+DURATION is representable - the writer derives the end",
  );
  assert.equal(
    reject(["DTSTART:20260810T140000Z", "DTEND:20260810T143000Z"]),
    null,
    "a normal event is untouched",
  );
  assert.equal(
    reject([]),
    null,
    "no DTSTART at all stays with the writer's backstop (exception-delta shape)",
  );
  assert.equal(
    reject(["DTSTART;VALUE=DATE:20260810"]),
    null,
    "a DATE start with no end is one day per RFC 5545 §3.6.1, not endless",
  );
});

test("a DATE start with no end at all is the RFC's one-day event on the wire", () => {
  // §3.6.1 gives this shape a defined duration, so reading it is not
  // inventing an end - holding it would strand ordinary imported
  // all-day items (a holiday .ics writes exactly this).
  const node = emit(vcal(["DTSTART;VALUE=DATE:20260810"]));
  assert.equal(readPathFrom(node, ["AllDayEvent"]), "1");
  assert.equal(readPathFrom(node, ["StartTime"]), "20260810T000000Z");
  assert.equal(readPathFrom(node, ["EndTime"]), "20260811T000000Z");
});

test("an override with a start but no derivable end is held, like a bad master", () => {
  // An occurrence rides the same writer as its master, so an override
  // the codec cannot time would reach the wire with EndTime = now().
  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    "UID:override-endless",
    "DTSTAMP:20260801T090000Z",
    "DTSTART:20260810T140000Z",
    "DTEND:20260810T144500Z",
    "RRULE:FREQ=DAILY;COUNT=5",
    "SUMMARY:series",
    "END:VEVENT",
    "BEGIN:VEVENT",
    "UID:override-endless",
    "DTSTAMP:20260801T090000Z",
    "RECURRENCE-ID:20260812T140000Z",
    "DTSTART:20260812T160000Z",
    "SUMMARY:moved occurrence with no end",
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");
  const reason = clientRejectReason({ blob: ics, syncRecurrence: true });
  assert.match(String(reason), /occurrence/, "the reason names the occurrence");
  assert.match(String(reason), /no end/);
});

test("inbound: a Change carrying EndTime drops a stale DURATION from the blob", async () => {
  const change = parseAdNode(`<ApplicationData>
    <StartTime xmlns='Calendar'>20260810T150000Z</StartTime>
    <EndTime xmlns='Calendar'>20260810T160000Z</EndTime>
  </ApplicationData>`);
  const after = await applicationDataToIcal({
    adNode: change,
    existingIcal: vcal(["DTSTART:20260810T140000Z", "DURATION:PT1H"]),
    serverID: "srv-duration",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: null,
    userEmail: null,
  });
  assert.ok(!/(^|\r?\n)DURATION[:;]/.test(after), "DURATION is gone");
  assert.match(after, /DTEND[:;][^\r\n]*20260810T160000Z/);
});
