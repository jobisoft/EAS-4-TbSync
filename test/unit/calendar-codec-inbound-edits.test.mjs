/**
 * Ported from PR #345 (tomaskovacik) to the node:test layer; fixtures
 * kept verbatim (several are live-server captures), expectations
 * re-verified against current master.
 */

// Inbound EAS -> iCal: edit/reschedule a single event, edit a whole
// recurring series, and edit/delete a single occurrence within a series
// (the 16.1 standalone per-instance <Change> path - applyInstanceChange/
// applyInstanceDelete - as opposed to the embedded <Exceptions> path
// already covered in calendar-codec.test.mjs).

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import ICAL from "../../src/vendor/ical.min.js";
import {
  applicationDataToIcal,
  applyInstanceChange,
  applyInstanceDelete,
} from "../../src/modules/eas/calendar-codec.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";

before(() => ensureLoaded());

const COMMON = {
  serverID: "server-id-1",
  asVersion: "16.1",
  defaultTimezone: "UTC",
  uid: null,
  userEmail: "user@example.invalid",
};

const ADD_SINGLE_EVENT = `<ApplicationData>
  <AllDayEvent xmlns='Calendar'>0</AllDayEvent>
  <DtStamp xmlns='Calendar'>20260801T090000Z</DtStamp>
  <StartTime xmlns='Calendar'>20260801T100000Z</StartTime>
  <Subject xmlns='Calendar'>test-single</Subject>
  <UID xmlns='Calendar'>0400-single-uid</UID>
  <EndTime xmlns='Calendar'>20260801T103000Z</EndTime>
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

test("single event: a Change reschedules Start/End and edits Subject", async () => {
  const afterAdd = await applicationDataToIcal({
    adNode: parseAdNode(ADD_SINGLE_EVENT),
    existingIcal: null,
    ...COMMON,
  });

  const reschedule = parseAdNode(`<ApplicationData>
    <StartTime xmlns='Calendar'>20260801T140000Z</StartTime>
    <EndTime xmlns='Calendar'>20260801T143000Z</EndTime>
    <Subject xmlns='Calendar'>test-single (moved)</Subject>
  </ApplicationData>`);

  const afterReschedule = await applicationDataToIcal({
    adNode: reschedule,
    existingIcal: afterAdd,
    ...COMMON,
  });

  const vevent = masterVevent(afterReschedule);
  assert.equal(vevent.getFirstPropertyValue("summary"), "test-single (moved)");
  assert.equal(
    vevent.getFirstPropertyValue("dtstart").toString(),
    "2026-08-01T14:00:00Z",
  );
  assert.equal(
    vevent.getFirstPropertyValue("dtend").toString(),
    "2026-08-01T14:30:00Z",
  );
});

const ADD_RECURRING_SERIES = `<ApplicationData>
  <AllDayEvent xmlns='Calendar'>0</AllDayEvent>
  <DtStamp xmlns='Calendar'>20260801T090000Z</DtStamp>
  <StartTime xmlns='Calendar'>20260801T100000Z</StartTime>
  <Subject xmlns='Calendar'>test-series</Subject>
  <UID xmlns='Calendar'>0400-series-uid</UID>
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

test("whole series: a Change with a new Recurrence.Until shortens the RRULE", async () => {
  const afterAdd = await applicationDataToIcal({
    adNode: parseAdNode(ADD_RECURRING_SERIES),
    existingIcal: null,
    ...COMMON,
  });

  const shortened = parseAdNode(`<ApplicationData>
    <Recurrence xmlns='Calendar'>
      <Type xmlns='Calendar'>1</Type>
      <Interval xmlns='Calendar'>1</Interval>
      <Until xmlns='Calendar'>20260901T100000Z</Until>
      <DayOfWeek xmlns='Calendar'>32</DayOfWeek>
      <FirstDayOfWeek xmlns='Calendar'>0</FirstDayOfWeek>
    </Recurrence>
  </ApplicationData>`);

  const afterChange = await applicationDataToIcal({
    adNode: shortened,
    existingIcal: afterAdd,
    ...COMMON,
  });

  const vevent = masterVevent(afterChange);
  const rrule = vevent.getFirstPropertyValue("rrule");
  assert.ok(rrule, "expected an RRULE");
  assert.equal(rrule.until.toString(), "2026-09-01T10:00:00Z");
});

function overrideByRecurrenceId(icalString, isoDate) {
  const vcal = new ICAL.Component(ICAL.parse(icalString));
  return vcal
    .getAllSubcomponents("vevent")
    .find(
      (v) => v.getFirstPropertyValue("recurrence-id")?.toString() === isoDate,
    );
}

test("single occurrence (16.1 standalone Change): applyInstanceChange reschedules just that instance", async () => {
  const afterAdd = await applicationDataToIcal({
    adNode: parseAdNode(ADD_RECURRING_SERIES),
    existingIcal: null,
    ...COMMON,
  });

  // Second occurrence of the weekly series: 2026-08-08T10:00:00Z.
  const instanceUtc = new Date("2026-08-08T10:00:00Z");
  const occurrenceEdit = parseAdNode(`<ApplicationData>
    <StartTime xmlns='Calendar'>20260808T130000Z</StartTime>
    <EndTime xmlns='Calendar'>20260808T133000Z</EndTime>
    <Subject xmlns='Calendar'>test-series (one occurrence moved)</Subject>
  </ApplicationData>`);

  const afterInstanceChange = await applyInstanceChange({
    ical: afterAdd,
    adNode: occurrenceEdit,
    instanceUtc,
    asVersion: "16.1",
    defaultTimezone: "UTC",
    userEmail: "user@example.invalid",
  });

  const master = masterVevent(afterInstanceChange);
  assert.equal(
    master.getFirstPropertyValue("summary"),
    "test-series",
    "the master's own Subject must be untouched by an occurrence edit",
  );

  const override = overrideByRecurrenceId(
    afterInstanceChange,
    "2026-08-08T10:00:00Z",
  );
  assert.ok(override, "expected an override VEVENT for the edited occurrence");
  assert.equal(
    override.getFirstPropertyValue("summary"),
    "test-series (one occurrence moved)",
  );
  assert.equal(
    override.getFirstPropertyValue("dtstart").toString(),
    "2026-08-08T13:00:00Z",
  );
});

test("single occurrence (16.1 standalone Change): applyInstanceDelete adds an EXDATE and drops any override", async () => {
  const afterAdd = await applicationDataToIcal({
    adNode: parseAdNode(ADD_RECURRING_SERIES),
    existingIcal: null,
    ...COMMON,
  });

  const instanceUtc = new Date("2026-08-08T10:00:00Z");
  const afterDelete = applyInstanceDelete({ ical: afterAdd, instanceUtc });

  const master = masterVevent(afterDelete);
  const exdates = master
    .getAllProperties("exdate")
    .map((p) => p.getFirstValue().toString());
  assert.deepEqual(exdates, ["2026-08-08T10:00:00Z"]);

  assert.equal(
    overrideByRecurrenceId(afterDelete, "2026-08-08T10:00:00Z"),
    undefined,
    "a deleted instance must not also leave a stale override behind",
  );
});
