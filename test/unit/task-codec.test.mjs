/**
 * Ported from PR #345 (tomaskovacik) to the node:test layer; fixtures
 * kept verbatim (several are live-server captures), expectations
 * re-verified against current master.
 */

// task-codec.mjs: EAS Tasks (codepage 9) <-> iCal VTODO. Same shape as
// calendar-codec.mjs's applicationDataToIcal/appendApplicationDataFromIcal
// but for VTODO - not covered by any calendar-codec test, and not
// touched at all until now.

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import ICAL from "../../src/vendor/ical.min.js";
import {
  applicationDataToIcal,
  appendApplicationDataFromIcal,
  readEasServerIdFromIcal,
  stampEasServerId,
} from "../../src/modules/eas/task-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

before(() => ensureLoaded());

const ADD_TASK = `<ApplicationData>
  <Subject xmlns='Tasks'>test-task</Subject>
  <Body xmlns='AirSyncBase'><Type xmlns='AirSyncBase'>1</Type><Data xmlns='AirSyncBase'>a%20task%20body</Data></Body>
  <Importance xmlns='Tasks'>2</Importance>
  <Sensitivity xmlns='Tasks'>2</Sensitivity>
  <UtcStartDate xmlns='Tasks'>2026-08-01T10:00:00.000Z</UtcStartDate>
  <StartDate xmlns='Tasks'>2026-08-01T10:00:00.000Z</StartDate>
  <UtcDueDate xmlns='Tasks'>2026-08-03T10:00:00.000Z</UtcDueDate>
  <DueDate xmlns='Tasks'>2026-08-03T10:00:00.000Z</DueDate>
  <Complete xmlns='Tasks'>0</Complete>
  <ReminderSet xmlns='Tasks'>1</ReminderSet>
  <ReminderTime xmlns='Tasks'>2026-08-01T09:00:00.000Z</ReminderTime>
  <Categories xmlns='Tasks'>
    <Category xmlns='Tasks'>Work</Category>
  </Categories>
</ApplicationData>`;

function firstVtodo(icalString) {
  return new ICAL.Component(ICAL.parse(icalString)).getFirstSubcomponent(
    "vtodo",
  );
}

test("applicationDataToIcal: maps Subject/Body/Importance/Sensitivity/Start/Due/Categories", async () => {
  const ical = await applicationDataToIcal({
    adNode: parseAdNode(ADD_TASK),
    existingIcal: null,
    serverID: "server-id-task-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: false,
    msTodoCompat: false,
    uid: null,
  });

  const vtodo = firstVtodo(ical);
  assert.equal(vtodo.getFirstPropertyValue("summary"), "test-task");
  assert.equal(vtodo.getFirstPropertyValue("description"), "a task body");
  assert.equal(vtodo.getFirstPropertyValue("priority"), 1); // Importance 2 -> PRIORITY 1
  assert.equal(vtodo.getFirstPropertyValue("class"), "PRIVATE"); // Sensitivity 2
  assert.equal(
    vtodo.getFirstPropertyValue("dtstart").toString(),
    "2026-08-01T10:00:00Z",
  );
  assert.equal(
    vtodo.getFirstPropertyValue("due").toString(),
    "2026-08-03T10:00:00Z",
  );
  assert.deepEqual(vtodo.getFirstProperty("categories").getValues(), ["Work"]);
  assert.equal(readEasServerIdFromIcal(ical), "server-id-task-1");
});

test("applicationDataToIcal: Complete=1 sets STATUS/PERCENT-COMPLETE/COMPLETED", async () => {
  const completedAd = ADD_TASK.replace(
    "<Complete xmlns='Tasks'>0</Complete>",
    "<Complete xmlns='Tasks'>1</Complete><DateCompleted xmlns='Tasks'>2026-08-02T10:00:00.000Z</DateCompleted>",
  );
  const ical = await applicationDataToIcal({
    adNode: parseAdNode(completedAd),
    existingIcal: null,
    serverID: "server-id-task-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: false,
    msTodoCompat: false,
    uid: null,
  });
  const vtodo = firstVtodo(ical);
  assert.equal(vtodo.getFirstPropertyValue("status"), "COMPLETED");
  assert.equal(vtodo.getFirstPropertyValue("percent-complete"), 100);
  assert.equal(
    vtodo.getFirstPropertyValue("completed").toString(),
    "2026-08-02T10:00:00Z",
  );
});

test("applicationDataToIcal: Complete is merge-aware - clears a prior COMPLETED state when a later delta says Complete=0", async () => {
  const commonArgs = {
    serverID: "server-id-task-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: false,
    msTodoCompat: false,
    uid: null,
  };
  const completedAd = ADD_TASK.replace(
    "<Complete xmlns='Tasks'>0</Complete>",
    "<Complete xmlns='Tasks'>1</Complete><DateCompleted xmlns='Tasks'>2026-08-02T10:00:00.000Z</DateCompleted>",
  );
  const afterComplete = await applicationDataToIcal({
    adNode: parseAdNode(completedAd),
    existingIcal: null,
    ...commonArgs,
  });

  const reopened = await applicationDataToIcal({
    adNode: parseAdNode(
      `<ApplicationData><Complete xmlns='Tasks'>0</Complete></ApplicationData>`,
    ),
    existingIcal: afterComplete,
    ...commonArgs,
  });
  const vtodo = firstVtodo(reopened);
  assert.equal(vtodo.getFirstPropertyValue("status"), null);
  assert.equal(vtodo.getFirstPropertyValue("percent-complete"), null);
  assert.equal(vtodo.getFirstPropertyValue("completed"), null);
});

test("applicationDataToIcal: a delta without <Complete> leaves an existing completion state untouched", async () => {
  const commonArgs = {
    serverID: "server-id-task-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: false,
    msTodoCompat: false,
    uid: null,
  };
  const completedAd = ADD_TASK.replace(
    "<Complete xmlns='Tasks'>0</Complete>",
    "<Complete xmlns='Tasks'>1</Complete>",
  );
  const afterComplete = await applicationDataToIcal({
    adNode: parseAdNode(completedAd),
    existingIcal: null,
    ...commonArgs,
  });

  const afterSubjectEdit = await applicationDataToIcal({
    adNode: parseAdNode(
      `<ApplicationData><Subject xmlns='Tasks'>renamed</Subject></ApplicationData>`,
    ),
    existingIcal: afterComplete,
    ...commonArgs,
  });
  const vtodo = firstVtodo(afterSubjectEdit);
  assert.equal(vtodo.getFirstPropertyValue("summary"), "renamed");
  assert.equal(vtodo.getFirstPropertyValue("status"), "COMPLETED");
});

test("stampEasServerId / readEasServerIdFromIcal round-trip", async () => {
  const ical = await applicationDataToIcal({
    adNode: parseAdNode(ADD_TASK),
    existingIcal: null,
    serverID: "original-id",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: false,
    msTodoCompat: false,
    uid: null,
  });
  const restamped = stampEasServerId(ical, "new-id");
  assert.equal(readEasServerIdFromIcal(restamped), "new-id");
});

test("appendApplicationDataFromIcal: outbound round-trip via the real WBXML encoder/decoder", async () => {
  const ical = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VTODO
UID:0400-task-outbound-uid
SUMMARY:test-task-outbound
PRIORITY:1
STATUS:COMPLETED
PERCENT-COMPLETE:100
DTSTART:20260801T100000Z
DUE:20260803T100000Z
COMPLETED:20260802T100000Z
CATEGORIES:Work,Home
END:VTODO
END:VCALENDAR
`;

  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendApplicationDataFromIcal({
    builder: w,
    ical,
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: false,
  });
  w.switchpage("AirSync");
  w.ctag();
  const node = parseAdNode(decodeWBXML(w.getBytes()));

  assert.equal(readPathFrom(node, ["Subject"]), "test-task-outbound");
  assert.equal(readPathFrom(node, ["Importance"]), "2"); // PRIORITY 1 -> Importance 2
  assert.equal(readPathFrom(node, ["Complete"]), "1");
  assert.equal(
    readPathFrom(node, ["UtcStartDate"]),
    "2026-08-01T10:00:00.000Z",
  );
  assert.equal(readPathFrom(node, ["UtcDueDate"]), "2026-08-03T10:00:00.000Z");
  assert.deepEqual(
    node.children
      .find((c) => c.tagName === "Categories")
      ?.children.filter((c) => c.tagName === "Category")
      .map((c) => c.textContent),
    ["Work", "Home"],
  );
});
