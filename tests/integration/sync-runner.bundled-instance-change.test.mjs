// jobisoft#334's root cause, at the orchestration level: appendCommands
// (sync-runner.mjs) writes the master's own <Change> for a modified item,
// then - still inside the same <Commands> block, i.e. the same <Sync>
// request - calls codec.appendInstanceChanges to append one more <Change>
// per edited/deleted occurrence, addressed by the SAME ServerId as the
// master (16.1's InstanceId shape). calendar-codec.outbound-edits.test.mjs
// already proves appendInstanceChanges' own output is correct in
// isolation; what's untested until now is that sync-runner.mjs actually
// bundles it alongside the master Change rather than sending it
// separately - which is exactly the shape Exchange rejects live (both
// commands come back Status 6, "invalid item").
//
// This is a characterization test: it asserts today's actual behavior
// (the bundling happens, and a same-ServerId double-rejection leaves the
// item queued for retry rather than silently dropped or falsely cleared)
// so a future fix - splitting the occurrence Change into its own <Sync>
// request - is forced to touch this file rather than land unverified.

import { test, beforeAll } from "vitest";
import assert from "node:assert/strict";
import "../support/webext-shim.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "../support/xml-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";
import { realCalendarCodec } from "../support/real-calendar-codec.mjs";
import {
  appendCommands,
  applyResponses,
} from "../../src/modules/eas/sync-runner.mjs";

beforeAll(() => ensureLoaded());

// Master series plus one edited (not deleted) occurrence - the same
// fixture shape as calendar-codec.outbound-edits.test.mjs's
// appendInstanceChanges "overrides" test, so this test is provably
// exercising the same live-rejected wire shape, just one layer up.
const SERIES_WITH_OVERRIDE = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
UID:0400-bundled-series-uid
DTSTAMP:20260801T090000Z
DTSTART:20260801T100000Z
DTEND:20260801T103000Z
SUMMARY:test-bundled-series
RRULE:FREQ=WEEKLY;INTERVAL=1;UNTIL=20260901T100000Z;BYDAY=SA
END:VEVENT
BEGIN:VEVENT
UID:0400-bundled-series-uid
RECURRENCE-ID:20260808T100000Z
DTSTAMP:20260801T090000Z
DTSTART:20260808T140000Z
DTEND:20260808T143000Z
SUMMARY:test-bundled-series (one occurrence moved)
END:VEVENT
END:VCALENDAR
`;

test("appendCommands bundles the master Change and the occurrence's InstanceId Change under one ServerId in a single Commands block (the shape Exchange rejects)", () => {
  const mod = {
    entry: { itemId: "item-1", parentId: "calendar1" },
    serverID: "server-id-bundled",
    item: { id: "item-1", blob: SERIES_WITH_OVERRIDE },
  };

  const w = createWBXML("AirSync");
  appendCommands(w, {
    adds: [],
    mods: [mod],
    dels: [],
    separator: ",",
    asVersion: "16.1",
    codec: realCalendarCodec(),
    defaultTimezone: "UTC",
    syncRecurrence: true,
    userEmail: "kovacik@dgtfactory.com",
    fallbackOrganizerName: undefined,
    eventLog: () => {},
  });

  const commandsNode = parseAdNode(decodeWBXML(w.getBytes()));
  const changeNodes = commandsNode.children.filter(
    (c) => c.tagName === "Change",
  );

  // Both the master's own edit and the occurrence's InstanceId edit went
  // out as <Change> commands sharing the master's ServerId, in the same
  // request - the actual bundling jobisoft#334 is about.
  assert.equal(changeNodes.length, 2);
  assert.equal(readPathFrom(changeNodes[0], ["ServerId"]), "server-id-bundled");
  assert.equal(readPathFrom(changeNodes[1], ["ServerId"]), "server-id-bundled");
  const masterAd = changeNodes[0].children.find(
    (c) => c.tagName === "ApplicationData",
  );
  assert.equal(readPathFrom(masterAd, ["InstanceId"]), null);
  const occurrenceAd = changeNodes[1].children.find(
    (c) => c.tagName === "ApplicationData",
  );
  assert.equal(
    readPathFrom(occurrenceAd, ["InstanceId"]),
    "20260808T100000Z",
  );
});

test("applyResponses: when both bundled Changes come back Status 6 (live Exchange rejection), the item is queued for retry, not silently cleared", async () => {
  const changelogRemoveCalls = [];
  const provider = {
    changelogRemove: async (args) => changelogRemoveCalls.push(args),
    reportEventLog: () => {},
  };
  const ctx = { accountId: "acc1", folderId: "folder1", provider };

  const mod = {
    entry: { itemId: "item-1", parentId: "calendar1" },
    serverID: "server-id-bundled",
    item: { id: "item-1", blob: SERIES_WITH_OVERRIDE },
  };
  const sent = { adds: [], mods: [mod], dels: [] };

  // The server rejects each of the two bundled Change commands with its
  // own Status 6 response node, both stamped with the shared ServerId -
  // exactly what was captured live during the jobisoft#334 investigation.
  const responses = {
    adds: [],
    changes: [
      parseAdNode(
        `<Change><ServerId>server-id-bundled</ServerId><Status>6</Status></Change>`,
      ),
      parseAdNode(
        `<Change><ServerId>server-id-bundled</ServerId><Status>6</Status></Change>`,
      ),
    ],
    deletes: [],
  };

  const failedItems = new Set();
  await applyResponses(ctx, responses, sent, failedItems, {
    hadResponsesElement: true,
  });

  // pushPhase's tail re-stage relies on this: the item is retried next
  // sync rather than dropped.
  assert.ok(failedItems.has("item-1"));
  // The changelog entry must NOT be cleared on a real failure - clearing
  // it here would silently drop the user's edit (both the reschedule and
  // the moved occurrence) with no further retry, which is worse than the
  // live symptom (a stuck retry loop) this bundling bug actually causes.
  assert.equal(changelogRemoveCalls.length, 0);
});
