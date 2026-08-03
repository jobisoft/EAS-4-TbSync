// MS-ASCMD allows a Sync push reply to come back Status 1 at the
// collection level with no per-command <Responses> element at all - a
// terser "everything you sent was accepted" reply some servers send
// instead of ACKing each command individually. sendSync/parseSyncResponse
// turns that into `responses: null`; pushPhase then falls back to
// `{ adds: [], changes: [], deletes: [] }` and tells applyResponses via
// `hadResponsesElement: false`.
//
// This is the exact mechanism behind a delete that looked successful
// live when it wasn't (see TEST-PLAN.md / project memory): with no
// per-command Responses to match against, applyResponses' fallback loops
// treat every sent mod and delete as an implicit success (no ack found ==
// no ack needed), clearing their changelog entries unconditionally. Sent
// *adds* are the odd one out: an Add can only be resolved by reading back
// the ServerId the server assigned it, and with no <Responses> element
// there's no ServerId anywhere - so adds are silently left pending
// instead, unlike mods/dels.
//
// A characterization test: this documents today's actual (asymmetric,
// optimistic-for-mods/dels) behavior so a future change to it is
// deliberate, not accidental.

import { test } from "vitest";
import assert from "node:assert/strict";
import { applyResponses } from "../../src/modules/eas/sync-runner.mjs";

test("applyResponses with no <Responses> element: mods and deletes are treated as silently successful, adds are left pending", async () => {
  const changelogRemoveCalls = [];
  const eventLogs = [];
  const provider = {
    changelogRemove: async (args) => changelogRemoveCalls.push(args),
    reportEventLog: (args) => eventLogs.push(args),
  };
  const indexMap = [
    { uid: "del-item", serverId: "server-id-del" },
    { uid: "mod-item", serverId: "server-id-mod" },
  ];
  const ctx = { accountId: "acc1", folderId: "folder1", provider, indexMap };

  const sent = {
    adds: [
      {
        entry: { itemId: "add-item", parentId: "calendar1" },
        clientId: "c-1",
        item: { id: "add-item", blob: "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n" },
      },
    ],
    mods: [
      {
        entry: { itemId: "mod-item", parentId: "calendar1" },
        serverID: "server-id-mod",
        item: { id: "mod-item", blob: "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n" },
      },
    ],
    dels: [
      {
        entry: { itemId: "del-item", parentId: "calendar1" },
        serverID: "server-id-del",
      },
    ],
  };

  const failedItems = new Set();
  await applyResponses(
    ctx,
    { adds: [], changes: [], deletes: [] },
    sent,
    failedItems,
    { hadResponsesElement: false },
  );

  // The mod and the delete both get their changelog entries cleared with
  // no server confirmation at all - the false-positive-success fallback.
  assert.deepEqual(
    changelogRemoveCalls.map((c) => c.itemId).sort(),
    ["del-item", "mod-item"],
  );
  // The delete's indexMap entry is removed right along with it.
  assert.equal(
    indexMap.find((e) => e.uid === "del-item"),
    undefined,
  );

  // The add is NOT resolved - no ServerId to stamp it with, so it's left
  // as-is (still pending in the changelog, unstamped) rather than guessed
  // at. It will be re-sent as an Add again on the next push.
  assert.equal(
    changelogRemoveCalls.some((c) => c.itemId === "add-item"),
    false,
  );

  // Nothing was marked failed either - the fallback path has no notion
  // of per-item failure when there's nothing to compare against.
  assert.equal(failedItems.size, 0);

  // The debug log documenting the fallback fired, counting all 3 sent
  // items (1 add + 1 mod + 1 del) even though only 2 were actually
  // resolved.
  assert.equal(eventLogs.length, 1);
  assert.equal(eventLogs[0].level, "debug");
  assert.match(eventLogs[0].message, /no <Responses>.*clearing 3 sent items/);
});
