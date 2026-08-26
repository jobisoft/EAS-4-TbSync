import { strict as assert } from "node:assert";
import { test } from "node:test";

import { applyResponses } from "../../src/modules/eas/sync-runner.mjs";

// The runner reads response nodes with `childByTag`, which only iterates
// `children` and reads `tagName` / `textContent`, so a plain object stands
// in for a parsed node without pretending to be a DOM element.
const node = (fields) => ({
  children: Object.entries(fields).map(([tagName, textContent]) => ({
    tagName,
    textContent,
    children: [],
  })),
});

const BLOB =
  "BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:evt-1\r\nSUMMARY:Standup\r\nEND:VEVENT\r\nEND:VCALENDAR";

function harness() {
  const log = [];
  const removed = [];
  const stored = [];
  const ctx = {
    accountId: "1",
    folderId: "f-11",
    targetID: "f-11",
    itemKind: {
      changelogKind: "event",
      codec: { stampEasServerId: (blob, id) => `${blob}\r\nSID:${id}` },
    },
    provider: { reportEventLog: (e) => log.push(e) },
    queue: {
      remove: async (r) => removed.push(r),
      markServerWrite: async () => {},
    },
    store: { update: async (id, blob) => stored.push({ id, blob }) },
    indexMap: { set: () => {} },
  };
  const sent = {
    adds: [
      {
        clientId: "c1",
        item: { id: "evt-1", blob: BLOB },
        entry: {
          parentId: "f-11",
          itemId: "evt-1",
          kind: "event",
          status: "added_by_user",
        },
      },
    ],
    mods: [],
    dels: [],
  };
  return { ctx, sent, log, removed, stored };
}

const conflict = node({ ClientId: "c1", Status: "7" });

test("a conflicted add is retired instead of retried", async () => {
  const { ctx, sent, removed } = harness();
  const failedItems = new Set();
  await applyResponses(
    ctx,
    { adds: [conflict], changes: [], deletes: [] },
    sent,
    failedItems,
  );
  assert.deepEqual(removed, [
    { parentId: "f-11", itemId: "evt-1", kind: "event" },
  ]);
});

test("a conflicted add stays out of failedItems, which re-stages", async () => {
  const { ctx, sent } = harness();
  const failedItems = new Set();
  await applyResponses(
    ctx,
    { adds: [conflict], changes: [], deletes: [] },
    sent,
    failedItems,
  );
  assert.equal(failedItems.size, 0);
});

test("a conflicted add is warned about, naming the item and what became of it", async () => {
  const { ctx, sent, log } = harness();
  await applyResponses(
    ctx,
    { adds: [conflict], changes: [], deletes: [] },
    sent,
    new Set(),
  );
  const warnings = log.filter((e) => e.level === "warning");
  assert.equal(warnings.length, 1);
  assert.match(warnings[0].message, /evt-1/);
  assert.match(warnings[0].message, /Status 7/);
  assert.match(warnings[0].message, /has been retired/);
  assert.equal(warnings[0].details, BLOB);
});

test("the item itself is left alone - only the queued push goes", async () => {
  const { ctx, sent, stored } = harness();
  await applyResponses(
    ctx,
    { adds: [conflict], changes: [], deletes: [] },
    sent,
    new Set(),
  );
  assert.deepEqual(stored, []);
});

test("any other add failure is still re-staged rather than retired", async () => {
  const { ctx, sent, removed } = harness();
  const failedItems = new Set();
  await applyResponses(
    ctx,
    { adds: [node({ ClientId: "c1", Status: "6" })], changes: [], deletes: [] },
    sent,
    failedItems,
  );
  assert.deepEqual(removed, []);
  assert.deepEqual([...failedItems], ["evt-1"]);
});

test("an accepted add is unaffected", async () => {
  const { ctx, sent, removed, stored } = harness();
  const failedItems = new Set();
  await applyResponses(
    ctx,
    {
      adds: [node({ ClientId: "c1", ServerId: "11:14", Status: "1" })],
      changes: [],
      deletes: [],
    },
    sent,
    failedItems,
  );
  assert.equal(failedItems.size, 0);
  assert.deepEqual(removed, [
    { parentId: "f-11", itemId: "evt-1", kind: "event" },
  ]);
  assert.equal(stored[0].blob, `${BLOB}\r\nSID:11:14`);
});
