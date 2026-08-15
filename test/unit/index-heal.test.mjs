/**
 * The daily reconcile of a folder's identity map.
 *
 * Two things it must get right, and both are asymmetric: a mapping wrongly
 * kept is inert almost always, while a mapping wrongly dropped can send a
 * queued delete to the wrong item on the server - or, if enough of them go,
 * duplicate the folder on the next re-download.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  healFolderIndex,
  healIsDue,
  HEAL_INTERVAL_MS,
} from "../../src/modules/eas/index-heal.mjs";

const stamp = (blob) => {
  if (blob === "unparsable") throw new Error("no");
  const m = /SERVERID:(\S+)/.exec(blob);
  return m ? m[1] : null;
};
const store = (items) => ({ list: async () => items });

test("it is due when never run, and once a day after", () => {
  const now = 1_000_000_000_000;
  assert.equal(healIsDue({}, now), true, "a folder that has never healed");
  assert.equal(healIsDue({ custom: { indexHealed: now } }, now), false);
  assert.equal(
    healIsDue({ custom: { indexHealed: now - HEAL_INTERVAL_MS } }, now),
    true,
  );
  // A stamp that is not a number says nothing, so it counts as never.
  assert.equal(healIsDue({ custom: { indexHealed: "soon" } }, now), true);
});

test("mappings the items carry are restored", async () => {
  const out = await healFolderIndex({
    store: store([
      { id: "a", blob: "SERVERID:1" },
      { id: "b", blob: "SERVERID:2" },
    ]),
    readStamp: stamp,
    stored: [],
    queuedIds: new Set(),
  });
  assert.deepEqual(out.entries.sort((x, y) => x.uid < y.uid ? -1 : 1), [
    { uid: "a", serverId: "1" },
    { uid: "b", serverId: "2" },
  ]);
});

test("an entry naming nothing is dropped", async () => {
  const out = await healFolderIndex({
    store: store([{ id: "a", blob: "SERVERID:1" }]),
    readStamp: stamp,
    stored: [
      { uid: "a", serverId: "1" },
      { uid: "ghost", serverId: "9" },
    ],
    queuedIds: new Set(),
  });
  assert.equal(out.pruned, 1);
  assert.deepEqual(out.entries, [{ uid: "a", serverId: "1" }]);
});

test("but not one the user has a delete queued for", async () => {
  // The item is gone locally *because* the user deleted it; the mapping is
  // what addresses that delete when the push finally goes, and
  // `hasPendingUserDelete` reads it to tell an ordinary crossing apart from
  // real drift.
  const out = await healFolderIndex({
    store: store([]),
    readStamp: stamp,
    stored: [{ uid: "deleted-by-user", serverId: "7" }],
    queuedIds: new Set(["deleted-by-user"]),
  });
  assert.equal(out.pruned, 0);
  assert.deepEqual(out.entries, [{ uid: "deleted-by-user", serverId: "7" }]);
});

test("an item whose blob will not parse keeps its mapping", async () => {
  // The fill cannot read it, but the item is there and the entry is true.
  // Pruning on "no stamp found" instead of "no item" would lose it.
  const out = await healFolderIndex({
    store: store([{ id: "a", blob: "unparsable" }]),
    readStamp: stamp,
    stored: [{ uid: "a", serverId: "1" }],
    queuedIds: new Set(),
  });
  assert.equal(out.pruned, 0);
  assert.deepEqual(out.entries, [{ uid: "a", serverId: "1" }]);
});

test("a stale entry cannot outlive the id another item now holds", async () => {
  // The case that is not merely untidy: `serverIdFor("ghost")` would answer
  // "1", and a delete queued for the ghost would remove item a's copy on
  // the server.
  const out = await healFolderIndex({
    store: store([{ id: "a", blob: "SERVERID:1" }]),
    readStamp: stamp,
    stored: [{ uid: "ghost", serverId: "1" }],
    queuedIds: new Set(),
  });
  assert.deepEqual(out.entries, [{ uid: "a", serverId: "1" }]);
});

test("a folder that cannot be read changes nothing at all", async () => {
  // Not the same as an empty folder. Reading the emptiness as truth would
  // prune every mapping the folder has.
  const out = await healFolderIndex({
    store: {
      list: async () => {
        throw new Error("store unavailable");
      },
    },
    readStamp: stamp,
    stored: [{ uid: "a", serverId: "1" }],
    queuedIds: new Set(),
  });
  assert.equal(out, null);
});

test("a locally created item that was never pushed is left alone", async () => {
  // No stamp, so no mapping - and none is invented for it.
  const out = await healFolderIndex({
    store: store([{ id: "fresh", blob: "no stamp yet" }]),
    readStamp: stamp,
    stored: [],
    queuedIds: new Set(["fresh"]),
  });
  assert.deepEqual(out.entries, []);
  assert.equal(out.pruned, 0);
});
