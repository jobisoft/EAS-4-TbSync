/**
 * The folder's uid <-> serverId map.
 *
 * Everything here is about one question: after some sequence of events,
 * does a server id still resolve to the item it names? A wrong answer is
 * not a wrong answer - it is a duplicate item, or somebody else's item
 * deleted.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { createServerIdIndex } from "../../src/modules/eas/server-id-index.mjs";

test("it answers in both directions", () => {
  const ix = createServerIdIndex([{ uid: "a", serverId: "1" }]);
  assert.equal(ix.uidFor("1"), "a");
  assert.equal(ix.serverIdFor("a"), "1");
});

test("an id nobody claims resolves to nothing, not to undefined", () => {
  const ix = createServerIdIndex();
  assert.equal(ix.uidFor("nope"), null);
  assert.equal(ix.serverIdFor("nope"), null);
  // The callers branch on falsiness, but a null says "asked and answered"
  // where undefined reads like a missing implementation.
  assert.equal(ix.uidFor(null), null);
  assert.equal(ix.serverIdFor(undefined), null);
});

test("re-stamping an item retires the id it replaced", () => {
  // The whole point. Answering an invitation makes the server re-file the
  // meeting: an <Add> under a new id arrives with a <Delete> for the old
  // one in the same response. If the superseded id still resolved, that
  // delete would take the item the add had just landed on.
  const ix = createServerIdIndex([{ uid: "a", serverId: "1" }]);
  ix.set("a", "2");
  assert.equal(ix.uidFor("2"), "a");
  assert.equal(ix.uidFor("1"), null, "the superseded id still resolves");
  assert.equal(ix.serverIdFor("a"), "2");
});

test("removing an item takes both directions with it", () => {
  const ix = createServerIdIndex([{ uid: "a", serverId: "1" }]);
  ix.remove("a");
  assert.equal(ix.uidFor("1"), null);
  assert.equal(ix.serverIdFor("a"), null);
  assert.equal(ix.size, 0);
});

test("two items claiming one server id: the last claim answers", () => {
  const ix = createServerIdIndex();
  ix.set("a", "1");
  ix.set("b", "1");
  assert.equal(ix.uidFor("1"), "b");
  // And the loser keeps its own entry, because that is what addresses a
  // delete still queued for it. Dropping it would silently strand that
  // delete on the server.
  assert.equal(ix.serverIdFor("a"), "1");
});

test("the loser of a contested id cannot take the winner's entry with it", () => {
  const ix = createServerIdIndex();
  ix.set("a", "1");
  ix.set("b", "1");
  ix.remove("a");
  assert.equal(ix.uidFor("1"), "b", "removing the loser unmapped the winner");
});

test("what is stored round-trips, and loading it is not a change", () => {
  const stored = [
    { uid: "a", serverId: "1" },
    { uid: "b", serverId: "2" },
  ];
  const ix = createServerIdIndex(stored);
  assert.equal(ix.dirty, false, "a freshly loaded map wants writing back");
  assert.deepEqual(ix.toArray(), stored);
});

test("only a real change marks it for writing back", () => {
  const ix = createServerIdIndex([{ uid: "a", serverId: "1" }]);
  ix.set("a", "1");
  assert.equal(ix.dirty, false, "re-stating what it already held");
  ix.remove("nobody");
  assert.equal(ix.dirty, false, "removing what was never there");
  ix.set("a", "2");
  assert.equal(ix.dirty, true);
});

test("a stored shape it cannot read leaves it empty rather than broken", () => {
  // Nothing but this module writes the value, so these are defensive - but
  // an exception here fails the folder's whole sync.
  for (const junk of [null, undefined, "", 0, {}, "[]"]) {
    assert.equal(createServerIdIndex(junk).size, 0, `for ${String(junk)}`);
  }
  const half = createServerIdIndex([
    { uid: "a" },
    { serverId: "2" },
    null,
    { uid: "c", serverId: "3" },
  ]);
  assert.equal(half.size, 1, "a pair missing half of itself was kept");
  assert.equal(half.uidFor("3"), "c");
});

test("it stays correct over a resync's worth of traffic", () => {
  // A folder is pulled, half of it re-stamped, some of it deleted - then
  // every id is asked for again. Each of these has its own test above; the
  // point of this one is that they compose.
  const ix = createServerIdIndex();
  for (let i = 0; i < 200; i++) ix.set(`uid-${i}`, `srv-${i}`);
  for (let i = 0; i < 200; i += 2) ix.set(`uid-${i}`, `srv-new-${i}`);
  for (let i = 0; i < 200; i += 5) ix.remove(`uid-${i}`);

  for (let i = 0; i < 200; i++) {
    const removed = i % 5 === 0;
    const restamped = i % 2 === 0;
    const live = restamped ? `srv-new-${i}` : `srv-${i}`;
    assert.equal(ix.serverIdFor(`uid-${i}`), removed ? null : live);
    assert.equal(ix.uidFor(live), removed ? null : `uid-${i}`);
    if (restamped) {
      assert.equal(ix.uidFor(`srv-${i}`), null, `superseded srv-${i} resolved`);
    }
  }
  assert.equal(ix.size, 160);
  assert.equal(ix.toArray().length, 160);
});

test("a lost index is rebuilt from the stamps in the stored items", () => {
  // The blobs carry the same fact, item by item, and are written as the pass
  // goes; the index is written once at the end. So a crash, a failed flush,
  // or anything clearing the stored index leaves the blobs ahead of it - and
  // an EAS task or contact has no UID on the wire, so nothing else could
  // recognise the item when the server re-sends it.
  const ix = createServerIdIndex();
  ix.fill(
    [
      { id: "a", blob: "SERVERID:1" },
      { id: "b", blob: "SERVERID:2" },
      { id: "local-only", blob: "no stamp here" },
      { id: "broken", blob: "boom" },
    ],
    (blob) => {
      if (blob === "boom") throw new Error("unparsable");
      const m = /SERVERID:(\d+)/.exec(blob);
      return m ? m[1] : null;
    },
  );
  assert.equal(ix.uidFor("1"), "a");
  assert.equal(ix.uidFor("2"), "b");
  assert.equal(ix.size, 2, "an unstamped or unreadable item claimed an id");
});

test("the rebuild lets the blob correct a stale mapping", () => {
  const ix = createServerIdIndex([{ uid: "a", serverId: "old" }]);
  ix.fill([{ id: "a", blob: "SERVERID:new" }], (b) => /SERVERID:(\S+)/.exec(b)[1]);
  assert.equal(ix.serverIdFor("a"), "new");
  assert.equal(ix.uidFor("old"), null, "the superseded id still resolves");
});

test("duplicates in the store cannot steal each other's id", () => {
  const ix = createServerIdIndex();
  ix.fill(
    [
      { id: "first", blob: "SERVERID:1" },
      { id: "second", blob: "SERVERID:1" },
    ],
    (b) => /SERVERID:(\S+)/.exec(b)[1],
  );
  assert.equal(ix.uidFor("1"), "first", "first writer must win, for stability");
});
