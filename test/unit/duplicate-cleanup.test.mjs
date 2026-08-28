/**
 * Unit tests for removing the surplus copies of a duplicated UID.
 *
 * This is the one path in the add-on that deletes from a live mailbox
 * without a user edit behind it, so what it puts on the wire and what it
 * does with the answer are both pinned here.
 *
 * The SyncKey is the part worth the most attention. A `Sync` carrying
 * commands consumes it and hands back the next one; a cleanup that lost
 * the new key mid-run would leave the folder presenting a key the server
 * has moved past, which costs a full resync. Hence it is persisted after
 * every chunk, and before the per-item statuses are even read.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();

const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);

const { decodeWBXML } = await import("../../src/modules/wbxml.mjs");

/** Requests the stub was given, decoded to text. */
const sent = [];
/** Replies to hand back, in order; the last one repeats. */
let replies = [];

const real = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...real,
    easRequest: async ({ body }) => {
      sent.push(decodeWBXML(body));
      const xml = replies[Math.min(sent.length - 1, replies.length - 1)];
      return { doc: { documentElement: parseAdNode(xml) } };
    },
  },
});
const { deleteSurplusCopies } =
  await import("../../src/modules/eas/duplicate-cleanup.mjs");

/** A Sync reply that accepts everything and moves the key on. */
const accepted = (synckey) => `<?xml version="1.0" encoding="utf-8"?>
  <Sync xmlns="AirSync">
    <Collections><Collection>
      <SyncKey>${synckey}</SyncKey>
      <CollectionId>42</CollectionId>
      <Status>1</Status>
    </Collection></Collections>
  </Sync>`;

/** A Sync reply that reports on the ids named. */
const answering = (synckey, statuses) => `<?xml version="1.0" encoding="utf-8"?>
  <Sync xmlns="AirSync">
    <Collections><Collection>
      <SyncKey>${synckey}</SyncKey>
      <CollectionId>42</CollectionId>
      <Status>1</Status>
      <Responses>
        ${Object.entries(statuses)
          .map(
            ([id, status]) =>
              `<Delete><ServerId>${id}</ServerId><Status>${status}</Status></Delete>`,
          )
          .join("")}
      </Responses>
    </Collection></Collections>
  </Sync>`;

const persisted = [];

function run(serverIds, { synckey = "key-0", ...rest } = {}) {
  sent.length = 0;
  persisted.length = 0;
  return deleteSurplusCopies({
    account: {
      serverURL: "https://example.invalid",
      custom: { user: "u", deviceId: "d" },
    },
    asVersion: "16.1",
    collectionId: "42",
    className: "Calendar",
    filterType: "7",
    conflict: "1",
    synckey,
    serverIds,
    persistSyncKey: async (key) => persisted.push(key),
    // The pause between chunks is for a throttling mailbox, not for
    // anything under test here.
    paceMs: 0,
    ...rest,
  });
}

test("each surplus copy goes out as a Delete naming its ServerId", async () => {
  replies = [accepted("key-1")];
  await run(["S1", "S2", "S3"]);
  assert.equal(sent.length, 1);
  const ids = [...sent[0].matchAll(/<Delete><ServerId>(.*?)<\/ServerId>/g)].map(
    (m) => m[1],
  );
  assert.deepEqual(ids, ["S1", "S2", "S3"]);
});

test("the request states GetChanges 0 - a cleanup is not a pull", async () => {
  replies = [accepted("key-1")];
  await run(["S1"]);
  assert.match(sent[0], /<GetChanges>0<\/GetChanges>/);
});

test("the copy that stays is never named", async () => {
  // The keeper is excluded before this is called; the guard is that
  // nothing here re-derives the list.
  replies = [accepted("key-1")];
  await run(["S1", "S2"]);
  assert.ok(!sent[0].includes("KEEPER"));
});

test("the SyncKey advances across chunks and is persisted for each", async () => {
  replies = [accepted("key-1"), accepted("key-2"), accepted("key-3")];
  const ids = Array.from({ length: 60 }, (_, i) => `S${i}`);
  const { deleted } = await run(ids);
  assert.equal(sent.length, 3, "60 ids is three chunks of 25");
  assert.match(sent[0], /<SyncKey>key-0<\/SyncKey>/);
  assert.match(sent[1], /<SyncKey>key-1<\/SyncKey>/);
  assert.match(sent[2], /<SyncKey>key-2<\/SyncKey>/);
  assert.deepEqual(persisted, ["key-1", "key-2", "key-3"]);
  assert.equal(deleted, 60);
});

test("a copy the server no longer has counts as removed", async () => {
  // Status 8 is "object not found". The list the user acted on may be
  // minutes old, and something else having removed a copy is not a
  // failure of this run.
  replies = [answering("key-1", { S1: "1", S2: "8" })];
  const { deleted, failed } = await run(["S1", "S2"]);
  assert.equal(deleted, 2);
  assert.deepEqual(failed, []);
});

test("silence is consent - an unanswered id succeeded", async () => {
  replies = [answering("key-1", { S2: "1" })];
  const { deleted, failed } = await run(["S1", "S2"]);
  assert.equal(deleted, 2);
  assert.deepEqual(failed, []);
});

test("a refused copy is reported with the status the server gave", async () => {
  replies = [answering("key-1", { S1: "1", S2: "3" })];
  const { deleted, failed } = await run(["S1", "S2"]);
  assert.equal(deleted, 1);
  assert.deepEqual(failed, [{ serverId: "S2", status: "3" }]);
});

test("a collection-level failure stops the run and keeps the key", async () => {
  replies = [
    accepted("key-1"),
    `<?xml version="1.0" encoding="utf-8"?>
     <Sync xmlns="AirSync"><Collections><Collection>
       <SyncKey>key-2</SyncKey><CollectionId>42</CollectionId><Status>3</Status>
     </Collection></Collections></Sync>`,
  ];
  const ids = Array.from({ length: 40 }, (_, i) => `S${i}`);
  await assert.rejects(run(ids), /Sync collection status 3/);
  // The first chunk's deletions are permanent and its key was banked
  // before the second chunk was ever sent.
  assert.deepEqual(persisted, ["key-1"]);
});

test("a folder that has never synced is refused outright", async () => {
  replies = [accepted("key-1")];
  await assert.rejects(
    run(["S1"], { synckey: "0" }),
    /refusing to delete on an unsynced folder/,
  );
  assert.equal(sent.length, 0, "nothing may go out");
});
