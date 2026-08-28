/**
 * Unit tests for spotting a UID the server holds more than once.
 *
 * The rule the tests pin is narrow on purpose: only the ServerIds a sync
 * actually heard the server claim for a UID count as evidence. The stamp
 * the local item already carried does not, because a server below 16.1
 * may re-mint every ServerId in a folder after a resync - and a rule that
 * compared against the stamp would then report the whole calendar as
 * duplicated and offer to delete it. The re-mint test below is that case.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";

const { duplicateClusters, noteUidClaim, titleFromBlob } =
  await import("../../src/modules/eas/duplicate-uids.mjs");

/** Build a claims map the way a sync would, from `[uid, serverId]` pairs. */
function claiming(...pairs) {
  const claims = new Map();
  for (const [uid, serverId] of pairs) noteUidClaim(claims, uid, serverId);
  return claims;
}

const bound = (map) => ({
  serverIdFor: (uid) => map[uid] ?? null,
  titleFor: async () => "",
});

test("one ServerId per UID is not a duplicate", async () => {
  const claims = claiming(["a", "S1"], ["b", "S2"], ["c", "S3"]);
  const found = await duplicateClusters(
    claims,
    bound({ a: "S1", b: "S2", c: "S3" }),
  );
  assert.deepEqual(found, []);
});

test("a UID claimed twice is reported, keeping the bound copy", async () => {
  const claims = claiming(["a", "S1"], ["a", "S2"]);
  const [cluster] = await duplicateClusters(claims, bound({ a: "S2" }));
  assert.equal(cluster.uid, "a");
  assert.equal(cluster.keeper, "S2");
  assert.deepEqual(cluster.surplus, ["S1"]);
});

test("the surplus count excludes the copy that stays", async () => {
  // Peters' cluster, in miniature: 533 copies means 532 to remove.
  const pairs = Array.from({ length: 533 }, (_, i) => ["a", `S${i}`]);
  const found = await duplicateClusters(
    claiming(...pairs),
    bound({ a: "S532" }),
  );
  assert.equal(found[0].surplus.length, 532);
  assert.ok(!found[0].surplus.includes("S532"));
});

test("the same ServerId claimed twice in one sync is one copy", async () => {
  // A pull can name an item in both Commands and Responses; that is one
  // item being talked about twice, not two items.
  const claims = claiming(["a", "S1"], ["a", "S1"]);
  assert.deepEqual(await duplicateClusters(claims, bound({ a: "S1" })), []);
});

test("a folder whose ServerIds were all re-minted reports nothing", async () => {
  // The resync case. Every item comes back under a new id, so every UID
  // has exactly one claim - and the old stamps, which do differ, are not
  // evidence and never reach this.
  const claims = claiming(["a", "NEW1"], ["b", "NEW2"], ["c", "NEW3"]);
  const found = await duplicateClusters(
    claims,
    bound({ a: "NEW1", b: "NEW2", c: "NEW3" }),
  );
  assert.deepEqual(found, []);
});

test("nothing is offered when no local copy is bound any more", async () => {
  // Deleted during the very sync that saw the copies: with no copy to
  // keep, every id would be a deletion candidate, and offering that on
  // the strength of a race is how a real item is lost.
  const claims = claiming(["a", "S1"], ["a", "S2"]);
  assert.deepEqual(await duplicateClusters(claims, bound({})), []);
});

test("a claim without a uid or a server id is not a claim", async () => {
  const claims = claiming(["a", "S1"], [null, "S2"], ["a", null], ["a", ""]);
  assert.deepEqual(await duplicateClusters(claims, bound({ a: "S1" })), []);
});

test("the title comes off the stored item, unfolded", async () => {
  // Folded mid-word, which is what a 75-octet line break does to peters'
  // longest title. Unfolding drops the CRLF and the one space that marks
  // the continuation, so the halves join with nothing between them.
  assert.equal(
    titleFromBlob(
      "BEGIN:VEVENT\r\nSUMMARY:Kultur\r\n konferenz Ruhr\r\nEND:VEVENT",
    ),
    "Kulturkonferenz Ruhr",
  );
  assert.equal(
    titleFromBlob("BEGIN:VCARD\r\nFN:Ada Lovelace\r\nEND:VCARD"),
    "Ada Lovelace",
  );
  assert.equal(titleFromBlob("SUMMARY;LANGUAGE=de:Titel"), "Titel");
  assert.equal(titleFromBlob("BEGIN:VEVENT\r\nEND:VEVENT"), "");
  assert.equal(titleFromBlob(null), "");
});
