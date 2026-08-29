/**
 * Delete the surplus copies a duplicate cluster left on the server.
 *
 * This is its own command rather than a run of changelog entries, because
 * the changelog cannot express it. An entry is keyed by item id - the UID -
 * and the push resolves that to a ServerId through the index map, which is
 * one-to-one and holds only the copy Thunderbird is bound to. A queued
 * delete for a duplicated UID would therefore address the one copy that
 * must be kept, and the surplus - which has no local item and no map entry
 * - is not reachable that way at all.
 *
 * So the ServerIds go out directly. Nothing local is touched: the local
 * item stays exactly as it is, bound to the copy it was already bound to,
 * and neither the store, the changelog nor the index map has anything to
 * say about ids that were never in them.
 *
 * The one piece of sync state this does share is the folder's SyncKey. A
 * `Sync` carrying commands consumes it and returns the next one, so the
 * caller's `persistSyncKey` runs after every chunk - skip that and the
 * next ordinary sync presents a key the server has moved past, draws a
 * Status 3 and resyncs the whole folder. It is also why this must not run
 * while that folder is syncing.
 *
 * `<GetChanges>0</GetChanges>` rides along with the commands, as it does on
 * every push: without it a non-zero SyncKey means "send me changes too",
 * and the answer would be a pull nobody is here to apply.
 */

import { easRequest } from "../network.mjs";
import { buildSyncBody } from "./sync-body.mjs";
import { childByTag, readPath, readPathFrom } from "./wbxml-helpers.mjs";

/** Deletes per request. The server's own window size governs how many
 *  commands it will take in one go, and 25 is what the push uses. */
const CHUNK = 25;

/** Between chunks. A 16.x mailbox throttles by request rate, and a
 *  cleanup is the one operation here that can fire twenty requests with
 *  nothing in between - 533 copies is 22 of them. */
const PACE_MS = 1000;

/** Both endings that mean the copy is gone: the server deleted it, or it
 *  was not there to delete. A cluster the user is shown may be minutes
 *  old, and something else removing a copy in between is not a failure. */
const DONE = new Set(["1", "8"]);

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * Send `<Delete>` for each of `serverIds`, in chunks, advancing the
 * folder's SyncKey as it goes.
 *
 * Returns `{deleted, failed}` - counts, with `failed` carrying the status
 * the server gave. Stops at the first chunk that fails outright, keeping
 * what earlier chunks achieved: every delete already acknowledged is
 * permanent, and the rest are still listed for a second attempt.
 */
export async function deleteSurplusCopies({
  account,
  asVersion,
  collectionId,
  className,
  filterType,
  conflict,
  synckey,
  serverIds,
  persistSyncKey,
  onProgress,
  paceMs = PACE_MS,
}) {
  let key = String(synckey ?? "0");
  if (key === "0") {
    // Nothing has been synced under this key, so the ServerIds we were
    // given belong to a state the server no longer agrees we are in.
    throw new Error("refusing to delete on an unsynced folder");
  }
  const failed = [];
  let deleted = 0;

  for (let i = 0; i < serverIds.length; i += CHUNK) {
    const chunk = serverIds.slice(i, i + CHUNK);
    const body = buildSyncBody({
      synckey: key,
      collectionId,
      asVersion,
      withCommands: {
        adds: [],
        mods: [],
        dels: chunk.map((serverID) => ({ serverID })),
        asVersion,
      },
      className,
      filterType,
      conflict,
    });
    const { doc } = await easRequest({
      account,
      command: "Sync",
      body,
      asVersion,
    });
    if (!doc) throw new Error("empty Sync response to a delete");

    const topStatus = readPath(doc, ["Status"]);
    if (topStatus !== null && topStatus !== "1") {
      throw new Error(`Sync top status ${topStatus}`);
    }
    // Walked rather than looked up by tag name: `Status` and `SyncKey`
    // appear at more than one depth in a Sync response, and a flat search
    // would answer from whichever came first.
    const collections = childByTag(doc.documentElement, "Collections");
    const collection = collections && childByTag(collections, "Collection");
    if (!collection) throw new Error("Sync response without a Collection");
    const collStatus = readPathFrom(collection, ["Status"]) ?? "1";
    if (collStatus !== "1") {
      throw new Error(`Sync collection status ${collStatus}`);
    }
    const nextKey = readPathFrom(collection, ["SyncKey"]);
    if (!nextKey) throw new Error("Sync response without a SyncKey");
    key = nextKey;
    // Before the statuses are read, not after: the key has moved whatever
    // the individual deletes said, and losing it costs a full resync.
    await persistSyncKey?.(key);

    const responses = childByTag(collection, "Responses");
    const seen = new Map();
    for (const node of responses ? Array.from(responses.children) : []) {
      if (node.tagName !== "Delete") continue;
      const id = readPathFrom(node, ["ServerId"]);
      if (id) seen.set(id, readPathFrom(node, ["Status"]) ?? "1");
    }
    for (const id of chunk) {
      // Silence is consent: [MS-ASCMD] 2.2.3.155 has the server answer
      // only the commands it has something to report, so a ServerId
      // missing from Responses succeeded.
      const status = seen.get(id) ?? "1";
      if (DONE.has(status)) deleted += 1;
      else failed.push({ serverId: id, status });
    }
    onProgress?.({ deleted, failed: failed.length, total: serverIds.length });
    if (i + CHUNK < serverIds.length && paceMs) await sleep(paceMs);
  }

  return { deleted, failed };
}
