/**
 * Where a pending user edit waits until the next sync can push it.
 *
 * For an address book that is still the host's `folder.changelog`: the host
 * observes Thunderbird's book events, so the host is where an edit is first
 * known and the queue may as well live there. `hostQueue` below is a thin
 * adapter over those RPCs.
 *
 * For a calendar we supply, it is here. The platform hands us the edit
 * directly and holds the user's save until we answer, so the record must be
 * durable before we do - and it must not depend on anything outside this
 * add-on being alive. An enabled EAS provider keeps working calendars with
 * the host absent, and a record that needed the host would be unmakeable on
 * every host reload, update and background suspend.
 *
 * ## Sessions
 *
 * A folder row outlives its bindings: deselect and reselect, delete the
 * calendar and let the next sync re-create it, and the queue from before
 * belongs to something that is gone. Pushing those edits into the new
 * binding would resurrect items the user deleted along with the calendar.
 *
 * The host names the current binding in `folder.sessionId` and mints a new
 * one whenever it ends one. So every key here is a session id and nothing
 * else: finding our queue means looking up the session the row names, and a
 * queue whose session no row names is garbage by that fact alone. `sweep()`
 * drops those. No teardown message is needed, and none is trusted -
 * Disconnect and Remove have to work when this add-on is broken, so they
 * cannot depend on it doing anything.
 *
 * The consequence worth stating: edits made while the host is down are
 * filed under the last session we saw. If that binding was torn down in the
 * meantime, they go with it. That is the correct answer, not a compromise.
 */

import { serialize } from "../../vendor/tbsync/storage-queue.mjs";
import {
  isUserEntry,
  moveToTailUpdater,
  recordUserEditUpdater,
  removeEntryUpdater,
} from "../../vendor/tbsync/changelog-core.mjs";

/** One key per session. The value carries the account and folder it belongs
 *  to, so a sweep can report what it dropped without parsing keys. */
const QUEUE_PREFIX = "queue.";

/** Which folder each calendar we supply belongs to, as
 *  `targetID -> {accountId, folderId, sessionId}`.
 *
 *  The item hooks are handed a calendar id and nothing else, and must not
 *  ask the host who owns it - that is the round trip this whole module
 *  exists to remove. So the answer is kept here, refreshed every time we
 *  legitimately have the folder rows in hand (a sync, a lifecycle event).
 *  Stale is survivable and self-correcting: a wrong session files the edit
 *  where the next sweep will find it, which is exactly what a torn-down
 *  binding deserves. */
const BINDINGS_KEY = "queue.bindings";

const queueKey = (sessionId) => `${QUEUE_PREFIX}${sessionId}`;

async function readQueue(sessionId) {
  const key = queueKey(sessionId);
  const rv = await browser.storage.local.get({ [key]: null });
  const bag = rv[key];
  return Array.isArray(bag?.entries) ? bag.entries : [];
}

/** Read, transform, write - serialised against every other storage mutation
 *  in this extension context, so a sync draining the queue and a hook adding
 *  to it cannot interleave. Returning the same array short-circuits the
 *  write, as it does on the host. */
function mutate(binding, updater) {
  const { accountId, folderId, sessionId } = binding;
  return serialize(async () => {
    const key = queueKey(sessionId);
    const rv = await browser.storage.local.get({ [key]: null });
    const bag = rv[key];
    const before = Array.isArray(bag?.entries) ? bag.entries : [];
    const after = updater(before) ?? before;
    if (after === before) return before;
    await browser.storage.local.set({
      [key]: { accountId, folderId, sessionId, entries: after },
    });
    return after;
  });
}

/**
 * The queue for one binding. `binding` must carry the folder's CURRENT
 * `sessionId` - read it from the folder row, never from a variable that has
 * outlived a sync.
 */
export function localQueue(binding) {
  return {
    owner: "local",

    /** Everything waiting to be pushed, oldest first. */
    async pending() {
      const entries = await readQueue(binding.sessionId);
      return entries.filter((e) => isUserEntry(e?.status));
    },

    /** Every row, unfiltered - what GET_CHANGELOG answers with. Identical
     *  to `pending()` in practice, since nothing writes a pre-tag here, but
     *  the two questions are not the same question. */
    async entries() {
      return readQueue(binding.sessionId);
    },

    /** Fold a user edit into whatever is already queued for that item.
     *  Returns the resulting number of pending entries, which the caller
     *  reports to the host for the needs-sync badge. */
    async record({ parentId, itemId, kind, op, detail }) {
      const after = await mutate(binding, (entries) => {
        const result = recordUserEditUpdater(entries, {
          parentId,
          itemId,
          kind,
          op,
          detail,
          now: Date.now(),
        });
        return result.entries;
      });
      return after.filter((e) => isUserEntry(e?.status)).length;
    },

    /** Drop the queued edit for this item - pushed, or established as
     *  unpushable. */
    async remove({ parentId, itemId, kind }) {
      await mutate(binding, (entries) =>
        removeEntryUpdater(entries, { parentId, itemId, kind }),
      );
    },

    /** Send failed items behind the rest, so the next sync tries the
     *  healthy ones first. */
    async moveToTail(items) {
      if (!items?.length) return;
      await mutate(binding, (entries) => moveToTailUpdater(entries, items));
    },

    /** Announce a write of our own so an observer does not log it as the
     *  user's - meaningless here, and deliberately so. Nothing observes a
     *  calendar we supply: our sync writes go to `<id>#cache`, which fires
     *  no item hooks at all. The suppression this names is structural, not
     *  something a marker has to arrange. Kept so both queues answer the
     *  same calls. */
    async markServerWrite() {},

    async count() {
      const entries = await readQueue(binding.sessionId);
      return entries.filter((e) => isUserEntry(e?.status)).length;
    },
  };
}

/**
 * The same shape, backed by the host's changelog RPCs. Used for address
 * books, whose changes the host observes and owns.
 */
export function hostQueue({ provider, accountId, folderId, changelog }) {
  return {
    owner: "host",

    async pending() {
      const entries = Array.isArray(changelog) ? changelog : [];
      return entries.filter((e) => isUserEntry(e?.status));
    },

    // No `record` here on purpose. An address-book edit is not something
    // this add-on is ever handed - the host observes the book and writes
    // the entry itself, which is the whole reason that queue is the host's.
    // A method to report one would have no caller and no way to be right.

    async remove({ parentId, itemId, kind }) {
      await provider.changelogRemove({
        accountId,
        folderId,
        parentId,
        itemId,
        kind,
      });
    },

    async moveToTail(items) {
      if (!items?.length) return;
      await provider.changelogMoveToTail({ accountId, folderId, items });
    },

    async markServerWrite({ parentId, itemId, kind, status }) {
      await provider.changelogMarkServerWrite({
        accountId,
        folderId,
        parentId,
        itemId,
        kind,
        status,
      });
    },

    async count() {
      const entries = Array.isArray(changelog) ? changelog : [];
      return entries.filter((e) => isUserEntry(e?.status)).length;
    },
  };
}

// ── Bindings ──────────────────────────────────────────────────────────────

/** Record which folder each of these calendars belongs to. Called whenever
 *  folder rows are in hand for an honest reason; cheap enough to do every
 *  time and worth more than the bytes it costs, since it is what lets an
 *  item hook answer without the host. */
export function rememberBindings(list) {
  return serialize(async () => {
    const rv = await browser.storage.local.get({ [BINDINGS_KEY]: {} });
    const map = rv[BINDINGS_KEY] ?? {};
    let dirty = false;
    for (const { targetID, accountId, folderId, sessionId } of list) {
      if (!targetID || !sessionId) continue;
      const prior = map[targetID];
      if (
        prior?.accountId === accountId &&
        prior?.folderId === folderId &&
        prior?.sessionId === sessionId
      ) {
        continue;
      }
      map[targetID] = { accountId, folderId, sessionId };
      dirty = true;
    }
    if (dirty) await browser.storage.local.set({ [BINDINGS_KEY]: map });
    return map;
  });
}

/** The folder a calendar belongs to, as last recorded. Null when we have
 *  never been told - a calendar that no sync of ours ever bound. */
export async function lookupBinding(targetID) {
  if (!targetID) return null;
  const rv = await browser.storage.local.get({ [BINDINGS_KEY]: {} });
  return rv[BINDINGS_KEY]?.[targetID] ?? null;
}

// ── Sweeping ──────────────────────────────────────────────────────────────

/**
 * Drop every queue whose session no folder row names, and every binding
 * pointing at one. `live` is the set of session ids currently in the host's
 * folder rows.
 *
 * This is the entire teardown path. A folder deselected, an account
 * disconnected, a calendar deleted, a whole account removed while this
 * add-on was uninstalled - all of them end the same way: the host stops
 * naming a session, and the next time we look we do not recognise it.
 *
 * Only ever called with rows actually read from the host. Sweeping against
 * an empty or partial list would delete live queues, so a caller that could
 * not read the rows must not call this at all.
 */
export function sweep(liveSessionIds) {
  const live = liveSessionIds instanceof Set
    ? liveSessionIds
    : new Set(liveSessionIds ?? []);
  return serialize(async () => {
    const all = await browser.storage.local.get(null);
    const drop = [];
    for (const [key, bag] of Object.entries(all)) {
      if (!key.startsWith(QUEUE_PREFIX)) continue;
      if (key === BINDINGS_KEY) continue;
      const sessionId = key.slice(QUEUE_PREFIX.length);
      if (live.has(sessionId)) continue;
      drop.push({
        key,
        accountId: bag?.accountId ?? null,
        folderId: bag?.folderId ?? null,
        entries: Array.isArray(bag?.entries) ? bag.entries.length : 0,
      });
    }

    const bindings = all[BINDINGS_KEY] ?? {};
    const keptBindings = {};
    let bindingsDirty = false;
    for (const [targetID, b] of Object.entries(bindings)) {
      if (live.has(b?.sessionId)) keptBindings[targetID] = b;
      else bindingsDirty = true;
    }

    if (drop.length) await browser.storage.local.remove(drop.map((d) => d.key));
    if (bindingsDirty) {
      await browser.storage.local.set({ [BINDINGS_KEY]: keptBindings });
    }
    return drop;
  });
}
