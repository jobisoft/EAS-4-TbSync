/**
 * Which local item is this server id, and what does the server call this
 * item - one map for one folder, answering both directions.
 *
 * Every item passes through the sync at least once, coming down from the
 * server or going up from the user, and the mapping is recorded there. So
 * the map is complete as long as it is kept, and the whole of keeping it
 * honest is this: **only `set` and `remove` write to either table**, and
 * `set` replaces. A uid maps to exactly one server id, so an id the server
 * has superseded stops resolving the moment the item is re-stamped.
 *
 * Two tables rather than one because both questions are asked on the hot
 * path - the pull asks "which item is this id" for every command, the push
 * asks "what is this item called" for every queued edit - and a folder can
 * hold thousands of items. They cannot drift apart because nothing else can
 * write to them.
 *
 * What is stored is the array of `{uid, serverId}` pairs in the folder's
 * `custom.indexMap`, which the host keeps and never interprets; `toArray`
 * is what goes back there, once per pass and only when something changed.
 */

export function createServerIdIndex(stored) {
  /** uid -> serverId. What the server calls each item we hold. */
  const byUid = new Map();
  /** serverId -> uid. The same pairs, read the other way. */
  const byServerId = new Map();
  let dirty = false;

  function set(uid, serverId) {
    if (!uid || !serverId) return;
    const prev = byUid.get(uid);
    if (prev === serverId) return;
    // Withdraw the old reverse entry, but only while it is still ours to
    // withdraw: if two items claim one server id, the reverse direction
    // belongs to whoever claimed it last.
    if (prev !== undefined && byServerId.get(prev) === uid) {
      byServerId.delete(prev);
    }
    byUid.set(uid, serverId);
    byServerId.set(serverId, uid);
    dirty = true;
  }

  for (const e of Array.isArray(stored) ? stored : []) {
    set(e?.uid, e?.serverId);
  }
  // Loading what was already stored is not a change to it.
  dirty = false;

  return {
    get dirty() {
      return dirty;
    },
    get size() {
      return byUid.size;
    },

    /** The local item this server id names, or null. */
    uidFor(serverId) {
      return (serverId ? byServerId.get(serverId) : null) ?? null;
    },

    /** What the server calls this local item, or null. */
    serverIdFor(uid) {
      return (uid ? byUid.get(uid) : null) ?? null;
    },

    set,

    remove(uid) {
      const prev = byUid.get(uid);
      if (prev === undefined) return;
      byUid.delete(uid);
      // The loser of a contested server id must not take the winner's
      // reverse entry with it.
      if (byServerId.get(prev) === uid) byServerId.delete(prev);
      dirty = true;
    },

    /**
     * Rebuild from the server ids stamped into the stored blobs.
     *
     * The index is written once per pass while the blobs are stamped item by
     * item, so a crash or a failed flush - or anything that clears the
     * stored index directly - leaves the blobs holding mappings the index
     * has lost. Nothing else would ever put them back, and the cost of not
     * putting them back is a duplicate of every item the server re-sends:
     * an EAS task or contact carries no `UID` on the wire, so `applyAdd` has
     * no second way to recognise its own item.
     *
     * The blob wins over anything already held, because the blob is the
     * authority; among blobs carrying the same id the first wins, which
     * keeps the result stable for a store that already holds duplicates. An
     * item with no stamp was created locally and never pushed, so it belongs
     * to no server id at all. A blob the codec cannot read simply cannot
     * answer.
     */
    fill(items, readStamp) {
      const claimed = new Set();
      for (const it of items ?? []) {
        if (!it?.id || !it?.blob) continue;
        let stamped;
        try {
          stamped = readStamp(it.blob);
        } catch {
          continue;
        }
        if (!stamped || claimed.has(stamped)) continue;
        claimed.add(stamped);
        set(it.id, stamped);
      }
    },

    toArray() {
      return [...byUid].map(([uid, serverId]) => ({ uid, serverId }));
    },
  };
}
