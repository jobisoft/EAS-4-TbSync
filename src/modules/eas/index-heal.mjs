/**
 * Reconcile a folder's identity map against the items themselves.
 *
 * The map answers "which local item is this server id?", and every pulled
 * command depends on it: a miss means the item is new, and the pull acts on
 * that by creating one. It is maintained by the sync, which is enough while
 * every departure passes through us - but an item can leave without our
 * sync seeing it. Another add-on or the user deleting during the seconds a
 * provider is restarting, a profile restored underneath us, a server below
 * 16.1 re-minting its ids after a resync: none of those reach a hook, and
 * nothing else ever removes the entry.
 *
 * Mostly those strays are inert - the id resolves to a uid, the item is not
 * there, and the command is treated as new, which is right. One case is
 * not. A stray entry still names a server id that another item may since
 * have been given, and a delete queued for the stray's uid would then be
 * addressed at the live item and remove *that* from the server.
 *
 * This is the tidying pass, not the safety net. A map that is entirely lost
 * is repaired the moment it is noticed, because every command would
 * otherwise duplicate its item; see `findExistingByServerId`. What is left
 * for here is drift too small to notice and strays nothing detects at all,
 * neither of which is urgent - hence once a day, in a slot the host
 * promises will not collide with a sync.
 */

import { createServerIdIndex } from "./server-id-index.mjs";

/** How stale a folder's last reconcile may be before it is redone. */
export const HEAL_INTERVAL_MS = 24 * 60 * 60_000;

export function healIsDue(folder, now = Date.now()) {
  const last = Number(folder?.custom?.indexHealed ?? 0);
  if (!Number.isFinite(last)) return true;
  return now - last >= HEAL_INTERVAL_MS;
}

/**
 * Rebuild one folder's map from the stamps in its items, and drop entries
 * that name nothing.
 *
 * Returns `{ rebuilt, pruned, entries }`, or null when the folder could not
 * be read - which is deliberately not the same as "the folder is empty".
 *
 * `queuedIds` is the set of item ids with a changelog entry. An entry whose
 * item is gone is kept when one exists: a delete the user has made and we
 * have not yet pushed looks exactly like a stray, and it is the mapping
 * that will address the delete when it finally goes.
 */
export async function healFolderIndex({ store, readStamp, stored, queuedIds }) {
  let items;
  try {
    items = await store.list();
  } catch {
    // A folder we cannot read tells us nothing about which entries are
    // stale. Pruning against it would drop every mapping the folder has.
    return null;
  }

  const index = createServerIdIndex(stored);
  const before = index.size;
  index.fill(items, readStamp);

  // Existence, not stamps: a blob the codec cannot parse is skipped by the
  // fill, but its item is still there and its mapping is still true.
  const present = new Set(items.map((it) => it?.id).filter(Boolean));
  let pruned = 0;
  for (const { uid } of index.toArray()) {
    if (present.has(uid) || queuedIds.has(uid)) continue;
    index.remove(uid);
    pruned += 1;
  }

  return { rebuilt: index.size - before + pruned, pruned, entries: index.toArray() };
}
