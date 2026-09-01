/**
 * Sync a resource shortly after the user stops editing it.
 *
 * Both of the account's resource kinds use this. A calendar's edits arrive
 * as item hooks, a card's through the address-book watcher's user-edit
 * callback, and from here on the two are the same thing: a person changed
 * something, and the account should not wait for its scheduled run to say
 * so.
 *
 * How long it waits is the account's `syncOnChange`, in seconds, with `0`
 * meaning never.
 *
 * `host` and `report` are passed in rather than held here, because the two
 * callers reach them differently - one from module state wired by the
 * background page, the other from the provider instance it is a method of.
 */

/** The window an account that predates the setting keeps, and what an
 *  unreadable one falls back to. Matches the provider's
 *  `DEFAULT_SYNC_ON_CHANGE`. */
const DEFAULT_QUIET_SECONDS = 15;

/** Pending sync per resource, keyed by target id, so edits in one calendar
 *  or address book never postpone another's. */
const syncAfterChangeTimers = new Map();

/** How long this account wants a resource to stay quiet before its edits
 *  are synced, in seconds, `0` meaning never. Read rather than remembered:
 *  the caller is already asking the host to store the pending count, so this
 *  is a second call beside one we make anyway, and it means a change to the
 *  setting takes effect on the very next edit. */
export async function quietSecondsFor(host, accountId) {
  try {
    const { account } = (await host?.getAccount(accountId)) ?? {};
    const stored = Number(account?.custom?.syncOnChange);
    return Number.isFinite(stored) ? stored : DEFAULT_QUIET_SECONDS;
  } catch {
    return DEFAULT_QUIET_SECONDS;
  }
}

/**
 * Arm the sync for one resource.
 *
 * Armed from a user edit and nowhere else, which is what keeps it from
 * feeding itself. A calendar's own sync writes go to `<id>#cache` and raise
 * no item hook; an address book's are announced to the watcher first and
 * consumed there. Either way what reaches this is a person editing
 * something.
 *
 * Re-arming on each edit rather than syncing on the first is the whole
 * point - a burst of writes is one sync at the end of it.
 *
 * Scoped to the resource that changed. `requestSync` takes a target id and
 * the host resolves it to that one folder, whichever kind it is, so a change
 * in one does not drag the account's others along.
 *
 * Nothing here can fail a save: the edit is already durable by the time this
 * is called, and the worst case is that it waits for the scheduled sync
 * exactly as it did before this existed.
 *
 * `kind` names the resource in the log and does nothing else.
 */
export function armSyncAfterChange(host, report, targetID, seconds, kind) {
  clearTimeout(syncAfterChangeTimers.get(targetID));
  if (!seconds) return;
  syncAfterChangeTimers.set(
    targetID,
    setTimeout(async () => {
      syncAfterChangeTimers.delete(targetID);
      try {
        if (!host) return;
        report?.({
          level: "debug",
          message: `[${kind}-sync] syncing ${targetID}, ${seconds}s after the edits stopped`,
        });
        await host.requestSync({ parentId: targetID });
      } catch (err) {
        // Includes the ordinary cases: the account was disconnected while
        // the timer ran, or a sync is already going and the host refused
        // this one. Neither loses anything - the edit stays queued.
        report?.({
          level: "debug",
          message: `[${kind}-sync] sync after change of ${targetID} did not run: ${err?.message ?? String(err)}`,
        });
      }
    }, seconds * 1000),
  );
}
