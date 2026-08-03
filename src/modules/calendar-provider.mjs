/**
 * EAS as a calendar provider.
 *
 * Until now every EAS folder synced into a plain `storage` calendar, which
 * cannot tell one kind of write from another: the user editing an event and
 * the sync writing down what the server just sent look identical. TbSync
 * compensates by hand - `changelogMarkServerWrite` pre-tags an entry so the
 * observer ignores the write it is about to make - and even then the record
 * it keeps is item-level, so a series that changed one occurrence is re-sent
 * whole.
 *
 * A provider calendar has the two directions as separate objects:
 *
 *   user edits    → `calendar.provider.onItem{Created,Updated,Removed}`
 *   our own sync  → written to `<id>#cache`, which fires none of those
 *
 * That is structural rather than tagged, and `onItemUpdated` is handed the
 * *previous* item, so we can see which occurrence of a series moved instead
 * of guessing. Measured, not assumed: both directions were driven on a
 * throwaway calendar before this was written.
 *
 * ## Why the hooks do not push
 *
 * A hook fires the instant the user hits save and the platform waits for an
 * answer. Pushing to Exchange there would make every edit a network round
 * trip and would fail the edit outright when the connection is down. So we
 * accept the item and let the scheduled sync push it - the same model TbSync
 * has always had.
 *
 * ## Where the record goes
 *
 * Into the host changelog, the same queue as always, via
 * `changelogRecordUserEdit`. The provider holds no state of its own: it does
 * not own the account and folder rows, and a queue that lived here would be
 * lost on every reload, update or background suspend - exactly the edits the
 * changelog exists to keep.
 *
 * What changes is only who writes the entry. For a calendar we supply, the
 * host stops observing and we report, which lets the entry carry what the
 * item looked like *before* the edit. Nothing on the host side can
 * reconstruct that once the new version has been written, and it is what
 * makes a per-exception push possible.
 *
 * Address books are untouched - they have no provider API, so the observer
 * still owns those entries.
 */

import { exceptionFingerprint } from "./eas/calendar-codec.mjs";

/** Report an edit to the host, which folds it into the folder's changelog.
 *
 *  Awaited by the hooks: the platform is holding the user's save until we
 *  answer, and answering before the record is durable would lose the edit if
 *  anything went wrong in between. A failure is logged rather than thrown -
 *  refusing the edit would be worse than a missed sync, and the next full
 *  compare still finds it.
 *
 *  `detail` carries only what cannot be re-derived: a fingerprint of the
 *  exception set as it was *before* the edit. Not the previous item itself -
 *  that is kilobytes per pending entry in a folder row, to answer a question
 *  ("which exceptions differ?") that a fixed-size digest answers just as
 *  well. */
async function record(calendarId, itemId, op, { type, oldIcal } = {}) {
  if (!calendarId || !itemId || !host) return;
  const before = oldIcal ? exceptionFingerprint(oldIcal) : null;
  // One calendar type serves both, so the item says which it is - a task
  // folder's edits must not be queued as events.
  const kind = type === "task" ? "task" : "event";
  try {
    await host.changelogRecordUserEdit({
      parentId: calendarId,
      itemId,
      kind,
      op,
      ...(before ? { detail: { exceptions: before } } : {}),
    });
    report?.({
      level: "debug",
      message:
        `[${kind}-sync] recorded user ${op} of ${itemId}` +
        (before
          ? ` (${before.exdates.length} cancelled, ${before.overrides.length} override(s) before the edit)`
          : ""),
    });
  } catch (err) {
    report?.({
      level: "warning",
      message: `[${kind}-sync] could not record user ${op} of ${itemId}: ${err?.message ?? String(err)}`,
    });
  }
}

let registered = false;

/** Register the item hooks. Safe to call more than once.
 *
 *  Each hook must return the item, or the platform treats the edit as
 *  failed and the user's change disappears from the UI. Returning it
 *  unchanged is us accepting the edit as-is; the sync will correct it
 *  later if the server disagrees. */
export function registerCalendarProvider() {
  if (registered) return;
  const provider = messenger.calendar?.provider;
  if (!provider) {
    console.warn("[eas] calendar.provider API missing; not a provider");
    return;
  }

  const options = { returnFormat: "ical" };

  provider.onItemCreated.addListener(async (calendar, item) => {
    await record(calendar?.id, item?.id, "created", { type: item?.type });
    return { item };
  }, options);

  provider.onItemUpdated.addListener(async (calendar, item, oldItem) => {
    await record(calendar?.id, item?.id, "updated", {
      type: item?.type,
      oldIcal: oldItem?.item ?? null,
    });
    return { item };
  }, options);

  provider.onItemRemoved.addListener(async (calendar, item) => {
    await record(calendar?.id, item?.id, "deleted", { type: item?.type });
    return {};
  }, options);

  // Fired by the calendar's own refresh - the reload button, and the
  // periodic refresh Thunderbird runs for provider calendars. TbSync owns
  // scheduling, so this is a request to sync, not a sync of its own.
  provider.onSync.addListener((calendar) => {
    onSyncRequested?.(calendar?.id);
  });

  // A user asking for a clean slate through the calendar UI. The folder's
  // sync key has to go, which only the host can do, so it is handed up the
  // same way.
  provider.onResetSync.addListener((calendar) => {
    onResetRequested?.(calendar?.id);
  });

  registered = true;
}

/* Wired by background.mjs, which is the only thing that knows how to reach
 * the host. Kept as plain hooks so this module has no dependency on the
 * provider object. */
let onSyncRequested = null;
let onResetRequested = null;
let report = null;
let host = null;

export function setSyncHandlers({
  onSync,
  onReset,
  reportEventLog,
  provider,
} = {}) {
  onSyncRequested = onSync ?? null;
  onResetRequested = onReset ?? null;
  report = reportEventLog ?? null;
  host = provider ?? null;
}
