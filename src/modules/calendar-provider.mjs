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
 * Into our own storage, via `change-queue.mjs`. The hook is holding the
 * user's save, so the record has to be durable before we answer - and it
 * cannot be conditional on anything outside this add-on being alive.
 *
 * It used to go to the host, over `changelogRecordUserEdit`. That made the
 * record depend on the host answering, and when it did not, the only two
 * available answers were both wrong: refuse the edit and destroy the user's
 * work, or accept it and let the queue never hear of it. It accepted, and
 * logged a warning nobody reads. An enabled EAS provider keeps its
 * calendars working with or without the host, so that window opened on
 * every host reload, update and background suspend.
 *
 * Writing it here removes the dependency rather than narrowing it. What the
 * host still gets is a count for its needs-sync badge, sent best-effort: if
 * that does not arrive, a dot is missing until the next sync, and the edit
 * is safe either way.
 *
 * Address books are untouched - they have no provider API, so the observer
 * still owns those entries and they stay in the host's changelog.
 */

import { exceptionFingerprint } from "./eas/calendar-codec.mjs";
import {
  localQueue,
  lookupBinding,
  rememberBindings,
} from "./eas/change-queue.mjs";

/** Every folder of ours that is bound to a calendar, as
 *  `targetID -> {accountId, folderId, targetName, targetColor}`. The host owns
 *  the folder table; this is a read of it, refreshed on demand rather than
 *  cached, since a target id changes whenever a folder is rebound.
 *
 *  The remembered name and colour ride along because the lifecycle listener
 *  needs them to tell a real change from the platform re-announcing what we
 *  already know. */
async function ourTargets() {
  const out = new Map();
  if (!host) return out;
  const bindings = [];
  for (const { accountId } of await host.listAccounts()) {
    const { folders = [] } = (await host.getAccount(accountId)) ?? {};
    for (const f of folders) {
      if (!f?.targetID) continue;
      if (f.targetType !== "calendars" && f.targetType !== "tasks") continue;
      out.set(f.targetID, {
        accountId,
        folderId: f.folderId,
        sessionId: f.sessionId ?? null,
        targetName: f.targetName ?? null,
        targetColor: f.targetColor ?? null,
      });
      bindings.push({
        targetID: f.targetID,
        accountId,
        folderId: f.folderId,
        sessionId: f.sessionId ?? null,
      });
    }
  }
  // We have the rows in hand for an honest reason, so bank what an item
  // hook will need when it has no way to ask: which folder a calendar id
  // belongs to, and which binding of it is current.
  await rememberBindings(bindings).catch((err) =>
    report?.({
      level: "debug",
      message: `[target] could not cache folder bindings: ${err?.message ?? String(err)}`,
    }),
  );
  return out;
}

/** Queue an edit the platform just handed us, in our own storage.
 *
 *  Awaited by the hooks: the platform is holding the user's save until we
 *  answer, and answering before the record is durable would lose the edit if
 *  anything went wrong in between. Nothing outside this add-on is involved,
 *  which is the point - see the module header.
 *
 *  `detail` carries only what cannot be re-derived: a fingerprint of the
 *  exception set as it was *before* the edit. Not the previous item itself -
 *  that is kilobytes per pending entry, to answer a question ("which
 *  exceptions differ?") that a fixed-size digest answers just as well. */
async function record(calendarId, itemId, op, { type, oldIcal } = {}) {
  if (!calendarId || !itemId) return;
  const before = oldIcal ? exceptionFingerprint(oldIcal) : null;
  // One calendar type serves both, so the item says which it is - a task
  // folder's edits must not be queued as events.
  const kind = type === "task" ? "task" : "event";

  // Which folder this calendar belongs to, and which binding of it is
  // current. Normally already known - every sync banks it - and then this
  // costs one storage read and no round trip, which is the point.
  //
  // A miss is worth one attempt at the host before giving up: it means we
  // have never held this folder's row (a resource enabled since our last
  // sync of it), and the host is the only place that answer exists. If the
  // host cannot answer either, there is genuinely nothing to file against
  // and saying so is the honest outcome - inventing a binding would push
  // the edit into whatever folder the id later turns out to belong to.
  let binding = await lookupBinding(calendarId);
  if (!binding?.sessionId && host) {
    await ourTargets().catch(() => {});
    binding = await lookupBinding(calendarId);
  }
  if (!binding?.sessionId) {
    report?.({
      level: "warning",
      message:
        `[${kind}-sync] cannot record user ${op} of ${itemId}: ` +
        `calendar ${calendarId} is not bound to a folder we know`,
    });
    return;
  }

  let pending;
  try {
    pending = await localQueue(binding).record({
      parentId: calendarId,
      itemId,
      kind,
      op,
      ...(before ? { detail: { exceptions: before } } : {}),
    });
  } catch (err) {
    // Storage itself failed - out of quota, or shutting down mid-write.
    // Nothing is recoverable here and the edit is genuinely lost, so this
    // is an error, not the warning the host round trip used to log.
    report?.({
      level: "error",
      message: `[${kind}-sync] FAILED to queue user ${op} of ${itemId}: ${err?.message ?? String(err)}`,
    });
    return;
  }

  report?.({
    level: "debug",
    message:
      `[${kind}-sync] queued user ${op} of ${itemId}` +
      (before
        ? ` (${before.exdates.length} cancelled, ${before.overrides.length} override(s) before the edit)`
        : ""),
  });

  // The host paints a needs-sync badge and cannot count a queue it does not
  // hold. Best-effort on purpose: the edit is already safe, and a badge
  // that lags until the next sync is not worth failing a save over.
  await host
    ?.updateFolder({
      accountId: binding.accountId,
      folderId: binding.folderId,
      patch: { custom: { pendingUserChanges: pending } },
    })
    .catch((err) =>
      report?.({
        level: "debug",
        message: `[${kind}-sync] could not update the pending count: ${err?.message ?? String(err)}`,
      }),
    );
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

  // The calendar asking to be refreshed: the Reload button, and the timer
  // Thunderbird arms for every calendar that says it can refresh - ours does.
  // We do not sync ourselves; the host owns the schedule, the folder status it
  // paints into the manager, and the lock that keeps two runs apart. Awaited,
  // so the calendar's spinner ends when the sync does rather than at once.
  provider.onSync.addListener(async (calendar) => {
    const targetID = calendar?.id;
    if (!targetID || !host) return;
    try {
      await host.requestSync({ parentId: targetID });
    } catch (err) {
      // Name it. Once the binding is gone nothing can resolve the id back to
      // a folder, so a bare uuid leaves the reader correlating by timestamp -
      // and this is the line that says a live calendar is asking for syncs we
      // cannot serve.
      const name = calendar?.name ? `"${calendar.name}" ` : "";
      report?.({
        level: "warning",
        message: `[target] sync request for ${name}${targetID} failed: ${err?.message ?? String(err)}`,
      });
    }
  });

  // A user asking, through the calendar's own properties, to start over. The
  // cache has been emptied, so our sync key and index describe items that are
  // no longer there - clear them before asking for the sync, or the server
  // answers "nothing has changed" and the calendar stays empty.
  provider.onResetSync.addListener(async (calendar) => {
    const targetID = calendar?.id;
    if (!targetID || !host) return;
    const mine = (await ourTargets()).get(targetID);
    if (!mine) return;
    try {
      await host.updateFolder({
        accountId: mine.accountId,
        folderId: mine.folderId,
        patch: { custom: { synckey: "0", indexMap: [] } },
      });
      await host.requestSync({ parentId: targetID });
    } catch (err) {
      report?.({
        level: "warning",
        message: `[target] reset of ${mine.accountId}/${mine.folderId} failed: ${err?.message ?? String(err)}`,
      });
    }
  });

  registerTargetLifecycle();
  registered = true;
}

/**
 * The calendars we supply are ours to keep an eye on: a rename or a recolour
 * has to reach the folder row, and a deletion has to clear the binding. The
 * host cannot do either - it has no calendar API, and it could not tell a
 * deletion from our own extension restarting even if it had one.
 *
 * Name and colour are mirrored for the same reason: both outlive the calendar
 * they describe. Disabling a resource deletes the calendar, so when the user
 * enables it again the only record of what they had called it, and what colour
 * they had given it, is the one kept here. Nothing can be recovered from the
 * server - ActiveSync's folder hierarchy carries neither.
 */
function registerTargetLifecycle() {
  messenger.calendar.calendars.onUpdated.addListener(
    async (calendar, changes) => {
      if (!changes) return;
      const renamed = "name" in changes;
      const recoloured = "color" in changes;
      if (!renamed && !recoloured) return;
      const mine = (await ourTargets()).get(calendar?.id);
      if (!mine) return;

      // The platform re-announces properties it has not changed (creating a
      // calendar alone fires several), so compare before writing rather than
      // patching the folder row on every event.
      const patch = {};
      if (renamed && mine.targetName !== changes.name) {
        patch.targetName = changes.name;
      }
      if (recoloured && mine.targetColor !== changes.color) {
        patch.targetColor = changes.color;
      }
      if (!Object.keys(patch).length) return;

      await host
        ?.updateFolder({
          accountId: mine.accountId,
          folderId: mine.folderId,
          patch,
        })
        .catch((err) =>
          report?.({
            level: "warning",
            message: `[target] could not mirror ${Object.keys(patch).join(" + ")} of ${calendar?.id}: ${err?.message ?? String(err)}`,
          }),
        );
    },
  );

  messenger.calendar.calendars.onRemoved.addListener(async (id) => {
    // This fires for a calendar the user deleted - and also for every one of
    // ours when our own calendar type unregisters, which happens on each
    // reload, update or disable. The two are indistinguishable in the event,
    // so ask afterwards: a real deletion leaves nothing behind, while an
    // unregistering type leaves a force-disabled placeholder with the same
    // id. If the question cannot be answered at all we are being torn down,
    // and doing nothing is the right answer.
    let calendar;
    try {
      calendar = await messenger.calendar.calendars.get(id);
    } catch (err) {
      report?.({
        level: "debug",
        message: `[target] ${id} unreadable after onRemoved (${err?.message ?? err}); assuming our own shutdown`,
      });
      return;
    }
    if (calendar) {
      report?.({
        level: "debug",
        message: `[target] ${id} still registered after onRemoved; our calendar type is unregistering`,
      });
      return;
    }

    const mine = (await ourTargets()).get(id);
    if (!mine) {
      report?.({
        level: "debug",
        message: `[target] ${id} is gone but is not one of our targets`,
      });
      return;
    }
    report?.({
      level: "info",
      message: `[target] ${id} was deleted; clearing the binding of ${mine.accountId}/${mine.folderId}`,
    });

    // Our sync state described a calendar that no longer exists. Clear it
    // before handing the folder back, or the next bind starts from a sync key
    // the server will answer with "nothing has changed" - leaving the user a
    // calendar that stays empty. The host cannot do this: the sync key and
    // the index are ours.
    await host
      ?.updateFolder({
        accountId: mine.accountId,
        folderId: mine.folderId,
        patch: { custom: { synckey: "0", indexMap: [] } },
      })
      .catch((err) =>
        report?.({
          level: "warning",
          message: `[target] could not reset the sync state of ${mine.accountId}/${mine.folderId}: ${err?.message ?? String(err)}`,
        }),
      );
    await host?.folderTargetRemoved({ targetID: id }).catch((err) =>
      report?.({
        level: "warning",
        message: `[target] could not report the removal of ${id}: ${err?.message ?? String(err)}`,
      }),
    );
  });
}

/* Wired by background.mjs, which is the only thing that knows how to reach
 * the host. Kept as plain hooks so this module has no dependency on the
 * provider object. */
let report = null;
let host = null;

export function setSyncHandlers({ reportEventLog, provider } = {}) {
  report = reportEventLog ?? null;
  host = provider ?? null;
}
