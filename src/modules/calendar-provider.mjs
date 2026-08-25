/**
 * EAS as a calendar provider.
 *
 * A provider calendar keeps the two directions apart as separate objects:
 *
 *   user edits    → `calendar.provider.onItem{Created,Updated,Removed}`
 *   our own sync  → written to `<id>#cache`, which fires none of those
 *
 * The separation is structural rather than tagged, and `onItemUpdated` is
 * handed the *previous* item, so which occurrence of a series moved is
 * visible rather than guessed at.
 *
 * ## Why the hooks do not push
 *
 * A hook fires the instant the user hits save and the platform waits for an
 * answer. Pushing to Exchange there would make every edit a network round
 * trip and would fail the edit outright when the connection is down. So the
 * item is accepted and the scheduled sync pushes it.
 *
 * ## Where the record goes
 *
 * Into our own storage, via `change-queue.mjs`. The hook is holding the
 * user's save, so the record must be durable before we answer, and it must
 * not be conditional on anything outside this add-on being alive: these
 * calendars keep working with the host absent, so a record that needed the
 * host would be unmakeable on every host reload, update and suspend.
 *
 * The host gets a count for its needs-sync badge, best-effort. If it does
 * not arrive, a dot is missing until the next sync; the edit is safe.
 *
 * Address books have no provider API, so they are watched instead - see
 * the vendored `address-book.mjs`. Same queue, same sessions; only the way an edit
 * is noticed differs.
 */

import ICAL from "../vendor/ical.min.js";
import {
  announceableOf,
  differingPropertyNames,
  exceptionFingerprint,
  isReceivedMeeting,
  pinEasStamps,
} from "./eas/calendar-codec.mjs";
import { clientSchedulingFor } from "./eas/client-scheduling.mjs";
import {
  localQueue,
  lookupBinding,
  rememberBindings,
} from "../vendor/tbsync/change-queue.mjs";

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
      // Every resource is banked, because the address-book observer needs
      // its bindings for exactly the same reason the item hooks need
      // theirs. Only calendars go into the map this function returns.
      bindings.push({
        targetID: f.targetID,
        accountId,
        folderId: f.folderId,
        sessionId: f.sessionId ?? null,
        targetType: f.targetType,
      });
      if (f.targetType !== "calendars" && f.targetType !== "tasks") continue;
      out.set(f.targetID, {
        accountId,
        folderId: f.folderId,
        sessionId: f.sessionId ?? null,
        targetName: f.targetName ?? null,
        targetColor: f.targetColor ?? null,
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

/**
 * Note that this edit owes the attendees a message, on the accounts where
 * sending it is our job.
 *
 * A separate row from the queued edit, in its own family, because it
 * answers a different question and outlives the push that clears the edit -
 * the message cannot go until the pull has settled what the server actually
 * kept. See `changelog-core.mjs` for why the family is part of a row's
 * identity.
 *
 * This is the only place the item's previous version exists, which is why
 * the decision is made here rather than at send time. `from` is the meeting
 * as the attendees currently hold it; the phase that sends compares the
 * settled item against it and stays quiet if nothing they care about moved.
 *
 * A create records no `from` - nobody has been told the meeting exists, so
 * there is nothing to have changed from, and the invitation carries whatever
 * the item finally says. A delete records nothing at all: a deletion is
 * never announced, and the shared updaters drop any note it had.
 *
 * Best-effort throughout. A note we fail to write costs one notification;
 * failing the user's save over it would cost the edit.
 */
async function noteSendMail({ binding, calendarId, itemId, kind, op, oldIcal, ical }) {
  if (kind !== "event" || op === "deleted") return;
  try {
    const scheduling = await clientSchedulingFor(calendarId);
    if (!scheduling) return;

    // Ours to send only if we are the organiser. An invitation somebody
    // else sent us is answered, never re-announced - that is item 46's job
    // and it runs from the item, not from a note.
    if (isReceivedMeeting(ical, scheduling.user)) return;

    const now = announceableOf(ical);
    const from = op === "created" ? null : announceableOf(oldIcal);
    // Nobody to tell. Both sides are asked because an edit that removes the
    // last attendee still owes them a cancellation.
    if (!now?.attendees?.length && !from?.attendees?.length) return;

    await localQueue(binding).recordSendMail({
      parentId: calendarId,
      itemId,
      kind,
      status: op === "created" ? "added_for_sendMail" : "modified_for_sendMail",
      ...(from ? { detail: { from } } : {}),
    });
    report?.({
      level: "debug",
      message:
        `[event-sync] ${itemId} owes its attendees a message ` +
        `(${op === "created" ? "invitation" : "update"})`,
    });
  } catch (err) {
    report?.({
      level: "warning",
      message:
        `[event-sync] could not note the message owed for ${itemId}: ` +
        `${err?.message ?? String(err)}`,
    });
  }
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
async function record(
  calendarId,
  itemId,
  op,
  { type, oldIcal, ical, flags } = {},
) {
  if (!calendarId || !itemId) {
    // Was a silent return, and that silence hid a data-loss bug: a
    // UI-created item arrives with no id, so every one of them was
    // dropped here without a trace. `identify` gives a create its id
    // before this point, so reaching it now means either the platform
    // named no calendar or minting failed - and for a create that is an
    // edit nobody will ever push, the same thing the storage failure
    // below calls an error.
    report?.({
      level: op === "created" ? "error" : "warning",
      message:
        `[${type === "task" ? "task" : "event"}-sync] cannot record user ` +
        `${op}: no ${calendarId ? "usable item id" : "calendar id"}` +
        (op === "created" ? " - this edit will never reach the server" : ""),
    });
    return;
  }
  const before = oldIcal ? exceptionFingerprint(oldIcal) : null;
  // One calendar type serves both, so the item says which it is - a task
  // folder's edits must not be queued as events.
  const kind = type === "task" ? "task" : "event";

  // Which folder this calendar belongs to, and which binding of it is
  // current. Normally already banked by a sync, so this is one storage read
  // and no round trip.
  //
  // A miss means we have never held this folder's row - a resource enabled
  // since our last sync of it - and the host is the only place that answer
  // exists, so it is worth one attempt. If the host cannot answer either,
  // there is nothing to file against: inventing a binding would push the
  // edit into whatever folder the id turns out to belong to.
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
      `[${kind}-sync] queued user ${op} of ${itemId}${flags ?? ""}` +
      (before
        ? ` (${before.exdates.length} cancelled, ${before.overrides.length} override(s) before the edit)`
        : ""),
  });

  await noteSendMail({ binding, calendarId, itemId, kind, op, oldIcal, ical });

  // The host paints a needs-sync badge and cannot count a queue it does not
  // hold. Best-effort on purpose: the edit is already safe, and a badge
  // that lags until the next sync is not worth failing a save over.
  await host
    ?.updateFolder({
      accountId: binding.accountId,
      folderId: binding.folderId,
      patch: { localChanges: pending },
    })
    .catch((err) =>
      report?.({
        level: "debug",
        message: `[${kind}-sync] could not update the pending count: ${err?.message ?? String(err)}`,
      }),
    );

  armSyncAfterChange(
    calendarId,
    await quietSecondsFor(binding.accountId),
    kind,
  );
}

/** The window an account that predates the setting keeps, and what an
 *  unreadable one falls back to. Matches the provider's
 *  `DEFAULT_SYNC_ON_CHANGE`. */
const DEFAULT_QUIET_SECONDS = 15;

/** Pending sync per calendar, keyed by target id, so edits in one calendar
 *  never postpone another's. */
const syncAfterChangeTimers = new Map();

/** How long this account wants a calendar to stay quiet before its edits
 *  are synced, in seconds, `0` meaning never. Read rather than remembered:
 *  `record` is already asking the host to store the pending count, so this
 *  is a second call beside one we make anyway, and it means a change to the
 *  setting takes effect on the very next edit. */
async function quietSecondsFor(accountId) {
  try {
    const { account } = (await host?.getAccount(accountId)) ?? {};
    const stored = Number(account?.custom?.syncOnChange);
    return Number.isFinite(stored) ? stored : DEFAULT_QUIET_SECONDS;
  } catch {
    return DEFAULT_QUIET_SECONDS;
  }
}

/**
 * Sync a calendar shortly after the user stops editing it.
 *
 * Without this an edit waits for the next scheduled sync, which since the
 * calendars declare `scheduling: "server"` is also how long an invitation
 * waits: Thunderbird no longer mails the attendees, the server does, and
 * the server only learns of the meeting when we push it.
 *
 * Armed from `record` and nowhere else, which is what keeps it from feeding
 * itself: our own sync writes go to `<id>#cache` and raise no item hook, so
 * every call here is a person editing something.
 *
 * Re-arming on each edit rather than syncing on the first is the whole
 * point - a burst of writes is one sync at the end of it.
 *
 * Scoped to the calendar that changed. `requestSync` takes a target id and
 * the host resolves it to that one folder, so a change in one calendar does
 * not drag the account's other resources along.
 *
 * Nothing here can fail a save: the edit is already durable by the time
 * this is called, and the worst case is that it waits for the scheduled
 * sync exactly as it did before this existed.
 */
function armSyncAfterChange(targetID, seconds, kind) {
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

let registered = false;

/** Which of the platform's flags this write arrived with, for the log.
 *  `invitation` means iTIP processing is writing the item - the user
 *  accepted an emailed invitation - and `offline` that this is a replay
 *  of something queued while Thunderbird was offline. Neither changes
 *  what we do: a replayed edit is still a real edit, and responding to
 *  an invitation properly is MeetingResponse, which we do not implement
 *  yet. They are logged because an edit that arrived either way is
 *  otherwise indistinguishable from a plain one. */
function flagsOf(hookOptions) {
  const flags = Object.entries(hookOptions ?? {})
    .filter(([, on]) => on)
    .map(([name]) => name);
  return flags.length ? ` (${flags.join(", ")})` : "";
}

/** Give a newly created item an id, and hand it back so the platform
 *  adopts ours.
 *
 *  Thunderbird decides an edit is an *addition* precisely by the absence
 *  of an id, and the experiment mints one only after this hook has run.
 *  So every item created in Thunderbird's own dialog reaches us with
 *  `id: null` - and an id is exactly what our queue files against, since
 *  we record now and push later. Left alone, such an item was dropped
 *  and never synced.
 *
 *  The platform rebuilds the item from returned props whenever they
 *  carry a `type`, so setting the UID in both the props and the iCal is
 *  what makes its own fallback id never happen and keeps the two in
 *  agreement. A provider that pushed inline would instead return the
 *  item the server named; this is the same mechanism.
 *
 *  Returns the props to hand back, or null when the item already has an
 *  id and nothing needs saying. */
export function identify(item) {
  if (!item || item.id) return null;
  const uid = crypto.randomUUID();
  try {
    const vcal = new ICAL.Component(ICAL.parse(item.item));
    // EVERY component, not just the master. A recurring item arrives with
    // its modified occurrences beside it, and they are all id-less
    // together - the platform clears the id on the whole series at once.
    // Stamping only the first would leave the overrides orphaned, which
    // is both an iCal violation (RFC 5545 binds an override to its master
    // by UID) and a rejected save: the platform refuses to rebuild an
    // exception whose id does not match its parent's.
    const comps = vcal
      .getAllSubcomponents()
      .filter((c) => c.name === "vevent" || c.name === "vtodo");
    if (!comps.length) return null;
    for (const comp of comps) comp.updatePropertyWithValue("uid", uid);
    return { ...item, id: uid, item: vcal.toString() };
  } catch {
    // A blob we cannot parse is not worth failing the user's save over -
    // the caller keeps the item as it came and logs the refusal.
    return null;
  }
}

/** The version of this item the calendar already holds, as iCal, or null.
 *
 *  An item's id is its iCal UID, so this is the UID lookup - one local
 *  read, no network, and no item hook fires for it. Called while the
 *  platform holds the user's save, which is why it stays a single get.
 *
 *  Best-effort: a lookup we cannot do leaves `prior` null, and the caller
 *  then treats the item as new. That errs towards stripping an identity
 *  rather than inventing one, which is the safer of the two. */
async function priorIcalOf(calendarId, item) {
  if (!calendarId || !item?.id) return null;
  try {
    const existing = await messenger.calendar.items.get(calendarId, item.id, {
      returnFormat: "ical",
    });
    return existing?.item ?? null;
  } catch {
    return null;
  }
}

/** Hold our own stamps to what the calendar already had - see
 *  `pinEasStamps`. Returns the item the platform should store: the same
 *  object when nothing needed doing, so the common path allocates nothing.
 *
 *  Logged when it bites, at info: something outside this add-on wrote to a
 *  field only we should write, and this line is the only trace of it. */
async function guardStamps(item, priorIcal) {
  const ical = item?.item;
  if (typeof ical !== "string") return item;
  const guarded = pinEasStamps({ builtIcal: ical, priorIcal });
  if (guarded === ical) return item;
  report?.({
    level: "info",
    message:
      `[event-sync] restored the EAS stamps on ${item.id}: the incoming ` +
      `item ${priorIcal ? "did not carry the ones it was stored with" : "carried stamps it has no claim to"}` +
      alsoChanges(ical, priorIcal),
  });
  return { ...item, item: guarded };
}

/** How many property names the line above will print before it stops. A
 *  rewrite that touches everything says so in a word instead of listing
 *  the whole item. */
const CHANGED_NAMES_SHOWN = 8;

/** What this write changes besides our stamps, named for the log.
 *
 *  The question the line exists to answer: who wrote this? A name is
 *  enough to tell an alarm being acknowledged (`x-moz-lastack`) from
 *  something rewriting recurrence (`rrule`) - and "nothing else" is the
 *  most useful answer of all, because a write that changes nothing but
 *  the stamps we just restored is one we should not be pushing.
 *
 *  Never throws. It runs inside the item hook, which holds the user's
 *  save: a hook that does not return the item makes the platform treat
 *  the edit as failed and the user's change disappears. Diagnostics are
 *  not worth that, so anything unexpected here degrades to saying
 *  nothing. */
function alsoChanges(ical, priorIcal) {
  if (!priorIcal) return "";
  try {
    const names = differingPropertyNames(priorIcal, ical);
    if (names == null) return "";
    if (!names.length) return "; nothing else differs from the stored copy";
    const shown = names.slice(0, CHANGED_NAMES_SHOWN).join(", ");
    const rest = names.length - CHANGED_NAMES_SHOWN;
    return `; it also changes ${shown}${rest > 0 ? ` and ${rest} more` : ""}`;
  } catch (err) {
    console.debug("[eas] could not diff an incoming item:", err);
    return "";
  }
}

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

  provider.onItemCreated.addListener(async (calendar, item, hookOptions) => {
    // An item created in Thunderbird's dialog has no id yet; give it one
    // and hand the props back, so the platform adopts our id instead of
    // minting its own after we have already filed the queue entry.
    const base = identify(item) ?? item;
    // An id is a UID here, so "is this already in the calendar?" is one
    // read. It usually is not - but an import of an item we already sync
    // arrives as a create, and the platform's own importer expects the
    // calendar to refuse a duplicate id, which ours does not. Without this
    // the item would be overwritten by a copy that has no identity, and a
    // synced event would quietly detach from the server.
    const prior = await priorIcalOf(calendar?.id, base);
    const guarded = await guardStamps(base, prior);
    await record(calendar?.id, guarded?.id ?? item?.id, "created", {
      type: item?.type,
      ical: guarded?.item ?? item?.item,
      flags: flagsOf(hookOptions),
    });
    return guarded === item ? { item } : guarded;
  }, options);

  provider.onItemUpdated.addListener(
    async (calendar, item, oldItem, hookOptions) => {
      const oldIcal = oldItem?.item ?? null;
      const guarded = await guardStamps(item, oldIcal);
      await record(calendar?.id, item?.id, "updated", {
        type: item?.type,
        oldIcal,
        ical: guarded?.item ?? item?.item,
        flags: flagsOf(hookOptions),
      });
      return guarded === item ? { item } : guarded;
    },
    options,
  );

  provider.onItemRemoved.addListener(async (calendar, item, hookOptions) => {
    await record(calendar?.id, item?.id, "deleted", {
      type: item?.type,
      flags: flagsOf(hookOptions),
    });
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
