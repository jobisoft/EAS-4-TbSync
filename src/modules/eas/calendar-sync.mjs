/**
 * EAS calendar / task sync entry points. Wires the Calendar and Tasks
 * itemKinds into the shared `runItemSync` framework. Codecs live in
 * `calendar-codec.mjs` and `task-codec.mjs`; local store reads/writes
 * go through the vendored `calendar.mjs`.
 */

import { runItemSync } from "./sync-runner.mjs";
import * as calendarStore from "../../vendor/tbsync/calendar.mjs";
import * as eventCodec from "./calendar-codec.mjs";
import * as taskCodec from "./task-codec.mjs";
import { ensureLoaded as ensureTimezoneMappingLoaded } from "./timezone-mapping.mjs";

function makeCodec(modCodec) {
  return {
    applicationDataToBlob({
      adNode,
      existingBlob,
      serverID,
      asVersion,
      defaultTimezone,
      syncRecurrence,
      msTodoCompat,
      uid,
      userEmail,
      eventLog,
      nativePlainText,
    }) {
      return modCodec.applicationDataToIcal({
        adNode,
        existingIcal: existingBlob,
        serverID,
        asVersion,
        defaultTimezone,
        syncRecurrence,
        msTodoCompat,
        uid,
        userEmail,
        eventLog,
        nativePlainText,
      });
    },
    appendApplicationDataFromBlob({
      builder,
      blob,
      asVersion,
      defaultTimezone,
      syncRecurrence,
      suppressExceptions,
      userEmail,
      fallbackOrganizerName,
      eventLog,
    }) {
      return modCodec.appendApplicationDataFromIcal({
        builder,
        ical: blob,
        asVersion,
        defaultTimezone,
        syncRecurrence,
        suppressExceptions,
        userEmail,
        fallbackOrganizerName,
        eventLog,
      });
    },
    readEasServerIdFromBlob: modCodec.readEasServerIdFromIcal,
    stampEasServerId: modCodec.stampEasServerId,
    // Optional: names why an item cannot go on the wire without changing
    // its meaning - the runner then holds it as a client-side rejection.
    clientRejectReason: modCodec.clientRejectReason,
    serverRejectReason: modCodec.serverRejectReason,
    // Optional 16.1 exception methods (calendar only). The runner falls
    // back to a normal master-update for codecs that don't implement them.
    applyInstanceChange: modCodec.applyInstanceChange,
    applyInstanceDelete: modCodec.applyInstanceDelete,
    listInstanceCommands: modCodec.listInstanceCommands,
  };
}

/** Every write the sync makes goes to the storage behind the calendar, not
 *  the calendar itself. That is what keeps server data out of the user-edit
 *  hooks: `<id>#cache` fires no `calendar.provider.onItem*`, while the
 *  calendar shows the result immediately because both share one id.
 *
 *  Reads use the same handle for symmetry - it is the same items either
 *  way, and one id in this factory is one thing to get wrong instead of
 *  two. */
function calendarStoreFactory(targetID, type) {
  const writeID = calendarStore.cacheId(targetID);
  return {
    async list() {
      const all = await calendarStore.listItems(writeID, type);
      return all.map((it) => ({ id: it.id, blob: it.item }));
    },
    async get(id) {
      const it = await calendarStore.getItem(writeID, id);
      return it ? { id: it.id, blob: it.item } : null;
    },
    async create(id, blob) {
      const created = await calendarStore.createItem(writeID, {
        id,
        type,
        ical: blob,
      });
      return created.id;
    },
    async update(id, blob) {
      await calendarStore.updateItem(writeID, id, { ical: blob });
    },
    async delete(id) {
      await calendarStore.deleteItem(writeID, id);
    },
  };
}

export const calendarItemKind = {
  className: "Calendar",
  filterType: "0",
  changelogKind: "event",
  codec: makeCodec(eventCodec),
  storeFactory: (targetID) => calendarStoreFactory(targetID, "event"),
};

export const taskItemKind = {
  className: "Tasks",
  filterType: "0",
  changelogKind: "task",
  codec: makeCodec(taskCodec),
  storeFactory: (targetID) => calendarStoreFactory(targetID, "task"),
};

async function getDefaultTimezone() {
  try {
    return (await messenger.calendar.timezones.currentZone) || "UTC";
  } catch (err) {
    console.debug(
      "[eas] getDefaultTimezone: messenger.calendar.timezones.currentZone failed:",
      err,
    );
    return "UTC";
  }
}

/** Run `action` with the local TB calendar temporarily writable, then
 *  restore the folder's effective read-only state once the action
 *  resolves. The toggle removes the sync write path's dependence on
 *  any privileged "bypass readOnly" affordance: writes always go to a
 *  writable calendar, and the user-facing read-only state only
 *  re-engages after the sync settles. Any user edit that lands during
 *  the brief writable window is captured by the calendar observer and
 *  reverted on the next sync via the runner's `revertLocalChanges`
 *  pass (sync-runner.mjs).
 *
 *  The restore runs in a `finally` so a thrown sync error doesn't
 *  leave the calendar mis-flagged. The restore itself is best-effort
 *  (any error is logged and swallowed) so the action's outcome — the
 *  thing the caller actually cares about — propagates unchanged. */
async function withWritableCalendar(folder, action) {
  if (!folder?.targetID) return action();
  const finalRO = !!folder.readOnly || !!folder.downloadOnly;
  try {
    await calendarStore.setCalendarReadOnly(folder.targetID, false);
    return await action();
  } finally {
    try {
      await calendarStore.setCalendarReadOnly(folder.targetID, finalRO);
    } catch (err) {
      console.debug(
        `[eas] withWritableCalendar: failed to restore readOnly on ${folder.targetID}:`,
        err,
      );
    }
  }
}

/** The calendar window this account asked for, as the `FilterType` of the
 *  Options block.
 *
 *  Shared rather than derived twice: the block is sticky per collection
 *  and a new one REPLACES it, so any request that states options and gets
 *  this wrong silently resets the window for every request after it. The
 *  item kind's own `filterType` is the tasks/contacts answer and is not
 *  it - see the `<Options>` unit tests. */
export function calendarFilterType(account) {
  return String(account?.custom?.synclimit ?? "7");
}

export async function syncCalendarFolder({
  provider,
  account,
  folder,
  accountId,
  folderId,
  asVersion,
}) {
  const filterType = calendarFilterType(account);
  const defaultTimezone = await getDefaultTimezone();
  await ensureTimezoneMappingLoaded((level, message) =>
    provider.reportEventLog({ level, accountId, folderId, message }),
  );
  return withWritableCalendar(folder, () =>
    runItemSync({
      provider,
      account,
      folder,
      accountId,
      folderId,
      asVersion,
      itemKind: { ...calendarItemKind, filterType },
      defaultTimezone,
    }),
  );
}

export async function syncTaskFolder({
  provider,
  account,
  folder,
  accountId,
  folderId,
  asVersion,
}) {
  const defaultTimezone = await getDefaultTimezone();
  await ensureTimezoneMappingLoaded((level, message) =>
    provider.reportEventLog({ level, accountId, folderId, message }),
  );
  return withWritableCalendar(folder, () =>
    runItemSync({
      provider,
      account,
      folder,
      accountId,
      folderId,
      asVersion,
      itemKind: taskItemKind,
      defaultTimezone,
    }),
  );
}
