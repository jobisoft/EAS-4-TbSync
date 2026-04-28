/**
 * EAS calendar / task sync entry points. Wires the Calendar and Tasks
 * itemKinds into the shared `runItemSync` framework. Codecs live in
 * `calendar-codec.mjs` and `task-codec.mjs`; local store reads/writes
 * go through `calendar-store.mjs`.
 */

import { runItemSync } from "./sync-runner.mjs";
import * as calendarStore from "../calendar-store.mjs";
import * as eventCodec from "./calendar-codec.mjs";
import * as taskCodec  from "./task-codec.mjs";

function makeCodec(modCodec) {
  return {
    applicationDataToBlob({ adNode, serverID, asVersion, defaultTimezone, syncRecurrence, msTodoCompat, uid }) {
      return modCodec.applicationDataToIcal({ adNode, serverID, asVersion, defaultTimezone, syncRecurrence, msTodoCompat, uid });
    },
    appendApplicationDataFromBlob({ builder, blob, asVersion, defaultTimezone, syncRecurrence }) {
      return modCodec.appendApplicationDataFromIcal({ builder, ical: blob, asVersion, defaultTimezone, syncRecurrence });
    },
    readEasServerIdFromBlob: modCodec.readEasServerIdFromIcal,
    stampEasServerId: modCodec.stampEasServerId,
    // Optional 16.1 exception methods (calendar only). The runner falls
    // back to a normal master-update for codecs that don't implement them.
    applyInstanceChange: modCodec.applyInstanceChange,
    applyInstanceDelete: modCodec.applyInstanceDelete,
    appendInstanceChanges: modCodec.appendInstanceChanges,
  };
}

function calendarStoreFactory(targetID, type) {
  return {
    async list() {
      const all = await calendarStore.listItems(targetID, type);
      return all.map(it => ({ id: it.id, blob: it.item }));
    },
    async get(id) {
      const it = await calendarStore.getItem(targetID, id);
      return it ? { id: it.id, blob: it.item } : null;
    },
    async create(id, blob) {
      const created = await calendarStore.createItem(targetID, { id, type, ical: blob });
      return created.id;
    },
    async update(id, blob) {
      await calendarStore.updateItem(targetID, id, { ical: blob });
    },
    async delete(id) {
      await calendarStore.deleteItem(targetID, id);
    },
  };
}

const calendarItemKind = {
  className: "Calendar",
  filterType: "0",
  changelogKind: "event",
  mapField: "itemMap",
  codec: makeCodec(eventCodec),
  storeFactory: targetID => calendarStoreFactory(targetID, "event"),
};

const taskItemKind = {
  className: "Tasks",
  filterType: "0",
  changelogKind: "task",
  mapField: "itemMap",
  codec: makeCodec(taskCodec),
  storeFactory: targetID => calendarStoreFactory(targetID, "task"),
};

async function getDefaultTimezone() {
  try {
    const z = await messenger.calendar.timezones.getCurrent();
    return z?.id ?? "UTC";
  } catch {
    return "UTC";
  }
}

export async function syncCalendarFolder({ provider, account, folder, accountId, folderId, asVersion }) {
  const filterType = String(account.custom?.synclimit ?? "7");
  const defaultTimezone = await getDefaultTimezone();
  return runItemSync({
    provider, account, folder, accountId, folderId, asVersion,
    itemKind: { ...calendarItemKind, filterType },
    defaultTimezone,
  });
}

export async function syncTaskFolder({ provider, account, folder, accountId, folderId, asVersion }) {
  const defaultTimezone = await getDefaultTimezone();
  return runItemSync({
    provider, account, folder, accountId, folderId, asVersion,
    itemKind: taskItemKind,
    defaultTimezone,
  });
}
