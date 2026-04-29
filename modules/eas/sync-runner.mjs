/**
 * Generic EAS Sync framework. Drives a single-folder sync pass for any
 * item kind (Contacts / Calendar / Tasks) given an `itemKind` config:
 *
 *   {
 *     className,        "Contacts" | "Calendar" | "Tasks" - AS 2.5 Class
 *     filterType,       FilterType for the Sync Options
 *     changelogKind,    "contact" | "event" | "task" - kind for markServerWrite
 *     mapField,         folder.custom field name for the local-id → serverId map
 *     codec: {
 *       applicationDataToBlob({ adNode, serverID, asVersion, defaultTimezone, separator, uid }),
 *       appendApplicationDataFromBlob({ builder, blob, asVersion, defaultTimezone, separator }),
 *       readEasServerIdFromBlob(blob),
 *       stampEasServerId(blob, serverID),
 *     },
 *     storeFactory(targetID) → {
 *       list()                  → [{id, blob}]
 *       get(id)                 → {id, blob} | null
 *       create(id, blob)        → realId   (asserts id match in createItem flow)
 *       update(id, blob)        → void
 *       delete(id)              → void
 *     },
 *   }
 *
 * The "blob" is whatever string format the codec uses (vCard for
 * contacts, iCal for events/tasks). The runner never inspects it.
 *
 * Mirrors the legacy `EAS-4-TbSync/content/includes/sync.js` flow:
 *   1. Bootstrap synckey when "0".
 *   2. GetItemEstimate, then pull-loop with WindowSize batches.
 *   3. Push the host changelog as Add/Change/Delete with adaptive batch
 *      shrinking on collection-level Status 4/6.
 *   4. Two-second pause, then a follow-up pull.
 *
 * Recovery:
 *   - Status 3: reset synckey to "0" and retry the pass once.
 *   - Status 12: clear the account's foldersynckey and return a warning
 *     so the next account-level sync re-runs FolderSync.
 */

import ICAL from "../../vendor/ical.min.js";
import { easRequest } from "../network.mjs";
import { createWBXML } from "../wbxml.mjs";
import { readPath, readPathFrom } from "./wbxml-helpers.mjs";
import { runGetItemEstimate } from "./get-item-estimate.mjs";
import { ok, warning as warningStatus, error as errorStatus } from "../../vendor/tbsync/provider.mjs";

const STATUS_OK = "1";
const STATUS_RESYNC = "3";
const STATUS_MALFORMED = "4";
const STATUS_INVALID = "6";
const STATUS_CONFLICT = "7";
const STATUS_OBJECT_NOT_FOUND = "8";
const STATUS_FOLDER_HIERARCHY = "12";
// Server temporarily unavailable / busy. Legacy paused autosync for 30
// minutes on this; we mirror that by writing `noAutosyncUntil` on the
// account so the host's autosync ticker skips it for the duration.
const STATUS_BUSY = "110";
const BUSY_BACKOFF_MS = 30 * 60 * 1000;

const MAX_PULL_BATCHES = 50;
const POST_PUSH_WAIT_MS = 2000;

/* ── Recurrence diagnostic logging ────────────────────────────────────
 * Emit a debug-level event-log entry whenever the runner touches a
 * recurring item or processes a 16.1 per-instance exception. The full
 * iCal blob is attached as `details.ical` (or before/after pair for
 * exceptions) so the user can inspect what shape the data is in at
 * each step. Gated behind level: "debug" - production captures stay
 * clean unless the user opts in. */
function blobHasRecurrence(blob) {
  if (typeof blob !== "string") return false;
  return /\n(?:RRULE|EXDATE|RECURRENCE-ID)[;:]/.test(blob);
}

/**
 * Extended Debug log for recurrence-related events.
 */
function logRecurrence(ctx, message, details) {
  //ctx.provider.reportEventLog({
  //  level: "debug",
  //  accountId: ctx.accountId,
  //  folderId: ctx.folderId,
  //  message: `[${ctx.itemKind.changelogKind}-sync] recurrence: ${message}`,
  //  details,
  //});
}

/** Pull WindowSize + initial push batch size. Migrated from the legacy
 *  `extensions.eas4tbsync.maxitems` pref (default 50) into
 *  `browser.storage.local["maxItems"]`; default 25 when unset. */
async function readMaxItems() {
  const { maxItems } = await browser.storage.local.get({ maxItems: 25 });
  const n = Number(maxItems);
  return Number.isFinite(n) && n > 0 ? n : 25;
}

async function readMsTodoCompat() {
  const { msTodoCompat } = await browser.storage.local.get({ msTodoCompat: false });
  return msTodoCompat === true;
}

/* ── Entry point ──────────────────────────────────────────────────── */

export async function runItemSync({
  provider, account, folder, accountId, folderId, asVersion,
  itemKind, defaultTimezone,
}) {
  if (!folder.targetID) return errorStatus("No local target bound to folder");
  const collectionId = folder.custom?.serverID;
  if (!collectionId) return errorStatus("Folder is missing serverID");

  let workingFolder = folder;
  for (let attempt = 0; attempt < 2; attempt++) {
    const result = await runOneSync({
      provider, account, folder: workingFolder, accountId, folderId,
      asVersion, collectionId, itemKind, defaultTimezone,
    });
    if (result.code === "RESYNC" && attempt === 0) {
      const reset = { synckey: "0", [itemKind.mapField]: {} };
      await provider.updateFolder({
        accountId, folderId,
        patch: { custom: reset },
      });
      workingFolder = { ...workingFolder, custom: { ...(workingFolder.custom ?? {}), ...reset } };
      continue;
    }
    if (result.code === "HIERARCHY") {
      await provider.updateAccount({
        accountId,
        patch: { custom: { foldersynckey: "0" } },
      });
      return warningStatus("Folder hierarchy changed on the server - refresh the folder list and retry");
    }
    if (result.code === "BUSY") {
      // Server signalled "temporarily unavailable" (Sync Status 110).
      // Suppress autosync for 30 minutes via the host-recognized
      // top-level `noAutosyncUntil` field; the user can still trigger a
      // manual sync, which will retry the request immediately.
      await provider.updateAccount({
        accountId,
        patch: { noAutosyncUntil: Date.now() + BUSY_BACKOFF_MS },
      }).catch(() => { });
      provider.reportEventLog({
        level: "warning",
        accountId, folderId,
        message: `[${itemKind.changelogKind}-sync] server busy (Status 110); autosync paused for 30 min`,
      });
      return warningStatus("Server busy - autosync paused for 30 minutes");
    }
    return result.status ?? ok();
  }
  return errorStatus("Repeated synckey reset - giving up");
}

/* ── One full sync pass ───────────────────────────────────────────── */

async function runOneSync({
  provider, account, folder, accountId, folderId, asVersion,
  collectionId, itemKind, defaultTimezone,
}) {
  const separator = String(account.custom?.seperator ?? "10");
  let synckey = String(folder.custom?.synckey ?? "0");
  const maxItems = await readMaxItems();
  const msTodoCompat = await readMsTodoCompat();

  const ctx = {
    provider, account, accountId, folderId, folder,
    targetID: folder.targetID, collectionId, separator, asVersion,
    defaultTimezone,
    syncRecurrence: account.custom?.syncrecurrence === true,
    msTodoCompat,
    itemKind,
    store: itemKind.storeFactory(folder.targetID),
    synckey,
    idMap: { ...(folder.custom?.[itemKind.mapField] ?? {}) },
    idMapDirty: false,
    syncKeyDirty: false,
    byServerId: null,
    maxItems,
  };

  // 1) Bootstrap if needed.
  if (synckey === "0" || !synckey) {
    const boot = await sendSync({
      account, asVersion,
      body: buildSyncBody({
        synckey: "0", collectionId, asVersion, withChanges: false,
        withCommands: null, className: itemKind.className, filterType: itemKind.filterType,
      }),
    });
    if (boot.code === "RESYNC")    return await finishWith(ctx, { code: "RESYNC" });
    if (boot.code === "HIERARCHY") return await finishWith(ctx, { code: "HIERARCHY" });
    if (boot.code === "BUSY")      return await finishWith(ctx, { code: "BUSY" });
    if (boot.error)                return await finishWith(ctx, { status: errorStatus(boot.error) });
    ctx.synckey = boot.synckey;
    ctx.syncKeyDirty = true;
  }

  ctx.byServerId = await snapshotByServerId(ctx);

  // 2) Pull pass.
  const firstPull = await pullPhase(ctx);
  if (firstPull.code) return await finishWith(ctx, firstPull);

  // 3) Push pass.
  const changelog = Array.isArray(folder.changelog) ? folder.changelog : [];
  const userEdits = changelog.filter(e =>
    e?.status === "added_by_user" ||
    e?.status === "modified_by_user" ||
    e?.status === "deleted_by_user"
  );
  let pushed = { changedAnything: false };
  if (userEdits.length) {
    pushed = await pushPhase(ctx, userEdits);
    if (pushed.code) return await finishWith(ctx, pushed);
  }

  // 4) Follow-up pull, after a brief settle window.
  if (pushed.changedAnything) {
    await sleep(POST_PUSH_WAIT_MS);
    const second = await pullPhase(ctx);
    if (second.code) return await finishWith(ctx, second);
  }

  return await finishWith(ctx, { status: ok() });
}

async function finishWith(ctx, result) {
  if (!ctx.syncKeyDirty && !ctx.idMapDirty) return result;
  const patch = {};
  if (ctx.syncKeyDirty) patch.synckey = ctx.synckey;
  if (ctx.idMapDirty)   patch[ctx.itemKind.mapField] = ctx.idMap;
  try {
    await ctx.provider.updateFolder({
      accountId: ctx.accountId, folderId: ctx.folderId,
      patch: { custom: patch },
    });
  } catch (err) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId, folderId: ctx.folderId,
      message: `[${ctx.itemKind.changelogKind}-sync] flush failed: ${err?.message ?? String(err)}`,
    });
  }
  return result;
}

/* ── Pull phase ───────────────────────────────────────────────────── */

async function pullPhase(ctx) {
  const estimate = await runGetItemEstimate({
    account: ctx.account, asVersion: ctx.asVersion,
    collectionId: ctx.collectionId, synckey: ctx.synckey,
    className: ctx.itemKind.className, filterType: ctx.itemKind.filterType,
  });
  let itemsDone = 0;
  let itemsTotal = estimate ?? 0;
  reportProgress(ctx, itemsDone, itemsTotal);

  for (let batch = 0; batch < MAX_PULL_BATCHES; batch++) {
    const body = buildSyncBody({
      synckey: ctx.synckey, collectionId: ctx.collectionId,
      asVersion: ctx.asVersion, withChanges: true, withCommands: null,
      className: ctx.itemKind.className, filterType: ctx.itemKind.filterType,
      windowSize: ctx.maxItems,
    });
    const r = await sendSync({ account: ctx.account, asVersion: ctx.asVersion, body });
    if (r.code === "RESYNC") return { code: "RESYNC" };
    if (r.code === "HIERARCHY") return { code: "HIERARCHY" };
    if (r.code === "BUSY") return { code: "BUSY" };
    if (r.error) return { status: errorStatus(r.error) };

    if (r.commands) {
      const processed = await applyServerCommands(ctx, r.commands);
      itemsDone += processed;
      if (itemsDone > itemsTotal) itemsTotal = itemsDone;
      reportProgress(ctx, itemsDone, itemsTotal);
    }
    ctx.synckey = r.synckey;
    ctx.syncKeyDirty = true;
    if (!r.moreAvailable) return {};
  }
  return { status: errorStatus("MoreAvailable loop exceeded safety cap") };
}

/* ── Push phase ───────────────────────────────────────────────────── */

async function pushPhase(ctx, userEdits) {
  const failedItems = new Set();
  let batchSize = ctx.maxItems;
  let pending = userEdits.slice();
  let changedAnything = false;
  let itemsDone = 0;
  const itemsTotal = userEdits.length;
  reportProgress(ctx, itemsDone, itemsTotal);

  while (pending.length) {
    const slice = [];
    while (pending.length && slice.length < batchSize) {
      const e = pending.shift();
      if (failedItems.has(e.itemId)) continue;
      slice.push(e);
    }
    if (!slice.length) break;

    const built = await buildPushBatch(ctx, slice);
    if (!built.adds.length && !built.mods.length && !built.dels.length) {
      itemsDone += slice.length;
      reportProgress(ctx, itemsDone, itemsTotal);
      continue;
    }

    // Trace any recurring items going out so the user can correlate
    // server responses with the iCal we just sent. 16.1 sends each
    // exception as its own <Change> with the same ServerId; that
    // expansion happens inside appendInstanceChanges and isn't visible
    // here - the master mod log entry is the anchor.
    if (ctx.syncRecurrence) {
      for (const a of built.adds) {
        if (blobHasRecurrence(a.item.blob)) {
          logRecurrence(ctx, `push add: itemId=${a.item.id}, clientId=${a.clientId}`, { ical: a.item.blob });
        }
      }
      for (const m of built.mods) {
        if (blobHasRecurrence(m.item.blob)) {
          logRecurrence(ctx, `push update: itemId=${m.item.id}, serverID=${m.serverID}`, { ical: m.item.blob });
        }
      }
    }

    const r = await sendSync({
      account: ctx.account, asVersion: ctx.asVersion,
      body: buildSyncBody({
        synckey: ctx.synckey, collectionId: ctx.collectionId,
        asVersion: ctx.asVersion, withChanges: false,
        withCommands: {
          ...built, separator: ctx.separator, asVersion: ctx.asVersion,
          codec: ctx.itemKind.codec, defaultTimezone: ctx.defaultTimezone,
          syncRecurrence: ctx.syncRecurrence,
        },
        className: ctx.itemKind.className, filterType: ctx.itemKind.filterType,
      }),
    });

    if (r.code === "RESYNC") return { code: "RESYNC" };
    if (r.code === "HIERARCHY") return { code: "HIERARCHY" };
    if (r.code === "BUSY") return { code: "BUSY" };

    if (r.code === "MALFORMED") {
      if (slice.length > 1) {
        batchSize = Math.max(1, Math.floor(batchSize / 5));
        pending = slice.concat(pending);
        continue;
      }
      failedItems.add(slice[0].itemId);
      ctx.provider.reportEventLog({
        level: "warning",
        accountId: ctx.accountId, folderId: ctx.folderId,
        message: `[${ctx.itemKind.changelogKind}-sync] dropping item ${slice[0].itemId} after Status ${r.collStatus} on a single-item batch`,
      });
      itemsDone += 1;
      reportProgress(ctx, itemsDone, itemsTotal);
      continue;
    }

    if (r.error) return { status: errorStatus(r.error) };

    ctx.synckey = r.synckey;
    ctx.syncKeyDirty = true;
    changedAnything = true;

    if (r.responses) await applyResponses(ctx, r.responses, built);
    if (r.commands) await applyServerCommands(ctx, r.commands);

    itemsDone += slice.length;
    reportProgress(ctx, itemsDone, itemsTotal);
  }
  return { changedAnything };
}

async function buildPushBatch(ctx, slice) {
  const adds = [];
  const mods = [];
  const dels = [];
  for (const entry of slice) {
    if (entry.status === "added_by_user") {
      const it = await ctx.store.get(entry.itemId);
      if (!it?.blob) continue;
      const clientId = `c-${Date.now().toString(36)}-${adds.length}`;
      adds.push({ entry, clientId, item: it });
    } else if (entry.status === "modified_by_user") {
      const it = await ctx.store.get(entry.itemId);
      if (!it?.blob) continue;
      const serverID = ctx.itemKind.codec.readEasServerIdFromBlob(it.blob)
                    ?? ctx.idMap[entry.itemId];
      if (!serverID) continue;
      mods.push({ entry, serverID, item: it });
    } else if (entry.status === "deleted_by_user") {
      const serverID = ctx.idMap[entry.itemId];
      if (!serverID) {
        await ctx.provider.changelogRemove({
          accountId: ctx.accountId, folderId: ctx.folderId,
          parentId: entry.parentId, itemId: entry.itemId,
        });
        continue;
      }
      dels.push({ entry, serverID });
    }
  }
  return { adds, mods, dels };
}

/* ── Apply responses to our push ──────────────────────────────────── */

async function applyResponses(ctx, responses, sent) {
  for (const node of responses.adds) {
    const clientId = readPathFrom(node, ["ClientId"]);
    const serverId = readPathFrom(node, ["ServerId"]);
    const status   = readPathFrom(node, ["Status"]);
    const sentEntry = sent.adds.find(a => a.clientId === clientId);
    if (!sentEntry) continue;
    if (status !== STATUS_OK || !serverId) continue;
    const stamped = ctx.itemKind.codec.stampEasServerId(sentEntry.item.blob, serverId);
    await ctx.provider.changelogMarkServerWrite({
      accountId: ctx.accountId, folderId: ctx.folderId,
      parentId: ctx.targetID, itemId: sentEntry.item.id,
      status: "modified_by_server", kind: ctx.itemKind.changelogKind,
    });
    await ctx.store.update(sentEntry.item.id, stamped);
    mergeIdMap(ctx, { [sentEntry.item.id]: serverId });
    await ctx.provider.changelogRemove({
      accountId: ctx.accountId, folderId: ctx.folderId,
      parentId: sentEntry.entry.parentId, itemId: sentEntry.entry.itemId,
    });
  }
  for (const node of responses.changes) {
    const status = readPathFrom(node, ["Status"]);
    if (status === STATUS_OK) continue;
    if (status === STATUS_CONFLICT || status === STATUS_OBJECT_NOT_FOUND) {
      const serverId = readPathFrom(node, ["ServerId"]);
      const sentEntry = sent.mods.find(m => m.serverID === serverId);
      if (sentEntry) {
        await ctx.provider.changelogRemove({
          accountId: ctx.accountId, folderId: ctx.folderId,
          parentId: sentEntry.entry.parentId, itemId: sentEntry.entry.itemId,
        });
      }
    }
  }
  for (const node of responses.deletes) {
    const status = readPathFrom(node, ["Status"]);
    const serverId = readPathFrom(node, ["ServerId"]);
    const sentEntry = sent.dels.find(d => d.serverID === serverId);
    if (!sentEntry) continue;
    if (status === STATUS_OK || status === STATUS_OBJECT_NOT_FOUND) {
      await ctx.provider.changelogRemove({
        accountId: ctx.accountId, folderId: ctx.folderId,
        parentId: sentEntry.entry.parentId, itemId: sentEntry.entry.itemId,
      });
      removeIdMap(ctx, sentEntry.entry.itemId);
    }
  }
  for (const m of sent.mods) {
    const ack = responses.changes.find(n => readPathFrom(n, ["ServerId"]) === m.serverID);
    const status = ack ? readPathFrom(ack, ["Status"]) : STATUS_OK;
    if (!ack || status === STATUS_OK) {
      await ctx.provider.changelogRemove({
        accountId: ctx.accountId, folderId: ctx.folderId,
        parentId: m.entry.parentId, itemId: m.entry.itemId,
      });
    }
  }
  for (const d of sent.dels) {
    const ack = responses.deletes.find(n => readPathFrom(n, ["ServerId"]) === d.serverID);
    if (!ack) {
      await ctx.provider.changelogRemove({
        accountId: ctx.accountId, folderId: ctx.folderId,
        parentId: d.entry.parentId, itemId: d.entry.itemId,
      });
      removeIdMap(ctx, d.entry.itemId);
    }
  }
}

/* ── Apply server commands ───────────────────────────────────────── */

async function applyServerCommands(ctx, commands) {
  let processed = 0;
  for (const node of commands.adds)        { await applyAdd(ctx, node);    processed++; }
  for (const node of commands.changes)     { await applyChange(ctx, node); processed++; }
  for (const node of commands.deletes)     { await applyDelete(ctx, node); processed++; }
  for (const node of commands.softDeletes) { await applyDelete(ctx, node); processed++; }
  return processed;
}

async function applyAdd(ctx, addNode) {
  const serverID = readPathFrom(addNode, ["ServerId"]);
  if (!serverID) return;
  const ad = childByTag(addNode, "ApplicationData");
  if (!ad) return;
  const existing = ctx.byServerId.get(serverID);
  if (existing) return applyChangeFromAd(ctx, ad, existing);

  const newId = crypto.randomUUID();
  const blob = await ctx.itemKind.codec.applicationDataToBlob({
    adNode: ad, serverID, asVersion: ctx.asVersion,
    separator: ctx.separator, defaultTimezone: ctx.defaultTimezone,
    syncRecurrence: ctx.syncRecurrence,
    msTodoCompat: ctx.msTodoCompat,
    uid: newId,
  });
  await ctx.provider.changelogMarkServerWrite({
    accountId: ctx.accountId, folderId: ctx.folderId,
    parentId: ctx.targetID, itemId: newId,
    status: "added_by_server", kind: ctx.itemKind.changelogKind,
  });
  const createdId = await ctx.store.create(newId, blob);
  if (createdId !== newId) {
    throw new Error(`store.create id mismatch: expected ${newId}, got ${createdId}`);
  }
  await verifyRoundTrip(ctx, newId, blob, "create");
  ctx.byServerId.set(serverID, { itemId: newId, blob });
  mergeIdMap(ctx, { [newId]: serverID });
  if (blobHasRecurrence(blob)) {
    logRecurrence(ctx, `pull add: itemId=${newId}, serverID=${serverID}`, { ical: blob });
  }
}

async function applyChange(ctx, changeNode) {
  const serverID = readPathFrom(changeNode, ["ServerId"]);
  if (!serverID) return;
  const ad = childByTag(changeNode, "ApplicationData");
  if (!ad) return;
  const existing = ctx.byServerId.get(serverID);
  if (!existing) return applyAdd(ctx, changeNode);
  // 16.1 per-instance Change: ApplicationData carries <InstanceId> and
  // is scoped to a single occurrence of the master event referenced by
  // ServerId. Route to the codec's exception path; bail back to the
  // normal master update if the codec doesn't support it.
  const instanceId = readPathFrom(ad, ["InstanceId"]);
  if (instanceId && ctx.syncRecurrence && ctx.itemKind.codec.applyInstanceChange) {
    return applyExceptionChange(ctx, ad, existing, instanceId);
  }
  return applyChangeFromAd(ctx, ad, existing);
}

async function applyExceptionChange(ctx, ad, existing, instanceId) {
  const instanceUtc = parseEasInstanceId(instanceId);
  if (!instanceUtc) return applyChangeFromAd(ctx, ad, existing);

  const deleted = readPathFrom(ad, ["Deleted"]) === "1";
  const codec = ctx.itemKind.codec;
  let nextBlob;
  if (deleted) {
    nextBlob = codec.applyInstanceDelete?.({ ical: existing.blob, instanceUtc });
  } else {
    nextBlob = codec.applyInstanceChange?.({
      ical: existing.blob,
      adNode: ad,
      instanceUtc,
      asVersion: ctx.asVersion,
      defaultTimezone: ctx.defaultTimezone,
    });
  }
  if (!nextBlob || nextBlob === existing.blob) {
    logRecurrence(ctx, `pull 16.1 exception ${deleted ? "delete" : "change"} no-op: itemId=${existing.itemId}, instance=${instanceId}`, {
      ical: existing.blob,
    });
    return;
  }

  await ctx.provider.changelogMarkServerWrite({
    accountId: ctx.accountId, folderId: ctx.folderId,
    parentId: ctx.targetID, itemId: existing.itemId,
    status: "modified_by_server", kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.update(existing.itemId, nextBlob);
  await verifyRoundTrip(ctx, existing.itemId, nextBlob, deleted ? "exception-delete" : "exception-update");
  logRecurrence(ctx, `pull 16.1 exception ${deleted ? "delete" : "change"} applied: itemId=${existing.itemId}, instance=${instanceId}`, {
    before: existing.blob,
    after: nextBlob,
  });
  // Re-key by the master's serverId (unchanged); keep the new blob.
  const masterServerId = codec.readEasServerIdFromBlob(nextBlob)
                       ?? codec.readEasServerIdFromBlob(existing.blob);
  if (masterServerId) {
    ctx.byServerId.set(masterServerId, { itemId: existing.itemId, blob: nextBlob });
  }
}

/** Parse an EAS InstanceId string ("YYYYMMDDTHHMMSSZ" or extended-ISO)
 *  into a JS Date. EAS encodes the original master occurrence in UTC. */
function parseEasInstanceId(s) {
  if (!s) return null;
  const compact = String(s).replace(/[-:]/g, "");
  const m = /^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?Z?$/.exec(compact);
  if (!m) return null;
  const [, y, mo, d, h = "0", mi = "0", se = "0"] = m;
  return new Date(Date.UTC(+y, +mo - 1, +d, +h, +mi, +se));
}

async function applyChangeFromAd(ctx, ad, existing) {
  const blob = await ctx.itemKind.codec.applicationDataToBlob({
    adNode: ad,
    serverID: ctx.itemKind.codec.readEasServerIdFromBlob(existing.blob),
    asVersion: ctx.asVersion,
    separator: ctx.separator,
    defaultTimezone: ctx.defaultTimezone,
    syncRecurrence: ctx.syncRecurrence,
    msTodoCompat: ctx.msTodoCompat,
    uid: existing.itemId,
  });
  await ctx.provider.changelogMarkServerWrite({
    accountId: ctx.accountId, folderId: ctx.folderId,
    parentId: ctx.targetID, itemId: existing.itemId,
    status: "modified_by_server", kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.update(existing.itemId, blob);
  await verifyRoundTrip(ctx, existing.itemId, blob, "update");
  ctx.byServerId.set(ctx.itemKind.codec.readEasServerIdFromBlob(existing.blob),
                     { itemId: existing.itemId, blob });
  if (blobHasRecurrence(blob) || blobHasRecurrence(existing.blob)) {
    logRecurrence(ctx, `pull update: itemId=${existing.itemId}`, {
      before: existing.blob,
      after: blob,
    });
  }
}

async function applyDelete(ctx, delNode) {
  const serverID = readPathFrom(delNode, ["ServerId"]);
  if (!serverID) return;
  const existing = ctx.byServerId.get(serverID);
  if (!existing) return;
  await ctx.provider.changelogMarkServerWrite({
    accountId: ctx.accountId, folderId: ctx.folderId,
    parentId: ctx.targetID, itemId: existing.itemId,
    status: "deleted_by_server", kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.delete(existing.itemId);
  ctx.byServerId.delete(serverID);
  removeIdMap(ctx, existing.itemId);
}

/* ── Sync request building ────────────────────────────────────────── */

function buildSyncBody({
  synckey, collectionId, asVersion, withChanges, withCommands,
  className, filterType, windowSize,
}) {
  const w = createWBXML();
  w.switchpage("AirSync");
  w.otag("Sync");
    w.otag("Collections");
      w.otag("Collection");
        if (asVersion === "2.5") w.atag("Class", className);
        w.atag("SyncKey", synckey);
        w.atag("CollectionId", collectionId);
        if (withChanges) {
          w.atag("DeletesAsMoves");
          w.atag("GetChanges");
          w.atag("WindowSize", String(windowSize ?? 25));
          if (asVersion !== "2.5") {
            w.otag("Options");
              w.atag("FilterType", String(filterType));
              w.switchpage("AirSyncBase");
              w.otag("BodyPreference");
                w.atag("Type", "1");
              w.ctag();
              w.switchpage("AirSync");
            w.ctag();
          }
        }
        if (withCommands) appendCommands(w, withCommands);
      w.ctag();
    w.ctag();
  w.ctag();
  return w.getBytes();
}

function appendCommands(w, { adds, mods, dels, separator, asVersion, codec, defaultTimezone, syncRecurrence }) {
  if (!adds.length && !mods.length && !dels.length) return;
  w.otag("Commands");
  for (const a of adds) {
    w.otag("Add");
      w.atag("ClientId", a.clientId);
      w.otag("ApplicationData");
        codec.appendApplicationDataFromBlob({
          builder: w, blob: a.item.blob, asVersion, separator, defaultTimezone, syncRecurrence,
        });
        w.switchpage("AirSync");
      w.ctag();
    w.ctag();
  }
  for (const m of mods) {
    w.otag("Change");
      w.atag("ServerId", m.serverID);
      w.otag("ApplicationData");
        codec.appendApplicationDataFromBlob({
          builder: w, blob: m.item.blob, asVersion, separator, defaultTimezone, syncRecurrence,
        });
        w.switchpage("AirSync");
      w.ctag();
    w.ctag();
    // 16.1 sends each modified/deleted exception as its own <Change>
    // with the master's ServerId. Codec opts in via `appendInstanceChanges`;
    // contacts and tasks don't expose it.
    if (syncRecurrence && asVersion === "16.1" && codec.appendInstanceChanges) {
      codec.appendInstanceChanges({
        builder: w, blob: m.item.blob, serverID: m.serverID,
        asVersion, defaultTimezone, syncRecurrence,
      });
      // codec leaves the builder on the AirSync codepage when it returns.
    }
  }
  for (const d of dels) {
    w.otag("Delete");
      w.atag("ServerId", d.serverID);
    w.ctag();
  }
  w.ctag();
}

/* ── Sync response parsing ────────────────────────────────────────── */

async function sendSync({ account, asVersion, body }) {
  const { doc } = await easRequest({ account, command: "Sync", body, asVersion });
  if (!doc) return { error: "Empty Sync response" };
  return parseSyncResponse(doc);
}

function parseSyncResponse(doc) {
  const top = readPath(doc, ["Status"]);
  if (top && top !== STATUS_OK) {
    if (top === STATUS_BUSY) return { code: "BUSY", topStatus: top };
    return { error: `Sync top status ${top}` };
  }
  const collection = doc.getElementsByTagName("Collection")[0];
  if (!collection) return { error: "Sync response missing Collection" };

  const collStatus = readPathFrom(collection, ["Status"]) ?? STATUS_OK;
  if (collStatus !== STATUS_OK) {
    if (collStatus === STATUS_RESYNC)           return { code: "RESYNC", collStatus };
    if (collStatus === STATUS_FOLDER_HIERARCHY) return { code: "HIERARCHY", collStatus };
    if (collStatus === STATUS_BUSY)             return { code: "BUSY", collStatus };
    if (collStatus === STATUS_MALFORMED || collStatus === STATUS_INVALID) {
      return { code: "MALFORMED", collStatus };
    }
    return { error: `Sync collection status ${collStatus}`, collStatus };
  }

  const synckey = readPathFrom(collection, ["SyncKey"]) ?? "0";
  const moreAvailable = !!childByTag(collection, "MoreAvailable");

  let commands = null;
  const cmdNode = childByTag(collection, "Commands");
  if (cmdNode) {
    commands = {
      adds:        Array.from(cmdNode.children).filter(c => c.tagName === "Add"),
      changes:     Array.from(cmdNode.children).filter(c => c.tagName === "Change"),
      deletes:     Array.from(cmdNode.children).filter(c => c.tagName === "Delete"),
      softDeletes: Array.from(cmdNode.children).filter(c => c.tagName === "SoftDelete"),
    };
  }

  let responses = null;
  const respNode = childByTag(collection, "Responses");
  if (respNode) {
    responses = {
      adds:    Array.from(respNode.children).filter(c => c.tagName === "Add"),
      changes: Array.from(respNode.children).filter(c => c.tagName === "Change"),
      deletes: Array.from(respNode.children).filter(c => c.tagName === "Delete"),
    };
  }

  return { synckey, moreAvailable, commands, responses };
}

/* ── Helpers ──────────────────────────────────────────────────────── */

async function snapshotByServerId(ctx) {
  const all = await ctx.store.list();
  const map = new Map();
  for (const it of all) {
    const sid = ctx.itemKind.codec.readEasServerIdFromBlob(it.blob);
    if (sid) map.set(sid, { itemId: it.id, blob: it.blob });
  }
  return map;
}

function mergeIdMap(ctx, additions) {
  for (const [k, v] of Object.entries(additions)) ctx.idMap[k] = v;
  ctx.idMapDirty = true;
}

function removeIdMap(ctx, itemId) {
  if (!(itemId in ctx.idMap)) return;
  delete ctx.idMap[itemId];
  ctx.idMapDirty = true;
}

function reportProgress(ctx, itemsDone, itemsTotal) {
  ctx.provider.reportProgress({
    accountId: ctx.accountId, folderId: ctx.folderId,
    itemsDone, itemsTotal,
  });
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function childByTag(node, tag) {
  if (!node?.children) return null;
  for (const c of node.children) if (c.tagName === tag) return c;
  return null;
}

/** Read the just-written item back through the store and compare it
 *  semantically to what we passed in. Property order in iCal/vCard is
 *  not significant (RFC 5545 / 6350), and Thunderbird rewrites a few
 *  envelope-level properties (PRODID, VERSION) plus reorders parameters,
 *  so we parse both sides, normalize each property to a canonical
 *  string, and diff the resulting multisets at the inner-component
 *  level (VEVENT / VTODO / VCARD). We only log properties that were
 *  dropped, added, or whose value/parameters changed. Soft-fails on
 *  read or parse errors. */
async function verifyRoundTrip(ctx, itemId, expected, op) {
  const kind = ctx.itemKind.changelogKind;
  const target = roundTripTargetFor(kind);
  if (!target) return;
  let actual = null;
  try {
    const got = await ctx.store.get(itemId);
    actual = got?.blob ?? null;
  } catch (err) {
    ctx.provider.reportEventLog({
      level: "debug",
      accountId: ctx.accountId, folderId: ctx.folderId,
      message: `[${kind}-sync] roundtrip readback (${op}) failed for ${itemId}: ${err?.message ?? String(err)}`,
    });
    return;
  }
  if (actual === expected) return;
  if (actual == null) {
    ctx.provider.reportEventLog({
      level: "debug",
      accountId: ctx.accountId, folderId: ctx.folderId,
      message: `[${kind}-sync] roundtrip readback (${op}) returned no item for ${itemId}`,
    });
    return;
  }

  const diff = diffComponentProperties(expected, actual, target);
  if (!diff) return;
  if (!diff.dropped.length && !diff.added.length && !diff.changed.length) return;

  const lines = [`[${kind}-sync] roundtrip mismatch on ${op} of ${itemId}`];
  if (diff.dropped.length) lines.push("dropped: " + diff.dropped.join(" | "));
  if (diff.added.length)   lines.push("added:   " + diff.added.join(" | "));
  if (diff.changed.length) lines.push("changed: " + diff.changed.join(" | "));
  ctx.provider.reportEventLog({
    level: "debug",
    accountId: ctx.accountId, folderId: ctx.folderId,
    message: lines.join("\n"),
  });
}

/** Resolve the kind to a parse target: which top-level component to
 *  parse and which subcomponent (if any) to compare. Returns null for
 *  kinds that don't have a known structured form. */
function roundTripTargetFor(kind) {
  if (kind === "event")   return { outer: "vcalendar", inner: "vevent" };
  if (kind === "task")    return { outer: "vcalendar", inner: "vtodo" };
  if (kind === "contact") return { outer: "vcard",     inner: null };
  return null;
}

/** Parse two iCal/vCard strings and return an order-insensitive diff
 *  between their inner-component properties. Returns null if either
 *  side fails to parse. */
function diffComponentProperties(expectedStr, actualStr, target) {
  const e = innerProps(expectedStr, target);
  const a = innerProps(actualStr, target);
  if (!e || !a) return null;

  const dropped = [];
  const added   = [];
  const changed = [];

  // Group both sides by property name. For each name compare sorted
  // canonical strings so multi-occurrence props (CATEGORIES, ATTENDEE)
  // diff cleanly without caring about order.
  const names = new Set([...e.keys(), ...a.keys()]);
  for (const name of names) {
    const eList = (e.get(name) ?? []).slice().sort();
    const aList = (a.get(name) ?? []).slice().sort();
    if (eList.length === 0 && aList.length > 0) {
      added.push(...aList);
      continue;
    }
    if (aList.length === 0 && eList.length > 0) {
      dropped.push(...eList);
      continue;
    }
    if (eList.length === aList.length && eList.every((s, i) => s === aList[i])) continue;
    // Same name, different content: report as changed.
    changed.push(`${name}: ${eList.join(",")} → ${aList.join(",")}`);
  }
  return { dropped, added, changed };
}

function innerProps(text, target) {
  let comp;
  try { comp = new ICAL.Component(ICAL.parse(text)); }
  catch { return null; }
  // For iCal we descend into VEVENT/VTODO; for vCard the top-level
  // component itself holds the properties (no inner wrapper).
  const inner = target.inner ? comp.getFirstSubcomponent(target.inner) : comp;
  if (!inner) return null;
  const map = new Map();
  for (const p of inner.getAllProperties()) {
    const name = p.name;
    const line = canonicalPropertyString(p);
    if (!map.has(name)) map.set(name, []);
    map.get(name).push(line);
  }
  return map;
}

/** Render a property in a parameter-order-independent canonical form so
 *  TB's parameter reordering (which is also legal per RFC 5545) doesn't
 *  trigger a diff. Falls back to toICALString() if the structured form
 *  isn't available. */
function canonicalPropertyString(prop) {
  try {
    const j = prop.toJSON();   // [name, paramsObj, valueType, ...values]
    const name = j[0];
    const params = j[1] ?? {};
    const valueType = j[2];
    const values = j.slice(3);
    const paramKeys = Object.keys(params).sort();
    const paramStr = paramKeys
      .map(k => `;${k.toUpperCase()}=${stringifyValue(params[k])}`)
      .join("");
    const valStr = values.map(stringifyValue).join(",");
    return `${name.toUpperCase()}${paramStr}${valueType ? "" : ""}:${valStr}`;
  } catch {
    try { return prop.toICALString(); }
    catch { return `${prop.name}:${stringifyValue(prop.getFirstValue())}`; }
  }
}

function stringifyValue(v) {
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (Array.isArray(v)) return v.map(stringifyValue).join(",");
  if (typeof v === "object" && typeof v.toString === "function") return v.toString();
  return String(v);
}
