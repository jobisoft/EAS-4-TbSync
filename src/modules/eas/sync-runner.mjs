/**
 * Generic EAS Sync framework. Drives a single-folder sync pass for any
 * item kind (Contacts / Calendar / Tasks) given an `itemKind` config:
 *
 *   {
 *     className,        "Contacts" | "Calendar" | "Tasks" - AS 2.5 Class
 *     filterType,       FilterType for the Sync Options
 *     changelogKind,    "contact" | "event" | "task" - kind for markServerWrite
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
import { fetchServerItem } from "./item-operations.mjs";
import { easCommandLikelyAvailable } from "./allowed-commands.mjs";
import {
  localQueue,
  rememberBindings,
} from "../../vendor/tbsync/change-queue.mjs";
import {
  ok,
  warning as warningStatus,
  error as errorStatus,
  accountRerun,
} from "../../vendor/tbsync/provider.mjs";

const STATUS_OK = "1";
const STATUS_RESYNC = "3";
const STATUS_MALFORMED = "4";
const STATUS_TEMP_SERVER = "5"; // Temporary server issues / invalid item - soft fail
const STATUS_INVALID = "6";
// Server's changes win. Legacy treated this as a silent OK; the instance
// phase instead absorbs the changes the reply carries and tries once more.
const STATUS_CONFLICT = "7";
const STATUS_OBJECT_NOT_FOUND = "8";
const STATUS_FOLDER_HIERARCHY = "12";
// Top-level only: "something on the server caused a retriable error", whose
// documented resolution is "resend the request" ([MS-ASCMD] §2.2.3.177.17).
const STATUS_RETRY = "16";
// Server temporarily unavailable / busy. Legacy paused autosync for 30
// minutes on this; we mirror that by writing `noAutosyncUntil` on the
// account so the host's autosync ticker skips it for the duration.
const STATUS_BUSY = "110";
const BUSY_BACKOFF_MS = 30 * 60 * 1000;

// Top-level Sync.Status codes that indicate a malformed wire-level
// payload ([MS-ASCMD] §2.2.3.167.16). Distinct from collection-level
// MALFORMED (4/6) because the server rejected the *transport* shape, not
// per-item data - usually fatal until a code change ships.
const STATUS_PROTOCOL_FAULT = new Set(["101", "102", "103"]);

// Top-level "client / device / user not allowed" codes. Legacy bundled
// these as `global.clientdenied`; the new code surfaces each with a
// short human-readable label localized via the provider's _locales (see
// `eas.sync.error.accessDenied.<code>`) so the user has a starting point
// for the fix (admin disabled the device, user has no mailbox, …). See
// [MS-ASCMD] §2.2.3.167.16.
const STATUS_ACCESS_DENIED = new Set([
  "109",
  "112",
  "126",
  "127",
  "128",
  "129",
  "130",
  "131",
]);

// Server demands re-Provisioning. 141 = DeviceNotProvisionable,
// 142 = DeviceNotProvisioned, 143 = PolicyRefresh, 144 = InvalidPolicyKey.
// In-band variant of HTTP 449 - same recovery: mark the account as needing
// Provision and let the host re-run the account sync.
const STATUS_PROVISION_REQUIRED = new Set(["141", "142", "143", "144"]);

// How often one instance command may be re-sent before it is reported. Per
// command, not per pass: each occurrence is an independent request, and one
// exception exhausting its attempts says nothing about the next.
const MAX_INSTANCE_RETRIES = 6;

// Statuses worth re-sending an instance command for, and how long to wait
// first. Membership decides whether, the value decides how long, so there is
// no second list to keep in step with this one.
//
// Everything absent is reported on the first reply. Status 6 in particular is
// documented as "the client has sent a malformed or invalid item ... this is
// not a transient condition" - asking again cannot make it true, and six
// pointless round trips would only delay the report.
//
// The delays split on whether the reply told us anything. A 5 or a 16 carries
// no usable body, so nothing on our side has changed and an instant resend is
// the same request a moment later. A 7 arrives with the server's newer copy of
// the master, which `applyServerCommands` has already applied by the time the
// status is judged, so the next attempt is against different data and waiting
// would add nothing.
const INSTANCE_RETRY_DELAY_MS = Object.freeze({
  [STATUS_TEMP_SERVER]: 1000,
  [STATUS_CONFLICT]: 0,
  [STATUS_RETRY]: 1000,
});

/* ── DEV: fixture injection ─────────────────────────────────────────────
 *
 * Set DEV_FIXTURE_ADD_XML to inject it as part of every inbound Sync.
 * The runner processes it just like a real server-pushed Add.
 */
const DEV_FIXTURE_ADD_XML = null;
// `<Add xmlns="AirSync">
//   <ServerId>DEV-FIXTURE-RECURRENCE-BUG</ServerId>
//   <ApplicationData>
//     <AllDayEvent xmlns="Calendar">0</AllDayEvent>
//     <TimeZone xmlns="Calendar">xP///1cALgAgAEUAdQByAG8AcABlACAAUwB0AGEAbgBkAGEAcgBkACAAVABpAG0AZQAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAFAAMAAAAAAAAAAAAAACgAVQBUAEMAKwAwADEAOgAwADAAKQAgAEEAbQBzAHQAZQByAGQAYQBtACwAIABCAGUAcgBsAGkAbgAsACAAQgAAAAMAAAAFAAIAAAAAAAAAxP///w==</TimeZone>
//     <DtStamp xmlns="Calendar">20260309T131533Z</DtStamp>
//     <StartTime xmlns="Calendar">20260304T133000Z</StartTime>
//     <Subject xmlns="Calendar">Kind (dev fixture)</Subject>
//     <UID xmlns="Calendar">dev-fixture-recurrence-bug</UID>
//     <OrganizerName xmlns="Calendar">Organisator</OrganizerName>
//     <OrganizerEmail xmlns="Calendar">org@acme.com</OrganizerEmail>
//     <Attendees xmlns="Calendar"/>
//     <Location xmlns="AirSyncBase"/>
//     <EndTime xmlns="Calendar">20260304T140000Z</EndTime>
//     <Recurrence xmlns="Calendar">
//       <Type>1</Type><Interval>1</Interval><DayOfWeek>62</DayOfWeek><FirstDayOfWeek>0</FirstDayOfWeek>
//     </Recurrence>
//     <Body xmlns="AirSyncBase"><Type>1</Type><EstimatedDataSize>2</EstimatedDataSize><Data>%0D%0A</Data></Body>
//     <Sensitivity xmlns="Calendar">0</Sensitivity>
//     <BusyStatus xmlns="Calendar">2</BusyStatus>
//     <Reminder xmlns="Calendar"/>
//     <MeetingStatus xmlns="Calendar">0</MeetingStatus>
//     <NativeBodyType xmlns="AirSyncBase">1</NativeBodyType>
//     <ResponseRequested xmlns="Calendar">1</ResponseRequested>
//     <ResponseType xmlns="Calendar">1</ResponseType>
//   </ApplicationData>
// </Add>`;

function parseDevFixtureAdd() {
  try {
    const doc = new DOMParser().parseFromString(
      DEV_FIXTURE_ADD_XML,
      "application/xml",
    );
    return doc.documentElement;
  } catch {
    return null;
  }
}

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
  const { msTodoCompat } = await browser.storage.local.get({
    msTodoCompat: false,
  });
  return msTodoCompat === true;
}

/* ── Entry point ──────────────────────────────────────────────────── */

export async function runItemSync({
  provider,
  account,
  folder,
  accountId,
  folderId,
  asVersion,
  itemKind,
  defaultTimezone,
}) {
  if (!folder.targetID) return errorStatus("No local target bound to folder");
  const collectionId = folder.custom?.serverID;
  if (!collectionId) return errorStatus("Folder is missing serverID");

  // Bank which folder this target belongs to while we are holding the row.
  // An item hook is handed a calendar id and nothing else and must be able
  // to file the user's edit without asking anyone - so the answer has to be
  // here before it is needed, and a sync is when we legitimately have it.
  if (folder.sessionId) {
    await rememberBindings([
      {
        targetID: folder.targetID,
        accountId,
        folderId,
        sessionId: folder.sessionId,
        targetType: folder.targetType,
      },
    ]).catch((err) =>
      console.debug("[eas] could not bank the folder binding:", err),
    );
  }

  let workingFolder = folder;
  for (let attempt = 0; attempt < 2; attempt++) {
    const result = await runOneSync({
      provider,
      account,
      folder: workingFolder,
      accountId,
      folderId,
      asVersion,
      collectionId,
      itemKind,
      defaultTimezone,
    });
    if (result.code === "RESYNC" && attempt === 0) {
      const reset = { synckey: "0", indexMap: [] };
      await provider.updateFolder({
        accountId,
        folderId,
        patch: { custom: reset },
      });
      workingFolder = {
        ...workingFolder,
        custom: { ...(workingFolder.custom ?? {}), ...reset },
      };
      continue;
    }
    if (result.code === "HIERARCHY") {
      // Server signalled Sync Status 12: the FolderSync state we're
      // working against is stale. Reset foldersynckey so the next
      // FolderSync starts from "0", and signal an account-level rerun -
      // matches legacy `resyncAccount` for `Sync.12` (sync.js:751-753).
      // Host's sync-coordinator caps the rerun count for loop protection.
      await provider.updateAccount({
        accountId,
        patch: { custom: { foldersynckey: "0" } },
      });
      return accountRerun(
        "Folder hierarchy changed on the server - rerunning account sync",
      );
    }
    if (result.code === "PROVISION_REQUIRED") {
      // Server signalled Sync Status 141/142/143/144 in-band - the same
      // condition HTTP 449 surfaces out-of-band. Mark the account as
      // needing Provision; the next account rerun will hit the gate at
      // eas-provider.mjs::#doConnectAndDiscover and refresh the policy
      // before resuming FolderSync. Matches legacy `resyncAccount` for
      // `Sync.14X` (network.js:793-800). Host caps rerun count.
      await provider.updateAccount({
        accountId,
        patch: { custom: { provision: true, policykey: "0" } },
      });
      return accountRerun(
        "Server demands re-provisioning - rerunning account sync",
      );
    }
    if (result.code === "BUSY") {
      // Server signalled "temporarily unavailable" (Sync Status 110).
      // Suppress autosync for 30 minutes via the host-recognized
      // top-level `noAutosyncUntil` field; the user can still trigger a
      // manual sync, which will retry the request immediately.
      await provider
        .updateAccount({
          accountId,
          patch: { noAutosyncUntil: Date.now() + BUSY_BACKOFF_MS },
        })
        .catch((err) =>
          console.debug(
            `[eas] updateAccount(noAutosyncUntil) for ${accountId} failed:`,
            err,
          ),
        );
      provider.reportEventLog({
        level: "warning",
        accountId,
        folderId,
        message: `[${itemKind.changelogKind}-sync] server busy (Status 110); autosync paused for 30 min`,
      });
      return warningStatus("Server busy - autosync paused for 30 minutes");
    }
    return result.status ?? ok();
  }
  return errorStatus("Repeated synckey reset - giving up");
}

/* ── One full sync pass ───────────────────────────────────────────── */

/** The account's stated conflict preference, sent as `<Conflict>` in every
 *  Options block: "1" = the server's copy wins (the default, and what both
 *  Exchange and Z-Push do when nothing is sent - stated explicitly so we
 *  rely on a declared contract instead of a default), "0" = this device
 *  wins. Anything else in storage falls back to "1". */
function conflictPolicyOf(account) {
  return account?.custom?.conflict === "0" ? "0" : "1";
}

/** The queue this folder's pending edits live in.
 *
 *  Always ours, whatever the resource: we are handed a calendar edit
 *  directly and we watch our own address books, so in both cases the record
 *  is made here and never depends on the host being up.
 *
 *  `observed` is the one difference. A book fires an event for every write
 *  including ours, so our writes have to be announced with a pre-tag first;
 *  a calendar we supply fires nothing for them, so a tag there would never
 *  be consumed. The queue reads it off the folder's kind. */
function makeQueue({ folder, accountId, folderId }) {
  return localQueue({
    accountId,
    folderId,
    sessionId: folder.sessionId,
    observed: folder.targetType === "contacts",
  });
}

async function runOneSync({
  provider,
  account,
  folder,
  accountId,
  folderId,
  asVersion,
  collectionId,
  itemKind,
  defaultTimezone,
}) {
  const separator = String(account.custom?.seperator ?? "10");
  let synckey = String(folder.custom?.synckey ?? "0");
  const maxItems = await readMaxItems();
  const msTodoCompat = await readMsTodoCompat();
  // Effective read-only: server-imposed (`folder.readOnly`) OR user-toggled
  // (`folder.downloadOnly`). When set, we discard pending user-side edits
  // before pulling, and skip the push phase entirely. Matches legacy's
  // `revertLocalChanges` + downloadonly gate (legacy sync.js:349-378).
  const effectiveDownloadOnly = !!folder.readOnly || !!folder.downloadOnly;

  const ctx = {
    provider,
    account,
    accountId,
    folderId,
    folder,
    targetID: folder.targetID,
    collectionId,
    separator,
    asVersion,
    defaultTimezone,
    syncRecurrence: account.custom?.syncrecurrence === true,
    msTodoCompat,
    conflict: conflictPolicyOf(account),
    itemKind,
    store: itemKind.storeFactory(folder.targetID),
    synckey,
    // Pre-bound info-level event-log emitter the codec calls when it
    // converts/drops a VALARM that EAS can't represent, or to surface
    // ORGANIZER round-trip diagnostics. Built once here so every codec
    // call site can pass `eventLog: ctx.eventLog` instead of inlining
    // the closure.
    eventLog: (level, message) =>
      provider.reportEventLog({ level, accountId, folderId, message }),
    // Single per-folder index of `{uid, serverId}` pairs. The
    // upgrades.mjs drain runs before any sync RPC, so by the time we
    // get here the persisted shape is guaranteed to be an array (or
    // missing for a never-synced folder).
    indexMap: Array.isArray(folder.custom?.indexMap)
      ? folder.custom.indexMap.slice()
      : [],
    indexMapDirty: false,
    // Lazily-built `serverId -> itemId` view of the stored blobs, standing
    // behind the indexMap when it cannot answer. Null until something
    // misses; see `serverIdScan`. Per-pass, so the RESYNC retry - which
    // builds a fresh ctx - never reuses a stale one.
    serverIdScan: null,
    syncKeyDirty: false,
    maxItems,
    // Where this folder's pending edits live. A calendar we supply keeps
    // them in our own storage, namespaced by the binding the host names; an
    // address book keeps them in the host's folder row, because the host is
    // what observes it. Same five calls either way, so nothing below has to
    // know which folder it is working on.
    queue: makeQueue({ folder, accountId, folderId }),
  };
  // The queue as this sync found it. Kept because one pull-side decision
  // needs to ask about it synchronously - see `hasPendingUserDelete`.
  ctx.pendingAtStart = await ctx.queue.pending();

  // Read-only revert pre-step. Drops local edits before the pull so the
  // local store ends up matching the server. ItemOperations.Fetch lets us
  // re-pull a single item by serverId; falls back to a synckey reset when
  // the server didn't advertise ItemOperations (legacy behaviour at
  // sync.js:888-911 `revertLocalChangesViaResync`).
  if (effectiveDownloadOnly) {
    const heavyResetNeeded = await revertLocalChanges(ctx);
    if (heavyResetNeeded) {
      ctx.synckey = "0";
      synckey = "0";
      ctx.syncKeyDirty = true;
      ctx.indexMap = [];
      ctx.indexMapDirty = true;
    }
  }

  // 1) Bootstrap if needed.
  if (synckey === "0" || !synckey) {
    const boot = await sendSync({
      account,
      asVersion,
      body: buildSyncBody({
        synckey: "0",
        collectionId,
        asVersion,
        withChanges: false,
        withCommands: null,
        className: itemKind.className,
        filterType: itemKind.filterType,
      }),
    });
    if (boot.code === "RESYNC")
      return await finishWith(ctx, { code: "RESYNC" });
    if (boot.code === "HIERARCHY")
      return await finishWith(ctx, { code: "HIERARCHY" });
    if (boot.code === "PROVISION_REQUIRED")
      return await finishWith(ctx, { code: "PROVISION_REQUIRED" });
    if (boot.code === "BUSY") return await finishWith(ctx, { code: "BUSY" });
    if (boot.error)
      return await finishWith(ctx, { status: errorStatus(boot.error) });
    ctx.synckey = boot.synckey;
    ctx.syncKeyDirty = true;
  }

  // 2) Push pass, BEFORE the pull - the order [MS-ASCMD] shapes a Sync
  // around (client Commands are applied before GetChanges is answered).
  // Pulling first meant the pull could apply a server <Change> over a
  // pending local edit and the push then re-sent the server's own data;
  // with the push first, the pending edit reaches the server before
  // anything can overwrite it, and a genuine two-writer conflict is
  // decided by the only party that knows both sides - the server, per
  // the <Conflict> preference every request now states.
  //
  // Skipped entirely on a read-only folder. Any edits the user made
  // between the revert above and this point will be re-reverted on the
  // next sync; the runner is the single authority for upsync gating.
  let pushed = {};
  if (!effectiveDownloadOnly) {
    const pending = await ctx.queue.pending();
    // ActiveSync has no mailing list: [MS-ASCNTC] describes a contact and
    // nothing else, so a Thunderbird list in a synced book has nowhere to go.
    // The host watches address books generically and queues one anyway, and
    // handing it to the contact push path means `contacts.get` on a list id -
    // which throws, fails the whole folder, and keeps failing until the list
    // is deleted. Drop it here instead, and say so once, exactly as the
    // legacy provider did.
    //
    // Every changelog row carries a `kind`, so this needs no fallback.
    const userEdits = [];
    for (const e of pending) {
      if (e.kind !== ctx.itemKind.changelogKind) {
        ctx.provider.reportEventLog({
          level: "warning",
          accountId: ctx.accountId,
          folderId: ctx.folderId,
          message:
            `[${ctx.itemKind.changelogKind}-sync] skipping a ${e.kind} ` +
            `("${e.itemId}"): ActiveSync cannot store one, so it stays local`,
        });
        await ctx.queue.remove({
          parentId: e.parentId,
          itemId: e.itemId,
          kind: e.kind,
        });
        continue;
      }
      userEdits.push(e);
    }
    if (userEdits.length) {
      pushed = await pushPhase(ctx, userEdits);
      if (pushed.code) return await finishWith(ctx, pushed);
    }
  }

  // 3) Exceptions of the recurring masters this push sent. Each keys on
  // the master's ServerId and cannot share a request with it, so they
  // need their own pass - see instancePhase. Ahead of the pull below,
  // so that pull sees the server's finished state.
  let instanceFailed = 0;
  if (pushed.instanceMasters?.length) {
    const inst = await instancePhase(ctx, pushed.instanceMasters);
    if (inst.code) return await finishWith(ctx, inst);
    if (inst.status) return await finishWith(ctx, inst);
    instanceFailed = inst.failedCount ?? 0;
  }

  // 4) Pull pass. Running after the push makes everything it delivers
  // server-authoritative by construction: our own changes are already part
  // of the state it reports, so there is no echo to defend against and no
  // window in which it could overwrite a pending local edit. When a push
  // lost a conflict (Status 7), this is also where the server's winning
  // copy arrives - the losing edit is visibly replaced, never silently
  // dropped.
  const pull = await pullPhase(ctx);
  if (pull.code) return await finishWith(ctx, pull);

  // Final status — `warning` if push rejected any items so the user sees
  // a count in the manager; `ok` otherwise. The warning's text comes
  // from a localized message with a `$1` substitution for the count.
  // Mirrors legacy's `ServerRejectedSomeItems::N` warning at sync.js:721-723.
  // Instance-change rejections count towards the same total: an exception
  // the server refused is an element it did not accept, and without this
  // the folder reported a clean sync while the occurrence was missing
  // server-side.
  const failedCount = (pushed.failedCount ?? 0) + instanceFailed;
  if (failedCount > 0) {
    const msg = browser.i18n.getMessage(
      "eas.sync.warning.serverRejectedSomeItems",
      [String(failedCount)],
    );
    return await finishWith(ctx, { status: warningStatus(msg) });
  }
  return await finishWith(ctx, { status: ok() });
}

/* ── Read-only revert ─────────────────────────────────────────────────
 *
 * Walk the folder's user-side changelog and undo each entry so the local
 * store ends up matching the server. Called from `runOneSync` only when
 * the folder is effectively read-only (`folder.readOnly` from the server
 * or `folder.downloadOnly` toggled by the user).
 *
 * Per-entry semantics (mirrors legacy sync.js:730-886):
 *   - added_by_user      → delete the local item; drop changelog entry.
 *   - modified_by_user   → ItemOperations.Fetch the server's copy and
 *                          overwrite local; drop changelog entry. If the
 *                          fetch comes back empty (server says it's gone),
 *                          delete the local item too.
 *   - deleted_by_user    → ItemOperations.Fetch and re-create locally;
 *                          drop changelog entry. If the fetch returns no
 *                          item, just drop the entry — server agrees the
 *                          item is gone.
 *
 * Fallback when the server didn't advertise ItemOperations: delete every
 * `added_by_user` item locally, drop the entire changelog, and signal to
 * the caller that a synckey reset is needed. The next pull (with
 * SyncKey=0) re-emits every server item as an Add; existing-locally items
 * route through `applyChangeFromAd` which overwrites with the server
 * version, locally-deleted items get re-added by `applyAdd`.
 *
 * Returns `true` when the caller should reset `ctx.synckey` to "0"
 * (heavy fallback path), `false` otherwise.
 */
async function revertLocalChanges(ctx) {
  const userEdits = await ctx.queue.pending();
  if (userEdits.length === 0) return false;

  const codec = ctx.itemKind.codec;
  const canFetch = easCommandLikelyAvailable(ctx.account, "ItemOperations");

  if (!canFetch) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[${ctx.itemKind.changelogKind}-sync] read-only revert: server lacks ItemOperations; resetting synckey for ${userEdits.length} pending edit(s)`,
    });
    for (const e of userEdits) {
      if (e.status === "added_by_user") {
        try {
          await ctx.store.delete(e.itemId);
        } catch {
          /* item already gone — fine */
        }
      }
      await ctx.queue.remove({
        parentId: e.parentId,
        itemId: e.itemId,
        kind: e.kind,
      });
    }
    return true;
  }

  for (const e of userEdits) {
    // Before touching anything: this loop deletes local items and drops
    // changelog entries, so being interrupted half-way is the one place a
    // cancel could cost the user work. `fetchServerItem` rethrows a cancel
    // rather than reporting "gone", and this stops the loop entering the
    // next item at all.
    throwIfCancelled(ctx);
    if (e.status === "added_by_user") {
      try {
        await ctx.store.delete(e.itemId);
      } catch {
        /* already gone */
      }
      await ctx.queue.remove({
        parentId: e.parentId,
        itemId: e.itemId,
        kind: e.kind,
      });
      continue;
    }

    let serverID = findServerIdByUid(ctx, e.itemId);
    if (!serverID && e.status === "modified_by_user") {
      // For modify, the local blob still has the server-stamped ID.
      try {
        const it = await ctx.store.get(e.itemId);
        if (it?.blob) serverID = codec.readEasServerIdFromBlob(it.blob);
      } catch {
        /* fall through */
      }
    }
    if (!serverID) {
      // Nothing to fetch — drop the entry and let the regular sync
      // settle the local state.
      await ctx.queue.remove({
        parentId: e.parentId,
        itemId: e.itemId,
        kind: e.kind,
      });
      continue;
    }

    const properties = await fetchServerItem({
      account: ctx.account,
      asVersion: ctx.asVersion,
      collectionId: ctx.collectionId,
      serverID,
    });

    if (!properties) {
      // Fetch failed or server says gone. For a modify, this means the
      // server deleted the item out from under us; delete the local copy
      // to converge. For a delete, the server already agrees.
      if (e.status === "modified_by_user") {
        try {
          await ctx.store.delete(e.itemId);
        } catch {
          /* already gone */
        }
      }
      await ctx.queue.remove({
        parentId: e.parentId,
        itemId: e.itemId,
        kind: e.kind,
      });
      continue;
    }

    const blob = await codec.applicationDataToBlob({
      adNode: properties,
      serverID,
      asVersion: ctx.asVersion,
      separator: ctx.separator,
      defaultTimezone: ctx.defaultTimezone,
      syncRecurrence: ctx.syncRecurrence,
      msTodoCompat: ctx.msTodoCompat,
      uid: e.itemId,
      userEmail: ctx.account?.custom?.user,
      eventLog: ctx.eventLog,
    });

    await ctx.queue.markServerWrite({
      parentId: ctx.targetID,
      itemId: e.itemId,
      status: "modified_by_server",
      kind: ctx.itemKind.changelogKind,
    });

    if (e.status === "modified_by_user") {
      await ctx.store.update(e.itemId, blob);
    } else {
      // deleted_by_user — re-create.
      const createdId = await ctx.store.create(e.itemId, blob);
      if (createdId !== e.itemId) {
        throw new Error(
          `revert: store.create id mismatch: expected ${e.itemId}, got ${createdId}`,
        );
      }
      upsertIndexMap(ctx, e.itemId, serverID);
    }

    await ctx.queue.remove({
      parentId: e.parentId,
      itemId: e.itemId,
      kind: e.kind,
    });
  }
  return false;
}

async function finishWith(ctx, result) {
  // How much is still queued, for the host's needs-sync badge. Only our own
  // queue needs reporting - the host can count the one it holds itself -
  // and it rides along on the flush that was happening anyway. Sent on
  // every finish, including the early ones: a sync that bailed out still
  // changed how much is waiting, and a badge that only ever went up would
  // be worse than none.
  if (ctx.queue.owner === "local") {
    try {
      ctx.pendingCount = await ctx.queue.count();
    } catch {
      /* a count we cannot read is a badge we do not update */
    }
  }
  const patch = {};
  if (ctx.syncKeyDirty) patch.synckey = ctx.synckey;
  if (ctx.indexMapDirty) patch.indexMap = ctx.indexMap;
  if (ctx.pendingCount !== undefined) {
    patch.pendingUserChanges = ctx.pendingCount;
  }
  if (!Object.keys(patch).length) return result;
  try {
    await ctx.provider.updateFolder({
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      patch: { custom: patch },
    });
  } catch (err) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[${ctx.itemKind.changelogKind}-sync] flush failed: ${err?.message ?? String(err)}`,
    });
  }
  return result;
}

/* ── Pull phase ───────────────────────────────────────────────────── */

/** Stop here if the host has cancelled this account's sync.
 *
 *  The abort signal already kills the request in flight, so this covers the
 *  gaps between requests - a long pull, a batched push - where we would
 *  otherwise start work nobody is waiting for any more.
 *
 *  Every call site is *before* sending or storing, never between a server
 *  acknowledgement and the changelog entry it settles: unwinding there is
 *  how a user edit gets lost. Redoing a whole batch costs one request. */
function throwIfCancelled(ctx) {
  ctx.provider?.throwIfCancelled?.(ctx.accountId);
}

async function pullPhase(ctx) {
  const estimate = await runGetItemEstimate({
    account: ctx.account,
    asVersion: ctx.asVersion,
    collectionId: ctx.collectionId,
    synckey: ctx.synckey,
    className: ctx.itemKind.className,
    filterType: ctx.itemKind.filterType,
  });
  let itemsDone = 0;
  let itemsTotal = estimate ?? 0;
  reportProgress(ctx, itemsDone, itemsTotal);

  // Unbounded MoreAvailable loop, matching legacy sync.js:445. The server
  // controls termination via the absent <MoreAvailable/> tag; we trust it
  // to converge. Initial syncs of large folders (10k+ items) hit dozens
  // of iterations and any cap risks a spurious abort mid-pull.
  for (;;) {
    throwIfCancelled(ctx);
    const body = buildSyncBody({
      synckey: ctx.synckey,
      collectionId: ctx.collectionId,
      asVersion: ctx.asVersion,
      withChanges: true,
      withCommands: null,
      className: ctx.itemKind.className,
      filterType: ctx.itemKind.filterType,
      windowSize: ctx.maxItems,
      conflict: ctx.conflict,
    });
    const r = await sendSync({
      account: ctx.account,
      asVersion: ctx.asVersion,
      body,
    });
    if (r.code === "RESYNC") return { code: "RESYNC" };
    if (r.code === "HIERARCHY") return { code: "HIERARCHY" };
    if (r.code === "PROVISION_REQUIRED") return { code: "PROVISION_REQUIRED" };
    if (r.code === "BUSY") return { code: "BUSY" };
    if (r.error) return { status: errorStatus(r.error) };

    if (r.commands) {
      const processed = await applyServerCommands(ctx, r.commands);
      itemsDone += processed;
      if (itemsDone > itemsTotal) itemsTotal = itemsDone;
      reportProgress(ctx, itemsDone, itemsTotal);
    }
    // Empty-body responses leave the syncKey unchanged (legacy parity).
    if (r.synckey != null) {
      ctx.synckey = r.synckey;
      ctx.syncKeyDirty = true;
    }
    if (r.emptyBody) {
      ctx.provider.reportEventLog({
        level: "debug",
        accountId: ctx.accountId,
        folderId: ctx.folderId,
        message:
          "[eas:sync] empty-body Sync response on pull (no changes, syncKey unchanged)",
      });
    }
    if (!r.moreAvailable) return {};
  }
}

/* ── Push phase ───────────────────────────────────────────────────── */

async function pushPhase(ctx, userEdits) {
  const failedItems = new Set();
  // Recurring masters pushed in this pass whose blob carries overrides.
  // Only 16.1 needs them, and only the calendar codec can express them.
  const instanceMasters =
    ctx.asVersion === "16.1" &&
    ctx.syncRecurrence &&
    ctx.itemKind.codec.listInstanceCommands
      ? []
      : null;
  let batchSize = ctx.maxItems;
  let pending = userEdits.slice();
  let itemsDone = 0;
  const itemsTotal = userEdits.length;
  reportProgress(ctx, itemsDone, itemsTotal);

  while (pending.length) {
    throwIfCancelled(ctx);
    const slice = [];
    while (pending.length && slice.length < batchSize) {
      const e = pending.shift();
      if (failedItems.has(e.itemId)) continue;
      slice.push(e);
    }
    if (!slice.length) break;

    const built = await buildPushBatch(ctx, slice, failedItems);
    if (!built.adds.length && !built.mods.length && !built.dels.length) {
      itemsDone += slice.length;
      reportProgress(ctx, itemsDone, itemsTotal);
      continue;
    }

    // Trace any recurring items going out so the user can correlate
    // server responses with the iCal we just sent. 16.1 sends each
    // exception as its own request afterwards, from the instance phase -
    // the master's log entry here is the anchor for those.
    if (ctx.syncRecurrence) {
      for (const a of built.adds) {
        if (blobHasRecurrence(a.item.blob)) {
          logRecurrence(
            ctx,
            `push add: itemId=${a.item.id}, clientId=${a.clientId}`,
            { ical: a.item.blob },
          );
        }
      }
      for (const m of built.mods) {
        if (blobHasRecurrence(m.item.blob)) {
          logRecurrence(
            ctx,
            `push update: itemId=${m.item.id}, serverID=${m.serverID}`,
            { ical: m.item.blob },
          );
        }
      }
    }

    const r = await sendSync({
      account: ctx.account,
      asVersion: ctx.asVersion,
      body: buildSyncBody({
        synckey: ctx.synckey,
        collectionId: ctx.collectionId,
        asVersion: ctx.asVersion,
        withChanges: false,
        withCommands: {
          ...built,
          separator: ctx.separator,
          asVersion: ctx.asVersion,
          codec: ctx.itemKind.codec,
          defaultTimezone: ctx.defaultTimezone,
          syncRecurrence: ctx.syncRecurrence,
          userEmail: ctx.account?.custom?.user,
          fallbackOrganizerName:
            ctx.account?.custom?.fallbackOrganizerNames?.[ctx.collectionId],
          eventLog: ctx.eventLog,
        },
        className: ctx.itemKind.className,
        filterType: ctx.itemKind.filterType,
        conflict: ctx.conflict,
      }),
    });

    if (r.code === "RESYNC") return { code: "RESYNC" };
    if (r.code === "HIERARCHY") return { code: "HIERARCHY" };
    if (r.code === "PROVISION_REQUIRED") return { code: "PROVISION_REQUIRED" };
    if (r.code === "BUSY") return { code: "BUSY" };

    if (r.code === "MALFORMED") {
      if (slice.length > 1) {
        batchSize = Math.max(1, Math.floor(batchSize / 5));
        pending = slice.concat(pending);
        continue;
      }
      failedItems.add(slice[0].itemId);
      reportRejectedPushItem(
        ctx,
        "a single-item batch",
        r.collStatus,
        built.adds.find((a) => a.entry === slice[0]) ??
          built.mods.find((m) => m.entry === slice[0]) ?? { entry: slice[0] },
        "warning",
      );
      itemsDone += 1;
      reportProgress(ctx, itemsDone, itemsTotal);
      continue;
    }

    if (r.error) return { status: errorStatus(r.error) };

    // Empty-body responses leave the syncKey unchanged (legacy parity).
    if (r.synckey != null) {
      ctx.synckey = r.synckey;
      ctx.syncKeyDirty = true;
    }
    if (r.emptyBody) {
      ctx.provider.reportEventLog({
        level: "debug",
        accountId: ctx.accountId,
        folderId: ctx.folderId,
        message:
          "[eas:sync] empty-body Sync response on push (no ACKs, syncKey unchanged)",
      });
    }
    // Reset batch size after a successful round-trip so subsequent
    // batches return to the configured maxItems. Without this the
    // `batchSize / 5` shrinkage triggered by an earlier MALFORMED
    // would persist for the rest of the call, sending one item per
    // request even after the bad item was singleton-dropped.
    batchSize = ctx.maxItems;

    // Always run applyResponses: when the server omits <Responses> (or
    // returns an empty body), the per-sent fallback loops inside still
    // clear the changelog as legacy did. Pass `hadResponsesElement` so
    // applyResponses can emit a debug log when it's running on the
    // fallback path alone.
    const responses = r.responses ?? { adds: [], changes: [], deletes: [] };
    await applyResponses(ctx, responses, built, failedItems, {
      hadResponsesElement: r.responses != null,
      instanceMasters,
    });
    if (r.commands) await applyServerCommands(ctx, r.commands);

    // Modified masters are noted here rather than in applyResponses,
    // which sees only the changes the server *refused* - a successful one
    // is never visited there. So the test is the other way round: a
    // ServerId named in <Responses> is one whose master did not land,
    // whether it was rejected outright or conceded to the server's copy
    // (Status 7 / 8), and its exceptions have nothing to attach to.
    if (instanceMasters) {
      const rejected = new Set(
        responses.changes
          .map((node) => readPathFrom(node, ["ServerId"]))
          .filter(Boolean),
      );
      for (const m of built.mods) {
        if (rejected.has(m.serverID) || failedItems.has(m.entry.itemId)) {
          continue;
        }
        if (blobHasInstanceOverrides(m.item.blob)) {
          instanceMasters.push({
            serverID: m.serverID,
            blob: m.item.blob,
            // The exception set as it stood before this edit, recorded by
            // the provider when the user saved. Lets the instance phase
            // send only what changed instead of re-asserting the lot.
            previous: m.entry?.detail?.exceptions ?? null,
          });
        }
      }
    }

    itemsDone += slice.length;
    reportProgress(ctx, itemsDone, itemsTotal);
  }

  // Tail re-stage: any per-item failures collected during this push pass
  // (singleton-MALFORMED at collection level, plus per-item Response
  // failures from applyResponses) get moved to the changelog tail so the
  // next sync attempts the un-failed items first. Mirrors legacy's
  // updateFailedItems calling removeItemFromChangeLog(id, /*re-add*/ true)
  // (sync.js:1011).
  if (failedItems.size > 0) {
    const failedEntries = userEdits.filter((e) => failedItems.has(e.itemId));
    if (failedEntries.length > 0) {
      try {
        await ctx.queue.moveToTail(
          failedEntries.map((e) => ({
            parentId: e.parentId,
            itemId: e.itemId,
            kind: e.kind,
          })),
        );
      } catch (err) {
        ctx.provider.reportEventLog({
          level: "warning",
          accountId: ctx.accountId,
          folderId: ctx.folderId,
          message: `[${ctx.itemKind.changelogKind}-sync] moving failed items to the queue tail failed: ${err?.message ?? String(err)}`,
        });
      }
    }
  }

  return {
    failedCount: failedItems.size,
    instanceMasters: instanceMasters ?? [],
  };
}

/** Cheap pre-filter for the instance phase: does this blob carry anything
 *  `listInstanceCommands` could emit? A false positive costs one no-op
 *  call, a false negative is impossible - both shapes it walks leave one
 *  of these two markers in the iCal text. */
function blobHasInstanceOverrides(blob) {
  return (
    typeof blob === "string" &&
    (blob.includes("RECURRENCE-ID") || blob.includes("EXDATE"))
  );
}

/* ── Instance phase ───────────────────────────────────────────────────
 *
 * On 16.1 a recurrence exception is not embedded in its master. It is a
 * sibling `<Change>` or `<Delete>` carrying the master's ServerId and an
 * `<InstanceId>` naming the occurrence, so it cannot be sent until the
 * server has assigned that ServerId. For a master the user just created
 * that happens in the push we have only now finished, which is why this
 * runs as its own pass rather than inside `pushPhase`.
 *
 * One command per request, always. Exchange will not take two commands
 * against the same ServerId in one <Commands> block: it applies the
 * first, faults on the second, and answers a global Status 16 that
 * discards the response while keeping what the first one did. Every
 * exception of a master shares that master's ServerId, so they have to
 * travel one at a time. This is also why a master being *modified* sends
 * its exceptions here rather than alongside its own <Change>.
 *
 * Which commands to send is decided against the exception fingerprint the
 * changelog entry carried, so a master that changed on its own re-sends
 * none of its occurrences. An added item has no such baseline and sends
 * everything, which is correct: the server has just met it.
 */
async function instancePhase(ctx, masters) {
  const commands = [];
  for (const m of masters) {
    const built = ctx.itemKind.codec.listInstanceCommands({
      blob: m.blob,
      serverID: m.serverID,
      previous: m.previous ?? null,
      asVersion: ctx.asVersion,
      defaultTimezone: ctx.defaultTimezone,
      syncRecurrence: ctx.syncRecurrence,
      userEmail: ctx.account?.custom?.user,
      fallbackOrganizerName:
        ctx.account?.custom?.fallbackOrganizerNames?.[ctx.collectionId],
      eventLog: ctx.eventLog,
    });
    // The master's blob rides along so a rejection can be logged with the
    // series it belongs to, not just its ServerId.
    for (const command of built) commands.push({ command, blob: m.blob });
  }

  let failedCount = 0;
  for (const { command, blob } of commands) {
    const r = await sendInstanceCommand(ctx, command, blob);
    // Only what the caller has to act on stops the pass. Anything else
    // has already been reported against its own occurrence, and the
    // remaining commands are independent of it.
    if (r.code || r.status) return r;
    if (r.failed) failedCount += 1;
  }
  return { failedCount };
}

/** One instance command, one request. Applies whatever the reply carries
 *  before judging it, so a retry sees current state.
 *
 *  Returns `{ code }` or `{ status }` for the conditions the caller must
 *  handle, else `{ failed }`.
 *
 *  One budget, spent by any retry whatever its cause. Counting per cause is
 *  what this replaces: a global 16 and an item-level 7 are independent, they
 *  arrive in sequence when Exchange is still settling a master, and a
 *  per-cause allowance let the first spend the second's - which reported a
 *  recoverable conflict as a failure. Nothing has to classify a retry in
 *  order to count it now. Which statuses are worth retrying at all, and how
 *  long to wait first, is `INSTANCE_RETRY_DELAY_MS`. */
async function sendInstanceCommand(ctx, command, blob, attempt = 0) {
  const label = `instance ${command.kind} ${command.instanceId}`;
  const reject = (status) => {
    // A rejection leaves the master synced but this exception absent.
    // Counted so the folder reports it, but kept out of `failedItems` -
    // that re-stages the changelog entry, and the master's own push
    // succeeded.
    reportRejectedPushItem(
      ctx,
      label,
      status,
      { entry: { itemId: command.serverID }, item: { blob } },
      "warning",
    );
    return { failed: true };
  };

  /** Re-send for a status the server called transient, if there is budget
   *  left. Returns null when there is not, or when the status is one that
   *  asking again cannot change - the caller then reports it. */
  const retry = async (status) => {
    const delay = INSTANCE_RETRY_DELAY_MS[status];
    if (delay === undefined || attempt >= MAX_INSTANCE_RETRIES) return null;
    ctx.provider.reportEventLog({
      level: "info",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[${ctx.itemKind.changelogKind}-sync] ${label}: Status ${status}, ` +
        `re-sending (${attempt + 1}/${MAX_INSTANCE_RETRIES})`,
    });
    if (delay) await sleep(delay);
    return await sendInstanceCommand(ctx, command, blob, attempt + 1);
  };

  const r = await sendSync({
    account: ctx.account,
    asVersion: ctx.asVersion,
    body: buildSyncBody({
      synckey: ctx.synckey,
      collectionId: ctx.collectionId,
      asVersion: ctx.asVersion,
      withChanges: false,
      withInstanceCommand: command,
    }),
  });

  if (r.code === "RESYNC") return { code: "RESYNC" };
  if (r.code === "HIERARCHY") return { code: "HIERARCHY" };
  if (r.code === "PROVISION_REQUIRED") return { code: "PROVISION_REQUIRED" };
  if (r.code === "BUSY") return { code: "BUSY" };
  if (r.error) {
    // Errors carrying no status of their own - protocol fault, access
    // denied, a reply we cannot parse - are account-level and sink the sync.
    if (!r.topStatus) return { status: errorStatus(r.error) };
    // A global status against this one command is not the sync's failure:
    // the master itself is already on the server, and the other commands
    // are unaffected.
    return (await retry(r.topStatus)) ?? reject(r.topStatus);
  }

  if (r.synckey) {
    ctx.synckey = r.synckey;
    ctx.syncKeyDirty = true;
  }

  // Exchange returns items it has modified as <Change> commands in this
  // same reply - a non-zero synckey with no <GetChanges> is treated as
  // GetChanges=1, so every request implicitly asks for them. Applied
  // before the status is judged: taking the synckey above tells the
  // server we have them, and a retry needs them in hand.
  if (r.commands) await applyServerCommands(ctx, r.commands);

  // We sent exactly one command, so at most one response node concerns
  // us - and only failures are reported at all. A moved occurrence goes
  // out as <Change>, a cancelled one as <Delete>.
  const node =
    (r.responses?.changes ?? [])[0] ?? (r.responses?.deletes ?? [])[0] ?? null;
  const status = node ? readPathFrom(node, ["Status"]) : null;
  if (!status || status === STATUS_OK) return { failed: false };

  // A Status 7 here is our own doing, not another client's: the master was
  // pushed moments ago, and Exchange enriches an item after accepting it.
  // The reply that rejects the command carries those very changes, and they
  // were applied above, so the next attempt is judged against the state the
  // server just handed us - which is why this one needs no delay.
  return (await retry(status)) ?? reject(status);
}

/** A queued edit that no sync can carry out, for the reason given. Skipping
 *  such a row without removing it leaves it queued for good and the account
 *  reported dirty forever, so it is retired here and said out loud.
 *
 *  Matches what the `deleted_by_user` branch already does for its own dead
 *  end, so all three statuses behave alike: nothing is passed over in
 *  silence. */
async function dropUnsatisfiableEntry(ctx, entry, reason) {
  ctx.provider.reportEventLog({
    level: "warning",
    accountId: ctx.accountId,
    folderId: ctx.folderId,
    message:
      `[${ctx.itemKind.changelogKind}-sync] dropping a queued ` +
      `${entry.status.replace("_by_user", "")} of "${entry.itemId}": ${reason}`,
  });
  await ctx.queue.remove({
    parentId: entry.parentId,
    itemId: entry.itemId,
    kind: entry.kind,
  });
}

/** An item the codec cannot put on the wire without changing its meaning
 *  is treated like a server rejection, client-side: warned about with the
 *  item and the reason, counted into the folder's "server did not accept
 *  N elements" warning, and held in the queue so it retries every sync
 *  until the user changes or removes it. Deliberate retry-forever - the
 *  same visibility decision the task-recurrence rejection made: a lie on
 *  the wire is worse than a nagging warning. NOT dropUnsatisfiableEntry,
 *  which removes the entry permanently.
 *
 *  Returns true when the entry was held. */
async function holdIfUnrepresentable(ctx, entry, it, operation, failedItems) {
  const reason = ctx.itemKind.codec.clientRejectReason?.({
    blob: it.blob,
    syncRecurrence: ctx.syncRecurrence,
    asVersion: ctx.asVersion,
  });
  if (!reason) return false;
  ctx.provider.reportEventLog({
    level: "warning",
    accountId: ctx.accountId,
    folderId: ctx.folderId,
    message:
      `[${ctx.itemKind.changelogKind}-sync] cannot send ${operation} for ` +
      `local item ${entry.itemId}: ${reason}: ` +
      `${summarizeBlobForLog(it.blob, ctx.itemKind.changelogKind)}`,
    details: it.blob,
  });
  failedItems.add(entry.itemId);
  return true;
}

async function buildPushBatch(ctx, slice, failedItems) {
  const adds = [];
  const mods = [];
  const dels = [];
  for (const entry of slice) {
    if (entry.status === "added_by_user") {
      const it = await ctx.store.get(entry.itemId);
      if (!it?.blob) {
        await dropUnsatisfiableEntry(
          ctx,
          entry,
          "there is no such local item, so nothing can be sent for it",
        );
        continue;
      }
      if (await holdIfUnrepresentable(ctx, entry, it, "add", failedItems)) {
        continue;
      }
      const clientId = `c-${Date.now().toString(36)}-${adds.length}`;
      adds.push({ entry, clientId, item: it });
    } else if (entry.status === "modified_by_user") {
      const it = await ctx.store.get(entry.itemId);
      if (!it?.blob) {
        await dropUnsatisfiableEntry(
          ctx,
          entry,
          "there is no such local item, so nothing can be sent for it",
        );
        continue;
      }
      if (await holdIfUnrepresentable(ctx, entry, it, "change", failedItems)) {
        continue;
      }
      const serverID =
        ctx.itemKind.codec.readEasServerIdFromBlob(it.blob) ??
        findServerIdByUid(ctx, entry.itemId);
      // The item is here, but neither its own stamp nor the indexMap can say
      // what the server calls it, so a Change has no address to carry. Not
      // something waiting to resolve itself: on an incremental sync the
      // server sends nothing for an item it already has, so nothing will
      // ever fill the gap. Same dead end the delete branch below describes,
      // and it needs the same recovery - a syncKey reset, after which the
      // server re-offers the item as an Add.
      //
      // Reaching it means the local item was never stamped: in practice a
      // folder whose legacy upgrade did not finish. Ordinary edits cannot,
      // since an add edited before it is pushed stays `added_by_user`.
      if (!serverID) {
        await dropUnsatisfiableEntry(
          ctx,
          entry,
          "no ServerId in the item or the indexMap, so a change cannot be " +
            "addressed - a syncKey reset is what recovers this",
        );
        continue;
      }
      mods.push({ entry, serverID, item: it });
    } else if (entry.status === "deleted_by_user") {
      const serverID = findServerIdByUid(ctx, entry.itemId);
      if (!serverID) {
        // No serverID resolvable from the indexMap and the local item
        // is already gone (so no blob X-EAS-SERVERID fallback). Drop
        // the changelog entry; the server keeps its copy. The
        // indexMap is event-driven only (no snapshot self-heal), so
        // recovery requires the server to re-push the item as an Add
        // (e.g. after a syncKey reset / RESYNC), which would re-run
        // applyAdd → upsertIndexMap. Otherwise the item stays orphaned
        // server-side. Common causes: item was created and deleted
        // between syncs (never registered server-side anyway), or an
        // earlier event hook failed to upsert (a real bug worth
        // chasing if it recurs).
        ctx.provider.reportEventLog({
          level: "info",
          accountId: ctx.accountId,
          folderId: ctx.folderId,
          message: `[${ctx.itemKind.changelogKind}-sync] dropped delete for itemId=${entry.itemId}: no serverID in indexMap`,
        });
        await ctx.queue.remove({
          parentId: entry.parentId,
          itemId: entry.itemId,
          kind: entry.kind,
        });
        continue;
      }
      dels.push({ entry, serverID });
    }
  }
  return { adds, mods, dels };
}

/* ── Apply responses to our push ──────────────────────────────────── */

async function applyResponses(ctx, responses, sent, failedItems, opts = {}) {
  const { hadResponsesElement = true, instanceMasters = null } = opts;
  if (!hadResponsesElement) {
    const sentCount = sent.adds.length + sent.mods.length + sent.dels.length;
    ctx.provider.reportEventLog({
      level: "debug",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[eas:sync] no <Responses> in push reply; clearing ${sentCount} sent items via no-ACK fallback`,
    });
  }
  for (const node of responses.adds) {
    const clientId = readPathFrom(node, ["ClientId"]);
    const serverId = readPathFrom(node, ["ServerId"]);
    const status = readPathFrom(node, ["Status"]);
    const sentEntry = sent.adds.find((a) => a.clientId === clientId);
    if (!sentEntry) continue;
    if (status !== STATUS_OK || !serverId) {
      // Per-item Add failure: mark for the final aggregate warning so
      // the user knows this edit didn't make it. Tail-re-staging in
      // pushPhase will move the changelog entry behind the good ones
      // for the next sync.
      failedItems.add(sentEntry.entry.itemId);
      reportRejectedPushItem(ctx, "add", status, sentEntry, "warning");
      continue;
    }
    const stamped = ctx.itemKind.codec.stampEasServerId(
      sentEntry.item.blob,
      serverId,
    );
    await ctx.queue.markServerWrite({
      parentId: ctx.targetID,
      itemId: sentEntry.item.id,
      status: "modified_by_server",
      kind: ctx.itemKind.changelogKind,
    });
    await ctx.store.update(sentEntry.item.id, stamped);
    // Register the just-stamped item in the indexMap so any follow-up
    // server-pushed Change for this ServerID matches the existing
    // local item via applyChangeFromAd instead of falling through to
    // applyAdd and creating a duplicate.
    upsertIndexMap(ctx, sentEntry.item.id, serverId);
    // On 16.1 an exception is not part of the master's payload - it is a
    // separate <Change> keyed on the master's ServerId, which only exists
    // once the server has acked this Add. Note the pair down for the
    // instance phase; nothing can send them before this point.
    if (instanceMasters && blobHasInstanceOverrides(sentEntry.item.blob)) {
      // No `previous`: the server has just learned about this item, so
      // every exception it carries is new to it.
      instanceMasters.push({ serverID: serverId, blob: sentEntry.item.blob });
    }
    await ctx.queue.remove({
      parentId: sentEntry.entry.parentId,
      itemId: sentEntry.entry.itemId,
      kind: sentEntry.entry.kind,
    });
  }
  for (const node of responses.changes) {
    const status = readPathFrom(node, ["Status"]);
    if (status === STATUS_OK) continue;
    if (status === STATUS_CONFLICT || status === STATUS_OBJECT_NOT_FOUND) {
      // Status 7 (server-wins conflict) and Status 8 (object-not-found)
      // are explicit "drop the local edit" signals - both legacy and new
      // discard the changelog entry without flagging as a failure.
      const serverId = readPathFrom(node, ["ServerId"]);
      const sentEntry = sent.mods.find((m) => m.serverID === serverId);
      if (sentEntry) {
        await ctx.queue.remove({
          parentId: sentEntry.entry.parentId,
          itemId: sentEntry.entry.itemId,
          kind: sentEntry.entry.kind,
        });
      }
      continue;
    }
    // Any other non-success status (4 / 5 / 6 / …) is a real failure.
    // Mark for tail-re-stage + final warning. The fallback below will
    // see the non-OK status and skip the changelog removal.
    const serverId = readPathFrom(node, ["ServerId"]);
    const sentEntry = sent.mods.find((m) => m.serverID === serverId);
    if (sentEntry) {
      failedItems.add(sentEntry.entry.itemId);
      reportRejectedPushItem(ctx, "change", status, sentEntry, "warning");
    }
  }
  for (const node of responses.deletes) {
    const status = readPathFrom(node, ["Status"]);
    const serverId = readPathFrom(node, ["ServerId"]);
    const sentEntry = sent.dels.find((d) => d.serverID === serverId);
    if (!sentEntry) continue;
    if (status === STATUS_OK || status === STATUS_OBJECT_NOT_FOUND) {
      await ctx.queue.remove({
        parentId: sentEntry.entry.parentId,
        itemId: sentEntry.entry.itemId,
        kind: sentEntry.entry.kind,
      });
      removeFromIndexMap(ctx, sentEntry.entry.itemId);
    } else {
      // Still not tracked - legacy didn't either ("What can we do about
      // failed deletes? SyncLog" - sync.js:1073), and its soft-fail path
      // never reached updateFailedItems, so the item was neither counted
      // nor re-staged. All that changes here is that the dump legacy
      // sent to the console is now visible in the log, at the level that
      // matches how quiet it was.
      reportRejectedPushItem(ctx, "delete", status, sentEntry, "debug");
    }
  }
  for (const m of sent.mods) {
    const ack = responses.changes.find(
      (n) => readPathFrom(n, ["ServerId"]) === m.serverID,
    );
    const status = ack ? readPathFrom(ack, ["Status"]) : STATUS_OK;
    if (!ack || status === STATUS_OK) {
      await ctx.queue.remove({
        parentId: m.entry.parentId,
        itemId: m.entry.itemId,
        kind: m.entry.kind,
      });
    }
    // Status 7/8: changelogRemove already happened in the per-response
    // loop above. Other non-success: leave entry untouched (failedItems
    // tracking happens above, tail-re-stage runs at end of pushPhase).
  }
  for (const d of sent.dels) {
    const ack = responses.deletes.find(
      (n) => readPathFrom(n, ["ServerId"]) === d.serverID,
    );
    if (!ack) {
      await ctx.queue.remove({
        parentId: d.entry.parentId,
        itemId: d.entry.itemId,
        kind: d.entry.kind,
      });
      removeFromIndexMap(ctx, d.entry.itemId);
    }
  }
}

/* ── Apply server commands ───────────────────────────────────────── */

async function applyServerCommands(ctx, commands) {
  let processed = 0;
  // Once, at the top: the loops below write to the store, and a check
  // between two of those writes would leave the batch half-applied for no
  // benefit - the server will re-send what we did not acknowledge.
  throwIfCancelled(ctx);
  for (const node of commands.adds) {
    await applyAdd(ctx, node);
    processed++;
  }
  for (const node of commands.changes) {
    await applyChange(ctx, node);
    processed++;
  }
  for (const node of commands.deletes) {
    await applyDelete(ctx, node);
    processed++;
  }
  for (const node of commands.softDeletes) {
    await applyDelete(ctx, node);
    processed++;
  }
  return processed;
}

async function applyAdd(ctx, addNode) {
  const serverID = readPathFrom(addNode, ["ServerId"]);
  if (!serverID) return;
  const ad = childByTag(addNode, "ApplicationData");
  if (!ad) return;
  await maybeRecordFallbackOrganizerName(ctx, ad);
  const existing = await findExistingByServerId(ctx, serverID);
  if (existing) return applyChangeFromAd(ctx, ad, existing, serverID);

  const newId = crypto.randomUUID();
  const blob = await ctx.itemKind.codec.applicationDataToBlob({
    adNode: ad,
    serverID,
    asVersion: ctx.asVersion,
    separator: ctx.separator,
    defaultTimezone: ctx.defaultTimezone,
    syncRecurrence: ctx.syncRecurrence,
    msTodoCompat: ctx.msTodoCompat,
    uid: newId,
    userEmail: ctx.account?.custom?.user,
    eventLog: ctx.eventLog,
  });
  await ctx.queue.markServerWrite({
    parentId: ctx.targetID,
    itemId: newId,
    status: "added_by_server",
    kind: ctx.itemKind.changelogKind,
  });
  const createdId = await ctx.store.create(newId, blob);
  if (createdId !== newId) {
    throw new Error(
      `store.create id mismatch: expected ${newId}, got ${createdId}`,
    );
  }
  await verifyRoundTrip(ctx, newId, blob, "create");
  upsertIndexMap(ctx, newId, serverID);
  if (blobHasRecurrence(blob)) {
    logRecurrence(ctx, `pull add: itemId=${newId}, serverID=${serverID}`, {
      ical: blob,
    });
  }
}

async function applyChange(ctx, changeNode) {
  const serverID = readPathFrom(changeNode, ["ServerId"]);
  if (!serverID) return;
  const ad = childByTag(changeNode, "ApplicationData");
  if (!ad) return;
  await maybeRecordFallbackOrganizerName(ctx, ad);
  const existing = await findExistingByServerId(ctx, serverID);
  if (!existing) return declineChangeForUnknownItem(ctx, serverID);
  // 16.1 per-instance Change: ApplicationData carries <InstanceId> and
  // is scoped to a single occurrence of the master event referenced by
  // ServerId. Route to the codec's exception path; bail back to the
  // normal master update if the codec doesn't support it.
  const instanceId = readPathFrom(ad, ["InstanceId"]);
  if (
    instanceId &&
    ctx.syncRecurrence &&
    ctx.itemKind.codec.applyInstanceChange
  ) {
    return applyExceptionChange(ctx, ad, existing, instanceId, serverID);
  }
  return applyChangeFromAd(ctx, ad, existing, serverID);
}

async function applyExceptionChange(ctx, ad, existing, instanceId, serverID) {
  const instanceUtc = parseEasInstanceId(instanceId);
  if (!instanceUtc) return applyChangeFromAd(ctx, ad, existing, serverID);

  const deleted = readPathFrom(ad, ["Deleted"]) === "1";
  const codec = ctx.itemKind.codec;
  let nextBlob;
  if (deleted) {
    nextBlob = codec.applyInstanceDelete?.({
      ical: existing.blob,
      instanceUtc,
    });
  } else {
    nextBlob = codec.applyInstanceChange?.({
      ical: existing.blob,
      adNode: ad,
      instanceUtc,
      asVersion: ctx.asVersion,
      defaultTimezone: ctx.defaultTimezone,
      userEmail: ctx.account?.custom?.user,
    });
  }
  if (!nextBlob || nextBlob === existing.blob) {
    logRecurrence(
      ctx,
      `pull 16.1 exception ${deleted ? "delete" : "change"} no-op: itemId=${existing.itemId}, instance=${instanceId}`,
      {
        ical: existing.blob,
      },
    );
    return;
  }

  // This path rewrites the blob from itself, so an item that arrived here
  // without a stamp would be stored without one again. The command knows
  // the id, so restore it while we are writing anyway - same repair as
  // applyChangeFromAd, which this function otherwise bypasses.
  if (serverID && !codec.readEasServerIdFromBlob(nextBlob)) {
    ctx.provider.reportEventLog({
      level: "debug",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[${ctx.itemKind.changelogKind}-sync] itemId=${existing.itemId} carried ` +
        `no ServerId stamp; restamping from the server command (${serverID})`,
    });
    nextBlob = codec.stampEasServerId(nextBlob, serverID);
  }

  await ctx.queue.markServerWrite({
    parentId: ctx.targetID,
    itemId: existing.itemId,
    status: "modified_by_server",
    kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.update(existing.itemId, nextBlob);
  await verifyRoundTrip(
    ctx,
    existing.itemId,
    nextBlob,
    deleted ? "exception-delete" : "exception-update",
  );
  logRecurrence(
    ctx,
    `pull 16.1 exception ${deleted ? "delete" : "change"} applied: itemId=${existing.itemId}, instance=${instanceId}`,
    {
      before: existing.blob,
      after: nextBlob,
    },
  );
  // Re-key by the master's serverId (unchanged); keep the new blob. The
  // caller's id is the last fallback rather than the first choice: this
  // path rewrites one occurrence of an existing series, so the blob's own
  // stamp is the value to keep. It only runs out when the blob lost it.
  const masterServerId =
    codec.readEasServerIdFromBlob(nextBlob) ??
    codec.readEasServerIdFromBlob(existing.blob) ??
    serverID;
  if (masterServerId) {
    upsertIndexMap(ctx, existing.itemId, masterServerId);
  }
}

/** Parse an EAS InstanceId string ("YYYYMMDDTHHMMSSZ" or extended-ISO)
 *  into a JS Date. EAS encodes the original master occurrence in UTC. */
function parseEasInstanceId(s) {
  if (!s) return null;
  // Fraction as well as separators - see parseEasUtc. A null here is not
  // inert: the caller falls back to treating a per-occurrence change as a
  // change to the whole series.
  const compact = String(s).replace(/[-:]|\.\d+/g, "");
  const m = /^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?Z?$/.exec(compact);
  if (!m) return null;
  const [, y, mo, d, h = "0", mi = "0", se = "0"] = m;
  return new Date(Date.UTC(+y, +mo - 1, +d, +h, +mi, +se));
}

async function applyChangeFromAd(ctx, ad, existing, serverID = null) {
  // The caller's ServerId wins over the blob's. It comes from the command
  // we are applying, so it is what the server calls this item right now;
  // the stamp in the blob is a cache of the same value, kept so a rebind
  // can rebuild the index from the items alone.
  //
  // The two can disagree: anything that replaces an item's body wholesale
  // - an import, another add-on, a re-imported test fixture - drops the
  // X- property while the index keeps the mapping. Reading the cache first
  // meant handing the codec a null and killing the whole folder sync at
  // the moment the server was telling us the answer. The push path never
  // had this problem: it has always resolved blob ?? index.
  const stamped = ctx.itemKind.codec.readEasServerIdFromBlob(existing.blob);
  const id = serverID ?? stamped;
  if (id && !stamped) {
    // Not silent: the blob is written back stamped below, so this is the
    // only trace that an item had lost its identity. If it shows up on an
    // account nobody re-imported into, something else is eating the
    // property and that is a separate bug.
    ctx.provider.reportEventLog({
      level: "debug",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[${ctx.itemKind.changelogKind}-sync] itemId=${existing.itemId} carried ` +
        `no ServerId stamp; restamping from the server command (${id})`,
    });
  }
  // Pass `existingBlob` so the codec merges the partial AD into the
  // current local blob instead of rebuilding from scratch. Exchange
  // routinely echoes Changes carrying only the modified fields (e.g.
  // just <DtStamp>); without merge, untouched fields would be lost.
  const blob = await ctx.itemKind.codec.applicationDataToBlob({
    adNode: ad,
    existingBlob: existing.blob,
    serverID: id,
    asVersion: ctx.asVersion,
    separator: ctx.separator,
    defaultTimezone: ctx.defaultTimezone,
    syncRecurrence: ctx.syncRecurrence,
    msTodoCompat: ctx.msTodoCompat,
    uid: existing.itemId,
    userEmail: ctx.account?.custom?.user,
    eventLog: ctx.eventLog,
  });
  await ctx.queue.markServerWrite({
    parentId: ctx.targetID,
    itemId: existing.itemId,
    status: "modified_by_server",
    kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.update(existing.itemId, blob);
  await verifyRoundTrip(ctx, existing.itemId, blob, "update");
  const masterServerId =
    ctx.itemKind.codec.readEasServerIdFromBlob(blob) ??
    ctx.itemKind.codec.readEasServerIdFromBlob(existing.blob);
  if (masterServerId) {
    upsertIndexMap(ctx, existing.itemId, masterServerId);
  }
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
  const existing = await findExistingByServerId(ctx, serverID);
  if (!existing) return;
  await ctx.queue.markServerWrite({
    parentId: ctx.targetID,
    itemId: existing.itemId,
    status: "deleted_by_server",
    kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.delete(existing.itemId);
  removeFromIndexMap(ctx, existing.itemId);
}

/* ── Sync request building ────────────────────────────────────────── */

function buildSyncBody({
  synckey,
  collectionId,
  asVersion,
  withChanges,
  withCommands,
  withInstanceCommand,
  className,
  filterType,
  windowSize,
  conflict,
}) {
  // The account's conflict preference is stated on every request that has
  // an Options block or carries Commands - not on AS 2.5, which keeps the
  // server default: that branch is minimal-touch by policy and untestable.
  const sendConflict = conflict != null && asVersion !== "2.5";
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
      // FilterType narrows Calendar pulls to a window (e.g. last 2 weeks).
      // Only meaningful for Calendar - Contacts/Tasks have no time axis.
      // Legacy gates this on `type == "Calendar"` (sync.js:401); we mirror
      // by emitting the tag only for the Calendar class.
      if (className === "Calendar") w.atag("FilterType", String(filterType));
      w.atag("Class", className);
      if (sendConflict) w.atag("Conflict", conflict);
      w.switchpage("AirSyncBase");
      w.otag("BodyPreference");
      w.atag("Type", "1");
      w.ctag();
      w.switchpage("AirSync");
      w.ctag();
    } else if (className === "Calendar") {
      // AS 2.5 has no Class/BodyPreference inside Options (those tags were
      // introduced in AS 12). Calendar is the one folder type that still
      // benefits from a FilterType - without it the server treats the
      // initial pull as "every event ever". Matches legacy sync.js:409-412.
      w.otag("Options");
      w.atag("FilterType", String(filterType));
      w.ctag();
    }
  }
  // A Commands-only push batch states the preference too - it is the
  // request the server resolves conflicts IN. Options precedes Commands in
  // the Collection schema. Instance-command requests are deliberately
  // excluded: an exception <Change> always follows our own master push in
  // the same sync, and with an explicit <Conflict> Exchange conflict-checks
  // it against exactly that master change and discards it with Status 7
  // (measured on Exchange Online; without the element the same request is
  // accepted). A conflict verdict against our own change is never wanted.
  if (!withChanges && withCommands && sendConflict) {
    w.otag("Options");
    w.atag("Conflict", conflict);
    w.ctag();
  }
  if (withCommands) appendCommands(w, withCommands);
  if (withInstanceCommand) appendInstanceCommand(w, withInstanceCommand);
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** `<Commands>` holding a single per-instance command - see the instance
 *  phase for why it is never more than one. The descriptor writes the
 *  element itself and leaves the builder on AirSync, same contract as the
 *  codec calls in `appendCommands`. */
function appendInstanceCommand(w, command) {
  w.otag("Commands");
  command.emit(w);
  w.ctag();
}

function appendCommands(
  w,
  {
    adds,
    mods,
    dels,
    separator,
    asVersion,
    codec,
    defaultTimezone,
    syncRecurrence,
    userEmail,
    fallbackOrganizerName,
    eventLog,
  },
) {
  if (!adds.length && !mods.length && !dels.length) return;
  w.otag("Commands");
  for (const a of adds) {
    w.otag("Add");
    w.atag("ClientId", a.clientId);
    w.otag("ApplicationData");
    codec.appendApplicationDataFromBlob({
      builder: w,
      op: "add",
      blob: a.item.blob,
      asVersion,
      separator,
      defaultTimezone,
      syncRecurrence,
      userEmail,
      fallbackOrganizerName,
      eventLog,
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
      builder: w,
      op: "change",
      blob: m.item.blob,
      asVersion,
      separator,
      defaultTimezone,
      syncRecurrence,
      userEmail,
      fallbackOrganizerName,
      eventLog,
    });
    w.switchpage("AirSync");
    w.ctag();
    w.ctag();
    // A 16.1 master's exceptions are *not* emitted here. They share this
    // master's ServerId, and Exchange rejects a second command against a
    // ServerId in the same block - the instance phase sends them one
    // request at a time once this push has landed.
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
  const { doc } = await easRequest({
    account,
    command: "Sync",
    body,
    asVersion,
  });
  // MS-ASCMD §2.2.3.165.2: an empty 200 OK is a valid "no changes,
  // syncKey unchanged" response. Match legacy EAS4's
  // allowEmptyResponse behaviour by returning a no-op success result;
  // callers detect it via `emptyBody` for logging.
  if (!doc) {
    return {
      synckey: null,
      moreAvailable: false,
      commands: null,
      responses: null,
      emptyBody: true,
    };
  }
  return parseSyncResponse(doc);
}

function parseSyncResponse(doc) {
  const top = readPath(doc, ["Status"]);
  if (top && top !== STATUS_OK) {
    if (top === STATUS_BUSY) return { code: "BUSY", topStatus: top };
    if (top === STATUS_RESYNC) return { code: "RESYNC", topStatus: top };
    if (top === STATUS_FOLDER_HIERARCHY)
      return { code: "HIERARCHY", topStatus: top };
    if (STATUS_PROVISION_REQUIRED.has(top))
      return { code: "PROVISION_REQUIRED", topStatus: top };
    if (STATUS_PROTOCOL_FAULT.has(top)) {
      return {
        error: browser.i18n.getMessage("eas.sync.error.protocolFault", [top]),
      };
    }
    if (STATUS_ACCESS_DENIED.has(top)) {
      const reason = browser.i18n.getMessage(
        `eas.sync.error.accessDenied.${top}`,
      );
      return {
        error: browser.i18n.getMessage("eas.sync.error.accessDenied", [
          reason,
          top,
        ]),
      };
    }
    // `topStatus` rides along so a caller that knows what to do with a
    // particular code can act on it; everyone else sees only `error` and
    // behaves as before.
    return { error: `Sync top status ${top}`, topStatus: top };
  }
  const collection = doc.getElementsByTagName("Collection")[0];
  if (!collection) return { error: "Sync response missing Collection" };

  const collStatus = readPathFrom(collection, ["Status"]) ?? STATUS_OK;
  // STATUS_CONFLICT at collection level: server-wins conflict, legacy
  // treated this as silent OK and continued. Per-item conflicts have
  // their own handling further down. Fall through to the success path.
  if (collStatus !== STATUS_OK && collStatus !== STATUS_CONFLICT) {
    if (collStatus === STATUS_RESYNC) return { code: "RESYNC", collStatus };
    if (collStatus === STATUS_FOLDER_HIERARCHY)
      return { code: "HIERARCHY", collStatus };
    if (STATUS_PROVISION_REQUIRED.has(collStatus))
      return { code: "PROVISION_REQUIRED", collStatus };
    if (collStatus === STATUS_BUSY) return { code: "BUSY", collStatus };
    if (
      collStatus === STATUS_MALFORMED ||
      collStatus === STATUS_INVALID ||
      collStatus === STATUS_TEMP_SERVER ||
      collStatus === STATUS_OBJECT_NOT_FOUND
    ) {
      return { code: "MALFORMED", collStatus };
    }
    return { error: `Sync collection status ${collStatus}`, collStatus };
  }

  const synckey = readPathFrom(collection, ["SyncKey"]);
  if (!synckey) {
    throw new Error(
      browser.i18n.getMessage("eas.sync.error.missingField", [
        "Sync.Collections.Collection.SyncKey",
      ]),
    );
  }
  const moreAvailable = !!childByTag(collection, "MoreAvailable");

  let commands = null;
  const cmdNode = childByTag(collection, "Commands");
  if (cmdNode) {
    commands = {
      adds: Array.from(cmdNode.children).filter((c) => c.tagName === "Add"),
      changes: Array.from(cmdNode.children).filter(
        (c) => c.tagName === "Change",
      ),
      deletes: Array.from(cmdNode.children).filter(
        (c) => c.tagName === "Delete",
      ),
      softDeletes: Array.from(cmdNode.children).filter(
        (c) => c.tagName === "SoftDelete",
      ),
    };
  }

  if (DEV_FIXTURE_ADD_XML) {
    const fixtureAdd = parseDevFixtureAdd();
    if (fixtureAdd) {
      if (!commands) {
        commands = { adds: [], changes: [], deletes: [], softDeletes: [] };
      }
      commands.adds.push(fixtureAdd);
    }
  }

  let responses = null;
  const respNode = childByTag(collection, "Responses");
  if (respNode) {
    responses = {
      adds: Array.from(respNode.children).filter((c) => c.tagName === "Add"),
      changes: Array.from(respNode.children).filter(
        (c) => c.tagName === "Change",
      ),
      deletes: Array.from(respNode.children).filter(
        (c) => c.tagName === "Delete",
      ),
    };
  }

  return { synckey, moreAvailable, commands, responses };
}

/* ── Helpers ──────────────────────────────────────────────────────── */

/** Has the user deleted the item this ServerId refers to, with the delete
 *  still waiting in the changelog?
 *
 *  The local item is gone, so `findExistingByServerId` cannot answer - but
 *  the indexMap still maps that ServerId, because it is only cleaned once
 *  the delete has been acknowledged. With the push running before the
 *  pull, an acked delete leaves nothing behind by the time server
 *  commands are applied - so this only answers "yes" in the window where
 *  the push could not get rid of the item (the server refused the delete,
 *  or the push itself failed) and the pull then hands us a `<Change>` for
 *  it. That mapping-plus-queued-entry pair is what makes the question
 *  answerable.
 *
 *  It does not survive a heavy reset, which empties the indexMap before
 *  re-pulling: in that window a queued delete is invisible here. */
function hasPendingUserDelete(ctx, serverId) {
  const entry = ctx.indexMap.find((e) => e.serverId === serverId);
  if (!entry) return false;
  // The queue as it stood when this sync began. Reading it live would need
  // this to be async for no gain: the indexMap check above has already
  // excluded every delete this sync managed to push, because acknowledging
  // one removes its mapping.
  return ctx.pendingAtStart.some(
    (e) => e?.itemId === entry.uid && e?.status === "deleted_by_user",
  );
}

/** A `<Change>` names an item the server believes we already hold, and EAS
 *  is stateful enough for that to mean something: our SyncKey acknowledged
 *  the item. Not finding it locally therefore has exactly two readings.
 *
 *  Either a delete for it is still queued - this sync's push failed to
 *  land it, and the change arriving now is for an item the user already
 *  removed - which is ordinary and silent; the delete retries next sync.
 *  Or the two states have drifted, which is a defect in the sync engine,
 *  ours or the server's, and the only evidence of it is this moment;
 *  hence the warning. (With the push ahead of the pull, an *acked* delete
 *  cannot produce this situation - neither tested server sends changes
 *  for an item it has agreed is gone.)
 *
 *  Neither reading justifies re-creating the item from the change: that
 *  would hide the first case behind an orphaned copy and the second behind
 *  no message at all. */
function declineChangeForUnknownItem(ctx, serverID) {
  if (hasPendingUserDelete(ctx, serverID)) return;
  ctx.eventLog(
    "warning",
    `ignored <Change> for ${serverID}: no local item, and no delete queued for it - local state has drifted from the server`,
  );
}

/** Look up the local item by its EAS server-side id. Returns
 *  `{ itemId, blob }` or null.
 *
 *  The indexMap is a cache, not the authority: the server id is also
 *  stamped into every blob we store, which is what the push side already
 *  reads back (see `buildPushBatch`). So a miss here falls through to the
 *  blobs rather than concluding the item is new.
 *
 *  That distinction is the whole point. A resync empties the indexMap
 *  before re-pulling the collection, so without the fallback every item
 *  arrives as an <Add> that matches nothing and gets re-created beside the
 *  copy already in the address book or calendar - a full duplicate set. */
async function findExistingByServerId(ctx, serverId) {
  const entry = ctx.indexMap.find((e) => e.serverId === serverId);
  if (entry) {
    const item = await ctx.store.get(entry.uid);
    if (item) return { itemId: entry.uid, blob: item.blob };
    // Mapped to an item that is no longer there - fall through and let
    // the blobs have the final say.
  }

  const itemId = (await serverIdScan(ctx)).get(serverId);
  if (!itemId) return null;
  const item = await ctx.store.get(itemId);
  if (!item) return null;
  // Put the mapping back so the rest of this pass - and, once the pass
  // flushes indexMapDirty, later ones - take the fast path again.
  upsertIndexMap(ctx, itemId, serverId);
  return { itemId, blob: item.blob };
}

/** `serverId -> itemId` built from the stamps in the stored blobs.
 *
 *  Built at most once per pass and only when something actually misses, so
 *  a healthy incremental sync never reads the store in bulk. Scanning per
 *  miss instead would be quadratic - a 5000-item resync misses 5000 times.
 */
async function serverIdScan(ctx) {
  if (ctx.serverIdScan) return ctx.serverIdScan;

  const map = new Map();
  try {
    for (const it of await ctx.store.list()) {
      if (!it?.blob) continue;
      let stamped;
      try {
        stamped = ctx.itemKind.codec.readEasServerIdFromBlob(it.blob);
      } catch {
        continue; // unparsable blob - it just cannot answer
      }
      // Items with no stamp were created locally and never pushed, so they
      // belong to no server id. First writer wins on a collision, which
      // keeps the result stable for a store that already carries
      // duplicates.
      if (stamped && !map.has(stamped)) map.set(stamped, it.id);
    }
  } catch (err) {
    // A store we cannot read leaves us exactly where we were without the
    // fallback; failing the sync over it would be worse.
    console.warn(
      "[eas-4-tbsync] server-id scan failed; identity falls back to the indexMap alone:",
      err?.message ?? err,
    );
  }

  ctx.serverIdScan = map;
  return map;
}

function findServerIdByUid(ctx, uid) {
  return ctx.indexMap.find((e) => e.uid === uid)?.serverId ?? null;
}

function upsertIndexMap(ctx, uid, serverId) {
  const existing = ctx.indexMap.find((e) => e.uid === uid);
  if (existing) {
    if (existing.serverId !== serverId) {
      existing.serverId = serverId;
      ctx.indexMapDirty = true;
    }
    return;
  }
  ctx.indexMap.push({ uid, serverId });
  ctx.indexMapDirty = true;
}

function removeFromIndexMap(ctx, uid) {
  const idx = ctx.indexMap.findIndex((e) => e.uid === uid);
  if (idx === -1) return;
  ctx.indexMap.splice(idx, 1);
  ctx.indexMapDirty = true;
}

/** When an inbound calendar event tells us the local user IS the
 *  organizer (MeetingStatus has the R bit cleared) and carries an
 *  OrganizerName, capture that name into
 *  `account.custom.fallbackOrganizerNames[<collectionId>]` for the
 *  upsync path to fall back on when the local iCal has no ORGANIZER CN.
 *  Mirrors legacy calendarsync.js:260-261, scoped per-calendar via the
 *  EAS folder serverId (collectionId). No-op for tasks (no MeetingStatus
 *  / OrganizerName), no-op when the value is unchanged. */
async function maybeRecordFallbackOrganizerName(ctx, adNode) {
  if (!adNode) return;
  const orgName = readPathFrom(adNode, ["OrganizerName"]);
  if (!orgName) return;
  const meetingStatus = readPathFrom(adNode, ["MeetingStatus"]);
  if (!meetingStatus) return;
  const ms = parseInt(meetingStatus, 10) || 0;
  if (ms & 0x2) return; // R bit: received from another organizer.

  const key = ctx.collectionId;
  if (!key) return;
  const current = ctx.account.custom?.fallbackOrganizerNames?.[key];
  if (current === orgName) return;

  const next = {
    ...(ctx.account.custom?.fallbackOrganizerNames ?? {}),
    [key]: orgName,
  };
  await ctx.provider.updateAccount({
    accountId: ctx.accountId,
    patch: { custom: { fallbackOrganizerNames: next } },
  });
  // Mirror into the in-memory ctx.account so subsequent calls in this
  // same sync see the updated value and skip the dedupe-no-op write.
  if (!ctx.account.custom) ctx.account.custom = {};
  ctx.account.custom.fallbackOrganizerNames = next;
}

function reportProgress(ctx, itemsDone, itemsTotal) {
  ctx.provider.reportProgress({
    accountId: ctx.accountId,
    folderId: ctx.folderId,
    itemsDone,
    itemsTotal,
  });
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
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
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[${kind}-sync] roundtrip readback (${op}) failed for ${itemId}: ${err?.message ?? String(err)}`,
    });
    return;
  }
  if (actual === expected) return;
  if (actual == null) {
    ctx.provider.reportEventLog({
      level: "debug",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[${kind}-sync] roundtrip readback (${op}) returned no item for ${itemId}`,
    });
    return;
  }

  const diff = diffComponentProperties(expected, actual, target);
  if (!diff) return;
  if (!diff.dropped.length && !diff.added.length && !diff.changed.length)
    return;

  const lines = [`[${kind}-sync] roundtrip mismatch on ${op} of ${itemId}`];
  if (diff.dropped.length) lines.push("dropped: " + diff.dropped.join(" | "));
  if (diff.added.length) lines.push("added:   " + diff.added.join(" | "));
  if (diff.changed.length) lines.push("changed: " + diff.changed.join(" | "));
  ctx.provider.reportEventLog({
    level: "debug",
    accountId: ctx.accountId,
    folderId: ctx.folderId,
    message: lines.join("\n"),
  });
}

/** Resolve the kind to a parse target: which top-level component to
 *  parse and which subcomponent (if any) to compare. Returns null for
 *  kinds that don't have a known structured form. */
function roundTripTargetFor(kind) {
  if (kind === "event") return { outer: "vcalendar", inner: "vevent" };
  if (kind === "task") return { outer: "vcalendar", inner: "vtodo" };
  if (kind === "contact") return { outer: "vcard", inner: null };
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
  const added = [];
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
    if (eList.length === aList.length && eList.every((s, i) => s === aList[i]))
      continue;
    // Same name, different content: report as changed.
    changed.push(`${name}: ${eList.join(",")} → ${aList.join(",")}`);
  }
  return { dropped, added, changed };
}

/* ── Rejected-item reporting ──────────────────────────────────────────
 *
 * A push the server refuses is counted (`failedItems`) and re-staged at
 * the tail of the queue. The folder's own status says only "did not
 * accept N elements", so these entries name the item (issue #319): a short
 * summary in the message to keep the log scannable, and the blob itself in
 * the details.
 *
 * The request and response halves are already logged as `[eas:net] send /
 * receive Sync`, so the item is the one thing left to say.
 */

/** One entry naming an item the server refused. `sentEntry` is an
 *  element of `built.adds` / `.mods` / `.dels`; deletes carry no `item`,
 *  so they report without a summary or details. */
function reportRejectedPushItem(ctx, operation, status, sentEntry, level) {
  const itemId = sentEntry?.entry?.itemId ?? sentEntry?.item?.id ?? "unknown";
  const localStatus = sentEntry?.entry?.status;
  const blob = sentEntry?.item?.blob;
  const summary = summarizeBlobForLog(blob, ctx.itemKind.changelogKind);
  ctx.provider.reportEventLog({
    level,
    accountId: ctx.accountId,
    folderId: ctx.folderId,
    message:
      `[${ctx.itemKind.changelogKind}-sync] server rejected ${operation} for ` +
      `local item ${itemId}` +
      (localStatus ? ` (${localStatus})` : "") +
      ` (Status ${status ?? "unknown"})` +
      (summary ? `: ${summary}` : ""),
    details: typeof blob === "string" && blob ? blob : null,
  });
}

/** Enough of an item to recognize it in a log line. Never throws: a blob
 *  we cannot parse is exactly the kind that gets rejected, and the
 *  details still carry it verbatim. */
function summarizeBlobForLog(blob, kind) {
  if (typeof blob !== "string" || !blob) return "";
  const target = roundTripTargetFor(kind);
  if (!target) return "";
  const inner = innerComponent(blob, target);
  if (!inner) return "";
  const fields =
    kind === "contact"
      ? ["fn", "n", "email"]
      : ["summary", "dtstart", "dtend", "due", "uid"];
  const parts = [];
  for (const name of fields) {
    const value = inner.getFirstPropertyValue(name);
    if (value != null && String(value) !== "") {
      parts.push(`${name.toUpperCase()}=${truncateForLog(String(value))}`);
    }
  }
  return parts.join(" ");
}

function truncateForLog(value, max = 80) {
  const singleLine = value.replace(/\s+/g, " ").trim();
  return singleLine.length > max
    ? `${singleLine.slice(0, Math.max(0, max - 3))}...`
    : singleLine;
}

/** Parse a stored blob and descend to the component that carries the
 *  properties: VEVENT/VTODO for iCal, the vCard itself for contacts.
 *  Returns null if the text does not parse or the inner component is
 *  absent. */
function innerComponent(text, target) {
  let comp;
  try {
    comp = new ICAL.Component(ICAL.parse(text));
  } catch {
    return null;
  }
  return target.inner ? comp.getFirstSubcomponent(target.inner) : comp;
}

function innerProps(text, target) {
  const inner = innerComponent(text, target);
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
    const j = prop.toJSON(); // [name, paramsObj, valueType, ...values]
    const name = j[0];
    const params = j[1] ?? {};
    const valueType = j[2];
    const values = j.slice(3);
    const paramKeys = Object.keys(params).sort();
    const paramStr = paramKeys
      .map((k) => `;${k.toUpperCase()}=${stringifyValue(params[k])}`)
      .join("");
    const valStr = values.map(stringifyValue).join(",");
    return `${name.toUpperCase()}${paramStr}${valueType ? "" : ""}:${valStr}`;
  } catch {
    try {
      return prop.toICALString();
    } catch {
      return `${prop.name}:${stringifyValue(prop.getFirstValue())}`;
    }
  }
}

function stringifyValue(v) {
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (Array.isArray(v)) return v.map(stringifyValue).join(",");
  if (typeof v === "object" && typeof v.toString === "function")
    return v.toString();
  return String(v);
}
