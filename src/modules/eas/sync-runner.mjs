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
import { easRequest, RETRY_LATER_BACKOFF_MS } from "../network.mjs";
import { canonicalPropertyString } from "./calendar-codec.mjs";
import { readPath, readPathFrom } from "./wbxml-helpers.mjs";
import { runGetItemEstimate } from "./get-item-estimate.mjs";
import { fetchServerItem, fetchServerItems } from "./item-operations.mjs";
import { easCommandLikelyAvailable } from "./allowed-commands.mjs";
import { accountUserAddress } from "./settings.mjs";
import {
  USERRESPONSE_TO_PARTSTAT,
  announceableOf,
  droppedAttendees,
  isReceivedMeeting,
  sameAnnounceable,
  preserveSelfPartstat,
  repliedPartstatOf,
  stampRepliedPartstat,
  selfUserResponses,
  serverKnownPartstat,
} from "./calendar-codec.mjs";
import { sendMeetingResponse } from "./meeting-response.mjs";
import {
  rememberClientScheduling,
  versionNeedsClientScheduling,
} from "./client-scheduling.mjs";
import {
  buildMeetingResponseMime,
  sendMail,
} from "./meeting-response-mail.mjs";
import { buildMeetingRequestMime } from "./meeting-request-mail.mjs";
import { createServerIdIndex } from "./server-id-index.mjs";
import {
  duplicateClusters,
  noteUidClaim,
  titleFromBlob,
} from "./duplicate-uids.mjs";
import { deleteSurplusCopies } from "./duplicate-cleanup.mjs";
import { blobHasInstanceOverrides, buildSyncBody } from "./sync-body.mjs";
import {
  localQueue,
  rememberBindings,
} from "../../vendor/tbsync/change-queue.mjs";
import {
  CHANGELOG_KINDS,
  SERVER_TAG_STATUSES,
} from "../../vendor/tbsync/changelog-core.mjs";
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
// "Come back later", in both spellings [MS-ASCMD] 2.2.2 gives it: 111
// (ServerErrorRetryLater) is the named one, and 110 (ServerError) is kept
// because real servers send it when they mean busy - the legacy add-on's
// 30-minute pause was built on observing exactly that. Autosync is
// suppressed via `noAutosyncUntil` on the account; a pre-14.0 server says
// the same thing as HTTP 503, which arrives as a thrown wire error and is
// folded into this path by the caller.
const BUSY_STATUSES = new Set(["110", "111"]);

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
// Only statuses that carry no verdict are here. A 5 or a 16 says the server
// hit a fault, not that it judged the command, so nothing on our side has
// changed and the same request a moment later is the right answer.
//
// A 7 is a verdict and is handled where it is read, not here: we declare
// server-wins in every Options block, so a conflict is that policy working
// and re-sending would be asking the server to change its mind.
const INSTANCE_RETRY_DELAY_MS = Object.freeze({
  [STATUS_TEMP_SERVER]: 1000,
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
 *  `browser.storage.local["maxItems"]`; default 25 when unset.
 *
 *  Never below `MIN_MAX_ITEMS`. A window of one or two asks a throttled
 *  service for the same folder in hundreds of requests where a handful
 *  would do, and Exchange Online budgets Sync commands by count: an account
 *  configured that low spends its allowance on paging alone and is answered
 *  HTTP 503 for the rest of the hour. The floor is applied here rather than
 *  in the options page because a value can also arrive from the v4
 *  migration, which no dialog ever sees. It does not constrain the push's
 *  own shrink after a rejected batch - going down to a single item is how
 *  that isolates the item the server refused. */
const MIN_MAX_ITEMS = 5;

async function readMaxItems() {
  const { maxItems } = await browser.storage.local.get({ maxItems: 25 });
  const n = Number(maxItems);
  if (!Number.isFinite(n) || n <= 0) return 25;
  return Math.max(MIN_MAX_ITEMS, n);
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
    let result;
    try {
      result = await runOneSync({
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
    } catch (err) {
      // HTTP 503 is "retry later" said at the transport - the shape
      // pre-14.0 servers use for Status 111, and throttling Exchanges for
      // any version - so it takes the same pause as the in-band statuses
      // instead of failing the folder red. Everything else stays thrown.
      if (err?.status !== 503) throw err;
      result = {
        code: "BUSY",
        httpStatus: 503,
        retryAfterMs: err.retryAfterMs ?? null,
      };
    }
    if (result.code === "RESYNC") {
      // Status 3: the server does not recognise the sync key we sent, so
      // our view of this collection is void. Starting over re-downloads
      // every item in the folder, which is the most expensive thing a sync
      // does and used to happen in silence - a flood of "pull add" lines
      // with nothing to say why. Said out loud, because the alternative is
      // reading it as data churn.
      //
      // Both attempts are announced, as v4 announced every FOLDER_RERUN
      // (core.js:188). `attempt` exists for nothing but this bound - it is
      // v4's `maxFolderReruns = 2` moved into the provider - so the second
      // line is the only trace that a folder gave up rather than settled.
      provider.reportEventLog({
        level: attempt === 0 ? "warning" : "error",
        accountId,
        folderId,
        message:
          `[${itemKind.changelogKind}-sync] the server refused our sync key` +
          (attempt === 0
            ? " - starting this folder over, so everything in it is downloaded again"
            : " again, on the folder we had just started over - giving up on it this run"),
      });
      if (attempt === 0) {
        // Only the sync key. [MS-ASCMD] asks for two things on Status 3 -
        // return to SyncKey 0, and either drop or re-push the items added
        // since the last good sync (the changelog survives this patch, so
        // they go out on the retry). Forgetting which local item each
        // server id names is neither, and the whole folder is about to
        // arrive as <Add> commands that have to match them.
        const reset = { synckey: "0" };
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
      // Twice refused, including once on a folder we had just started from
      // scratch. Nothing synced and nothing here can recover it, so the
      // loop ends at the error below rather than falling through to
      // `result.status ?? ok()` - a RESYNC result carries no status, so
      // that path reported success for a folder that never synced, and
      // left the error meant for this case unreachable. v4 ended the same
      // way, as `ERROR: "resync-loop"` once `maxFolderReruns` was spent
      // (core.js:164-166).
      break;
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
      // The server asked us to come back later - Sync Status 110/111 in
      // the body, or HTTP 503 at the transport. Suppress autosync via the
      // host-recognized top-level `noAutosyncUntil` field, for as long as
      // the server's own Retry-After asked when it sent one; the user can
      // still trigger a manual sync, which retries immediately.
      const pauseMs = result.retryAfterMs ?? RETRY_LATER_BACKOFF_MS;
      const pauseMin = Math.max(1, Math.round(pauseMs / 60_000));
      const via = result.httpStatus
        ? `HTTP ${result.httpStatus}`
        : `Status ${result.topStatus ?? result.collStatus}`;
      await provider
        .updateAccount({
          accountId,
          patch: { noAutosyncUntil: Date.now() + pauseMs },
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
        message: `[${itemKind.changelogKind}-sync] server asks to retry later (${via}); autosync paused for ${pauseMin} min`,
      });
      return warningStatus(
        `Server busy - autosync paused for ${pauseMin} minutes`,
      );
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
    // The folder's uid <-> serverId map, both directions, over what was
    // stored last time. It is the only answer to "which local item is
    // this?" - see `server-id-index.mjs`, and item 50 for what happens
    // when it is lost.
    indexMap: createServerIdIndex(folder.custom?.indexMap),
    // uid -> Set of every ServerId this sync heard the server claim for
    // it. More than one means the server holds the item twice; see
    // `duplicate-uids.mjs` for why only what this sync saw may count.
    uidClaims: new Map(),
    fullPull: false,
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

  // Bank whether this calendar's attendees are ours to notify. The item
  // hooks need the answer and cannot ask the host for it, so the sync -
  // which holds both the target and the negotiated version - leaves it
  // where they can read it. See `client-scheduling.mjs`.
  if (ctx.itemKind.changelogKind === "event") {
    await rememberClientScheduling(
      ctx.targetID,
      versionNeedsClientScheduling(asVersion),
      accountUserAddress(account),
    );
  }

  // Read-only revert pre-step. Drops local edits before the pull so the
  // local store ends up matching the server. ItemOperations.Fetch lets us
  // re-pull a single item by serverId; falls back to a synckey reset when
  // the server didn't advertise ItemOperations (legacy behaviour at
  // sync.js:888-911 `revertLocalChangesViaResync`).
  if (effectiveDownloadOnly) {
    const heavyResetNeeded = await revertLocalChanges(ctx);
    if (heavyResetNeeded) {
      // The sync key goes, the map stays. Re-downloading the collection
      // does not move a single local item, so which one each server id
      // names is as true afterwards as before - and every re-sent <Add>
      // has to find its item or it creates a second copy of it.
      ctx.synckey = "0";
      synckey = "0";
      ctx.syncKeyDirty = true;
    }
  }

  // 1) Bootstrap if needed.
  if (synckey === "0" || !synckey) {
    // Everything the server holds is about to arrive, so what this sync
    // does not see, the folder does not have. That is the one condition
    // under which a stored duplicate finding may be cleared.
    ctx.fullPull = true;
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
    if (boot.code === "BUSY") {
      return await finishWith(ctx, {
        code: "BUSY",
        topStatus: boot.topStatus,
        collStatus: boot.collStatus,
      });
    }
    if (boot.error)
      return await finishWith(ctx, { status: errorStatus(boot.error) });
    ctx.synckey = boot.synckey;
    ctx.syncKeyDirty = true;
  }

  // 1b) Remove server copies the user asked to be rid of, before anything
  // else looks at the collection. Inside the sync rather than beside it:
  // the host serialises syncs per account and defers a request made while
  // one is running, which is what stops an autosync - or a sync armed by
  // an edit - reaching this folder's SyncKey while the deletes are
  // advancing it. Running first also means the pull that follows sees only
  // the copy that stays.
  await duplicateCleanupPhase(ctx);

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
  // Queued edits to meetings somebody else organised. Held back from the
  // push - there is no truthful way to send one - and answered after the
  // pull instead, which is where their ServerId comes from. See
  // `invitationPhase`.
  const invitationAnswers = [];
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
        // Two different situations share this branch and must not share a
        // message: a *known* foreign kind is a real item EAS has no wire
        // format for (the mailing-list case above), while a kind outside
        // the vocabulary is a bug in whatever queued it, and calling it
        // "unsupported" would send the reader to the protocol instead of
        // to the writer.
        const message = CHANGELOG_KINDS.includes(e.kind)
          ? `[${ctx.itemKind.changelogKind}-sync] skipping a ${e.kind} ` +
            `("${e.itemId}"): ActiveSync cannot store one, so it stays local`
          : `[${ctx.itemKind.changelogKind}-sync] skipping a queued edit ` +
            `of "${e.itemId}": unknown changelog kind ` +
            `${JSON.stringify(e.kind)} - a bug in whatever queued it`;
        ctx.provider.reportEventLog({
          level: "warning",
          accountId: ctx.accountId,
          folderId: ctx.folderId,
          message,
        });
        await ctx.queue.remove({
          parentId: e.parentId,
          itemId: e.itemId,
          kind: e.kind,
        });
        continue;
      }
      // A meeting somebody else organised cannot be sent as an Add or a
      // Change *on 16.x*: there the client may state neither the organizer
      // ([MS-ASCAL] 2.2.2.35) nor an attendee's status (2.2.2.5), and the
      // server fills both in with the current user - so what arrives is not
      // their meeting changed but ours, re-invited to everyone on it. The
      // answer is the one thing we may say, and it goes as a
      // MeetingResponse after the pull.
      //
      // Below 16.0 the opposite holds. `AttendeeStatus` is a request
      // element there ("the client MUST NOT include [it] ... when protocol
      // version 16.0 or 16.1 is used"), and we restate the organizer we
      // already hold rather than leaving the server to substitute one. So
      // the answer travels as an ordinary <Change> - which is the only
      // route that lets a user change their mind, because a 14.1 server
      // refuses a second MeetingResponse for the same meeting.
      //
      // A delete is exempt: the user removed it from their calendar, which
      // is a plain <Delete> and says nothing about the meeting itself.
      // There is also nothing left to read - the item is already gone.
      if (
        parseFloat(ctx.asVersion) >= 16 &&
        ctx.itemKind.changelogKind === "event" &&
        e.status !== "deleted_by_user"
      ) {
        const item = await ctx.store.get(e.itemId);
        if (
          item &&
          isReceivedMeeting(item.blob, accountUserAddress(ctx.account))
        ) {
          invitationAnswers.push(e);
          continue;
        }
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
  if (pushed.followUpMasters?.length) {
    const follow = await followUpPhase(ctx, pushed.followUpMasters);
    if (follow.code) return await finishWith(ctx, follow);
    if (follow.status) return await finishWith(ctx, follow);
    instanceFailed += follow.failedCount ?? 0;
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

  // 5) Answers to invitations, after the pull because that is what gives
  // an item its ServerId - MeetingResponse addresses one by RequestId, and
  // an invitation Thunderbird filed from the email has none until we have
  // seen the server's copy. Nothing here waits for the answer to take
  // effect: Exchange applies it asynchronously and the next sync reports it.
  if (invitationAnswers.length) await invitationPhase(ctx, invitationAnswers);

  // 6) The messages we owe the attendees of meetings we organise, on the
  // versions where the server sends none. Last, because it is the only
  // phase that must see the item exactly as this sync left it.
  await organizedMeetingPhase(ctx);

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

    let serverID = ctx.indexMap.serverIdFor(e.itemId);
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
      // Nothing to fetch, so the server's copy cannot be re-read and the
      // user's edit stands locally while the folder is read-only - the
      // two have diverged and only the next full pull settles it. Said
      // out loud, because dropping a queue entry in silence is how a
      // folder quietly stops matching the server.
      ctx.provider.reportEventLog({
        level: "warning",
        accountId: ctx.accountId,
        folderId: ctx.folderId,
        message:
          `[${ctx.itemKind.changelogKind}-sync] cannot revert local item ` +
          `${e.itemId}: it has no ServerId, so the server's copy cannot be ` +
          `fetched; the local edit stays until the next full pull`,
      });
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
      // Plain text, matching the Sync Options BodyPreference - see there
      // for why asking Exchange for HTML makes it rewrite the note.
      bodyType: "1",
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

    // The same judgement the pull makes: this response was asked for plain
    // text, so an item the server holds as HTML arrived flattened here too.
    // Reverting to a flattening would drop the editor's copy of a note the
    // user never touched.
    const reverted = await resolveBody(ctx, properties, serverID);

    const blob = await codec.applicationDataToBlob({
      adNode: reverted.adNode,
      serverID,
      asVersion: ctx.asVersion,
      separator: ctx.separator,
      defaultTimezone: ctx.defaultTimezone,
      msTodoCompat: ctx.msTodoCompat,
      uid: e.itemId,
      userEmail: accountUserAddress(ctx.account),
      eventLog: ctx.eventLog,
      nativePlainText: reverted.nativePlainText,
    });

    await ctx.queue.markServerWrite({
      parentId: ctx.targetID,
      itemId: e.itemId,
      status: SERVER_TAG_STATUSES[1],
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
      ctx.indexMap.set(e.itemId, serverID);
    }

    await ctx.queue.remove({
      parentId: e.parentId,
      itemId: e.itemId,
      kind: e.kind,
    });
  }
  return false;
}

/** Hand any duplicate clusters this sync saw to the provider, which
 *  collects them across the account's folders and offers the cleanup.
 *
 *  Reporting only, and best-effort with it: a calendar that has synced is
 *  a calendar that has synced, and a folder must not go red because the
 *  thing that noticed a mailbox problem tripped over. */
async function duplicateFinding(ctx) {
  if (!ctx.uidClaims.size) return null;
  try {
    const clusters = await duplicateClusters(ctx.uidClaims, {
      serverIdFor: (uid) => ctx.indexMap.serverIdFor(uid),
      titleFor: async (uid) => titleFromBlob((await ctx.store.get(uid))?.blob),
    });
    if (!clusters.length) {
      // A pull that saw the whole folder and found nothing duplicated is
      // proof that an older finding is spent. An incremental one is not,
      // so it leaves what is stored alone.
      return ctx.fullPull ? [] : null;
    }
    const copies = clusters.reduce((n, c) => n + c.surplus.length, 0);
    ctx.eventLog(
      "warning",
      `[${ctx.itemKind.changelogKind}-sync] the server holds ${clusters.length} ` +
        `item(s) more than once - ${copies} surplus copy/copies, ` +
        `offering to remove them`,
    );
    return clusters;
  } catch (err) {
    ctx.eventLog(
      "debug",
      `[${ctx.itemKind.changelogKind}-sync] duplicate scan failed: ${err?.message ?? String(err)}`,
    );
    return null;
  }
}

/**
 * Send the deletes the duplicates window asked for.
 *
 * The selection is read from the folder row, not from the window: the
 * window names UIDs, and which ServerIds those stand for is decided here
 * against the finding the sync made. A UID no longer in the finding is
 * silently nothing to do.
 *
 * The SyncKey stays the runner's. `deleteSurplusCopies` hands each new one
 * back through `persistSyncKey`, and the flush at the end of this sync
 * writes it with everything else, so there is no second writer for it.
 */
async function duplicateCleanupPhase(ctx) {
  const wanted = ctx.folder.custom?.duplicatesPending;
  if (!Array.isArray(wanted) || !wanted.length) return;
  const finding = Array.isArray(ctx.folder.custom?.duplicates)
    ? ctx.folder.custom.duplicates
    : [];
  ctx.eventLog(
    "debug",
    `[${ctx.itemKind.changelogKind}-sync] duplicate cleanup: ` +
      `${wanted.length} requested, ${finding.length} cluster(s) on record`,
  );
  const asked = new Set(wanted);
  const clusters = finding.filter((c) => asked.has(c.uid));
  const serverIds = clusters.flatMap((c) => c.surplus ?? []);

  let outcome = { deleted: 0, failed: [] };
  try {
    if (serverIds.length) {
      outcome = await deleteSurplusCopies({
        account: ctx.account,
        asVersion: ctx.asVersion,
        collectionId: ctx.collectionId,
        className: ctx.itemKind.className,
        filterType: ctx.itemKind.filterType,
        conflict: ctx.conflict,
        synckey: ctx.synckey,
        serverIds,
        // Written through to the folder row, not just banked on `ctx`.
        // The flush at the end of a sync is the usual writer, but it only
        // runs if the sync gets there: a throw in a later phase would
        // leave the row holding the key from before these deletes, while
        // the server has moved several steps past it. That costs a
        // Status 3 and a full re-download of the folder - the one thing a
        // mailbox with hundreds of copies can least afford.
        persistSyncKey: async (synckey) => {
          ctx.synckey = synckey;
          ctx.syncKeyDirty = true;
          await ctx.provider.updateFolder({
            accountId: ctx.accountId,
            folderId: ctx.folderId,
            patch: { custom: { synckey } },
          });
        },
        // Chunk by chunk, so the window that asked for this can say how
        // far it has got. A cluster the size of the one this was written
        // for is twenty-odd requests, and a button that only goes quiet
        // looks like a button that did nothing.
        onProgress: ({ deleted, total }) =>
          ctx.provider.reportDuplicateProgress?.({
            accountId: ctx.accountId,
            folderId: ctx.folderId,
            deleted,
            total,
          }),
      });
    }
    ctx.eventLog(
      "info",
      `[${ctx.itemKind.changelogKind}-sync] removed ${outcome.deleted} of ` +
        `${serverIds.length} surplus server copy/copies`,
    );
  } catch (err) {
    // Whatever was acknowledged before this is gone from the server for
    // good, and the SyncKey that went with it has already been taken. The
    // request stays on the row so the next sync finishes the job.
    ctx.eventLog(
      "warning",
      `[${ctx.itemKind.changelogKind}-sync] removing surplus copies stopped: ${err?.message ?? String(err)}`,
    );
    return;
  }
  for (const f of outcome.failed) {
    ctx.eventLog(
      "warning",
      `[${ctx.itemKind.changelogKind}-sync] the server refused to remove the ` +
        `surplus copy ${f.serverId} (Status ${f.status})`,
    );
  }
  // Both the request and the clusters it answered leave the row together.
  // A refused copy is dropped with them on purpose: it is the server's
  // verdict on that item, and re-offering it every sync would be a prompt
  // the user cannot act on.
  ctx.folder = {
    ...ctx.folder,
    custom: {
      ...ctx.folder.custom,
      duplicates: finding.filter((c) => !asked.has(c.uid)),
      duplicatesPending: [],
    },
  };
  await ctx.provider.updateFolder({
    accountId: ctx.accountId,
    folderId: ctx.folderId,
    patch: {
      custom: {
        duplicates: ctx.folder.custom.duplicates,
        duplicatesPending: [],
      },
    },
  });
}

async function finishWith(ctx, result) {
  const duplicates = await duplicateFinding(ctx);
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
  const custom = {};
  if (ctx.syncKeyDirty) custom.synckey = ctx.synckey;
  if (ctx.indexMap.dirty) custom.indexMap = ctx.indexMap.toArray();
  // The finding lives on the folder row rather than in the provider's
  // memory: it is what the window offers, what the cleanup deletes
  // against, and - since it is stored where every other piece of folder
  // state is - what can be read back without the window.
  if (duplicates !== null) custom.duplicates = duplicates;
  const patch = Object.keys(custom).length ? { custom } : {};
  if (ctx.pendingCount !== undefined) patch.localChanges = ctx.pendingCount;
  if (!Object.keys(patch).length) return result;
  try {
    await ctx.provider.updateFolder({
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      patch,
    });
  } catch (err) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[${ctx.itemKind.changelogKind}-sync] flush failed: ${err?.message ?? String(err)}`,
    });
  }
  // After the row is written, never before: the offer reads the finding
  // back from the folder, so announcing it first would open a window on
  // state that is not there yet.
  if (duplicates?.length) {
    await ctx.provider.offerDuplicateCleanup?.({
      accountId: ctx.accountId,
      folderId: ctx.folderId,
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

  // The server controls termination via the absent <MoreAvailable/> tag,
  // and it is trusted to converge: initial syncs of large folders (10k+
  // items) hit dozens of iterations, so any cap on the count risks a
  // spurious abort mid-pull. Bounded instead by whether the round trip
  // achieved anything - see the check at the foot of the loop.
  for (;;) {
    throwIfCancelled(ctx);
    const sentKey = ctx.synckey;
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
    if (r.code === "BUSY")
      return { code: "BUSY", topStatus: r.topStatus, collStatus: r.collStatus };
    if (r.error) return { status: errorStatus(r.error) };

    let applied = 0;
    if (r.commands) {
      applied = await applyServerCommands(ctx, r.commands);
      itemsDone += applied;
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

    // "More available" and yet nothing arrived and the sync key did not
    // move: the next request would be byte-for-byte the one just sent,
    // against state the server says is unchanged, so the answer can only
    // be the same again. Measured on Exchange Online (#356): 362 items
    // estimated, one delivered, then sixty identical empty windows until
    // the mailbox was answered HTTP 503 with x-ms-asthrottle: SyncCommands
    // for the rest of the hour - a budget counted in commands, so a loop
    // that fetches nothing still spends it all.
    //
    // Both halves are required. An empty window on its own is not a fault:
    // a server walking a change set whose items all fall outside the
    // FilterType can legitimately send nothing while advancing the key,
    // and stopping there would turn one sync of a large folder into many.
    //
    // Treated as the pull being over, exactly like the absent tag: there is
    // nothing more to be had this pass. Nothing is reset either - the key we
    // hold is the one the server last stated, so the next sync resumes from
    // here, costs two requests instead of sixty, and delivers the rest the
    // moment the server starts sending again.
    //
    // Said out loud all the same. A folder that stops short this way looks
    // from the outside like one that finished, and the difference - 362
    // items estimated against one delivered - would otherwise exist nowhere
    // a bug report could reach.
    if (applied === 0 && (r.synckey == null || r.synckey === sentKey)) {
      ctx.eventLog(
        "warning",
        `[${ctx.itemKind.changelogKind}-sync] the server says more items are ` +
          `available for collection ${ctx.collectionId} but sent none, and ` +
          `left the sync key at ${sentKey}; stopping this folder after ` +
          `${itemsDone} of ${itemsTotal} - asking again cannot change the answer`,
      );
      return {};
    }
  }
}

/* ── Push phase ───────────────────────────────────────────────────── */

async function pushPhase(ctx, userEdits) {
  const failedItems = new Set();
  // Recurring masters pushed in this pass whose blob carries overrides.
  // Only 16.1 needs them, and only the calendar codec can express them.
  const instanceMasters =
    ctx.asVersion === "16.1" && ctx.itemKind.codec.listInstanceCommands
      ? []
      : null;
  // The ≤14.x counterpart: masters ADDED this pass whose blob carries
  // exceptions. Their Add went out without the embedded <Exceptions>
  // wrapper (see `suppressExceptions` in the codec) because at least one
  // server family acks an Add-with-exceptions and silently discards the
  // wrapper - for an all-day series, the recurrence with it. Once the
  // ack has assigned a ServerId, followUpPhase re-sends each as one full
  // <Change>, exceptions embedded - the ordinary modify payload, which
  // the same servers keep.
  const followUpMasters = ctx.asVersion !== "16.1" ? [] : null;
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
    {
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
          userEmail: accountUserAddress(ctx.account),
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
    if (r.code === "BUSY")
      return { code: "BUSY", topStatus: r.topStatus, collStatus: r.collStatus };

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
      followUpMasters,
    });
    if (r.commands) await applyServerCommands(ctx, r.commands);
    await mailAnsweredMeetings(ctx, built, failedItems);

    // Modified masters are noted here rather than in applyResponses, which
    // acts on the changes the server refused. A master that did not land
    // has nothing for its exceptions to attach to, so it is held back from
    // the instance phase - but that is decided on the status, never on
    // whether the server mentioned it.
    //
    // [MS-ASCMD] does not say a successful change goes unreported, and
    // servers differ: Exchange lists only failures, Kerio Connect answers
    // every <Change> with <ServerId> and <Status>1. Reading presence as
    // failure meant every successfully-changed master on such a server was
    // skipped, so a moved occurrence produced no instance command at all.
    if (instanceMasters) {
      const rejected = new Set(
        responses.changes
          .filter((node) => {
            const status = readPathFrom(node, ["Status"]);
            return status && status !== STATUS_OK;
          })
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
    followUpMasters: followUpMasters ?? [],
  };
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
/** Answer the invitations the push phase held back.
 *
 *  Each entry is a queued edit to a meeting somebody else organises. The
 *  only thing we may tell the server about one is the user's answer, and
 *  MeetingResponse is how - [MS-ASCMD] 2.2.2.10. It addresses the item by
 *  `RequestId`, which is its ServerId, so this runs after the pull: an
 *  invitation Thunderbird filed from the email has no ServerId at all until
 *  the server's own copy has come down and been adopted onto it.
 *
 *  Nothing here waits. Exchange acknowledges receipt and applies the answer
 *  to the calendar afterwards - measured at absent after 24 s and present
 *  after 34 s - so the item this sync just pulled will usually still show
 *  the old state. The next sync reports it, and holding this one open to
 *  watch for it is what makes a client freeze.
 *
 *  An entry is dropped when it can never be sent, and kept when it might be
 *  sent later. The difference matters: a kept entry lights the needs-sync
 *  badge until it goes, so keeping one that can never go leaves a badge
 *  nobody can clear. */
async function invitationPhase(ctx, entries) {
  // Responding to an invitation from the calendar item is 14.0 and later.
  // Below that the only way is the meeting request in the Inbox, and we do
  // not sync mail - so these can never be sent, at all, ever.
  if (parseFloat(ctx.asVersion) < 14) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[event-sync] dropping ${entries.length} invitation answer(s): ` +
        `ActiveSync ${ctx.asVersion} can only answer from the meeting ` +
        `request in the Inbox, which is not a folder we sync`,
    });
    for (const entry of entries) await dropEntry(ctx, entry);
    return;
  }
  if (!easCommandLikelyAvailable(ctx.account, "MeetingResponse")) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[event-sync] dropping ${entries.length} invitation answer(s): ` +
        `the server does not advertise MeetingResponse`,
    });
    for (const entry of entries) await dropEntry(ctx, entry);
    return;
  }

  for (const entry of entries) {
    throwIfCancelled(ctx);
    // Re-read rather than trust the copy taken before the pull: the pull
    // may have just adopted the server's version onto this item, which is
    // how it gets a ServerId in the first place. `preserveSelfPartstat`
    // is what keeps the user's answer across that.
    const item = await ctx.store.get(entry.itemId);
    if (!item) {
      await dropEntry(ctx, entry);
      continue;
    }
    // The series and every occurrence answered on its own. Answering from
    // the calendar rather than from the message writes an override and
    // leaves the master untouched, so reading only the master misses the
    // ordinary way of answering a recurring invitation.
    const answers = selfUserResponses(
      item.blob,
      accountUserAddress(ctx.account),
    ).filter((a) => {
      // Say nothing the server already knows. Its own ResponseType comes
      // back on every pull, so an answer it has applied stops being resent
      // on the next unrelated edit - and the organizer stops being mailed
      // a reply they have already had. A stamp that disagrees, or is
      // missing entirely, means send: it can cost a duplicate reply, never
      // a lost answer.
      const known = serverKnownPartstat(a.responseType);
      return !known || known !== USERRESPONSE_TO_PARTSTAT[a.userResponse];
    });
    if (!answers.length) {
      // Not an answer: NEEDS-ACTION, an answer the server already has, or
      // an edit to something else on a meeting we cannot push anyway.
      // Nothing to send and nothing owed.
      ctx.eventLog(
        "debug",
        `[event-sync] ${entry.itemId} carries no answer to send; dropping the queued edit`,
      );
      await dropEntry(ctx, entry);
      continue;
    }
    const serverID = ctx.itemKind.codec.readEasServerIdFromBlob(item.blob);
    if (!serverID) {
      // The user answered before we had ever seen the server's copy. It
      // will arrive, be adopted onto this item, and bring an id with it -
      // so this stays queued and goes on the next sync.
      ctx.eventLog(
        "info",
        `[event-sync] holding the answer to ${entry.itemId}: the server has not ` +
          `named this meeting yet, so there is nothing to address a response to`,
      );
      continue;
    }

    // Answering one occurrence needs InstanceId, which is 14.1 and later.
    // Below that there is no way to name an occurrence at all, so such an
    // answer can never be sent - and a queue entry that can never go is a
    // needs-sync badge that can never clear. It is dropped, out loud,
    // exactly as the two version and command checks above do it. The user
    // is told what to do instead, because the answer really is gone.
    let sendable = answers;
    if (parseFloat(ctx.asVersion) < 14.1) {
      const perOccurrence = answers.filter((a) => a.instanceId);
      if (perOccurrence.length) {
        ctx.provider.reportEventLog({
          level: "warning",
          accountId: ctx.accountId,
          folderId: ctx.folderId,
          message:
            `[event-sync] dropping ${perOccurrence.length} answer(s) to single ` +
            `occurrences of ${entry.itemId}: ActiveSync ${ctx.asVersion} cannot ` +
            `name an occurrence, so answer the whole series instead`,
        });
      }
      sendable = answers.filter((a) => !a.instanceId);
      if (!sendable.length) {
        await dropEntry(ctx, entry);
        continue;
      }
    }

    let allLanded = true;
    let lastStatus = null;
    // The id the server moves the item to. [MS-ASCMD] 3.1.5.6: on an accept
    // or a tentative accept the server "will add or update the corresponding
    // calendar item and return its server ID in the CalendarId element", and
    // the client "updates that item with the returned server ID". A decline
    // carries none, because there the server deletes the calendar item.
    let calendarId = null;
    for (const answer of sendable) {
      throwIfCancelled(ctx);
      const result = await sendMeetingResponse({
        account: ctx.account,
        asVersion: ctx.asVersion,
        collectionId: ctx.collectionId,
        serverID,
        userResponse: answer.userResponse,
        instanceId: answer.instanceId,
      });
      const answerStatus = result?.status ?? null;
      if (answerStatus === "1" && result?.calendarId) {
        calendarId = result.calendarId;
      }
      const what = answer.instanceId
        ? `the occurrence on ${answer.rid}`
        : "the series";
      if (answerStatus === "1") {
        ctx.eventLog(
          "info",
          `[event-sync] answered ${what} of ${entry.itemId} with ` +
            `UserResponse ${answer.userResponse}`,
        );
        await mailTheOrganizer(ctx, entry, item, answer);
      } else {
        allLanded = false;
        lastStatus = answerStatus;
        ctx.eventLog(
          "warning",
          `[event-sync] the server refused the answer to ${what} of ` +
            `${entry.itemId} with Status ${answerStatus}`,
        );
      }
    }
    // Take the server at its word about where the item now lives, rather
    // than waiting for the pull to say the same thing. It arrives as an
    // <Add> under the new id plus a <Delete> for the old one, in that
    // order and in one response, so until the map is re-pointed the delete
    // resolves to the item the add has just landed on and takes it. Doing
    // it here also means the item is addressable again straight away,
    // without a round trip.
    if (calendarId && calendarId !== serverID) {
      const fresh = await ctx.store.get(entry.itemId);
      if (fresh) {
        const restamped = ctx.itemKind.codec.stampEasServerId(
          fresh.blob,
          calendarId,
        );
        await ctx.queue.markServerWrite({
          parentId: ctx.targetID,
          itemId: entry.itemId,
          status: SERVER_TAG_STATUSES[1],
          kind: ctx.itemKind.changelogKind,
        });
        await ctx.store.update(entry.itemId, restamped);
        ctx.indexMap.set(entry.itemId, calendarId);
        ctx.eventLog(
          "info",
          `[event-sync] the answer moved ${entry.itemId} to ${calendarId}; ` +
            `re-stamped it rather than waiting for the pull to say so`,
        );
      }
    }

    // The queued edit stands for every answer on the item, so it is only
    // spent once they have all gone. One refusal keeps it, and the ones
    // that did land are not sent again: the server's own ResponseType
    // comes back on the next pull and filters them out above.
    if (allLanded) {
      await dropEntry(ctx, entry);
      continue;
    }
    const status = lastStatus;
    // 2 is an invalid meeting request and 4 a meeting that cannot be
    // responded to - a cancelled one, or somebody else's delegate. Neither
    // improves by being retried, and a queue entry that can never go is a
    // badge that can never clear. 3 is a server-side sync-state problem
    // that does improve, so that one waits.
    if (status === "2" || status === "4") {
      ctx.provider.reportEventLog({
        level: "warning",
        accountId: ctx.accountId,
        folderId: ctx.folderId,
        message:
          `[event-sync] the server refused the answer to ${entry.itemId} ` +
          `(Status ${status}); it will not be retried`,
      });
      await dropEntry(ctx, entry);
      continue;
    }
    ctx.eventLog(
      "warning",
      `[event-sync] could not answer the invitation ${entry.itemId} ` +
        `(${status ? `Status ${status}` : "no response"}); leaving it queued`,
    );
  }
}

/** Take a queued edit out of the queue, having established that nothing
 *  more can be done with it. */
/** After a push pass below 16.0, tell the organiser about any answer that
 *  just went out as an `AttendeeStatus`.
 *
 *  The push is what carries the answer there, so this is the equivalent of
 *  the MeetingResponse-then-SendMail order the spec describes: never before
 *  the server has taken it, and only for what actually landed.
 *
 *  Only an answer the server did not already know is worth a message, which
 *  is the same test the response phase uses - otherwise every later edit to
 *  an answered meeting mails the organiser again. */
async function mailAnsweredMeetings(ctx, sent, failedItems) {
  if (parseFloat(ctx.asVersion) >= 16) return;
  if (ctx.itemKind.changelogKind !== "event") return;
  const userEmail = accountUserAddress(ctx.account);
  if (!userEmail) return;

  // `sent` is the built batch: { adds, mods, dels }, each carrying the
  // changelog entry and the item as it went out. A delete says nothing
  // about the meeting, so only adds and mods are of interest.
  for (const built of [...(sent?.adds ?? []), ...(sent?.mods ?? [])]) {
    const entry = built.entry;
    if (!entry || failedItems?.has?.(entry.itemId)) continue;
    const item = await ctx.store.get(entry.itemId);
    if (!item || !isReceivedMeeting(item.blob, userEmail)) continue;
    const answer = selfUserResponses(item.blob, userEmail).find(
      (a) => !a.instanceId,
    );
    if (!answer) continue;
    // Not the server's ResponseType: below 16.0 it never leaves 5, because
    // the reply is ours to send and the server never sees one. Our own
    // record of what the organiser was last told is the only thing that
    // can answer "have they heard this already?".
    const want = USERRESPONSE_TO_PARTSTAT[answer.userResponse];
    if (repliedPartstatOf(item.blob) === want) continue;
    const told = await mailTheOrganizer(ctx, entry, item, answer);
    if (!told) continue;
    const fresh = await ctx.store.get(entry.itemId);
    if (!fresh) continue;
    await ctx.queue.markServerWrite({
      parentId: ctx.targetID,
      itemId: entry.itemId,
      status: SERVER_TAG_STATUSES[1],
      kind: ctx.itemKind.changelogKind,
    });
    await ctx.store.update(
      entry.itemId,
      stampRepliedPartstat(fresh.blob, want),
    );
  }
}

/**
 * Send what the organiser owes the attendees, for the notes this sync's
 * edits left behind.
 *
 * **After the push and the pull**, which is the only correct point: the pull
 * is where a push that lost a `Status 7` conflict is replaced by the
 * server's winning copy, so a message built any earlier could announce a
 * change the server threw away. Comparing the settled item against the note
 * makes every no-op case disappear without anything having to recognise it.
 *
 * One attempt each. A message we could not send is a warning in the event
 * log and nothing else - never a retry, never a failed folder - and the note
 * is dropped either way, so nothing can accumulate that never clears. The
 * cost is that a transport failure loses that notification for good; the
 * user's recovery is to save the meeting again.
 */
async function organizedMeetingPhase(ctx) {
  if (ctx.itemKind.changelogKind !== "event") return;
  if (!versionNeedsClientScheduling(ctx.asVersion)) return;
  const notes = await ctx.queue.sendMailPending();
  if (!notes.length) return;
  const userEmail = accountUserAddress(ctx.account);
  if (!userEmail) return;

  // A note whose edit is still queued belongs to a push that did not land,
  // so the server does not have the change yet. Leaving it costs a sync;
  // sending would announce something the server never accepted, and would
  // spend the single attempt doing it.
  const stillQueued = new Set((await ctx.queue.pending()).map((e) => e.itemId));

  for (const note of notes) {
    if (stillQueued.has(note.itemId)) continue;
    try {
      await sendOwedMessages(ctx, note, userEmail);
    } catch (err) {
      ctx.eventLog(
        "warning",
        `[event-sync] could not tell the attendees of ${note.itemId}: ` +
          `${err?.message ?? String(err)}`,
      );
    }
    await ctx.queue.removeSendMail({
      parentId: note.parentId,
      itemId: note.itemId,
      kind: note.kind,
    });
  }
}

/** The messages one note turns into, and their sending. */
async function sendOwedMessages(ctx, note, userEmail) {
  const item = await ctx.store.get(note.itemId);
  // Gone during the pull: the server deleted it, so there is no meeting to
  // announce and a cancellation is not ours to invent.
  if (!item?.blob) return;
  if (isReceivedMeeting(item.blob, userEmail)) return;

  const now = announceableOf(item.blob);
  if (!now) return;
  const from = note.detail?.from ?? null;
  const isInvitation = note.status === "added_for_sendMail";

  if (!isInvitation && from && sameAnnounceable(from, now)) {
    ctx.eventLog(
      "debug",
      `[event-sync] ${note.itemId} settled back to what the attendees ` +
        `already hold; nothing sent`,
    );
    return;
  }

  const dropped = droppedAttendees(from, now);
  const send = async (method, recipients) => {
    if (!recipients.length) return;
    const mime = buildMeetingRequestMime({
      blob: item.blob,
      method,
      recipients,
      userEmail,
      userName: ctx.account.custom?.userDisplayName ?? "",
      now: new Date(),
    });
    if (!mime) return;
    await sendMail({
      account: ctx.account,
      asVersion: ctx.asVersion,
      mime,
      clientId: `eas-${method.toLowerCase()}-${Date.now().toString(36)}`,
    });
    ctx.eventLog(
      "info",
      `[event-sync] told ${recipients.length} attendee(s) of ` +
        `${note.itemId}: ${method}`,
    );
  };

  // A cancelled meeting is cancelled for everyone who was ever told about
  // it, including anyone dropped in the same edit.
  if (now.status === "CANCELLED") {
    await send("CANCEL", [...new Set([...now.attendees, ...dropped])]);
    return;
  }
  await send("CANCEL", dropped);
  await send("REQUEST", now.attendees);
}

/** Tell the organiser, on the versions where that is the client's job.
 *
 *  [MS-ASCMD]: the SendMail step "applies only to protocol versions 2.5,
 *  12.0, 12.1, 14.0, and 14.1" - from 16.0 the server generates the reply.
 *  So below 16.0 a Status 1 means the user's own calendar was updated and
 *  nothing else: without this the organiser is never told at all.
 *
 *  After the response and never before, as the guidance requires, so the
 *  invitee's calendar and what the organiser has been told cannot disagree.
 *  Best-effort: a reply we could not send is worth a warning, not a failed
 *  folder, and never a retry - the answer itself has already landed, and
 *  re-running the response would be a second MeetingResponse.
 */
async function mailTheOrganizer(ctx, entry, item, answer) {
  if (parseFloat(ctx.asVersion) >= 16) return false;
  // An occurrence is answered against the series it belongs to, and the
  // reply carries no RECURRENCE-ID, so one message per item is right.
  if (answer.instanceId) return false;
  const userEmail = accountUserAddress(ctx.account);
  if (!userEmail) return false;

  const mime = buildMeetingResponseMime({
    blob: item.blob,
    userResponse: answer.userResponse,
    userEmail,
    userName: ctx.account.custom?.userDisplayName ?? "",
    now: new Date(),
  });
  if (!mime) return false;

  try {
    const status = await sendMail({
      account: ctx.account,
      asVersion: ctx.asVersion,
      mime,
      clientId: `tbsync-${crypto.randomUUID()}`,
    });
    if (status === null || status === "1") {
      ctx.eventLog(
        "info",
        `[event-sync] told the organiser of ${entry.itemId} - ActiveSync ` +
          `${ctx.asVersion} leaves that to the client`,
      );
      return true;
    }
    ctx.eventLog(
      "warning",
      `[event-sync] the answer to ${entry.itemId} reached the server, but ` +
        `the reply to the organiser was refused (SendMail Status ${status})`,
    );
    return false;
  } catch (err) {
    ctx.eventLog(
      "warning",
      `[event-sync] the answer to ${entry.itemId} reached the server, but ` +
        `the reply to the organiser could not be sent: ${err?.message ?? err}`,
    );
  }
  return false;
}

async function dropEntry(ctx, entry) {
  await ctx.queue.remove({
    parentId: entry.parentId,
    itemId: entry.itemId,
    kind: entry.kind,
  });
}

async function instancePhase(ctx, masters) {
  const commands = [];
  for (const m of masters) {
    const built = ctx.itemKind.codec.listInstanceCommands({
      blob: m.blob,
      serverID: m.serverID,
      previous: m.previous ?? null,
      asVersion: ctx.asVersion,
      defaultTimezone: ctx.defaultTimezone,
      userEmail: accountUserAddress(ctx.account),
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

/** ≤14.x counterpart of the instance phase: for each master whose Add
 *  deliberately went out bare, send one full-payload <Change> - the
 *  ordinary modify writer, exceptions embedded - now that the ack has
 *  assigned the ServerId it needs. One command per request, through
 *  `sendInstanceCommand`, which brings the status handling and the
 *  bounded retries with it; a refused follow-up leaves the master synced
 *  and its exceptions absent, counted into the folder's rejected total -
 *  the same trade the 16.1 instance phase makes. */
async function followUpPhase(ctx, masters) {
  let failedCount = 0;
  for (const m of masters) {
    const command = {
      kind: "change",
      serverID: m.serverID,
      instanceId: "(embedded exceptions)",
      emit(builder) {
        builder.otag("Change");
        builder.atag("ServerId", m.serverID);
        builder.otag("ApplicationData");
        ctx.itemKind.codec.appendApplicationDataFromBlob({
          builder,
          op: "change",
          blob: m.blob,
          asVersion: ctx.asVersion,
          separator: ctx.separator,
          defaultTimezone: ctx.defaultTimezone,
          userEmail: accountUserAddress(ctx.account),
          fallbackOrganizerName:
            ctx.account?.custom?.fallbackOrganizerNames?.[ctx.collectionId],
          eventLog: ctx.eventLog,
        });
        builder.switchpage("AirSync");
        builder.ctag();
        builder.ctag();
      },
    };
    const r = await sendInstanceCommand(ctx, command, m.blob);
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
 *  One budget, spent by any retry whatever its cause, so nothing has to
 *  classify a retry in order to count it. Which statuses are worth retrying
 *  at all, and how long to wait first, is `INSTANCE_RETRY_DELAY_MS`. */
export async function sendInstanceCommand(ctx, command, blob, attempt = 0) {
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
  if (r.code === "BUSY")
    return { code: "BUSY", topStatus: r.topStatus, collStatus: r.collStatus };
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

  // This request asked for no changes, so a reply carrying <Commands> is
  // not expected. It is still applied if one arrives, and that is not
  // belt-and-braces: the synckey taken above tells the server we have
  // them, so anything dropped here is dropped for good.
  if (r.commands) await applyServerCommands(ctx, r.commands);

  // We sent exactly one command, so at most one response node concerns
  // us - and only failures are reported at all. A moved occurrence goes
  // out as <Change>, a cancelled one as <Delete>.
  const node =
    (r.responses?.changes ?? [])[0] ?? (r.responses?.deletes ?? [])[0] ?? null;
  const status = node ? readPathFrom(node, ["Status"]) : null;
  if (!status || status === STATUS_OK) return { failed: false };

  // Server-wins conflict, the policy we declare in every Options block, so
  // retiring the command here is what "the server wins" means. Not a
  // failure: the sync did what it was told to do.
  //
  // The server's copy is not in this reply - the request asked for no
  // changes - and does not need to be. The pull at the end of this same
  // sync brings it, which is where every other server-side truth arrives.
  // The user sees their edit undone, and making it again is a new edit the
  // next push carries like any other.
  if (status === STATUS_CONFLICT) {
    // Warning, not info: [MS-ASCMD] 2.2.3.177.17 resolves a 7 with "inform
    // the user that the change they made to the item has been overwritten
    // by a server change", and the user's own edit has just been undone.
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[${ctx.itemKind.changelogKind}-sync] ${label}: the server's copy ` +
        `won (Status 7); it has been taken and the command dropped`,
    });
    return { failed: false };
  }

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
        ctx.indexMap.serverIdFor(entry.itemId);
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
      const serverID = ctx.indexMap.serverIdFor(entry.itemId);
      if (!serverID) {
        // The map cannot say what the server calls this item, and the
        // local item is already gone, so its blob cannot either. Drop the
        // changelog entry; the server keeps its copy. Nothing re-derives
        // the mapping - recovery needs the server to offer the item as an
        // Add again, which puts it back through `applyAdd`. Common causes:
        // the item was created and deleted between syncs, so the server
        // never heard of it at all, or the map has lost an entry it should
        // have (item 50 - worth chasing if it recurs).
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

export async function applyResponses(
  ctx,
  responses,
  sent,
  failedItems,
  opts = {},
) {
  const {
    hadResponsesElement = true,
    instanceMasters = null,
    followUpMasters = null,
  } = opts;
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
    if (status === STATUS_CONFLICT) {
      // The server already holds an item matching this one. Unlike a
      // conflict on a <Change>, the response does not say which: an Add
      // response names the item by our own ClientId, and a refused Add
      // carries no ServerId to adopt. So there is no address to turn this
      // into a change, and no way to learn one from this reply.
      //
      // That leaves nothing a retry can improve - the identical Add draws
      // the identical refusal on every later sync - so the entry is
      // retired. Only the queued push is: the item itself is untouched and
      // stays exactly as the user left it. What it does not have is a
      // server identity, so it stays local-only until something gives it
      // one, and that is why this is a warning rather than a silent drop.
      //
      // Kept out of `failedItems` on purpose: that set re-stages the entry
      // for the next sync, which is the one thing this branch has decided
      // against. It has a second reader, though - below 16.0
      // `mailAnsweredMeetings` skips whatever is in it - so an answered
      // invitation retired here still reaches the organiser. That is the
      // right way round: the user did accept, and below 16.0 telling the
      // organiser is ours to do because the server never does. What stays
      // behind is the calendar item, not the answer.
      reportRejectedPushItem(
        ctx,
        "add",
        status,
        sentEntry,
        "warning",
        "the server already holds a matching item, and a refused add names " +
          "no ServerId to adopt, so the queued add has been retired",
      );
      await ctx.queue.remove({
        parentId: sentEntry.entry.parentId,
        itemId: sentEntry.entry.itemId,
        kind: sentEntry.entry.kind,
      });
      continue;
    }
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
      status: SERVER_TAG_STATUSES[1],
      kind: ctx.itemKind.changelogKind,
    });
    await ctx.store.update(sentEntry.item.id, stamped);
    // Register the just-stamped item in the indexMap so any follow-up
    // server-pushed Change for this ServerID matches the existing
    // local item via applyChangeFromAd instead of falling through to
    // applyAdd and creating a duplicate.
    ctx.indexMap.set(sentEntry.item.id, serverId);
    // On 16.1 an exception is not part of the master's payload - it is a
    // separate <Change> keyed on the master's ServerId, which only exists
    // once the server has acked this Add. Note the pair down for the
    // instance phase; nothing can send them before this point.
    if (instanceMasters && blobHasInstanceOverrides(sentEntry.item.blob)) {
      // No `previous`: the server has just learned about this item, so
      // every exception it carries is new to it.
      instanceMasters.push({ serverID: serverId, blob: sentEntry.item.blob });
    }
    // ≤14.x mirror of the above: the Add deliberately went out without
    // its exceptions, and only now is there a ServerId to hang them on.
    if (followUpMasters && blobHasInstanceOverrides(sentEntry.item.blob)) {
      followUpMasters.push({ serverID: serverId, blob: sentEntry.item.blob });
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
    if (status === STATUS_CONFLICT) {
      // Server-wins conflict. Dropping the local edit is the whole point of
      // the policy we declared, and the server's version arrives on the
      // pull, so there is nothing to keep and nothing to warn about.
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
    if (status === STATUS_OBJECT_NOT_FOUND) {
      // The server has no such item, so - unlike a conflict - nothing is
      // coming down to replace what the user just wrote. Dropping the entry
      // here would lose the edit *and* leave the item sitting locally with
      // a ServerId that names nothing.
      //
      // They edited it, which is as clear a statement as we get that they
      // want it. So it is re-queued as an add and the next push re-creates
      // it: `created` over a `modified_by_user` entry is exactly the "late
      // create after modify" transition the changelog already defines. The
      // stale stamp and index entry are corrected when that add is acked.
      //
      // This is the one signal of a locally-held item the server has
      // dropped that costs nothing to collect - it arrives on its own, and
      // only for an item the user cared enough to touch.
      const serverId = readPathFrom(node, ["ServerId"]);
      const sentEntry = sent.mods.find((m) => m.serverID === serverId);
      if (sentEntry) {
        await ctx.queue.record({
          parentId: sentEntry.entry.parentId,
          itemId: sentEntry.entry.itemId,
          kind: sentEntry.entry.kind,
          op: "created",
        });
        ctx.indexMap.remove(sentEntry.entry.itemId);
        ctx.eventLog(
          "info",
          `[${ctx.itemKind.changelogKind}-sync] the server no longer has ` +
            `${sentEntry.entry.itemId}; re-creating it from the local copy ` +
            `rather than discarding the edit`,
        );
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
      ctx.indexMap.remove(sentEntry.entry.itemId);
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
    // Status 7/8: the queue entry was already dropped in the per-response
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
      ctx.indexMap.remove(d.entry.itemId);
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

  // Items whose NativeBodyType says the server holds an HTML note the Sync
  // payload flattened. They are applied like everything else - the
  // flattening becomes the note, so nothing is held back - and their rich
  // form is fetched afterwards, for the whole window in one request instead
  // of one request per item mid-loop. Keyed by ServerId so an add followed
  // by a change in the same window records once; local to this call so the
  // early returns in the phases above us can never drop it silently.
  const noteBacklog = new Map();

  for (const node of commands.adds) {
    await applyAdd(ctx, node, noteBacklog);
    processed++;
  }
  for (const node of commands.changes) {
    await applyChange(ctx, node, noteBacklog);
    processed++;
  }
  for (const node of commands.deletes) {
    await applyDelete(ctx, node, noteBacklog);
    processed++;
  }
  for (const node of commands.softDeletes) {
    await applyDelete(ctx, node, noteBacklog);
    processed++;
  }
  await upgradeNoteBacklog(ctx, noteBacklog);
  return processed;
}

/** Fetch the rich form of every backlogged note in one request per chunk and
 *  re-apply those items as ordinary changes. Every item already exists with
 *  the server's own flattening as its note, so failure at any point - the
 *  gate, a transport error, a missing result - costs formatting freshness
 *  only, never data: the remainder simply keeps the plain text, with one
 *  warning naming the count rather than one per item. */
async function upgradeNoteBacklog(ctx, noteBacklog) {
  if (noteBacklog.size === 0) return;

  const entries = [...noteBacklog.values()];
  if (!easCommandLikelyAvailable(ctx.account, "ItemOperations")) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[${ctx.itemKind.changelogKind}-sync] server does not offer ` +
        `ItemOperations; ${entries.length} HTML note(s) stay as plain text`,
    });
    return;
  }

  // A window is normally at most maxItems anyway; the cap is for a
  // user-raised maxItems, where 25 HTML bodies is already a sizeable
  // response. No shrink ladder: the request carries only ids, so there is
  // no poisonous item to bisect for, and shrinking would re-create the very
  // burst this batching exists to prevent.
  const chunkSize = Math.min(ctx.maxItems, 25);
  for (let at = 0; at < entries.length; at += chunkSize) {
    throwIfCancelled(ctx);
    const chunk = entries.slice(at, at + chunkSize);
    let result = null;
    let failure = null;
    try {
      result = await fetchServerItems({
        account: ctx.account,
        asVersion: ctx.asVersion,
        collectionId: ctx.collectionId,
        serverIDs: chunk.map((e) => e.serverID),
        bodyType: "2",
      });
    } catch (err) {
      failure = err;
    }

    // A failed request - thrown, or answered with a non-1 Status - ends the
    // upgrade for everything still waiting. The notes are already stored as
    // the server's own plain text, so there is nothing to repair and nothing
    // to retry; asking again is how the burst this exists to prevent comes
    // back.
    if (!result || (result.status && result.status !== "1")) {
      const remaining = entries.length - at;
      ctx.provider.reportEventLog({
        level: "warning",
        accountId: ctx.accountId,
        folderId: ctx.folderId,
        message:
          `[${ctx.itemKind.changelogKind}-sync] could not fetch the HTML form ` +
          `of ${remaining} note(s); they keep the server's plain-text rendering`,
        details: failure
          ? String(failure?.message ?? failure)
          : `ItemOperations Status ${result?.status}`,
      });
      return;
    }

    for (const entry of chunk) {
      const properties = result.items.get(entry.serverID);
      if (!properties) {
        ctx.provider.reportEventLog({
          level: "warning",
          accountId: ctx.accountId,
          folderId: ctx.folderId,
          message:
            `[${ctx.itemKind.changelogKind}-sync] could not fetch the HTML note ` +
            `for ${entry.serverID}; it keeps the server's plain-text rendering`,
        });
        continue;
      }
      // Resolved now, not when the entry was recorded: pass 1 wrote this
      // item moments ago, and for an add the row did not exist then.
      const existing = await findExistingByServerId(ctx, entry.serverID);
      if (!existing) continue;
      await applyChangeFromAd(ctx, properties, existing, entry.serverID, {
        adNode: properties,
        nativePlainText: entry.flattened,
      });
    }
  }
}

async function applyAdd(ctx, addNode, noteBacklog = null) {
  const serverID = readPathFrom(addNode, ["ServerId"]);
  if (!serverID) {
    // Skipped, and the sync key still advances - the server will not
    // offer this item again, so it is simply missing locally from here
    // on. Nothing can be done about it, but it must not be invisible.
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[${ctx.itemKind.changelogKind}-sync] server <Add> without a ServerId, skipped - the item will be missing locally`,
    });
    return;
  }
  let ad = childByTag(addNode, "ApplicationData");
  if (!ad) return;
  // Read off the wire, before the body is resolved and before anything is
  // built from it: an item shaped like this must not enter the calendar at
  // all. The sync key still advances, so the server will not offer it
  // again until it changes, and what never arrives cannot be edited into
  // something we would have to send back.
  const refusal = ctx.itemKind.codec.serverRejectReason?.({ adNode: ad });
  if (refusal) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message:
        `[${ctx.itemKind.changelogKind}-sync] skipping ${serverID} sent by ` +
        `the server: ${refusal} - it stays out of the calendar`,
    });
    return;
  }
  // A note the server holds as HTML is applied with the flattening in hand
  // and upgraded after the loop, one request for the whole window; the
  // resolver's inline fetch is for callers with no window to batch over.
  let resolved;
  if (noteBacklog && needsHtmlRefetch(ctx, ad)) {
    recordNoteForUpgrade(noteBacklog, ad, serverID);
    resolved = { adNode: ad, nativePlainText: null };
  } else {
    resolved = await resolveBody(ctx, ad, serverID);
  }
  ad = resolved.adNode;
  await maybeRecordFallbackOrganizerName(ctx, ad);
  const existing = await findExistingByServerId(ctx, serverID);
  if (existing) {
    noteUidClaim(ctx.uidClaims, existing.itemId, serverID);
    return applyChangeFromAd(ctx, ad, existing, serverID, resolved);
  }

  // The UID the server gave this item. Every mailbox holding the same
  // meeting names it the same way, and it is the one identifier that
  // survives the server deleting and re-creating an item when somebody
  // answers an invitation - the ServerId does not. So it is preferred to a
  // minted one, and Thunderbird's item ids are iCal UIDs, which makes the
  // next lookup possible at all.
  const serverUid = readPathFrom(ad, ["UID"]);
  noteUidClaim(ctx.uidClaims, serverUid, serverID);
  const kind = ctx.itemKind.changelogKind;
  let newId;
  if (serverUid) {
    newId = serverUid;
    ctx.eventLog(
      "debug",
      `[${kind}-sync] pull add ${serverID}: the server named this item ` +
        `${serverUid}, and that is the id it gets locally`,
    );
  } else {
    newId = crypto.randomUUID();
    ctx.eventLog(
      "debug",
      `[${kind}-sync] pull add ${serverID}: the server sent no UID, so the ` +
        `item gets a minted id, ${newId}`,
    );
  }

  // Thunderbird's own copy, filed from the invitation email before we ever
  // synced this folder: same meeting, same UID, no ServerId. Adopting it is
  // what stops the user ending up with the invitation they accepted and a
  // second copy of it that we added beside it.
  if (serverUid) {
    const twin = await ctx.store.get(serverUid);
    if (twin) {
      ctx.eventLog(
        "info",
        `[${ctx.itemKind.changelogKind}-sync] adopting the local copy of ${serverUid} ` +
          `instead of adding a second (server id ${serverID})`,
      );
      return applyChangeFromAd(
        ctx,
        ad,
        { itemId: serverUid, blob: twin.blob },
        serverID,
        resolved,
      );
    }
  }
  const blob = await ctx.itemKind.codec.applicationDataToBlob({
    adNode: ad,
    serverID,
    asVersion: ctx.asVersion,
    separator: ctx.separator,
    defaultTimezone: ctx.defaultTimezone,
    msTodoCompat: ctx.msTodoCompat,
    uid: newId,
    userEmail: accountUserAddress(ctx.account),
    eventLog: ctx.eventLog,
    nativePlainText: resolved.nativePlainText,
  });
  await ctx.queue.markServerWrite({
    parentId: ctx.targetID,
    itemId: newId,
    status: SERVER_TAG_STATUSES[0],
    kind: ctx.itemKind.changelogKind,
  });
  const createdId = await ctx.store.create(newId, blob);
  if (createdId !== newId) {
    throw new Error(
      `store.create id mismatch: expected ${newId}, got ${createdId}`,
    );
  }
  await verifyRoundTrip(ctx, newId, blob, "create");
  ctx.indexMap.set(newId, serverID);
  if (blobHasRecurrence(blob)) {
    logRecurrence(ctx, `pull add: itemId=${newId}, serverID=${serverID}`, {
      ical: blob,
    });
  }
}

async function applyChange(ctx, changeNode, noteBacklog = null) {
  const serverID = readPathFrom(changeNode, ["ServerId"]);
  if (!serverID) {
    ctx.provider.reportEventLog({
      level: "warning",
      accountId: ctx.accountId,
      folderId: ctx.folderId,
      message: `[${ctx.itemKind.changelogKind}-sync] server <Change> without a ServerId, skipped - the local copy stays as it is`,
    });
    return;
  }
  const ad = childByTag(changeNode, "ApplicationData");
  if (!ad) return;
  await maybeRecordFallbackOrganizerName(ctx, ad);
  const existing = await findExistingByServerId(ctx, serverID);
  if (!existing) return declineChangeForUnknownItem(ctx, serverID);
  noteUidClaim(ctx.uidClaims, existing.itemId, serverID);
  // 16.1 per-instance Change: ApplicationData carries <InstanceId> and
  // is scoped to a single occurrence of the master event referenced by
  // ServerId. Route to the codec's exception path; bail back to the
  // normal master update if the codec doesn't support it.
  const instanceId = readPathFrom(ad, ["InstanceId"]);
  if (instanceId && ctx.itemKind.codec.applyInstanceChange) {
    return applyExceptionChange(ctx, ad, existing, instanceId, serverID);
  }
  return applyChangeFromAd(ctx, ad, existing, serverID, null, noteBacklog);
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
    nextBlob = await codec.applyInstanceChange?.({
      ical: existing.blob,
      adNode: ad,
      instanceUtc,
      asVersion: ctx.asVersion,
      defaultTimezone: ctx.defaultTimezone,
      userEmail: accountUserAddress(ctx.account),
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
    status: SERVER_TAG_STATUSES[1],
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
    ctx.indexMap.set(existing.itemId, masterServerId);
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

/** The (NativeBodyType << 4 | Type) pair, normalised to the four cases the
 *  note logic distinguishes. A missing Type is 2.5, where Body is plain and
 *  has no Type at all; a server that reports no NativeBodyType makes no
 *  claim, so the payload is taken at face value; RTF folds to plain, because
 *  we cannot decompress it and a conversion we stored would be pushed back
 *  as the note. */
function bodyCase(ad) {
  const rawType = readPathFrom(ad, ["Body", "Type"]);
  const rawNative = readPathFrom(ad, ["NativeBodyType"]);
  const type = rawType === "2" ? 2 : 1;
  const native = rawNative == null ? type : rawNative === "2" ? 2 : 1;
  return (native << 4) | type;
}

/** True when the server holds this item's note as HTML but the payload in
 *  hand is a flattening - the one case worth a second request. Contacts are
 *  excluded outright: a vCard NOTE is text and nothing else, so fetching
 *  HTML for one would put markup in the note and push it back as the
 *  server's plain-text note, destroying the real one. */
function needsHtmlRefetch(ctx, ad) {
  if (ctx.itemKind.className === "Contacts") return false;
  return bodyCase(ad) === 0x21;
}

/** Remember an item for the post-loop note upgrade. The flattening is
 *  captured now, from this payload, so the tooltip can keep the server's own
 *  rendering whatever the fetch later returns. Last write wins - an add
 *  followed by a change in one window records once, with the newer text. */
function recordNoteForUpgrade(noteBacklog, ad, serverID) {
  if (!serverID) return;
  noteBacklog.set(serverID, {
    serverID,
    flattened: readPathFrom(ad, ["Body", "Data"]),
  });
}

/** The note in the form the server actually holds it.
 *
 *  A server answers in the `Type` the request asked for whatever it stores,
 *  and reports what it really has separately, in `NativeBodyType`
 *  ([MS-ASAIRS] 2.2.2.32). The two only disagree when the server converted
 *  the body to satisfy us - so `Type` says how the note arrived and
 *  `NativeBodyType` says what it is.
 *
 *  Returns the ApplicationData to decode from, plus the server's own
 *  flattening of a rich note when we were given one: it is text we were
 *  handed rather than text we computed, so it beats a local conversion for
 *  the tooltip, and it costs nothing because the first response carried it.
 */
async function resolveBody(
  ctx,
  ad,
  serverID,
  refetched = false,
  plainSoFar = null,
) {
  // Only the note-bearing calendar classes can hold HTML. A vCard NOTE is
  // text and nothing else - `readNote` stores whatever Data arrives without
  // reading Type, and the contact writer always emits Type 1 - so fetching
  // HTML for a contact would put markup in the note and push it back as the
  // server's plain-text note, destroying the real one.
  if (ctx.itemKind.className === "Contacts") {
    return { adNode: ad, nativePlainText: null };
  }

  switch (bodyCase(ad)) {
    // Holds plain, gave plain - and holds HTML, gave HTML. Either way the
    // payload is the native form.
    case 0x11:
      return { adNode: ad, nativePlainText: null };
    case 0x22:
      return { adNode: ad, nativePlainText: plainSoFar };

    // Holds HTML, gave us a flattening. Ask for the real thing, once.
    case 0x21: {
      if (refetched || !serverID) return { adNode: ad, nativePlainText: null };
      const flattened = readPathFrom(ad, ["Body", "Data"]);
      let properties = null;
      let failure = null;
      try {
        properties = await fetchServerItem({
          account: ctx.account,
          asVersion: ctx.asVersion,
          collectionId: ctx.collectionId,
          serverID,
          bodyType: "2",
        });
      } catch (err) {
        failure = err;
      }
      // Could not ask, or the item is gone. The payload in hand is the
      // server's flattening of a note it holds as HTML, and storing it would
      // clear the ALTREP - so the next push would replace the server's rich
      // note with its own flattening, and one throttling window would strip
      // the formatting from every rich note in the batch. Drop the Body and
      // leave the note alone: an absent Body means "this response says
      // nothing about the note", which is exactly the truth here.
      if (!properties) {
        const body = childByTag(ad, "Body");
        if (body) ad.removeChild(body);
        ctx.provider.reportEventLog({
          level: "warning",
          accountId: ctx.accountId,
          folderId: ctx.folderId,
          message:
            `[${ctx.itemKind.changelogKind}-sync] could not fetch the HTML note for ` +
            `${serverID}; leaving the note untouched rather than replacing it ` +
            `with the server's plain-text rendering`,
          details: failure
            ? String(failure?.message ?? failure)
            : "no item returned",
        });
        return { adNode: ad, nativePlainText: null };
      }
      return resolveBody(ctx, properties, serverID, true, flattened);
    }

    // Holds plain, gave us HTML anyway. Honour what arrived: it is what the
    // server chose to hand us, and it is what we hand back unchanged.
    default:
      return { adNode: ad, nativePlainText: null };
  }
}

async function applyChangeFromAd(
  ctx,
  ad,
  existing,
  serverID = null,
  resolved = null,
  noteBacklog = null,
) {
  if (!resolved) {
    if (noteBacklog && needsHtmlRefetch(ctx, ad)) {
      recordNoteForUpgrade(noteBacklog, ad, serverID);
      resolved = { adNode: ad, nativePlainText: null };
    } else {
      resolved = await resolveBody(ctx, ad, serverID);
    }
  }
  ad = resolved.adNode;
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
  let blob = await ctx.itemKind.codec.applicationDataToBlob({
    adNode: ad,
    existingBlob: existing.blob,
    serverID: id,
    asVersion: ctx.asVersion,
    separator: ctx.separator,
    defaultTimezone: ctx.defaultTimezone,
    msTodoCompat: ctx.msTodoCompat,
    uid: existing.itemId,
    userEmail: accountUserAddress(ctx.account),
    eventLog: ctx.eventLog,
    nativePlainText: resolved.nativePlainText,
  });
  // The user's own answer is not the server's to overwrite. Responses are
  // sent after this pull, so without this the server's copy - which does
  // not know about an answer just given - would replace it, and the phase
  // below would then read that back and send it. The user accepts, the
  // calendar looks right, and the organizer never hears.
  if (ctx.itemKind.changelogKind === "event") {
    blob = preserveSelfPartstat({
      builtIcal: blob,
      priorIcal: existing.blob,
      userEmail: accountUserAddress(ctx.account),
    });
  }
  await ctx.queue.markServerWrite({
    parentId: ctx.targetID,
    itemId: existing.itemId,
    status: SERVER_TAG_STATUSES[1],
    kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.update(existing.itemId, blob);
  await verifyRoundTrip(ctx, existing.itemId, blob, "update");
  const masterServerId =
    ctx.itemKind.codec.readEasServerIdFromBlob(blob) ??
    ctx.itemKind.codec.readEasServerIdFromBlob(existing.blob);
  if (masterServerId) {
    ctx.indexMap.set(existing.itemId, masterServerId);
  }
  await noteExceptionsDroppedByServer(ctx, ad, existing, blob);
  if (blobHasRecurrence(blob) || blobHasRecurrence(existing.blob)) {
    logRecurrence(ctx, `pull update: itemId=${existing.itemId}`, {
      before: existing.blob,
      after: blob,
    });
  }
}

/** Queue a series whose exceptions the server has thrown away.
 *
 *  [MS-ASCAL] 2.2.2.22 and 2.2.2.42: at 16.0 and 16.1, changing a series'
 *  recurrence pattern or its start or end times deletes every exception on
 *  the item. Exchange does exactly that - one hour added to a master's end
 *  left it holding neither the moved occurrences nor the cancelled one -
 *  and it then restates the series with no `<Exceptions>` at all.
 *
 *  The merge above keeps ours, because an AD that does not mention
 *  exceptions is the same shape as a partial echo that simply has nothing
 *  to say about them. So the two copies disagree and nothing says so: the
 *  local calendar still shows every override, and they are gone the moment
 *  anything re-reads the folder.
 *
 *  Read from what the server sent rather than from a rule about when it
 *  does this: a restated `<Recurrence>` with no `<Exceptions>` beside it is
 *  a series the server holds bare. A partial echo carries neither and is
 *  passed over.
 *
 *  Queued rather than sent. The push that caused this has already run, and
 *  a command against an exception the server has just deleted is refused
 *  with Status 7 - measured, every one of three. The entry names an empty
 *  set as the baseline, which is what the server now holds, so the next
 *  push re-asserts the lot. */
export async function noteExceptionsDroppedByServer(ctx, ad, existing, blob) {
  if (
    ctx.asVersion !== "16.1" ||
    ctx.itemKind.changelogKind !== "event" ||
    !childByTag(ad, "Recurrence") ||
    childByTag(ad, "Exceptions") ||
    !blobHasInstanceOverrides(blob)
  ) {
    return;
  }
  await ctx.queue.record({
    parentId: ctx.targetID,
    itemId: existing.itemId,
    kind: ctx.itemKind.changelogKind,
    op: "updated",
    detail: { exceptions: { exdates: [], overrides: [] } },
  });
  ctx.provider.reportEventLog({
    level: "debug",
    accountId: ctx.accountId,
    folderId: ctx.folderId,
    message:
      `[event-sync] ${existing.itemId}: the server restated the series ` +
      `without the exceptions we hold; queued to re-assert them`,
  });
}

async function applyDelete(ctx, delNode, noteBacklog = null) {
  const serverID = readPathFrom(delNode, ["ServerId"]);
  if (!serverID) return;
  // A deletion outranks a pending note upgrade - fetching the rich note of
  // an item this same window removed would resurrect it as a change.
  noteBacklog?.delete(serverID);
  const existing = await findExistingByServerId(ctx, serverID);
  if (!existing) return;
  await ctx.queue.markServerWrite({
    parentId: ctx.targetID,
    itemId: existing.itemId,
    status: SERVER_TAG_STATUSES[2],
    kind: ctx.itemKind.changelogKind,
  });
  await ctx.store.delete(existing.itemId);
  ctx.indexMap.remove(existing.itemId);
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
    if (BUSY_STATUSES.has(top)) return { code: "BUSY", topStatus: top };
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
    if (BUSY_STATUSES.has(collStatus)) return { code: "BUSY", collStatus };
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
 *  answerable. */
function hasPendingUserDelete(ctx, serverId) {
  const uid = ctx.indexMap.uidFor(serverId);
  if (!uid) return false;
  // The queue as it stood when this sync began. Reading it live would need
  // this to be async for no gain: the indexMap check above has already
  // excluded every delete this sync managed to push, because acknowledging
  // one removes its mapping.
  return ctx.pendingAtStart.some(
    (e) => e?.itemId === uid && e?.status === "deleted_by_user",
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
 *  A miss means the item is new to us, and the pull acts on that by
 *  creating it - so the map answering for everything we already hold is
 *  what stands between a re-download and a duplicated folder. It is kept
 *  across a sync key reset for exactly that reason. */
async function findExistingByServerId(ctx, serverId) {
  let itemId = ctx.indexMap.uidFor(serverId);
  if (!itemId && ctx.indexMap.size === 0 && !ctx.indexRebuilt) {
    // A miss against an *empty* index is the one shape that says the index
    // was lost rather than that the item is new: anything the folder holds
    // would be in it. Rebuild from the blobs, once, and only then.
    //
    // It has to be here. An EAS task or contact carries no `UID` on the
    // wire, so the twin adopt below cannot recognise the item either, and a
    // server re-sending the folder would mint a second copy of every one.
    ctx.indexRebuilt = true;
    let items = [];
    try {
      items = await ctx.store.list();
    } catch (err) {
      // Leaves us where we were without the rebuild; failing the folder
      // over it would be worse.
      console.warn(
        "[eas-4-tbsync] could not read the store to rebuild the index:",
        err?.message ?? err,
      );
    }
    ctx.indexMap.fill(items, (blob) =>
      ctx.itemKind.codec.readEasServerIdFromBlob(blob),
    );
    if (ctx.indexMap.size) {
      ctx.eventLog(
        "info",
        `[${ctx.itemKind.changelogKind}-sync] the stored index was empty; ` +
          `rebuilt ${ctx.indexMap.size} mapping(s) from the items themselves`,
      );
    }
    itemId = ctx.indexMap.uidFor(serverId);
  }
  if (!itemId) return null;
  const item = await ctx.store.get(itemId);
  // Mapped to an item that is no longer there. The mapping stays: a delete
  // the user has made and we have not yet pushed looks exactly like this,
  // and `hasPendingUserDelete` reads it to tell that apart from drift.
  return item ? { itemId, blob: item.blob } : null;
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
 *  so they report without a summary or details.
 *
 *  `note` is for a rejection we act on rather than retry, so the line can
 *  say what became of the queued edit instead of leaving the user to wait
 *  for a repair that is not coming. */
function reportRejectedPushItem(
  ctx,
  operation,
  status,
  sentEntry,
  level,
  note,
) {
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
      (summary ? `: ${summary}` : "") +
      (note ? ` - ${note}` : ""),
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
