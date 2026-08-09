/**
 * Provider-side completion of the host's legacy import.
 *
 * TbSync's importer lifts the host-owned fields out of the legacy
 * `<profile>/TbSync/*.json` files but copies every provider field into
 * `account.custom` verbatim, so an imported account still carries whatever
 * shape the legacy add-on wrote: `host` + `https` where the current code
 * wants `server`, credentials still in `nsILoginManager`,
 * `allowedEasCommands` as a comma-separated string, EAS ServerIds stored as
 * Thunderbird item ids. Converting all of that is this module's job.
 *
 * The trigger is the host's `legacyMigrationPending` flag, not anything
 * about this add-on's own install history. The importer's only guard is
 * the absence of the host's account storage, and it never consumes the
 * legacy files, so anything that clears that storage - reinstalling TbSync
 * is enough - makes the next host boot re-import the legacy snapshot over
 * accounts that were already converted. No event reaches this side when
 * that happens, which is why the flag has to be polled rather than
 * reacted to.
 *
 * `runStartupMigrations` therefore runs on every port open (see
 * `onConnectedToHost` in eas-provider.mjs) and costs one `listAccounts`
 * when nothing is flagged. The host blocks flagged accounts until we clear
 * them, and an account whose conversion throws keeps its flag and is tried
 * again next boot - so every step below has to be idempotent.
 *
 * Two triggers, one per kind of data, on the principle that the record of
 * a conversion belongs with the data it converted:
 *
 *   - Account `custom` and the Thunderbird resources bound to it live in
 *     the host and in the address book / calendar. Their trigger is the
 *     host's flag, which the host re-sets every time it re-imports.
 *   - This add-on's own global settings live in `storage.local`. Their
 *     trigger is `schemaVersion` in that same storage, so marker and data
 *     are wiped together and can never disagree.
 */

import {
  easTypeToFolderType,
  finalizeFolderListForPush,
  iconForServerType,
} from "./eas-provider.mjs";
import * as calendarStore from "../vendor/tbsync/calendar.mjs";
import * as eventCodec from "./eas/calendar-codec.mjs";
import * as taskCodec from "./eas/task-codec.mjs";
import { localQueue } from "../vendor/tbsync/change-queue.mjs";
import {
  CHANGELOG_KINDS,
  isUserEntry,
  SERVER_TAG_STATUSES,
} from "../vendor/tbsync/changelog-core.mjs";

/** Coerce the legacy `Map<uid, serverId>` JSON shape into the new array
 *  of `{uid, serverId}` records. Returns a fresh array — caller is free
 *  to mutate. */
function buildIndexMap(value) {
  if (Array.isArray(value)) return value.slice();
  if (!value || typeof value !== "object") return [];
  return Object.entries(value).map(([uid, serverId]) => ({ uid, serverId }));
}

/** Legacy prefs to lift into `browser.storage.local`. Global rather than
 *  per-account, and driven by the schema ladder rather than by any
 *  account's state - see `MIGRATIONS` rung 2. */
const PREF_MIGRATIONS = [
  {
    keys: {
      "extensions.eas4tbsync.timeout": "timeout",
      "extensions.eas4tbsync.maxitems": "maxItems",
    },
    validate: (v) => typeof v === "number" && Number.isFinite(v) && v > 0,
    transform: (v) => v,
    logValue: (v) => ` (${v})`,
  },
  {
    keys: {
      "extensions.eas4tbsync.oauth.clientID": "oauth.clientID",
      "extensions.eas4tbsync.clientID.useragent": "tbsync.useragent",
      "extensions.eas4tbsync.clientID.type": "tbsync.type",
    },
    validate: (v) => typeof v === "string" && !!v.trim(),
    transform: (v) => v.trim(),
    logValue: () => "",
  },
  {
    keys: {
      "extensions.eas4tbsync.msTodoCompat": "msTodoCompat",
    },
    defaultValue: null,
    validate: (v) => typeof v === "boolean",
    transform: (v) => v,
    logValue: (v) => ` (${v})`,
  },
];

/* ── Provider-local storage schema ──────────────────────────────────── */

const SCHEMA_KEY = "schemaVersion";

/** Shape of this add-on's own `storage.local`. Independent of the add-on
 *  version - ship releases freely and bump this only when the stored shape
 *  actually changes. Also independent of any other provider's number: a
 *  `2` here and a `2` in google-4-tbsync describe different storages and
 *  must never be compared. */
const SCHEMA_VERSION = 4;

/** Steps that raise storage from the previous version to the keyed one,
 *  applied in ascending order. `name` appears in the event log so a
 *  support log shows the sequence rather than only its side effects. A
 *  rung with no `run` is legal and just bumps the number. */
const MIGRATIONS = {
  2: { name: "lift-legacy-prefs", run: liftLegacyPrefs },
  3: { name: "repair-unconverted-accounts", run: repairUnconvertedAccounts },
  4: { name: "adopt-host-changelogs", run: adoptHostChangelogs },
};

let inFlight = null;

/** Bring this installation up to date, in one pass under one upgrade lock:
 *  the storage schema ladder, then the accounts the host flagged.
 *
 *  Self-coalescing - a second caller while the first is mid-flight awaits
 *  the same Promise - and re-runnable, so a host that restarts and
 *  re-imports is picked up on the next port open. */
export function runStartupMigrations(provider) {
  if (inFlight) return inFlight;
  // Clear the latch when the run settles, however it settles - including
  // the common case where there was nothing to do. "Nothing to convert" is
  // only ever true of the moment it was asked: the host re-imports long
  // after we first connect, and a latch left holding a resolved Promise
  // would turn every later port open into a silent no-op.
  inFlight = runAll(provider).finally(() => {
    inFlight = null;
  });
  return inFlight;
}

async function runAll(provider) {
  // The lock goes up before anything is read, so the provider is never
  // serviceable with either phase outstanding. A run with nothing to do
  // costs one lock round-trip; port opens are rare.
  let lockAcquired = false;
  try {
    await provider.setProviderUpgradeLock(true);
    lockAcquired = true;

    await runStorageSchemaMigrations(provider);
    await convertFlaggedAccounts(provider);
  } finally {
    if (lockAcquired) {
      await provider
        .setProviderUpgradeLock(false)
        .catch((err) =>
          console.warn(
            "[eas-4-tbsync] failed to release upgrade lock:",
            err?.message ?? String(err),
          ),
        );
    }
  }
}

/** Walk the storage ladder from whatever version is recorded up to
 *  `SCHEMA_VERSION`.
 *
 *  Absent (or non-integer) means 1: storage exists but nothing has been
 *  migrated. Writing that before running anything gives a crash mid-rung a
 *  recorded state to resume from, and makes "have we ever run here?"
 *  answerable from storage rather than inferred from a side effect.
 *
 *  Each rung is stamped on success, so a failure at 3 keeps 2 banked; a
 *  rung that throws leaves the version alone and is retried on the next
 *  startup, which is why every `run` has to be idempotent. */
async function runStorageSchemaMigrations(provider) {
  const rv = await browser.storage.local.get({ [SCHEMA_KEY]: null });
  let version = rv[SCHEMA_KEY];
  if (!Number.isInteger(version) || version < 1) {
    version = 1;
    await browser.storage.local.set({ [SCHEMA_KEY]: version });
  }

  for (let next = version + 1; next <= SCHEMA_VERSION; next++) {
    const step = MIGRATIONS[next];
    const label = step ? ` (${step.name})` : "";
    try {
      if (step?.run) await step.run(provider);
      await browser.storage.local.set({ [SCHEMA_KEY]: next });
      provider.reportEventLog({
        level: "debug",
        message: `[upgrade] storage schema ${next - 1} -> ${next}${label}`,
      });
    } catch (err) {
      provider.reportEventLog({
        level: "warning",
        message: `[upgrade] storage schema ${next - 1} -> ${next}${label} failed, retrying on next start: ${err?.message ?? String(err)}`,
      });
      return;
    }
  }
}

/** Rung 2. Carry the settings a v4 user explicitly customised out of the
 *  legacy pref branch and into this add-on's storage.
 *
 *  Guarded by the ladder rather than by per-key checks: the Options page
 *  removes a key to mean "use the default", so an absent key cannot be
 *  read as "never set". Marker and settings share `storage.local`, so a
 *  reinstall wipes both together and re-adopting the v4 values has nothing
 *  to overwrite. `getUserPref` returns only prefs with a user value, never
 *  the defaults the legacy add-on registered, so an untouched v4 profile
 *  lifts nothing at all. */
async function liftLegacyPrefs(provider) {
  for (const migration of PREF_MIGRATIONS) {
    await liftPref(provider, migration);
  }
}

/** Rung 3. Convert accounts that the host left in legacy shape before
 *  `legacyMigrationPending` existed: the importer re-ran under a build that
 *  had no flag to set, so nothing has ever asked for their conversion and
 *  nothing ever would.
 *
 *  No detection heuristic - every step of the conversion is individually
 *  guarded, so running it over an already-converted account is a sequence
 *  of early returns. Accounts that *are* flagged belong to the flag path
 *  and are skipped here to avoid converting them twice in one run. */
async function repairUnconvertedAccounts(provider) {
  const accounts = await provider.listAccounts();
  const stale = accounts.filter((acc) => !acc.legacyMigrationPending);
  if (!stale.length) return;

  // Every account is attempted before the rung reports failure. Letting
  // the first throw escape would leave the accounts behind it untouched,
  // and a permanently failing one would then block the repair of all the
  // others for good, since the rung is never stamped and always restarts
  // from the same place.
  let failed = 0;
  for (const acc of stale) {
    try {
      await convertAccountData(provider, acc);
    } catch (err) {
      failed++;
      provider.reportEventLog({
        level: "warning",
        accountId: acc.accountId,
        message: `[upgrade] repair failed: ${err?.message ?? String(err)}`,
      });
    }
  }
  if (failed) {
    throw new Error(
      `${failed} of ${stale.length} account(s) could not be repaired`,
    );
  }
}

/** Rung 4. Take over any edits the host is still holding for us, into the
 *  local queue keyed by the folder's binding.
 *
 *  A host folder row can carry queued edits for any resource - unsynced work
 *  of the user's, which has to end up somewhere we will read it.
 *
 *  Import first, remove second. A crash in between leaves an entry in both
 *  places, and the next run re-imports it onto a queue that already folds
 *  duplicates by identity, so the repeat is a no-op rather than a second
 *  copy. The other order would lose entries outright. */
async function adoptHostChangelogs(provider) {
  for (const { accountId } of await provider.listAccounts()) {
    const { folders = [] } = (await provider.getAccount(accountId)) ?? {};
    for (const folder of folders) {
      const entries = Array.isArray(folder.changelog) ? folder.changelog : [];
      if (!entries.length) continue;
      if (!folder.sessionId) {
        // No session, but edits to adopt: the host has not (yet) minted
        // session ids for its rows. Continuing here and letting the ladder
        // stamp the schema would strand these rows forever - the ladder
        // never looks back. Throw instead: the rung stays unstamped and
        // the whole adoption retries on the next start, when the host has
        // caught up. Observed for real on 10 Aug 2026 during migration
        // testing - "unreachable in practice" was wrong.
        throw new Error(
          `cannot adopt ${entries.length} queued edit(s) of folder ` +
            `${folder.folderId}: the folder has no session id yet`,
        );
      }

      const queue = localQueue({
        accountId,
        folderId: folder.folderId,
        sessionId: folder.sessionId,
        observed: folder.targetType === "contacts",
      });
      // A single-kind folder names its rows' kind itself. 5.0.13 was seen
      // stamping a task DELETION as kind "event" in the wild (the gold
      // migration baseline, 10 Aug 2026); adopted verbatim, such a row
      // survives adoption only to be dropped at push as "not this
      // folder's kind" - and the pull then resurrects the item the user
      // deleted. Contacts folders hold several kinds and cannot be
      // healed this way; calendars and tasks hold exactly one.
      const folderKind = { calendars: "event", tasks: "task" }[
        folder.targetType
      ];
      let adopted = 0;
      let refused = 0;
      let healed = 0;
      for (const e of entries) {
        if (!isUserEntry(e?.status)) continue;
        if (folderKind && e.kind !== folderKind) {
          e.kind = folderKind;
          healed++;
        }
        // `record` throws on a kind outside CHANGELOG_KINDS. The inbox is
        // whatever the v4 importer wrote, and one malformed row must not
        // abort the whole migration - skip it, count it, and let the
        // summary line below name the loss.
        if (!CHANGELOG_KINDS.includes(e.kind)) {
          refused++;
          continue;
        }
        await queue.record({
          parentId: e.parentId,
          itemId: e.itemId,
          kind: e.kind,
          op: OP_FOR_STATUS[e.status],
          detail: e.detail,
        });
        adopted++;
      }
      await provider.updateFolder({
        accountId,
        folderId: folder.folderId,
        patch: { changelog: [] },
      });
      provider.reportEventLog({
        level: refused ? "warning" : "info",
        accountId,
        folderId: folder.folderId,
        message:
          `[upgrade] adopted ${adopted} queued edit(s) from the host` +
          (healed ? `; healed ${healed} kind(s) to the folder's own` : "") +
          (refused
            ? `; refused ${refused} with an unknown changelog kind`
            : ""),
      });
    }
  }
}

/** The op that produced each user status - `record` speaks ops, a stored
 *  row states the status it reached. Replaying the op onto an empty queue
 *  reproduces the row, which is what makes the import idempotent. */
const OP_FOR_STATUS = {
  added_by_user: "created",
  modified_by_user: "updated",
  deleted_by_user: "deleted",
};

/* ── Host-flag driven account conversion ────────────────────────────── */

/** Convert every account the host flagged, clearing each flag as it
 *  succeeds. One account failing must not affect the others, so failures
 *  are contained per account rather than aborting the phase. */
async function convertFlaggedAccounts(provider) {
  const pending = (await provider.listAccounts()).filter(
    (acc) => acc.legacyMigrationPending,
  );
  if (!pending.length) return;

  provider.reportEventLog({
    level: "debug",
    message: `[upgrade] converting ${pending.length} legacy-imported account(s)`,
  });
  for (const acc of pending) {
    await convertAccount(provider, acc);
  }
}

/** Convert one flagged account and tell the host it is finished. A throw
 *  anywhere leaves the flag set, so the account stays blocked and is
 *  retried next boot rather than syncing against half-converted data. */
async function convertAccount(provider, acc) {
  try {
    await convertAccountData(provider, acc);
    // Last, so the flag only clears once every step above has landed.
    await provider.legacyMigrationDone({ accountId: acc.accountId });
  } catch (err) {
    provider.reportEventLog({
      level: "warning",
      accountId: acc.accountId,
      message: `[upgrade] legacy conversion failed - account stays blocked and is retried on the next boot: ${err?.message ?? String(err)}`,
    });
    return;
  }
  provider.reportEventLog({
    level: "info",
    accountId: acc.accountId,
    message: `[upgrade] legacy conversion complete`,
  });
}

/** The conversion itself, with no flag handling, so it can serve both the
 *  flag path and the rung-3 repair. Throws on the first step that fails -
 *  the caller decides what that means. */
async function convertAccountData(provider, acc) {
  await liftHostAndHttpsToServer(provider, acc);
  await liftCredentials(provider, acc);
  await normalizeAllowedEasCommands(provider, acc);
  await fixFolders(provider, acc);
  await liftAccountIcon(provider, acc);

  // Legacy EAS4 stored each contact's EAS ServerId as the TB card's
  // UID (`card.primaryKey === serverId`) and several extra fields in
  // the property bag via `setProperty()`. The new code expects the
  // ServerId in an `X-EAS-SERVERID` vCard property and the extras in
  // matching `X-EAS-*` properties. Without this migration, an
  // upgraded user would see duplicates after the first delta sync
  // and silent edit/delete failures on legacy cards. See Phase 3
  // audit row 3.11 for the full rationale.
  await migrateContactsForAccount(provider, acc);

  // Legacy EAS4 stored each event/task's EAS ServerId as the
  // calendar item's id (`item.primaryKey === serverId`, see legacy
  // sync.js:1042). The new code expects the ServerId in an
  // `X-EAS-SERVERID` iCal property and a matching
  // `folder.custom.indexMap` entry. Without this migration, an
  // upgraded user would see duplicates after the first delta sync
  // and silent edit/delete failures on legacy events/tasks. Same
  // shape as the contact migration above; simpler because Lightning
  // already stores X-EAS-* properties in the iCal blob (no XPCOM
  // experiment needed).
  await migrateCalendarItemsForAccount(provider, acc);
}

// ── Upgrade helpers for legacy migrations─────────────────────────────────────

async function liftPref(provider, { keys, validate, transform, logValue }) {
  for (const [legacyKey, storageKey] of Object.entries(keys)) {
    const value = await browser.LegacyPrefs.getUserPref(legacyKey);
    if (!validate(value)) continue;

    const newValue = transform(value);
    await browser.storage.local.set({ [storageKey]: newValue });

    provider.reportEventLog({
      level: "debug",
      message: `[upgrade] lifted legacy '${legacyKey}' pref${logValue(newValue)} into storage.local['${storageKey}']`,
    });
  }
}

/** Normalize legacy `allowedEasCommands` from a comma-separated string
 *  into the canonical deduped array. After this runs, the rest of the
 *  provider can assume an array on `account.custom.allowedEasCommands`
 *  without sniffing for the legacy string form. Some EAS frontends emit
 *  MS-ASProtocolCommands twice, so the legacy raw header value can carry
 *  the same command twice - `Set` collapses those at upgrade time. */
async function normalizeAllowedEasCommands(provider, acc) {
  const cmds = acc.custom?.allowedEasCommands;
  if (Array.isArray(cmds)) return;
  if (typeof cmds !== "string" || !cmds.length) return;
  const arr = [
    ...new Set(
      cmds
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean),
    ),
  ];
  await provider.updateAccount({
    accountId: acc.accountId,
    patch: { custom: { allowedEasCommands: arr } },
  });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] normalized legacy allowedEasCommands string into a deduped array of ${arr.length} command(s)`,
  });
}

/** Re-derive host-shape per-folder fields the legacy migration couldn't:
 *  - `targetType` from EAS-specific `custom.type` (legacy stored task
 *    folders with the same `"calendar"` string Lightning uses for both
 *    calendars and task lists, so the host migrator's static
 *    `calendar → calendars` mapping mis-labels tasks as calendars).
 *  - Trash visibility flags via `finalizeFolderListForPush`. */
async function fixFolders(provider, acc) {
  const rv = await provider.getAccount(acc.accountId);
  const folders = rv?.folders ?? [];
  if (!folders.length) return;
  const retyped = folders.map((f) => ({
    ...f,
    targetType: easTypeToFolderType(f.custom?.type) ?? f.targetType,
  }));
  const patched = await finalizeFolderListForPush(retyped);
  await provider.pushFolderList({ accountId: acc.accountId, folders: patched });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] re-derived targetType + trash visibility for ${patched.length} folder(s)`,
  });
}

async function liftAccountIcon(provider, acc) {
  if (acc.icon) return;
  const servertype = acc.custom?.servertype;
  const icon = iconForServerType(servertype);
  if (!icon) return;
  await provider.updateAccount({
    accountId: acc.accountId,
    patch: { icon },
  });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] set per-account icon for servertype="${servertype}"`,
  });
}

async function liftHostAndHttpsToServer(provider, acc) {
  if (acc.custom?.server) return;
  const host = acc.custom?.host;
  if (!host) return;
  const protocol = acc.custom?.https ? "https://" : "http://";
  let url = protocol + host;
  while (url.endsWith("/")) url = url.slice(0, -1);
  if (!url.endsWith("Microsoft-Server-ActiveSync"))
    url += "/Microsoft-Server-ActiveSync";
  await provider.updateAccount({
    accountId: acc.accountId,
    patch: { custom: { server: url, host: null, https: null } },
  });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] lifted legacy host+https to server="${url}"`,
  });
}

async function liftCredentials(provider, acc) {
  /** nsILoginManager realm legacy used for EAS credentials
   *  The origin is namespaced per-account as "TbSync#<accountID>" rather
   *  than the actual server hostname - legacy decoupled the credential
   *  from the host so Autodiscover-driven host changes don't orphan it.
   *  The legacy accountID survives the host's profile migration unchanged,
   *  so we can reach the entry by reusing `account.accountId` here.
   */
  const LEGACY_LOGIN_REALM = "TbSync/EAS";

  const c = acc.custom ?? {};
  const isOAuthLegacy = c.servertype === "office365";

  if (isOAuthLegacy && c.refreshToken) return;
  if (!isOAuthLegacy && c.password) return;

  const user = c.user;
  if (!user) {
    provider.reportEventLog({
      level: "warning",
      accountId: acc.accountId,
      message: `[upgrade] cannot lift credentials: missing legacy user`,
    });
    return;
  }

  const origin = `TbSync#${acc.accountId}`;
  const stored = await browser.LegacyLoginManager.getLoginInfo({
    origin,
    httpRealm: LEGACY_LOGIN_REALM,
    username: user,
  });
  if (stored == null) {
    provider.reportEventLog({
      level: "warning",
      accountId: acc.accountId,
      message: `[upgrade] no legacy nsILoginManager entry for (${origin}, ${LEGACY_LOGIN_REALM}, ${user})`,
    });
    return;
  }

  if (isOAuthLegacy) {
    let refreshToken = "";
    try {
      refreshToken = JSON.parse(stored)?.refresh ?? "";
    } catch {
      /* malformed blob; refreshToken stays empty */
    }
    if (!refreshToken) {
      provider.reportEventLog({
        level: "warning",
        accountId: acc.accountId,
        message: `[upgrade] legacy OAuth token blob has no 'refresh' field`,
      });
      return;
    }
    await provider.updateAccount({
      accountId: acc.accountId,
      patch: {
        custom: {
          refreshToken,
          authenticatedUserEmail: c.authenticatedUserEmail ?? null,
        },
      },
    });
    provider.reportEventLog({
      level: "debug",
      accountId: acc.accountId,
      message: `[upgrade] lifted legacy OAuth refresh token from nsILoginManager`,
    });
    return;
  }

  await provider.updateAccount({
    accountId: acc.accountId,
    patch: {
      custom: {
        password: stored,
      },
    },
  });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] lifted legacy basic-auth password from nsILoginManager`,
  });
}

/* ── Contact vCard migration (5.0.3) ─────────────────────────────────── */

/** Walk every selected contacts folder on the account and re-shape each
 *  legacy card into the new vCard layout. Idempotent: cards that already
 *  carry an `X-EAS-SERVERID` property are skipped. */
/** Every contacts folder is attempted, but one failing fails the account.
 *  The caller uses that to decide whether the legacy flag may clear or the
 *  schema rung may stamp - and neither may happen over a folder whose
 *  cards still carry their ServerId as the card UID, because the sync path
 *  has no fallback to the legacy property bag and would duplicate them. */
async function migrateContactsForAccount(provider, acc) {
  const rv = await provider.getAccount(acc.accountId);
  const folders = rv?.folders ?? [];
  let failed = 0;
  for (const folder of folders) {
    if (folder.targetType !== "contacts") continue;
    if (!folder.targetID) continue;
    try {
      await migrateContactsForFolder(provider, acc.accountId, folder);
    } catch (err) {
      failed++;
      provider.reportEventLog({
        level: "warning",
        accountId: acc.accountId,
        folderId: folder.folderId,
        message: `[upgrade] folder migration failed: ${err?.message ?? String(err)}`,
      });
    }
  }
  if (failed) {
    throw new Error(`${failed} contact folder(s) could not be migrated`);
  }
}

async function migrateContactsForFolder(provider, accountId, folder) {
  // Read every non-list card via the LegacyAbProperties experiment, which
  // also surfaces any property-bag stamps (Spouse, Yomi*, Children, …)
  // that legacy wrote via setProperty() and that don't appear in the
  // WebExtension-visible vCard.
  let stamps;
  try {
    stamps = await browser.LegacyAbProperties.readEasStamps(folder.targetID);
  } catch (err) {
    // Not knowing whether this folder holds legacy cards is not the same
    // as knowing it doesn't. Rethrow so the account stays blocked: the
    // sync path never reads the property bag, so proceeding would sync
    // cards whose identity we failed to look at and duplicate every one.
    // This is also how the eventual removal of the Experiment surfaces -
    // as a blocked account with a reason, not silent duplication.
    throw new Error(
      `LegacyAbProperties.readEasStamps failed: ${err?.message ?? String(err)}`,
    );
  }
  if (!Array.isArray(stamps) || stamps.length === 0) return;

  // indexMap is an array of {uid, serverId}. Coerce older installs'
  // object shape on read, then operate on the array.
  const indexMap = buildIndexMap(folder.custom?.indexMap);
  const hasContactMapping = (uid) => indexMap.some((e) => e.uid === uid);
  const upsertContactMapping = (uid, serverId) => {
    const ex = indexMap.find((e) => e.uid === uid);
    if (ex) {
      ex.serverId = serverId;
    } else {
      indexMap.push({ uid, serverId });
    }
  };
  let migratedCount = 0;
  let alreadyMigratedCount = 0;

  for (const { contactId, stamps: legacyStamps } of stamps) {
    try {
      const card = await messenger.contacts.get(contactId);
      const oldVCard =
        typeof card?.vCard === "string" && card.vCard.trim() ? card.vCard : "";
      if (!oldVCard) continue;

      // Idempotency guard: if X-EAS-SERVERID is already present we've
      // already migrated this card on a previous run.
      if (hasVCardProperty(oldVCard, "X-EAS-SERVERID")) {
        // Even if the vCard is migrated, the indexMap might be stale -
        // make sure the entry is present so the deleted_by_user path can
        // resolve the ServerId.
        if (!hasContactMapping(contactId)) {
          upsertContactMapping(contactId, contactId);
          migratedCount++;
        } else {
          alreadyMigratedCount++;
        }
        continue;
      }

      const newVCard = buildMigratedVCard(oldVCard, contactId, legacyStamps);
      if (newVCard === oldVCard) continue;

      // Pre-tag the changelog so the address-book observer treats the
      // upcoming `messenger.contacts.update` as self-inflicted and does
      // NOT produce a `modified_by_user` entry. NB: this also replaces
      // any existing entry for the card - any unsynced pre-upgrade user
      // edit's *push intent* is dropped (the data itself is preserved
      // because the migration read-modify-writes the vCard).
      await localQueue({
        accountId,
        folderId: folder.folderId,
        sessionId: folder.sessionId,
        observed: true,
      }).markServerWrite({
        parentId: folder.targetID,
        itemId: contactId,
        status: SERVER_TAG_STATUSES[1],
        kind: "contact",
      });

      await messenger.contacts.update(contactId, { vCard: newVCard });

      // Legacy convention: card.UID === EAS ServerId. The indexMap
      // entry is needed by `buildPushBatch`'s `deleted_by_user` path,
      // which only consults the idIndex (the codec's vCard-blob
      // fallback doesn't help for deletes - the local card is gone
      // by then).
      upsertContactMapping(contactId, contactId);
      migratedCount++;
    } catch (err) {
      provider.reportEventLog({
        level: "warning",
        accountId,
        folderId: folder.folderId,
        message: `[upgrade] contact ${contactId}: ${err?.message ?? String(err)}`,
      });
    }
  }

  if (migratedCount > 0) {
    await provider.updateFolder({
      accountId,
      folderId: folder.folderId,
      patch: { custom: { indexMap } },
    });
    provider.reportEventLog({
      level: "debug",
      accountId,
      folderId: folder.folderId,
      message: `[upgrade] migrated ${migratedCount} contact card(s) to new vCard shape (${alreadyMigratedCount} already migrated)`,
    });
  } else if (alreadyMigratedCount > 0) {
    provider.reportEventLog({
      level: "debug",
      accountId,
      folderId: folder.folderId,
      message: `[upgrade] all ${alreadyMigratedCount} contact card(s) already migrated`,
    });
  }
}

/** Map a legacy property-bag key to its `X-EAS-*` vCard counterpart.
 *  Most keys carry the `EAS-` prefix in legacy storage and just need
 *  `X-` prepended; `Children` is the odd one - legacy stored it without
 *  any prefix at all. */
function legacyKeyToVCardKey(legacyKey) {
  if (legacyKey === "Children") return "X-EAS-CHILDREN";
  if (legacyKey.startsWith("EAS-")) {
    return "X-EAS-" + legacyKey.slice("EAS-".length).toUpperCase();
  }
  // Fallback: shouldn't happen for fields in LEGACY_PROPERTY_BAG_FIELDS,
  // but be defensive.
  return "X-EAS-" + legacyKey.toUpperCase();
}

/** Insert `X-EAS-SERVERID` and any `X-EAS-*` properties derived from the
 *  legacy property bag, just before `END:VCARD`. Line-based (rather than
 *  parsing through ICAL.js) so the rest of the vCard's existing
 *  formatting / line-folding is preserved untouched. */
function buildMigratedVCard(vCardString, serverId, legacyStamps) {
  const lines = vCardString.split(/\r?\n/);
  const endIdx = lines.findIndex((l) => /^END:VCARD\s*$/i.test(l));
  if (endIdx === -1) return vCardString; // malformed; skip

  const inserts = [];
  if (!hasVCardProperty(vCardString, "X-EAS-SERVERID")) {
    inserts.push(`X-EAS-SERVERID:${escapeVCardValue(serverId)}`);
  }
  for (const [legacyKey, value] of Object.entries(legacyStamps ?? {})) {
    const vCardKey = legacyKeyToVCardKey(legacyKey);
    // Skip if the new key is already on the card (partial migration).
    if (hasVCardProperty(vCardString, vCardKey)) continue;
    inserts.push(`${vCardKey}:${escapeVCardValue(value)}`);
  }
  if (inserts.length === 0) return vCardString;

  // RFC 6350 §3.2 line termination is CRLF; preserve.
  return [...lines.slice(0, endIdx), ...inserts, ...lines.slice(endIdx)].join(
    "\r\n",
  );
}

function hasVCardProperty(vCardString, key) {
  // Match `KEY:` or `KEY;params:` at line start, case-insensitive.
  const re = new RegExp(
    `^${key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}[:;]`,
    "im",
  );
  return re.test(vCardString);
}

function escapeVCardValue(s) {
  // RFC 6350 §3.4: escape backslash, comma, newline. Single-value text
  // properties don't need to escape semicolons (those are compound-value
  // delimiters).
  return String(s)
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "")
    .replace(/,/g, "\\,");
}

/* ── Calendar / Task item migration ──────────────────────────────────── */

/** Walk every events/tasks folder on the account and re-shape each
 *  legacy item by stamping `X-EAS-SERVERID` onto its iCal and adding
 *  a matching `folder.custom.indexMap` entry. Idempotent: items that
 *  already carry `X-EAS-SERVERID` are skipped. */
/** Attempts every calendar/task folder, then fails the account if any of
 *  them failed - see `migrateContactsForAccount` for why partial success
 *  must not be reported as success. */
async function migrateCalendarItemsForAccount(provider, acc) {
  const rv = await provider.getAccount(acc.accountId);
  const folders = rv?.folders ?? [];
  let failed = 0;
  for (const folder of folders) {
    const itemKind = itemKindForFolder(folder);
    if (!itemKind || !folder.targetID) continue;
    try {
      await migrateCalendarItemsForFolder(
        provider,
        acc.accountId,
        folder,
        itemKind,
      );
    } catch (err) {
      failed++;
      provider.reportEventLog({
        level: "warning",
        accountId: acc.accountId,
        folderId: folder.folderId,
        message: `[upgrade] calendar/task folder migration failed: ${err?.message ?? String(err)}`,
      });
    }
  }
  if (failed) {
    throw new Error(`${failed} calendar/task folder(s) could not be migrated`);
  }
}

/** Map a folder's targetType to the codec / store-type / changelog-kind
 *  triple used by the per-item migration. Returns null for non-calendar
 *  folders. */
function itemKindForFolder(folder) {
  if (folder.targetType === "calendars") {
    return { storeType: "event", codec: eventCodec, changelogKind: "event" };
  }
  if (folder.targetType === "tasks") {
    return { storeType: "task", codec: taskCodec, changelogKind: "task" };
  }
  return null;
}

async function migrateCalendarItemsForFolder(
  provider,
  accountId,
  folder,
  itemKind,
) {
  let items;
  try {
    items = await calendarStore.listItems(folder.targetID, itemKind.storeType);
  } catch (err) {
    // Same reasoning as the contact stamps above: an unreadable folder is
    // not an empty one, and letting it pass would clear the flag over
    // items whose ServerId is still stored as the item id.
    throw new Error(
      `calendar items.list failed: ${err?.message ?? String(err)}`,
    );
  }
  if (!Array.isArray(items) || items.length === 0) return;

  // indexMap is an array of {uid, serverId}. Coerce older installs'
  // object shape on read.
  const indexMap = buildIndexMap(folder.custom?.indexMap);
  const hasItemMapping = (uid) => indexMap.some((e) => e.uid === uid);
  const upsertItemMapping = (uid, serverId) => {
    const ex = indexMap.find((e) => e.uid === uid);
    if (ex) {
      ex.serverId = serverId;
    } else {
      indexMap.push({ uid, serverId });
    }
  };
  let migratedCount = 0;
  let alreadyMigratedCount = 0;

  for (const node of items) {
    const id = node?.id;
    const ical =
      typeof node?.item === "string" && node.item.trim() ? node.item : "";
    if (!id || !ical) continue;
    try {
      // Idempotency guard: skip if X-EAS-SERVERID is already on the
      // master VEVENT/VTODO.
      if (itemKind.codec.readEasServerIdFromIcal(ical)) {
        if (!hasItemMapping(id)) {
          upsertItemMapping(id, id);
          migratedCount++;
        } else {
          alreadyMigratedCount++;
        }
        continue;
      }

      const stamped = itemKind.codec.stampEasServerId(ical, id);
      if (stamped === ical) continue; // codec couldn't parse the iCal; skip.

      // Through the cache, so the write fires no item hook and cannot be
      // mistaken for the user editing every item in the folder. A pre-tag
      // would not do here: our own calendars are not observed, so nothing
      // would ever consume one - the suppression has to be structural,
      // exactly as it is for a sync write.
      await calendarStore.updateItem(
        calendarStore.cacheId(folder.targetID),
        id,
        {
          ical: stamped,
        },
      );

      // Legacy convention: item.id === EAS ServerId. Populate the
      // indexMap so `buildPushBatch`'s deleted_by_user path (which only
      // consults the idIndex, not the iCal blob - the local item is
      // gone by then) can resolve the ServerId.
      upsertItemMapping(id, id);
      migratedCount++;
    } catch (err) {
      provider.reportEventLog({
        level: "warning",
        accountId,
        folderId: folder.folderId,
        message: `[upgrade] ${itemKind.changelogKind} item ${id}: ${err?.message ?? String(err)}`,
      });
    }
  }

  if (migratedCount > 0) {
    await provider.updateFolder({
      accountId,
      folderId: folder.folderId,
      patch: { custom: { indexMap } },
    });
    provider.reportEventLog({
      level: "debug",
      accountId,
      folderId: folder.folderId,
      message: `[upgrade] migrated ${migratedCount} ${itemKind.changelogKind} item(s) to new iCal shape (${alreadyMigratedCount} already migrated)`,
    });
  } else if (alreadyMigratedCount > 0) {
    provider.reportEventLog({
      level: "debug",
      accountId,
      folderId: folder.folderId,
      message: `[upgrade] all ${alreadyMigratedCount} ${itemKind.changelogKind} item(s) already migrated`,
    });
  }
}
