/**
 * Provider-side completion of the host's legacy import.
 *
 * TbSync's importer lifts the host-owned fields out of the legacy
 * `<profile>/TbSync/*.json` files but copies every provider field into
 * `account.custom` verbatim, so an imported account still carries whatever
 * shape the legacy add-on wrote: `host` + `https` where the current code
 * wants `server`, credentials still in `nsILoginManager`,
 * `allowedEasCommands` as a comma-separated string. Converting the account's
 * settings is this module's job.
 *
 * What it deliberately does NOT convert is the account's data: the cards and
 * items in the bound address book and calendars, and any edit the legacy
 * add-on had queued for them. Those were written by code that identified
 * them differently, and an imported account does not sync at all until the
 * user reconnects it - which deletes the local copies and rebuilds them from
 * the server. Converting them first would be work done on copies about to be
 * replaced, and every attempt to do it had to guess at data it could not
 * read back.
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
 *   - Account `custom` lives in the host. Its trigger is the host's flag,
 *     which the host re-sets every time it re-imports.
 *   - This add-on's own global settings live in `storage.local`. Their
 *     trigger is `schemaVersion` in that same storage, so marker and data
 *     are wiped together and can never disagree.
 */

import {
  easTypeToFolderType,
  finalizeFolderListForPush,
  iconForServerType,
} from "./eas-provider.mjs";

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
  // Rung 4 ran when this version still adopted the host's imported change
  // queues. It no longer does - an imported account does not sync until it
  // is reconnected, which replaces its resources - but the number stays
  // spent: installations that reached 4 must not be walked over it again.
  4: { name: "adopt-host-changelogs" },
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
