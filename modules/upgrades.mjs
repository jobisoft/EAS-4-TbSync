/**
 * Provider-local one-shot upgrades.
 *
 * Runs work that has to happen exactly once after the user updates the
 * provider across a "split version" - typically a one-time data-shape
 * migration that the host's legacy migration deliberately couldn't do
 * because it's provider-specific.
 *
 * The trigger is `runtime.onInstalled` (with `reason === "update"` and a
 * `previousVersion` set), wired up in [background.mjs](../background.mjs).
 * Fresh installs never fire any upgrade. The list of pending upgrade IDs
 * persists in `browser.storage.local` under UPGRADE_QUEUE_KEY so a
 * partial run (host crash, network outage) is retried on the next
 * host-connect via the boot-time stale drain.
 *
 * While a drain is in flight, the host treats every account belonging
 * to this provider as "upgrading" - refuses every user-initiated RPC
 * and skips autosync ticks. The lock is acquired before the first
 * upgrade body runs and released in a `finally` so a crashing upgrade
 * still releases it.
 */

import {
  easTypeToFolderType,
  finalizeFolderListForPush,
} from "./eas-provider.mjs";

const UPGRADE_QUEUE_KEY = "eas.upgradeQueue";

/** Ordered list of split versions. An upgrade is *applicable* to an
 *  `(previousVersion, currentVersion)` pair iff
 *  `previousVersion < splitVersion <= currentVersion`. Strict on the
 *  prev side so a user already on the split doesn't re-run; inclusive
 *  on the cur side so installing exactly at the split triggers it. */
export const UPGRADES = [
  {
    splitVersion: "4.20",
    id: "eas.legacy-migration",
    run: async (provider) => {
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

      for (const migration of PREF_MIGRATIONS) {
        await liftPref(provider, migration);
      }

      const accounts = await provider.listAccounts();
      for (const acc of accounts) {
        try {
          await liftHostAndHttpsToServer(provider, acc);
          await liftCredentials(provider, acc);
          await normalizeAllowedEasCommands(provider, acc);
          await fixFolders(provider, acc);
        } catch (err) {
          provider.reportEventLog({
            level: "warning",
            accountId: acc.accountId,
            message: `[upgrade] failed to lift legacy state: ${err?.message ?? String(err)}`,
          });
        }
      }

      // Legacy EAS4 stored each contact's EAS ServerId as the TB card's
      // UID (`card.primaryKey === serverId`) and several extra fields in
      // the property bag via `setProperty()`. The new code expects the
      // ServerId in an `X-EAS-SERVERID` vCard property and the extras in
      // matching `X-EAS-*` properties. Without this migration, an
      // upgraded user would see duplicates after the first delta sync
      // and silent edit/delete failures on legacy cards. See Phase 3
      // audit row 3.11 for the full rationale.
      for (const acc of accounts) {
        try {
          await migrateContactsForAccount(provider, acc);
        } catch (err) {
          provider.reportEventLog({
            level: "warning",
            accountId: acc.accountId,
            message: `[upgrade] contact vCard migration failed: ${err?.message ?? String(err)}`,
          });
        }
      }      
    },
  },
];

/** Dotted-decimal version comparison. Sufficient for the version
 *  strings the legacy add-on shipped (e.g. `"4.17.2.ews.16.1"`) and
 *  the new add-on (`"5.0"`) - any non-numeric segment becomes NaN,
 *  which only matters if the *first differing* segment is non-numeric.
 *  In practice the legacy → new transition diverges at the first
 *  segment (4 → 5), so the comparator short-circuits before reaching
 *  any non-numeric tail. */
export function compareVersions(a, b) {
  const pa = String(a).split(".").map(Number);
  const pb = String(b).split(".").map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const x = pa[i] ?? 0,
      y = pb[i] ?? 0;
    if (x !== y) return x - y;
  }
  return 0;
}

let inFlight = null;

/** Drain `eas.upgradeQueue` against the UPGRADES table. Idempotent
 *  (each upgrade body is itself idempotent) and self-coalescing - a
 *  second caller while the first is mid-flight just awaits the same
 *  Promise.
 *
 *  The host upgrade lock is acquired before any upgrade body runs and
 *  released in `finally`, so:
 *    - User-initiated RPCs against this provider's accounts are refused
 *      while the drain is running.
 *    - Autosync ticks skip those accounts.
 *    - A throw inside an upgrade still releases the lock; the failed
 *      upgrade ID stays in the queue and is retried next boot. */
export function runUpgrades(provider) {
  if (inFlight) return inFlight;
  inFlight = (async () => {
    const rv = await browser.storage.local.get({ [UPGRADE_QUEUE_KEY]: [] });
    const queue = rv[UPGRADE_QUEUE_KEY];
    if (!queue.length) return;

    let lockAcquired = false;
    try {
      await provider.setProviderUpgradeLock(true);
      lockAcquired = true;
      provider.reportEventLog({
        level: "debug",
        message: `[upgrade] entering upgrade mode - sync and account/resource modifications are paused (${queue.length} upgrade(s) pending)`,
      });

      const remaining = [];
      for (const id of queue) {
        const upgrade = UPGRADES.find((u) => u.id === id);
        if (!upgrade) continue; // unknown id - silently drop
        try {
          provider.reportEventLog({
            level: "debug",
            message: `[upgrade] ${id} starting`,
          });
          await upgrade.run(provider);
          provider.reportEventLog({
            level: "debug",
            message: `[upgrade] ${id} done`,
          });
        } catch (err) {
          provider.reportEventLog({
            level: "error",
            message: `[upgrade] ${id} failed: ${err?.message ?? String(err)}`,
          });
          remaining.push(id);
        }
      }

      await browser.storage.local.set({ [UPGRADE_QUEUE_KEY]: remaining });
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
        provider.reportEventLog({
          level: "debug",
          message: `[upgrade] exiting upgrade mode - sync and account/resource modifications re-enabled`,
        });
      }
      inFlight = null;
    }
  })();
  return inFlight;
}

/** Compute the set of upgrades triggered by an update transition and
 *  merge their IDs into the persistent queue. No-op when nothing
 *  applies. Returns the new queue length. */
export async function enqueueUpgradesForUpdate(
  previousVersion,
  currentVersion,
) {
  const triggered = UPGRADES.filter(
    (u) =>
      compareVersions(previousVersion, u.splitVersion) < 0 &&
      compareVersions(u.splitVersion, currentVersion) <= 0,
  ).map((u) => u.id);
  if (!triggered.length) return 0;
  const rv = await browser.storage.local.get({ [UPGRADE_QUEUE_KEY]: [] });
  const next = Array.from(new Set([...rv[UPGRADE_QUEUE_KEY], ...triggered]));
  await browser.storage.local.set({ [UPGRADE_QUEUE_KEY]: next });
  return next.length;
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
async function migrateContactsForAccount(provider, acc) {
  const rv = await provider.getAccount(acc.accountId);
  const folders = rv?.folders ?? [];
  for (const folder of folders) {
    if (folder.targetType !== "contacts") continue;
    if (!folder.targetID) continue;
    try {
      await migrateContactsForFolder(provider, acc.accountId, folder);
    } catch (err) {
      provider.reportEventLog({
        level: "warning",
        accountId: acc.accountId,
        folderId: folder.folderId,
        message: `[upgrade] folder migration failed: ${err?.message ?? String(err)}`,
      });
    }
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
    provider.reportEventLog({
      level: "warning",
      accountId,
      folderId: folder.folderId,
      message: `[upgrade] LegacyAbProperties.readEasStamps failed: ${err?.message ?? String(err)}`,
    });
    return;
  }
  if (!Array.isArray(stamps) || stamps.length === 0) return;

  const contactMap = { ...(folder.custom?.contactMap ?? {}) };
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
        // Even if the vCard is migrated, the contactMap might be stale -
        // make sure the entry is present so the deleted_by_user path can
        // resolve the ServerId.
        if (!contactMap[contactId]) {
          contactMap[contactId] = contactId;
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
      await provider.changelogMarkServerWrite({
        accountId,
        folderId: folder.folderId,
        parentId: folder.targetID,
        itemId: contactId,
        status: "modified_by_server",
        kind: "contact",
      });

      await messenger.contacts.update(contactId, { vCard: newVCard });

      // Legacy convention: card.UID === EAS ServerId. The contactMap
      // entry is needed by `buildPushBatch`'s `deleted_by_user` path,
      // which only consults `idMap` (the codec's vCard-blob fallback
      // doesn't help for deletes - the local card is gone by then).
      contactMap[contactId] = contactId;
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
      patch: { custom: { contactMap } },
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
