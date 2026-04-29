/**
 * Per-account read-only "Global Address List" address book, backed by
 * the EAS `Search` command via `addressBooks.provider.onSearchRequest`.
 *
 * Lifecycle:
 *   - On every host-port boot and on `onAccountEnabled`, we register a
 *     listener for each enabled EAS account whose OPTIONS-negotiated
 *     `allowedEasCommands` includes `Search`. Registering the listener
 *     creates the read-only directory (Thunderbird API contract) keyed
 *     by the deterministic id `eas-gal-<accountId>`.
 *   - On `onAccountDisabled` / `onAccountDeleted`, we deregister the
 *     listener so live searches stop hitting the server. The directory
 *     itself is removed via `addressBooks.delete` when possible; if the
 *     API rejects the call, the empty directory is left behind for the
 *     user to remove manually.
 *
 * Idempotency: registration is keyed by accountId in a module-scoped
 * map, so re-entry from boot + onAccountEnabled is safe.
 */

import { runGalSearch } from "./eas/gal-search.mjs";
import { easCommandAdvertised } from "./eas/allowed-commands.mjs";
import { isOAuthAccount, primeAuth } from "./eas/oauth.mjs";

const MIN_QUERY_LENGTH = 3;

const listeners = new Map(); // accountId → { callback, addressBookId }

function galAddressBookId(accountId) {
  return `eas-gal-${accountId}`;
}

function galAddressBookName(account) {
  // Suffix the account name so multiple accounts don't collide in the
  // directory tree. Localized via the same i18n that backs the rest of
  // the UI; falls back to English when the key is missing.
  const suffix =
    browser.i18n.getMessage("gal.addressBookSuffix") || "Global Address List";
  const base = account.accountName || account.accountId;
  return `${base} - ${suffix}`;
}

function searchSupported(account) {
  // The per-account toggle defaults to "enabled" - undefined / missing
  // counts as on, only an explicit `false` disables. New accounts get
  // GAL automatically; existing-pre-toggle accounts behave unchanged.
  if (account.custom?.galenabled === false) return false;
  return easCommandAdvertised(account, "Search");
}

/** Register the per-account onSearchRequest listener. No-op when the
 *  account has no Search capability or a listener is already in place. */
export async function enableGal({ provider, account }) {
  if (!account || !account.accountId) return;
  if (!searchSupported(account)) return;
  if (listeners.has(account.accountId)) return;

  const accountId = account.accountId;
  const addressBookId = galAddressBookId(accountId);
  const asVersion = account.custom?.asversion;

  const callback = async (_node, searchString) => {
    const query = String(searchString ?? "").trim();
    if (query.length < MIN_QUERY_LENGTH) {
      return { results: [], isCompleteResult: true };
    }
    try {
      // Reload the account each time so we pick up token / server-URL
      // changes that happened since enableGal ran. `getAccount` returns
      // a `{ account, folders }` wrapper - unwrap before use.
      const rv = await provider.getAccount(accountId);
      const fresh = rv?.account;
      if (!fresh || !searchSupported(fresh)) {
        return { results: [], isCompleteResult: true };
      }
      // Seed the OAuth auth cache for this account if needed. The
      // provider does this at the top of every on* hook that hits the
      // network; the GAL search callback runs outside those hooks, so
      // we have to prime explicitly before issuing the EAS request.
      if (isOAuthAccount(fresh.custom)) {
        primeAuth(accountId, {
          refreshToken: fresh.custom?.refreshToken,
          servertype: fresh.custom?.servertype,
        });
      }
      const results = await runGalSearch({
        account: fresh,
        asVersion: fresh.custom?.asversion ?? asVersion,
        query,
        companyName: fresh.accountName,
      });
      return { results, isCompleteResult: true };
    } catch (err) {
      provider.reportEventLog?.({
        level: "warning",
        accountId,
        message: `[gal] search failed: ${err?.message ?? String(err)}`,
      });
      return { results: [], isCompleteResult: true };
    }
  };

  try {
    messenger.addressBooks.provider.onSearchRequest.addListener(callback, {
      addressBookName: galAddressBookName(account),
      id: addressBookId,
      isSecure: true,
    });
  } catch (err) {
    provider.reportEventLog?.({
      level: "warning",
      accountId,
      message: `[gal] failed to register search listener: ${err?.message ?? String(err)}`,
    });
    return;
  }
  listeners.set(accountId, { callback, addressBookId });
}

/** Deregister the listener and (best-effort) drop the directory. */
export async function disableGal({ provider, accountId }) {
  const entry = listeners.get(accountId);
  if (!entry) return;
  listeners.delete(accountId);

  try {
    messenger.addressBooks.provider.onSearchRequest.removeListener(
      entry.callback,
    );
  } catch (err) {
    provider.reportEventLog?.({
      level: "debug",
      accountId,
      message: `[gal] removeListener failed (likely already gone): ${err?.message ?? String(err)}`,
    });
  }

  try {
    await messenger.addressBooks.delete(entry.addressBookId);
  } catch {
    // The directory may persist if the API does not allow deletion of
    // provider-created books; that's acceptable - searches simply stop
    // returning anything once the listener is gone.
  }
}

/** Iterate every enabled account and ensure GAL is registered. Called
 *  once when the provider's host port opens. */
export async function enableGalForAllAccounts(provider) {
  let accounts;
  try {
    accounts = await provider.listAccounts();
  } catch {
    return;
  }
  if (!Array.isArray(accounts)) return;
  for (const acc of accounts) {
    if (acc?.enabled === false) continue;
    await enableGal({ provider, account: acc });
  }
}
