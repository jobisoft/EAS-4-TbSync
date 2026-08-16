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
 *   - An account whose credentials the server has rejected keeps its
 *     listener but answers with no results, so searching resumes on its
 *     own once the account is authenticated again. Deregistering instead
 *     would need something to put the listener back, and an
 *     authentication failure leaves the account enabled.
 *
 * Idempotency: registration is keyed by accountId in a module-scoped
 * map, so re-entry from boot + onAccountEnabled is safe.
 */

import { ERR } from "../vendor/tbsync/provider.mjs";
import { easCommandAdvertised } from "./eas/allowed-commands.mjs";
import { isOAuthAccount, primeAuth } from "./eas/oauth.mjs";

const MIN_QUERY_LENGTH = 3;

/** How long an answer stays usable. Long enough to cover a compose
 *  session's typing and re-typing, short enough that a colleague added to
 *  the directory today is findable within the same sitting. */
const CACHE_TTL_MS = 120_000;

const listeners = new Map(); // accountId → { callback, addressBookId }

/* ── Answers we have already paid for ─────────────────────────────────
 *
 * Worth having because Thunderbird asks the same thing repeatedly: it
 * starts a fresh search on every keystroke and discards any answer that
 * arrives after the next one, so a query typed, deleted and retyped is
 * asked several times over - and it asks each GAL directory separately, so
 * a machine with four EAS accounts pays four round trips per keystroke.
 * Measured against a Z-Push GAL, one search took ~1.7s, longer than the
 * gap between keystrokes: without this, the answer to what the user is
 * typing reliably arrives too late to be shown.
 *
 * It never sees a query Thunderbird can answer itself: a complete result
 * for "abc" licenses it to narrow locally, so "abcd" never reaches us.
 *
 * What is stored is the promise rather than the answer, which makes one
 * structure do both jobs: a second asker for a question still in flight
 * waits on the request already running, and once it settles the same entry
 * is the cached answer. Awaiting a settled promise costs nothing.
 *
 * Entries carry the moment they expire, computed once when stored, and are
 * dropped by the lookup that finds them stale. Nothing sweeps.
 */

const answers = new Map(); // `${accountId}\n${query}` → { eol, promise }

function cacheKey(accountId, query) {
  // Case-folded: the server matches case-insensitively, so "Biel" and
  // "biel" are the same question and should not be asked twice.
  return `${accountId}\n${query.toLowerCase()}`;
}

/** The pending or settled answer for this query, or null when there is
 *  none or the one we had has expired. Expiry is applied here, which is
 *  what makes a sweeper unnecessary. */
function cachedAnswer(key) {
  const hit = answers.get(key);
  if (!hit) return null;
  if (hit.eol <= Date.now()) {
    answers.delete(key);
    return null;
  }
  return hit.promise;
}

/** Keep a running search under its query. A failure is not kept: handing
 *  the same rejection to every caller for the whole TTL would leave the
 *  GAL looking dead long after the network came back. */
function rememberAnswer(key, promise) {
  const entry = { eol: Date.now() + CACHE_TTL_MS, promise, answer: null };
  answers.set(key, entry);
  promise.then(
    (answer) => {
      // Kept on the entry as well as in the promise so `narrowFrom` can
      // read a finished answer without awaiting - it must not block on a
      // search still running for some other prefix.
      entry.answer = answer;
    },
    () => {
      if (answers.get(key)?.promise === promise) answers.delete(key);
    },
  );
  return promise;
}

/** Everything we hold about one result, lower-cased once, for matching. */
function haystack(contact) {
  return Object.values(contact).join(" ").toLowerCase();
}

/** Answer `query` from a wider answer we already have, or null.
 *
 *  Thunderbird does this itself for the compose autocomplete - a complete
 *  result for "bie" licenses it to narrow locally and "biel" never reaches
 *  us - but the address book search has no such step and asks us for every
 *  keystroke. On a GAL that takes ~1.7s per search that is the whole of the
 *  delay, so we do the same narrowing here.
 *
 *  Only from a *complete* answer: a truncated one is the first hundred of
 *  a larger set, and narrowing it would hide everyone the server held back.
 *  Longest prefix first, since it is the smallest set to filter and the
 *  most recently confirmed.
 *
 *  Verified against ekir's GAL before being written: the server's answer
 *  for "biel" is exactly this filter applied to its answer for "bie" (10
 *  of 76), and for "bielz" exactly this filter applied to "biel" (1 of 10).
 *  The assumption is that a longer query matches a subset of a shorter
 *  one, which holds for any substring or prefix matching. */
function narrowFrom(accountId, query) {
  const needle = query.toLowerCase();
  const prefix = `${accountId}\n`;
  let best = null;
  for (const [key, entry] of answers) {
    if (!key.startsWith(prefix) || !entry.answer) continue;
    if (entry.eol <= Date.now()) continue;
    const cachedQuery = key.slice(prefix.length);
    if (cachedQuery.length >= needle.length) continue;
    if (!needle.startsWith(cachedQuery)) continue;
    const { total, delivered } = entry.answer;
    if (total == null || total > delivered) continue; // truncated - unsafe
    if (!best || cachedQuery.length > best.q.length) {
      best = { q: cachedQuery, answer: entry.answer };
    }
  }
  if (!best) return null;
  const results = best.answer.results.filter((c) =>
    haystack(c).includes(needle),
  );
  // Complete by construction: it is every match in a set that was itself
  // complete.
  return { results, total: results.length, delivered: results.length };
}

/** Drop an account's answers, so a re-enabled account is never served from
 *  the period it was gone. */
function forgetAnswers(accountId) {
  const prefix = `${accountId}\n`;
  for (const key of answers.keys()) {
    if (key.startsWith(prefix)) answers.delete(key);
  }
}

let renameWatcherInstalled = false;

function galAddressBookId(accountId) {
  return `eas-gal-${accountId}`;
}

/** Inverse of `galAddressBookId`: returns the accountId encoded in a GAL
 *  directory id, or null for non-GAL ids. */
function accountIdFromGalAddressBookId(id) {
  if (typeof id !== "string") return null;
  const m = id.match(/^eas-gal-(.+)$/);
  return m ? m[1] : null;
}

function galAddressBookDefaultName(account) {
  // Suffix the account name so multiple accounts don't collide in the
  // directory tree. Localized via the same i18n that backs the rest of
  // the UI; falls back to English when the key is missing.
  const suffix =
    browser.i18n.getMessage("gal.addressBookSuffix") || "Global Address List";
  const base = account.accountName || account.accountId;
  return `${base} - ${suffix}`;
}

/** The display name to use when (re)creating the GAL directory. Prefers
 *  the user's locally-applied name (cached in `account.custom.galName`)
 *  so a rename survives the directory being torn down and recreated on
 *  next provider boot. */
function galAddressBookName(account) {
  const cached = account.custom?.galName;
  if (typeof cached === "string" && cached.trim()) return cached;
  return galAddressBookDefaultName(account);
}

function searchSupported(account) {
  // The per-account toggle defaults to "enabled" - undefined / missing
  // counts as on, only an explicit `false` disables. New accounts get
  // GAL automatically; existing-pre-toggle accounts behave unchanged.
  if (account.custom?.galenabled === false) return false;
  return easCommandAdvertised(account, "Search");
}

/** Install a single global watcher that mirrors local renames of any GAL
 *  directory back into `account.custom.galName`, so the rename survives
 *  the next teardown / recreation cycle. Idempotent - safe to call from
 *  the EAS provider's constructor. The listener filters by directory id
 *  prefix; non-GAL books are ignored. */
export function installRenameWatcher(provider) {
  if (renameWatcherInstalled) return;
  renameWatcherInstalled = true;

  messenger.addressBooks.onUpdated.addListener(async (node) => {
    const accountId = accountIdFromGalAddressBookId(node?.id);
    if (!accountId) return;
    if (typeof node.name !== "string" || !node.name) return;
    try {
      const rv = await provider.getAccount(accountId);
      const acc = rv?.account;
      if (!acc) return;
      if (acc.custom?.galName === node.name) return;
      await provider.updateAccount({
        accountId,
        patch: { custom: { galName: node.name } },
      });
    } catch (err) {
      console.debug("[eas] gal rename watcher update failed:", err);
    }
  });
}

/* ── Withholding an empty answer ──────────────────────────────────────
 *
 * Thunderbird drops our "ask me again" flag when the answer is empty.
 * `AbAutoCompleteSearch.onSearchFinished` records the directory in
 * `result.asyncDirectories` - the list the next search re-queries - but
 * that line sits inside `if (cards.length)`, so an empty result never gets
 * on it. The next keystroke then takes the reuse path, which *replaces*
 * the directory list with the previous result's instead of walking the
 * address book manager again, finds it empty, and returns before searching
 * anything. Our callback is never called again, however many characters
 * the user types.
 *
 * That is the whole of issue #344 for anyone whose server declines short
 * queries - Exchange wants four characters and answers three with a bare
 * <Result/> - and who also has a local contact matching the same prefix,
 * because a local match is what makes the previous result RESULT_SUCCESS
 * and so eligible for reuse. With no local match the result is
 * RESULT_NOMATCH, reuse is skipped, and the same typing works.
 *
 * So an empty answer is not returned at all: the promise is left pending,
 * `onSearchFinished` never runs for it, and the search stays
 * RESULT_SUCCESS_ONGOING. Reuse requires RESULT_SUCCESS exactly, so the
 * next keystroke searches afresh and reaches us. This is what happens by
 * accident when a debugger holds the search open, which is how the cause
 * was found.
 *
 * At most one is held per address book: the next request releases the one
 * before it. Releasing late costs nothing - Thunderbird has moved on and
 * discards the answer at its own `this._result != result` guard.
 *
 * Remove this once Thunderbird records the flag for empty results. It
 * exploits an implementation detail rather than a contract, and it leaves
 * one search per address book permanently open if the user stops typing
 * on a query the server declined.
 */
/** How long an empty autocomplete answer is held before being let go.
 *  Long enough to cover any amount of typing, and matched to the cache TTL
 *  so an answer is held for exactly as long as it would have been reused. */
const WITHHOLD_MS = CACHE_TTL_MS;

/** Hand back `answer`, unless it is the empty one the autocomplete cannot
 *  remember - that one is held for a while instead of returned.
 *
 *  Thunderbird records "ask me again" only for an answer that carried at
 *  least one card: in `AbAutoCompleteSearch.onSearchFinished` the
 *  `result.asyncDirectories.push(dir)` that the next search reads sits
 *  inside `if (cards.length)`. An empty answer is therefore forgotten, and
 *  the next keystroke takes the reuse path - which replaces the directory
 *  list with the previous result's rather than rebuilding it from the
 *  address book manager - finds it empty, and never asks us again, however
 *  many characters follow. That is the whole of #344 for anyone whose
 *  server declines short queries, and it only bites when a local card also
 *  matches, because that is what makes the previous result RESULT_SUCCESS
 *  and so eligible for reuse.
 *
 *  Not answering leaves the search RESULT_SUCCESS_ONGOING. Reuse requires
 *  RESULT_SUCCESS exactly, so the next keystroke searches afresh and
 *  reaches us. Holding on a timer rather than releasing on the next request
 *  keeps this free of shared state: the API tells us nothing about which
 *  window is asking - `onSearchRequest` passes the address book, the string
 *  and the query, and no caller identity - so anything keyed per account
 *  would let one compose window end another's hold and bring the bug back
 *  there. A timer belongs to its own search and to nothing else.
 *
 *  `abQuery` keeps this to the one caller that has the bug. Of the four
 *  things in Thunderbird that search a directory, only the autocomplete's
 *  async branch passes null:
 *
 *    AbAutoCompleteSearch:599  dir.search(null, str, listener)   <- this one
 *    AbAutoCompleteSearch:225  directory.search(query, str, ...) sync path
 *    AddrBookDataAdapter:38    dir.search(query, str, this)      address book
 *    ext-addressBook:1302      book.item.search(query, str, ...) contacts.query
 *
 *  The last two wait for us - `contacts.query` awaits one promise per book,
 *  resolved in `onSearchFinished` - so withholding from them would hang a
 *  caller with nothing to do with this bug.
 *
 *  Remove all of this once Thunderbird records the flag for empty results;
 *  `patches/` carries that fix. It exploits an implementation detail rather
 *  than a contract, and it costs one held search per query for the TTL. */
function answerOrWithhold(abQuery, answer) {
  if (abQuery != null) return answer;
  if (answer.results.length || answer.isCompleteResult) return answer;
  return new Promise((resolve) => {
    const timer = setTimeout(
      () => resolve({ results: [], isCompleteResult: false }),
      WITHHOLD_MS,
    );
    // No-op in the browser, where setTimeout returns a number. Under
    // node:test it stops a pending hold from keeping the process alive.
    timer?.unref?.();
  });
}

/** Register the per-account onSearchRequest listener. No-op when the
 *  account has no Search capability or a listener is already in place. */
export async function enableGal({ provider, account }) {
  if (!account || !account.accountId) return;
  if (!searchSupported(account)) return;
  if (listeners.has(account.accountId)) return;

  const accountId = account.accountId;
  const addressBookId = galAddressBookId(accountId);

  const callback = async (_node, searchString, abQuery = null) => {
    const query = String(searchString ?? "").trim();
    if (query.length < MIN_QUERY_LENGTH) {
      // `isCompleteResult: false`, deliberately: `true` tells Thunderbird
      // the empty answer is final for this prefix, and its autocomplete
      // then narrows the cached result locally instead of asking again on
      // the next keystroke - which is why the GAL appeared to need one
      // more character than MIN_QUERY_LENGTH (#344: gate 3, observed 4).
      // "Incomplete" keeps it querying, so the search genuinely fires the
      // moment the query is long enough.
      return answerOrWithhold(abQuery, {
        results: [],
        isCompleteResult: false,
      });
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
      // The server has rejected this account's credentials. Searching runs
      // outside the sync path - it fires on every keystroke in a compose
      // window - so nothing else stops it presenting the same rejected
      // credentials over and over, which is how a server decides to lock
      // an account out. Checked here rather than by deregistering the
      // listener so searches resume by themselves once the account is
      // authenticated again.
      if (fresh.error === ERR.AUTH) {
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
      const key = cacheKey(accountId, query);
      const { results, total, delivered } = await (cachedAnswer(key) ??
        narrowFrom(accountId, query) ??
        rememberAnswer(
          key,
          provider.runGalSearch({
            accountId,
            query,
            companyName: fresh.accountName,
          }),
        ));
      // `isCompleteResult: true` licenses Thunderbird to stop asking and
      // narrow this set locally as the user types on. Two things stop us
      // saying that, and both mean "ask again":
      //
      //   - the server stated no <Total>, so we know nothing about the
      //     match set - the spec says it MUST state one, and a server that
      //     does not has left us unable to tell "that is everyone" from
      //     "that is the first hundred of many";
      //   - it found more than it sent, having capped the answer at the
      //     <Range> we asked for.
      //
      // Getting this wrong towards `true` costs a user a colleague they
      // cannot find, with nothing to show why. Towards `false` it costs
      // some extra requests.
      const isCompleteResult = !(total == null || total > delivered);
      return answerOrWithhold(abQuery, { results, isCompleteResult });
    } catch (err) {
      provider.reportEventLog?.({
        level: "warning",
        accountId,
        message: `[gal] search failed: ${err?.message ?? String(err)}`,
      });
      // Incomplete, not complete: a search that failed has told us nothing
      // about the match set, and calling the empty answer final would let
      // Thunderbird narrow it locally for every further character - so one
      // timeout would silence the GAL for the rest of the typing, long
      // after the network recovered.
      return answerOrWithhold(abQuery, {
        results: [],
        isCompleteResult: false,
      });
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
  forgetAnswers(accountId);

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
  } catch (err) {
    // The directory may persist if the API does not allow deletion of
    // provider-created books; that's acceptable - searches simply stop
    // returning anything once the listener is gone.
    console.debug(
      `[eas] addressBooks.delete(${entry.addressBookId}) failed:`,
      err,
    );
  }
}

/** Iterate every enabled account and ensure GAL is registered. Called
 *  once when the provider's host port opens. */
export async function enableGalForAllAccounts(provider) {
  let accounts;
  try {
    accounts = await provider.listAccounts();
  } catch (err) {
    console.debug("[eas] enableGalForAllAccounts: listAccounts failed:", err);
    return;
  }
  if (!Array.isArray(accounts)) return;
  for (const acc of accounts) {
    if (acc?.enabled === false) continue;
    await enableGal({ provider, account: acc });
  }
}
