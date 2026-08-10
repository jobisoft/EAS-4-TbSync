/**
 * Attendee availability, backed by the EAS `ResolveRecipients` command
 * via `calendar.provider.onFreeBusy`.
 *
 * Lifecycle, and where it differs from the GAL next door: registering
 * this listener adds ONE provider to Thunderbird's global free/busy
 * service, not a per-account directory. So there is a single listener
 * for the whole add-on, held while any account can answer and dropped
 * when none can - and the routing happens inside it.
 *
 * Routing is needed because the call carries only an address: no event,
 * no calendar, no account. Every account whose mailbox matches is asked
 * and their intervals are contributed together - not reconciled against
 * each other, because Thunderbird already combines what several
 * providers say about one person and two accounts that can both see a
 * colleague will agree anyway. An address that matches nothing is
 * answered with silence rather than handed to a server that has no
 * relationship with it.
 */

import { ERR } from "../vendor/tbsync/provider.mjs";
import { easCommandAdvertised } from "./eas/allowed-commands.mjs";
import { isOAuthAccount, primeAuth } from "./eas/oauth.mjs";
import {
  accountCanAnswerFor,
  alignWindow,
  intervalsFromMergedFreeBusy,
} from "./eas/free-busy.mjs";
import { runResolveRecipients } from "./eas/resolve-recipients.mjs";

/** How long an answer stays good. The dialog re-asks as the user edits
 *  attendees and times, and these servers throttle; a minute is long
 *  enough to absorb that editing and short enough that a colleague who
 *  just accepted something shows up while the dialog is still open. */
const CACHE_TTL_MS = 60 * 1000;

let listener = null;
/** key → { at, promise } - `promise` so identical lookups in flight
 *  share one request instead of queuing a second. */
const cache = new Map();

export function freeBusySupported(account) {
  // Mirrors the GAL toggle: undefined / missing counts as enabled, only
  // an explicit false disables.
  if (account?.custom?.freebusyenabled === false) return false;
  // AS 2.5 has no Availability element in ResolveRecipients at all.
  if (account?.custom?.asversion === "2.5") return false;
  return easCommandAdvertised(account, "ResolveRecipients");
}

function cacheKey(accountId, address, start, end) {
  return `${accountId}\u0000${address}\u0000${start.getTime()}\u0000${end.getTime()}`;
}

function pruneCache(now) {
  for (const [key, entry] of cache) {
    if (now - entry.at > CACHE_TTL_MS) cache.delete(key);
  }
}

/** Drop everything an account contributed - on disable, on delete, and
 *  whenever its toggle changes, so a re-enable is not answered from
 *  results gathered under the old setting. */
export function forgetFreeBusyCache(accountId) {
  if (!accountId) {
    cache.clear();
    return;
  }
  const prefix = `${accountId}\u0000`;
  for (const key of cache.keys()) {
    if (key.startsWith(prefix)) cache.delete(key);
  }
}

/** One account's answer for one address, as raw MergedFreeBusy digits.
 *  Returns null when the account cannot answer - which is not an error,
 *  and must not become one. */
async function askAccount({ provider, account, address, window }) {
  const accountId = account.accountId;
  // Re-read the account: tokens and server URLs change under us, and
  // this runs outside the sync path where that is normally refreshed.
  const rv = await provider.getAccount(accountId);
  const fresh = rv?.account;
  if (!fresh || !freeBusySupported(fresh)) return null;
  // Presenting rejected credentials on every keystroke is how a server
  // decides to lock an account out - the same guard the GAL search has.
  if (fresh.error === ERR.AUTH) return null;
  if (isOAuthAccount(fresh.custom)) {
    primeAuth(accountId, {
      refreshToken: fresh.custom?.refreshToken,
      servertype: fresh.custom?.servertype,
    });
  }
  const reply = await runResolveRecipients({
    account: fresh,
    asVersion: fresh.custom?.asversion ?? "14.1",
    to: address,
    start: window.start,
    end: window.end,
  });
  // Three statuses have to agree, and the middle one matters most: a
  // per-response status other than 1 is how the server says the address
  // was ambiguous, and the reply then describes whichever recipient it
  // listed first - someone else's day, painted into the grid as this
  // person's. Absence is still tolerated, as everywhere else here.
  const ok =
    reply?.status === "1" &&
    (reply.responseStatus === null || reply.responseStatus === "1") &&
    reply.availabilityStatus === "1";
  if (!ok) {
    provider.reportEventLog?.({
      level: "debug",
      accountId,
      message:
        `[free-busy] no availability for ${address}: status=${reply?.status} ` +
        `response=${reply?.responseStatus} availability=${reply?.availabilityStatus}`,
    });
    return null;
  }
  return reply.mergedFreeBusy ?? null;
}

/** Everything this add-on knows about `address` in `[start, end)`. */
async function lookup({ provider, address, start, end, types }) {
  let accounts;
  try {
    accounts = await provider.listAccounts();
  } catch {
    return [];
  }
  if (!Array.isArray(accounts)) return [];

  const matching = accounts.filter(
    (a) =>
      a?.enabled !== false &&
      freeBusySupported(a) &&
      accountCanAnswerFor(a, address),
  );
  if (!matching.length) {
    // Said out loud: no request is the correct answer for an address
    // that belongs to nobody here, and silence is otherwise
    // indistinguishable from the feature being broken.
    provider.reportEventLog?.({
      level: "debug",
      message: `[free-busy] ${address} matches no account - not asked`,
    });
    return [];
  }

  const window = alignWindow(start, end);
  const now = Date.now();
  pruneCache(now);

  let served = 0;
  const answers = await Promise.all(
    matching.map((account) => {
      const key = cacheKey(
        account.accountId,
        address,
        window.start,
        window.end,
      );
      const hit = cache.get(key);
      if (hit) {
        served++;
        return hit.promise;
      }
      // Cached before it resolves, so a second identical lookup joins
      // this request rather than starting another. A rejection is
      // cached too - a server that just refused will refuse again, and
      // retrying it per keystroke is the thing to avoid.
      const promise = askAccount({ provider, account, address, window }).catch(
        (err) => {
          provider.reportEventLog?.({
            level: "debug",
            accountId: account.accountId,
            message: `[free-busy] lookup failed for ${address}: ${err?.message ?? String(err)}`,
          });
          return null;
        },
      );
      cache.set(key, { at: now, promise });
      return promise;
    }),
  );

  const intervals = answers.flatMap((mergedFreeBusy) =>
    intervalsFromMergedFreeBusy({
      mergedFreeBusy,
      askedStart: window.start,
      wantStart: start,
      wantEnd: end,
      types,
    }),
  );

  // One line per lookup, naming what was asked and where the answer came
  // from. A cache hit issues no request at all, so without this the
  // difference between "answered from cache" and "never ran" is
  // invisible in the log.
  const busy = intervals.filter((i) => i.type !== "free").length;
  provider.reportEventLog?.({
    level: "debug",
    message:
      `[free-busy] ${address} ${window.start.toISOString()}..` +
      `${window.end.toISOString()} -> ${intervals.length} interval(s), ` +
      `${busy} not free, from ${matching.length} account(s)` +
      (served ? ` (${served} from cache)` : ""),
  });
  return intervals;
}

/** Register the single listener, if any account can answer. Idempotent. */
export async function refreshFreeBusyListener(provider) {
  let accounts = [];
  try {
    accounts = await provider.listAccounts();
  } catch (err) {
    console.debug("[eas] refreshFreeBusyListener: listAccounts failed:", err);
    return;
  }
  if (!Array.isArray(accounts)) return;
  const wanted = accounts.some(
    (a) => a?.enabled !== false && freeBusySupported(a),
  );
  if (wanted === !!listener) return;

  if (!wanted) {
    try {
      messenger.calendar.provider.onFreeBusy.removeListener(listener);
    } catch (err) {
      console.debug("[eas] free-busy removeListener failed:", err);
    }
    listener = null;
    forgetFreeBusyCache();
    return;
  }

  // The event arrives as four positional arguments - the schema
  // describes one object, but the parent fires
  // `fire.async(attendee, start, end, types)`. Both shapes are accepted
  // so the listener does not depend on which one wins: reading the
  // object shape off a string yields undefined everywhere, and the
  // symptom is a silently empty grid rather than an error.
  listener = async (first, maybeStart, maybeEnd, maybeTypes) => {
    const opts =
      first && typeof first === "object"
        ? first
        : {
            addressee: first,
            start: maybeStart,
            end: maybeEnd,
            types: maybeTypes,
          };
    try {
      return await lookup({
        provider,
        // Normalised once, here: routing already compares case-folded,
        // and an un-normalised address would otherwise cache and ask
        // twice for one person typed two ways.
        address: String(opts.addressee ?? "")
          .replace(/^mailto:/i, "")
          .trim()
          .toLowerCase(),
        start: new Date(opts.start),
        end: new Date(opts.end),
        types: opts.types,
      });
    } catch (err) {
      // Never throw into the dialog: an empty answer leaves the grid
      // blank, an exception leaves the user with a broken one.
      provider.reportEventLog?.({
        level: "debug",
        message: `[free-busy] listener failed: ${err?.message ?? String(err)}`,
      });
      return [];
    }
  };
  try {
    messenger.calendar.provider.onFreeBusy.addListener(listener);
  } catch (err) {
    listener = null;
    provider.reportEventLog?.({
      level: "warning",
      message: `[free-busy] could not register the listener: ${err?.message ?? String(err)}`,
    });
  }
}
