/**
 * Negotiate an EAS protocol version with the server. The OPTIONS probe
 * returns the server's supported versions and commands; we pick the
 * highest version present in both the server's list and our supported
 * set.
 *
 * Preference order is 16.1 → 14.1 → 14.0 → 2.5. 16.1 leads: it is the
 * protocol's current form (per-instance exception commands, saner
 * all-day semantics per [MS-ASCAL] §2.2.2.1).
 *
 * What this produces is a *suggestion*, not the version an account runs
 * on. An account runs on what its config says: the version the user
 * pinned, or this suggestion when the config says `auto`. Only connecting
 * reads that - see `#doConnectAndDiscover` in eas-provider.
 */

import { ERR, withCode } from "../../vendor/tbsync/provider.mjs";
import { easOptions, NET_ERR } from "../network.mjs";

const SUPPORTED = ["16.1", "14.1", "14.0", "2.5"];

/** Thrown when the OPTIONS probe could not tell us what the server
 *  speaks.
 *
 *  Typed, because the caller has somewhere to go from here that a
 *  transport error does not deserve: a version the user pinned, or one
 *  this account was negotiated onto before, will both serve. Only when
 *  there is no such version is this a dead end. */
export const NO_OPTIONS_ANSWER = Symbol("NO_OPTIONS_ANSWER");

export function isNoOptionsAnswer(err) {
  return err?.easReason === NO_OPTIONS_ANSWER;
}

function noOptionsAnswer(cause) {
  const err = withCode(
    new Error(
      browser.i18n.getMessage("eas.connect.error.optionsUnavailable") ||
        "The server did not answer the ActiveSync version probe.",
    ),
    ERR.UNKNOWN_COMMAND,
  );
  err.easReason = NO_OPTIONS_ANSWER;
  err.cause = cause;
  return err;
}

/** What the server says it speaks, or a failure the caller can act on.
 *
 *  A refused credential and a redirect stay exactly what they are. Every
 *  other failure becomes `NO_OPTIONS_ANSWER`, because a server that will
 *  not answer this one request may serve every other perfectly well, and
 *  the connect has somewhere to go from there. */
async function probeVersions({ account }) {
  try {
    return await easOptions({ account });
  } catch (err) {
    if (err?.code === NET_ERR.AUTH) throw err;
    if (err?.code === NET_ERR.HOST_REDIRECT) throw err;
    throw noOptionsAnswer(err);
  }
}

/** The best overlap between what a server advertises and what we speak,
 *  or null when there is none. Derived rather than stored: every OPTIONS
 *  answer rewrites the advertised list, so a stored copy would say the
 *  same thing and could only drift from it. */
export function suggestedAsVersion(serverVersions) {
  return SUPPORTED.find((v) => (serverVersions ?? []).includes(v)) ?? null;
}

export async function negotiateAsVersion({ account }) {
  const probe = await probeVersions({ account });
  const serverVersions = probe.versions ?? [];
  if (serverVersions.length === 0) {
    throw noOptionsAnswer(null);
  }
  const asVersion = suggestedAsVersion(serverVersions);
  if (!asVersion) {
    throw withCode(
      new Error(
        `No mutually-supported EAS version (server: ${serverVersions.join(",")})`,
      ),
      ERR.UNKNOWN_COMMAND,
    );
  }
  return {
    asVersion,
    allowedAsVersions: serverVersions,
    allowedCommands: probe.commands ?? [],
  };
}
