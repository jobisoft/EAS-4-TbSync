/**
 * Which calendars we have to mail the attendees for ourselves.
 *
 * A 14.x server does not send the invitation, the update or the
 * cancellation when we push a meeting - measured, see the item-48 notes -
 * and the provider declares `scheduling: "server"`, so Thunderbird does not
 * send one either. On those accounts the message is ours to send. From 16.0
 * the server does it and we must stay silent, and 2.5 cannot send at all.
 *
 * ## Why this is a cache rather than a question
 *
 * The answer is needed inside the item hooks, which are the only place the
 * previous version of an edited item exists. A hook is holding the user's
 * save and must work with the host absent - that is the whole reason
 * `record()` writes to our own storage instead of asking - so it cannot ask
 * the host what version an account negotiated.
 *
 * So the sync banks it: it knows the version and the target, and it runs
 * often enough that the answer is never stale for long. A calendar cannot
 * be edited before a sync has bound it, so the value exists by the time any
 * hook can fire.
 *
 * Absent means **no**. A missing answer sends nothing, which costs a
 * notification nobody was expecting; guessing yes would mail every attendee
 * of every meeting on an account that had already told them itself.
 */

const KEY = "eas.clientScheduling";

/** The versions where sending is the client's job: 14.0 and 14.1.
 *
 *  Not "below 16". 2.5 is below 16 and cannot do this at all - `sendMail`
 *  builds WBXML where 2.5 wants raw MIME, and the codec emits no attendee
 *  block there - so a note recorded on a 2.5 account could never be sent
 *  and would be dropped unsent on every sync. */
export function versionNeedsClientScheduling(asVersion) {
  const v = parseFloat(asVersion);
  return Number.isFinite(v) && v >= 14 && v < 16;
}

/** Bank the answer for one bound calendar, as `{ user }` - present means
 *  yes, absent means no.
 *
 *  The account's own address rides along because the hook needs it for the
 *  same reason it needs the version, and cannot ask for it either: telling
 *  a meeting we organise from an invitation somebody sent us is a
 *  comparison against that address.
 *
 *  Called by the sync, the only place holding the target, the negotiated
 *  version and the address at once. Best-effort: a write that fails costs a
 *  notification, never an edit. */
export async function rememberClientScheduling(targetID, needed, user = "") {
  if (!targetID) return;
  try {
    const rv = await browser.storage.local.get({ [KEY]: {} });
    const map = rv[KEY] ?? {};
    const want = needed && user ? { user } : null;
    if ((map[targetID]?.user ?? null) === (want?.user ?? null)) return;
    if (want) map[targetID] = want;
    else delete map[targetID];
    await browser.storage.local.set({ [KEY]: map });
  } catch {
    /* best-effort, see above */
  }
}

/** `{ user }` when this calendar's attendees are ours to notify, else null.
 *  Null is the answer for every 16.x and 2.5 calendar, for one we have
 *  never synced, and for anything that went wrong reading it. */
export async function clientSchedulingFor(targetID) {
  if (!targetID) return null;
  try {
    const rv = await browser.storage.local.get({ [KEY]: {} });
    const found = (rv[KEY] ?? {})[targetID];
    return found?.user ? found : null;
  } catch {
    return null;
  }
}

/** Forget a calendar that is no longer bound, so the map does not grow one
 *  entry per binding this profile has ever had. */
export async function forgetClientScheduling(targetID) {
  await rememberClientScheduling(targetID, false);
}
