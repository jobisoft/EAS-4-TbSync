/**
 * One UID, several ServerIds: the same item held more than once by the
 * server.
 *
 * Thunderbird stores exactly one item per UID, so a mailbox in this state
 * is invisible locally - the pull adopts each copy onto the one local item
 * in turn and the user sees nothing wrong, while every copy stays on the
 * server and comes down again on the next full sync. Peters' calendar had
 * three such clusters, one of them 533 copies deep.
 *
 * What counts as evidence is deliberately narrow: only the ServerIds this
 * sync actually saw the server claim for a UID. The stamp the local item
 * was already carrying is not evidence, because a server below 16.1 may
 * re-mint its ServerIds after a resync - every item in the folder then
 * arrives under a new id, and a rule that compared against the stamp would
 * report the whole calendar as duplicated and offer to delete it.
 * Re-minting still gives each item exactly one id, so counting claims is
 * immune to it.
 *
 * The consequence is that a cluster split across syncs - one copy last
 * week, one today - is not seen until a sync brings more than one of them
 * down together. A full pull always does.
 *
 * The copy to keep is the one Thunderbird ended up bound to, which is the
 * last the pull adopted. It needs no choosing: it is what the index map
 * says, and keeping anything else would mean rewriting the local item.
 */

/** Note that the server named `serverId` as holding `uid`, in this sync. */
export function noteUidClaim(claims, uid, serverId) {
  if (!claims || !uid || !serverId) return;
  let ids = claims.get(uid);
  if (!ids) claims.set(uid, (ids = new Set()));
  ids.add(serverId);
}

/** Pull a human-readable title out of a stored item.
 *
 *  Blobs are iCalendar for calendars and vCard for address books, so one
 *  of `SUMMARY` and `FN` is the title and neither appears in the other
 *  format. Folded continuation lines are joined first, because a long
 *  title is wrapped at 75 octets and would otherwise be shown cut off. */
export function titleFromBlob(blob) {
  if (!blob) return "";
  const unfolded = String(blob).replace(/\r?\n[ \t]/g, "");
  const hit = /^(?:SUMMARY|FN)(?:;[^:\r\n]*)?:(.*)$/im.exec(unfolded);
  return hit ? hit[1].trim() : "";
}

/**
 * One entry per duplicated UID, or an empty array when the sync saw none.
 *
 * `serverIdFor` and `titleFor` are passed in rather than read from a sync
 * context so this is testable on its own; the sync supplies the index map
 * and the local store.
 */
export async function duplicateClusters(claims, { serverIdFor, titleFor }) {
  const clusters = [];
  for (const [uid, ids] of claims ?? []) {
    if (ids.size < 2) continue;
    const keeper = serverIdFor(uid);
    // No keeper means nothing local is bound to this UID any more - the
    // item was deleted during the very sync that saw the copies. Leaving
    // it out is the safe answer: with no surviving copy to keep, every
    // id in the cluster would be a candidate for deletion, and offering
    // that on the strength of a race is how real items get lost.
    if (!keeper) continue;
    const surplus = [...ids].filter((id) => id !== keeper);
    if (!surplus.length) continue;
    clusters.push({ uid, keeper, surplus, title: await titleFor(uid) });
  }
  return clusters;
}
