/**
 * Turning EAS availability into the intervals Thunderbird's attendee
 * grid wants, and back.
 *
 * `MergedFreeBusy` is a string of digits, one per time slot, laid end to
 * end from the StartTime that was asked for. Everything below follows
 * from what the two server families actually do, measured 10 Aug 2026
 * against Z-Push 14.1 and Exchange Online 16.1:
 *
 *   - A slot is 30 minutes. Fixed, not the window divided by the string's
 *     length: 3 hours answered 6 digits and a day answered 48, but a
 *     20-minute window answered 2 - dividing there would invent 10-minute
 *     slots and put every busy block in the wrong place.
 *   - The grid starts at the requested StartTime, unsnapped, and a slot
 *     counts as busy if the event overlaps it at all. Asking 09:07 for an
 *     event at 10:00-11:00 marked three slots busy, 09:37 through 11:07.
 *     So a request must start on a slot boundary or an hour-long meeting
 *     reads as 90 minutes.
 *   - A partial slot is the one place the families disagree: Z-Push
 *     widens the window and answers anyway, Exchange Online refuses the
 *     whole request with Status 5. Asking only for whole slots satisfies
 *     both, which is why alignment happens here rather than in a
 *     per-server branch.
 */

export const SLOT_MS = 30 * 60 * 1000;

/** [MS-ASCMD] MergedFreeBusy digits, in the vocabulary the
 *  `calendar.provider.onFreeBusy` listener answers with. Anything else
 *  the server invents reads as "unknown" rather than being guessed at. */
const TYPE_FOR_DIGIT = {
  0: "free",
  1: "tentative",
  2: "busy",
  3: "unavailable",
  4: "unknown",
};

/** The whole-slot window covering `[start, end)`. Both server families
 *  answer identically for one of these, and neither has to widen or
 *  refuse it. */
export function alignWindow(start, end) {
  const from = Math.floor(start.getTime() / SLOT_MS) * SLOT_MS;
  let to = Math.ceil(end.getTime() / SLOT_MS) * SLOT_MS;
  // A zero-length ask still needs a slot to be answerable at all.
  if (to <= from) to = from + SLOT_MS;
  return { start: new Date(from), end: new Date(to) };
}

/**
 * Expand `MergedFreeBusy` into intervals, clipped back to what the
 * caller actually asked about.
 *
 * `askedStart` is the aligned start the request carried - slot 0 begins
 * there. `wantStart`/`wantEnd` are Thunderbird's original bounds; the
 * outer slots are trimmed to them so the grid does not show availability
 * for time nobody asked about. Neighbouring slots of the same type merge
 * into one interval, which is both what the UI wants and far fewer
 * objects for a day-long window.
 *
 * The string's own length rules: the server is the authority on how much
 * it answered, so a reply longer or shorter than the window is read for
 * as many slots as it has rather than being stretched to fit.
 */
export function intervalsFromMergedFreeBusy({
  mergedFreeBusy,
  askedStart,
  wantStart,
  wantEnd,
  types,
}) {
  if (typeof mergedFreeBusy !== "string" || !mergedFreeBusy) return [];
  const base = askedStart.getTime();
  const lower = wantStart ? wantStart.getTime() : -Infinity;
  const upper = wantEnd ? wantEnd.getTime() : Infinity;
  const wanted = types?.length ? new Set(types) : null;

  const out = [];
  for (let i = 0; i < mergedFreeBusy.length; i++) {
    const type = TYPE_FOR_DIGIT[mergedFreeBusy[i]] ?? "unknown";
    const from = Math.max(base + i * SLOT_MS, lower);
    const to = Math.min(base + (i + 1) * SLOT_MS, upper);
    if (to <= from) continue; // entirely outside what was asked
    const last = out[out.length - 1];
    if (last && last.type === type && last.to === from) last.to = to;
    else out.push({ type, from, to });
  }

  return out
    .filter((iv) => !wanted || wanted.has(iv.type))
    .map((iv) => ({
      start: new Date(iv.from).toISOString(),
      end: new Date(iv.to).toISOString(),
      type: iv.type,
    }));
}

/** Does this account's mailbox put it in a position to answer for
 *  `address`? True for the account's own address, and for anything
 *  sharing its domain - colleagues, which is who the attendee grid is
 *  usually about. An account with no known address matches nothing:
 *  guessing from the login would send strangers' addresses to a server
 *  that has no business seeing them. */
export function accountCanAnswerFor(account, address) {
  const own = account?.custom?.userSmtpAddress;
  if (!own || typeof address !== "string") return false;
  const a = address.trim().toLowerCase();
  const mine = String(own).trim().toLowerCase();
  // Both sides must actually be addresses. Without this, a bare domain
  // typed into the attendee field would match on the domain compare and
  // be sent to the server as if it were a person.
  if (!a.includes("@") || !mine.includes("@")) return false;
  if (a === mine) return true;
  const domain = (s) => s.slice(s.lastIndexOf("@") + 1);
  const d = domain(a);
  return !!d && d === domain(mine);
}
