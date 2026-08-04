/**
 * iCal RRULE → ActiveSync `<Recurrence>` field derivation, shared by the
 * calendar and task codecs.
 *
 * Only the *mapping* lives here: which EAS recurrence type an RRULE is, and
 * which of the qualifying fields that type needs. Emitting is left to each
 * codec, because the two disagree on everything else - `[MS-ASCAL]` and
 * `[MS-ASTASK]` are different namespaces with different element sets (a task
 * carries `<Start>`, an event does not), and they format `Until` differently.
 *
 * Shared because they had drifted. The task codec was a stripped-down copy
 * that emitted the type but none of the fields qualifying it, so every task
 * recurrence except daily was rejected with Status 6 while the identical rule
 * on an event was accepted. Splitting mapping from emitting is what stops one
 * side quietly losing a case the other handles.
 *
 * Which fields a type requires, per `[MS-ASTASK]` 2.2.2.31 and `[MS-ASCAL]`
 * 2.2.2.35:
 *
 *   0  daily        -
 *   1  weekly       DayOfWeek
 *   2  monthly      DayOfMonth
 *   3  monthly nth  DayOfWeek + WeekOfMonth
 *   5  yearly       DayOfMonth + MonthOfYear
 *   6  yearly nth   DayOfWeek + WeekOfMonth + MonthOfYear
 *
 * Daily needing nothing is exactly why it was the only kind of recurring task
 * that worked.
 */

/** Sunday-first, matching the EAS `DayOfWeek` bitmask: SU=1 … SA=64. */
const ICAL_DAYS = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];

/**
 * @param {object} rruleProp   the iCal RRULE property
 * @param {object} startProp   the item's DTSTART property, used to fill in a
 *                             qualifying field the rule leaves implicit
 * @returns {null | {
 *   type: number, dayOfMonth: number|null, dayOfWeek: number|null,
 *   monthOfYear: number|null, weekOfMonth: number|null,
 *   interval: number, count: number|null, until: object|null,
 * }}  `until` is the raw value, for the caller to format.
 */
export function rruleToEas(rruleProp, startProp) {
  const r = rruleProp?.getFirstValue?.(); // ICAL.Recur
  if (!r) return null;

  const startDate = startProp?.getFirstValue?.();
  let type = 0;
  let monthDays = r.parts?.BYMONTHDAY ?? [];
  let weekDays = (r.parts?.BYDAY ?? []).slice();
  let months = r.parts?.BYMONTH ?? [];
  const weeks = [];

  // Unpack ±NDD style days into weekDays + weekOfMonth.
  for (let i = 0; i < weekDays.length; i++) {
    const m = /^([+-]?\d*)(SU|MO|TU|WE|TH|FR|SA)$/.exec(weekDays[i]);
    if (!m) continue;
    const n = parseInt(m[1] || "0", 10);
    const dow = ICAL_DAYS.indexOf(m[2]) + 1;
    weekDays[i] = dow;
    if (n) weeks[i] = n === -1 ? 5 : n;
  }

  // A rule may leave its qualifying field implicit - `FREQ=WEEKLY` with no
  // BYDAY means "the weekday DTSTART falls on". EAS has no such shorthand and
  // rejects the rule outright, so the implied value is filled in here.
  if (r.freq === "WEEKLY") {
    type = 1;
    if (!weekDays.length && startDate) {
      weekDays = [startDate.dayOfWeek?.() ?? 1];
    }
  } else if (r.freq === "MONTHLY" && weeks.length) {
    type = 3;
  } else if (r.freq === "MONTHLY") {
    type = 2;
    if (!monthDays.length && startDate) monthDays = [startDate.day];
  } else if (r.freq === "YEARLY" && weeks.length) {
    type = 6;
  } else if (r.freq === "YEARLY") {
    type = 5;
    if (!monthDays.length && startDate) monthDays = [startDate.day];
    if (!months.length && startDate) months = [startDate.month];
  }

  let dayOfWeek = null;
  if (weekDays.length) {
    let bits = 0;
    for (const d of weekDays) bits |= 1 << (d - 1);
    dayOfWeek = bits;
  }

  return {
    type,
    dayOfMonth: monthDays[0] ?? null,
    dayOfWeek,
    monthOfYear: months.length ? months[0] : null,
    weekOfMonth: weeks.length ? weeks[0] : null,
    interval: r.interval ?? 1,
    // Both may be set; which one wins is the caller's to decide, since it is
    // an emit rule (`Occurrences` xor `Until`) rather than a fact about the
    // rule.
    count: r.count ?? null,
    until: r.until ?? null,
  };
}
