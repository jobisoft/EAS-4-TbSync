/**
 * iCal RRULE ⇆ ActiveSync `<Recurrence>` mapping, shared by the calendar and
 * task codecs.
 *
 * Only the *mapping* lives here: which EAS recurrence type an RRULE is, which
 * of the qualifying fields that type needs, and the reverse. Emitting is left
 * to each codec, because the two disagree on everything else - `[MS-ASCAL]`
 * and `[MS-ASTASK]` are different namespaces with different element sets (a
 * task carries `<Start>`, an event does not), and they format `Until`
 * differently. Reading has no such split, so `easToRrule` is shared whole.
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

import { readPathFrom } from "./wbxml-helpers.mjs";

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

/**
 * The reverse: an ActiveSync `<Recurrence>` node → an RRULE string, or null if
 * the node carries no type this maps.
 *
 * Shared verbatim. Both codecs carried byte-equivalent copies of this - the
 * only differences were blank lines and one statement reorder - which is the
 * same duplication that let the outbound halves drift until a task could only
 * recur daily. One copy means a mapping fixed here is fixed for both.
 *
 * @param {object} recNode  the `<Recurrence>` node
 * @returns {string|null}
 */
export function easToRrule(recNode) {
  const type = readPathFrom(recNode, ["Type"]);
  const freq = {
    0: "DAILY",
    1: "WEEKLY",
    2: "MONTHLY",
    3: "MONTHLY",
    5: "YEARLY",
    6: "YEARLY",
  }[type];
  if (!freq) return null;
  const parts = [`FREQ=${freq}`];
  const interval = readPathFrom(recNode, ["Interval"]);
  if (interval) parts.push(`INTERVAL=${interval}`);

  const dow = readPathFrom(recNode, ["DayOfWeek"]);
  if (dow) {
    const bits = parseInt(dow, 10) || 0;
    const week = readPathFrom(recNode, ["WeekOfMonth"]);
    const days = [];
    for (let i = 0; i < 7; i++) if (bits & (1 << i)) days.push(ICAL_DAYS[i]);
    if (days.length) {
      // WeekOfMonth 5 is "last", which iCal spells -1.
      const prefix = week === "5" ? "-1" : week ? String(week) : "";
      parts.push("BYDAY=" + days.map((d) => prefix + d).join(","));
    }
  }
  const dom = readPathFrom(recNode, ["DayOfMonth"]);
  if (dom) parts.push(`BYMONTHDAY=${dom}`);
  const moy = readPathFrom(recNode, ["MonthOfYear"]);
  if (moy) parts.push(`BYMONTH=${moy}`);

  const occ = readPathFrom(recNode, ["Occurrences"]);
  if (occ) parts.push(`COUNT=${occ}`);
  const until = readPathFrom(recNode, ["Until"]);
  // Both EAS date shapes reach here, because this is shared: an event's
  // `Until` is compact basic (20261001T000000Z) and needs nothing removed,
  // while a task's is extended with milliseconds
  // (2026-10-01T00:00:00.000Z). iCal's DATE-TIME has no fractional part, so
  // the fraction has to go along with the separators - dropping only the
  // separators left `20261001T000000.000Z`, which ICAL.js does not accept and
  // normalises by discarding the fraction *and* the `Z`. That silently turned
  // a UTC UNTIL into a floating one, which RFC 5545 forbids next to a UTC
  // DTSTART.
  if (until) parts.push(`UNTIL=${until.replace(/[-:]|\.\d+/g, "")}`);

  return parts.join(";");
}
