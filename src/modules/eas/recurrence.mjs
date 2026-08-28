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

/**
 * Recurrence elements neither codec models, and the property each is parked
 * in so a push can hand it straight back.
 *
 * `Regenerate` schedules the next occurrence from the *completion* of the
 * last one rather than from the pattern - Outlook's "regenerating task".
 * `DeadOccur` marks an item as a spent occurrence rather than the live
 * series, which is how Exchange records history for a task series that has
 * no exceptions. Both are `[MS-ASTASK]` only. `CalendarType` states which
 * calendar system the rule is computed in and `IsLeapMonth` qualifies a
 * monthly rule inside one, and those two exist in both namespaces.
 *
 * None is representable here. iCalendar has no completion-relative
 * recurrence at all; a spent occurrence would be a RECURRENCE-ID override,
 * which is a different shape from the separate item with its own ServerId
 * that Exchange actually sends; and the non-Gregorian pair maps to RFC 7529
 * `RSCALE`, which neither ical.js nor Thunderbird implements. All four are
 * invisible in the UI.
 *
 * They are kept anyway, because a push rebuilds `<Recurrence>` from the
 * RRULE and the server replaces the block wholesale - measured on both 14.1
 * and 16.1 by sending a weekly rule, changing it to daily, and finding the
 * omitted `DayOfWeek` gone from the server's own copy. So an edit that has
 * nothing to do with recurrence used to destroy them.
 *
 * Here rather than in one codec for the reason in the module header: this
 * started as a task-only fix, and the calendar codec had the identical hole.
 *
 * `FirstDayOfWeek` is not one of these. It is the one element of the set
 * iCalendar can express, so carrying it is not enough - see
 * `FIRST_DAY_OF_WEEK`.
 */
export const UNMAPPED_RECURRENCE = Object.freeze({
  Regenerate: "x-eas-regenerate",
  DeadOccur: "x-eas-deadoccur",
  CalendarType: "x-eas-calendartype",
  IsLeapMonth: "x-eas-isleapmonth",
});

/** `FirstDayOfWeek`, kept apart from the four above because it is handled
 *  rather than merely carried: `easToRrule` renders it as `WKST` so the
 *  rule expands correctly here, and `firstDayOfWeekOf` decides what goes
 *  back. It is also written *last* in the block, where both servers put it,
 *  while the four above come first - so it cannot ride the same emit loop.
 *
 *  The stamp is what lets the push tell "the mailbox said Sunday" apart
 *  from "this rule was written here". */
export const FIRST_DAY_OF_WEEK = Object.freeze({
  tag: "FirstDayOfWeek",
  prop: "x-eas-firstdayofweek",
});

/** Park the elements above on `comp` (a VEVENT or VTODO), and clear the ones
 *  the server has stopped sending - a stale stamp would be re-asserted for
 *  ever. Call it only when the AD actually carried a `<Recurrence>`: absence
 *  means "unchanged", not "gone". */
export function keepUnmappedRecurrence(comp, recNode) {
  for (const [tag, prop] of [
    ...Object.entries(UNMAPPED_RECURRENCE),
    [FIRST_DAY_OF_WEEK.tag, FIRST_DAY_OF_WEEK.prop],
  ]) {
    comp.removeAllProperties(prop);
    const value = readPathFrom(recNode, [tag]);
    if (value != null && value !== "") comp.addPropertyWithValue(prop, value);
  }
}

/** What `keepUnmappedRecurrence` parked for the head of the block, as
 *  `[tag, value]` pairs. Empty for an item the server said nothing about. */
export function unmappedRecurrenceOf(comp) {
  const out = [];
  for (const [tag, prop] of Object.entries(UNMAPPED_RECURRENCE)) {
    const value = comp.getFirstPropertyValue(prop);
    if (value != null && value !== "") out.push([tag, String(value)]);
  }
  return out;
}

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
  // The day the week starts on, which decides how an interval-above-1
  // weekly rule groups its days. EAS numbers it 0=Sunday..6=Saturday,
  // which is exactly `ICAL_DAYS` - and exactly Thunderbird's own
  // `calIRecurrenceRule.weekStart`, so nothing has to be converted.
  //
  // Mapped so an interval-above-1 weekly rule expands here the way the
  // mailbox shows it. What goes back out is not read from this rule -
  // `firstDayOfWeekOf` decides that, because Thunderbird stamps `weekStart`
  // from the `calendar.week.start` preference on any edit and offers no
  // control for it (the dialog's `week-start` input is `type="hidden"`).
  const fdow = readPathFrom(recNode, ["FirstDayOfWeek"]);
  const wkst = ICAL_DAYS[parseInt(fdow, 10)];
  if (wkst) parts.push(`WKST=${wkst}`);

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

/**
 * The `FirstDayOfWeek` to send, for the *end* of the block, or null.
 *
 * An item the server gave us echoes back what the server last said, so a
 * mailbox's own week start is never overwritten by ours - Thunderbird
 * stamps `weekStart` from `calendar.week.start` on any edit of a rule and
 * offers no control for it, so the local value is not the user's choice in
 * any meaningful sense.
 *
 * Otherwise it goes out only if the rule genuinely names one, and that has
 * to be read off the serialised rule rather than the parsed object: ical.js
 * gives every parsed rule a `wkst`, defaulting to Monday, so the parsed
 * value cannot tell a rule that states a week start from one that says
 * nothing about it. `toICALString` round-trips the source and can - it
 * keeps `WKST=MO` when it was written and adds nothing when it was not.
 *
 * Reading the parsed value instead put `FirstDayOfWeek: 1` on every
 * recurring item authored here, on every push, at every frequency -
 * including yearly rules on a fixed date, where a week start means nothing
 * at all. Nobody had expressed that preference; ical.js had.
 *
 * ical.js counts 1=Sunday..7=Saturday where EAS counts 0=Sunday, hence the
 * offset.
 */
export function firstDayOfWeekOf(comp, rruleProp = null) {
  const stamped = comp.getFirstPropertyValue(FIRST_DAY_OF_WEEK.prop);
  if (stamped != null && stamped !== "") return String(stamped);
  const written = rruleProp?.toICALString?.() ?? "";
  if (!/;WKST=/i.test(written)) return null;
  const wkst = rruleProp?.getFirstValue?.()?.wkst;
  return wkst ? String(wkst - 1) : null;
}
