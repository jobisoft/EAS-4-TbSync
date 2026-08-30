/**
 * EAS Tasks (codepage 9) + AirSyncBase Body (17) ⇆ iCal VTODO codec.
 *
 * Mirrors the legacy `EAS-4-TbSync/content/includes/tasksync.js` mapping.
 * Round-trips: Subject, Body, Importance, Sensitivity, Categories,
 * StartDate/DueDate (with UtcStart/UtcDue pairing), Complete +
 * DateCompleted, ReminderSet + ReminderTime, basic recurrence (RRULE).
 *
 * EAS Tasks use extended-ISO date strings (YYYY-MM-DDTHH:MM:SS.sssZ) for
 * dates; events use compact basic ISO. Reminder times are absolute UTC.
 */

import ICAL from "../../vendor/ical.min.js";
import { readPathFrom, isFiletimeZero } from "./wbxml-helpers.mjs";
import {
  readBodyIntoDescription,
  appendBodyFromDescription,
} from "./body-codec.mjs";
import {
  rruleToEas,
  easToRrule,
  keepUnmappedRecurrence,
  unmappedRecurrenceOf,
  firstDayOfWeekOf,
} from "./recurrence.mjs";
import {
  guessTimezoneByCurrentOffset,
  getIcalTimezone,
} from "./timezone-mapping.mjs";

const X_EAS_SERVERID = "X-EAS-SERVERID";

const IMPORTANCE_TO_PRIORITY = { 0: "9", 1: "5", 2: "1" };
const PRIORITY_TO_IMPORTANCE = { 9: "0", 5: "1", 1: "2" };

const SENSITIVITY_TO_CLASS = {
  0: "PUBLIC",
  1: "PRIVATE",
  2: "PRIVATE",
  3: "CONFIDENTIAL",
};
const CLASS_TO_SENSITIVITY = { PUBLIC: "0", PRIVATE: "2", CONFIDENTIAL: "3" };

/* ── Reader: ApplicationData → iCal VTODO ──────────────────────────── */

export async function applicationDataToIcal({
  adNode,
  existingIcal,
  serverID,
  asVersion,
  defaultTimezone,
  syncRecurrence,
  msTodoCompat,
  uid,
  nativePlainText = null,
}) {
  // Merge mode: parse the existing iCal and overlay only fields the AD
  // mentions. Fall through to a fresh build for server-pushed Adds or
  // when the existing blob is unparseable.
  let vcal = null;
  let vtodo = null;
  if (existingIcal) {
    vcal = parseVCalendar(existingIcal);
    if (vcal) vtodo = vcal.getFirstSubcomponent("vtodo");
  }
  if (!vcal || !vtodo) {
    vcal = newVCalendar();
    vtodo = new ICAL.Component(["vtodo", [], []]);
    vcal.addSubcomponent(vtodo);
  }
  if (uid) vtodo.updatePropertyWithValue("uid", uid);
  vtodo.updatePropertyWithValue(X_EAS_SERVERID.toLowerCase(), serverID);

  const subject = readPathFrom(adNode, ["Subject"]);
  if (subject) vtodo.updatePropertyWithValue("summary", subject);

  // Body (codepage-aware).
  await readBodyIntoDescription(vtodo, adNode, {
    useAirSyncBase: asVersion !== "2.5",
    nativePlainText,
  });

  // Reminder is read up-front so the MS To-Do compatibility hack can pin
  // DTSTART/DUE to the reminder time before they are written.
  const reminderTime =
    readPathFrom(adNode, ["ReminderSet"]) === "1"
      ? readDate(adNode, "ReminderTime")
      : null;
  const msTodoOverride = msTodoCompat === true && !!reminderTime;

  // StartDate / DueDate with Utc* pairing for tz hint extraction.
  // `dtstartSource` is held over for the Reminder block below so a
  // relative-to-START alarm can anchor against whichever value
  // ultimately became DTSTART.
  let dtstartSource = null;
  if (msTodoOverride) {
    // MS To-Do only ships date-only due dates; pin both ends to the
    // reminder so Lightning renders the task on the correct day.
    writeUtcDateProp(vtodo, "dtstart", reminderTime);
    writeUtcDateProp(vtodo, "due", reminderTime);
    dtstartSource = reminderTime;
  } else {
    const startUtc = readDate(adNode, "UtcStartDate");
    const startLocal = readDate(adNode, "StartDate");
    const dueUtc = readDate(adNode, "UtcDueDate");
    const dueLocal = readDate(adNode, "DueDate");

    // Recover the server-hinted IANA zone for each pair via the
    // moment-in-time offset between the local-clock and UTC forms.
    // Mirrors legacy tasksync.js:48-69.
    const startTz = guessTzFromPair(startUtc, startLocal);
    const dueTz = guessTzFromPair(dueUtc, dueLocal);

    // If the server sends Recurrence without StartDate, fall back to
    // DueDate as the recurrence anchor — Lightning won't render a
    // recurring VTODO without a DTSTART. Mirrors legacy
    // tasksync.js:70-79. Only a no-op when both are absent (the
    // recurrence will then be lost on the local side, same as legacy).
    dtstartSource = startUtc;
    let dtstartTz = startTz;
    if (!dtstartSource && dueUtc && childByTag(adNode, "Recurrence")) {
      dtstartSource = dueUtc;
      dtstartTz = dueTz;
    }
    if (dtstartSource) {
      writeUtcDateProp(vtodo, "dtstart", dtstartSource, dtstartTz);
    }
    if (dueUtc) writeUtcDateProp(vtodo, "due", dueUtc, dueTz);
  }

  // Importance → PRIORITY.
  const importance = readPathFrom(adNode, ["Importance"]);
  if (importance && IMPORTANCE_TO_PRIORITY[importance]) {
    vtodo.updatePropertyWithValue(
      "priority",
      IMPORTANCE_TO_PRIORITY[importance],
    );
  }
  // Sensitivity → CLASS.
  const sens = readPathFrom(adNode, ["Sensitivity"]);
  if (sens && SENSITIVITY_TO_CLASS[sens]) {
    vtodo.updatePropertyWithValue("class", SENSITIVITY_TO_CLASS[sens]);
  }

  // Complete + DateCompleted. Merge-aware: <Complete> in the AD signals
  // the server's authoritative completion state.
  if (childByTag(adNode, "Complete")) {
    vtodo.removeAllProperties("status");
    vtodo.removeAllProperties("percent-complete");
    vtodo.removeAllProperties("completed");
    const complete = readPathFrom(adNode, ["Complete"]);
    if (complete === "1") {
      vtodo.updatePropertyWithValue("status", "COMPLETED");
      vtodo.updatePropertyWithValue("percent-complete", 100);
      const dc = readDate(adNode, "DateCompleted");
      if (dc) writeUtcDateProp(vtodo, "completed", dc);
    }
  }

  // Reminder. Three branches mirroring legacy tasksync.js:88-115:
  //   - msTodoOverride: alarm offset = 0 from a synthesised DTSTART.
  //   - DTSTART known (either real or DueDate-synthesised): store the
  //     VALARM as a Duration TRIGGER with RELATED=START so the alarm
  //     slides if the user moves DTSTART. Mirrors legacy lines 103-107.
  //   - No DTSTART anchor: store as an absolute DATE-TIME TRIGGER.
  //     Mirrors legacy lines 108-113.
  // Merge-aware: <ReminderSet> / <ReminderTime> in the AD signals the
  // server's authoritative alarm state - clear existing VALARMs first.
  const hasReminderTag =
    childByTag(adNode, "ReminderSet") != null ||
    childByTag(adNode, "ReminderTime") != null;
  if (hasReminderTag) {
    for (const a of vtodo.getAllSubcomponents("valarm")) {
      vtodo.removeSubcomponent(a);
    }
  }
  if (reminderTime) {
    if (msTodoOverride) {
      appendStartRelativeAlarm(vtodo, 0);
    } else if (dtstartSource) {
      const startMs = parseExtendedIso(dtstartSource)?.getTime();
      const reminderMs = parseExtendedIso(reminderTime)?.getTime();
      if (startMs != null && reminderMs != null) {
        const offsetSec = Math.round((reminderMs - startMs) / 1000);
        appendStartRelativeAlarm(vtodo, offsetSec);
      } else {
        appendAbsoluteAlarm(vtodo, reminderTime);
      }
    } else {
      appendAbsoluteAlarm(vtodo, reminderTime);
    }
  }

  // Categories. Merge-aware: clear existing when <Categories> is in
  // the AD; the AD's children replace the prior set.
  if (childByTag(adNode, "Categories")) {
    vtodo.removeAllProperties("categories");
    const cats = collectChildren(adNode, "Categories", "Category");
    if (cats.length) {
      const prop = new ICAL.Property("categories", vtodo);
      prop.setValues(cats);
      vtodo.addProperty(prop);
    }
  }

  // Recurrence (RRULE only; tasks have no exceptions in EAS).
  if (syncRecurrence) {
    const recNode = childByTag(adNode, "Recurrence");
    if (recNode) {
      vtodo.removeAllProperties("rrule");
      const rrule = easToRrule(recNode);
      if (rrule && /^FREQ=[A-Z]+/.test(rrule)) {
        const prop = new ICAL.Property("rrule", vtodo);
        prop.setValue(ICAL.Recur.fromString(rrule));
        vtodo.addProperty(prop);
      }
      keepUnmappedRecurrence(vtodo, recNode);
    }
  }

  return vcal.toString();
}

/* ── Writer: iCal VTODO → ApplicationData WBXML ────────────────────── */

export function appendApplicationDataFromIcal({
  builder,
  ical,
  asVersion,
  defaultTimezone,
  syncRecurrence,
}) {
  const vtodo = parseFirstVtodo(ical);
  if (!vtodo) return;

  // Caller hands us the builder on the AirSync codepage; switch into
  // Tasks so the tag tokens resolve.
  builder.switchpage("Tasks");

  builder.atag("Subject", stringOf(vtodo.getFirstPropertyValue("summary")));

  // Body.
  appendBodyFromDescription(builder, vtodo, asVersion, "Tasks");

  // Importance.
  const priority = stringOf(vtodo.getFirstPropertyValue("priority"));
  builder.atag("Importance", PRIORITY_TO_IMPORTANCE[priority] ?? "1");

  // Start / Due (extended-ISO with Z).
  const startProp = vtodo.getFirstProperty("dtstart");
  let localStart = null;
  if (startProp) {
    const utc = toExtendedIsoUtc(startProp.getFirstValue());
    builder.atag("UtcStartDate", utc);
    localStart = fakeLocalAsUtc(startProp.getFirstValue());
    builder.atag("StartDate", localStart);
  }
  const rawDueProp = vtodo.getFirstProperty("due");
  const dueProp = rawDueProp ?? startProp;
  if (dueProp) {
    builder.atag("UtcDueDate", toExtendedIsoUtc(dueProp.getFirstValue()));
    builder.atag("DueDate", fakeLocalAsUtc(dueProp.getFirstValue()));
  }

  // Categories.
  const catsProp = vtodo.getFirstProperty("categories");
  if (catsProp) {
    const cats = catsProp.getValues();
    if (cats.length) {
      builder.otag("Categories");
      for (const c of cats) builder.atag("Category", String(c));
      builder.ctag();
    } else if (asVersion !== "16.1") {
      builder.atag("Categories");
    }
  } else if (asVersion !== "16.1") {
    builder.atag("Categories");
  }

  // Recurrence outbound (RRULE only; tasks have no exceptions). The
  // anchor `<Start>` carries: on ≤14.x the real UTC instant, on 16.x the
  // fake-local form legacy has always sent - see `appendRecurrence`.
  // A task without a DTSTART anchors its series on DUE - Exchange models
  // recurring tasks on the start/due pair, so the due date is a faithful
  // anchor and the rule survives instead of being silently dropped. A
  // rule with NEITHER anchor never gets here: `clientRejectReason` holds
  // the item before the push.
  const anchorProp = startProp ?? rawDueProp;
  if (syncRecurrence && anchorProp) {
    const rrule = vtodo.getFirstProperty("rrule");
    if (rrule) {
      const anchorLocal =
        localStart ?? fakeLocalAsUtc(anchorProp.getFirstValue());
      const anchorUtc = toExtendedIsoUtc(anchorProp.getFirstValue());
      appendRecurrence(
        builder,
        rrule,
        anchorProp,
        anchorLocal,
        anchorUtc,
        asVersion,
        unmappedRecurrenceOf(vtodo),
        firstDayOfWeekOf(vtodo, rrule),
      );
    }
  }

  // Complete.
  const status = stringOf(vtodo.getFirstPropertyValue("status"));
  if (status === "COMPLETED") {
    builder.atag("Complete", "1");
    const completedProp = vtodo.getFirstProperty("completed");
    if (completedProp)
      builder.atag(
        "DateCompleted",
        toExtendedIsoUtc(completedProp.getFirstValue()),
      );
  } else {
    builder.atag("Complete", "0");
  }

  // Sensitivity.
  const cls = stringOf(vtodo.getFirstPropertyValue("class"));
  builder.atag("Sensitivity", CLASS_TO_SENSITIVITY[cls] ?? "0");

  // Reminder.
  const alarm = vtodo.getFirstSubcomponent("valarm");
  if (alarm && (startProp || dueProp)) {
    const reminderTime = absoluteReminderTime(alarm, startProp ?? dueProp);
    if (reminderTime) {
      builder.atag("ReminderTime", reminderTime);
      builder.atag("ReminderSet", "1");
    } else {
      builder.atag("ReminderSet", "0");
    }
  } else {
    builder.atag("ReminderSet", "0");
  }
}

/* ── ID stamping ───────────────────────────────────────────────────── */

/** Names why a task cannot go on the wire without changing its meaning,
 *  or null when it can. Two cases: EAS has no sub-daily recurrence
 *  frequencies, and a recurring task needs an anchor - the series is
 *  emitted against DTSTART or, failing that, DUE. Only when a rule has
 *  neither is the item held. Gated on `syncRecurrence` like the
 *  emission itself. */
export function clientRejectReason({ blob, syncRecurrence }) {
  if (!syncRecurrence || typeof blob !== "string") return null;

  // A moved occurrence, which a task cannot carry. [MS-ASTASK] declares no
  // exception element at any version and this codec neither writes nor
  // reads one, so an override would be dropped in silence and its
  // occurrence would sit back on the rule's own instant rather than the
  // one the item states. An event has somewhere to put it; a task has not.
  //
  // The calendar guard is what produces these: it restates a set of loose
  // dates as a rule with an override moving each occurrence onto its date,
  // and it does that for a VTODO as readily as for a VEVENT.
  const vcal = parseVCalendar(blob);
  if (
    vcal
      ?.getAllSubcomponents("vtodo")
      .some((c) => c.getFirstProperty("recurrence-id"))
  ) {
    return (
      "this task moves individual occurrences, and ActiveSync can state " +
      "a task's occurrences only as a rule"
    );
  }

  if (!blob.includes("RRULE")) return null;
  const vtodo = parseFirstVtodo(blob);
  const rrule = vtodo?.getFirstProperty("rrule");
  if (!rrule) return null;
  const freq = String(rrule.getFirstValue()?.freq ?? "").toUpperCase();
  if (freq === "HOURLY" || freq === "MINUTELY" || freq === "SECONDLY") {
    return `EAS cannot represent a recurrence below daily (FREQ=${freq})`;
  }
  if (!vtodo.getFirstProperty("dtstart") && !vtodo.getFirstProperty("due")) {
    return (
      "a recurring task needs a start or a due date - EAS anchors the " +
      "series on one of them"
    );
  }
  return null;
}

export function readEasServerIdFromIcal(ical) {
  const v = parseFirstVtodo(ical);
  if (!v) return null;
  const x = v.getFirstPropertyValue(X_EAS_SERVERID.toLowerCase());
  return x ? String(x) : null;
}

export function stampEasServerId(ical, serverID) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return ical;
  const vtodo = vcal.getFirstSubcomponent("vtodo");
  if (!vtodo) return ical;
  vtodo.updatePropertyWithValue(X_EAS_SERVERID.toLowerCase(), serverID);
  return vcal.toString();
}

/* ── Date helpers (extended ISO) ───────────────────────────────────── */

/** Add a date-time property to the VTODO. With no `tzid`, emits as
 *  UTC (`...Z`). With a known IANA `tzid`, converts the UTC moment to
 *  wall-clock in that zone via `time.convertToZone(icalTimezone)` and
 *  serialises the property with `TZID=<tzid>` — same shape legacy
 *  produced via `utc.getInTimezone(timezone)` (Lightning's
 *  calITimezoneService is itself ICAL.js-backed). */
function writeUtcDateProp(vtodo, name, easDateStr, tzid) {
  const d = parseExtendedIso(easDateStr);
  if (!d) return;
  // Replace any existing property of this name so merge-mode partial
  // Changes don't duplicate dtstart/due/completed on the master.
  vtodo.removeAllProperties(name);
  const time = new ICAL.Time({
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    hour: d.getUTCHours(),
    minute: d.getUTCMinutes(),
    second: d.getUTCSeconds(),
    isDate: false,
  });
  time.zone = ICAL.Timezone.utcTimezone;
  const prop = new ICAL.Property(name, vtodo);
  const tz = tzid ? getIcalTimezone(tzid) : null;
  if (tz) {
    time.convertToZone(tz);
    prop.setValue(time);
    prop.setParameter("tzid", tzid);
  } else {
    prop.setValue(time);
  }
  vtodo.addProperty(prop);
}

function parseExtendedIso(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Recover the server-hinted IANA tzid from an EAS task date pair
 *  (`<UtcStartDate>` + `<StartDate>` or `<UtcDueDate>` + `<DueDate>`).
 *  Legacy `tasksync.js:53` / `:65` computes `offset = (UtcDate -
 *  LocalDate) / 60000` and passes it to `guessTimezoneByCurrentOffset`
 *  for an IANA-zone lookup. We mirror that calculation literally —
 *  including the JS `new Date(string)` parsing convention where a
 *  trailing `Z` means UTC and its absence means local-clock. */
function guessTzFromPair(utc, local) {
  if (!utc || !local) return null;
  const u = parseExtendedIso(utc);
  const l = parseExtendedIso(local);
  if (!u || !l) return null;
  const offsetMin = Math.round((u.getTime() - l.getTime()) / 60000);
  return guessTimezoneByCurrentOffset(offsetMin, u);
}

function toExtendedIsoUtc(value) {
  const d = value instanceof ICAL.Time ? value.toJSDate() : new Date(value);
  return d.toISOString();
}

/** EAS Tasks "local" date strings: encode local time as if it were UTC.
 *  Mirrors `getIsoUtcString(date, true, true)` from legacy tools.js. */
function fakeLocalAsUtc(value) {
  const d = value instanceof ICAL.Time ? value.toJSDate() : new Date(value);
  const pad = (n) => String(n).padStart(2, "0");
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T` +
    `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}.000Z`
  );
}

/* ── Alarm helpers ─────────────────────────────────────────────────── */

function appendStartRelativeAlarm(vtodo, offsetSeconds) {
  const alarm = new ICAL.Component(["valarm", [], []]);
  alarm.updatePropertyWithValue("action", "DISPLAY");
  const trig = new ICAL.Property("trigger", alarm);
  trig.setValue(ICAL.Duration.fromSeconds(offsetSeconds));
  trig.setParameter("related", "START");
  alarm.addProperty(trig);
  vtodo.addSubcomponent(alarm);
}

function appendAbsoluteAlarm(vtodo, easUtcStr) {
  const d = parseExtendedIso(easUtcStr);
  if (!d) return;
  const alarm = new ICAL.Component(["valarm", [], []]);
  alarm.updatePropertyWithValue("action", "DISPLAY");
  const trig = new ICAL.Property("trigger", alarm);
  const time = new ICAL.Time({
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    hour: d.getUTCHours(),
    minute: d.getUTCMinutes(),
    second: d.getUTCSeconds(),
  });
  time.zone = ICAL.Timezone.utcTimezone;
  trig.setValue(time);
  trig.setParameter("value", "DATE-TIME");
  alarm.addProperty(trig);
  vtodo.addSubcomponent(alarm);
}

function absoluteReminderTime(alarm, anchorProp) {
  const trig = alarm.getFirstProperty("trigger");
  if (!trig) return null;
  const v = trig.getFirstValue();
  if (v instanceof ICAL.Time) return toExtendedIsoUtc(v);
  if (v instanceof ICAL.Duration && anchorProp) {
    const anchor = anchorProp.getFirstValue();
    if (!(anchor instanceof ICAL.Time)) return null;
    const out = new Date(anchor.toJSDate().getTime() + v.toSeconds() * 1000);
    return out.toISOString();
  }
  return null;
}

/* ── Body ──────────────────────────────────────────────────────────── */

/* ── Recurrence (RRULE only; tasks have no exceptions) ────────────── */

/** `[MS-ASTASK]` 2.2.2.31. Every element qualifying the type has to be here:
 *  without them the server rejects the whole Add with Status 6, and only
 *  daily needs none. A rejected entry is re-staged and retried on every
 *  sync, leaving the folder in `warning` until the task is deleted.
 *
 *  `<Start>` is task-only - an event's recurrence is anchored on its own
 *  StartTime - and `Until` uses this codec's own formatter. Everything else is
 *  in the same relative order as the calendar codec, which the server accepts.
 *
 *  **`<Start>` on ≤14.x is a real UTC instant, not the fake-local form used
 *  for `StartDate`/`DueDate`.** Exchange 14.1 derives the stored
 *  `UtcStartDate` of a *recurring* task from this element and reads it as
 *  UTC, so sending local wall-clock numerals moved every recurring task by
 *  the sender's offset: a task at 08:00Z came back at 10:00Z in Europe/Berlin,
 *  and again on each round trip's worth of edits. Measured by making the three
 *  values disagree - `UtcStartDate` 08:00Z, `StartDate` 10:00, `Start` 08:00Z -
 *  and observing the task return at 08:00Z; with `Start` at 10:00 it returned
 *  at 10:00Z. A non-recurring task was never affected, since it sends no
 *  `<Start>` and the server then keeps `UtcStartDate` verbatim.
 *
 *  **16.x keeps the legacy fake-local value, and must.** The same round trip
 *  on an Office 365 16.1 account returns `DTSTART` 08:00Z and `DUE` 09:00Z
 *  untouched for all nine recurrence types, so 16.1 reads `<Start>` the way
 *  legacy always assumed. The two servers genuinely disagree about this
 *  element; sending the real UTC instant to both would move a currently
 *  correct round trip rather than fix a second broken one.
 *
 *  Not emitted: `Regenerate` and `DeadOccur` (we never regenerate a task), and
 *  `CalendarType` (this server takes monthly and yearly *events* without it). */
function appendRecurrence(
  builder,
  rruleProp,
  startProp,
  localStart,
  utcStart,
  asVersion,
  unmapped = [],
  firstDayOfWeek = null,
) {
  const rec = rruleToEas(rruleProp, startProp);
  if (!rec) return;

  const legacy = !String(asVersion ?? "").startsWith("16.");

  builder.otag("Recurrence");
  // Ahead of <Type>, which is where 16.1 puts them in the block it sends
  // us. The Tasks codepage assigns their tokens after the qualifying
  // elements instead, so the two disagree - and the server is the half that
  // was observed. Section 15.3 holds the order, against two different 16.1
  // servers: get it wrong and the push comes back Status 6.
  for (const [tag, value] of unmapped) builder.atag(tag, value);
  builder.atag("Type", String(rec.type));
  builder.atag("Start", legacy && utcStart ? utcStart : localStart);
  if (rec.dayOfMonth !== null) {
    builder.atag("DayOfMonth", String(rec.dayOfMonth));
  }
  if (rec.dayOfWeek !== null) builder.atag("DayOfWeek", String(rec.dayOfWeek));
  builder.atag("Interval", String(rec.interval));
  if (rec.monthOfYear !== null) {
    builder.atag("MonthOfYear", String(rec.monthOfYear));
  }
  if (rec.count) builder.atag("Occurrences", String(rec.count));
  else if (rec.until) builder.atag("Until", toExtendedIsoUtc(rec.until));
  if (rec.weekOfMonth !== null) {
    builder.atag("WeekOfMonth", String(rec.weekOfMonth));
  }
  // Last in the block, which is where both servers put it - unlike the
  // elements above, which they send first. Taken from what the server last
  // said rather than from the rule's own WKST: Thunderbird stamps that from
  // a preference on any edit, and has no control for it.
  if (firstDayOfWeek != null) {
    builder.atag("FirstDayOfWeek", String(firstDayOfWeek));
  }
  builder.ctag();
}

/* ── ICAL.js plumbing ──────────────────────────────────────────────── */

function newVCalendar() {
  const vcal = new ICAL.Component(["vcalendar", [], []]);
  vcal.updatePropertyWithValue("prodid", "-//tbsync-eas//EN");
  vcal.updatePropertyWithValue("version", "2.0");
  return vcal;
}

function parseVCalendar(ical) {
  if (!ical) return null;
  try {
    return new ICAL.Component(ICAL.parse(ical));
  } catch {
    return null;
  }
}

/** A date element, with the Windows FILETIME zero read as absent - see
 *  `isFiletimeZero`. Every date this codec reads goes through here. */
function readDate(adNode, tag) {
  const v = readPathFrom(adNode, [tag]);
  return v && !isFiletimeZero(v) ? v : null;
}

function parseFirstVtodo(ical) {
  const vcal = parseVCalendar(ical);
  return vcal?.getFirstSubcomponent("vtodo") ?? null;
}

function stringOf(v) {
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (typeof v.toString === "function") return v.toString();
  return String(v);
}

function childByTag(node, tag) {
  if (!node?.children) return null;
  for (const c of node.children) if (c.tagName === tag) return c;
  return null;
}

function collectChildren(adNode, wrapperTag, childTag) {
  const wrapper = childByTag(adNode, wrapperTag);
  if (!wrapper) return [];
  const out = [];
  for (const c of wrapper.children) {
    if (c.tagName === childTag) {
      const t = c.textContent;
      if (t != null) {
        try {
          out.push(decodeURIComponent(t));
        } catch {
          out.push(t);
        }
      }
    }
  }
  return out;
}
