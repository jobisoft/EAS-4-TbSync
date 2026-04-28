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
import { readPathFrom } from "./wbxml-helpers.mjs";

const X_EAS_SERVERID = "X-EAS-SERVERID";

const IMPORTANCE_TO_PRIORITY = { "0": "9", "1": "5", "2": "1" };
const PRIORITY_TO_IMPORTANCE = { "9": "0", "5": "1", "1": "2" };

const SENSITIVITY_TO_CLASS = { "0": "PUBLIC", "1": "PRIVATE", "2": "PRIVATE", "3": "CONFIDENTIAL" };
const CLASS_TO_SENSITIVITY = { PUBLIC: "0", PRIVATE: "2", CONFIDENTIAL: "3" };

/* ── Reader: ApplicationData → iCal VTODO ──────────────────────────── */

export function applicationDataToIcal({ adNode, serverID, asVersion, defaultTimezone, syncRecurrence, msTodoCompat, uid }) {
  const vcal = newVCalendar();
  const vtodo = new ICAL.Component(["vtodo", [], []]);
  vcal.addSubcomponent(vtodo);
  if (uid) vtodo.updatePropertyWithValue("uid", uid);
  vtodo.updatePropertyWithValue(X_EAS_SERVERID.toLowerCase(), serverID);

  const subject = readPathFrom(adNode, ["Subject"]);
  if (subject) vtodo.updatePropertyWithValue("summary", subject);

  // Body (codepage-aware).
  if (asVersion === "2.5") {
    const data = readPathFrom(adNode, ["Body"]);
    if (data) vtodo.updatePropertyWithValue("description", data);
  } else {
    const data = readPathFrom(adNode, ["Body", "Data"]);
    if (data) vtodo.updatePropertyWithValue("description", data);
  }

  // Reminder is read up-front so the MS To-Do compatibility hack can pin
  // DTSTART/DUE to the reminder time before they are written.
  const reminderTime = readPathFrom(adNode, ["ReminderSet"]) === "1"
    ? readPathFrom(adNode, ["ReminderTime"])
    : null;
  const msTodoOverride = msTodoCompat === true && !!reminderTime;

  // StartDate / DueDate with Utc* pairing for offset extraction.
  if (msTodoOverride) {
    // MS To-Do only ships date-only due dates; pin both ends to the
    // reminder so Lightning renders the task on the correct day.
    writeUtcDateProp(vtodo, "dtstart", reminderTime);
    writeUtcDateProp(vtodo, "due",     reminderTime);
  } else {
    const startUtc = readPathFrom(adNode, ["UtcStartDate"]);
    if (startUtc) writeUtcDateProp(vtodo, "dtstart", startUtc);
    const dueUtc = readPathFrom(adNode, ["UtcDueDate"]);
    if (dueUtc) writeUtcDateProp(vtodo, "due", dueUtc);
  }

  // Importance → PRIORITY.
  const importance = readPathFrom(adNode, ["Importance"]);
  if (importance && IMPORTANCE_TO_PRIORITY[importance]) {
    vtodo.updatePropertyWithValue("priority", IMPORTANCE_TO_PRIORITY[importance]);
  }
  // Sensitivity → CLASS.
  const sens = readPathFrom(adNode, ["Sensitivity"]);
  if (sens && SENSITIVITY_TO_CLASS[sens]) {
    vtodo.updatePropertyWithValue("class", SENSITIVITY_TO_CLASS[sens]);
  }

  // Complete + DateCompleted.
  const complete = readPathFrom(adNode, ["Complete"]);
  if (complete === "1") {
    vtodo.updatePropertyWithValue("status", "COMPLETED");
    vtodo.updatePropertyWithValue("percent-complete", 100);
    const dc = readPathFrom(adNode, ["DateCompleted"]);
    if (dc) writeUtcDateProp(vtodo, "completed", dc);
  }

  // Reminder.
  if (reminderTime) {
    if (msTodoOverride) appendStartRelativeAlarm(vtodo, 0);
    else                appendAbsoluteAlarm(vtodo, reminderTime);
  }

  // Categories.
  const cats = collectChildren(adNode, "Categories", "Category");
  if (cats.length) {
    const prop = new ICAL.Property("categories", vtodo);
    prop.setValues(cats);
    vtodo.addProperty(prop);
  }

  // Recurrence (RRULE only; tasks have no exceptions in EAS).
  if (syncRecurrence) {
    const recNode = childByTag(adNode, "Recurrence");
    if (recNode) {
      const rrule = recurrenceToRrule(recNode);
      if (rrule && /^FREQ=[A-Z]+/.test(rrule)) {
        const prop = new ICAL.Property("rrule", vtodo);
        prop.setValue(ICAL.Recur.fromString(rrule));
        vtodo.addProperty(prop);
      }
    }
  }

  return vcal.toString();
}

/* ── Writer: iCal VTODO → ApplicationData WBXML ────────────────────── */

export function appendApplicationDataFromIcal({ builder, ical, asVersion, defaultTimezone, syncRecurrence }) {
  const vtodo = parseFirstVtodo(ical);
  if (!vtodo) return;

  // Caller hands us the builder on the AirSync codepage; switch into
  // Tasks so the tag tokens resolve.
  builder.switchpage("Tasks");

  builder.atag("Subject", stringOf(vtodo.getFirstPropertyValue("summary")));

  // Body.
  appendBody(builder, vtodo, asVersion);

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
  const dueProp = vtodo.getFirstProperty("due") ?? startProp;
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

  // Recurrence outbound (RRULE only; tasks have no exceptions). Need a
  // localStart to put inside <Start> per legacy.
  if (syncRecurrence && localStart) {
    const rrule = vtodo.getFirstProperty("rrule");
    if (rrule) appendRecurrence(builder, rrule, startProp, localStart);
  }

  // Complete.
  const status = stringOf(vtodo.getFirstPropertyValue("status"));
  if (status === "COMPLETED") {
    builder.atag("Complete", "1");
    const completedProp = vtodo.getFirstProperty("completed");
    if (completedProp) builder.atag("DateCompleted", toExtendedIsoUtc(completedProp.getFirstValue()));
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

function writeUtcDateProp(vtodo, name, easDateStr) {
  const d = parseExtendedIso(easDateStr);
  if (!d) return;
  const time = new ICAL.Time({
    year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(),
    hour: d.getUTCHours(), minute: d.getUTCMinutes(), second: d.getUTCSeconds(),
    isDate: false,
  });
  time.zone = ICAL.Timezone.utcTimezone;
  const prop = new ICAL.Property(name, vtodo);
  prop.setValue(time);
  vtodo.addProperty(prop);
}

function parseExtendedIso(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function toExtendedIsoUtc(value) {
  const d = (value instanceof ICAL.Time) ? value.toJSDate() : new Date(value);
  return d.toISOString();
}

/** EAS Tasks "local" date strings: encode local time as if it were UTC.
 *  Mirrors `getIsoUtcString(date, true, true)` from legacy tools.js. */
function fakeLocalAsUtc(value) {
  const d = (value instanceof ICAL.Time) ? value.toJSDate() : new Date(value);
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T` +
         `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}.000Z`;
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
    year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(),
    hour: d.getUTCHours(), minute: d.getUTCMinutes(), second: d.getUTCSeconds(),
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

function appendBody(builder, vtodo, asVersion) {
  const desc = stringOf(vtodo.getFirstPropertyValue("description"));
  if (asVersion === "16.1" && !desc) return;
  if (asVersion === "2.5") {
    builder.atag("Body", desc ?? "");
    return;
  }
  builder.switchpage("AirSyncBase");
  builder.otag("Body");
    builder.atag("Type", "1");
    if (asVersion !== "16.1") builder.atag("EstimatedDataSize", String((desc ?? "").length));
    builder.atag("Data", desc ?? "");
  builder.ctag();
  builder.switchpage("Tasks");
}

/* ── Recurrence (RRULE only; tasks have no exceptions) ────────────── */

function recurrenceToRrule(recNode) {
  const type = readPathFrom(recNode, ["Type"]);
  const freq = ({ "0": "DAILY", "1": "WEEKLY", "2": "MONTHLY", "3": "MONTHLY",
                  "5": "YEARLY", "6": "YEARLY" })[type];
  if (!freq) return null;
  const parts = [`FREQ=${freq}`];
  const interval = readPathFrom(recNode, ["Interval"]);
  if (interval) parts.push(`INTERVAL=${interval}`);
  const dow = readPathFrom(recNode, ["DayOfWeek"]);
  if (dow) {
    const bits = parseInt(dow, 10) || 0;
    const week = readPathFrom(recNode, ["WeekOfMonth"]);
    const ical = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
    const days = [];
    for (let i = 0; i < 7; i++) if (bits & (1 << i)) days.push(ical[i]);
    if (days.length) {
      const prefix = week === "5" ? "-1" : (week ? String(week) : "");
      parts.push("BYDAY=" + days.map(d => prefix + d).join(","));
    }
  }
  const dom = readPathFrom(recNode, ["DayOfMonth"]);
  if (dom) parts.push(`BYMONTHDAY=${dom}`);
  const moy = readPathFrom(recNode, ["MonthOfYear"]);
  if (moy) parts.push(`BYMONTH=${moy}`);
  const occ = readPathFrom(recNode, ["Occurrences"]);
  if (occ) parts.push(`COUNT=${occ}`);
  const until = readPathFrom(recNode, ["Until"]);
  if (until) parts.push(`UNTIL=${until.replace(/[-:]/g, "")}`);
  return parts.join(";");
}

function appendRecurrence(builder, rruleProp, startProp, localStart) {
  const r = rruleProp.getFirstValue();
  if (!r) return;
  let type = 0;
  if      (r.freq === "DAILY")   type = 0;
  else if (r.freq === "WEEKLY")  type = 1;
  else if (r.freq === "MONTHLY") type = 2;
  else if (r.freq === "YEARLY")  type = 5;
  builder.otag("Recurrence");
    builder.atag("Type", String(type));
    builder.atag("Start", localStart);
    builder.atag("Interval", String(r.interval ?? 1));
    if (r.count) builder.atag("Occurrences", String(r.count));
    else if (r.until) builder.atag("Until", toExtendedIsoUtc(r.until));
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
  try { return new ICAL.Component(ICAL.parse(ical)); }
  catch { return null; }
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
        try { out.push(decodeURIComponent(t)); }
        catch { out.push(t); }
      }
    }
  }
  return out;
}
