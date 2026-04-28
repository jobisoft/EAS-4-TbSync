/**
 * EAS Calendar (codepage 4) + AirSyncBase Body (17) ⇆ iCal VEVENT codec.
 *
 * Mirrors the legacy `EAS-4-TbSync/content/includes/calendarsync.js` field
 * map. Round-trips the common field set: Subject, Location, Body,
 * Start/End/AllDay, BusyStatus, Sensitivity, Reminder, Categories,
 * Organizer, Attendees, MeetingStatus, ResponseType, UID, recurrence.
 *
 * Recurrence handling is one-shot RRULE only (no embedded `<Exceptions>`
 * yet - inbound exceptions and outbound RECURRENCE-ID/EXDATE deltas are a
 * follow-up). The 16.1 InstanceId-per-Change exception path is not
 * emitted; on the inbound side we ignore InstanceId-bearing changes for
 * now (the master event is still kept in sync).
 *
 * The TimeZone blob (≤14.x only; 16.1 uses UTC times) is encoded /
 * decoded via `TimeZoneBlob` in `timezone-blob.mjs`. When the server's
 * blob is all-zero we fall back to the host's default IANA zone.
 *
 * Local items carry the EAS server-assigned ServerId on a custom
 * `X-EAS-SERVERID` property so pull/push paths can find the local item
 * without a separate id map (mirrors the contact-codec's approach).
 */

import ICAL from "../../vendor/ical.min.js";
import { readPathFrom } from "./wbxml-helpers.mjs";
import { TimeZoneBlob, isAllZero } from "./timezone-blob.mjs";

const X_EAS_SERVERID    = "X-EAS-SERVERID";
const X_EAS_RESPONSETYPE = "X-EAS-RESPONSETYPE";
const X_EAS_MEETINGSTATUS = "X-EAS-MEETINGSTATUS";

// EAS BusyStatus → iCal TRANSP. Tentative (1) maps to "no TRANSP" so the
// caller falls back to STATUS=TENTATIVE; the codec mirrors legacy here.
const BUSYSTATUS_TO_TRANSP = {
  "0": "TRANSPARENT", "1": null, "2": "OPAQUE", "3": "OPAQUE", "4": "OPAQUE",
};
const TRANSP_TO_BUSYSTATUS = { TRANSPARENT: "0", OPAQUE: "2" };

// EAS Sensitivity → iCal CLASS.
const SENSITIVITY_TO_CLASS = { "0": "PUBLIC", "1": "PRIVATE", "2": "PRIVATE", "3": "CONFIDENTIAL" };
const CLASS_TO_SENSITIVITY = { PUBLIC: "0", PRIVATE: "2", CONFIDENTIAL: "3" };

// EAS AttendeeStatus → iCal PARTSTAT.
const ATTENDEESTATUS_TO_PARTSTAT = {
  "0": "NEEDS-ACTION", "2": "TENTATIVE", "3": "ACCEPTED", "4": "DECLINED", "5": "ACCEPTED",
};

/* ── Reader: ApplicationData → iCal VEVENT ─────────────────────────── */

export function applicationDataToIcal({ adNode, serverID, asVersion, defaultTimezone, syncRecurrence, uid }) {
  const vcal = newVCalendar();
  const vevent = new ICAL.Component(["vevent", [], []]);
  vcal.addSubcomponent(vevent);

  if (uid) vevent.updatePropertyWithValue("uid", uid);
  vevent.updatePropertyWithValue(X_EAS_SERVERID.toLowerCase(), serverID);

  populateVeventFromAd({ adNode, vevent, asVersion, defaultTimezone });

  // Recurrence + 2.5/14.x exceptions. Gated on the account-level
  // syncRecurrence flag.
  if (syncRecurrence) {
    const recNode = childByTag(adNode, "Recurrence");
    if (recNode) {
      const rrule = recurrenceToRrule(recNode);
      if (rrule && /^FREQ=[A-Z]+/.test(rrule)) {
        const prop = new ICAL.Property("rrule", vevent);
        prop.setValue(ICAL.Recur.fromString(rrule));
        vevent.addProperty(prop);
      }
    }
    appendInboundExceptions({ adNode, vcal, vevent, asVersion, defaultTimezone });
  }

  return vcal.toString();
}

/** Populate a VEVENT (master or override) from an EAS <ApplicationData>
 *  or <Exception> node. The set of fields is the same on both — legacy
 *  reuses `setThunderbirdItemFromWbxml` for both paths.
 *  Returns nothing; mutates `vevent`. */
function populateVeventFromAd({ adNode, vevent, asVersion, defaultTimezone }) {
  // Subject / Location.
  const subject = readPathFrom(adNode, ["Subject"]);
  if (subject) vevent.updatePropertyWithValue("summary", subject);

  const locDisplay = readPathFrom(adNode, ["Location", "DisplayName"])
                  ?? readPathFrom(adNode, ["Location"]);
  if (locDisplay) vevent.updatePropertyWithValue("location", locDisplay);

  // Body (codepage-aware; AirSyncBase ≥14.x).
  if (useAirSyncBaseBody(asVersion)) {
    const data = readPathFrom(adNode, ["Body", "Data"]);
    if (data) vevent.updatePropertyWithValue("description", data);
  } else {
    const data = readPathFrom(adNode, ["Body"]);
    if (data) vevent.updatePropertyWithValue("description", data);
  }

  // Resolve effective timezone (sticks to UTC on 16.1; otherwise use the
  // TimeZone blob to derive an IANA zone via std/dst offset matching, or
  // fall back to the host's default zone).
  const tzId = resolveTimezone(adNode, asVersion, defaultTimezone);
  const allDay = readPathFrom(adNode, ["AllDayEvent"]) === "1";

  // Start / End. EAS sends UTC strings; convert on the way in.
  const startUtc = readPathFrom(adNode, ["StartTime"]);
  const endUtc   = readPathFrom(adNode, ["EndTime"]);
  if (startUtc) writeDateProp(vevent, "dtstart", startUtc, tzId, allDay);
  if (endUtc)   writeDateProp(vevent, "dtend",   endUtc,   tzId, allDay);

  // DtStamp - preserve when present (AS ≤ 14.x); 16.1 omits.
  const dtStamp = readPathFrom(adNode, ["DtStamp"]);
  if (dtStamp) writeDateProp(vevent, "dtstamp", dtStamp, "UTC", false);

  // BusyStatus + STATUS interplay (legacy logic, see calendarsync.js).
  const busy = readPathFrom(adNode, ["BusyStatus"]);
  const transp = busy ? BUSYSTATUS_TO_TRANSP[busy] : undefined;
  if (transp) vevent.updatePropertyWithValue("transp", transp);

  // Sensitivity → CLASS.
  const sens = readPathFrom(adNode, ["Sensitivity"]);
  if (sens && SENSITIVITY_TO_CLASS[sens]) {
    vevent.updatePropertyWithValue("class", SENSITIVITY_TO_CLASS[sens]);
  }

  // Reminder → VALARM (DISPLAY, offset relative to start in minutes).
  const reminderMinutes = readPathFrom(adNode, ["Reminder"]);
  if (reminderMinutes != null && reminderMinutes !== "" && startUtc) {
    appendDisplayAlarm(vevent, parseInt(reminderMinutes, 10));
  }

  // Categories.
  const cats = collectChildren(adNode, "Categories", "Category");
  if (cats.length) {
    const prop = new ICAL.Property("categories", vevent);
    prop.setValues(cats);
    vevent.addProperty(prop);
  }

  // Organizer (omitted by server in 16.1).
  const orgEmail = readPathFrom(adNode, ["OrganizerEmail"]);
  const orgName  = readPathFrom(adNode, ["OrganizerName"]);
  if (orgEmail) {
    const prop = new ICAL.Property("organizer", vevent);
    prop.setValue("mailto:" + orgEmail);
    if (orgName) prop.setParameter("cn", orgName);
    vevent.addProperty(prop);
  }

  // Attendees.
  const attendees = collectAttendees(adNode);
  for (const a of attendees) {
    const prop = new ICAL.Property("attendee", vevent);
    prop.setValue("mailto:" + a.email);
    if (a.cn) prop.setParameter("cn", a.cn);
    if (a.role) prop.setParameter("role", a.role);
    if (a.partstat) prop.setParameter("partstat", a.partstat);
    if (a.cutype) prop.setParameter("cutype", a.cutype);
    vevent.addProperty(prop);
  }

  // Opaque pass-throughs.
  const respType = readPathFrom(adNode, ["ResponseType"]);
  if (respType) vevent.updatePropertyWithValue(X_EAS_RESPONSETYPE.toLowerCase(), respType);
  const meetingStatus = readPathFrom(adNode, ["MeetingStatus"]);
  if (meetingStatus) {
    vevent.updatePropertyWithValue(X_EAS_MEETINGSTATUS.toLowerCase(), meetingStatus);
    // Map MeetingStatus to STATUS (CONFIRMED / CANCELLED).
    const ms = parseInt(meetingStatus, 10) || 0;
    if (ms & 0x4) vevent.updatePropertyWithValue("status", "CANCELLED");
    else if (ms & 0x1) vevent.updatePropertyWithValue("status", "CONFIRMED");
  } else if (busy === "1") {
    // Tentative-only state: leave TRANSP unset and set STATUS=TENTATIVE.
    vevent.updatePropertyWithValue("status", "TENTATIVE");
  }
}

/** Public entry point for the 16.1 InstanceId path: called from the
 *  sync runner when an inbound `<Change>` carries an `<InstanceId>`.
 *  Locates or creates the override VEVENT keyed by RECURRENCE-ID, then
 *  populates it from `adNode`. For deletions, the runner adds an EXDATE
 *  via `addExdateToMaster` instead. */
export function applyInstanceChange({ ical, adNode, instanceUtc, asVersion, defaultTimezone }) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return ical;
  const master = vcal.getFirstSubcomponent("vevent");
  if (!master) return ical;

  removeExdate(master, instanceUtc);

  // Drop any existing override for this RECURRENCE-ID before re-creating.
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    const rid = sub.getFirstPropertyValue("recurrence-id");
    if (rid && rid.toString() === instanceUtcToIcalString(instanceUtc)) {
      vcal.removeSubcomponent(sub);
    }
  }

  const override = new ICAL.Component(["vevent", [], []]);
  vcal.addSubcomponent(override);
  const masterUid = stringOf(master.getFirstPropertyValue("uid"));
  if (masterUid) override.updatePropertyWithValue("uid", masterUid);
  // RECURRENCE-ID anchors the override to the original master occurrence.
  const ridProp = new ICAL.Property("recurrence-id", override);
  ridProp.setValue(instanceUtcToIcalTime(instanceUtc));
  override.addProperty(ridProp);
  populateVeventFromAd({ adNode: adNode, vevent: override, asVersion, defaultTimezone });
  return vcal.toString();
}

/** Outbound 16.1: emit one `<Change ServerId=master>` per current
 *  EXDATE / RECURRENCE-ID override on the master. Idempotent — re-asserts
 *  the full exception set on every push of a recurring master.
 *
 *  Limitation: a user un-deleting an EXDATE or removing an override
 *  cannot be expressed in EAS without comparing against the server's
 *  last-known state (which we don't currently snapshot). The unwanted
 *  EXDATE / override stays server-side until manually re-edited there.
 *
 *  Caller (sync runner) hands us the builder on the AirSync codepage
 *  after closing the master `<Change>`. We emit zero or more sibling
 *  `<Change>` commands and leave the builder on AirSync.
 */
export function appendInstanceChanges({ builder, blob, serverID, asVersion, defaultTimezone, syncRecurrence }) {
  if (asVersion !== "16.1") return;
  const vcal = parseVCalendar(blob);
  if (!vcal) return;
  const master = vcal.getFirstSubcomponent("vevent")
              ?? null;
  // parseVCalendar's first vevent may be an override if iCal order is
  // unusual; reuse the master picker instead.
  const masterVevent = pickMasterVevent(vcal) ?? master;
  if (!masterVevent) return;

  const masterUid = stringOf(masterVevent.getFirstPropertyValue("uid"));
  const exdates = collectExdates(masterVevent);
  const overrides = [];
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    if (sub === masterVevent) continue;
    const subUid = stringOf(sub.getFirstPropertyValue("uid"));
    const rid = sub.getFirstProperty("recurrence-id");
    if (subUid === masterUid && rid) overrides.push(sub);
  }
  if (!exdates.length && !overrides.length) return;

  for (const ex of exdates) {
    builder.otag("Change");
      builder.atag("ServerId", serverID);
      builder.otag("ApplicationData");
        builder.switchpage("AirSyncBase");
        builder.atag("InstanceId", icalTimeToBasicUtc(ex));
        builder.switchpage("Calendar");
        builder.atag("Deleted", "1");
        builder.switchpage("AirSync");
      builder.ctag();
    builder.ctag();
  }
  for (const override of overrides) {
    const rid = override.getFirstPropertyValue("recurrence-id");
    builder.otag("Change");
      builder.atag("ServerId", serverID);
      builder.otag("ApplicationData");
        builder.switchpage("AirSyncBase");
        builder.atag("InstanceId", icalTimeToBasicUtc(rid));
        // appendApplicationDataFromIcal switches to Calendar at entry
        // and may bounce to AirSyncBase for Body / Location, but always
        // returns to Calendar before the closing tag.
        appendApplicationDataFromIcal({
          builder, ical: override, asVersion, defaultTimezone, syncRecurrence,
          isException: true,
        });
        builder.switchpage("AirSync");
      builder.ctag();
    builder.ctag();
  }
}

function pickMasterVevent(vcal) {
  const all = vcal.getAllSubcomponents("vevent");
  for (const v of all) if (!v.getFirstProperty("recurrence-id")) return v;
  return all[0] ?? null;
}

/** Add an EXDATE to the master VEVENT (16.1 InstanceId-with-Deleted=1). */
export function applyInstanceDelete({ ical, instanceUtc }) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return ical;
  const master = vcal.getFirstSubcomponent("vevent");
  if (!master) return ical;

  // Drop any existing override at this RECURRENCE-ID — server says it's
  // gone now.
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    const rid = sub.getFirstPropertyValue("recurrence-id");
    if (rid && rid.toString() === instanceUtcToIcalString(instanceUtc)) {
      vcal.removeSubcomponent(sub);
    }
  }
  addExdate(master, instanceUtc);
  return vcal.toString();
}

/* ── Writer: iCal VEVENT → ApplicationData WBXML ───────────────────── */

export function appendApplicationDataFromIcal({
  builder, ical, asVersion, defaultTimezone, syncRecurrence, isException = false,
}) {
  // Exception bodies are emitted from a vevent we've already parsed;
  // accept either a string or a pre-parsed component for nested calls.
  let vevent = ical;
  if (typeof ical === "string") vevent = parseFirstVevent(ical);
  else if (ical && ical.name !== "vevent") vevent = null;
  if (!vevent) return;
  const vcal = vevent.parent;

  // Caller hands us the builder on the AirSync codepage; switch into
  // Calendar so the tag tokens resolve. The body / location / inner
  // helpers switch to AirSyncBase as needed and switch back here.
  builder.switchpage("Calendar");

  // Outbound timezone (≤14.x only; never inside an exception body —
  // legacy emits this only on the master).
  if (asVersion !== "16.1" && !isException) {
    const blob = buildTimezoneBlob(vevent, defaultTimezone);
    builder.atag("TimeZone", blob.easTimeZone64);
  }

  const dtstart = vevent.getFirstProperty("dtstart");
  const dtend   = vevent.getFirstProperty("dtend");
  const allDay  = isAllDayProp(dtstart) && isAllDayProp(dtend);
  builder.atag("AllDayEvent", allDay ? "1" : "0");

  // Body.
  appendBody(builder, vevent, asVersion);

  // BusyStatus from TRANSP (or TENTATIVE STATUS).
  const status = stringOf(vevent.getFirstPropertyValue("status"));
  if (status === "TENTATIVE") {
    builder.atag("BusyStatus", "1");
  } else {
    const transp = stringOf(vevent.getFirstPropertyValue("transp"));
    builder.atag("BusyStatus", TRANSP_TO_BUSYSTATUS[transp] ?? "2");
  }

  // Organizer (≤14.x; not inside an exception).
  if (asVersion !== "16.1" && !isException) {
    const orgProp = vevent.getFirstProperty("organizer");
    if (orgProp) {
      const cn = orgProp.getParameter("cn");
      if (cn) builder.atag("OrganizerName", cn);
      const email = stripMailto(orgProp.getFirstValue());
      if (email) builder.atag("OrganizerEmail", email);
    }
  }

  // DtStamp (≤14.x).
  if (asVersion !== "16.1") {
    const ds = vevent.getFirstProperty("dtstamp");
    builder.atag("DtStamp", ds ? toBasicUtc(ds.getFirstValue()) : nowBasicUtc());
  }

  // EndTime. AS 16.1 all-day uses a "fake local as UTC" form
  // (`YYYYMMDDT000000Z` from the local-clock date, no TZ conversion) -
  // mirrors legacy `getIsoUtcString(date, false, true, true)` so the
  // user-intended date isn't shifted by ±1 day in non-UTC zones.
  builder.atag("EndTime", endTimeFor(dtend, asVersion, allDay));

  // Location.
  const location = stringOf(vevent.getFirstPropertyValue("location"));
  if (asVersion !== "16.1") {
    builder.atag("Location", location);
  } else if (location) {
    builder.switchpage("AirSyncBase");
    builder.otag("Location");
      builder.atag("DisplayName", location);
    builder.ctag();
    builder.switchpage("Calendar");
  }

  // Reminder.
  const alarm = vevent.getFirstSubcomponent("valarm");
  if (alarm) {
    const minutes = alarmMinutes(alarm, dtstart);
    if (minutes != null && minutes >= 0) builder.atag("Reminder", String(minutes));
  }

  // Sensitivity.
  const cls = stringOf(vevent.getFirstPropertyValue("class"));
  builder.atag("Sensitivity", CLASS_TO_SENSITIVITY[cls] ?? "0");

  // Subject + StartTime.
  builder.atag("Subject", stringOf(vevent.getFirstPropertyValue("summary")));
  builder.atag("StartTime", startTimeFor(dtstart, asVersion, allDay));

  // UID (forbidden in 16.1; not inside exceptions either - legacy
  // suppresses UID inside <Exception>, even on 2.5/14.x).
  if (asVersion !== "16.1" && !isException) {
    const uid = stringOf(vevent.getFirstPropertyValue("uid"));
    if (uid) builder.atag("UID", uid);
  }

  // MeetingStatus + Attendees. Legacy comment: Exchange 2010 doesn't
  // support MeetingStatus inside <Exception>, so skip both fields when
  // emitting an exception body.
  if (!isException) {
    const attendees = vevent.getAllProperties("attendee");
    if (attendees.length === 0) {
      builder.atag("MeetingStatus", "0");
    } else {
      const cancelled = status === "CANCELLED";
      const orgProp = vevent.getFirstProperty("organizer");
      const isReceived = orgProp ? !!stripMailto(orgProp.getFirstValue()) &&
                                    !ownerMatchesOrganizer(orgProp) : false;
      if (cancelled) builder.atag("MeetingStatus", isReceived ? "7" : "5");
      else            builder.atag("MeetingStatus", isReceived ? "3" : "1");

      builder.otag("Attendees");
      for (const a of attendees) {
        builder.otag("Attendee");
          builder.atag("Email", stripMailto(a.getFirstValue()));
          const cn = a.getParameter("cn") ?? stripMailto(a.getFirstValue()).split("@")[0];
          builder.atag("Name", cn);
          if (asVersion !== "2.5") {
            const role = a.getParameter("role");
            const cutype = a.getParameter("cutype");
            let type = "2";
            if (cutype === "RESOURCE" || cutype === "ROOM" || role === "NON-PARTICIPANT") type = "3";
            else if (role === "REQ-PARTICIPANT" || role === "CHAIR") type = "1";
            builder.atag("AttendeeType", type);
          }
        builder.ctag();
      }
      builder.ctag();
    }
  }

  // Categories. Legacy emits Categories on the master only.
  if (!isException) {
    const catsProp = vevent.getFirstProperty("categories");
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
  }

  // Recurrence + 2.5/14.x <Exceptions> block. Master only; gated on
  // syncRecurrence. 16.1 sends exceptions as separate <Change> commands
  // at the orchestrator level, so the master payload itself never
  // carries an <Exceptions> wrapper on 16.1.
  if (syncRecurrence && !isException) {
    const rrule = vevent.getFirstProperty("rrule");
    if (rrule) appendRecurrence(builder, rrule, dtstart);
    if (asVersion !== "16.1") {
      appendOutboundExceptions({
        builder, vcal, vevent, asVersion, defaultTimezone, syncRecurrence,
      });
    }
  }
}

/* ── ID stamping ───────────────────────────────────────────────────── */

export function readEasServerIdFromIcal(ical) {
  const v = parseFirstVevent(ical);
  if (!v) return null;
  const x = v.getFirstPropertyValue(X_EAS_SERVERID.toLowerCase());
  return x ? String(x) : null;
}

export function stampEasServerId(ical, serverID) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return ical;
  const vevent = vcal.getFirstSubcomponent("vevent");
  if (!vevent) return ical;
  vevent.updatePropertyWithValue(X_EAS_SERVERID.toLowerCase(), serverID);
  return vcal.toString();
}

/* ── Helpers: timezone resolution ──────────────────────────────────── */

function resolveTimezone(adNode, asVersion, defaultTimezone) {
  if (asVersion === "16.1") return "UTC";
  const blobB64 = readPathFrom(adNode, ["TimeZone"]);
  if (!blobB64 || isAllZero(blobB64)) return defaultTimezone || "UTC";
  // The blob's standardName is a Windows zone ID. Without a Windows→IANA
  // map we keep UTC times in the iCal value but skip TZID; callers that
  // need exact local times can layer a map later.
  return defaultTimezone || "UTC";
}

function buildTimezoneBlob(vevent, defaultTimezone) {
  const blob = new TimeZoneBlob();
  // Without a working IANA→Windows offset/switchdate computation here,
  // emit a zero-bias blob and rely on UTC times in StartTime/EndTime.
  // Servers tolerate this and treat the times as authoritative.
  blob.utcOffset = 0;
  blob.standardBias = 0;
  blob.daylightBias = 0;
  blob.standardName = "UTC";
  blob.daylightName = "UTC";
  return blob;
}

/* ── Helpers: dates ────────────────────────────────────────────────── */

function writeDateProp(vevent, name, easUtc, tzId, allDay) {
  const prop = new ICAL.Property(name, vevent);
  if (allDay) {
    // EAS UTC → date-only (drop time).
    const d = parseEasUtc(easUtc);
    if (!d) return;
    const date = new ICAL.Time({
      year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(),
      isDate: true,
    });
    prop.setValue(date);
  } else {
    const d = parseEasUtc(easUtc);
    if (!d) return;
    const time = new ICAL.Time({
      year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(),
      hour: d.getUTCHours(), minute: d.getUTCMinutes(), second: d.getUTCSeconds(),
      isDate: false,
    });
    if (tzId === "UTC") time.zone = ICAL.Timezone.utcTimezone;
    prop.setValue(time);
    if (tzId && tzId !== "UTC") prop.setParameter("tzid", tzId);
  }
  vevent.addProperty(prop);
}

function parseEasUtc(s) {
  if (!s) return null;
  // Accept extended ISO and basic compact forms.
  const compact = s.replace(/[-:]/g, "");
  const m = /^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?Z?$/.exec(compact);
  if (!m) return null;
  const [, y, mo, d, h = "0", mi = "0", se = "0"] = m;
  return new Date(Date.UTC(+y, +mo - 1, +d, +h, +mi, +se));
}

function toBasicUtc(value) {
  if (!value) return nowBasicUtc();
  const d = (value instanceof ICAL.Time) ? value.toJSDate() : new Date(value);
  return formatBasicUtc(d);
}

function formatBasicUtc(d) {
  const pad = n => String(n).padStart(2, "0");
  return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}` +
         `T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

function nowBasicUtc() { return formatBasicUtc(new Date()); }

/** Read a property's date as `YYYYMMDDT000000Z` from the *local-clock*
 *  year/month/day, with no UTC conversion. Mirrors legacy
 *  `getIsoUtcString(date, false, true, true)` for AS 16.1 all-day. */
function fakeLocalAsUtcDate(prop) {
  if (!prop) return nowBasicUtc();
  const v = prop.getFirstValue();
  const pad = n => String(n).padStart(2, "0");
  if (v instanceof ICAL.Time) {
    return `${v.year}${pad(v.month)}${pad(v.day)}T000000Z`;
  }
  const d = new Date(v);
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}T000000Z`;
}

function startTimeFor(dtstart, asVersion, allDay) {
  if (asVersion === "16.1" && allDay) return fakeLocalAsUtcDate(dtstart);
  return dtstart ? toBasicUtc(dtstart.getFirstValue()) : nowBasicUtc();
}

function endTimeFor(dtend, asVersion, allDay) {
  if (asVersion === "16.1" && allDay) return fakeLocalAsUtcDate(dtend);
  return dtend ? toBasicUtc(dtend.getFirstValue()) : nowBasicUtc();
}

function isAllDayProp(prop) {
  if (!prop) return false;
  const v = prop.getFirstValue();
  return v instanceof ICAL.Time && v.isDate;
}

/* ── Helpers: 2.5/14.x <Exceptions> round-trip ─────────────────────── */

/** Inbound: parse `<Exceptions><Exception>` children of `adNode`. For
 *  each, read `<ExceptionStartTime>` (the ORIGINAL occurrence date) and
 *  either add an EXDATE to the master (`Deleted=1`) or build an override
 *  VEVENT keyed by RECURRENCE-ID. Mirrors legacy `setItemRecurrence` at
 *  sync.js:1344-1372. */
function appendInboundExceptions({ adNode, vcal, vevent, asVersion, defaultTimezone }) {
  const wrapper = childByTag(adNode, "Exceptions");
  if (!wrapper) return;
  const masterUid = stringOf(vevent.getFirstPropertyValue("uid"));

  for (const exc of wrapper.children) {
    if (exc.tagName !== "Exception") continue;
    const startStr = readPathFrom(exc, ["ExceptionStartTime"]);
    if (!startStr) continue;
    const ridDate = parseEasUtc(startStr);
    if (!ridDate) continue;

    if (readPathFrom(exc, ["Deleted"]) === "1") {
      addExdate(vevent, ridDate);
      continue;
    }

    const override = new ICAL.Component(["vevent", [], []]);
    vcal.addSubcomponent(override);
    if (masterUid) override.updatePropertyWithValue("uid", masterUid);
    const ridProp = new ICAL.Property("recurrence-id", override);
    ridProp.setValue(jsDateToIcalUtcTime(ridDate));
    override.addProperty(ridProp);
    populateVeventFromAd({ adNode: exc, vevent: override, asVersion, defaultTimezone });
  }
}

/** Outbound: emit a `<Exceptions>` wrapper from the VCALENDAR's EXDATEs
 *  on the master plus any sibling override VEVENTs (subcomponents that
 *  share the master's UID and carry RECURRENCE-ID). 2.5/14.x only —
 *  16.1 sends each exception as its own `<Change>` at the runner level.
 *  Mirrors legacy `getItemRecurrence` at sync.js:1488-1505. */
function appendOutboundExceptions({ builder, vcal, vevent, asVersion, defaultTimezone, syncRecurrence }) {
  if (!vcal) return;
  const masterUid = stringOf(vevent.getFirstPropertyValue("uid"));
  const exdates = collectExdates(vevent);
  const overrides = [];
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    if (sub === vevent) continue;
    const subUid = stringOf(sub.getFirstPropertyValue("uid"));
    const rid = sub.getFirstProperty("recurrence-id");
    if (subUid === masterUid && rid) overrides.push(sub);
  }
  if (!exdates.length && !overrides.length) return;

  builder.otag("Exceptions");
  for (const ex of exdates) {
    builder.otag("Exception");
      builder.atag("ExceptionStartTime", icalTimeToBasicUtc(ex));
      builder.atag("Deleted", "1");
    builder.ctag();
  }
  for (const override of overrides) {
    const rid = override.getFirstPropertyValue("recurrence-id");
    builder.otag("Exception");
      builder.atag("ExceptionStartTime", icalTimeToBasicUtc(rid));
      // Recurse into the writer in exception-mode. We're already on the
      // Calendar codepage; the recursive call may switch to AirSyncBase
      // (Body / Location 16.1) and switches back to Calendar before
      // returning, so we resume cleanly here.
      appendApplicationDataFromIcal({
        builder, ical: override, asVersion, defaultTimezone, syncRecurrence,
        isException: true,
      });
    builder.ctag();
  }
  builder.ctag();
}

function addExdate(vevent, jsDate) {
  const prop = new ICAL.Property("exdate", vevent);
  prop.setValue(jsDateToIcalUtcTime(jsDate));
  vevent.addProperty(prop);
}

function removeExdate(vevent, jsDate) {
  const target = formatBasicUtc(jsDate);
  for (const p of vevent.getAllProperties("exdate")) {
    const v = p.getFirstValue();
    if (v instanceof ICAL.Time && icalTimeToBasicUtc(v) === target) {
      vevent.removeProperty(p);
    }
  }
}

function collectExdates(vevent) {
  const out = [];
  for (const p of vevent.getAllProperties("exdate")) {
    const v = p.getFirstValue();
    if (v instanceof ICAL.Time) out.push(v);
  }
  return out;
}

function jsDateToIcalUtcTime(d) {
  const t = new ICAL.Time({
    year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(),
    hour: d.getUTCHours(), minute: d.getUTCMinutes(), second: d.getUTCSeconds(),
    isDate: false,
  });
  t.zone = ICAL.Timezone.utcTimezone;
  return t;
}

function icalTimeToBasicUtc(t) {
  if (t instanceof ICAL.Time) {
    const d = t.toJSDate();
    return formatBasicUtc(d);
  }
  return formatBasicUtc(new Date(t));
}

function instanceUtcToIcalTime(jsDate) {
  return jsDateToIcalUtcTime(jsDate);
}

function instanceUtcToIcalString(jsDate) {
  return jsDateToIcalUtcTime(jsDate).toString();
}

/* ── Helpers: alarms ───────────────────────────────────────────────── */

function appendDisplayAlarm(vevent, minutesBeforeStart) {
  const alarm = new ICAL.Component(["valarm", [], []]);
  alarm.updatePropertyWithValue("action", "DISPLAY");
  const trig = new ICAL.Property("trigger", alarm);
  const dur = new ICAL.Duration({ minutes: Math.abs(minutesBeforeStart), isNegative: minutesBeforeStart > 0 });
  trig.setValue(dur);
  alarm.addProperty(trig);
  vevent.addSubcomponent(alarm);
}

function alarmMinutes(alarm, dtstartProp) {
  const trig = alarm.getFirstProperty("trigger");
  if (!trig) return null;
  const v = trig.getFirstValue();
  if (v instanceof ICAL.Duration) {
    const total = v.toSeconds();
    return Math.round(-total / 60);   // EAS minutes = before start, positive
  }
  if (v instanceof ICAL.Time && dtstartProp) {
    const start = dtstartProp.getFirstValue();
    if (!(start instanceof ICAL.Time)) return null;
    const diffSec = start.subtractDateTz(v).toSeconds();
    return Math.round(diffSec / 60);
  }
  return null;
}

/* ── Helpers: attendees ────────────────────────────────────────────── */

function collectAttendees(adNode) {
  const out = [];
  const wrapper = childByTag(adNode, "Attendees");
  if (!wrapper) return out;
  for (const a of wrapper.children) {
    if (a.tagName !== "Attendee") continue;
    const email = readPathFrom(a, ["Email"]);
    if (!email) continue;
    const item = { email, cn: readPathFrom(a, ["Name"]) };
    const status = readPathFrom(a, ["AttendeeStatus"]);
    if (status) item.partstat = ATTENDEESTATUS_TO_PARTSTAT[status] ?? "NEEDS-ACTION";
    const type = readPathFrom(a, ["AttendeeType"]);
    if (type === "1")      { item.role = "REQ-PARTICIPANT"; item.cutype = "INDIVIDUAL"; }
    else if (type === "2") { item.role = "OPT-PARTICIPANT"; item.cutype = "INDIVIDUAL"; }
    else if (type === "3") { item.role = "NON-PARTICIPANT"; item.cutype = "RESOURCE"; }
    out.push(item);
  }
  return out;
}

function ownerMatchesOrganizer(/* orgProp */) {
  // We don't have the account user here; the caller can override
  // `MeetingStatus` later if needed. Default: assume not-organizer (3/7)
  // when an organizer is present at all - matches legacy fallback.
  return false;
}

function stripMailto(s) {
  if (!s) return "";
  return String(s).replace(/^mailto:/i, "");
}

/* ── Helpers: recurrence ───────────────────────────────────────────── */

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
    const days = [];
    const ical = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
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

function appendRecurrence(builder, rruleProp, dtstartProp) {
  const r = rruleProp.getFirstValue();   // ICAL.Recur
  if (!r) return;

  const freq = r.freq;
  const startDate = dtstartProp?.getFirstValue();
  let type = 0;
  let monthDays = r.parts?.BYMONTHDAY ?? [];
  let weekDays  = (r.parts?.BYDAY ?? []).slice();
  let months    = r.parts?.BYMONTH ?? [];
  const weeks   = [];

  // Unpack ±NDD style days into weekDays + weekOfMonth.
  for (let i = 0; i < weekDays.length; i++) {
    const m = /^([+-]?\d*)(SU|MO|TU|WE|TH|FR|SA)$/.exec(weekDays[i]);
    if (!m) continue;
    const n = parseInt(m[1] || "0", 10);
    const dow = ["SU","MO","TU","WE","TH","FR","SA"].indexOf(m[2]) + 1;
    weekDays[i] = dow;
    if (n) weeks[i] = n === -1 ? 5 : n;
  }

  if (freq === "WEEKLY") {
    type = 1;
    if (!weekDays.length && startDate) weekDays = [(startDate.dayOfWeek?.() ?? 1)];
  } else if (freq === "MONTHLY" && weeks.length) {
    type = 3;
  } else if (freq === "MONTHLY") {
    type = 2;
    if (!monthDays.length && startDate) monthDays = [startDate.day];
  } else if (freq === "YEARLY" && weeks.length) {
    type = 6;
  } else if (freq === "YEARLY") {
    type = 5;
    if (!monthDays.length && startDate) monthDays = [startDate.day];
    if (!months.length && startDate) months = [startDate.month];
  }

  builder.otag("Recurrence");
    builder.atag("Type", String(type));
    if (monthDays[0]) builder.atag("DayOfMonth", String(monthDays[0]));
    if (weekDays.length) {
      let bits = 0;
      for (const d of weekDays) bits |= 1 << (d - 1);
      builder.atag("DayOfWeek", String(bits));
    }
    builder.atag("Interval", String(r.interval ?? 1));
    if (months.length) builder.atag("MonthOfYear", String(months[0]));
    if (r.count) builder.atag("Occurrences", String(r.count));
    else if (r.until) builder.atag("Until", toBasicUtc(r.until));
    if (weeks.length) builder.atag("WeekOfMonth", String(weeks[0]));
  builder.ctag();
}

/* ── Helpers: body codepage ───────────────────────────────────────── */

function appendBody(builder, vevent, asVersion) {
  const desc = stringOf(vevent.getFirstPropertyValue("description"));
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
  builder.switchpage("Calendar");
}

function useAirSyncBaseBody(asVersion) {
  return asVersion !== "2.5";
}

/* ── Helpers: ICAL.js plumbing ─────────────────────────────────────── */

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

function parseFirstVevent(ical) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return null;
  // Pick the master: the first vevent without RECURRENCE-ID. Fallback
  // to the very first vevent if no master is identifiable (defensive).
  const all = vcal.getAllSubcomponents("vevent");
  for (const v of all) {
    if (!v.getFirstProperty("recurrence-id")) return v;
  }
  return all[0] ?? null;
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
