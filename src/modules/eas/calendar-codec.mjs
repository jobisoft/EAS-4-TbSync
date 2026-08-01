/**
 * EAS Calendar (codepage 4) + AirSyncBase Body (17) ⇆ iCal VEVENT codec.
 *
 * Mirrors the legacy `EAS-4-TbSync/content/includes/calendarsync.js` field
 * map. Round-trips the common field set: Subject, Location, Body,
 * Start/End/AllDay, BusyStatus, Sensitivity, Reminder, Categories,
 * Organizer, Attendees, MeetingStatus, ResponseType, UID, recurrence.
 *
 * Recurrence covers the RRULE plus exceptions, which reach us in two
 * shapes. Embedded: an `<Exceptions>` block on the master item, read by
 * `appendInboundExceptions` and written by `appendOutboundExceptions`
 * (outbound on 2.5/14.x only). Per-instance: one `<Change>` per
 * occurrence carrying `<InstanceId>`, handled by `applyInstanceChange` /
 * `applyInstanceDelete` and emitted by `appendInstanceChanges` (16.1
 * only). All of it is gated on the account's `syncRecurrence` option.
 *
 * The TimeZone blob (every version, but not for all-day events on 16.1,
 * which are floating dates) is encoded / decoded via `TimeZoneBlob` in
 * `timezone-blob.mjs`. When the server's blob is all-zero we fall back to
 * the host's default IANA zone.
 *
 * Local items carry the EAS server-assigned ServerId on a custom
 * `X-EAS-SERVERID` property so pull/push paths can find the local item
 * without a separate id map (mirrors the contact-codec's approach).
 */

import ICAL from "../../vendor/ical.min.js";
import { readPathFrom } from "./wbxml-helpers.mjs";
import { TimeZoneBlob, isAllZero } from "./timezone-blob.mjs";
import {
  guessTimezoneByStdDstOffset,
  tzInfoForBlob,
  getIcalTimezone,
} from "./timezone-mapping.mjs";

const X_EAS_SERVERID = "X-EAS-SERVERID";
const X_EAS_RESPONSETYPE = "X-EAS-RESPONSETYPE";
const X_EAS_MEETINGSTATUS = "X-EAS-MEETINGSTATUS";

// EAS BusyStatus → iCal TRANSP. Tentative (1) maps to "no TRANSP" so the
// caller falls back to STATUS=TENTATIVE; the codec mirrors legacy here.
const BUSYSTATUS_TO_TRANSP = {
  0: "TRANSPARENT",
  1: null,
  2: "OPAQUE",
  3: "OPAQUE",
  4: "OPAQUE",
};
const TRANSP_TO_BUSYSTATUS = { TRANSPARENT: "0", OPAQUE: "2" };

// EAS Sensitivity → iCal CLASS.
const SENSITIVITY_TO_CLASS = {
  0: "PUBLIC",
  1: "PRIVATE",
  2: "PRIVATE",
  3: "CONFIDENTIAL",
};
const CLASS_TO_SENSITIVITY = { PUBLIC: "0", PRIVATE: "2", CONFIDENTIAL: "3" };

// EAS AttendeeStatus → iCal PARTSTAT.
const ATTENDEESTATUS_TO_PARTSTAT = {
  0: "NEEDS-ACTION",
  2: "TENTATIVE",
  3: "ACCEPTED",
  4: "DECLINED",
  5: "ACCEPTED",
};

/* ── Reader: ApplicationData → iCal VEVENT ─────────────────────────── */

export function applicationDataToIcal({
  adNode,
  existingIcal,
  serverID,
  asVersion,
  defaultTimezone,
  syncRecurrence,
  uid,
  userEmail,
  eventLog,
}) {
  // Merge mode: parse the existing iCal and overlay only fields the AD
  // mentions. Fall through to a fresh build when there's no existing
  // blob (server-pushed Add) or the existing blob is unparseable.
  let vcal = null;
  let vevent = null;
  if (existingIcal) {
    vcal = parseVCalendar(existingIcal);
    if (vcal) {
      // Pick the master (no RECURRENCE-ID) so partial Changes don't
      // clobber recurrence overrides. Fallback to the first vevent.
      const all = vcal.getAllSubcomponents("vevent");
      vevent =
        all.find((v) => !v.getFirstProperty("recurrence-id")) ?? all[0] ?? null;
    }
  }
  if (!vcal || !vevent) {
    vcal = newVCalendar();
    vevent = new ICAL.Component(["vevent", [], []]);
    vcal.addSubcomponent(vevent);
  }

  if (uid) vevent.updatePropertyWithValue("uid", uid);
  vevent.updatePropertyWithValue(X_EAS_SERVERID.toLowerCase(), serverID);

  if (eventLog) {
    const orgEmailRaw = readPathFrom(adNode, ["OrganizerEmail"]);
    const orgNameRaw = readPathFrom(adNode, ["OrganizerName"]);
    const hasOrgInfo =
      childByTag(adNode, "OrganizerEmail") ||
      childByTag(adNode, "OrganizerName");
    eventLog(
      "info",
      `[calendar-codec] receive OrganizerInfo: present=${!!hasOrgInfo} OrganizerEmail=${JSON.stringify(orgEmailRaw ?? null)} OrganizerName=${JSON.stringify(orgNameRaw ?? null)}`,
    );
  }

  populateVeventFromAd({
    adNode,
    vevent,
    asVersion,
    defaultTimezone,
    userEmail,
  });

  // Recurrence + 2.5/14.x exceptions. Gated on the account-level
  // syncRecurrence flag. Only touch RRULE / exceptions when the AD
  // mentions them; otherwise leave whatever the existing blob carried.
  if (syncRecurrence) {
    const recNode = childByTag(adNode, "Recurrence");
    if (recNode) {
      vevent.removeAllProperties("rrule");
      const rrule = recurrenceToRrule(recNode);
      if (rrule && /^FREQ=[A-Z]+/.test(rrule)) {
        const prop = new ICAL.Property("rrule", vevent);
        prop.setValue(ICAL.Recur.fromString(rrule));
        vevent.addProperty(prop);
      }
    }
    if (childByTag(adNode, "Exceptions")) {
      // Clear existing override vevents (anything with RECURRENCE-ID)
      // so the AD's exception set replaces them wholesale.
      for (const sub of vcal.getAllSubcomponents("vevent")) {
        if (sub.getFirstProperty("recurrence-id")) {
          vcal.removeSubcomponent(sub);
        }
      }
      appendInboundExceptions({
        adNode,
        vcal,
        vevent,
        asVersion,
        defaultTimezone,
      });
    }
  }

  return vcal.toString();
}

/** Populate a VEVENT (master or override) from an EAS <ApplicationData>
 *  or <Exception> node. The set of fields is the same on both - legacy
 *  reuses `setThunderbirdItemFromWbxml` for both paths.
 *  Returns nothing; mutates `vevent`. */
function populateVeventFromAd({
  adNode,
  vevent,
  asVersion,
  defaultTimezone,
  userEmail,
}) {
  // Subject / Location.
  const subject = readPathFrom(adNode, ["Subject"]);
  if (subject) vevent.updatePropertyWithValue("summary", subject);

  const locDisplay =
    readPathFrom(adNode, ["Location", "DisplayName"]) ??
    readPathFrom(adNode, ["Location"]);
  if (locDisplay) vevent.updatePropertyWithValue("location", locDisplay);

  // Body (codepage-aware; AirSyncBase ≥14.x).
  if (useAirSyncBaseBody(asVersion)) {
    const data = readPathFrom(adNode, ["Body", "Data"]);
    if (data) vevent.updatePropertyWithValue("description", data);
  } else {
    const data = readPathFrom(adNode, ["Body"]);
    if (data) vevent.updatePropertyWithValue("description", data);
  }

  // Resolve effective timezone: from the TimeZone blob when the item
  // carried a real one, otherwise the host's default zone (servers that
  // send an all-zero blob).
  //
  // An all-day event does not end up with a zone at all. iCalendar has a
  // form for exactly this - a DATE value, `isDate: true`, written as
  // `DTSTART;VALUE=DATE:20260915` with no TZID. RFC 5545 §3.3.4 calls
  // these floating: they mean the same calendar day in every zone, which
  // is what a birthday or a bank holiday actually is. `writeDateProp`
  // always stores that form for all-day.
  //
  // The zone is still needed to *read* the boundary, though, and only
  // then - which is what `fromBlob` selects. Exchange ≤14.x encodes an
  // all-day boundary as midnight-in-zone expressed as UTC, so recovering
  // the calendar date from `20230831T220000Z` needs the zone; it is then
  // discarded. An all-day event on AS 16.1 carries no blob and needs no
  // zone, because its wire value is already the date.
  const { tzId, fromBlob } = resolveTimezone(adNode, defaultTimezone);
  const allDay = readPathFrom(adNode, ["AllDayEvent"]) === "1";

  // Start / End. EAS sends UTC strings; convert on the way in.
  const startUtc = readPathFrom(adNode, ["StartTime"]);
  const endUtc = readPathFrom(adNode, ["EndTime"]);
  if (startUtc) writeDateProp(vevent, "dtstart", startUtc, tzId, allDay, fromBlob);
  if (endUtc) writeDateProp(vevent, "dtend", endUtc, tzId, allDay, fromBlob);

  // DtStamp - preserve when present. A 16.x *client* MUST NOT send it
  // ([MS-ASCAL] §2.2.2.18), which is why the writer omits it there, but
  // the server does send it on 16.1 - observed on every item of an
  // Exchange Online initial sync.
  const dtStamp = readPathFrom(adNode, ["DtStamp"]);
  if (dtStamp) writeDateProp(vevent, "dtstamp", dtStamp, "UTC", false, false);

  // BusyStatus → TRANSP. STATUS is computed below from BusyStatus +
  // MeetingStatus together (legacy calendarsync.js:235-265).
  const busy = readPathFrom(adNode, ["BusyStatus"]);
  const transp = busy ? BUSYSTATUS_TO_TRANSP[busy] : undefined;
  if (transp) vevent.updatePropertyWithValue("transp", transp);

  // Sensitivity → CLASS.
  const sens = readPathFrom(adNode, ["Sensitivity"]);
  if (sens && SENSITIVITY_TO_CLASS[sens]) {
    vevent.updatePropertyWithValue("class", SENSITIVITY_TO_CLASS[sens]);
  }

  // Reminder → VALARM (DISPLAY, offset relative to start in minutes).
  // Merge-aware: the presence of <Reminder> in the AD signals the
  // server's authoritative alarm state - clear any existing VALARMs
  // first so we never stack alarms on a partial Change. An empty
  // <Reminder/> means "no alarm".
  const hasReminderTag = childByTag(adNode, "Reminder") != null;
  if (hasReminderTag) {
    for (const a of vevent.getAllSubcomponents("valarm")) {
      vevent.removeSubcomponent(a);
    }
    vevent.removeAllProperties("x-moz-lastack");
  }
  const reminderMinutes = readPathFrom(adNode, ["Reminder"]);
  if (reminderMinutes != null && reminderMinutes !== "" && startUtc) {
    appendDisplayAlarm(vevent, parseInt(reminderMinutes, 10));
    const startDate = parseEasUtc(startUtc);
    if (startDate && startDate.getTime() < Date.now()) {
      vevent.updatePropertyWithValue("x-moz-lastack", nowBasicUtc());
    }
  }

  // Categories. Merge-aware: clear existing when <Categories> is in
  // the AD; the AD's children replace the prior set.
  if (childByTag(adNode, "Categories")) {
    vevent.removeAllProperties("categories");
    const cats = collectChildren(adNode, "Categories", "Category");
    if (cats.length) {
      const prop = new ICAL.Property("categories", vevent);
      prop.setValues(cats);
      vevent.addProperty(prop);
    }
  }

  // Organizer. Merge-aware on either OrganizerEmail or OrganizerName
  // being present in the AD - Exchange Online sends both on 16.1, on
  // every item of an initial sync, so absence means "no change" rather
  // than "unset".
  const orgEmail = readPathFrom(adNode, ["OrganizerEmail"]);
  const orgName = readPathFrom(adNode, ["OrganizerName"]);
  const hasOrgInfo =
    childByTag(adNode, "OrganizerEmail") || childByTag(adNode, "OrganizerName");
  if (hasOrgInfo) {
    vevent.removeAllProperties("organizer");
    if (orgEmail) {
      const prop = new ICAL.Property("organizer", vevent);
      prop.setValue("mailto:" + orgEmail);
      if (orgName) prop.setParameter("cn", orgName);
      vevent.addProperty(prop);
    }
  }

  // Attendees. ResponseType is the event-level fallback for the
  // self-attendee's PARTSTAT when the per-attendee AttendeeStatus is
  // missing (legacy calendarsync.js:200-206). Default for everyone
  // else is NEEDS-ACTION (legacy line 206). Merge-aware: when
  // <Attendees> is present in the AD, the AD's set replaces the
  // existing set wholesale (matches EAS partial-Change semantics).
  const respType = readPathFrom(adNode, ["ResponseType"]);
  if (childByTag(adNode, "Attendees")) {
    vevent.removeAllProperties("attendee");
    const attendees = collectAttendees(adNode, userEmail, respType);
    for (const a of attendees) {
      const prop = new ICAL.Property("attendee", vevent);
      prop.setValue("mailto:" + a.email);
      if (a.cn) prop.setParameter("cn", a.cn);
      if (a.role) prop.setParameter("role", a.role);
      if (a.partstat) prop.setParameter("partstat", a.partstat);
      if (a.cutype) prop.setParameter("cutype", a.cutype);
      vevent.addProperty(prop);
    }
  }

  // Pass-through ResponseType so upsync round-trips the original value.
  if (respType)
    vevent.updatePropertyWithValue(X_EAS_RESPONSETYPE.toLowerCase(), respType);

  // STATUS computed from BusyStatus + MeetingStatus together. Mirrors
  // legacy calendarsync.js:244-265:
  //   - BusyStatus=1 (tentative) seeds tbStatus = TENTATIVE.
  //   - MeetingStatus M (0x1) means "is a meeting"; C (0x4) means
  //     "cancelled". M+C → CANCELLED (overrides TENTATIVE). M alone →
  //     CONFIRMED, but only when not already TENTATIVE.
  //   - The R bit (0x2) is "received from another organizer"; legacy
  //     uses it to populate a calendar-level fallbackOrganizerName,
  //     which the WebExtension calendar API doesn't expose. Skip.
  let tbStatus = busy === "1" ? "TENTATIVE" : null;
  const meetingStatus = readPathFrom(adNode, ["MeetingStatus"]);
  if (meetingStatus) {
    vevent.updatePropertyWithValue(
      X_EAS_MEETINGSTATUS.toLowerCase(),
      meetingStatus,
    );
    const ms = parseInt(meetingStatus, 10) || 0;
    if (ms & 0x1) {
      if (ms & 0x4) tbStatus = "CANCELLED";
      else if (!tbStatus) tbStatus = "CONFIRMED";
    }
  }
  if (tbStatus) vevent.updatePropertyWithValue("status", tbStatus);
}

/** Public entry point for the 16.1 InstanceId path: called from the
 *  sync runner when an inbound `<Change>` carries an `<InstanceId>`.
 *  Locates or creates the override VEVENT keyed by RECURRENCE-ID, then
 *  populates it from `adNode`. For deletions, the runner adds an EXDATE
 *  via `addExdateToMaster` instead. */
export function applyInstanceChange({
  ical,
  adNode,
  instanceUtc,
  asVersion,
  defaultTimezone,
  userEmail,
}) {
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
  populateVeventFromAd({
    adNode: adNode,
    vevent: override,
    asVersion,
    defaultTimezone,
    userEmail,
  });
  return vcal.toString();
}

/** Outbound 16.1: emit one `<Change ServerId=master>` per current
 *  EXDATE / RECURRENCE-ID override on the master. Idempotent - re-asserts
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
export function appendInstanceChanges({
  builder,
  blob,
  serverID,
  asVersion,
  defaultTimezone,
  syncRecurrence,
  userEmail,
  fallbackOrganizerName,
  eventLog,
}) {
  if (asVersion !== "16.1") return;
  const vcal = parseVCalendar(blob);
  if (!vcal) return;
  const master = vcal.getFirstSubcomponent("vevent") ?? null;
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
    // InstanceId belongs to the command, not to its payload: [MS-ASAIRS]
    // §2.2.2.25 makes it a child of airsync:Change (or airsync:Delete),
    // and [MS-ASCMD] Change says a 16.x calendar item is identified by
    // the master's ServerId together with it. Inside ApplicationData it
    // is only ever valid in a *response*, for an orphan instance.
    builder.switchpage("AirSyncBase");
    builder.atag("InstanceId", icalTimeToBasicUtc(ex));
    builder.switchpage("AirSync");
    builder.otag("ApplicationData");
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
    // Sibling of ServerId, not part of the payload - see the EXDATE
    // branch above for the citations.
    builder.switchpage("AirSyncBase");
    builder.atag("InstanceId", icalTimeToBasicUtc(rid));
    builder.switchpage("AirSync");
    builder.otag("ApplicationData");
    // appendApplicationDataFromIcal switches to Calendar at entry
    // and may bounce to AirSyncBase for Body / Location, but always
    // returns to Calendar before the closing tag.
    appendApplicationDataFromIcal({
      builder,
      ical: override,
      asVersion,
      defaultTimezone,
      syncRecurrence,
      isException: true,
      userEmail,
      fallbackOrganizerName,
      eventLog,
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

  // Drop any existing override at this RECURRENCE-ID - server says it's
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
  builder,
  ical,
  asVersion,
  defaultTimezone,
  syncRecurrence,
  isException = false,
  userEmail,
  fallbackOrganizerName,
  eventLog,
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

  const dtstart = vevent.getFirstProperty("dtstart");
  const dtend = vevent.getFirstProperty("dtend");
  const allDay = isAllDayProp(dtstart) && isAllDayProp(dtend);
  // Source zone for all-day Start/End on ≤14.x — must match the zone we
  // describe in the TimeZone blob (same precedence as buildTimezoneBlob)
  // so the server reads the boundary back on the right calendar day.
  const allDaySourceTzid = pickSourceTzid(vevent) ?? defaultTimezone ?? "UTC";

  // Outbound timezone. [MS-ASCAL] §2.2.2.44 lists the element as
  // supported in every protocol version, 16.1 included, and states no
  // restriction on client use - unlike UID and DtStamp, where the same
  // document says outright that a 16.x client MUST NOT send them.
  // Exchange Online sends it on 16.1 too, on every timed event.
  //
  // Two cases get none:
  //  - an embedded ≤14.x <Exception>, whose parent payload already
  //    carries the master's blob. A 16.1 per-instance <Change> is its own
  //    top-level command (see `appendInstanceChanges`) and needs its own.
  //  - an all-day event on 16.1. [MS-ASCAL] §2.2.2.1 is explicit: with
  //    AllDayEvent set to 1 the client "MUST NOT include the TimeZone
  //    element", and the server will not send one either - "a client
  //    SHOULD interpret this event to be at the given date(s) regardless
  //    of the time zone used". That is the floating semantics iCalendar
  //    gives a DATE value, so the two formats agree. Sending a blob would
  //    also flip `fromBlob` on the echo, converting a value that has to
  //    be read verbatim.
  const embeddedException = isException && asVersion !== "16.1";
  const floatingAllDay = allDay && asVersion === "16.1";
  if (!embeddedException && !floatingAllDay) {
    const blob = buildTimezoneBlob(vevent, defaultTimezone);
    builder.atag("TimeZone", blob.easTimeZone64);
  }

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

  // Organizer (≤14.x; not inside an exception). When the iCal ORGANIZER
  // has no CN= parameter, fall back to the per-folder name captured by
  // the inbound side from prior server responses (Phase 5 row 5.4.5,
  // stored on `account.custom.fallbackOrganizerNames[<collectionId>]`).
  // Mirrors legacy's `calendar.fallbackOrganizerName` consumption (note:
  // legacy didn't actually wire the fallback into the OrganizerName emit
  // either - it just stashed it; we lift it on emit here).
  if (asVersion !== "16.1" && !isException) {
    const orgProp = vevent.getFirstProperty("organizer");
    const localEmail = orgProp ? stripMailto(orgProp.getFirstValue()) : null;
    const localCn = orgProp?.getParameter("cn") ?? null;
    let emittedEmail = null;
    let emittedName = null;
    if (orgProp) {
      const cn = localCn;
      const name = cn || fallbackOrganizerName;
      if (name) {
        builder.atag("OrganizerName", name);
        emittedName = name;
      }
      if (localEmail) {
        builder.atag("OrganizerEmail", localEmail);
        emittedEmail = localEmail;
      }
    }
    if (eventLog) {
      eventLog(
        "info",
        `[calendar-codec] push OrganizerInfo: local ORGANIZER email=${JSON.stringify(localEmail)} cn=${JSON.stringify(localCn)} → emitted OrganizerEmail=${JSON.stringify(emittedEmail)} OrganizerName=${JSON.stringify(emittedName)}`,
      );
    }
  }

  // DtStamp (≤14.x).
  if (asVersion !== "16.1") {
    const ds = vevent.getFirstProperty("dtstamp");
    builder.atag(
      "DtStamp",
      ds ? toBasicUtc(ds.getFirstValue()) : nowBasicUtc(),
    );
  }

  // EndTime. All-day: 16.1 uses the "fake local as UTC" form; ≤14.x uses
  // local midnight in the TimeZone-blob zone expressed as UTC. Both avoid
  // the ±1-day shift a naive UTC conversion would cause in non-UTC zones.
  builder.atag("EndTime", endTimeFor(dtend, asVersion, allDay, allDaySourceTzid));

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

  // Reminder. `alarmMinutes` is responsible for surfacing info-level
  // event-log entries when an absolute VALARM is converted or a
  // negative-offset alarm is dropped.
  // For events with no VALARM on AS 16.1, emit empty <Reminder/> to
  // explicitly clear the server-side default reminder.
  // Per [MS-ASCAL] §2.2.2.38, the empty-tag form is only documented
  // as supported on 16.0/16.1.
  const alarm = vevent.getFirstSubcomponent("valarm");
  if (alarm) {
    const minutes = alarmMinutes(alarm, dtstart, eventLog);
    if (minutes != null && minutes >= 0)
      builder.atag("Reminder", String(minutes));
  } else if (asVersion === "16.1") {
    builder.atag("Reminder");
  }

  // Sensitivity.
  const cls = stringOf(vevent.getFirstPropertyValue("class"));
  builder.atag("Sensitivity", CLASS_TO_SENSITIVITY[cls] ?? "0");

  // Subject + StartTime.
  builder.atag("Subject", stringOf(vevent.getFirstPropertyValue("summary")));
  builder.atag("StartTime", startTimeFor(dtstart, asVersion, allDay, allDaySourceTzid));

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
      // Legacy emits an empty <Attendees/> container on ≤14.x to force
      // the server to clear its copy of the attendee list (otherwise
      // server-side stale attendees survive the upsync). We skip the
      // empty container on 16.1, assuming the server treats absence as
      // no-change - inherited from legacy calendarsync.js:497-498 and not
      // verified against [MS-ASCAL]. Note the server does send
      // <Attendees/>, empty included, on 16.1: observed on every item of
      // an Exchange Online initial sync. Whether omitting it outbound
      // actually clears the list there is untested.
      if (asVersion !== "16.1") {
        builder.atag("Attendees");
      }
    } else {
      const cancelled = status === "CANCELLED";
      const orgProp = vevent.getFirstProperty("organizer");
      // R bit (received-from-another-organizer): the local user is NOT
      // the organizer iff the ORGANIZER email differs from
      // `account.custom.user`. The previous code's
      // `ownerMatchesOrganizer` was a hardcoded `false`, which made
      // every attendee'd event look "received" — so events the user
      // organized were emitted as MeetingStatus=3 (received) instead of
      // 1 (organizer). Mirrors legacy calendarsync.js:447-450.
      const orgEmail = orgProp
        ? stripMailto(orgProp.getFirstValue()).toLowerCase()
        : "";
      const userEmailLower = userEmail
        ? String(userEmail).toLowerCase()
        : "";
      const isReceived =
        !!orgEmail && (!userEmailLower || orgEmail !== userEmailLower);
      if (cancelled) builder.atag("MeetingStatus", isReceived ? "7" : "5");
      else builder.atag("MeetingStatus", isReceived ? "3" : "1");

      builder.otag("Attendees");
      for (const a of attendees) {
        builder.otag("Attendee");
        builder.atag("Email", stripMailto(a.getFirstValue()));
        const cn =
          a.getParameter("cn") ?? stripMailto(a.getFirstValue()).split("@")[0];
        builder.atag("Name", cn);
        if (asVersion !== "2.5") {
          const role = a.getParameter("role");
          const cutype = a.getParameter("cutype");
          let type = "2";
          if (
            cutype === "RESOURCE" ||
            cutype === "ROOM" ||
            role === "NON-PARTICIPANT"
          )
            type = "3";
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
        builder,
        vcal,
        vevent,
        asVersion,
        defaultTimezone,
        syncRecurrence,
        userEmail,
        fallbackOrganizerName,
        eventLog,
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
  // The master, not merely the first VEVENT - `readEasServerIdFromIcal`
  // skips overrides when it reads the stamp back, so stamping the first
  // component would write it somewhere the reader never looks.
  const vevent = pickMasterVevent(vcal);
  if (!vevent) return ical;
  vevent.updatePropertyWithValue(X_EAS_SERVERID.toLowerCase(), serverID);
  return vcal.toString();
}

/* ── Helpers: timezone resolution ──────────────────────────────────── */

/** Resolve the effective IANA tzid for an inbound event and report
 *  whether it came from a real (non-empty) TimeZone blob. `fromBlob`
 *  drives the all-day boundary reading: a real blob means the server
 *  encoded all-day Start/End as midnight-in-zone-expressed-as-UTC (real
 *  Exchange ≤14.x), which must be converted back through the zone; a
 *  missing/empty blob (all-day events on AS 16.1, and
 *  Z-Push/Kopano/Grommunio, which send an all-zero blob) means the date
 *  is a literal "fake local as UTC" and must be read verbatim — see
 *  writeDateProp. Keyed on the item rather than the version: Exchange
 *  sends a blob on 16.1 for timed events. */
function resolveTimezone(adNode, defaultTimezone) {
  const blobB64 = readPathFrom(adNode, ["TimeZone"]);
  if (!blobB64 || isAllZero(blobB64)) {
    return { tzId: defaultTimezone || "UTC", fromBlob: false };
  }
  const blob = new TimeZoneBlob();
  blob.easTimeZone64 = blobB64;
  // utcOffset is "minutes from local to UTC" (e.g. -60 for CET); daylight
  // shifts by daylightBias (typically -60 again for European DST). Match
  // legacy calendarsync.js:106-107.
  const stdOffset = blob.utcOffset;
  const dstOffset = blob.daylightBias + blob.utcOffset;
  const stdName = blob.standardName;
  // The SYSTEMTIME transition dates distinguish zones that share an
  // offset but switch DST on different dates; pass them so the resolver
  // can pick the right one instead of collapsing every UTC+1/+2 zone onto
  // a single fallback (which makes recurring events drift by an hour
  // around the mismatched transition).
  const dstRule = dstRuleFromBlob(blob);
  const tzid = guessTimezoneByStdDstOffset(stdOffset, dstOffset, stdName, dstRule);
  return { tzId: tzid || defaultTimezone || "UTC", fromBlob: true };
}

/** Extract the nth-weekday DST transition rule from a decoded TimeZone
 *  blob's StandardDate / DaylightDate SYSTEMTIMEs, in the same shape the
 *  timezone-mapping resolver stores per IANA zone. Returns null when the
 *  blob carries no DST (wMonth === 0). */
function dstRuleFromBlob(blob) {
  const std = blob.standardDate;
  const dst = blob.daylightDate;
  if (!std || !dst || std.wMonth === 0 || dst.wMonth === 0) return null;
  return {
    std: { month: std.wMonth, weekOfMonth: std.wDay, dayOfWeek: std.wDayOfWeek },
    dst: { month: dst.wMonth, weekOfMonth: dst.wDay, dayOfWeek: dst.wDayOfWeek },
  };
}

function buildTimezoneBlob(vevent, defaultTimezone) {
  const sourceTzid =
    pickSourceTzid(vevent) ?? defaultTimezone ?? "UTC";
  const tzInfo = tzInfoForBlob(sourceTzid);

  const blob = new TimeZoneBlob();
  blob.utcOffset = tzInfo.std.offset;
  blob.standardBias = 0;
  // Only advertise DST when we can also supply the SYSTEMTIME transition
  // dates. A non-zero daylightBias with all-zero StandardDate/DaylightDate
  // is contradictory: per the Windows TIME_ZONE_INFORMATION rules a zero
  // wMonth means "no DST" (DaylightBias then ignored), but some servers
  // honour the bias and apply DST year-round (or not at all), shifting the
  // event by an hour for part of the year. Keep the blob consistent.
  const hasDstRule = !!(tzInfo.std.switchdate && tzInfo.dst.switchdate);
  blob.daylightBias = hasDstRule ? tzInfo.dst.offset - tzInfo.std.offset : 0;
  blob.standardName = tzInfo.stdWinName;
  blob.daylightName = tzInfo.dstWinName;

  // SYSTEMTIME-shaped switch dates, only when both std and dst rules exist
  // (no-DST zones leave both SYSTEMTIMEs zero-filled and daylightBias=0).
  if (hasDstRule) {
    const std = blob.standardDate;
    std.wMonth = tzInfo.std.switchdate.month;
    std.wDay = tzInfo.std.switchdate.weekOfMonth;
    std.wDayOfWeek = tzInfo.std.switchdate.dayOfWeek;
    std.wHour = tzInfo.std.switchdate.hour;
    std.wMinute = tzInfo.std.switchdate.minute;
    std.wSecond = tzInfo.std.switchdate.second;

    const dst = blob.daylightDate;
    dst.wMonth = tzInfo.dst.switchdate.month;
    dst.wDay = tzInfo.dst.switchdate.weekOfMonth;
    dst.wDayOfWeek = tzInfo.dst.switchdate.dayOfWeek;
    dst.wHour = tzInfo.dst.switchdate.hour;
    dst.wMinute = tzInfo.dst.switchdate.minute;
    dst.wSecond = tzInfo.dst.switchdate.second;
  }

  return blob;
}

/* ── Helpers: dates ────────────────────────────────────────────────── */

/** Pick the source TZID for the outbound TimeZone blob. Matches the
 *  legacy precedence: dtstart's TZID, then dtend's TZID, then explicit
 *  UTC, then null (caller falls back to the host's default zone for
 *  floating / unknown values). */
function pickSourceTzid(vevent) {
  for (const name of ["dtstart", "dtend"]) {
    const prop = vevent?.getFirstProperty(name);
    if (!prop) continue;
    const tzid = prop.getParameter("tzid");
    if (tzid && tzid !== "floating") return tzid;
    const value = prop.getFirstValue?.();
    if (value?.zone?.tzid === "UTC" || value?.isUTC) return "UTC";
  }
  return null;
}

function writeDateProp(vevent, name, easUtc, tzId, allDay, fromBlob) {
  // Replace any existing property of this name so merge-mode partial
  // Changes don't duplicate dtstart/dtend/dtstamp on the master.
  vevent.removeAllProperties(name);
  const prop = new ICAL.Property(name, vevent);
  if (allDay) {
    const d = parseEasUtc(easUtc);
    if (!d) return;
    // The calendar date of an all-day boundary depends on how the server
    // encodes it. Real Exchange ≤14.x sends the midnight that begins the
    // day in the event's TimeZone, expressed as UTC ([MS-ASCAL] §2.2.2.39);
    // for a non-UTC zone that instant lands on the previous/next UTC
    // calendar day, so taking getUTCDate() directly shifts the event by
    // ±1 day. Convert into the resolved zone first and take the wall-clock
    // date there. Items that carry NO real TimeZone blob — all-day events
    // on AS 16.1, and Z-Push/Kopano/Grommunio (all-zero blob) — instead
    // send a "fake local
    // as UTC" YYYYMMDDT000000Z whose UTC date already IS the intended date
    // (it mirrors what we emit outbound), so it must be read verbatim;
    // converting it would shift the date by a day for users west of UTC.
    // `fromBlob` is the discriminator: convert only when a real blob gave
    // us the zone, otherwise read the UTC date verbatim.
    let year = d.getUTCFullYear();
    let month = d.getUTCMonth() + 1;
    let day = d.getUTCDate();
    if (fromBlob && tzId && tzId !== "UTC") {
      const zone = getIcalTimezone(tzId);
      if (zone) {
        const utc = new ICAL.Time({
          year,
          month,
          day,
          hour: d.getUTCHours(),
          minute: d.getUTCMinutes(),
          second: d.getUTCSeconds(),
          isDate: false,
        });
        utc.zone = ICAL.Timezone.utcTimezone;
        const local = utc.convertToZone(zone);
        year = local.year;
        month = local.month;
        day = local.day;
      }
    }
    const date = new ICAL.Time({ year, month, day, isDate: true });
    prop.setValue(date);
  } else {
    const d = parseEasUtc(easUtc);
    if (!d) return;
    // Build the UTC instant first so the wall-clock numerals match the
    // EAS-on-the-wire string. For a TZID-tagged property the wall-clock
    // numerals must be in the named zone (RFC 5545 §3.3.5), so convert
    // before serialising. Without the conversion, ICAL.js reads
    // `DTSTART;TZID=America/Los_Angeles:20260430T003000` as "Apr 30 00:30
    // in LA" — the same numerals tagged with the wrong meaning, shifted
    // from the intended UTC instant by the user's offset.
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

    if (!tzId || tzId === "UTC") {
      prop.setValue(time);
    } else {
      const targetZone = getIcalTimezone(tzId);
      if (targetZone) {
        const local = time.convertToZone(targetZone);
        prop.setValue(local);
        prop.setParameter("tzid", tzId);
      } else {
        // Zone wasn't in the loaded set — keep the value as UTC so the
        // calendar app still renders the correct instant in the user's
        // local zone, just without the TZID hint.
        prop.setValue(time);
      }
    }
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
  const d = value instanceof ICAL.Time ? value.toJSDate() : new Date(value);
  return formatBasicUtc(d);
}

function formatBasicUtc(d) {
  const pad = (n) => String(n).padStart(2, "0");
  return (
    `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}` +
    `T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`
  );
}

function nowBasicUtc() {
  return formatBasicUtc(new Date());
}

/** Read a property's date as `YYYYMMDDT000000Z` from the *local-clock*
 *  year/month/day, with no UTC conversion. Mirrors legacy
 *  `getIsoUtcString(date, false, true, true)` for AS 16.1 all-day. */
function fakeLocalAsUtcDate(prop) {
  if (!prop) return nowBasicUtc();
  return fakeLocalAsUtcFromValue(prop.getFirstValue());
}

/** Same, for a value that is already in hand rather than a property. */
function fakeLocalAsUtcFromValue(v) {
  const pad = (n) => String(n).padStart(2, "0");
  if (v instanceof ICAL.Time) {
    return `${v.year}${pad(v.month)}${pad(v.day)}T000000Z`;
  }
  const d = new Date(v);
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}T000000Z`;
}

/** RRULE UNTIL for the wire. A DATE-valued UNTIL belongs to an all-day
 *  series and is floating, so converting it through the host zone moves
 *  it back a day for anyone east of UTC - the final occurrence then
 *  silently disappears. [MS-ASCAL] §2.2.2.1 also requires Until to carry
 *  no time component when AllDayEvent is 1, alongside StartTime and
 *  EndTime. Emit the wall-clock date, exactly as `fakeLocalAsUtcDate`
 *  already does for all-day Start/End.
 *
 *  Keyed on the value, not on `allDay` or the protocol version: a DATE is
 *  floating whichever version is in play, and the arithmetic is wrong on
 *  ≤14.x too, where the spec rule does not reach. */
function untilFor(until) {
  return until?.isDate ? fakeLocalAsUtcFromValue(until) : toBasicUtc(until);
}

/** Midnight (local) of an all-day date in the event's source zone,
 *  expressed as UTC — the form AS ≤14.x expects for all-day Start/End
 *  ([MS-ASCAL] §2.2.2.39). Mirrors the inbound all-day reader so the
 *  round-trip is stable for non-UTC zones. Falls back to treating the
 *  date as UTC midnight when the zone isn't in the loaded set. */
function allDayMidnightUtc(prop, sourceTzid) {
  const v = prop?.getFirstValue();
  if (!(v instanceof ICAL.Time)) return nowBasicUtc();
  const zone =
    sourceTzid && sourceTzid !== "UTC" ? getIcalTimezone(sourceTzid) : null;
  const localMidnight = new ICAL.Time({
    year: v.year,
    month: v.month,
    day: v.day,
    hour: 0,
    minute: 0,
    second: 0,
    isDate: false,
  });
  localMidnight.zone = zone ?? ICAL.Timezone.utcTimezone;
  return formatBasicUtc(localMidnight.toJSDate());
}

function startTimeFor(dtstart, asVersion, allDay, sourceTzid) {
  if (allDay) {
    // 16.1 uses the "fake local as UTC" form; ≤14.x expects local
    // midnight in the blob's zone expressed as UTC. A floating date-only
    // value would otherwise be turned into UTC via the host's local zone
    // (toJSDate), shifting the date by ±1 day for non-UTC users.
    if (asVersion === "16.1") return fakeLocalAsUtcDate(dtstart);
    return allDayMidnightUtc(dtstart, sourceTzid);
  }
  return dtstart ? toBasicUtc(dtstart.getFirstValue()) : nowBasicUtc();
}

function endTimeFor(dtend, asVersion, allDay, sourceTzid) {
  if (allDay) {
    if (asVersion === "16.1") return fakeLocalAsUtcDate(dtend);
    return allDayMidnightUtc(dtend, sourceTzid);
  }
  return dtend ? toBasicUtc(dtend.getFirstValue()) : nowBasicUtc();
}

function isAllDayProp(prop) {
  if (!prop) return false;
  const v = prop.getFirstValue();
  return v instanceof ICAL.Time && v.isDate;
}

/* ── Helpers: embedded <Exceptions> round-trip ─────────────────────── */

/** Inbound: parse `<Exceptions><Exception>` children of `adNode`. For
 *  each, read the ORIGINAL occurrence date and either add an EXDATE to
 *  the master (`Deleted=1`) or build an override VEVENT keyed by
 *  RECURRENCE-ID. Mirrors legacy `setItemRecurrence` at
 *  sync.js:1344-1372. */
function appendInboundExceptions({
  adNode,
  vcal,
  vevent,
  asVersion,
  defaultTimezone,
}) {
  const wrapper = childByTag(adNode, "Exceptions");
  if (!wrapper) return;
  const masterUid = stringOf(vevent.getFirstPropertyValue("uid"));

  for (const exc of wrapper.children) {
    if (exc.tagName !== "Exception") continue;
    // Which occurrence this exception overrides: 16.x identifies it with
    // AirSyncBase <InstanceId>, 2.5/14.x with Calendar
    // <ExceptionStartTime>. Both are UTC.
    const startStr =
      readPathFrom(exc, ["InstanceId"]) ||
      readPathFrom(exc, ["ExceptionStartTime"]);
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
    populateVeventFromAd({
      adNode: exc,
      vevent: override,
      asVersion,
      defaultTimezone,
    });
  }
}

/** Outbound: emit a `<Exceptions>` wrapper from the VCALENDAR's EXDATEs
 *  on the master plus any sibling override VEVENTs (subcomponents that
 *  share the master's UID and carry RECURRENCE-ID). 2.5/14.x only -
 *  16.1 sends each exception as its own `<Change>` at the runner level.
 *  Mirrors legacy `getItemRecurrence` at sync.js:1488-1505. */
function appendOutboundExceptions({
  builder,
  vcal,
  vevent,
  asVersion,
  defaultTimezone,
  syncRecurrence,
  userEmail,
  fallbackOrganizerName,
  eventLog,
}) {
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
      builder,
      ical: override,
      asVersion,
      defaultTimezone,
      syncRecurrence,
      isException: true,
      userEmail,
      fallbackOrganizerName,
      eventLog,
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
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    hour: d.getUTCHours(),
    minute: d.getUTCMinutes(),
    second: d.getUTCSeconds(),
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
  const dur = new ICAL.Duration({
    minutes: Math.abs(minutesBeforeStart),
    isNegative: minutesBeforeStart > 0,
  });
  trig.setValue(dur);
  alarm.addProperty(trig);
  vevent.addSubcomponent(alarm);
}

function alarmMinutes(alarm, dtstartProp, eventLog) {
  const trig = alarm.getFirstProperty("trigger");
  if (!trig) return null;
  const v = trig.getFirstValue();
  let minutes;
  let wasAbsolute = false;
  if (v instanceof ICAL.Duration) {
    minutes = Math.round(-v.toSeconds() / 60); // EAS minutes = before start, positive
  } else if (v instanceof ICAL.Time && dtstartProp) {
    const start = dtstartProp.getFirstValue();
    if (!(start instanceof ICAL.Time)) return null;
    minutes = Math.round(start.subtractDateTz(v).toSeconds() / 60);
    wasAbsolute = true;
  } else {
    return null;
  }
  if (eventLog) {
    if (wasAbsolute) {
      eventLog(
        "info",
        `[calendar-sync] converted absolute VALARM trigger to relative offset (${minutes} min before start) - EAS supports relative alarms only`,
      );
    }
    if (minutes < 0) {
      eventLog(
        "info",
        "[calendar-sync] dropped VALARM scheduled after event start - EAS supports alarms before start only",
      );
    }
  }
  return minutes;
}

/* ── Helpers: attendees ────────────────────────────────────────────── */

function collectAttendees(adNode, userEmail, fallbackResponseType) {
  const out = [];
  const wrapper = childByTag(adNode, "Attendees");
  if (!wrapper) return out;
  const userEmailLower = userEmail ? String(userEmail).toLowerCase() : null;
  for (const a of wrapper.children) {
    if (a.tagName !== "Attendee") continue;
    const email = readPathFrom(a, ["Email"]);
    if (!email) continue;
    const item = { email, cn: readPathFrom(a, ["Name"]) };
    const status = readPathFrom(a, ["AttendeeStatus"]);
    const isSelf =
      userEmailLower && email.toLowerCase() === userEmailLower;
    if (status) {
      item.partstat = ATTENDEESTATUS_TO_PARTSTAT[status] ?? "NEEDS-ACTION";
    } else if (isSelf && fallbackResponseType) {
      // Legacy calendarsync.js:203-204: when AttendeeStatus is missing
      // for the self-attendee, fall back to the event-level
      // ResponseType.
      item.partstat =
        ATTENDEESTATUS_TO_PARTSTAT[fallbackResponseType] ?? "NEEDS-ACTION";
    } else {
      // Legacy line 206: explicit default for missing status.
      item.partstat = "NEEDS-ACTION";
    }
    const type = readPathFrom(a, ["AttendeeType"]);
    if (type === "1") {
      item.role = "REQ-PARTICIPANT";
      item.cutype = "INDIVIDUAL";
    } else if (type === "2") {
      item.role = "OPT-PARTICIPANT";
      item.cutype = "INDIVIDUAL";
    } else if (type === "3") {
      item.role = "NON-PARTICIPANT";
      item.cutype = "RESOURCE";
    }
    out.push(item);
  }
  return out;
}

function stripMailto(s) {
  if (!s) return "";
  return String(s).replace(/^mailto:/i, "");
}

/* ── Helpers: recurrence ───────────────────────────────────────────── */

function recurrenceToRrule(recNode) {
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
    const ical = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
    for (let i = 0; i < 7; i++) if (bits & (1 << i)) days.push(ical[i]);
    if (days.length) {
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
  if (until) parts.push(`UNTIL=${until.replace(/[-:]/g, "")}`);

  return parts.join(";");
}

function appendRecurrence(builder, rruleProp, dtstartProp) {
  const r = rruleProp.getFirstValue(); // ICAL.Recur
  if (!r) return;

  const freq = r.freq;
  const startDate = dtstartProp?.getFirstValue();
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
    const dow = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"].indexOf(m[2]) + 1;
    weekDays[i] = dow;
    if (n) weeks[i] = n === -1 ? 5 : n;
  }

  if (freq === "WEEKLY") {
    type = 1;
    if (!weekDays.length && startDate)
      weekDays = [startDate.dayOfWeek?.() ?? 1];
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
  else if (r.until) builder.atag("Until", untilFor(r.until));
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
  if (asVersion !== "16.1")
    builder.atag("EstimatedDataSize", String((desc ?? "").length));
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
  try {
    return new ICAL.Component(ICAL.parse(ical));
  } catch {
    return null;
  }
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
