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
 * (outbound on 2.5/14.x only). Per-instance: one `<Change>` or `<Delete>`
 * per occurrence carrying `<InstanceId>`, handled by `applyInstanceChange`
 * / `applyInstanceDelete` and built by `listInstanceCommands` (16.1
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
import { TimeZoneBlob, isAllZero } from "./timezone-blob.mjs";
import {
  guessTimezoneByStdDstOffset,
  tzInfoForBlob,
  getIcalTimezone,
} from "./timezone-mapping.mjs";

const X_EAS_SERVERID = "X-EAS-SERVERID";
const X_EAS_RESPONSETYPE = "X-EAS-RESPONSETYPE";
const X_EAS_MEETINGSTATUS = "X-EAS-MEETINGSTATUS";
/** The answer we last mailed to the organiser.
 *
 *  Below 16.0 the client sends that mail, and the server never records the
 *  answer in its own ResponseType - it stays 5, "not responded", however
 *  many times we push an AttendeeStatus. So "does the server already know
 *  this?" cannot be the test there, and without one every later edit to an
 *  answered meeting mails the organiser again. This is our own record of
 *  what they have been told. It is an X-EAS-* property, so the stamp guard
 *  keeps outside writers off it and an adopt carries it across. */
const X_EAS_REPLIED = "X-EAS-REPLIED";

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

// EAS AttendeeStatus → iCal PARTSTAT ([MS-ASCAL] 2.2.2.5). 1 is not a
// value this element takes.
const ATTENDEESTATUS_TO_PARTSTAT = {
  0: "NEEDS-ACTION", // response unknown
  2: "TENTATIVE",
  3: "ACCEPTED",
  4: "DECLINED",
  5: "NEEDS-ACTION", // not responded
};

// EAS ResponseType → iCal PARTSTAT ([MS-ASCAL] 2.2.2.40). A separate table
// on purpose: the two enums agree on 2, 3 and 4 and part company either
// side of them, so sharing one made 5 - "the user has not yet responded" -
// read as an acceptance, and wrote "you accepted" into the calendar for
// every invitation nobody had answered.
const RESPONSETYPE_TO_PARTSTAT = {
  0: "NEEDS-ACTION", // the user's response has not yet been received
  1: "ACCEPTED", // the user is the organizer; no reply is required
  2: "TENTATIVE",
  3: "ACCEPTED",
  4: "DECLINED",
  5: "NEEDS-ACTION", // not responded
};

/* ── Reader: ApplicationData → iCal VEVENT ─────────────────────────── */

export async function applicationDataToIcal({
  adNode,
  existingIcal,
  serverID,
  asVersion,
  defaultTimezone,
  syncRecurrence,
  uid,
  userEmail,
  eventLog,
  nativePlainText = null,
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
      "debug",
      `[calendar-codec] receive OrganizerInfo: present=${!!hasOrgInfo} OrganizerEmail=${JSON.stringify(orgEmailRaw ?? null)} OrganizerName=${JSON.stringify(orgNameRaw ?? null)}`,
    );
  }

  await populateVeventFromAd({
    adNode,
    vevent,
    asVersion,
    defaultTimezone,
    userEmail,
    nativePlainText,
  });

  // Recurrence + 2.5/14.x exceptions. Gated on the account-level
  // syncRecurrence flag. Only touch RRULE / exceptions when the AD
  // mentions them; otherwise leave whatever the existing blob carried.
  if (syncRecurrence) {
    const recNode = childByTag(adNode, "Recurrence");
    if (recNode) {
      vevent.removeAllProperties("rrule");
      const rrule = easToRrule(recNode);
      if (rrule && /^FREQ=[A-Z]+/.test(rrule)) {
        const prop = new ICAL.Property("rrule", vevent);
        prop.setValue(ICAL.Recur.fromString(rrule));
        vevent.addProperty(prop);
      }
      keepUnmappedRecurrence(vevent, recNode);
    }
    if (childByTag(adNode, "Exceptions")) {
      // Clear the whole existing exception set so the AD's replaces it: the
      // override vevents (anything with RECURRENCE-ID) *and* the master's
      // EXDATEs, which are how a cancelled occurrence is stored. Clearing
      // only the overrides made every re-delivery of the series add another
      // copy of each EXDATE, and `listInstanceCommands` then emitted a
      // duplicate <Delete> for each copy.
      //
      // It also drops an EXDATE the server no longer has - an occurrence
      // un-cancelled elsewhere - which nothing else on the inbound path
      // would ever remove.
      for (const sub of vcal.getAllSubcomponents("vevent")) {
        if (sub.getFirstProperty("recurrence-id")) {
          vcal.removeSubcomponent(sub);
        }
      }
      vevent.removeAllProperties("exdate");
      await appendInboundExceptions({
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
async function populateVeventFromAd({
  adNode,
  vevent,
  asVersion,
  defaultTimezone,
  // For an exception body: the master's AllDayEvent value. [MS-ASCAL]
  // §2.2.2.1 - an Exception without its own AllDayEvent "is assumed to
  // be the same as the value of the top-level AllDayEvent element", and
  // Exchange 16.1 does omit it on embedded exceptions. Reading the
  // absence as 0 turned an all-day override into a timed one, whose
  // midnight-UTC DTSTART then failed to match anything.
  inheritedAllDay = false,
  userEmail,
  nativePlainText = null,
}) {
  // Subject / Location.
  const subject = readPathFrom(adNode, ["Subject"]);
  if (subject) vevent.updatePropertyWithValue("summary", subject);

  const locDisplay =
    readPathFrom(adNode, ["Location", "DisplayName"]) ??
    readPathFrom(adNode, ["Location"]);
  if (locDisplay) vevent.updatePropertyWithValue("location", locDisplay);

  // Body (codepage-aware; AirSyncBase ≥14.x).
  await readBodyIntoDescription(vevent, adNode, {
    useAirSyncBase: useAirSyncBaseBody(asVersion),
    nativePlainText,
  });

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
  // then. Exchange ≤14.x encodes an all-day boundary as midnight-in-zone
  // expressed as UTC, so recovering the calendar date from
  // `20230831T220000Z` needs the zone; it is then discarded. An all-day
  // event on AS 16.1 carries no blob and needs no zone, because its wire
  // value is already the date. `writeDateProp` picks between the two on
  // the shape of the value - see there for why the blob cannot decide it.
  const { tzId } = resolveTimezone(adNode, defaultTimezone);
  const ownAllDay = readPathFrom(adNode, ["AllDayEvent"]);
  const allDay = ownAllDay == null ? inheritedAllDay : ownAllDay === "1";

  // Start / End. EAS sends UTC strings; convert on the way in.
  const startUtc = readPathFrom(adNode, ["StartTime"]);
  const endUtc = readPathFrom(adNode, ["EndTime"]);
  if (startUtc) writeDateProp(vevent, "dtstart", startUtc, tzId, allDay);
  if (endUtc) {
    writeDateProp(vevent, "dtend", endUtc, tzId, allDay);
    // An externally-authored blob may express its end as DURATION.
    // RFC 5545 forbids carrying both, and DTEND now holds the truth.
    vevent.removeAllProperties("duration");
  }

  // DtStamp - preserve when present. A 16.x *client* MUST NOT send it
  // ([MS-ASCAL] §2.2.2.18), which is why the writer omits it there, but
  // the server does send it on 16.1 - observed on every item of an
  // Exchange Online initial sync.
  const dtStamp = readPathFrom(adNode, ["DtStamp"]);
  if (dtStamp) writeDateProp(vevent, "dtstamp", dtStamp, "UTC", false);

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
export async function applyInstanceChange({
  ical,
  adNode,
  instanceUtc,
  asVersion,
  defaultTimezone,
  userEmail,
}) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return ical;
  // The master is the vevent without a RECURRENCE-ID, not simply the first
  // one - same rule, and same reason, as ad219dc: for a blob whose overrides
  // happen to come first this would otherwise strip an EXDATE from an
  // override.
  const master = pickMasterVevent(vcal);
  if (!master) return ical;
  // An all-day master's exceptions are DATEs, like its DTSTART - RFC 5545
  // §3.8.4.4 matches RECURRENCE-ID to the occurrence by value type, so a
  // DATE-TIME one binds nothing. A 16.1 all-day InstanceId is the
  // fake-local form (date at T000000Z), so no zone is needed to derive
  // the date.
  const allDay = !!master.getFirstPropertyValue("dtstart")?.isDate;

  removeExdate(master, instanceUtc);

  // Replace any existing override for this RECURRENCE-ID - but keep it as
  // the seed: a 16.1 per-instance <Change> is a delta against it and
  // omits every field that did not change.
  let existing = null;
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    const rid = sub.getFirstPropertyValue("recurrence-id");
    if (rid && namesInstance(rid, instanceUtc, null)) {
      if (!existing) existing = sub;
      vcal.removeSubcomponent(sub);
    }
  }

  const ridTime = allDay
    ? allDayIcalDate(instanceUtc, null)
    : instanceUtcToIcalTime(instanceUtc);
  const override = seedOverride({ master, existing, ridTime });
  vcal.addSubcomponent(override);
  const masterUid = stringOf(master.getFirstPropertyValue("uid"));
  if (masterUid) override.updatePropertyWithValue("uid", masterUid);
  // RECURRENCE-ID anchors the override to the original master occurrence.
  const ridProp = new ICAL.Property("recurrence-id", override);
  ridProp.setValue(ridTime);
  override.addProperty(ridProp);
  await populateVeventFromAd({
    adNode: adNode,
    vevent: override,
    asVersion,
    defaultTimezone,
    inheritedAllDay: allDay,
    userEmail,
  });
  return vcal.toString();
}

/** Outbound 16.1: one command per current EXDATE / RECURRENCE-ID override
 *  on the master, each naming the master's ServerId and the occurrence's
 *  InstanceId - `<Delete>` for a cancelled occurrence, `<Change>` for a
 *  moved one. Idempotent - re-asserts the full exception set on every push
 *  of a recurring master.
 *
 *  Returns descriptors rather than writing them, because Exchange will not
 *  take two commands against one ServerId in a single request: it applies
 *  the first, faults on the second and discards the whole response with a
 *  global Status 16. Batching is therefore the caller's decision, and the
 *  caller needs the InstanceId to name the occurrence in a log line.
 *
 *  Each descriptor's `emit(builder)` writes one complete command. The
 *  builder must be on the AirSync codepage on entry; it is left there.
 *
 *  `previous` is the exception fingerprint the item carried before the
 *  user's edit, taken from the changelog entry. When supplied, only
 *  exceptions that actually differ from it produce a command: re-asserting
 *  a cancellation Exchange already has earns a rejection and a round of
 *  retries, and a series whose master changed would otherwise re-send every
 *  occurrence it has. Without it - a newly added item, or an edit recorded
 *  before the entry carried one - everything is emitted, which is the older
 *  behaviour and always safe.
 *
 *  Limitation: a user un-deleting an EXDATE or removing an override still
 *  cannot be expressed in EAS. `previous` now tells us it happened, but
 *  there is no command for "make this occurrence ordinary again"; the
 *  unwanted EXDATE / override stays server-side until re-edited there.
 *
 *  @returns {Array<{kind: string, serverID: string, instanceId: string,
 *                   emit: (builder: object) => void}>}
 */
export function listInstanceCommands({
  blob,
  serverID,
  asVersion,
  defaultTimezone,
  syncRecurrence,
  userEmail,
  fallbackOrganizerName,
  eventLog,
  previous = null,
}) {
  if (asVersion !== "16.1") return [];
  const vcal = parseVCalendar(blob);
  if (!vcal) return [];
  const master = vcal.getFirstSubcomponent("vevent") ?? null;
  // parseVCalendar's first vevent may be an override if iCal order is
  // unusual; reuse the master picker instead.
  const masterVevent = pickMasterVevent(vcal) ?? master;
  if (!masterVevent) return [];

  const masterUid = stringOf(masterVevent.getFirstPropertyValue("uid"));
  const exdates = collectExdates(masterVevent);
  const overrides = [];
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    if (sub === masterVevent) continue;
    const subUid = stringOf(sub.getFirstPropertyValue("uid"));
    const rid = sub.getFirstProperty("recurrence-id");
    if (subUid === masterUid && rid) overrides.push(sub);
  }

  // What the server already had, if the caller knew. Same shape and same
  // digest function as exceptionFingerprint, so the two are comparable.
  const knownExdates = new Set(previous?.exdates ?? []);
  const knownOverrides = new Map(
    (previous?.overrides ?? []).map((o) => [o.rid, o.digest]),
  );

  const commands = [];
  // A cancelled occurrence is a <Delete>, not a <Change> carrying
  // <Deleted>. [MS-ASCAL] §2.2.2.16 allows `Deleted` only as a child of
  // `Exception`, which lives in the embedded <Exceptions> block that 16.x
  // replaced - there is no legal place for it here, and Exchange rejects
  // the command with Status 6. [MS-ASCMD] Delete gives the 16.x form:
  // ServerId plus airsyncbase:InstanceId, no ApplicationData at all,
  // because "the object is identified by both the ServerId element of the
  // master item as well as the airsyncbase:InstanceId element of the
  // specific occurrence".
  for (const ex of exdates) {
    const instanceId = instanceKey(ex);
    // Already cancelled server-side. Sending it again is rejected, and the
    // rejection costs a retry budget and shows up as a failed element.
    if (previous && knownExdates.has(instanceId)) continue;
    commands.push({
      kind: "delete",
      serverID,
      instanceId,
      emit(builder) {
        builder.otag("Delete");
        builder.atag("ServerId", serverID);
        builder.switchpage("AirSyncBase");
        builder.atag("InstanceId", instanceId);
        builder.switchpage("AirSync");
        builder.ctag();
      },
    });
  }
  for (const override of overrides) {
    const rid = override.getFirstPropertyValue("recurrence-id");
    const instanceId = instanceKey(rid);
    // Unchanged since the server last saw it: the whole point of carrying a
    // baseline. An override is pushed whole, so an identical digest means
    // there is nothing to say about this occurrence.
    if (
      previous &&
      knownOverrides.get(instanceId) === digestOf(override.toString())
    )
      continue;
    commands.push({
      kind: "change",
      serverID,
      instanceId,
      emit(builder) {
        builder.otag("Change");
        builder.atag("ServerId", serverID);
        // Sibling of ServerId, not part of the payload - see the EXDATE
        // branch above for the citations.
        builder.switchpage("AirSyncBase");
        builder.atag("InstanceId", instanceId);
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
      },
    });
  }
  return commands;
}

function cloneVevent(comp) {
  return new ICAL.Component(structuredClone(comp.toJSON()));
}

/** The component a rebuilt override starts from - never an empty one.
 *
 *  EAS payloads for an exception omit what did not change, and
 *  `populateVeventFromAd` is merge-aware throughout: it only touches a
 *  field whose wire element is present. Both were useless while the
 *  callers handed it a fresh empty VEVENT - every omitted field then
 *  came out blank, which is how a status-only delta from Exchange
 *  emptied an override's title and times (#342), and an override left
 *  without DTSTART later serialised as a malformed change the server
 *  rejected.
 *
 *  `existing` is the prior override when the wire is a delta against it
 *  (the 16.1 per-instance <Change>); null on the embedded ≤14.x path,
 *  where [MS-ASCAL] §2.2.2.21 defines absence as "same as the top-level
 *  element" - inheritance from the master. The master seed drops the
 *  series-defining properties and anchors the occurrence's own
 *  boundaries: its start is the instant the RECURRENCE-ID names, its end
 *  that plus the master's duration. */
function seedOverride({ master, existing, ridTime }) {
  if (existing) {
    const base = cloneVevent(existing);
    // The caller writes a fresh RECURRENCE-ID; the clone's would sit
    // beside it as a duplicate property.
    base.removeAllProperties("recurrence-id");
    return base;
  }
  const base = cloneVevent(master);
  base.removeAllProperties("rrule");
  base.removeAllProperties("exdate");
  base.removeAllProperties("recurrence-id");
  const mStart = master.getFirstPropertyValue("dtstart");
  // Through `eventTimingFor`, not a raw DTEND read - a DURATION master
  // must seed its overrides with the derived end.
  const mEnd = itemDateValue(eventTimingFor(master).end);
  if (ridTime && mStart instanceof ICAL.Time) {
    const start = ridTime.clone();
    base.removeAllProperties("dtstart");
    const ds = new ICAL.Property("dtstart", base);
    ds.setValue(start);
    base.addProperty(ds);
    if (mEnd instanceof ICAL.Time) {
      const end = ridTime.clone();
      end.addDuration(mEnd.subtractDate(mStart));
      base.removeAllProperties("dtend");
      // The clone may carry the master's DURATION; the override's end is
      // now this explicit DTEND, and iCal forbids holding both.
      base.removeAllProperties("duration");
      const de = new ICAL.Property("dtend", base);
      de.setValue(end);
      base.addProperty(de);
    }
  }
  return base;
}

function pickMasterVevent(vcal) {
  const all = vcal.getAllSubcomponents("vevent");
  for (const v of all) if (!v.getFirstProperty("recurrence-id")) return v;
  return all[0] ?? null;
}

/** The component our stamps belong on, for a blob that may hold either kind.
 *
 *  One calendar type serves events and tasks, so anything reached through
 *  the item hooks can be a VTODO - and a VTODO has no overrides in EAS, so
 *  the first one is the only one. Without this `pinEasStamps` stripped a
 *  task's stamps and then found no component to restore them onto, which
 *  deleted `X-EAS-SERVERID` from every task the user edited. */
function pickMasterItem(vcal) {
  return pickMasterVevent(vcal) ?? vcal.getFirstSubcomponent("vtodo");
}

/** Add an EXDATE to the master VEVENT (16.1 InstanceId-with-Deleted=1). */
export function applyInstanceDelete({ ical, instanceUtc }) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return ical;
  // The vevent without a RECURRENCE-ID - see applyInstanceChange above.
  const master = pickMasterVevent(vcal);
  if (!master) return ical;
  const allDay = !!master.getFirstPropertyValue("dtstart")?.isDate;

  // Drop any existing override at this RECURRENCE-ID - server says it's
  // gone now.
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    const rid = sub.getFirstPropertyValue("recurrence-id");
    if (rid && namesInstance(rid, instanceUtc, null)) {
      vcal.removeSubcomponent(sub);
    }
  }
  addExdate(master, instanceUtc, allDay);
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
  // ≤14.x only: leave the embedded <Exceptions> wrapper out of this
  // payload. Used by an <Add> whose blob carries exceptions - they are
  // sent as a follow-up <Change> once the server has assigned a
  // ServerId, mirroring the shape 16.1 forces for every exception. At
  // least one server family answers Status 1 to an Add-with-exceptions
  // and then silently discards the wrapper (and, for an all-day series,
  // the recurrence with it), so nothing may ride on the Add.
  suppressExceptions = false,
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

  const { dtstart, end, allDay } = eventTimingFor(vevent);

  // Outbound timezone. [MS-ASCAL] §2.2.2.44 lists the element as
  // supported in every protocol version, 16.1 included, and states no
  // restriction on client use - unlike UID and DtStamp, where the same
  // document says outright that a 16.x client MUST NOT send them.
  // Exchange Online sends it on 16.1 too, on every timed event.
  //
  // Two cases get none:
  //  - an exception, in every version. A TimeZone describes the series,
  //    not one occurrence: [MS-ASCAL] §2.2.2.21 lists what an Exception
  //    may contain and TimeZone is not on it. On ≤14.x that is moot,
  //    since the embedded <Exception> sits inside the master's payload
  //    and inherits its blob. On 16.1 the instance is its own top-level
  //    <Change> (see `listInstanceCommands`), and the reasoning here used
  //    to be that such a command therefore needs a blob of its own -
  //    plausible, and wrong. Exchange rejects the command outright with
  //    Status 6, "malformed or invalid item", which [MS-ASCMD] marks not
  //    transient and tells the client to stop sending. Removing this one
  //    element is the whole difference between test 3.4 failing and passing
  //    (`test/test_3_recurrence.py`); the payloads are otherwise
  //    byte-identical to the ones the server accepts when the same
  //    exceptions are first created.
  //  - an all-day event on 16.1. [MS-ASCAL] §2.2.2.1 is explicit: with
  //    AllDayEvent set to 1 the client "MUST NOT include the TimeZone
  //    element", and the server will not send one either - "a client
  //    SHOULD interpret this event to be at the given date(s) regardless
  //    of the time zone used". That is the floating semantics iCalendar
  //    gives a DATE value, so the two formats agree.
  //
  // An all-day event on ≤14.x DOES get a blob - but a UTC one, whatever
  // zone the user sits in. Its boundaries go out date-shaped (see
  // `startTimeFor`), and `date + UTC blob` names the same calendar day
  // under both reading disciplines in the wild: a blob-reading server
  // computes midnight UTC → that date, a date-digit reader takes the
  // digits → that date. A user-zone blob is what used to make the two
  // readers disagree.
  const floatingAllDay = allDay && asVersion === "16.1";
  if (!isException && !floatingAllDay) {
    const blob = allDay
      ? buildUtcTimezoneBlob()
      : buildTimezoneBlob(vevent, defaultTimezone);
    builder.atag("TimeZone", blob.easTimeZone64);
  }

  builder.atag("AllDayEvent", allDay ? "1" : "0");

  // Body.
  appendBodyFromDescription(builder, vevent, asVersion, "Calendar");

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
        "debug",
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

  // EndTime. All-day boundaries are date-shaped in every version - see
  // `startTimeFor` for why.
  builder.atag("EndTime", endTimeFor(end, asVersion, allDay));

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
      // the organizer iff the ORGANIZER email differs from this
      // account's own address. `userEmail` is that address as the
      // server named it, so an account whose login is not an email
      // still labels its own meetings correctly; with no address at all
      // every attendee'd event reads as received.
      const orgEmail = orgProp
        ? stripMailto(orgProp.getFirstValue()).toLowerCase()
        : "";
      const userEmailLower = userEmail ? String(userEmail).toLowerCase() : "";
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
          // The attendee's own answer. [MS-ASCAL] 2.2.2.5: a request
          // element on 12.0 through 14.1, and "the client MUST NOT include
          // the AttendeeStatus element in a command request when protocol
          // version 16.0 or 16.1 is used" - there it is MeetingResponse's
          // job. Below 16.0 this is how a user answers at all, and the only
          // way they can change an answer they have already given.
          if (!asVersion.startsWith("16.")) {
            const partstat = String(
              a.getParameter("partstat") ?? "",
            ).toUpperCase();
            const status = PARTSTAT_TO_ATTENDEESTATUS[partstat];
            if (status) builder.atag("AttendeeStatus", status);
          }
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
    if (rrule) {
      appendRecurrence(
        builder,
        rrule,
        dtstart,
        unmappedRecurrenceOf(vevent),
        firstDayOfWeekOf(vevent, rrule),
      );
    }
    if (asVersion !== "16.1" && !suppressExceptions) {
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

/** Names why an event cannot go on the wire without changing its
 *  meaning, or null when it can. The EAS Recurrence Type enum has no
 *  sub-daily frequencies, so an HOURLY/MINUTELY/SECONDLY rule would
 *  silently sync as DAILY - the runner holds such an item as a
 *  client-side rejection instead (warned, counted, retried every sync
 *  until the user changes or removes it). Gated on `syncRecurrence`:
 *  with the flag off no recurrence is emitted at all, so nothing is
 *  misrepresented. The regex pre-filter is safe - VTIMEZONE transition
 *  rules are only ever yearly. */
const SUB_DAILY_FREQ = /FREQ=(HOURLY|MINUTELY|SECONDLY)/;

export function clientRejectReason({ blob, syncRecurrence }) {
  if (typeof blob !== "string") return null;
  const vcal = parseVCalendar(blob);
  const master = vcal ? pickMasterVevent(vcal) : null;
  if (!master) return null;

  if (syncRecurrence && SUB_DAILY_FREQ.test(blob)) {
    const freq = String(
      master.getFirstProperty("rrule")?.getFirstValue()?.freq ?? "",
    ).toUpperCase();
    if (freq === "HOURLY" || freq === "MINUTELY" || freq === "SECONDLY") {
      return `EAS cannot represent a recurrence below daily (FREQ=${freq})`;
    }
  }

  // Timing, for the master and for every override - an occurrence rides
  // the same writer, so a bad one lands on the wire the same way. An
  // event with a start but no expressed end, or an end not after its
  // start, cannot go out without inventing data, so it is held like a
  // server rejection. A VEVENT with no DTSTART at all is left alone:
  // that is the status-only exception-delta shape, which is legitimate.
  for (const comp of vcal.getAllSubcomponents("vevent")) {
    const startVal = comp.getFirstPropertyValue("dtstart");
    if (!(startVal instanceof ICAL.Time)) continue;
    const where =
      comp === master
        ? "the event"
        : `the occurrence on ${comp.getFirstPropertyValue("recurrence-id")}`;
    const endVal = itemDateValue(eventTimingFor(comp).end);
    if (!(endVal instanceof ICAL.Time)) {
      return `${where} has a start but no end (neither DTEND nor DURATION)`;
    }
    if (endVal.compare(startVal) <= 0) {
      return `${where} does not end after it starts`;
    }
  }
  return null;
}

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

/** Resolve the effective IANA tzid for an inbound event, falling back to
 *  the host's default zone when the item carries no usable blob (all-day
 *  events on AS 16.1, and Z-Push/Kopano/Grommunio, which send an all-zero
 *  one). Keyed on the item rather than the version: Exchange sends a blob
 *  on 16.1 for timed events.
 *
 *  `fromBlob` reports whether a real blob supplied the zone. It is no
 *  longer what selects the all-day boundary encoding — Exchange 14.1 drops
 *  the TimeZone element when echoing an all-day item back, so the flag
 *  says "16.1-style" for a value that is not. `writeDateProp` decides on
 *  the value instead. */
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
  const tzid = guessTimezoneByStdDstOffset(
    stdOffset,
    dstOffset,
    stdName,
    dstRule,
  );
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
    std: {
      month: std.wMonth,
      weekOfMonth: std.wDay,
      dayOfWeek: std.wDayOfWeek,
    },
    dst: {
      month: dst.wMonth,
      weekOfMonth: dst.wDay,
      dayOfWeek: dst.wDayOfWeek,
    },
  };
}

function buildTimezoneBlob(vevent, defaultTimezone) {
  const sourceTzid = pickSourceTzid(vevent) ?? defaultTimezone ?? "UTC";
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

/** The blob for an all-day item on ≤14.x: all zeros - bias 0, no DST, no
 *  names. That is UTC, and it is also the exact form the Z-Push family
 *  itself emits for all-day, so both reading disciplines resolve the
 *  date-shaped boundaries to the same calendar day. Needs nothing from
 *  the timezone mapping, deliberately: an all-day date has no zone. */
function buildUtcTimezoneBlob() {
  return new TimeZoneBlob();
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

/** The calendar date of an all-day boundary, from the UTC instant EAS put
 *  on the wire. The encoding depends on the server, and the two encodings
 *  in the wild are told apart by the value itself:
 *
 *  - "fake local as UTC", YYYYMMDDT000000Z. The UTC date already IS
 *    the intended date. All-day events on AS 16.1 use it (§2.2.2.1
 *    forbids a TimeZone element there), as do Z-Push/Kopano/Grommunio
 *    (all-zero blob). It mirrors what `fakeLocalAsUtcDate` emits
 *    outbound. Read verbatim - converting would shift the date by a
 *    day for users west of UTC.
 *
 *  - midnight-in-zone expressed as UTC, e.g. 20261006T220000Z for the
 *    7th in Europe/Berlin. Real Exchange ≤14.x stores this form, so it
 *    arrives with items other clients created.
 *    For a non-UTC zone that instant lands on the previous/next UTC
 *    calendar day, so getUTCDate() shifts the event by ±1 day; the
 *    value has to be converted into the zone first.
 *
 *  A midnight-UTC time-of-day is the discriminator, NOT the presence of
 *  a real TimeZone blob. Exchange 14.1 was measured echoing an all-day
 *  item back with StartTime 20261006T220000Z and no TimeZone element at
 *  all, having accepted exactly that value from us moments earlier: a
 *  blob-keyed test reads such an echo verbatim and moves every all-day
 *  event one day earlier on the next pull. The value cannot lie the same
 *  way - 220000Z is not a date-shaped value under any encoding.
 *
 *  Every all-day boundary goes through this one derivation - DTSTART and
 *  DTEND here, RECURRENCE-ID and EXDATE in the exception readers - so an
 *  exception lands on the same calendar date as the occurrence it names. */
function allDayDateFromUtcInstant(d, tzId) {
  let year = d.getUTCFullYear();
  let month = d.getUTCMonth() + 1;
  let day = d.getUTCDate();
  const midnightUtc =
    d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0;
  if (!midnightUtc && tzId && tzId !== "UTC") {
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
  return { year, month, day };
}

/** Same, as a DATE-valued ICAL.Time - the stored form of every all-day
 *  property (RFC 5545 §3.8.4.4: a RECURRENCE-ID must match the value type
 *  of the DTSTART it binds to, so an all-day master's exceptions are DATEs
 *  or they bind nothing). */
function allDayIcalDate(d, tzId) {
  return new ICAL.Time({ ...allDayDateFromUtcInstant(d, tzId), isDate: true });
}

function writeDateProp(vevent, name, easUtc, tzId, allDay) {
  // Replace any existing property of this name so merge-mode partial
  // Changes don't duplicate dtstart/dtend/dtstamp on the master.
  vevent.removeAllProperties(name);
  const prop = new ICAL.Property(name, vevent);
  if (allDay) {
    const d = parseEasUtc(easUtc);
    if (!d) return;
    prop.setValue(allDayIcalDate(d, tzId));
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
  // Accept extended ISO and basic compact forms. The fraction goes with
  // the separators: both are what distinguishes extended from compact, and
  // EAS's extended form carries milliseconds (2026-10-01T00:00:00.000Z).
  // Stripping only the separators left a fraction the pattern below cannot
  // match, so a value this function says it accepts returned null.
  const compact = s.replace(/[-:]|\.\d+/g, "");
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

/** A boundary reaches the writer either as an iCal property or as an
 *  `ICAL.Time` already derived from one (DTSTART+DURATION); unwrap to
 *  the value either way. */
function itemDateValue(propOrValue) {
  return propOrValue?.getFirstValue ? propOrValue.getFirstValue() : propOrValue;
}

/** Read a date as `YYYYMMDDT000000Z` from the *local-clock*
 *  year/month/day, with no UTC conversion. Mirrors legacy
 *  `getIsoUtcString(date, false, true, true)` for AS 16.1 all-day. */
function fakeLocalAsUtcDate(propOrValue) {
  const v = itemDateValue(propOrValue);
  if (!v) return nowBasicUtc();
  return fakeLocalAsUtcFromValue(v);
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

function startTimeFor(dtstart, asVersion, allDay) {
  // All-day boundaries are date-shaped in EVERY version, not just 16.1
  // (where [MS-ASCAL] §2.2.2.1 mandates it). On ≤14.x the spec says only
  // that an all-day item begins "on midnight of the specified day" - in
  // no stated zone - and the two server families in the wild read the
  // value differently: real Exchange through the TimeZone blob, the
  // Z-Push family off the date digits. `YYYYMMDDT000000Z` with a UTC
  // blob (see the all-day branch in `appendApplicationDataFromIcal`)
  // lands on the same calendar date under both readings, in every user
  // zone. Midnight-in-zone - the previous ≤14.x form - shifted the whole
  // series a day on a date-digit reader whenever the zone sat east of
  // UTC. A floating date-only value must not go through toJSDate: that
  // resolves it in the host zone and shifts the date the same way.
  if (allDay) return fakeLocalAsUtcDate(dtstart);
  return dtstart ? toBasicUtc(dtstart.getFirstValue()) : nowBasicUtc();
}

function endTimeFor(end, asVersion, allDay) {
  if (allDay) return fakeLocalAsUtcDate(end);
  const v = itemDateValue(end);
  return v ? toBasicUtc(v) : nowBasicUtc();
}

function isAllDayProp(prop) {
  if (!prop) return false;
  const v = prop.getFirstValue();
  return v instanceof ICAL.Time && v.isDate;
}

/** Effective timing for the wire. RFC 5545 lets a VEVENT carry its end
 *  as DTEND or as DURATION (never both); EAS always needs an EndTime, so
 *  a DURATION event's end is DTSTART+DURATION.
 *
 *  A DATE-valued DTSTART with neither is not endless: §3.6.1 states its
 *  duration "is taken to be one day", so reading it that way is reading
 *  the item, not defaulting it. The DATE-TIME case gets no such
 *  treatment here - the spec makes it zero-length, which
 *  `clientRejectReason` holds as invalid rather than send.
 *
 *  `end` is therefore the DTEND property, a derived `ICAL.Time`, or null
 *  when a timed event expresses no end at all. All-day follows the
 *  start: a DATE DTSTART is all-day whether its end is a DATE DTEND, a
 *  DURATION, or the one-day default. */
function eventTimingFor(vevent) {
  const dtstart = vevent.getFirstProperty("dtstart");
  const dtend = vevent.getFirstProperty("dtend");
  const allDay = isAllDayProp(dtstart) && (!dtend || isAllDayProp(dtend));
  let end = dtend;
  // An all-day item whose DTEND equals its DTSTART is a day long, not
  // zero. DTEND is exclusive, so the shape is malformed - but Outlook and
  // a good many .ics exporters write it, and it means the same thing as
  // the missing DTEND below, which §3.6.1 gives a day. Reading it that way
  // is reading the item; holding it as unrepresentable would strand an
  // ordinary imported all-day event and warn about it on every sync.
  if (allDay && dtend) {
    const s = dtstart?.getFirstValue();
    const e = dtend.getFirstValue();
    if (
      s instanceof ICAL.Time &&
      e instanceof ICAL.Time &&
      e.compare(s) === 0
    ) {
      end = e.clone();
      end.addDuration(new ICAL.Duration({ days: 1 }));
    }
  }
  if (!end) {
    const startVal = dtstart?.getFirstValue();
    const dur = vevent.getFirstProperty("duration")?.getFirstValue();
    if (startVal instanceof ICAL.Time && dur instanceof ICAL.Duration) {
      end = startVal.clone();
      end.addDuration(dur);
    } else if (startVal instanceof ICAL.Time && allDay) {
      end = startVal.clone();
      end.addDuration(new ICAL.Duration({ days: 1 }));
    } else {
      end = null;
    }
  }
  return { dtstart, end, allDay };
}

/* ── Helpers: embedded <Exceptions> round-trip ─────────────────────── */

/** Inbound: parse `<Exceptions><Exception>` children of `adNode`. For
 *  each, read the ORIGINAL occurrence date and either add an EXDATE to
 *  the master (`Deleted=1`) or build an override VEVENT keyed by
 *  RECURRENCE-ID. Mirrors legacy `setItemRecurrence` at
 *  sync.js:1344-1372. */
async function appendInboundExceptions({
  adNode,
  vcal,
  vevent,
  asVersion,
  defaultTimezone,
}) {
  const wrapper = childByTag(adNode, "Exceptions");
  if (!wrapper) return;
  const masterUid = stringOf(vevent.getFirstPropertyValue("uid"));
  // An all-day master stores DATE-valued exceptions, derived from the wire
  // instant exactly the way its own DTSTART was - same discriminator, same
  // zone - so RECURRENCE-ID/EXDATE land on the dates the RRULE generates
  // and Thunderbird can bind them (RFC 5545 §3.8.4.4 matches by value
  // type).
  const allDay = readPathFrom(adNode, ["AllDayEvent"]) === "1";
  const { tzId } = resolveTimezone(adNode, defaultTimezone);

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
      addExdate(vevent, ridDate, allDay, tzId);
      continue;
    }

    const ridTime = allDay
      ? allDayIcalDate(ridDate, tzId)
      : jsDateToIcalUtcTime(ridDate);
    // Seeded from the master, never empty: §2.2.2.21 defines an absent
    // child element as "same as the top-level element", and this path
    // rebuilds every override from scratch on each re-delivery.
    const override = seedOverride({ master: vevent, existing: null, ridTime });
    vcal.addSubcomponent(override);
    if (masterUid) override.updatePropertyWithValue("uid", masterUid);
    const ridProp = new ICAL.Property("recurrence-id", override);
    ridProp.setValue(ridTime);
    override.addProperty(ridProp);
    await populateVeventFromAd({
      adNode: exc,
      vevent: override,
      asVersion,
      defaultTimezone,
      inheritedAllDay: allDay,
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
    // `instanceKey` keeps a DATE row date-shaped - the same encoding the
    // all-day master's own StartTime now uses in every version, so the
    // server lands master and exceptions on one occurrence grid.
    builder.atag("ExceptionStartTime", instanceKey(ex));
    builder.atag("Deleted", "1");
    builder.ctag();
  }
  for (const override of overrides) {
    const rid = override.getFirstPropertyValue("recurrence-id");
    builder.otag("Exception");
    builder.atag("ExceptionStartTime", instanceKey(rid));
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

/** Add an EXDATE, unless the same instant is already excluded. The guard is
 *  here rather than at the call sites because both of them can be reached
 *  for an occurrence we have already cancelled - the embedded `<Exceptions>`
 *  block on a re-delivered master, and a 16.1 per-instance `<Delete>` for an
 *  instance we EXDATE'd on an earlier sync. A duplicate is not inert: each
 *  copy becomes its own outbound `<Delete>` command. */
function addExdate(vevent, jsDate, allDay = false, tzId = null) {
  // `namesInstance` matches both stored forms, so a DATE row written now
  // and a DATE-TIME row from an older blob both count as "already
  // excluded" - each duplicate would otherwise become its own outbound
  // <Delete>.
  for (const existing of collectExdates(vevent)) {
    if (namesInstance(existing, jsDate, tzId)) return;
  }
  const prop = new ICAL.Property("exdate", vevent);
  prop.setValue(
    allDay ? allDayIcalDate(jsDate, tzId) : jsDateToIcalUtcTime(jsDate),
  );
  vevent.addProperty(prop);
}

function removeExdate(vevent, jsDate, tzId = null) {
  // Match-only, so no allDay parameter: `namesInstance` recognises both
  // stored forms on its own.
  for (const p of vevent.getAllProperties("exdate")) {
    const v = p.getFirstValue();
    if (v instanceof ICAL.Time && namesInstance(v, jsDate, tzId)) {
      vevent.removeProperty(p);
    }
  }
}

/**
 * Reduce a series to the smallest thing that identifies its exception set:
 * every cancelled instance, and every override keyed by RECURRENCE-ID with a
 * short digest of its content.
 *
 * This is what a changelog entry carries as the "before" side of a user
 * edit. Keeping the whole previous iCal would work too, but it puts several
 * kilobytes per pending item into the folder row for the sake of a
 * comparison that only ever asks *which* exceptions differ - and a bulk edit
 * would multiply that by the number of items touched. A digest answers the
 * same question at a fixed size.
 *
 * The digest deliberately covers the override's whole serialised form, so
 * any change to it registers; we never need to know *what* changed within an
 * override, only that it did, because an override is pushed whole.
 *
 * Returns null when the blob is unparseable or carries no recurrence - a
 * caller with no baseline falls back to sending the full set, which is what
 * happens today anyway.
 */
/* ── Invitations ───────────────────────────────────────────────────────
 *
 * A meeting somebody else organised cannot be changed through this
 * protocol. On 16.0/16.1 the client is forbidden from sending either half
 * of what such a change would have to say - [MS-ASCAL] 2.2.2.35 "the client
 * MUST NOT include the OrganizerEmail element", 2.2.2.5 the same for
 * AttendeeStatus - and the server substitutes the current user for both, so
 * an Add or a Change we sent would not read as "their meeting, changed" but
 * as "my meeting", re-inviting everyone on it. The one thing we may say
 * about such a meeting is the user's answer, and that goes as a
 * MeetingResponse.
 */

/** How the user's answer maps onto MeetingResponse's UserResponse. Anything
 *  absent from this table is not an answer: NEEDS-ACTION means the user has
 *  not replied, and there is no UserResponse value for "no reply". */
/** PARTSTAT → [MS-ASCAL] AttendeeStatus. NEEDS-ACTION is deliberately
 *  absent: 5 means "not responded", which is the server's to say, and
 *  writing it back would be us reporting an answer nobody gave. */
const PARTSTAT_TO_ATTENDEESTATUS = {
  TENTATIVE: "2",
  ACCEPTED: "3",
  DECLINED: "4",
};

const PARTSTAT_TO_USERRESPONSE = {
  ACCEPTED: 1,
  TENTATIVE: 2,
  DECLINED: 3,
};

/** Which attendee is us, on an item we may not have an address for.
 *
 *  `X-MOZ-INVITED-ATTENDEE` is Thunderbird's own answer, set when its iTIP
 *  processing matched an attendee to a local identity ("so we know who
 *  accepted the event") and deleted again before it sends anything, so no
 *  server can echo it back. It is the only answer available on an item
 *  Thunderbird filed from an emailed invitation, which carries none of our
 *  own properties at all. The account's address is the fallback for
 *  everything else. */
function selfAddress(vevent, userEmail) {
  const marked = stripMailto(
    stringOf(vevent.getFirstPropertyValue("x-moz-invited-attendee")),
  );
  return (marked || stringOf(userEmail)).toLowerCase();
}

/** The address of one ATTENDEE or ORGANIZER property, lower-cased and
 *  without its scheme, or "" when it carries none. */
function addressOfProperty(prop) {
  return stringOf(prop?.getFirstValue())
    .replace(/^mailto:/i, "")
    .trim()
    .toLowerCase();
}

/**
 * The fields whose change is worth telling the attendees about, as a plain
 * JSON-able bag, or null when the item carries no event at all.
 *
 * This is the whole definition of "worth announcing" and both sides of the
 * comparison use it: the hook records one of these as the item stood before
 * the user's edit, and the phase that sends the message builds another from
 * the item once the sync has settled. Equal means say nothing.
 *
 * Every value is normalised so that two spellings of one fact compare
 * equal, because anything that survives here becomes a message somebody
 * receives:
 *
 * - times as the UTC instant. The same moment written in two zones is not a
 *   reschedule, and a `TZID` respelled by a round trip must not mail
 *   anybody. All-day boundaries carry no zone at all and are compared as
 *   the calendar date, which is the only thing they mean.
 * - attendees as a sorted set of bare addresses. Thunderbird rewrites
 *   `PARTSTAT`, `RSVP` and `CN` on the property as replies arrive and as the
 *   dialog is used, and none of that is a change to who is invited.
 *
 * Deliberately absent: reminders, categories, colour, transparency, the
 * body. A user who moves an alarm has not changed the meeting.
 */
export function announceableOf(ical) {
  const vcal = parseVCalendar(ical);
  const master = vcal ? pickMasterVevent(vcal) : null;
  if (!master) return null;

  // The master alone. An override is one occurrence's business, and this
  // pass announces the series only - reading every component here would
  // turn any occurrence edit into a message to the whole series.
  //
  // Through `eventTimingFor`, never a raw DTEND read: an event can carry
  // DURATION instead, an all-day one can leave the end out entirely (§3.6.1
  // gives it a day), and Outlook writes an all-day DTEND equal to its
  // DTSTART meaning the same thing. All three are the same fact spelled
  // three ways, and that function is where the codec already knows it.
  const { dtstart, end, allDay } = eventTimingFor(master);
  const boundary = (v) => {
    const value = v?.getFirstValue ? v.getFirstValue() : v;
    if (!(value instanceof ICAL.Time)) return null;
    // An all-day boundary is a date, and `toJSDate` would give it a
    // midnight in whatever zone happens to be current - a day out for
    // anyone west of UTC, every time the item crossed the wire.
    if (value.isDate) return { date: value.toString() };
    const at = value.toJSDate?.();
    return at && !Number.isNaN(at.getTime()) ? { at: at.toISOString() } : null;
  };

  // Our own address is not somebody we invite, and Exchange returns the
  // organiser inside its own Attendees list - so the copy that comes back
  // can carry an ATTENDEE the saved one never had.
  const organizer = addressOfProperty(master.getFirstProperty("organizer"));
  const status = stringOf(master.getFirstPropertyValue("status"))
    .trim()
    .toUpperCase();

  return {
    start: boundary(dtstart),
    end: boundary(end),
    allDay,
    location: stringOf(master.getFirstPropertyValue("location")).trim(),
    summary: stringOf(master.getFirstPropertyValue("summary")).trim(),
    // The decode stamps CONFIRMED on everything the server calls a meeting,
    // while a locally-authored one carries no STATUS at all. Same fact.
    // Only TENTATIVE and CANCELLED say anything.
    status: status === "CONFIRMED" ? "" : status,
    // The schedule of a series is as announceable as its time. EXDATE is
    // deliberately absent: one occurrence removed is a deletion, and a
    // deletion is never announced.
    rrule: stringOf(master.getFirstPropertyValue("rrule")?.toString()) || null,
    attendees: [
      ...new Set(
        master
          .getAllProperties("attendee")
          .map(addressOfProperty)
          .filter((a) => a && a !== organizer),
      ),
    ].sort(),
  };
}

/** A stable string for one bag, whatever order its keys were built in and
 *  however deeply they nest. */
function canonicalBag(value) {
  if (Array.isArray(value)) return `[${value.map(canonicalBag).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value)
      .sort()
      .map((k) => `${JSON.stringify(k)}:${canonicalBag(value[k])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value ?? null);
}

/**
 * Do two bags describe the same meeting? This is the question that decides
 * whether anybody is mailed, so both ways of being wrong cost something
 * real: a false "different" mails every attendee about nothing, a false
 * "same" leaves them holding the old time.
 *
 * A bag whose **shape** differs is not comparable - it came from a build
 * that defined the fields differently. Answering "different" there would
 * mail every attendee of every pending meeting the first time somebody
 * updated the add-on, so an unreadable bag counts as no change. That is the
 * one place where the usual "when in doubt, send" bias is the wrong way
 * round: a missed notification is recoverable by saving the meeting again,
 * a mailshot is not.
 */
export function sameAnnounceable(a, b) {
  if (!a || !b) return false;
  const ka = Object.keys(a).sort().join();
  const kb = Object.keys(b).sort().join();
  if (ka !== kb) return true;
  return canonicalBag(a) === canonicalBag(b);
}

/** The attendees named in the earlier bag and gone from the later one.
 *  They are owed their own cancellation while the meeting goes ahead for
 *  everybody else, and their addresses exist nowhere but that earlier bag. */
export function droppedAttendees(from, now) {
  if (!from?.attendees) return [];
  const current = new Set(now?.attendees ?? []);
  return from.attendees.filter((a) => !current.has(a));
}

/** Is this a meeting somebody else organised?
 *
 *  Two signals, because neither sees every item.
 *
 *  `X-EAS-MEETINGSTATUS` is the server's own statement, recorded inbound on
 *  everything we pull: 0x1 says the item is a meeting rather than an
 *  appointment, and 0x2 - the R bit - says it came from another organizer.
 *  Authoritative, and blind to anything that has not been through that path.
 *
 *  An invitation Thunderbird filed from an emailed iTIP has not been: it is
 *  built from the message and carries no `X-EAS-*` at all. There the
 *  organizer is compared with `X-MOZ-INVITED-ATTENDEE`, which is the same
 *  test the platform itself uses to decide an item is an invitation
 *  (`calProviderBase.isInvitation`).
 *
 *  False whenever neither signal is conclusive: an item nobody has said
 *  anything about is ours. */
export function isReceivedMeeting(ical, userEmail = null) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return false;
  const master = pickMasterVevent(vcal);
  if (!master) return false;

  // When the server has spoken it is the answer, both ways. A clear R bit
  // is not silence: it says this meeting is the user's own, and it outranks
  // anything we could infer from addresses - which is what protects an
  // account whose organizer address is an alias of itself.
  const ms = parseInt(
    master.getFirstPropertyValue(X_EAS_MEETINGSTATUS.toLowerCase()),
    10,
  );
  if (Number.isFinite(ms)) return (ms & 0x1) === 0x1 && (ms & 0x2) === 0x2;

  const orgProp = master.getFirstProperty("organizer");
  if (!orgProp) return false; // names nobody, which is not the same as us
  const self = selfAddress(master, userEmail);
  if (!self) return false;
  return stripMailto(stringOf(orgProp.getFirstValue())).toLowerCase() !== self;
}

/** The self attendee of one component, by address. */
function selfAttendeeOf(comp, self) {
  return comp
    .getAllProperties("attendee")
    .find(
      (a) => stripMailto(stringOf(a.getFirstValue())).toLowerCase() === self,
    );
}

/** One component's answer as a MeetingResponse UserResponse, or null when
 *  it carries none. */
function userResponseOf(comp, self) {
  const partstat = selfAttendeeOf(comp, self)?.getParameter("partstat");
  return PARTSTAT_TO_USERRESPONSE[String(partstat).toUpperCase()] ?? null;
}

/** `MeetingResponse` names an occurrence by an `InstanceId` of exactly 24
 *  characters - `2026-09-08T09:00:00.000Z`. That is not the form the rest
 *  of the protocol uses: AirSyncBase's own `InstanceId`, and every other
 *  date on the wire, is the 16-character basic form. `instanceKey` speaks
 *  basic, so this widens it, and returns null for anything that is not the
 *  shape we expect rather than letting a malformed value reach the server -
 *  which answers `Status 2`, indistinguishable from a stale meeting. */
function extendedInstanceId(basic) {
  const m = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$/.exec(basic ?? "");
  if (!m) return null;
  const [, y, mo, d, h, mi, sec] = m;
  return `${y}-${mo}-${d}T${h}:${mi}:${sec}.000Z`;
}

/** Every answer this item carries: the series, and each occurrence the user
 *  answered on its own.
 *
 *  Answering one occurrence of a recurring invitation is what Thunderbird
 *  does whenever the answer is given from the calendar rather than from the
 *  message - it writes an override and leaves the master alone - so a
 *  master-only reading of the item misses the ordinary case entirely.
 *
 *  Each answer carries the `instanceId` naming its occurrence, or null for
 *  the series. An occurrence whose `RECURRENCE-ID` cannot be expressed as
 *  an `InstanceId` is skipped rather than answered against the series,
 *  which would apply the answer to every occurrence at once.
 *
 *  @returns {Array<{instanceId: string|null, rid: string|null,
 *                   userResponse: number, responseType: string|null}>}
 */
export function selfUserResponses(ical, userEmail = null) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return [];
  const master = pickMasterVevent(vcal);
  if (!master) return [];
  // Addresses live on the master: an override need not repeat the
  // X-MOZ-INVITED-ATTENDEE marker Thunderbird put there.
  const self = selfAddress(master, userEmail);
  if (!self) return [];

  const answers = [];
  const masterResponse = userResponseOf(master, self);
  if (masterResponse) {
    answers.push({
      instanceId: null,
      rid: null,
      userResponse: masterResponse,
      responseType: responseTypeOf(master),
    });
  }

  for (const sub of vcal.getAllSubcomponents("vevent")) {
    const ridProp = sub.getFirstProperty("recurrence-id");
    if (!ridProp) continue;
    const userResponse = userResponseOf(sub, self);
    if (!userResponse) continue;
    const rid = instanceKey(sub.getFirstPropertyValue("recurrence-id"));
    const instanceId = extendedInstanceId(rid);
    if (!instanceId) continue;
    answers.push({
      instanceId,
      rid,
      userResponse,
      // An override of the server's own making carries no ResponseType:
      // measured on a 14.1 mailbox, answering one occurrence sets the
      // *master* to 3 and leaves the exception unstamped. Falling back to
      // the master is therefore what "has the server heard this answer?"
      // means for an occurrence. It still says send whenever the two
      // disagree, which is what a decline on one occurrence of an accepted
      // series looks like.
      responseType: responseTypeOf(sub) ?? responseTypeOf(master),
    });
  }
  return answers;
}

/** What the server last told us about this component's response, as the
 *  PARTSTAT it maps to, or null when the server has never said. */
export function serverKnownPartstat(responseType) {
  if (responseType == null || responseType === "") return null;
  return RESPONSETYPE_TO_PARTSTAT[parseInt(responseType, 10)] ?? null;
}

function responseTypeOf(comp) {
  const v = comp.getFirstPropertyValue(X_EAS_RESPONSETYPE.toLowerCase());
  return v == null || v === "" ? null : String(v);
}

/** The PARTSTAT the user's answer implies, for comparing against what the
 *  server already knows. Inverse of `PARTSTAT_TO_USERRESPONSE`. */
export const USERRESPONSE_TO_PARTSTAT = Object.freeze({
  1: "ACCEPTED",
  2: "TENTATIVE",
  3: "DECLINED",
});

/** Carry the user's own answer across an adoption of the server's copy.
 *
 *  The response phase runs after the pull, so without this the pull would
 *  overwrite the answer with the server's - which does not know about it
 *  yet - and we would then read that back and send it. The user accepts,
 *  the calendar looks right, and the organizer never hears. It fails
 *  silently, which is why it is a rule and not a nicety.
 *
 *  Only a real answer wins: NEEDS-ACTION is the absence of one, and letting
 *  it win would erase an answer the server already knows about.
 *
 *  Contributed by Tomas Kovacik <kovacik@dgtfactory.com> in PR #339; the
 *  self-attendee lookup is widened here to Thunderbird's marker, so it also
 *  works on an item we hold no address for. */
/** Every `X-EAS-*` property on a component, whatever it is called. Matched
 *  by prefix rather than by a list, so a stamp added later is covered
 *  without anyone remembering to come back here. */
function easStampsOf(comp) {
  return comp
    .getAllProperties()
    .filter((p) => p.name.toLowerCase().startsWith("x-eas-"));
}

/** Hold our own stamps to what they were, on an item somebody else wrote.
 *
 *  `X-EAS-SERVERID` and its siblings are the item's identity on the server
 *  and our record of what the server said about it. Nothing outside this
 *  add-on has any business writing them - and nothing outside it ever
 *  legitimately does, because our own sync writes go to `<id>#cache`, which
 *  fires no item hooks at all. So every write that reaches a hook is
 *  somebody else's, and any difference in these properties is damage:
 *  Thunderbird rebuilding an item from an emailed invitation, a copied
 *  event carrying the identity of the one it was copied from, an import
 *  overwriting an item that is already synced.
 *
 *  One rule covers all of it. A local item never invents a server identity,
 *  and never loses one it already had:
 *
 *    - with a previous version, its stamps are restored exactly - one that
 *      was removed comes back, one that was altered reverts, one that was
 *      introduced is dropped;
 *    - with none, every stamp is stripped, because a genuinely new item has
 *      no identity to carry.
 *
 *  `priorIcal` is that previous version: the platform hands it to an update
 *  hook, and a create hook finds it by looking the incoming UID up - an
 *  import of an already-synced item arrives as a create, and stripping it
 *  would detach the item from the server rather than update it.
 *
 *  Stamps belong on the master; `stampEasServerId` puts them there and the
 *  reader skips overrides. So an override never keeps one, whoever wrote it.
 *
 *  Returns the input untouched when it cannot be parsed: a save must never
 *  fail over this. */
export function pinEasStamps({ builtIcal, priorIcal = null }) {
  const vcal = parseVCalendar(builtIcal);
  if (!vcal) return builtIcal;
  const comps = vcal
    .getAllSubcomponents()
    .filter((c) => c.name === "vevent" || c.name === "vtodo");
  if (!comps.length) return builtIcal;

  for (const comp of comps) {
    for (const prop of easStampsOf(comp)) comp.removeProperty(prop);
  }

  // Restored per component, not just onto the master. An occurrence of a
  // recurring meeting carries its own stamps - a 16.1 mailbox records the
  // answer to one occurrence as a ResponseType on that exception - and
  // stripping every component while restoring one loses them on every
  // local edit. The series keys as null so it matches its own prior self
  // rather than whichever occurrence happens to come first.
  const prior = priorIcal ? parseVCalendar(priorIcal) : null;
  const priorStamps = new Map();
  if (prior) {
    for (const comp of prior.getAllSubcomponents()) {
      if (comp.name !== "vevent" && comp.name !== "vtodo") continue;
      const stamps = easStampsOf(comp);
      if (!stamps.length) continue;
      const ridValue = comp.getFirstPropertyValue("recurrence-id");
      priorStamps.set(ridValue ? instanceKey(ridValue) : null, stamps);
    }
  }
  for (const comp of comps) {
    const ridValue = comp.getFirstPropertyValue("recurrence-id");
    const stamps = priorStamps.get(ridValue ? instanceKey(ridValue) : null);
    if (!stamps) continue;
    for (const prop of stamps) {
      comp.updatePropertyWithValue(prop.name, prop.getFirstValue());
    }
  }
  return vcal.toString();
}

export function preserveSelfPartstat({ builtIcal, priorIcal, userEmail }) {
  const priorCal = parseVCalendar(priorIcal);
  const priorMaster = priorCal ? pickMasterVevent(priorCal) : null;
  if (!priorMaster) return builtIcal;
  const self = selfAddress(priorMaster, userEmail);
  if (!self) return builtIcal;

  // Keyed by occurrence, with the series under null. An answer given on one
  // occurrence lives on its override and nowhere else, so carrying only the
  // master would drop it here - and the answer is read back out of this
  // item after the pull, so dropping it loses it for good.
  const keep = new Map();
  for (const comp of priorCal.getAllSubcomponents("vevent")) {
    const partstat = selfAttendeeOf(comp, self)?.getParameter("partstat");
    if (!partstat || String(partstat).toUpperCase() === "NEEDS-ACTION") {
      continue;
    }
    const ridValue = comp.getFirstPropertyValue("recurrence-id");
    keep.set(ridValue ? instanceKey(ridValue) : null, partstat);
  }
  if (!keep.size) return builtIcal;

  const vcal = parseVCalendar(builtIcal);
  if (!vcal) return builtIcal;
  let touched = false;
  for (const comp of vcal.getAllSubcomponents("vevent")) {
    const ridValue = comp.getFirstPropertyValue("recurrence-id");
    const partstat = keep.get(ridValue ? instanceKey(ridValue) : null);
    if (!partstat) continue;
    const builtSelf = selfAttendeeOf(comp, self);
    if (!builtSelf) continue;
    builtSelf.setParameter("partstat", partstat);
    touched = true;
  }
  return touched ? vcal.toString() : builtIcal;
}

export function exceptionFingerprint(ical) {
  const vcal = parseVCalendar(ical);
  if (!vcal) return null;
  const master = pickMasterVevent(vcal);
  if (!master) return null;

  const exdates = collectExdates(master)
    .map((t) => instanceKey(t))
    .sort();

  const overrides = [];
  for (const sub of vcal.getAllSubcomponents("vevent")) {
    const rid = sub.getFirstProperty("recurrence-id");
    if (!rid) continue;
    overrides.push({
      rid: instanceKey(sub.getFirstPropertyValue("recurrence-id")),
      digest: digestOf(sub.toString()),
    });
  }
  overrides.sort((a, b) => (a.rid < b.rid ? -1 : a.rid > b.rid ? 1 : 0));

  if (!exdates.length && !overrides.length && !master.getFirstProperty("rrule"))
    return null;
  return { exdates, overrides };
}

/** FNV-1a, as an unsigned 32-bit hex string. Not a security hash - it only
 *  has to make "this override changed" cheap and stable across restarts,
 *  which rules out anything seeded or address-derived. */
function digestOf(text) {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
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

/** Whether a stored RECURRENCE-ID / EXDATE names the occurrence EAS
 *  identified by the UTC instant `jsDate`. Two stored forms exist and
 *  both must keep matching: a DATE row (all-day, the form the writers
 *  emit) names the calendar date the instant derives to, and a DATE-TIME
 *  row (timed events, and every all-day blob written before the DATE
 *  form) names the instant itself. */
function namesInstance(t, jsDate, tzId) {
  if (!(t instanceof ICAL.Time)) return false;
  if (!t.isDate) return icalTimeToBasicUtc(t) === formatBasicUtc(jsDate);
  const d = allDayDateFromUtcInstant(jsDate, tzId);
  return t.year === d.year && t.month === d.month && t.day === d.day;
}

/** The instance identity a 16.1 exchange runs on - the wire InstanceId,
 *  and the key `exceptionFingerprint` and `listInstanceCommands` agree
 *  through. A DATE row keys as the fake-local form the protocol itself
 *  uses for all-day instances (`YYYYMMDDT000000Z`); a DATE-TIME row keys
 *  as its instant, which keeps old-form blobs comparable to fingerprints
 *  taken before the DATE form existed. */
function instanceKey(t) {
  return t instanceof ICAL.Time && t.isDate
    ? fakeLocalAsUtcFromValue(t)
    : icalTimeToBasicUtc(t);
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
    const isSelf = userEmailLower && email.toLowerCase() === userEmailLower;
    if (status) {
      item.partstat = ATTENDEESTATUS_TO_PARTSTAT[status] ?? "NEEDS-ACTION";
    } else if (isSelf && fallbackResponseType) {
      // Legacy calendarsync.js:203-204: when AttendeeStatus is missing
      // for the self-attendee, fall back to the event-level ResponseType -
      // through that element's own table, which is not the one above.
      item.partstat =
        RESPONSETYPE_TO_PARTSTAT[fallbackResponseType] ?? "NEEDS-ACTION";
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

/** Element order is `[MS-ASCAL]`'s and is load-bearing - the server validates
 *  against a sequence - so it is kept exactly as it was when this derivation
 *  moved into `recurrence.mjs`. */
function appendRecurrence(
  builder,
  rruleProp,
  dtstartProp,
  unmapped = [],
  firstDayOfWeek = null,
) {
  const rec = rruleToEas(rruleProp, dtstartProp);
  if (!rec) return;

  builder.otag("Recurrence");
  // Ahead of <Type>, where a server-authored block puts them - see
  // `UNMAPPED_RECURRENCE`. Empty unless the server stated something we do
  // not model, which for an event means the non-Gregorian pair.
  for (const [tag, value] of unmapped) builder.atag(tag, value);
  builder.atag("Type", String(rec.type));
  if (rec.dayOfMonth !== null) {
    builder.atag("DayOfMonth", String(rec.dayOfMonth));
  }
  if (rec.dayOfWeek !== null) builder.atag("DayOfWeek", String(rec.dayOfWeek));
  builder.atag("Interval", String(rec.interval));
  if (rec.monthOfYear !== null) {
    builder.atag("MonthOfYear", String(rec.monthOfYear));
  }
  if (rec.count) builder.atag("Occurrences", String(rec.count));
  else if (rec.until) builder.atag("Until", untilFor(rec.until));
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

/* ── Helpers: body codepage ───────────────────────────────────────── */

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

/** What we last told the organiser, or null. */
export function repliedPartstatOf(ical) {
  const vcal = parseVCalendar(ical);
  const master = vcal ? pickMasterVevent(vcal) : null;
  const v = master?.getFirstPropertyValue(X_EAS_REPLIED.toLowerCase());
  return v == null || v === "" ? null : String(v).toUpperCase();
}

/** Record what we just told them. Returns the blob to store. */
export function stampRepliedPartstat(ical, partstat) {
  const vcal = parseVCalendar(ical);
  const master = vcal ? pickMasterVevent(vcal) : null;
  if (!master) return ical;
  master.updatePropertyWithValue(
    X_EAS_REPLIED.toLowerCase(),
    String(partstat).toUpperCase(),
  );
  return vcal.toString();
}
