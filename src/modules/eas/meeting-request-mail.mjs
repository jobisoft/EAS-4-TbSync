/**
 * The invitation, the update and the cancellation - the messages an
 * organiser owes the attendees, and the mirror of `meeting-response-mail`.
 *
 * [MS-ASCMD] *Working with meeting requests*: to create a meeting a client
 * must both add the event to the organiser's calendar **and** "send an email
 * to prospective attendees". Two operations, not one - the server does not
 * derive the invitation from the pushed attendee list, which is exactly what
 * a 14.1 server was measured doing: nothing. From 16.0 it sends these
 * itself and we must stay silent.
 *
 * The `UID` in the iCalendar **must** equal the one on the calendar item.
 * That is what lets the organiser's mailbox reconcile the replies that come
 * back against the meeting they belong to.
 *
 * Updates and cancellations are not covered by any Microsoft document - the
 * blog promised a follow-up that was never written - so they follow RFC 5546
 * (iTIP): an update is a fresh `METHOD:REQUEST`, a cancellation is
 * `METHOD:CANCEL`.
 *
 * ## `SEQUENCE:0`
 *
 * iTIP wants `SEQUENCE` raised on every update, and a receiver is entitled
 * to ignore a `REQUEST` that does not raise it. EAS carries no `SEQUENCE` on
 * the wire, so keeping a real one would mean a new property on every item
 * and a new entry in the codec's stamp guard - a wire-format change to serve
 * a spec nicety. In practice receivers apply the newest `REQUEST` for a
 * `UID`. If updates are ever seen being ignored in the field, this is the
 * reason and the fix; it is not paid for up front.
 */

import ICAL from "../../vendor/ical.min.js";

const CRLF = "\r\n";

/** What each method calls itself in the subject. */
const SUBJECT_PREFIX = Object.freeze({
  REQUEST: "",
  CANCEL: "Canceled: ",
});

/** Properties that must not leave this profile.
 *
 *  `X-EAS-*` are our own bookkeeping - the item's identity on the server and
 *  what the server last said about it - and mean nothing to a recipient.
 *  `X-MOZ-*` are Thunderbird's, including which alarms this user has already
 *  dismissed. Everything else on the event is the meeting itself and
 *  travels. */
const PRIVATE_PREFIX = /^x-(eas|moz)-/i;

function headerSafe(text) {
  return String(text ?? "")
    .replace(/[\r\n]+/g, " ")
    .trim();
}

function addressOf(name, email) {
  const clean = headerSafe(name);
  return clean ? `${clean} <${email}>` : email;
}

/** The address one ATTENDEE or ORGANIZER carries, bare and lower-cased. */
function addressValue(prop) {
  return String(prop?.getFirstValue() ?? "")
    .replace(/^mailto:/i, "")
    .trim()
    .toLowerCase();
}

/**
 * One message for one meeting, as raw MIME, or null when there is nothing
 * to send.
 *
 * `recipients` are the addresses this message goes to, which is not always
 * the event's attendee list: somebody dropped from a meeting that is going
 * ahead gets their own `CANCEL` while everybody else gets the updated
 * `REQUEST`. For a `CANCEL` the iCalendar's attendee list is narrowed to the
 * recipients, so the message says who it is about.
 *
 * `now` is passed in rather than read, so `DTSTAMP` is testable.
 */
export function buildMeetingRequestMime({
  blob,
  method,
  recipients,
  userEmail,
  userName = "",
  now,
}) {
  if (!blob || !userEmail) return null;
  if (method !== "REQUEST" && method !== "CANCEL") return null;
  const to = [...new Set((recipients ?? []).map((r) => String(r).trim()))].filter(
    Boolean,
  );
  if (!to.length) return null;

  let vcal;
  let master;
  try {
    vcal = new ICAL.Component(ICAL.parse(blob));
    master =
      vcal
        .getAllSubcomponents("vevent")
        .find((v) => !v.getFirstProperty("recurrence-id")) ??
      vcal.getFirstSubcomponent("vevent");
  } catch {
    return null;
  }
  if (!master) return null;

  const uid = String(master.getFirstPropertyValue("uid") ?? "");
  if (!uid) return null;

  // The master only. Overrides are one occurrence's business and this pass
  // announces the series; sending them would tell attendees about changes
  // they were never given the context for.
  const event = new ICAL.Component(
    JSON.parse(JSON.stringify(master.toJSON())),
  );
  // Snapshotted before removing: ical.js hands back its live array, so
  // deleting while iterating steps over the element after each one taken -
  // which silently left every second private property in the message.
  for (const alarm of [...event.getAllSubcomponents("valarm")]) {
    event.removeSubcomponent(alarm);
  }
  for (const prop of [...event.getAllProperties()]) {
    if (PRIVATE_PREFIX.test(prop.name)) event.removeProperty(prop);
  }
  for (const name of ["dtstamp", "sequence", "status", "recurrence-id"]) {
    event.removeAllProperties(name);
  }
  event.addPropertyWithValue("dtstamp", ICAL.Time.fromJSDate(now, true));
  event.addPropertyWithValue("sequence", 0);

  if (method === "CANCEL") {
    // A cancelled meeting, addressed to the people it is being cancelled
    // for - which on a dropped attendee is not everybody on the event.
    event.addPropertyWithValue("status", "CANCELLED");
    // The prefix goes on the SUMMARY as well as the subject, which is what
    // Outlook does and what a recipient actually reads: a client renders an
    // iTIP message from the calendar item, not from the MIME headers, so a
    // cancellation whose prefix lived only in the subject arrived looking
    // exactly like the invitation it was cancelling. Measured on Microsoft.
    const titled = headerSafe(event.getFirstPropertyValue("summary") ?? "");
    event.removeAllProperties("summary");
    event.addPropertyWithValue("summary", `${SUBJECT_PREFIX.CANCEL}${titled}`);
    const keep = new Set(to.map((a) => a.toLowerCase()));
    for (const prop of [...event.getAllProperties("attendee")]) {
      if (!keep.has(addressValue(prop))) event.removeProperty(prop);
    }
  }

  const out = new ICAL.Component("vcalendar");
  out.addPropertyWithValue("version", "2.0");
  out.addPropertyWithValue("prodid", "-//TbSync//EAS-4-TbSync//EN");
  out.addPropertyWithValue("method", method);
  // Copied, never generated: Thunderbird keeps the definition in the blob
  // for a TZID-anchored DTSTART, and a receiver cannot resolve a zone it
  // was not given.
  for (const tz of vcal.getAllSubcomponents("vtimezone")) {
    out.addSubcomponent(tz);
  }
  out.addSubcomponent(event);

  const summary = headerSafe(master.getFirstPropertyValue("summary") ?? "");
  const boundary = `eas-request-${uid}`.replace(/[^A-Za-z0-9-]/g, "").slice(0, 60);

  return [
    `From: ${addressOf(userName, userEmail)}`,
    `To: ${to.join(", ")}`,
    `Subject: ${SUBJECT_PREFIX[method]}${summary || "(no subject)"}`,
    "MIME-Version: 1.0",
    `Content-Type: multipart/alternative; boundary="${boundary}"`,
    "",
    `--${boundary}`,
    "Content-Type: text/plain; charset=utf-8",
    "",
    // Deliberately bare, as the reply is: the calendar part carries the
    // meeting and anything written here is text the user did not write
    // appearing over their name. Giving the body a readable summary of
    // when and where is a follow-up, decided separately.
    `${SUBJECT_PREFIX[method]}${summary}`,
    "",
    `--${boundary}`,
    `Content-Type: text/calendar; charset=utf-8; method=${method}; name="meeting.ics"`,
    "",
    out.toString(),
    "",
    `--${boundary}--`,
    "",
  ].join(CRLF);
}
