/**
 * Item 34a — the account asks the server which mailbox it is, and the
 * answer decides two things: who a calendar tells Thunderbird it belongs
 * to, and whether an event goes out as one the user organises
 * (MeetingStatus 1) or merely attends (3).
 *
 * The EAS login cannot answer either question - it may be `DOMAIN\user`
 * or any other non-address - so `UserInformation/Get` is asked instead.
 * Two reply shapes exist in the wild and both are parsed here; the live
 * suite proves the request itself reaches a real server.
 */

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import {
  accountUserAddress,
  readUserInformation,
} from "../../src/modules/eas/settings.mjs";
import {
  appendApplicationDataFromIcal,
  applicationDataToIcal,
} from "../../src/modules/eas/calendar-codec.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

before(() => ensureLoaded());

/** The readers take a decoded document, so wrap a captured body the way
 *  the WBXML decoder hands it over: a root element with children.
 *
 *  With one correction the fixture parser cannot make on its own. A real
 *  document's `children` is a live DOM collection - iterable, indexable,
 *  and carrying NONE of Array's methods - while `parseAdNode` builds
 *  plain arrays. Code that reached for `.filter` therefore passed here
 *  and threw against a real server, so every collection is handed over
 *  array-like, exactly as the DOM would. */
const domLike = (node) => {
  if (!node?.children) return node;
  for (const child of node.children) domLike(child);
  const kids = node.children;
  node.children = {
    length: kids.length,
    [Symbol.iterator]: () => kids[Symbol.iterator](),
  };
  for (let i = 0; i < kids.length; i++) node.children[i] = kids[i];
  return node;
};

const doc = (xml) => ({ documentElement: domLike(parseAdNode(xml)) });

test("14.1+ reply: the primary address wins over the other SMTP entries", () => {
  const info = readUserInformation(
    doc(`<Settings>
      <Status>1</Status>
      <UserInformation>
        <Status>1</Status>
        <Get>
          <Accounts><Account>
            <UserDisplayName>John Bieling</UserDisplayName>
            <EmailAddresses>
              <SMTPAddress>alias@example.org</SMTPAddress>
              <PrimarySmtpAddress>john@example.org</PrimarySmtpAddress>
            </EmailAddresses>
          </Account></Accounts>
        </Get>
      </UserInformation>
    </Settings>`),
  );
  assert.equal(info.address, "john@example.org");
  assert.equal(info.displayName, "John Bieling");
});

test("12.1/14.0 reply: bare SMTPAddress entries, first one taken", () => {
  // No Accounts wrapper and no primary marker on these versions, so the
  // first address is the best answer available.
  const info = readUserInformation(
    doc(`<Settings>
      <Status>1</Status>
      <UserInformation>
        <Status>1</Status>
        <Get>
          <EmailAddresses>
            <SMTPAddress>john@example.org</SMTPAddress>
            <SMTPAddress>alias@example.org</SMTPAddress>
          </EmailAddresses>
        </Get>
      </UserInformation>
    </Settings>`),
  );
  assert.equal(info.address, "john@example.org");
  assert.equal(info.displayName, null);
});

test("a reply with no address yields null rather than throwing", () => {
  const empty = `<Settings><Status>1</Status></Settings>`;
  assert.equal(readUserInformation(doc(empty)), null, "no UserInformation");
  assert.equal(
    readUserInformation(
      doc(
        `<Settings><UserInformation><Status>2</Status></UserInformation></Settings>`,
      ),
    ),
    null,
    "a UserInformation without Get",
  );
  assert.equal(
    readUserInformation(
      doc(`<Settings><UserInformation><Get><Accounts><Account>
             <EmailAddresses/>
           </Account></Accounts></Get></UserInformation></Settings>`),
    ),
    null,
    "an empty address list",
  );
  assert.equal(readUserInformation(null), null, "no document at all");
});

test("accountUserAddress prefers the learned address and falls back to the login", () => {
  assert.equal(
    accountUserAddress({
      custom: { userSmtpAddress: "john@example.org", user: "DOMAIN\\jbieling" },
    }),
    "john@example.org",
  );
  assert.equal(
    accountUserAddress({ custom: { user: "DOMAIN\\jbieling" } }),
    "DOMAIN\\jbieling",
    "unlearned accounts keep the old behaviour",
  );
  assert.equal(accountUserAddress({}), undefined);
});

/* ── what the address is for: MeetingStatus ─────────────────────────── */

const withAttendees = (organizer) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    "UID:meeting-uid",
    "DTSTAMP:20260801T090000Z",
    "DTSTART:20260810T140000Z",
    "DTEND:20260810T150000Z",
    "SUMMARY:staff meeting",
    `ORGANIZER;CN=Someone:mailto:${organizer}`,
    "ATTENDEE;CN=Other:mailto:other@example.org",
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n");

function meetingStatusFor(ics, userEmail) {
  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendApplicationDataFromIcal({
    builder: w,
    ical: ics,
    asVersion: "14.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    userEmail,
  });
  w.switchpage("AirSync");
  w.ctag();
  return readPathFrom(parseAdNode(decodeWBXML(w.getBytes())), [
    "MeetingStatus",
  ]);
}

test("MeetingStatus: the mailbox owner organises (1), anyone else attends (3)", () => {
  assert.equal(
    meetingStatusFor(withAttendees("john@example.org"), "john@example.org"),
    "1",
    "the user's own meeting must not go out as one they received",
  );
  assert.equal(
    meetingStatusFor(withAttendees("boss@example.org"), "john@example.org"),
    "3",
  );
});

test("MeetingStatus: a login that is not an address mislabels - the bug this fixes", () => {
  // With only `DOMAIN\user` to compare against, the user's own meeting
  // reads as received. Pinned so the fallback's cost stays visible.
  assert.equal(
    meetingStatusFor(withAttendees("john@example.org"), "DOMAIN\\jbieling"),
    "3",
  );
  assert.equal(
    meetingStatusFor(withAttendees("john@example.org"), undefined),
    "3",
    "and an unknown address does the same",
  );
});

test("the self-attendee's PARTSTAT falls back to ResponseType by address", async () => {
  // The second consumer of the same value: with no AttendeeStatus, the
  // event-level ResponseType applies to the attendee that is us.
  const ical = await applicationDataToIcal({
    adNode: parseAdNode(`<ApplicationData>
      <StartTime xmlns='Calendar'>20260810T140000Z</StartTime>
      <EndTime xmlns='Calendar'>20260810T150000Z</EndTime>
      <Subject xmlns='Calendar'>staff meeting</Subject>
      <UID xmlns='Calendar'>meeting-uid</UID>
      <ResponseType xmlns='Calendar'>3</ResponseType>
      <Attendees xmlns='Calendar'>
        <Attendee><Email>john@example.org</Email><Name>John</Name></Attendee>
        <Attendee><Email>other@example.org</Email><Name>Other</Name></Attendee>
      </Attendees>
    </ApplicationData>`),
    existingIcal: null,
    serverID: "srv-meeting",
    asVersion: "14.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: null,
    userEmail: "john@example.org",
  });
  const self = ical
    .split(/\r?\n/)
    .find((l) => l.startsWith("ATTENDEE") && l.includes("john@example.org"));
  assert.ok(self, "expected the self attendee");
  assert.match(self, /PARTSTAT=ACCEPTED/, "ResponseType 3 is accepted");
  const other = ical
    .split(/\r?\n/)
    .find((l) => l.startsWith("ATTENDEE") && l.includes("other@example.org"));
  assert.match(other, /PARTSTAT=NEEDS-ACTION/, "only the self row inherits it");
});
