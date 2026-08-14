/**
 * The MeetingResponse request, on the wire.
 *
 * The subject is who tells the organiser. [MS-ASCMD] splits that by
 * protocol version and the split is total: up to 14.1 the client sends its
 * own mail afterwards, and from 16.0 the server does it - "If the
 * SendResponse element is not present, no email will be sent."
 *
 * Getting it wrong is silent on both sides. The server answers Status 1
 * either way, the user's own calendar is correct either way, and inside a
 * single tenant the organiser's tracking updates through the shared store
 * even with no message sent - so the omission only shows when the organiser
 * is on another system, which is where it was found (Rutger, 5.2.3).
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { buildMeetingResponseBody } from "../../src/modules/eas/meeting-response.mjs";
import { decodeWBXML } from "../../src/modules/wbxml.mjs";

const REQUEST = {
  collectionId: "7",
  serverID: "req-1",
  userResponse: 1,
};

const xml = (over = {}) =>
  decodeWBXML(buildMeetingResponseBody({ ...REQUEST, ...over }));

test("16.x asks the server to mail the organiser", () => {
  for (const v of ["16.0", "16.1"]) {
    assert.match(xml({ asVersion: v }), /<SendResponse\s*\/>/, `on ${v}`);
  }
});

test("below 16.0 it is left out, because the element does not exist there", () => {
  // Not a policy choice: SendResponse is listed for 16.0 and 16.1 only, and
  // on these versions the client sends the mail itself instead. Emitting an
  // unknown token risks the server rejecting the whole response, which
  // would cost the answer as well as the notification.
  for (const v of ["2.5", "12.0", "12.1", "14.0", "14.1"]) {
    assert.doesNotMatch(xml({ asVersion: v }), /SendResponse/, `on ${v}`);
  }
});

test("a version we cannot read does not get it either", () => {
  for (const v of [null, undefined, "", "auto", "garbage"]) {
    assert.doesNotMatch(xml({ asVersion: v }), /SendResponse/, `for ${v}`);
  }
});

test("the rest of the request is unchanged by it", () => {
  const text = xml({ asVersion: "16.1" });
  assert.match(text, /<UserResponse>1<\/UserResponse>/);
  assert.match(text, /<CollectionId>7<\/CollectionId>/);
  assert.match(text, /<RequestId>req-1<\/RequestId>/);
});

test("an occurrence is answered with both, and the server mails for it", () => {
  // The client-side path skips a single occurrence - its reply carries no
  // RECURRENCE-ID, so it could only describe the series. The server has the
  // InstanceId in this very request and has no such problem, so 16.x is not
  // held to that limit.
  const text = xml({ asVersion: "16.1", instanceId: "2026-09-08T09:00:00.000Z" });
  // The decoder percent-escapes the colons on its way back to text; what
  // went onto the wire is the plain value.
  assert.match(text, /<InstanceId>2026-09-08T09%3A00%3A00\.000Z<\/InstanceId>/);
  assert.match(text, /<SendResponse\s*\/>/);
});

test("declining is announced too", () => {
  // A decline the organiser never hears about is the same defect as an
  // accept they never hear about, and the server deletes the local item -
  // so nothing is left locally to notice it by.
  assert.match(xml({ asVersion: "16.1", userResponse: 3 }), /<SendResponse\s*\/>/);
});
