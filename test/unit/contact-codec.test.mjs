/**
 * Ported from PR #345 (tomaskovacik) to the node:test layer; fixtures
 * kept verbatim (several are live-server captures), expectations
 * re-verified against current master.
 */

// contact-codec.mjs: EAS Contacts <-> vCard 4.0. A representative
// slice, not exhaustive - the module maps a large number of individual
// fields (prefixed phones, four Custom slots, half a dozen pass-through
// properties); this covers name/email/phone/organization plus the
// serverID/merge-aware/outbound mechanics shared by all of them, which
// is where a regression is most likely to actually bite. Extending

import { test } from "node:test";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import ICAL from "../../src/vendor/ical.min.js";
import assert from "node:assert/strict";
import {
  applicationDataToVCard,
  appendApplicationDataFromVCard,
  readEasServerIdFromVCard,
  stampEasServerId,
} from "../../src/modules/eas/contact-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

const ADD_CONTACT = `<ApplicationData>
  <FirstName xmlns='Contacts'>Sample</FirstName>
  <LastName xmlns='Contacts'>User</LastName>
  <Email1Address xmlns='Contacts'>Sample%20User%20%3Cuser%40example.invalid%3E</Email1Address>
  <BusinessPhoneNumber xmlns='Contacts'>%2B49301234567</BusinessPhoneNumber>
  <CompanyName xmlns='Contacts'>Example%20GmbH</CompanyName>
  <JobTitle xmlns='Contacts'>President%20of%20space</JobTitle>
</ApplicationData>`;

test("applicationDataToVCard: maps Name/Email/Phone/Organization/JobTitle and stamps the ServerId", async () => {
  const vcard = await applicationDataToVCard({
    adNode: parseAdNode(ADD_CONTACT),
    existingVcard: null,
    serverID: "server-id-contact-1",
    asVersion: "16.1",
    separator: ",",
    uid: null,
  });

  const comp = new ICAL.Component(ICAL.parse(vcard));
  assert.deepEqual(comp.getFirstPropertyValue("n"), [
    "User",
    "Sample",
    "",
    "",
    "",
  ]);
  assert.equal(comp.getFirstPropertyValue("email"), "user@example.invalid");
  const tel = comp.getFirstProperty("tel");
  assert.equal(tel.getFirstValue(), "+49301234567");
  assert.equal(tel.getParameter("type"), "work");
  assert.deepEqual(comp.getFirstPropertyValue("org"), "Example GmbH");
  assert.equal(comp.getFirstPropertyValue("title"), "President of space");
  assert.equal(readEasServerIdFromVCard(vcard), "server-id-contact-1");
});

test("applicationDataToVCard: Name is merge-aware - a delta with no name tags leaves the existing N property untouched", async () => {
  const afterAdd = await applicationDataToVCard({
    adNode: parseAdNode(ADD_CONTACT),
    existingVcard: null,
    serverID: "server-id-contact-1",
    asVersion: "16.1",
    separator: ",",
    uid: null,
  });

  const afterPhoneOnlyChange = await applicationDataToVCard({
    adNode: parseAdNode(
      `<ApplicationData><MobilePhoneNumber xmlns='Contacts'>+49301234568</MobilePhoneNumber></ApplicationData>`,
    ),
    existingVcard: afterAdd,
    serverID: "server-id-contact-1",
    asVersion: "16.1",
    separator: ",",
    uid: null,
  });

  const comp = new ICAL.Component(ICAL.parse(afterPhoneOnlyChange));
  assert.deepEqual(comp.getFirstPropertyValue("n"), [
    "User",
    "Sample",
    "",
    "",
    "",
  ]);
  // Phones ARE merge-aware too, but at the whole-set level (any phone
  // tag present reasserts the complete set) - so BusinessPhoneNumber
  // from the original Add is gone now, replaced by just the Mobile one
  // this delta carries. That's the documented behavior, not a bug.
  const tels = comp.getAllProperties("tel");
  assert.equal(tels.length, 1);
  assert.equal(tels[0].getFirstValue(), "+49301234568");
});

test("stampEasServerId / readEasServerIdFromVCard round-trip", async () => {
  const vcard = await applicationDataToVCard({
    adNode: parseAdNode(ADD_CONTACT),
    existingVcard: null,
    serverID: "original-id",
    asVersion: "16.1",
    separator: ",",
    uid: null,
  });
  const restamped = stampEasServerId(vcard, "new-id");
  assert.equal(readEasServerIdFromVCard(restamped), "new-id");
});

test("appendApplicationDataFromVCard: outbound round-trip via the real WBXML encoder/decoder", () => {
  const vcard = [
    "BEGIN:VCARD",
    "VERSION:4.0",
    "N:User;Sample;;;",
    "EMAIL:user@example.invalid",
    "TEL;TYPE=work:+49301234567",
    "ORG:Example GmbH",
    "TITLE:President of space",
    "END:VCARD",
    "",
  ].join("\r\n");

  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendApplicationDataFromVCard({
    builder: w,
    vCard: vcard,
    asVersion: "16.1",
    separator: ",",
  });
  w.switchpage("AirSync");
  w.ctag();
  const node = parseAdNode(decodeWBXML(w.getBytes()));

  assert.equal(readPathFrom(node, ["FirstName"]), "Sample");
  assert.equal(readPathFrom(node, ["LastName"]), "User");
  assert.equal(readPathFrom(node, ["Email1Address"]), "user@example.invalid");
  assert.equal(readPathFrom(node, ["BusinessPhoneNumber"]), "+49301234567");
  assert.equal(readPathFrom(node, ["CompanyName"]), "Example GmbH");
  assert.equal(readPathFrom(node, ["JobTitle"]), "President of space");
});
