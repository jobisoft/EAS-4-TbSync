// contact-codec.mjs: EAS Contacts <-> vCard 4.0. A representative
// slice, not exhaustive - the module maps a large number of individual
// fields (prefixed phones, four Custom slots, half a dozen pass-through
// properties); this covers name/email/phone/organization plus the
// serverID/merge-aware/outbound mechanics shared by all of them, which
// is where a regression is most likely to actually bite. Extending
// field-by-field coverage is listed as open work in TEST-PLAN.md.

import { test } from "vitest";
import "../support/webext-shim.mjs";
import ICAL from "../../src/vendor/ical.min.js";
import assert from "node:assert/strict";
import {
  applicationDataToVCard,
  appendApplicationDataFromVCard,
  readEasServerIdFromVCard,
  stampEasServerId,
} from "../../src/modules/eas/contact-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "../support/xml-node.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

const ADD_CONTACT = `<ApplicationData>
  <FirstName xmlns='Contacts'>Tomas</FirstName>
  <LastName xmlns='Contacts'>Kovacik</LastName>
  <Email1Address xmlns='Contacts'>Tomas%20Kovacik%20%3Ckovacik%40dgtfactory.com%3E</Email1Address>
  <BusinessPhoneNumber xmlns='Contacts'>00421907813840</BusinessPhoneNumber>
  <CompanyName xmlns='Contacts'>DGT%20factory%2C%20a.%20s.</CompanyName>
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
    "Kovacik",
    "Tomas",
    "",
    "",
    "",
  ]);
  assert.equal(comp.getFirstPropertyValue("email"), "kovacik@dgtfactory.com");
  const tel = comp.getFirstProperty("tel");
  assert.equal(tel.getFirstValue(), "00421907813840");
  assert.equal(tel.getParameter("type"), "work");
  assert.deepEqual(comp.getFirstPropertyValue("org"), "DGT factory, a. s.");
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
      `<ApplicationData><MobilePhoneNumber xmlns='Contacts'>00421900000000</MobilePhoneNumber></ApplicationData>`,
    ),
    existingVcard: afterAdd,
    serverID: "server-id-contact-1",
    asVersion: "16.1",
    separator: ",",
    uid: null,
  });

  const comp = new ICAL.Component(ICAL.parse(afterPhoneOnlyChange));
  assert.deepEqual(comp.getFirstPropertyValue("n"), [
    "Kovacik",
    "Tomas",
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
  assert.equal(tels[0].getFirstValue(), "00421900000000");
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
    "N:Kovacik;Tomas;;;",
    "EMAIL:kovacik@dgtfactory.com",
    "TEL;TYPE=work:00421907813840",
    "ORG:DGT factory, a. s.",
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

  assert.equal(readPathFrom(node, ["FirstName"]), "Tomas");
  assert.equal(readPathFrom(node, ["LastName"]), "Kovacik");
  assert.equal(readPathFrom(node, ["Email1Address"]), "kovacik@dgtfactory.com");
  assert.equal(readPathFrom(node, ["BusinessPhoneNumber"]), "00421907813840");
  assert.equal(readPathFrom(node, ["CompanyName"]), "DGT factory, a. s.");
  assert.equal(readPathFrom(node, ["JobTitle"]), "President of space");
});
