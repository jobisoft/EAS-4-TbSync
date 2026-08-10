/**
 * #347 — a calendar note (Body) must land so BOTH Thunderbird views agree.
 * TB stores a rich note as `DESCRIPTION;ALTREP="data:text/html,…":<text>`:
 * the tooltip reads the value, the editor reads the ALTREP. So an inbound
 * plaintext Body must CLEAR any stale ALTREP (the reported bug), and an
 * inbound HTML Body (Type 2 — we now request HTML) must set the ALTREP and
 * a converted plaintext value. Outbound mirrors it.
 *
 * convertToPlainText is stubbed by webext-env (a documented stand-in), so
 * the converted value here is that stub's output, not Thunderbird's.
 */

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

import ICAL from "../../src/vendor/ical.min.js";
import {
  applicationDataToIcal,
  appendApplicationDataFromIcal,
} from "../../src/modules/eas/calendar-codec.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { readPathFrom } from "../../src/modules/eas/wbxml-helpers.mjs";

before(() => ensureLoaded());

const COMMON = {
  serverID: "srv-note",
  asVersion: "16.1",
  defaultTimezone: "UTC",
  syncRecurrence: true,
  uid: null,
  userEmail: null,
};

const ADD = (bodyXml) => `<ApplicationData>
  <StartTime xmlns='Calendar'>20260801T100000Z</StartTime>
  <EndTime xmlns='Calendar'>20260801T103000Z</EndTime>
  <Subject xmlns='Calendar'>note-event</Subject>
  <UID xmlns='Calendar'>note-uid</UID>
  ${bodyXml}
</ApplicationData>`;

// Body Data is percent-encoded on the wire (the WBXML decoder escapes it),
// so a faithful capture fixture encodes it too - and it is what lets HTML
// live inside <Data> without confusing the capture parser.
const plainBody = (text) =>
  `<Body xmlns='AirSyncBase'><Type>1</Type><Data>${encodeURIComponent(text)}</Data></Body>`;
const htmlBody = (html) =>
  `<Body xmlns='AirSyncBase'><Type>2</Type><Data>${encodeURIComponent(html)}</Data></Body>`;

const master = (ics) =>
  new ICAL.Component(ICAL.parse(ics))
    .getAllSubcomponents("vevent")
    .find((v) => !v.getFirstProperty("recurrence-id"));

const descProp = (ics) => master(ics).getFirstProperty("description");

test("#347: an inbound plaintext Body clears a stale ALTREP", async () => {
  // Seed the exact bad state: a note whose editor-facing ALTREP is out of
  // date - what a prior in-Thunderbird edit leaves behind.
  const seeded = await applicationDataToIcal({
    adNode: parseAdNode(ADD(htmlBody("<b>old</b>"))),
    existingIcal: null,
    ...COMMON,
  });
  assert.ok(descProp(seeded).getParameter("altrep"), "seed has an ALTREP");

  const after = await applicationDataToIcal({
    adNode: parseAdNode(
      `<ApplicationData>${plainBody("new plain note")}</ApplicationData>`,
    ),
    existingIcal: seeded,
    ...COMMON,
  });
  const prop = descProp(after);
  assert.equal(
    prop.getFirstValue(),
    "new plain note",
    "the value is the new note",
  );
  assert.equal(
    prop.getParameter("altrep"),
    undefined,
    "the stale ALTREP is gone - the editor now reads the value too",
  );
});

test("an inbound HTML Body sets the ALTREP and a converted plaintext value", async () => {
  const after = await applicationDataToIcal({
    adNode: parseAdNode(ADD(htmlBody("line one<br>line two"))),
    existingIcal: null,
    ...COMMON,
  });
  const prop = descProp(after);
  assert.equal(
    prop.getFirstValue(),
    "line one\nline two",
    "the value is the plaintext rendering (tooltip source)",
  );
  assert.equal(
    prop.getParameter("altrep"),
    "data:text/html," + encodeURIComponent("line one<br>line two"),
    "the HTML rides along as the ALTREP (editor source)",
  );
});

test("a payload with no Body leaves the note and its ALTREP untouched", async () => {
  const seeded = await applicationDataToIcal({
    adNode: parseAdNode(ADD(htmlBody("<i>keep</i>"))),
    existingIcal: null,
    ...COMMON,
  });
  const before = descProp(seeded);

  const after = await applicationDataToIcal({
    adNode: parseAdNode(
      `<ApplicationData><Subject xmlns='Calendar'>note-event renamed</Subject></ApplicationData>`,
    ),
    existingIcal: seeded,
    ...COMMON,
  });
  const prop = descProp(after);
  assert.equal(prop.getFirstValue(), before.getFirstValue(), "value unchanged");
  assert.equal(
    prop.getParameter("altrep"),
    before.getParameter("altrep"),
    "ALTREP unchanged - a delta that omits the Body changes nothing",
  );
});

test("outbound: an ALTREP note goes out as Type 2 HTML, a plain note as Type 1", () => {
  const emit = (ics) => {
    const w = createWBXML("AirSync");
    w.otag("Add");
    w.otag("ApplicationData");
    appendApplicationDataFromIcal({
      builder: w,
      ical: ics,
      asVersion: "16.1",
      defaultTimezone: "UTC",
      syncRecurrence: true,
      userEmail: null,
    });
    w.switchpage("AirSync");
    w.ctag();
    w.ctag();
    const node = parseAdNode(decodeWBXML(w.getBytes()));
    return node.children.find((c) => c.tagName === "ApplicationData");
  };

  const html = [
    "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT",
    "UID:o1\r\nDTSTAMP:20260801T090000Z",
    "DTSTART:20260801T100000Z\r\nDTEND:20260801T103000Z\r\nSUMMARY:o",
    'DESCRIPTION;ALTREP="data:text/html,%3Cb%3Ehi%3C/b%3E":hi',
    "END:VEVENT\r\nEND:VCALENDAR\r\n",
  ].join("\r\n");
  const adHtml = emit(html);
  assert.equal(readPathFrom(adHtml, ["Body", "Type"]), "2");
  assert.equal(readPathFrom(adHtml, ["Body", "Data"]), "<b>hi</b>");

  const plain = html.replace(
    /DESCRIPTION;ALTREP=[^\n]*/,
    "DESCRIPTION:just text",
  );
  const adPlain = emit(plain);
  assert.equal(readPathFrom(adPlain, ["Body", "Type"]), "1");
  assert.equal(readPathFrom(adPlain, ["Body", "Data"]), "just text");
});
