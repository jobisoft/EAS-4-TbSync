/**
 * #347 — a calendar note (Body) must land so BOTH Thunderbird views agree.
 * TB stores a rich note as `DESCRIPTION;ALTREP="data:text/html,…":<text>`:
 * the tooltip reads the value, the editor reads the ALTREP. So an inbound
 * plaintext Body must CLEAR any stale ALTREP (the reported bug), and an
 * inbound HTML Body must set the ALTREP and a converted plaintext value.
 * Outbound mirrors it: a note carrying an ALTREP goes back as Type 2.
 *
 * The codec is handed a Body it can trust - the sync runner reads
 * NativeBodyType and re-fetches an item the server holds as HTML before it
 * gets here - so these tests exercise placement, not that decision.
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

test("a plaintext payload replaces an existing ALTREP note entirely", async () => {
  // A plaintext Body reaching the codec means the server holds the note as
  // plain text: the sync runner re-fetches anything whose NativeBodyType
  // says HTML, so this branch never sees a flattened rich note. The whole
  // note is therefore the text, and an ALTREP left beside it would be the
  // stale editor copy of #347.
  const seeded = await applicationDataToIcal({
    adNode: parseAdNode(ADD(htmlBody("<b>mine</b>"))),
    existingIcal: null,
    ...COMMON,
  });
  assert.ok(descProp(seeded).getParameter("altrep"), "seeded with an ALTREP");

  const after = await applicationDataToIcal({
    adNode: parseAdNode(
      `<ApplicationData>${plainBody("now plain")}</ApplicationData>`,
    ),
    existingIcal: seeded,
    ...COMMON,
  });
  const prop = descProp(after);
  assert.equal(prop.getFirstValue(), "now plain", "the value is the payload");
  assert.equal(
    prop.getParameter("altrep"),
    undefined,
    "no ALTREP survives - both readers show the same plain text",
  );
});

test("#262: carriage returns never reach the stored note", async () => {
  // RFC 5545 §3.3.11 escapes a line break in a TEXT value and has no escape
  // for a carriage return, and ical.js writes one straight through - so a CR
  // ends the property mid-value for any reader that splits on it, and the
  // rest of the note becomes a line belonging to nothing. Exchange sends
  // CRLF between the lines of a note and after the last one, so this is the
  // ordinary case, not an exotic one.
  const plain = await applicationDataToIcal({
    adNode: parseAdNode(ADD(plainBody("first\r\nsecond\r\n"))),
    existingIcal: null,
    ...COMMON,
  });
  const value = descProp(plain).getFirstValue();
  assert.equal(value, "first\nsecond\n", "CRLF became a bare newline");
  assert.ok(!value.includes("\r"), "no carriage return survives");
  assert.ok(
    !plain
      .split("\r\n")
      .some(
        (line, i) =>
          i && line && !/^[ \t]/.test(line) && !/^[A-Z-]+[;:]/.test(line),
      ),
    "every line of the blob is a property or a fold continuation",
  );

  // The HTML branch trims its rendering as well - it is the tooltip's copy,
  // never what travels - so a trailing break cannot end the property either.
  const rich = await applicationDataToIcal({
    adNode: parseAdNode(ADD(htmlBody("one<br>two"))),
    existingIcal: null,
    ...COMMON,
  });
  const richValue = descProp(rich).getFirstValue();
  assert.ok(!richValue.includes("\r"), "no carriage return in the rendering");
  assert.equal(
    richValue,
    richValue.replace(/\s+$/, ""),
    "and no trailing break",
  );
});

test("#262: a CRLF server gets its line endings back, once", async () => {
  // Normalising CRLF away on the way in is not optional - a CR in the value
  // makes the blob unparseable - but the value then no longer matches what
  // the server holds, and pushing a value it did not give us is an edit it
  // acts on. The shape is remembered per item and restored on the way out.
  const stored = await applicationDataToIcal({
    adNode: parseAdNode(ADD(plainBody("first\r\nsecond"))),
    existingIcal: null,
    ...COMMON,
  });
  assert.equal(
    descProp(stored).getFirstValue(),
    "first\nsecond",
    "stored with bare newlines",
  );

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

  assert.equal(
    readPathFrom(emit(stored), ["Body", "Data"]),
    "first\r\nsecond",
    "the server gets the CRLF it sent us back",
  );

  // A line the user adds later is bare - it must not come back doubled.
  const edited = stored.replace("first\\nsecond", "first\\nsecond\\nthird");
  assert.equal(
    readPathFrom(emit(edited), ["Body", "Data"]),
    "first\r\nsecond\r\nthird",
    "every newline expands exactly once",
  );

  // A server that sends bare newlines is left alone.
  const lf = await applicationDataToIcal({
    adNode: parseAdNode(ADD(plainBody("alpha\nbeta"))),
    existingIcal: null,
    ...COMMON,
  });
  assert.equal(
    readPathFrom(emit(lf), ["Body", "Data"]),
    "alpha\nbeta",
    "no CRLF is invented for a server that never used it",
  );
});
