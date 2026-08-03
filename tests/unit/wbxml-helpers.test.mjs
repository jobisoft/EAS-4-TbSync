// wbxml-helpers.mjs directly: readPath/readPathFrom/readChildTexts, and
// specifically the percent-encoding + UTF-8 reinterpret round-trip the
// module's own docblock warns about ("ü" (UTF-8 0xC3 0xBC) surfacing as
// "Ã¼" without the second decode step) - exercised through the REAL
// WBXML encoder/decoder (wbxml.mjs), not a hand-crafted percent-encoded
// fixture, so this is a faithful end-to-end check of the actual wire
// round-trip, not just the helper's own logic in isolation.

import { test } from "vitest";
import assert from "node:assert/strict";
import {
  readPath,
  readPathFrom,
  readChildTexts,
} from "../../src/modules/eas/wbxml-helpers.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "../support/xml-node.mjs";

function docFromXml(xml) {
  return { documentElement: parseAdNode(xml) };
}

test("readPath: walks a multi-level path and returns the leaf's decoded text", () => {
  const doc = docFromXml(
    `<Root><A><B>hello</B></A></Root>`,
  );
  assert.equal(readPath(doc, ["A", "B"]), "hello");
});

test("readPath: a missing intermediate step returns null, not a throw", () => {
  const doc = docFromXml(`<Root><A/></Root>`);
  assert.equal(readPath(doc, ["A", "B", "C"]), null);
});

test("readPath: no documentElement at all returns null", () => {
  assert.equal(readPath({ documentElement: null }, ["A"]), null);
  assert.equal(readPath(null, ["A"]), null);
});

test("readPath: an empty path array returns the root's own text", () => {
  const doc = docFromXml(`<Root>direct-text</Root>`);
  assert.equal(readPath(doc, []), "direct-text");
});

test("readPathFrom: starts at a given node instead of the document root", () => {
  const doc = docFromXml(`<Root><Item><Name>widget</Name></Item></Root>`);
  const item = doc.documentElement.children.find((c) => c.tagName === "Item");
  assert.equal(readPathFrom(item, ["Name"]), "widget");
});

test("readPathFrom: a null starting node returns null immediately", () => {
  assert.equal(readPathFrom(null, ["Name"]), null);
});

test("readChildTexts: collects every direct child matching the tag, in order", () => {
  const doc = docFromXml(
    `<Categories><Category>Work</Category><Category>Personal</Category></Categories>`,
  );
  assert.deepEqual(readChildTexts(doc.documentElement, "Category"), [
    "Work",
    "Personal",
  ]);
});

test("readChildTexts: skips empty/whitespace-only entries and non-matching siblings", () => {
  const doc = docFromXml(
    `<Categories><Category>Work</Category><Category></Category><Other>ignored</Other></Categories>`,
  );
  assert.deepEqual(readChildTexts(doc.documentElement, "Category"), ["Work"]);
});

test("readChildTexts: no children at all returns an empty array, not a throw", () => {
  assert.deepEqual(readChildTexts({ children: undefined }, "Category"), []);
});

test("percent-encoding + UTF-8 reinterpret round-trip: accented text survives the real WBXML wire format", () => {
  const original = "Tomáš Kováčik — café";
  const w = createWBXML("AirSync");
  w.otag("Add");
  w.atag("ServerId", original);
  w.ctag();
  const xml = decodeWBXML(w.getBytes());
  const doc = docFromXml(xml);
  assert.equal(readPath(doc, ["ServerId"]), original);
});
