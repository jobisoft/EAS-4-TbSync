/**
 * wbxml-helpers' readers (readPath / readPathFrom / readChildTexts) and
 * the percent-encoding + UTF-8 reinterpret round-trip the module's own
 * docblock warns about ("ü" surfacing as "Ã¼" without the second decode
 * step). The round-trip case goes through the REAL WBXML encoder and
 * decoder, so it checks the actual wire format, not the helper in
 * isolation.
 *
 * Cases ported from PR #345 (tomaskovacik).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  readPath,
  readPathFrom,
  readChildTexts,
} from "../../src/modules/eas/wbxml-helpers.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "./support/ad-node.mjs";

const docFromXml = (xml) => ({ documentElement: parseAdNode(xml) });

test("readPath walks the path and returns the leaf's decoded text", () => {
  const doc = docFromXml(`<Root><A><B>hello</B></A></Root>`);
  assert.equal(readPath(doc, ["A", "B"]), "hello");
});

test("readPath: a missing step returns null, never throws", () => {
  const doc = docFromXml(`<Root><A/></Root>`);
  assert.equal(readPath(doc, ["A", "B", "C"]), null);
});

test("readPath: no documentElement returns null", () => {
  assert.equal(readPath({ documentElement: null }, ["A"]), null);
  assert.equal(readPath(null, ["A"]), null);
});

test("readPath: the empty path reads the root's own text", () => {
  const doc = docFromXml(`<Root>direct-text</Root>`);
  assert.equal(readPath(doc, []), "direct-text");
});

test("readPathFrom starts at the given node", () => {
  const doc = docFromXml(`<Root><Item><Name>widget</Name></Item></Root>`);
  const item = doc.documentElement.children.find((c) => c.tagName === "Item");
  assert.equal(readPathFrom(item, ["Name"]), "widget");
});

test("readPathFrom: a null start returns null", () => {
  assert.equal(readPathFrom(null, ["Name"]), null);
});

test("readChildTexts collects matching direct children, in order", () => {
  const doc = docFromXml(
    `<Categories><Category>Work</Category><Category>Personal</Category></Categories>`,
  );
  assert.deepEqual(readChildTexts(doc.documentElement, "Category"), [
    "Work",
    "Personal",
  ]);
});

test("readChildTexts skips empty entries and non-matching siblings", () => {
  const doc = docFromXml(
    `<Categories><Category>Work</Category><Category></Category><Other>x</Other></Categories>`,
  );
  assert.deepEqual(readChildTexts(doc.documentElement, "Category"), ["Work"]);
});

test("readChildTexts: a node without children yields an empty array", () => {
  assert.deepEqual(readChildTexts({ children: undefined }, "Category"), []);
});

test("accented text survives the real WBXML wire round-trip", () => {
  const original = "Tomáš Kováčik — café";
  const w = createWBXML("AirSync");
  w.otag("Add");
  w.atag("ServerId", original);
  w.ctag();
  const doc = docFromXml(decodeWBXML(w.getBytes()));
  assert.equal(readPath(doc, ["ServerId"]), original);
});
