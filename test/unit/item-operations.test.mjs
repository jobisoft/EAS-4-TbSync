/**
 * The batched ItemOperations.Fetch: one request carrying a <Fetch> per item,
 * results keyed by the ServerId each response node echoes.
 *
 * The builder is tested through the real WBXML round trip - build with
 * createWBXML, decode with decodeWBXML - so a codepage mistake fails here
 * rather than live. The reader is fed capture-shaped nodes from parseAdNode,
 * the same lightweight shape wbxml-helpers consume, which is exactly why it
 * walks `children` instead of using getElementsByTagName.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  buildFetchBody,
  readFetchResults,
} from "../../src/modules/eas/item-operations.mjs";
import { decodeWBXML } from "../../src/modules/wbxml.mjs";
import { parseAdNode } from "./support/ad-node.mjs";

const decodeToNode = (bytes) =>
  parseAdNode(decodeWBXML(bytes).replace(/^<\?xml[^?]*\?>/, ""));

const fetchNodes = (root) => {
  const out = [];
  const walk = (n) => {
    for (const c of n.children ?? []) {
      if (c.tagName === "Fetch") out.push(c);
      walk(c);
    }
  };
  walk(root);
  return out;
};

const leaf = (node, tag) => {
  for (const c of node.children ?? []) {
    if (c.tagName === tag) return c.textContent;
    const deeper = leaf(c, tag);
    if (deeper != null) return deeper;
  }
  return null;
};

test("one id builds one Fetch, three ids build three", () => {
  const one = decodeToNode(
    buildFetchBody({ collectionId: "C7", serverIDs: ["a:1"], bodyType: "2" }),
  );
  assert.equal(fetchNodes(one).length, 1);

  const three = decodeToNode(
    buildFetchBody({
      collectionId: "C7",
      serverIDs: ["a:1", "b:2", "c:3"],
      bodyType: "2",
    }),
  );
  const nodes = fetchNodes(three);
  assert.equal(nodes.length, 3, "one <Fetch> per id");
  for (const n of nodes) {
    assert.equal(leaf(n, "Store"), "Mailbox");
    assert.equal(leaf(n, "CollectionId"), "C7");
    assert.equal(leaf(n, "Type"), "2", "each Fetch carries the BodyPreference");
  }
});

test("ids with URL-hostile characters survive the wire round trip", () => {
  // Real ServerIds carry ':' and '+' (Exchange) and '%' once encoded; the
  // decoder percent-escapes inline strings, so what matters is that the
  // decoded capture shows the original id.
  const ids = ["U2f1ad:57a5+ff", "3:22"];
  const root = decodeToNode(
    buildFetchBody({ collectionId: "3", serverIDs: ids, bodyType: "1" }),
  );
  const seen = fetchNodes(root).map((n) =>
    decodeURIComponent(leaf(n, "ServerId")),
  );
  assert.deepEqual(seen, ids);
});

const RESPONSE = (inner) =>
  parseAdNode(`<ItemOperations><Status>1</Status><Response>${inner}</Response></ItemOperations>`);

const FETCH = (serverId, subject, status = "1") =>
  `<Fetch><Status>${status}</Status>` +
  (serverId ? `<ServerId>${serverId}</ServerId>` : "") +
  (subject
    ? `<Properties><Subject>${subject}</Subject></Properties>`
    : "") +
  `</Fetch>`;

test("results key by the echoed ServerId, whatever the order", () => {
  const root = RESPONSE(FETCH("b", "second") + FETCH("a", "first"));
  const { status, items } = readFetchResults(root, ["a", "b"]);
  assert.equal(status, "1");
  assert.equal(items.size, 2);
  assert.equal(leaf(items.get("a"), "Subject"), "first");
  assert.equal(leaf(items.get("b"), "Subject"), "second");
});

test("a failed Fetch is absent while its siblings survive", () => {
  const root = RESPONSE(
    FETCH("a", "good") + FETCH("b", "bad", "8") + FETCH("c", null),
  );
  const { items } = readFetchResults(root, ["a", "b", "c"]);
  assert.deepEqual(
    [...items.keys()],
    ["a"],
    "Status 8 and a Fetch with no Properties are both dropped",
  );
});

test("a non-1 top-level Status yields no items but reports itself", () => {
  const root = parseAdNode(
    `<ItemOperations><Status>2</Status></ItemOperations>`,
  );
  const { status, items } = readFetchResults(root, ["a"]);
  assert.equal(status, "2");
  assert.equal(items.size, 0);
  assert.deepEqual(readFetchResults(null, ["a"]), {
    status: null,
    items: new Map(),
  });
});

test("a lone un-echoed Fetch falls back to position", () => {
  // A server that omits the ServerId echo works today only because the
  // singular caller took the first node blindly; that tolerance must
  // survive the plural form.
  const root = RESPONSE(FETCH(null, "only"));
  const { items } = readFetchResults(root, ["the-id"]);
  assert.equal(leaf(items.get("the-id"), "Subject"), "only");

  // But with a count mismatch there is no honest key, so nothing is guessed.
  const two = RESPONSE(FETCH(null, "one"));
  assert.equal(readFetchResults(two, ["x", "y"]).items.size, 0);
});
