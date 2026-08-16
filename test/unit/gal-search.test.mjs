/**
 * Unit tests for what a GAL Search response says about its own completeness.
 *
 * [MS-ASCMD] Search: the server returns as many entries as `<Range>` asks
 * for - 100 by default, which is what we ask - and "MUST also indicate the
 * total number of entries that are found" in `<Total>`.
 *
 * That count is the whole point. `isCompleteResult: true` licenses
 * Thunderbird to stop asking and narrow the set locally as the user types
 * on (`AbAutoCompleteSearch.sys.mjs`: a longer string that extends a
 * RESULT_SUCCESS is filtered in JS, and the directory is re-queried only
 * when something reported the result incomplete). Claiming completeness
 * over a truncated answer therefore makes everyone past the hundredth
 * match permanently unreachable, with nothing to show it happened.
 *
 * Counting rows cannot substitute for reading `<Total>`: a GAL holding
 * exactly 100 matches and one holding 400 both return 100 rows.
 *
 * The response shape below is a capture from ekir (Z-Push, AS 14.1),
 * trimmed of contact details - <Total> sits inside <Store> beside <Range>.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();

const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);

/** A Search response holding `n` results and declaring `total` found. */
function searchReply(n, total) {
  const rows = Array.from(
    { length: n },
    (_, i) =>
      `<Result><Properties>` +
      `<DisplayName>Person ${i}</DisplayName>` +
      `<EmailAddress>p${i}@example.invalid</EmailAddress>` +
      `</Properties></Result>`,
  ).join("");
  const totalTag = total === null ? "" : `<Total>${total}</Total>`;
  return `<?xml version="1.0" encoding="utf-8"?>
    <Search><Status>1</Status><Response><Store><Status>1</Status>
      ${rows}<Range>0-${Math.max(n - 1, 0)}</Range>${totalTag}
    </Store></Response></Search>`;
}

/** The two things reply parsing asks a Document for. */
function fakeDoc(xml) {
  const root = parseAdNode(xml);
  const all = [];
  (function walk(node) {
    all.push(node);
    for (const c of node.children ?? []) walk(c);
  })(root);
  return {
    documentElement: root,
    getElementsByTagName: (tag) => all.filter((n) => n.tagName === tag),
  };
}

/* Installed once, before the module under test is imported: ESM caches the
 * import, so a per-test mock would bind it to the first test's stub. */
let reply = "";
const real = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...real,
    easRequest: async () => ({ doc: fakeDoc(reply) }),
  },
});
const { runGalSearch } = await import("../../src/modules/eas/gal-search.mjs");

const ACCOUNT = { serverURL: "https://example.invalid", user: "u" };
const run = (xml) => {
  reply = xml;
  return runGalSearch({
    account: ACCOUNT,
    asVersion: "14.1",
    query: "biel",
    companyName: "Example",
  });
};

test("a truncated answer reports the larger total", async () => {
  // The case the fix exists for: the server found 340 and sent the 100 we
  // asked for. Nothing in the rows themselves says so.
  const { results, total } = await run(searchReply(100, 340));
  assert.equal(results.length, 100);
  assert.equal(total, 340);
  assert.ok(total > results.length, "caller must be able to see the shortfall");
});

test("a complete answer reports a total equal to what it sent", async () => {
  // Measured on ekir: Total=10 with Range=0-9, ten rows.
  const { results, total } = await run(searchReply(10, 10));
  assert.equal(results.length, 10);
  assert.equal(total, 10);
});

test("exactly the range ceiling is not assumed to be truncated", async () => {
  // 100 rows and 100 found. Row-counting would call this truncated; the
  // server says otherwise, and the server is the one who knows.
  const { results, total } = await run(searchReply(100, 100));
  assert.equal(total, 100);
  assert.equal(results.length, 100);
  assert.ok(!(total > results.length));
});

test("an absent Total is reported as null, not as zero", async () => {
  // Measured: some servers answer an empty search with no <Total> at all.
  // Zero would read as "found none" and, worse, arithmetic on it would say
  // the answer is complete for the wrong reason.
  const { results, total } = await run(searchReply(0, null));
  assert.deepEqual(results, []);
  assert.equal(total, null);
});

test("the completeness rule: only a stated Total can license reuse", async () => {
  // The rule gal.mjs applies, asserted here because it is the whole point
  // of reading Total at all. Absent means unknown, and unknown must not
  // license Thunderbird to stop asking: the cost of being wrong that way
  // is a colleague nobody can find, against a few more requests.
  const complete = (total, delivered) => !(total == null || total > delivered);
  assert.equal(complete(340, 100), false, "found more than it sent");
  assert.equal(complete(10, 10), true, "sent everything it found");
  assert.equal(complete(100, 100), true, "exactly the ceiling, and stated");
  assert.equal(complete(null, 0), false, "said nothing, empty answer");
  assert.equal(complete(null, 7), false, "said nothing, some rows");
});

test("a row we cannot map still counts as delivered", async () => {
  // `readProperties` drops a Result carrying neither a name nor an address.
  // Completeness is about the server's match set, so judging it against the
  // rows we could map would call a complete answer truncated the moment one
  // row was unusable - and then never stop asking.
  reply = `<?xml version="1.0" encoding="utf-8"?>
    <Search><Status>1</Status><Response><Store><Status>1</Status>
      <Result><Properties><DisplayName>Real Person</DisplayName>
        <EmailAddress>real@example.invalid</EmailAddress></Properties></Result>
      <Result><Properties><Office>no name, no address</Office></Properties></Result>
      <Range>0-1</Range><Total>2</Total>
    </Store></Response></Search>`;
  const { results, total, delivered } = await runGalSearch({
    account: ACCOUNT,
    asVersion: "14.1",
    query: "biel",
    companyName: "Example",
  });
  assert.equal(results.length, 1, "only the usable row is mapped");
  assert.equal(delivered, 2, "but the server sent two");
  assert.equal(total, 2);
  assert.ok(!(total > delivered), "so this is complete, not truncated");
});

test("an empty body yields no results and no total", async () => {
  reply = "<Search><Status>1</Status></Search>";
  const { results, total } = await runGalSearch({
    account: ACCOUNT,
    asVersion: "14.1",
    query: "biel",
    companyName: "Example",
  });
  assert.deepEqual(results, []);
  assert.equal(total, null);
});
