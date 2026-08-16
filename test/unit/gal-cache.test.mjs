/**
 * Unit tests for what the GAL search callback decides before it asks the
 * server: whether it can answer from an earlier answer, and whether it may
 * tell Thunderbird to stop asking.
 *
 * Every one of these is a judgement made from measured wire behaviour, and
 * none of them is visible in the answer itself - a wrong call shows up as a
 * colleague who cannot be found, months later, on somebody else's server.
 *
 * The callback is reached the way Thunderbird reaches it: `enableGal`
 * registers it through `addressBooks.provider.onSearchRequest`, and the
 * fake below keeps the function it was handed.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

/** Capture the listener `enableGal` registers, and count what it asks the
 *  server. `answers` maps a query to the reply `runGalSearch` gives. */
let nextAccountId = 0;

async function galUnderTest(answers, { accountId } = {}) {
  // A fresh id per test: `listeners` is module state and `enableGal`
  // no-ops for an account that already has one, so reusing an id would
  // silently register nothing.
  accountId = accountId ?? `acct-${(nextAccountId += 1)}`;
  let search;
  // `enableGal` builds the directory's display name while assembling the
  // addListener arguments, and that reaches for i18n. The env deliberately
  // ships no `browser.i18n`, and the throw would be swallowed by the
  // registration's own try/catch - leaving no listener and a confusing
  // "search is not a function".
  globalThis.browser.i18n = { getMessage: () => "Global Address List" };
  globalThis.messenger.addressBooks = {
    onUpdated: { addListener() {} },
    provider: {
      onSearchRequest: {
        addListener: (fn) => {
          search = fn;
        },
        removeListener() {},
      },
    },
    delete: async () => {},
  };

  const asked = [];
  const account = {
    accountId,
    accountName: "Example",
    custom: { allowedEasCommands: ["Sync", "Search"] },
  };
  const provider = {
    getAccount: async () => ({ account }),
    reportEventLog() {},
    runGalSearch: async ({ query }) => {
      asked.push(query);
      const reply = answers[query];
      if (typeof reply === "function") return reply();
      return reply ?? { results: [], total: null, delivered: 0 };
    },
  };

  const { enableGal } = await import("../../src/modules/gal.mjs");
  await enableGal({ provider, account });
  return { search: (q) => search(null, q), asked };
}

const person = (name) => ({
  DisplayName: name,
  PrimaryEmail: `${name.toLowerCase()}@example.invalid`,
});

/** A complete answer: the server found exactly what it sent. */
const complete = (names) => ({
  results: names.map(person),
  total: names.length,
  delivered: names.length,
});

test("a repeated query is answered without asking again", async () => {
  const { search, asked } = await galUnderTest({ abc: complete(["Abel"]) });
  const first = await search("abc");
  const second = await search("abc");
  assert.deepEqual(asked, ["abc"], "the server was asked once");
  assert.equal(first.results.length, 1);
  assert.deepEqual(second.results, first.results);
});

test("a longer query is narrowed from a complete answer, not asked", async () => {
  // What the address book search needs: it has no narrowing of its own, so
  // every keystroke would otherwise be a round trip - ~1.7s each on a
  // measured Z-Push GAL.
  const { search, asked } = await galUnderTest({
    bie: complete(["Bieling", "Bierschenk", "Bietz"]),
  });
  await search("bie");
  const narrowed = await search("biel");
  assert.deepEqual(asked, ["bie"], "the longer query never reached the server");
  assert.deepEqual(
    narrowed.results.map((c) => c.DisplayName),
    ["Bieling"],
  );
  assert.equal(narrowed.isCompleteResult, true, "complete by construction");
});

test("a truncated answer is never narrowed from", async () => {
  // The safety-critical case. The server found 340 and sent 2; filtering
  // those 2 would answer the longer query with a slice of a slice and hide
  // everyone it held back. It must go and ask.
  const { search, asked } = await galUnderTest({
    bie: { results: [person("Bieling"), person("Bietz")], total: 340, delivered: 2 },
    biel: complete(["Bieling", "Bielz"]),
  });
  const wide = await search("bie");
  assert.equal(wide.isCompleteResult, false, "truncated is not complete");
  const narrow = await search("biel");
  assert.deepEqual(asked, ["bie", "biel"], "the server was asked again");
  assert.equal(narrow.results.length, 2, "and gave the answer we could not");
});

test("an answer with no stated total is never narrowed from", async () => {
  // Measured on Exchange 16.1: a three-character query is declined with a
  // bare <Result/> and no <Total>, and the four-character one then returns
  // people. Without a total we cannot tell "that is everyone" from "I did
  // not look", so the longer query has to be asked.
  const { search, asked } = await galUnderTest({
    cvj: { results: [], total: null, delivered: 0 },
    cvjm: complete(["CVJM Buero", "Flohmarkt CVJM", "CVJM Bonn"]),
  });
  const declined = await search("cvj");
  assert.equal(declined.isCompleteResult, false);
  const found = await search("cvjm");
  assert.deepEqual(asked, ["cvj", "cvjm"]);
  assert.equal(found.results.length, 3, "the people the short query missed");
});

test("every extension is served by the one answer we paid for", async () => {
  // Narrowed answers are not themselves cached - only what the server
  // actually told us is. So each further character filters the same
  // complete set rather than a filtered copy of a filtered copy, and the
  // server is asked exactly once however far the user types.
  const { search, asked } = await galUnderTest({
    bie: complete(["Bieling", "Bietz"]),
  });
  await search("bie");
  const four = await search("biel");
  const six = await search("bielin");
  assert.deepEqual(asked, ["bie"], "one request for the whole sequence");
  assert.deepEqual(
    four.results.map((c) => c.DisplayName),
    ["Bieling"],
  );
  assert.deepEqual(
    six.results.map((c) => c.DisplayName),
    ["Bieling"],
  );
  assert.equal(six.isCompleteResult, true);
});

test("below the length floor nothing is asked, and nothing is final", async () => {
  // Our own policy, not the server's: three characters before we search at
  // all. `false` keeps Thunderbird asking, so the search fires the moment
  // the query is long enough instead of a character later (#344).
  const { search, asked } = await galUnderTest({});
  const short = await search("ab");
  assert.deepEqual(asked, [], "no request below the floor");
  assert.deepEqual(short.results, []);
  assert.equal(short.isCompleteResult, false);
});

test("a failed search is not remembered as an answer", async () => {
  // One timeout must not silence the GAL for the whole cache lifetime.
  let attempt = 0;
  const { search, asked } = await galUnderTest({
    abc: () => {
      attempt += 1;
      if (attempt === 1) throw new Error("network went away");
      return complete(["Abel"]);
    },
  });
  const failed = await search("abc");
  assert.deepEqual(failed.results, [], "the failure answers empty");
  assert.equal(failed.isCompleteResult, false, "and never claims to be final");
  const retried = await search("abc");
  assert.deepEqual(asked, ["abc", "abc"], "the next search really asks again");
  assert.equal(retried.results.length, 1);
});
