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
  return {
    // The autocomplete's async branch passes a null query; every other
    // caller passes a real one, and must always be answered.
    search: (q) => search(null, q),
    query: (q) => search(null, q, "?(and(DisplayName,c,%s))"),
    asked,
  };
}

/** What a search did: its answer, or the fact that it is being withheld.
 *
 *  An empty answer is deliberately never returned - see `answerOrWithhold`
 *  in gal.mjs - so awaiting one hangs until a later search releases it.
 *  Tests must be able to say "this one is held" without waiting forever.
 *
 *  The race has to be against a real delay, not a resolved promise. The
 *  callback is async and awaits the account and the search before it can
 *  answer, so a microtask-sized racer wins against every answer and reports
 *  everything as held - which is exactly the false pass this helper had
 *  when it was written that way. */
const HELD = Symbol("withheld");
function settledOr(promise, ms = 50) {
  return Promise.race([
    promise,
    new Promise((resolve) => setTimeout(() => resolve(HELD), ms)),
  ]);
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
  const declined = search("cvj");
  assert.equal(await settledOr(declined), HELD, "the empty answer is held");
  const found = await search("cvjm");
  assert.deepEqual(asked, ["cvj", "cvjm"], "the server was asked both times");
  assert.equal(found.results.length, 3, "the people the short query missed");
  // `declined` stays held - nothing releases it but its own timer.
  assert.equal(await settledOr(declined), HELD, "still held");
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

test("below the length floor nothing is asked, and the answer is withheld", async () => {
  // Our own policy, not the server's: three characters before we search at
  // all. The empty answer is not returned, because Thunderbird forgets the
  // "ask me again" flag on an empty result and then never asks again
  // (#344) - holding the promise keeps the search ongoing, which is the
  // state that makes it ask.
  const { search, asked } = await galUnderTest({ abc: complete(["Abel"]) });
  const short = search("ab");
  assert.equal(await settledOr(short), HELD, "held, not answered");
  assert.deepEqual(asked, [], "and nothing was asked of the server");

  // Nothing releases it early - it is let go by its own timer - so a later
  // search of any kind leaves it alone.
  const next = await search("abc");
  assert.equal(next.results.length, 1, "and the real query is answered");
  assert.equal(await settledOr(short), HELD, "the held one is still held");
});

test("a hold belongs to its own search and nothing can end it early", async () => {
  // The API says nothing about who is asking - `onSearchRequest` hands over
  // the address book, the string and the query, and no caller identity - so
  // two compose windows are indistinguishable. A hold is therefore released
  // only by its own timer: anything keyed per account would let one window
  // end another's hold, its search would complete, and the bug would come
  // back there.
  const { search } = await galUnderTest({
    cvj: { results: [], total: null, delivered: 0 },
    abc: { results: [], total: null, delivered: 0 },
  });
  const first = search("cvj");
  assert.equal(await settledOr(first), HELD, "held");

  search("abc"); // another window, another question
  assert.equal(await settledOr(first), HELD, "still held");

  search("cvj"); // and even the same question again
  assert.equal(await settledOr(first), HELD, "still held");
});

test("contacts.query is answered even when the answer is empty", async () => {
  // Only the autocomplete's async branch is withheld from. `contacts.query`
  // awaits one promise per address book and would hang for good otherwise -
  // and it is nothing to do with the bug being worked around.
  const { query, asked } = await galUnderTest({
    cvj: { results: [], total: null, delivered: 0 },
  });
  const answered = await settledOr(query("cvj"));
  assert.notEqual(answered, HELD, "a real query is never withheld");
  assert.deepEqual(answered, { results: [], isCompleteResult: false });
  assert.deepEqual(asked, ["cvj"]);
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
  const failed = search("abc");
  assert.equal(await settledOr(failed), HELD, "a failure answers nothing yet");
  const retried = await search("abc");
  assert.deepEqual(asked, ["abc", "abc"], "the next search really asks again");
  assert.equal(retried.results.length, 1, "and gets a real answer");
});
