/**
 * Item 33 / #189 — attendee availability.
 *
 * The fixtures that matter are captures, not inventions: "002200" and
 * the 48-character day string came off a live probe on 10 Aug 2026,
 * asking about a mailbox with one known event at 10:00-11:00 UTC on
 * 2026-08-12, against Z-Push 14.1 and Exchange Online 16.1. The shorter
 * strings are constructed, to reach shapes no probe produced - an
 * aligned reply wider than the caller's window, an undocumented digit,
 * a truncated answer. What the digits mean, and where the two families
 * diverge, is in the memory note `eas-freebusy-mergedfreebusy`.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  accountCanAnswerFor,
  alignWindow,
  intervalsFromMergedFreeBusy,
  SLOT_MS,
} from "../../src/modules/eas/free-busy.mjs";

const D = (s) => new Date(s);
const DAY = "2026-08-12";

/** The window as the request would carry it, plus the mapping, so a test
 *  reads the way the runtime does. */
const expand = (mergedFreeBusy, wantStart, wantEnd, types) => {
  const asked = alignWindow(wantStart, wantEnd);
  return intervalsFromMergedFreeBusy({
    mergedFreeBusy,
    askedStart: asked.start,
    wantStart,
    wantEnd,
    types,
  });
};

test("slots are 30 minutes and start at the asked-for time", () => {
  assert.equal(SLOT_MS, 30 * 60 * 1000);
  const w = alignWindow(D(`${DAY}T09:00:00Z`), D(`${DAY}T12:00:00Z`));
  assert.equal(w.start.toISOString(), `${DAY}T09:00:00.000Z`);
  assert.equal(w.end.toISOString(), `${DAY}T12:00:00.000Z`);
});

test("an unaligned window is widened to whole slots", () => {
  // Z-Push answers an unaligned window by starting its grid there, which
  // smears a one-hour event across three slots; Exchange refuses a
  // partial slot outright. Asking whole slots is what makes them agree.
  const w = alignWindow(D(`${DAY}T09:07:00Z`), D(`${DAY}T12:07:00Z`));
  assert.equal(w.start.toISOString(), `${DAY}T09:00:00.000Z`);
  assert.equal(w.end.toISOString(), `${DAY}T12:30:00.000Z`);
  assert.equal((w.end - w.start) % SLOT_MS, 0);
});

test("a sub-slot window still asks for one whole slot", () => {
  // The 20-minute ask that Exchange rejected with Status 5.
  const w = alignWindow(D(`${DAY}T10:00:00Z`), D(`${DAY}T10:20:00Z`));
  assert.equal(w.start.toISOString(), `${DAY}T10:00:00.000Z`);
  assert.equal(w.end.toISOString(), `${DAY}T10:30:00.000Z`);
});

test("a zero-length window is still answerable", () => {
  const at = D(`${DAY}T10:00:00Z`);
  const w = alignWindow(at, at);
  assert.equal(w.end - w.start, SLOT_MS);
});

test("the live 09:00-12:00 reply resolves to the event's real hour", () => {
  // Captured: "002200" for a 3-hour window with one event 10:00-11:00.
  const out = expand("002200", D(`${DAY}T09:00:00Z`), D(`${DAY}T12:00:00Z`));
  assert.deepEqual(out, [
    {
      start: `${DAY}T09:00:00.000Z`,
      end: `${DAY}T10:00:00.000Z`,
      type: "free",
    },
    {
      start: `${DAY}T10:00:00.000Z`,
      end: `${DAY}T11:00:00.000Z`,
      type: "busy",
    },
    {
      start: `${DAY}T11:00:00.000Z`,
      end: `${DAY}T12:00:00.000Z`,
      type: "free",
    },
  ]);
});

test("neighbouring slots of one type merge into a single interval", () => {
  const out = expand("000000", D(`${DAY}T09:00:00Z`), D(`${DAY}T12:00:00Z`));
  assert.equal(out.length, 1, "six free slots are one free interval");
  assert.equal(out[0].start, `${DAY}T09:00:00.000Z`);
  assert.equal(out[0].end, `${DAY}T12:00:00.000Z`);
});

test("the live day-long reply puts the event at slots 20-21", () => {
  // Captured: 48 characters, busy at indices 20 and 21 = 10:00-11:00.
  const digits = "0".repeat(20) + "22" + "0".repeat(26);
  assert.equal(digits.length, 48);
  const out = expand(digits, D(`${DAY}T00:00:00Z`), D("2026-08-13T00:00:00Z"));
  const busy = out.filter((i) => i.type === "busy");
  assert.deepEqual(busy, [
    {
      start: `${DAY}T10:00:00.000Z`,
      end: `${DAY}T11:00:00.000Z`,
      type: "busy",
    },
  ]);
});

test("intervals are clipped back to the range Thunderbird asked about", () => {
  // Asking 09:07-12:07 widens to 09:00-12:30, so the reply covers more
  // than the caller wants; the grid must not claim knowledge outside it.
  const out = expand("0022000", D(`${DAY}T09:07:00Z`), D(`${DAY}T12:07:00Z`));
  assert.equal(out[0].start, `${DAY}T09:07:00.000Z`, "clipped at the front");
  assert.equal(out[out.length - 1].end, `${DAY}T12:07:00.000Z`, "and the back");
  const busy = out.filter((i) => i.type === "busy");
  assert.deepEqual(busy, [
    {
      start: `${DAY}T10:00:00.000Z`,
      end: `${DAY}T11:00:00.000Z`,
      type: "busy",
    },
  ]);
});

test("every documented digit maps, and anything else reads as unknown", () => {
  const out = expand("01234", D(`${DAY}T09:00:00Z`), D(`${DAY}T11:30:00Z`));
  assert.deepEqual(
    out.map((i) => i.type),
    ["free", "tentative", "busy", "unavailable", "unknown"],
  );
  const odd = expand("9", D(`${DAY}T09:00:00Z`), D(`${DAY}T09:30:00Z`));
  assert.equal(odd[0].type, "unknown", "an undocumented digit is not guessed");
});

test("only the types the caller asked for come back", () => {
  const out = expand("002200", D(`${DAY}T09:00:00Z`), D(`${DAY}T12:00:00Z`), [
    "busy",
  ]);
  assert.deepEqual(
    out.map((i) => i.type),
    ["busy"],
  );
});

test("a reply is read for its own length, not stretched to the window", () => {
  // The server is the authority on how much it answered. A short string
  // covers the slots it has and no more.
  const out = expand("22", D(`${DAY}T09:00:00Z`), D(`${DAY}T12:00:00Z`));
  assert.deepEqual(out, [
    {
      start: `${DAY}T09:00:00.000Z`,
      end: `${DAY}T10:00:00.000Z`,
      type: "busy",
    },
  ]);
  assert.deepEqual(
    expand("", D(`${DAY}T09:00:00Z`), D(`${DAY}T12:00:00Z`)),
    [],
  );
  assert.deepEqual(
    expand(null, D(`${DAY}T09:00:00Z`), D(`${DAY}T12:00:00Z`)),
    [],
    "a missing MergedFreeBusy is silence, not an error",
  );
});

/* ── routing: which accounts may answer for an address ──────────────── */

const acct = (address) => ({ custom: { userSmtpAddress: address } });

test("an account answers for its own address and its domain, nothing else", () => {
  const a = acct("john@example.org");
  assert.equal(accountCanAnswerFor(a, "john@example.org"), true, "itself");
  assert.equal(accountCanAnswerFor(a, "JOHN@Example.ORG"), true, "case");
  assert.equal(accountCanAnswerFor(a, " john@example.org "), true, "spacing");
  assert.equal(
    accountCanAnswerFor(a, "kollege@example.org"),
    true,
    "colleague",
  );
  assert.equal(
    accountCanAnswerFor(a, "someone@other.org"),
    false,
    "a stranger's address is never handed to this server",
  );
  assert.equal(accountCanAnswerFor(a, "not-an-address"), false);
  assert.equal(accountCanAnswerFor(a, ""), false);
  assert.equal(
    accountCanAnswerFor(a, "example.org"),
    false,
    "a bare domain is not a person - it must not reach the server",
  );
  assert.equal(accountCanAnswerFor(a, "a@b@example.org"), true, "last @ wins");
  assert.equal(
    accountCanAnswerFor(a, "someone@mail.example.org"),
    false,
    "a subdomain is a different mail domain",
  );
});

test("an account that has not learned its own address answers for nobody", () => {
  // Guessing from the login would leak addresses to a server with no
  // relationship to them - and the login may not even be an address.
  assert.equal(
    accountCanAnswerFor({ custom: { user: "DOMAIN\\jbieling" } }, "x@y.org"),
    false,
  );
  assert.equal(accountCanAnswerFor({}, "x@y.org"), false);
});

test("two accounts on one domain both match, so both get asked", () => {
  const address = "kollege@example.org";
  const accounts = [
    acct("john@example.org"),
    acct("john@other.org"),
    acct("second@example.org"),
  ];
  const matching = accounts.filter((a) => accountCanAnswerFor(a, address));
  assert.equal(matching.length, 2, "answers merge rather than first-wins");
});

/* ── the reply reader ───────────────────────────────────────────────── */

test("the reader tolerates a reply that names no availability", async () => {
  // A recipient the server cannot resolve comes back without the
  // Availability block at all - absence, not an error (the #337 lesson).
  const { installWebextEnv } = await import("./support/webext-env.mjs");
  installWebextEnv();
  const { readResolveRecipients } =
    await import("../../src/modules/eas/resolve-recipients.mjs");
  const { parseAdNode } = await import("./support/ad-node.mjs");
  const doc = (xml) => ({ documentElement: parseAdNode(xml) });

  const full = readResolveRecipients(
    doc(`<ResolveRecipients><Status>1</Status><Response>
           <To>a@b.org</To><Status>1</Status>
           <Recipient><DisplayName>A B</DisplayName>
             <Availability><Status>1</Status>
               <MergedFreeBusy>002200</MergedFreeBusy></Availability>
           </Recipient></Response></ResolveRecipients>`),
  );
  assert.equal(full.mergedFreeBusy, "002200");
  assert.equal(full.availabilityStatus, "1");

  const noAvailability = readResolveRecipients(
    doc(`<ResolveRecipients><Status>1</Status><Response>
           <To>a@b.org</To><Status>1</Status>
           <Recipient><DisplayName>A B</DisplayName></Recipient>
         </Response></ResolveRecipients>`),
  );
  assert.equal(noAvailability.status, "1");
  assert.equal(noAvailability.mergedFreeBusy, null);
  assert.equal(noAvailability.availabilityStatus, null);

  const refused = readResolveRecipients(
    doc(`<ResolveRecipients><Status>5</Status></ResolveRecipients>`),
  );
  assert.equal(refused.status, "5");
  assert.equal(refused.mergedFreeBusy, null);

  assert.equal(readResolveRecipients(null), null, "no document at all");
});
