/**
 * `syncrecurrence` moves one way only: off to on, never back.
 *
 * Off means recurrence rules and their exceptions are left off the wire in
 * both directions, so a series reads as a single entry on whichever side did
 * not author it. It is a state a carried-over profile can bring in, and the
 * two rules below are what keep an account from being put back into it.
 *
 * The migration is the one thing that switches it on by itself, and it may:
 * the account's resources are deleted and rebuilt from the server around it,
 * so the pull that follows carries the recurrence in.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { installWebextEnv } from "./support/webext-env.mjs";
import { EasProvider } from "../../src/modules/eas-provider.mjs";

installWebextEnv();
// The handler logs through i18n, and the env ships no `browser.i18n`.
globalThis.browser.i18n = { getMessage: (key) => key };

const onMigrate = EasProvider.prototype.onMigrateLegacyAccount;

/** Just enough provider to run the handler: what it reads, what it writes. */
function providerHolding(custom) {
  const writes = [];
  const logs = [];
  return {
    writes,
    logs,
    async getAccount(accountId) {
      return { account: { accountId, custom }, folders: [] };
    },
    async updateAccount(args) {
      writes.push(args);
    },
    reportEventLog(payload) {
      logs.push(payload);
    },
  };
}

test("a migrated account that has recurrence off is switched on", async () => {
  const p = providerHolding({ syncrecurrence: false });
  await onMigrate.call(p, { accountId: "1" });
  assert.equal(p.writes.length, 1);
  assert.deepEqual(p.writes[0], {
    accountId: "1",
    patch: { custom: { syncrecurrence: true } },
  });
});

test("an account that never carried the setting is switched on too", async () => {
  const p = providerHolding({});
  await onMigrate.call(p, { accountId: "1" });
  assert.equal(p.writes.length, 1);
  assert.equal(p.writes[0].patch.custom.syncrecurrence, true);
});

test("an account that already has it on is not written to at all", async () => {
  const p = providerHolding({ syncrecurrence: true });
  await onMigrate.call(p, { accountId: "1" });
  assert.deepEqual(p.writes, []);
  assert.deepEqual(p.logs, []);
});

test("switching it on is recorded, so the account's history shows it", async () => {
  const p = providerHolding({ syncrecurrence: false });
  await onMigrate.call(p, { accountId: "1" });
  assert.equal(p.logs.length, 1);
  assert.equal(p.logs[0].level, "info");
  assert.equal(p.logs[0].accountId, "1");
});

/* ── The config popup's half of the rule ─────────────────────────────────
 *
 * The control is offered while the setting is off, so it can be switched on,
 * and withheld once it is on. A withheld row must not write its key back, or
 * the ratchet would be the checkbox's value rather than a rule - and the
 * dialog would send `false` for a row nobody could see.
 */

const CONFIG_MJS = readFileSync(
  new URL("../../src/dialogs/config/config.mjs", import.meta.url),
  "utf-8",
);
const CONFIG_HTML = readFileSync(
  new URL("../../src/dialogs/config/config.html", import.meta.url),
  "utf-8",
);

test("the row the rule hides exists, and is what the rule names", () => {
  assert.match(CONFIG_HTML, /id="sync-recurrence-row"/);
  assert.match(CONFIG_MJS, /\$\("sync-recurrence-row"\)\.hidden = /);
});

test("the save patch never carries syncRecurrence unconditionally", () => {
  const literal = CONFIG_MJS.slice(
    CONFIG_MJS.indexOf("const patch = {"),
    CONFIG_MJS.indexOf("};", CONFIG_MJS.indexOf("const patch = {")),
  );
  assert.ok(literal.length > 0, "could not find the patch literal");
  assert.ok(
    !literal.includes("syncRecurrence"),
    "syncRecurrence is in the unconditional patch literal, so a hidden row " +
      "would still write it back",
  );
  assert.match(
    CONFIG_MJS,
    /if \(!\$\("sync-recurrence-row"\)\.hidden\) \{\s*patch\.syncRecurrence = /,
  );
});
