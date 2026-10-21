import { test, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { installWebextEnv } from "./support/webext-env.mjs";
import { EasProvider } from "../../src/modules/eas-provider.mjs";
installWebextEnv();
browser.runtime.getManifest = () => ({});
messenger.addressBooks = { onUpdated: { addListener() {} } };
let stored, calendars, custom, operations, enabled;
browser.storage = { local: {
  async get(defaults) { return { ...structuredClone(defaults), ...structuredClone(stored) }; },
  async set(values) { operations.push("storage"); Object.assign(stored, structuredClone(values)); },
  async remove(key) { operations.push(`remove:${key}`); delete stored[key]; },
} };
messenger.calendar.calendars = {
  async get(id) { return calendars[id]; },
  async update(id, { capabilities }) {
    operations.push(`capability:${id}:${capabilities.imapInvitations}`);
    Object.assign(calendars[id].capabilities, capabilities);
  },
};
const folders = [
  { targetID: "default", targetType: "calendars", custom: { type: "8" } },
  { targetID: "shared", targetType: "calendars", custom: { type: "13" } },
  { targetID: "tasks", targetType: "tasks", custom: { type: "9" } },
];
const provider = new EasProvider();
provider.getAccount = async () => ({ account: { accountId: "account", custom, enabled }, folders });
provider.updateAccount = async ({ patch }) => Object.assign(custom, patch.custom);
beforeEach(() => {
  stored = { "eas.clientScheduling": { existing: { user: "untouched" } } };
  custom = { userSmtpAddress: "me@example.org" };
  enabled = false;
  calendars = Object.fromEntries(["default", "shared", "tasks"].map(id => [id, { capabilities: { scheduling: "server", imapInvitations: false } }]));
  operations = [];
});

test("IMAP invitations default off and accept only an explicit boolean true", async () => {
  assert.equal((await provider.getAccountForConfig("account")).imapInvitations, false);
  await provider.saveAccountFromConfig({ accountId: "account", patch: { imapInvitations: "true" } });
  assert.equal(custom.imapInvitations, false);
  assert.equal(stored["eas.imapInvitations.default"], undefined);
});

test("saving opt-in sets only the default calendar and banks policy before SMTP capability", async () => {
  await provider.saveAccountFromConfig({ accountId: "account", patch: { imapInvitations: true } });
  assert.equal(custom.imapInvitations, true);
  assert.deepEqual(stored["eas.imapInvitations.default"], { user: "me@example.org" });
  assert.equal(stored["eas.imapInvitations.shared"], undefined);
  assert.equal(calendars.default.capabilities.imapInvitations, true);
  assert.equal(calendars.shared.capabilities.imapInvitations, false);
  assert.equal(calendars.tasks.capabilities.imapInvitations, false);
  assert.ok(operations.indexOf("storage") < operations.indexOf("capability:default:true"));
  assert.deepEqual(stored["eas.clientScheduling"], { existing: { user: "untouched" } });
});

test("opt-out stops SMTP before dropping hook policy and preserves source snapshots", async () => {
  await provider.saveAccountFromConfig({ accountId: "account", patch: { imapInvitations: true } });
  stored["eas.clientInvitations.default"] = { source: { blob: "original" } };
  operations.length = 0;
  await provider.saveAccountFromConfig({ accountId: "account", patch: { imapInvitations: false } });
  assert.deepEqual(operations, ["capability:default:false", "remove:eas.imapInvitations.default", "remove:eas.imapInvitations.shared"]);
  assert.equal(stored["eas.imapInvitations.default"], undefined);
  assert.equal(stored["eas.clientInvitations.default"].source.blob, "original");
});

test("the opt-in checkbox and both user descriptions are present", () => {
  const html = readFileSync(new URL("../../src/dialogs/config/config.html", import.meta.url), "utf8");
  assert.match(html, /id="imap-invitations" type="checkbox"/);
  for (const locale of ["pl", "en-US"]) {
    const messages = JSON.parse(readFileSync(new URL(`../../src/_locales/${locale}/messages.json`, import.meta.url), "utf8"));
    assert.ok(messages["config.calendar.imapInvitations"].message);
    assert.ok(messages["config.calendar.imapInvitations.hint"].message);
  }
});

test("a connected account can save only IMAP policy without disconnecting or resetting cache", async () => {
  enabled = true;
  await provider.saveAccountFromConfig({ accountId: "account", patch: { imapInvitations: true } });
  assert.equal(custom.imapInvitations, true);
  assert.equal(calendars.default.capabilities.imapInvitations, true);
  await assert.rejects(provider.saveAccountFromConfig({ accountId: "account", patch: {
    imapInvitations: false, asVersionSelected: "14.1",
  } }), /tylko obsługę zaproszeń/);
  assert.equal(custom.imapInvitations, true);
  assert.equal(custom.asversionselected, undefined);
});

test("an older sync snapshot cannot revert the option saved on the connected account", async () => {
  const oldSnapshot = { accountId: "account", custom: { ...custom, imapInvitations: false } };
  custom.imapInvitations = true;
  await provider.updateImapInvitationMode(oldSnapshot, folders);
  assert.deepEqual(stored["eas.imapInvitations.default"], { user: "me@example.org" });
  assert.equal(calendars.default.capabilities.imapInvitations, true);
});

test("concurrent account policy updates cannot overwrite another calendar's policy", async () => {
  const other = new EasProvider();
  other.getAccount = async () => ({ account: {
    accountId: "other", custom: { imapInvitations: true, userSmtpAddress: "other@example.org" },
  } });
  custom.imapInvitations = true;
  calendars.other = { capabilities: { scheduling: "server" } };
  await Promise.all([
    provider.updateImapInvitationMode({ accountId: "account" }, folders),
    other.updateImapInvitationMode({ accountId: "other" }, [{
      targetID: "other", targetType: "calendars", custom: { type: "8" },
    }]),
  ]);
  assert.deepEqual(stored["eas.imapInvitations.default"], { user: "me@example.org" });
  assert.deepEqual(stored["eas.imapInvitations.other"], { user: "other@example.org" });
});
