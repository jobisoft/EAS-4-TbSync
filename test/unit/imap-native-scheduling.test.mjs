import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

// Execute the actual experiment policy with only the native platform replaced.
const sent = [];
const identity = { email: "me@example.org" };
const smtp = { sendItems(...args) { sent.push(args); return true; } };
class BaseCalendar {
  initProviderBase() {}
  getProperty(name) { return name === "itip.transport" ? smtp : null; }
}
const context = vm.createContext({
  ChromeUtils: {
    importESModule() {
      return {
        ExtensionCommon: {}, ExtensionUtils: {},
        cal: { provider: { BaseClass: BaseCalendar } },
        CalItipEmailTransport: class {
          sendItems(...args) { sent.push(args); return true; }
        },
        MailServices: { accounts: {
          allIdentities: [identity], findAccountForIdentity: () => ({ key: "imap" }),
        } },
      };
    },
    generateQI: () => () => {},
  },
});
const source = readFileSync(new URL("../../src/experiments/calendar/parent/ext-calendar-provider.js", import.meta.url), "utf8");
vm.runInContext(source.slice(0, source.indexOf("this.calendar_provider =")) +
  "\nglobalThis.Calendar = ExtCalendar;", context);
const makeCalendar = (enabled = true) => {
  const calendar = new context.Calendar({});
  calendar.capabilities = { scheduling: "server", organizer: "mailto:me@example.org", imapInvitations: enabled };
  return calendar;
};
const event = ({ marked = false, stamped = false, own = false, attendee = "me@example.org" } = {}) => ({
  organizer: { id: own ? "mailto:me@example.org" : "mailto:sender@example.org" },
  getAttendees: () => [{ id: `mailto:${attendee}` }],
  getProperty: name => name === "X-EAS-CLIENT-INVITATION" && marked ? "1" :
    name === "X-EAS-SERVERID" && stamped ? "sid" : null,
});

test("only an explicit opt-in permits client replies", () => {
  for (const enabled of [false, null, "true", 1]) {
    assert.equal(makeCalendar(enabled).canNotify("REPLY", event()), true);
  }
  const absent = makeCalendar();
  delete absent.capabilities.imapInvitations;
  assert.equal(absent.canNotify("REPLY", event()), true);
  assert.equal(makeCalendar().canNotify("REPLY", event()), false, "initial source before hook marker");
  assert.equal(makeCalendar().canNotify("REPLY", event({ marked: true, stamped: true })), false);
});

test("opt-out affects new invitations while existing marked sources still use SMTP", () => {
  const calendar = makeCalendar(false);
  assert.equal(calendar.canNotify("REPLY", event({ marked: true, stamped: true })), false);
  const before = sent.length;
  assert.equal(calendar.getProperty("itip.transport").sendItems([], {
    responseMethod: "REPLY", getItemList: () => [event({ marked: true, stamped: true })],
  }, { id: "mailto:me@example.org" }), true);
  assert.equal(sent.length, before + 1);
});

test("transport fails closed for an unknown identity or a different sender", () => {
  const calendar = makeCalendar();
  const before = sent.length;
  const reply = { responseMethod: "REPLY", getItemList: () => [event({ marked: true })] };
  const transport = calendar.getProperty("itip.transport");
  assert.equal(transport.sendItems([], reply, { id: "mailto:other@example.org" }), false);
  assert.equal(transport.sendItems([], reply, null), false);
  calendar.capabilities.organizer = "mailto:missing@example.org";
  const unknownReply = { responseMethod: "REPLY", getItemList: () => [event({ marked: true, attendee: "missing@example.org" })] };
  assert.equal(transport.sendItems([], unknownReply, { id: "mailto:missing@example.org" }), false);
  assert.equal(sent.length, before);
});

test("server invitations and our meetings never use SMTP", () => {
  const calendar = makeCalendar();
  assert.equal(calendar.canNotify("REPLY", event({ stamped: true })), true);
  assert.equal(calendar.canNotify("REPLY", event({ own: true, marked: true })), true);
  assert.equal(calendar.canNotify("REPLY", event({ marked: true, attendee: "other@example.org" })), true);
  for (const method of ["REQUEST", "CANCEL", "COUNTER"]) {
    assert.equal(calendar.canNotify(method, event({ marked: true })), true);
  }
});

test("transport delegates only client REPLY and preserves its original item", () => {
  sent.length = 0;
  const transport = makeCalendar().getProperty("itip.transport");
  const sourceItem = event({ marked: true, stamped: true });
  const itip = { responseMethod: "REPLY", getItemList: () => [sourceItem] };
  const recipients = [{ id: "mailto:sender@example.org" }];
  const sender = { id: "mailto:me@example.org" };
  assert.equal(transport.sendItems(recipients, itip, sender), true);
  assert.equal(sent.length, 1);
  assert.equal(sent[0][1], itip, "no UID or SEQUENCE rewrite");
  assert.equal(sent[0][2], sender);
  for (const method of ["REQUEST", "CANCEL"]) {
    transport.sendItems(recipients, { ...itip, responseMethod: method }, sender);
  }
  transport.sendItems(recipients, { ...itip, getItemList: () => [event({ stamped: true })] }, sender);
  assert.equal(sent.length, 1);
});

test("SMTP uses an existing matching identity, never the default other account", () => {
  const calendar = makeCalendar();
  assert.equal(calendar.getProperty("imip.identity"), identity);
  assert.equal(calendar.getProperty("imip.account").key, "imap");
  calendar.capabilities.organizer = "mailto:missing@example.org";
  assert.equal(calendar.getProperty("imip.identity"), null);
  assert.equal(calendar.getProperty("imip.account"), null);
});
