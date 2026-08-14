/**
 * Which accounts we send the attendees' mail for ourselves.
 *
 * This gate is the whole of the version policy, and both answers cost
 * something real. Say yes where the server already sends, and every
 * attendee of every meeting gets invited twice. Say no where it does not,
 * and organising a meeting notifies nobody - the bug this feature exists
 * to close.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { versionNeedsClientScheduling } from "../../src/modules/eas/client-scheduling.mjs";

test("14.0 and 14.1 are ours to send", () => {
  // Measured: a 14.1 server does not invite the attendees on push, and
  // with `scheduling: "server"` Thunderbird does not either.
  assert.equal(versionNeedsClientScheduling("14.0"), true);
  assert.equal(versionNeedsClientScheduling("14.1"), true);
});

test("16.x sends its own, so we must stay silent", () => {
  // Measured on cvjmbonn: the invitation arrives on the second pull with
  // nothing sent by us. Sending as well would invite everybody twice.
  assert.equal(versionNeedsClientScheduling("16.0"), false);
  assert.equal(versionNeedsClientScheduling("16.1"), false);
});

test("2.5 is excluded, and not because it is old", () => {
  // It is below 16 but cannot do this at all: `sendMail` builds WBXML
  // where 2.5 wants raw MIME, and the codec emits no attendee block there.
  // A note recorded on a 2.5 account could never be sent, so the gate is a
  // range rather than "below 16". Pinned 2.5 accounts exist - item 31.
  assert.equal(versionNeedsClientScheduling("2.5"), false);
});

test("an unsupported or unreadable version sends nothing", () => {
  // Absent means no, everywhere: a missing answer costs one notification,
  // while guessing yes mails every attendee of every meeting.
  for (const v of [null, undefined, "", "auto", "garbage", NaN, {}]) {
    assert.equal(versionNeedsClientScheduling(v), false, `for ${String(v)}`);
  }
});

test("the boundaries are exactly where the versions are", () => {
  assert.equal(versionNeedsClientScheduling("13.9"), false);
  assert.equal(versionNeedsClientScheduling("14"), true);
  assert.equal(versionNeedsClientScheduling("15.9"), true);
  assert.equal(versionNeedsClientScheduling("16"), false);
});
