/**
 * An item the server states with exceptions and no recurrence rule.
 *
 * Kerio Connect sends it for meetings imported from an Outlook
 * invitation. The exception keys land on the instants of a rule it holds
 * but does not state, so the occurrences between them are the rule's to
 * produce and are simply not in the message. Thunderbird refuses such an
 * item whole, which failed the whole folder on every sync (#355).
 *
 * It is skipped on sight rather than reconstructed, and these pin the
 * sight: what the reader refuses, what it must not refuse, and that the
 * pull consults it before it builds anything.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import { serverRejectReason } from "../../src/modules/eas/calendar-codec.mjs";
import { el } from "./support/ad-node.mjs";

const exception = (key, extra = []) =>
  el("Exception", [el("ExceptionStartTime", key), ...extra]);

const rule = () =>
  el("Recurrence", [
    el("Type", "1"),
    el("Interval", "2"),
    el("DayOfWeek", "16"),
  ]);

function adNode({ exceptions = null, recurrence = null } = {}) {
  const kids = [
    el("UID", "listed@eas-test.invalid"),
    el("Subject", "listed occurrences"),
    el("StartTime", "20261012T090000Z"),
    el("EndTime", "20261012T100000Z"),
  ];
  if (recurrence) kids.push(recurrence);
  if (exceptions) kids.push(el("Exceptions", exceptions));
  return el("ApplicationData", kids);
}

test("exceptions with no rule are refused", () => {
  const reason = serverRejectReason({
    adNode: adNode({
      exceptions: [
        exception("20261012T090000Z"),
        exception("20261026T090000Z"),
      ],
    }),
  });
  assert.match(String(reason), /no recurrence rule/);
});

test("a single cancelled occurrence with no rule is refused too", () => {
  // The other shape in the reporter's mailbox: one exception, and it is a
  // deletion keyed to the master's own start. Reading past the deletion
  // leaves an EXDATE excluding an instant nothing generates.
  const reason = serverRejectReason({
    adNode: adNode({
      exceptions: [exception("20261012T090000Z", [el("Deleted", "1")])],
    }),
  });
  assert.match(String(reason), /no recurrence rule/);
});

test("exceptions to a stated rule are taken", () => {
  assert.equal(
    serverRejectReason({
      adNode: adNode({
        recurrence: rule(),
        exceptions: [exception("20261012T090000Z")],
      }),
    }),
    null,
  );
});

test("an item with no exceptions is taken, ruled or not", () => {
  assert.equal(serverRejectReason({ adNode: adNode() }), null);
  assert.equal(
    serverRejectReason({ adNode: adNode({ recurrence: rule() }) }),
    null,
  );
});

test("nothing to read is not a refusal", () => {
  assert.equal(serverRejectReason({ adNode: null }), null);
});

test("the pull refuses before it builds anything", async () => {
  // The point of the refusal is that such an item never enters the
  // calendar, so the check has to come before the decode - reading it out
  // of the runner rather than restating it here.
  const { readFileSync } = await import("node:fs");
  const src = readFileSync(
    new URL("../../src/modules/eas/sync-runner.mjs", import.meta.url),
    "utf8",
  );
  const add = src.slice(src.indexOf("async function applyAdd("));
  const checked = add.indexOf("serverRejectReason");
  const built = add.indexOf("applicationDataToBlob");
  assert.ok(
    checked > 0,
    "the pull no longer asks whether the item is takeable",
  );
  assert.ok(
    checked < built,
    "the item is built before it is refused, so it entered the calendar",
  );
});
