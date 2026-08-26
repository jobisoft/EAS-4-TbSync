import { strict as assert } from "node:assert";
import { test } from "node:test";

import { overwriteDetails } from "../../src/modules/calendar-provider.mjs";

const STORED = "BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:a\r\nX-EAS-SERVERID:11:14\r\nEND:VEVENT\r\nEND:VCALENDAR";
const INCOMING = "BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:a\r\nX-MOZ-RECEIVED-DTSTAMP:20260826T110000Z\r\nEND:VEVENT\r\nEND:VCALENDAR";

test("both versions appear whole, each under its own label", () => {
  const details = overwriteDetails(INCOMING, STORED);
  assert.ok(details.includes(`--- stored ---\n${STORED}`));
  assert.ok(details.includes(`--- incoming ---\n${INCOMING}`));
});

test("the stored version comes first, so a reader sees the change in order", () => {
  const details = overwriteDetails(INCOMING, STORED);
  assert.ok(details.indexOf("--- stored ---") < details.indexOf("--- incoming ---"));
});

test("nothing is dropped from either version", () => {
  const details = overwriteDetails(INCOMING, STORED);
  assert.ok(details.includes("X-EAS-SERVERID:11:14"));
  assert.ok(details.includes("X-MOZ-RECEIVED-DTSTAMP:20260826T110000Z"));
});

test("without a stored version the label still names what is there", () => {
  for (const nothing of [null, undefined, ""]) {
    const details = overwriteDetails(INCOMING, nothing);
    assert.equal(details, `--- incoming ---\n${INCOMING}`);
    assert.ok(!details.includes("--- stored ---"));
  }
});
