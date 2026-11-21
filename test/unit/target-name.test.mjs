/**
 * Unit tests for following a folder that the server renamed.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  DEFAULT_FOLLOW_MODE,
  followModeOf,
  genResourceName,
  planNameFollow,
} from "../../src/modules/eas/target-name.mjs";

const ACCOUNT = "Work";

const row = (over = {}) => ({
  folderId: "f-1",
  displayName: "Calendar",
  targetName: "Calendar (Work)",
  targetID: "cal-1",
  targetType: "calendars",
  ...over,
});
const renamed = (over = {}) => ({ folderId: "f-1", displayName: "Events", ...over });
const plan = (oldRows, newDescriptors, mode) =>
  planNameFollow({ oldRows, newDescriptors, accountName: ACCOUNT, mode });

test("a resource is named after its folder and its account", () => {
  assert.equal(genResourceName("Calendar", "Work"), "Calendar (Work)");
});

test("an unmodified name follows the server", () => {
  assert.deepEqual(plan([row()], [renamed()], "unmodified"), [
    {
      folderId: "f-1",
      targetID: "cal-1",
      targetType: "calendars",
      name: "Events (Work)",
    },
  ]);
});

test("a name the user chose is left alone unless the mode is always", () => {
  const mine = row({ targetName: "My dates" });
  assert.deepEqual(plan([mine], [renamed()], "unmodified"), []);
  assert.deepEqual(plan([mine], [renamed()], "always").map((p) => p.name), [
    "Events (Work)",
  ]);
});

test("ignore never renames anything", () => {
  assert.deepEqual(plan([row()], [renamed()], "ignore"), []);
  assert.deepEqual(plan([row()], [renamed()], "nonsense"), []);
});

test("a name the server did not change is not touched", () => {
  assert.deepEqual(
    plan([row()], [renamed({ displayName: "Calendar" })], "always"),
    [],
  );
});

test("a folder without a stored name or without an old row has nothing to rename", () => {
  assert.deepEqual(
    plan([row({ targetName: null })], [renamed()], "always"),
    [],
  );
  assert.deepEqual(
    plan([row({ targetName: "" })], [renamed()], "always"),
    [],
  );
  assert.deepEqual(plan([], [renamed()], "always"), []);
  assert.deepEqual(plan(undefined, undefined, "always"), []);
});

test("a folder without a resource still gets the stored name", () => {
  const [item] = plan([row({ targetID: null })], [renamed()], "unmodified");
  assert.equal(item.targetID, null);
  assert.equal(item.name, "Events (Work)");
});

test("a name that is already the new one is not renamed again", () => {
  const already = row({ targetName: "Events (Work)" });
  assert.deepEqual(plan([already], [renamed()], "always"), []);
});

test("moving a folder into the trash changes its name, and it follows", () => {
  const old = row({ displayName: "Old", targetName: "Old (Work)" });
  const next = renamed({ displayName: "Trash | Old" });
  assert.deepEqual(plan([old], [next], "unmodified").map((p) => p.name), [
    "Trash | Old (Work)",
  ]);
});

test("several folders are judged one by one", () => {
  const a = row();
  const b = row({
    folderId: "f-2",
    displayName: "Tasks",
    targetName: "Mine",
    targetType: "tasks",
    targetID: "t-2",
  });
  const c = row({ folderId: "f-3", displayName: "Contacts", targetName: "Contacts (Work)" });
  const out = plan(
    [a, b, c],
    [
      renamed(),
      { folderId: "f-2", displayName: "To do" },
      { folderId: "f-3", displayName: "Contacts" },
    ],
    "unmodified",
  );
  assert.deepEqual(out.map((p) => p.folderId), ["f-1"]);
});

test("the mode of an account is read from its custom data", () => {
  assert.equal(followModeOf({ followServerRenames: "always" }), "always");
  assert.equal(followModeOf({ followServerRenames: "ignore" }), "ignore");
  assert.equal(followModeOf({}), DEFAULT_FOLLOW_MODE);
  assert.equal(followModeOf(undefined), DEFAULT_FOLLOW_MODE);
  assert.equal(followModeOf({ followServerRenames: "bogus" }), DEFAULT_FOLLOW_MODE);
  assert.equal(DEFAULT_FOLLOW_MODE, "unmodified");
});
