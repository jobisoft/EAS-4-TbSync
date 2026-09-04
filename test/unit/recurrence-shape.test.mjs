/**
 * The recurrence shapes a calendar of ours may hold.
 *
 * Two rules, asserted here: an item carries at most one `RRULE` and never
 * one beside `RDATE`s, and it carries `RDATE`s only if a server stated
 * them that way. Everything else has its dates restated as a rule with
 * modified occurrences, which is the only form the protocol carries.
 *
 * Run with `npm run test:unit`.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { webcrypto } from "node:crypto";

import ICAL from "../../src/vendor/ical.min.js";
import {
  conformRecurrence,
  nonConformingShape,
} from "../../src/modules/recurrence-shape.mjs";

globalThis.crypto ??= webcrypto;

const MASTER = [
  "BEGIN:VCALENDAR",
  "VERSION:2.0",
  "PRODID:-//unit//EN",
  "BEGIN:VEVENT",
  "UID:shape@eas-test.invalid",
  "DTSTAMP:20260801T120000Z",
  "DTSTART:20261106T090000Z",
  "DTEND:20261106T100000Z",
  "SUMMARY:unit master",
];

/** A blob. `overrides` are `[recurrenceId, movedTo]` day pairs. */
function blob({
  rules = [],
  dates = [],
  exdates = [],
  overrides = [],
  kind = "VEVENT",
} = {}) {
  const lines = MASTER.map((l) =>
    l === "BEGIN:VEVENT" ? `BEGIN:${kind}` : l,
  ).map((l) =>
    kind === "VTODO" && l.startsWith("DTEND:") ? l.replace("DTEND", "DUE") : l,
  );
  for (const r of rules) lines.push(`RRULE:${r}`);
  // A day gets the fixture's usual hour; a full instant is taken as given,
  // which is how two occurrences land on one day.
  const at = (d) => (d.includes("T") ? d : `${d}T090000Z`);
  for (const d of dates) lines.push(`RDATE:${at(d)}`);
  for (const d of exdates) lines.push(`EXDATE:${at(d)}`);
  lines.push(`END:${kind}`);
  for (const [rid, to] of overrides) {
    lines.push(
      `BEGIN:${kind}`,
      "UID:shape@eas-test.invalid",
      "DTSTAMP:20260801T120000Z",
      `RECURRENCE-ID:${rid}T090000Z`,
      `DTSTART:${to}T090000Z`,
      `SUMMARY:moved to ${to}`,
      `END:${kind}`,
    );
  }
  return lines.concat("END:VCALENDAR").join("\r\n");
}

let minted = 0;
const conform = (text) =>
  conformRecurrence(text, { newUid: () => `sib-${++minted}@eas-test.invalid` });

const pieces = (r) => [r.master, ...r.siblings.map((s) => s.ical)];

/** One piece, read back: its rule, its dates, and its overrides as
 *  `[recurrenceId, start]`. */
function parts(ical) {
  const vcal = new ICAL.Component(ICAL.parse(ical));
  const comps = vcal
    .getAllSubcomponents()
    .filter((c) => c.name === "vevent" || c.name === "vtodo");
  const master = comps.find((c) => !c.getFirstProperty("recurrence-id"));
  const values = (name) =>
    master
      .getAllProperties(name)
      .flatMap((p) => p.getValues())
      .map((v) => (v.start ?? v).toICALString());
  return {
    uid: master.getFirstPropertyValue("uid"),
    dtstart: master.getFirstPropertyValue("dtstart").toICALString(),
    dtend: master.getFirstPropertyValue("dtend")?.toICALString() ?? null,
    due: master.getFirstPropertyValue("due")?.toICALString() ?? null,
    rrule:
      master.getFirstProperty("rrule")?.getFirstValue()?.toString() ?? null,
    rdates: values("rdate"),
    exdates: values("exdate"),
    overrides: comps
      .filter((c) => c.getFirstProperty("recurrence-id"))
      .map((c) => [
        c.getFirstPropertyValue("recurrence-id").toICALString(),
        c.getFirstPropertyValue("dtstart").toICALString(),
        c.getFirstPropertyValue("uid"),
      ]),
  };
}

/* ── what conforms already ────────────────────────────────────────────── */

test("an item with one rule, or none, is left alone", () => {
  for (const text of [
    blob({ rules: ["FREQ=DAILY;COUNT=3"] }),
    blob({ rules: ["FREQ=DAILY;COUNT=3"], exdates: ["20261107"] }),
    blob({
      rules: ["FREQ=DAILY;COUNT=3"],
      overrides: [["20261107", "20261119"]],
    }),
    blob(),
  ]) {
    assert.equal(conform(text), null);
  }
});

/* ── restating dates as a rule ────────────────────────────────────────── */

test("a set that fits a rule exactly needs no overrides", () => {
  // 6, 20 Nov, 4 Dec - fourteen days apart.
  const r = conform(blob({ dates: ["20261120", "20261204"] }));
  assert.equal(r.siblings.length, 0);
  const p = parts(r.master);
  assert.equal(p.rrule, "FREQ=DAILY;COUNT=3;INTERVAL=14");
  assert.deepEqual(p.rdates, []);
  assert.deepEqual(p.overrides, []);
  assert.equal(p.dtstart, "20261106T090000Z");
});

test("seven days apart is spelled weekly", () => {
  const p = parts(conform(blob({ dates: ["20261113", "20261120"] })).master);
  assert.equal(p.rrule, "FREQ=WEEKLY;COUNT=3");
});

test("an irregular set gets a daily rule and an override per date", () => {
  const r = conform(blob({ dates: ["20261119", "20261225"] }));
  const p = parts(r.master);
  // Packed under the dates, so every override moves its occurrence later -
  // the direction the writer's emission order is built for.
  assert.equal(p.rrule, "FREQ=DAILY;COUNT=3");
  assert.deepEqual(p.rdates, []);
  assert.deepEqual(
    p.overrides.map(([rid, start]) => [rid, start]),
    [
      ["20261107T090000Z", "20261119T090000Z"],
      ["20261108T090000Z", "20261225T090000Z"],
    ],
  );
  // Each override replaces a slot the rule really produces, or the server
  // has nothing to attach it to.
  assert.equal(p.overrides[0][2], p.uid);
});

test("a single date is a plain event, not a series of one", () => {
  // Nothing to fit, and Thunderbird refuses an override on a
  // non-recurring item, so there must be neither rule nor override.
  const r = conform(
    blob({ rules: ["FREQ=DAILY;COUNT=2"], dates: ["20270303"] }),
  );
  const dates = parts(r.siblings[0].ical);
  assert.equal(dates.rrule, null);
  assert.deepEqual(dates.rdates, []);
  assert.deepEqual(dates.overrides, []);
  assert.equal(dates.dtstart, "20270303T090000Z");
});

test("two occurrences on one day keep their dates, for the codec to refuse", () => {
  // No daily rule sits under both and EAS has no finer frequency.
  const text = blob({ dates: ["20261106"] }).replace(
    "RDATE:20261106T090000Z",
    "RDATE:20261106T140000Z",
  );
  const p = parts(conform(text).master);
  assert.equal(p.rrule, null);
  assert.deepEqual(p.rdates, ["20261106T140000Z"]);
});

test("a date that already carries an override keeps it, re-pointed", () => {
  // An RDATE can only say that an occurrence exists. Anything about its
  // content - its own subject, a different time, its attendees - is an
  // override sitting on that date. Giving the dates a rule moves the
  // occurrences to its instants, and the override has to move with the
  // one it describes rather than be duplicated by a minted one.
  const r = conform(
    blob({
      dates: ["20261119", "20261225"],
      overrides: [["20261119", "20261120"]],
    }),
  );
  const p = parts(r.master);
  assert.equal(p.rrule, "FREQ=DAILY;COUNT=3");
  // Three occurrences, so at most two overrides - never three.
  assert.equal(p.overrides.length, 2);
  const [carried] = p.overrides.filter(([, start]) =>
    start.startsWith("20261120"),
  );
  assert.ok(carried, "the override on the listed date was lost");
  // Re-pointed onto the rule's own instant, or the server has nothing to
  // attach it to and the occurrence shows twice.
  assert.equal(carried[0], "20261107T090000Z");
  for (const [rid] of p.overrides) {
    assert.match(
      rid,
      /^202611(06|07|08)T090000Z$/,
      `${rid} is not an instant of the rule`,
    );
  }
});

test("an override on the first date survives the move to DTSTART", () => {
  // Two dates thirteen days apart fit a rule exactly, so nothing is
  // minted - and the override the item already had must still be there,
  // on the instant the rule puts its occurrence at.
  const r = conform(
    blob({
      dates: ["20261119"],
      overrides: [["20261106", "20261107"]],
    }),
  );
  const p = parts(r.master);
  assert.equal(p.rrule, "FREQ=DAILY;COUNT=2;INTERVAL=13");
  assert.deepEqual(
    p.overrides.map(([rid, start]) => [rid, start]),
    [["20261106T090000Z", "20261107T090000Z"]],
  );
});

test("nine listed dates each carrying an override yield nine, not eighteen", () => {
  // The reported shape: a server states a series it cannot express as a
  // rule by sending every occurrence as an exception, so each date it
  // names arrives with an override on it.
  const dates = [
    "20261113",
    "20261120",
    "20261127",
    "20261211",
    "20261218",
    "20270108",
    "20270115",
    "20270122",
  ];
  const r = conform(
    blob({
      dates,
      overrides: [["20261106", "20261106"], ...dates.map((d) => [d, d])],
    }),
  );
  const p = parts(r.master);
  assert.equal(p.rrule, "FREQ=DAILY;COUNT=9");
  assert.equal(p.overrides.length, 9);
  const slots = p.overrides.map(([rid]) => rid).sort();
  assert.equal(new Set(slots).size, 9, "two overrides landed on one instant");
  // Nine consecutive days from the first date, which is what a daily rule
  // of nine produces - and every override sits on one of them.
  const instants = new Set(
    Array.from(
      { length: 9 },
      (_, i) => `202611${String(6 + i).padStart(2, "0")}T090000Z`,
    ),
  );
  for (const rid of slots) {
    assert.ok(instants.has(rid), `${rid} is not an instant of the rule`);
  }
});

/* ── splitting ────────────────────────────────────────────────────────── */

test("two rules become two items, the first keeping the identity", () => {
  const r = conform(
    blob({
      rules: ["FREQ=DAILY;COUNT=3", "FREQ=MONTHLY;BYMONTHDAY=15;COUNT=2"],
    }),
  );
  const [first, second] = pieces(r).map(parts);
  assert.equal(first.uid, "shape@eas-test.invalid");
  assert.equal(first.rrule, "FREQ=DAILY;COUNT=3");
  assert.notEqual(second.uid, first.uid);
  assert.equal(second.rrule, "FREQ=MONTHLY;COUNT=2;BYMONTHDAY=15");
  // Its own rule's first instance: a server derives the series from the
  // start time it is given.
  assert.equal(second.dtstart, "20261115T090000Z");
  assert.equal(second.dtend, "20261115T100000Z");
});

test("a rule beside dates splits, and the dates become a rule of their own", () => {
  const r = conform(
    blob({ rules: ["FREQ=DAILY;COUNT=3"], dates: ["20261120", "20261204"] }),
  );
  const [ruled, listed] = pieces(r).map(parts);
  assert.equal(ruled.rrule, "FREQ=DAILY;COUNT=3");
  assert.deepEqual(ruled.rdates, []);
  assert.equal(listed.dtstart, "20261120T090000Z");
  assert.equal(listed.rrule, "FREQ=DAILY;COUNT=2;INTERVAL=14");
  assert.deepEqual(listed.rdates, []);
});

test("an override goes to the piece that produces its instant", () => {
  const r = conform(
    blob({
      rules: ["FREQ=DAILY;COUNT=3", "FREQ=MONTHLY;BYMONTHDAY=15;COUNT=2"],
      overrides: [
        ["20261107", "20261110"],
        ["20261215", "20261216"],
      ],
    }),
  );
  const [first, second] = pieces(r).map(parts);
  assert.deepEqual(
    first.overrides.map(([rid]) => rid),
    ["20261107T090000Z"],
  );
  assert.deepEqual(
    second.overrides.map(([rid]) => rid),
    ["20261215T090000Z"],
  );
  // Re-UIDed onto its new master, or it belongs to neither.
  assert.equal(second.overrides[0][2], second.uid);
});

test("an exclusion reaches every piece that would produce it", () => {
  const r = conform(
    blob({
      rules: ["FREQ=DAILY;COUNT=3", "FREQ=MONTHLY;BYMONTHDAY=15;COUNT=2"],
      exdates: ["20261107", "20261215"],
    }),
  );
  const [first, second] = pieces(r).map(parts);
  assert.deepEqual(first.exdates, ["20261107T090000Z"]);
  assert.deepEqual(second.exdates, ["20261215T090000Z"]);
});

test("a start no rule produces is an occurrence and keeps a home", () => {
  // DTSTART is 6 Nov and neither rule names it, so it cannot simply be
  // moved onto a rule's first instance - that would lose an occurrence.
  const r = conform(
    blob({
      rules: [
        "FREQ=MONTHLY;BYMONTHDAY=15;COUNT=2",
        "FREQ=MONTHLY;BYMONTHDAY=20;COUNT=2",
      ],
    }),
  );
  const all = pieces(r).map(parts);
  assert.deepEqual(all.map((p) => p.dtstart).sort(), [
    "20261106T090000Z",
    "20261115T090000Z",
    "20261120T090000Z",
  ]);
});

/* ── tasks ────────────────────────────────────────────────────────────── */

test("a task's far end moves with its start", () => {
  const r = conform(
    blob({
      kind: "VTODO",
      rules: ["FREQ=DAILY;COUNT=3", "FREQ=MONTHLY;BYMONTHDAY=15;COUNT=2"],
    }),
  );
  const second = parts(r.siblings[0].ical);
  // Or the piece is due an hour before it starts.
  assert.equal(second.dtstart, "20261115T090000Z");
  assert.equal(second.due, "20261115T100000Z");
});

/* ── what the guard deliberately leaves behind ────────────────────────── */

test("dates the guard keeps are refused by the codec, not sent short", async () => {
  // Two occurrences on one calendar day are the one set it cannot give a
  // rule, so it reports the shape and leaves them as dates. Nothing
  // downstream can carry that - the server stores such an item without
  // its dates - so the push has to refuse it rather than lose the
  // occurrences in silence.
  const { clientRejectReason } = await import(
    "../../src/modules/eas/calendar-codec.mjs"
  );
  const text = blob({ dates: ["20261119T090000Z", "20261119T140000Z"] });
  const stored = conform(text)?.master ?? text;
  assert.match(
    stored,
    /^RDATE/m,
    "the guard found a rule for a same-day set after all",
  );
  assert.match(
    clientRejectReason({ blob: stored, syncRecurrence: true }) ?? "",
    /same day/,
  );
});
