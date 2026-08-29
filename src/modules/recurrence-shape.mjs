/**
 * The recurrence shapes a calendar of ours may hold.
 *
 * Thunderbird stores shapes ActiveSync cannot carry - two `RRULE`s, an
 * `RRULE` beside `RDATE`s, a bare list of `RDATE`s - and measured, none
 * survives a push: 14.1 keeps only the last of two rules, 16.1 refuses
 * them, and a list of occurrences is stored without its dates. The same
 * occurrences are carriable as an ordinary series with modified
 * occurrences, which every server keeps. So two rules hold here:
 *
 *   1. an item carries at most one RRULE, and never an RRULE with RDATEs
 *   2. an item carries no RDATEs at all - they become an RRULE with
 *      overrides, one moving each occurrence onto its date
 *
 * One question decides every placement: *does this piece produce instant
 * T?* A date list answers by lookup, a rule by stepping its own iterator
 * until it reaches or passes T.
 */

import ICAL from "../vendor/ical.min.js";

/** How far a rule is stepped before the answer is taken to be "no".
 *  Iteration stops of its own accord at the instant being asked about, so
 *  this only bounds a rule that never reaches it. High enough that no
 *  ordinary rule meets it - a daily rule reaches this far in 273 years -
 *  because a rule cut short answers "does not generate", which would put
 *  an override on the wrong piece. */
const STEP_CAP = 100000;

const RECURRING_KINDS = new Set(["vevent", "vtodo"]);

/* ── Asking a rule or a list about one instant ─────────────────────── */

const stamp = (time) => (time ? time.toUnixTime() : null);
const sameInstant = (a, b) => stamp(a) === stamp(b);

/** An `RDATE` may hold a period rather than an instant. Its start is the
 *  occurrence either way. */
const instantOf = (value) => value?.start ?? value ?? null;

function ruleProduces(recur, dtstart, instant) {
  const want = stamp(instant);
  if (want === null) return false;
  const iterator = recur.iterator(dtstart);
  for (let step = 0; step < STEP_CAP; step++) {
    const next = iterator.next();
    if (!next) return false;
    const at = stamp(next);
    if (at === want) return true;
    if (at > want) return false;
  }
  return false;
}

/** The first `count` instants a rule produces, or fewer if it ends.
 *
 *  Cloned as they come: the iterator hands back the same `ICAL.Time` it
 *  keeps stepping, so collecting them without copying yields a list of
 *  references to one instant that ends up holding the last. */
function instantsOf(recur, dtstart, count) {
  const iterator = recur.iterator(dtstart);
  const out = [];
  while (out.length < count) {
    const next = iterator.next();
    if (!next) break;
    out.push(next.clone());
  }
  return out;
}

/** Every value of a multi-valued date property, as instants. One property
 *  can carry a comma-separated list, so the values are read rather than
 *  the properties counted. */
function datesOf(vcomp, name) {
  const out = [];
  for (const prop of vcomp.getAllProperties(name)) {
    for (const value of prop.getValues()) {
      const at = instantOf(value);
      if (at) out.push(at);
    }
  }
  return out;
}

const holds = (list, instant) => list.some((e) => sameInstant(e, instant));

/** Does this piece produce that instant? Its `DTSTART` counts: a start is
 *  an occurrence in its own right, whether or not the rule names it. */
function pieceProduces(piece, instant) {
  if (sameInstant(piece.dtstart, instant)) return true;
  if (holds(piece.dates, instant)) return true;
  return !!piece.recur && ruleProduces(piece.recur, piece.dtstart, instant);
}

/* ── Reading the item ──────────────────────────────────────────────── */

const isRecurring = (comp) => RECURRING_KINDS.has(comp.name);

function masterOf(vcal) {
  return (
    vcal
      .getAllSubcomponents()
      .find((c) => isRecurring(c) && !c.getFirstProperty("recurrence-id")) ??
    null
  );
}

const overridesOf = (vcal) =>
  vcal
    .getAllSubcomponents()
    .filter((c) => isRecurring(c) && c.getFirstProperty("recurrence-id"));

/** Why this item may not be held as it is, or null when it may be.
 *  Structural: it counts rules and dates and reads neither. */
export function nonConformingShape(vcomp) {
  if (!vcomp) return null;
  const rules = vcomp.getAllProperties("rrule").length;
  const dates = datesOf(vcomp, "rdate").length;
  if (rules > 1 && dates > 0) return "more than one RRULE, and RDATEs";
  if (rules > 1) return "more than one RRULE";
  if (rules === 1 && dates > 0) return "an RRULE together with RDATEs";
  if (dates > 0) return "occurrences stated as a list of RDATEs";
  return null;
}

/* ── Building components ───────────────────────────────────────────── */

/** A copy that shares nothing with the original.
 *
 *  `toJSON()` hands back the component's own jCal array rather than a
 *  copy, so wrapping it builds a second view of the same data and every
 *  edit made through one is visible through the other. Pieces are built by
 *  editing a copy of the master, so without this they edit each other. */
const cloneComponent = (comp) =>
  new ICAL.Component(JSON.parse(JSON.stringify(comp.toJSON())));

const cloneProperty = (prop) =>
  new ICAL.Property(JSON.parse(JSON.stringify(prop.toJSON())));

/** The calendar wrapper: its own properties, and anything that is not an
 *  event or a task - which is where `VTIMEZONE` lives. */
function wrapperLike(vcal) {
  const out = new ICAL.Component("vcalendar");
  for (const prop of vcal.getAllProperties())
    out.addProperty(cloneProperty(prop));
  for (const sub of vcal.getAllSubcomponents()) {
    if (!isRecurring(sub)) out.addSubcomponent(cloneComponent(sub));
  }
  return out;
}

/** Move a component's start, carrying whatever states its far end with it
 *  so the length is kept. An event states `DTEND` and a task `DUE`; a
 *  `DURATION` is relative already and needs no help. */
function moveStart(comp, from, to) {
  if (sameInstant(from, to)) return;
  const shift = to.subtractDate(from);
  for (const name of ["dtend", "due"]) {
    const end = comp.getFirstPropertyValue(name);
    if (!end) continue;
    const moved = end.clone();
    moved.addDuration(shift);
    comp.updatePropertyWithValue(name, moved);
  }
  comp.updatePropertyWithValue("dtstart", to);
}

/** The master as one occurrence of itself, at `when`, replacing the
 *  occurrence the rule puts at `slot`. */
function overrideFor(master, masterStart, uid, slot, when) {
  const comp = cloneComponent(master);
  comp.updatePropertyWithValue("uid", uid);
  comp.removeAllProperties("rrule");
  comp.removeAllProperties("rdate");
  comp.removeAllProperties("exdate");
  moveStart(comp, masterStart, when);
  comp.updatePropertyWithValue("recurrence-id", slot);
  return comp;
}

/* ── Turning a list of dates into a rule ───────────────────────────── */

const DAY_MS = 86400000;

/** Whole days between two instants, or null when they are not a whole
 *  number of days apart. */
function wholeDaysBetween(a, b) {
  const ms = b.toJSDate() - a.toJSDate();
  return ms > 0 && ms % DAY_MS === 0 ? ms / DAY_MS : null;
}

/** Rules worth trying for this set, cheapest first. Each is only a
 *  candidate: `fitRule` keeps one when its own instants *are* the dates,
 *  so a proposal that does not fit costs nothing but the check. */
function candidateRules(dates) {
  const n = dates.length;
  const out = [];
  const gap = wholeDaysBetween(dates[0], dates[1]);
  if (gap === 7) out.push(`FREQ=WEEKLY;COUNT=${n}`);
  if (gap) out.push(`FREQ=DAILY;INTERVAL=${gap};COUNT=${n}`);
  const day = dates[0].day;
  if (dates.every((d) => d.day === day)) {
    out.push(`FREQ=MONTHLY;BYMONTHDAY=${day};COUNT=${n}`);
  }
  const month = dates[0].month;
  if (dates.every((d) => d.day === day && d.month === month)) {
    out.push(`FREQ=YEARLY;COUNT=${n}`);
  }
  return out;
}

/** A rule whose own instants are exactly these dates, or null.
 *
 *  Proposed and then checked by expansion rather than reasoned about: a
 *  monthly rule fits three dates a month apart and not three that skip a
 *  month, and the difference is easier to measure than to argue. */
function fitRule(dates) {
  for (const text of candidateRules(dates)) {
    const recur = ICAL.Recur.fromString(text);
    const got = instantsOf(recur, dates[0], dates.length);
    if (got.length !== dates.length) continue;
    if (got.every((t, i) => sameInstant(t, dates[i]))) return recur;
  }
  return null;
}

/** Two occurrences on one calendar day cannot both sit under a daily rule,
 *  and EAS has no finer frequency - so there is no rule to pack under such
 *  a set and no safe order to write its overrides in. Left as dates, which
 *  is what the codec refuses. */
const sharesADay = (dates) =>
  new Set(dates.map((d) => `${d.year}-${d.month}-${d.day}`)).size !==
  dates.length;

/**
 * Give a piece whose occurrences are dates a rule that produces them.
 *
 * A fitted rule costs no overrides. Failing that it is daily from the
 * first date, which sits at or before every date - sorted distinct days
 * are at least a day apart - so every override moves its occurrence
 * later, the direction the writer's emission order is built for.
 *
 * A date may already carry an override, and usually does when the dates
 * came from a server: an `RDATE` can only say that an occurrence exists,
 * so anything about its content - a subject of its own, a different time,
 * its attendees - is an override sitting on it. Those are re-pointed onto
 * the instant the rule now puts that occurrence at, and are the reason
 * this cannot simply mint one per date: minting would leave the item with
 * two overrides for the same occurrence, one of them keyed to an instant
 * the new rule never produces.
 *
 * `carried` maps a date to the override already on it, and every one
 * taken is added to `claimed` so the caller does not place it a second
 * time. Returns false when the set can be given no rule at all.
 */
function ruleFor(piece, master, masterStart, carried, claimed) {
  const dates = [piece.dtstart, ...piece.dates].sort(
    (a, b) => stamp(a) - stamp(b),
  );
  piece.dates = [];
  piece.dtstart = dates[0];

  // One occurrence is not a series. Left as a plain single event, which is
  // also the only shape Thunderbird accepts for it: it refuses an override
  // on a non-recurring item outright.
  if (dates.length === 1) return true;

  if (sharesADay(dates)) {
    piece.dates = dates.slice(1);
    return false;
  }

  const fitted = fitRule(dates);
  piece.recur =
    fitted ?? ICAL.Recur.fromString(`FREQ=DAILY;COUNT=${dates.length}`);

  const slots = fitted
    ? dates
    : instantsOf(piece.recur, piece.dtstart, dates.length);
  for (let i = 0; i < dates.length; i++) {
    const own = carried.get(stamp(dates[i]));
    if (own) {
      // It already says what this occurrence is. All it needs is the
      // instant the rule now puts that occurrence at.
      const moved = cloneComponent(own);
      moved.updatePropertyWithValue("uid", piece.uid);
      moved.updatePropertyWithValue("recurrence-id", slots[i]);
      piece.overrides.push(moved);
      claimed.add(own);
      continue;
    }
    // Nothing about this occurrence differs except where the rule puts
    // it, so an override of the master at that date is the whole of it.
    if (i > 0 && !sameInstant(slots[i], dates[i])) {
      piece.overrides.push(
        overrideFor(master, masterStart, piece.uid, slots[i], dates[i]),
      );
    }
  }
  return true;
}

/* ── The entry point ───────────────────────────────────────────────── */

/** A piece left with one occurrence and an override on it is not a series
 *  with an exception - it is a single event at the override's time.
 *  Thunderbird refuses the other shape outright ("Exceptions were supplied
 *  to a non-recurring item") and EAS cannot state it either. */
function collapsed(piece) {
  if (piece.recur || piece.dates.length || piece.overrides.length !== 1) {
    return null;
  }
  const only = piece.overrides[0];
  const rid = only.getFirstPropertyValue("recurrence-id");
  if (!sameInstant(rid, piece.dtstart)) return null;
  const comp = cloneComponent(only);
  comp.removeAllProperties("recurrence-id");
  return comp;
}

function render(piece, vcal, master, masterStart) {
  const out = wrapperLike(vcal);
  const single = collapsed(piece);
  if (single) {
    out.addSubcomponent(single);
    return out.toString();
  }
  const comp = cloneComponent(master);
  comp.updatePropertyWithValue("uid", piece.uid);
  comp.removeAllProperties("rrule");
  comp.removeAllProperties("rdate");
  comp.removeAllProperties("exdate");
  moveStart(comp, masterStart, piece.dtstart);
  if (piece.recur) comp.addPropertyWithValue("rrule", piece.recur);
  for (const d of piece.dates) comp.addPropertyWithValue("rdate", d);
  for (const d of piece.exdates) comp.addPropertyWithValue("exdate", d);
  out.addSubcomponent(comp);
  for (const override of piece.overrides) out.addSubcomponent(override);
  return out.toString();
}

/**
 * Make an item conform, or say it already does.
 *
 * Returns null when there is nothing to do, which is the common case.
 * Otherwise `{ shape, master, siblings }`: the
 * master keeps the original `UID` and is stored under the id the caller
 * already has, and each sibling is a new item to write beside it.
 *
 * `newUid` mints a sibling's id, random by default: nothing links the
 * pieces, so a derived id would buy no mechanism and could collide if the
 * same item were split twice.
 */
export function conformRecurrence(
  ical,
  { newUid = () => crypto.randomUUID() } = {},
) {
  let vcal;
  try {
    vcal = new ICAL.Component(ICAL.parse(ical));
  } catch {
    return null;
  }
  const master = masterOf(vcal);
  const shape = master ? nonConformingShape(master) : null;
  if (!shape) return null;

  const masterStart = master.getFirstPropertyValue("dtstart");
  if (!masterStart) return null;

  const uid = master.getFirstPropertyValue("uid");
  const rules = master.getAllProperties("rrule").map((p) => p.getFirstValue());
  const exdates = datesOf(master, "exdate");
  const excluded = (t) => holds(exdates, t);

  // One piece per rule, each starting where its own rule first fires. The
  // master's start cannot serve a rule that does not name it: a server
  // derives the whole series from the start time it is given.
  const pieces = [];
  rules.forEach((recur, index) => {
    const start = instantsOf(recur, masterStart, 1)[0];
    if (!start) return;
    pieces.push({
      uid: index === 0 ? uid : newUid(),
      recur,
      dates: [],
      dtstart: start,
      exdates: [],
      overrides: [],
    });
  });

  // Occurrences no rule accounts for. The master's own start is one of
  // them when no rule names it - a start is an occurrence in its own
  // right, so dropping it would lose one.
  const ownedByARule = (t) =>
    rules.some((r) => ruleProduces(r, masterStart, t));
  const loose = datesOf(master, "rdate").filter(
    (d) => !ownedByARule(d) && !excluded(d),
  );
  if (!ownedByARule(masterStart) && !excluded(masterStart)) {
    loose.push(masterStart);
  }
  loose.sort((a, b) => stamp(a) - stamp(b));
  const claimed = new Set();
  if (loose.length) {
    // The overrides already sitting on those dates. Giving the dates a
    // rule moves the occurrences to its instants, and each of these has
    // to move with the occurrence it describes.
    const carried = new Map();
    for (const override of overridesOf(vcal)) {
      const rid = override.getFirstPropertyValue("recurrence-id");
      if (holds(loose, rid)) carried.set(stamp(rid), override);
    }
    pieces.push({
      uid: pieces.length ? newUid() : uid,
      recur: null,
      dates: loose.slice(1),
      dtstart: loose[0],
      exdates: [],
      overrides: [],
    });
    ruleFor(pieces[pieces.length - 1], master, masterStart, carried, claimed);
  }
  if (!pieces.length) return null;

  // An exclusion has to reach every piece that would otherwise produce it.
  for (const date of exdates) {
    for (const piece of pieces) {
      if (pieceProduces(piece, date)) piece.exdates.push(date);
    }
  }

  // An override belongs to the piece that produces its instant. Any other
  // piece producing it excludes it, so the occurrence appears once.
  for (const override of overridesOf(vcal)) {
    if (claimed.has(override)) continue;
    const rid = override.getFirstPropertyValue("recurrence-id");
    const owner = pieces.find((p) => pieceProduces(p, rid)) ?? pieces[0];
    const comp = cloneComponent(override);
    comp.updatePropertyWithValue("uid", owner.uid);
    owner.overrides.push(comp);
    for (const piece of pieces) {
      if (piece !== owner && pieceProduces(piece, rid)) piece.exdates.push(rid);
    }
  }

  const blobs = pieces.map((piece) => ({
    uid: piece.uid,
    ical: render(piece, vcal, master, masterStart),
  }));
  return { shape, master: blobs[0].ical, siblings: blobs.slice(1) };
}
