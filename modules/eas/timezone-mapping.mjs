/**
 * Windows ↔ IANA timezone mapping for EAS calendar items.
 *
 * EAS ≤14.x carries timezone information as a 172-byte blob whose
 * standardName/daylightName fields are Windows zone IDs (e.g. "W. Europe
 * Standard Time"). Thunderbird's calendar uses IANA tzids (e.g.
 * "Europe/Berlin"). Inbound and outbound conversions both need a map.
 *
 * Data sources (verbatim copies from the legacy EAS-4-TbSync,
 * originally from mj1856's TimeZoneConverter):
 *   - timezonedata/WindowsTimezone.csv (Windows → IANA, 514 lines)
 *   - timezonedata/Aliases.csv         (IANA alias chains, 115 lines)
 *
 * Inbound resolution mirrors `eas.tools.guessTimezoneByStdDstOffset` from
 * the legacy provider: four tiers - Windows-name, IANA-string parse,
 * abbreviation match, pure-offset fallback - with a "default-overtake"
 * branch when one Windows zone maps to multiple IANA zones (e.g. W.
 * Europe Standard Time → Berlin/Rome/Brussels).
 */

import ICAL from "../../vendor/ical.min.js";

const DAYS = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];

/** Frozen "no zone" info, used for UTC and as a no-DST collapse target. */
const UTC_INFO = Object.freeze({
  id: "UTC",
  offset: 0,
  abbreviation: "UTC",
  displayname: "UTC",
});
function utcTzInfo(tzid = "UTC") {
  return { std: { ...UTC_INFO, id: tzid, displayname: tzid }, dst: { ...UTC_INFO, id: tzid, displayname: tzid }, timezone: tzid };
}

let _loaded = null; // Promise<void> - resolves once state is built
let _state = null;  // { windowsToIana, ianaToWindows, cachedTimezoneData, defaultTimezone, defaultTimezoneInfo }

/** Per-tzid VTIMEZONE→info cache (filled lazily by loadTzInfo). */
const _tzInfoCache = new Map();

/** Idempotent on success: loads the two CSVs, walks every IANA zone the
 *  calendar service knows about, parses each VTIMEZONE definition, and
 *  builds the resolver's lookup tables. Concurrent calls share one
 *  in-flight load. On failure the cache is cleared and the rejection
 *  propagates, so the next call retries from scratch. */
export function ensureLoaded() {
  if (_state) return Promise.resolve();
  if (!_loaded) {
    _loaded = loadInternal();
    // Log + clear the cached promise on failure so the next caller retries.
    // The original rejection still propagates to whoever awaits this call.
    _loaded.catch((err) => {
      console.error("[eas] timezone-mapping load failed:", err);
      _loaded = null;
    });
  }
  return _loaded;
}

async function loadInternal() {
  const [winCsvLines, aliasCsvLines] = await Promise.all([
    fetchCsvLines("modules/eas/timezonedata/WindowsTimezone.csv"),
    fetchCsvLines("modules/eas/timezonedata/Aliases.csv"),
  ]);
  const currentZone =
    (await messenger.calendar.timezones.currentZone) || "UTC";
  const timezoneIds = (await messenger.calendar.timezones.timezoneIds) || [];
  
  // 1) Aliases.csv → "Africa/Abidjan" → ["Iceland", "Africa/Timbuktu", …]
  const aliasNames = {};
  for (const line of aliasCsvLines) {
    const cols = line.split(",");
    if (cols.length < 2) continue;
    aliasNames[cols[0].trim()] = cols[1].trim().split(/\s+/);
  }

  // 2) WindowsTimezone.csv → forward + reverse maps
  const windowsToIana = Object.create(null);
  const ianaToWindows = Object.create(null);
  for (const line of winCsvLines) {
    const cols = line.split(",");
    if (cols.length < 3) continue;
    const winName = cols[0].trim();
    const zoneType = cols[1].trim();
    const ianaList = cols[2].trim();
    if (zoneType === "001") windowsToIana[winName] = ianaList;
    // Reverse map: every IANA in the list maps back to this Windows zone,
    // and so do all of its IANA aliases. Same Windows zone wins regardless
    // of CSV row order because every row for one Windows zone carries the
    // same name in column 0.
    for (const ianaZone of ianaList.split(/\s+/)) {
      if (!ianaZone) continue;
      ianaToWindows[ianaZone] = winName;
      const aliases = aliasNames[ianaZone];
      if (aliases) {
        for (const alias of aliases) {
          if (alias) ianaToWindows[alias] = winName;
        }
      }
    }
  }

  // 3) Walk every IANA zone TB knows; extract std/dst offset, abbreviation
  //    and switch-date rule. Build the four lookup tables the resolver uses.
  const cached = {
    iana: Object.create(null),
    abbreviations: Object.create(null),
    bothOffsets: Object.create(null),
    stdOffset: Object.create(null),
  };

  for (const tzid of timezoneIds) {
    const info = await loadTzInfo(tzid);
    if (!info) continue;
    const both = `${info.std.offset}:${info.dst.offset}`;
    const stdKey = info.std.offset;
    // Only overwrite if the slot is empty OR this is the user's default
    // zone - so the default zone wins the offset-fallback ties.
    if (!cached.bothOffsets[both] || tzid === currentZone) {
      cached.bothOffsets[both] = tzid;
    }
    if (!cached.stdOffset[stdKey] || tzid === currentZone) {
      cached.stdOffset[stdKey] = tzid;
    }
    if (info.std.abbreviation) {
      cached.abbreviations[info.std.abbreviation] = tzid;
    }
    cached.iana[tzid] = info;
  }

  // 4) Pin UTC and the user's default zone in the offset/abbreviation
  //    tables so they win ties even when iteration order didn't favour them.
  cached.bothOffsets["0:0"] = "UTC";
  if (!cached.iana["UTC"]) cached.iana["UTC"] = utcTzInfo("UTC");
  const defaultInfo = cached.iana[currentZone] || utcTzInfo(currentZone);
  cached.iana[currentZone] = defaultInfo;
  if (defaultInfo.std.abbreviation) {
    cached.abbreviations[defaultInfo.std.abbreviation] = currentZone;
  }
  cached.bothOffsets[`${defaultInfo.std.offset}:${defaultInfo.dst.offset}`] =
    currentZone;
  cached.stdOffset[defaultInfo.std.offset] = currentZone;

  // Stamp the default zone's Windows name so the Tier-1 overtake branch
  // can compare against it (matches legacy provider.js:104).
  if (ianaToWindows[currentZone]) {
    defaultInfo.std.windowsZoneName = ianaToWindows[currentZone];
  }

  _state = {
    windowsToIana,
    ianaToWindows,
    cachedTimezoneData: cached,
    defaultTimezone: currentZone,
    defaultTimezoneInfo: defaultInfo,
  };
}

async function fetchCsvLines(relativePath) {
  const url = browser.runtime.getURL(relativePath);
  const text = await (await fetch(url)).text();
  return text.split(/\r?\n/).filter((l) => l && !l.startsWith("#"));
}

/** Parse one IANA zone's VTIMEZONE into {std, dst, timezone}. Returns
 *  null for zones whose definition can't be fetched/parsed. Caches per
 *  tzid; subsequent calls are O(1). */
async function loadTzInfo(tzid) {
  if (_tzInfoCache.has(tzid)) return _tzInfoCache.get(tzid);
  if (tzid === "UTC" || tzid === "floating") {
    const info = utcTzInfo(tzid);
    _tzInfoCache.set(tzid, info);
    return info;
  }
  const def = await messenger.calendar.timezones.getDefinition(tzid);
  if (!def) {
    _tzInfoCache.set(tzid, null);
    return null;
  }
  // The API returns just BEGIN:VTIMEZONE…; ICAL.parse needs a wrapping
  // VCALENDAR. Be defensive in case the wrapper is ever added upstream.
  const wrapped = def.includes("BEGIN:VCALENDAR")
    ? def
    : `BEGIN:VCALENDAR\r\n${def}\r\nEND:VCALENDAR`;
  const comp = new ICAL.Component(ICAL.parse(wrapped));
  const vtimezone =
    comp.name === "vtimezone" ? comp : comp.getFirstSubcomponent("vtimezone");
  if (!vtimezone) {
    _tzInfoCache.set(tzid, null);
    return null;
  }
  const std = parseTzSubcomponent(vtimezone, "standard", tzid);
  let dst = parseTzSubcomponent(vtimezone, "daylight", tzid);
  if (!std) {
    _tzInfoCache.set(tzid, null);
    return null;
  }
  // No-DST zones: collapse dst → std (matches legacy tools.js:454). Both
  // sides will then carry the same offset and a missing switchdate, so
  // the outbound blob writer leaves SYSTEMTIME zero-filled.
  if (!dst) dst = std;
  // Build an `ICAL.Timezone` from the parsed VTIMEZONE so callers can
  // ask `tz.utcOffset(time)` for a moment-in-time offset and convert
  // a UTC time into wall-clock parts via `time.convertToZone(tz)`.
  // This is the ICAL.js-native equivalent of legacy's
  // `calITimezoneService.getTimezone(tzid)` + `.getInTimezone(...)`.
  const icalTimezone = new ICAL.Timezone({ component: vtimezone, tzid });
  const info = { std, dst, timezone: tzid, icalTimezone };
  _tzInfoCache.set(tzid, info);
  return info;
}

/** Pick the latest STANDARD/DAYLIGHT subcomponent (zones change over time;
 *  the most-recent rule is what's currently in effect) and pull offset,
 *  abbreviation, and the SYSTEMTIME-shaped switch-date out of its RRULE. */
function parseTzSubcomponent(vtimezone, kind, tzid) {
  const subs = vtimezone.getAllSubcomponents(kind);
  if (!subs.length) return null;
  subs.sort((a, b) => {
    const da = String(a.getFirstPropertyValue("dtstart") ?? "");
    const db = String(b.getFirstPropertyValue("dtstart") ?? "");
    return db.localeCompare(da);
  });
  const sub = subs[0];

  // tzoffsetto: ICAL exposes "+HHMM" / "-HHMM" (no colon). Legacy parsed
  // the integer form: "+0530" → 530, "-0330" → -330, then split into
  // hours/minutes and negated to get "minutes from local to UTC".
  const tzoffsetto = String(sub.getFirstPropertyValue("tzoffsetto") ?? "");
  if (!tzoffsetto) return null;
  const o = parseInt(tzoffsetto.replace(":", ""), 10);
  if (Number.isNaN(o)) return null;
  const h = Math.trunc(o / 100);
  const m = o - h * 100;
  const offset = -1 * (h * 60 + m);

  const tzname = sub.getFirstPropertyValue("tzname");
  const abbreviation = tzname == null ? "" : String(tzname);

  const obj = {
    id: tzid,
    offset,
    abbreviation,
    displayname: tzid,
  };

  const rrule = sub.getFirstPropertyValue("rrule");
  const dtstart = sub.getFirstPropertyValue("dtstart");
  if (rrule && dtstart) {
    const rules = parseRRule(rrule);
    if (
      rules.FREQ === "YEARLY" &&
      rules.BYDAY &&
      rules.BYMONTH &&
      rules.BYDAY.length > 2
    ) {
      const month = parseInt(rules.BYMONTH, 10);
      const dayCode = rules.BYDAY.slice(-2);
      let weekOfMonth = parseInt(rules.BYDAY.slice(0, -2), 10);
      // SYSTEMTIME's wDay range is 1..5 where 5 means "last". Legacy
      // clamps anything out of range (and negative "last-N" rules) to 5.
      if (Number.isNaN(weekOfMonth) || weekOfMonth < 0 || weekOfMonth > 5) {
        weekOfMonth = 5;
      }
      let dayOfWeek = DAYS.indexOf(dayCode);
      if (dayOfWeek < 0) dayOfWeek = 0;
      const time = parseDtstartTime(dtstart);
      obj.switchdate = {
        month,
        dayOfWeek,
        weekOfMonth,
        hour: time.hour,
        minute: time.minute,
        second: time.second,
      };
    }
  }

  return obj;
}

function parseRRule(rrule) {
  const out = {};
  const text = typeof rrule === "string" ? rrule : (rrule.toString?.() ?? "");
  for (const part of text.split(";")) {
    const eq = part.indexOf("=");
    if (eq < 0) continue;
    out[part.slice(0, eq)] = part.slice(eq + 1);
  }
  return out;
}

function parseDtstartTime(dtstart) {
  // ICAL.Property.getFirstPropertyValue("dtstart") returns an ICAL.Time
  // for date-time properties; older code paths can return a raw string.
  if (dtstart && typeof dtstart === "object") {
    return {
      hour: dtstart.hour ?? 0,
      minute: dtstart.minute ?? 0,
      second: dtstart.second ?? 0,
    };
  }
  const s = String(dtstart ?? "");
  const m = s.match(/T(\d{2}):?(\d{2}):?(\d{2})/);
  if (!m) return { hour: 0, minute: 0, second: 0 };
  return {
    hour: parseInt(m[1], 10),
    minute: parseInt(m[2], 10),
    second: parseInt(m[3], 10),
  };
}

/** Inbound: given the std/dst offsets and the (possibly Windows) zone name
 *  carried in the EAS blob, return the IANA tzid that best matches.
 *  Mirrors legacy `eas.tools.guessTimezoneByStdDstOffset`.
 *
 *  Synchronous; the caller must `await ensureLoaded()` earlier in the
 *  sync entry point. Throws if called before initialisation. */
export function guessTimezoneByStdDstOffset(stdOffset, dstOffset, stdName) {
  if (!_state) {
    throw new Error("timezone-mapping: ensureLoaded() must be awaited first");
  }
  const {
    windowsToIana,
    cachedTimezoneData: cached,
    defaultTimezone,
    defaultTimezoneInfo,
  } = _state;

  // Tier 1: Windows zone-name lookup, with default-overtake when one
  // Windows zone maps to multiple IANA zones and the user is in one
  // of them (e.g. Berlin vs. Rome both belong to W. Europe Standard Time).
  const winIana = windowsToIana[stdName];
  if (winIana && cached.iana[winIana]?.std.offset === stdOffset) {
    if (
      defaultTimezoneInfo.std.windowsZoneName &&
      winIana !== defaultTimezone &&
      cached.iana[winIana].std.offset === defaultTimezoneInfo.std.offset &&
      stdName === defaultTimezoneInfo.std.windowsZoneName
    ) {
      return defaultTimezone;
    }
    return winIana;
  }

  // Tiers 2 & 3: split the std name on punctuation, test each chunk as a
  // literal IANA tzid then as an international abbreviation (CET, CAT, …).
  const parts = String(stdName ?? "")
    .replace(/[;,()\[\]]/g, " ")
    .split(/\s+/)
    .filter(Boolean);
  for (const part of parts) {
    if (cached.iana[part]?.std.offset === stdOffset) return part;
    const fromAbbr = cached.abbreviations[part];
    if (fromAbbr && cached.iana[fromAbbr]?.std.offset === stdOffset) {
      return fromAbbr;
    }
  }

  // Tier 4: pure offset fallback (both → std-only → default).
  const both = cached.bothOffsets[`${stdOffset}:${dstOffset}`];
  if (both) return both;
  const stdOnly = cached.stdOffset[stdOffset];
  if (stdOnly) return stdOnly;
  return defaultTimezone;
}

/** Outbound: given an IANA tzid (or "floating" / unknown), return the
 *  {std, dst} info plus the resolved Windows zone names that the blob
 *  writer needs. Falls back to the user's default zone for floating /
 *  unresolvable tzids (matches legacy calendarsync.js:299-303).
 *
 *  Synchronous; the caller must `await ensureLoaded()` earlier in the
 *  sync entry point. Throws if called before initialisation. */
export function tzInfoForBlob(tzid) {
  if (!_state) {
    throw new Error("timezone-mapping: ensureLoaded() must be awaited first");
  }
  const { ianaToWindows, defaultTimezoneInfo, cachedTimezoneData } = _state;

  let info = null;
  if (tzid && tzid !== "floating") {
    info = cachedTimezoneData.iana[tzid] ?? null;
  }
  if (!info) info = defaultTimezoneInfo;

  const stdWinName =
    ianaToWindows[info.std.displayname] ?? info.std.displayname;
  const dstWinName =
    ianaToWindows[info.dst.displayname] ?? info.dst.displayname;

  return { std: info.std, dst: info.dst, stdWinName, dstWinName };
}

/** Inbound (Tasks): given the moment-in-time offset between an EAS
 *  `<UtcStartDate>` / `<StartDate>` pair (or the equivalent due-date
 *  pair), return the IANA tzid that produces that offset at that
 *  exact moment. Mirrors legacy `eas.tools.guessTimezoneByCurrentOffset`
 *  ([tools.js]) which iterated `calITimezoneService.timezoneIds` and
 *  returned the first match.
 *
 *    offsetMinutes  signed minutes east of UTC at the given moment
 *    utcDate        JS `Date` for the moment (UTC)
 *
 *  Returns `_state.defaultTimezone` when no zone matches (legacy
 *  returned `null` and let upstream fall back; we centralise here so
 *  callers always get a valid tzid).
 *
 *  Synchronous; the caller must `await ensureLoaded()` earlier in the
 *  sync entry point. Throws if called before initialisation. */
export function guessTimezoneByCurrentOffset(offsetMinutes, utcDate) {
  if (!_state) {
    throw new Error("timezone-mapping: ensureLoaded() must be awaited first");
  }
  const time = ICAL.Time.fromJSDate(utcDate, false);
  const wantSec = offsetMinutes * 60;
  const ianaMap = _state.cachedTimezoneData.iana;
  // Sorted iteration matches legacy's deterministic "first match wins"
  // semantics (Lightning's calITimezoneService also returned tzids in
  // a stable, alphabetical order).
  for (const tzid of Object.keys(ianaMap).sort()) {
    const info = ianaMap[tzid];
    if (!info?.icalTimezone) continue;
    if (info.icalTimezone.utcOffset(time) === wantSec) return tzid;
  }
  return _state.defaultTimezone;
}

/** Fetch the cached `ICAL.Timezone` for an IANA tzid, or null if the
 *  zone isn't in the loaded set (or the input is `null` / `"UTC"` /
 *  `"floating"`). Callers building iCal properties with a `TZID=`
 *  parameter use this to convert UTC moments to wall-clock in the
 *  named zone via `time.convertToZone(tz)`.
 *
 *  Synchronous; the caller must `await ensureLoaded()` earlier. */
export function getIcalTimezone(tzid) {
  if (!tzid || tzid === "UTC" || tzid === "floating") return null;
  if (!_state) {
    throw new Error("timezone-mapping: ensureLoaded() must be awaited first");
  }
  return _state.cachedTimezoneData.iana[tzid]?.icalTimezone ?? null;
}
