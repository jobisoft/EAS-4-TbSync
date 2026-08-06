"""Item bodies the tests push, and the cleanup that finds them again.

Every probe carries a stable UID and a SUMMARY beginning with `PROBE`. Both
matter after a crash: the UID makes a re-run overwrite its own litter instead
of adding to it, and the marker lets `sweep` recognise anything left behind
without touching the account's real data.

Larger fixtures - the recurrence and timezone series - are files in
`fixtures/`, because their value is in being byte-stable across runs.
"""

import os

FIXTURES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")

MARKER = "PROBE"

BERLIN = [
    "BEGIN:VTIMEZONE",
    "TZID:Europe/Berlin",
    "BEGIN:DAYLIGHT",
    "TZOFFSETFROM:+0100",
    "TZOFFSETTO:+0200",
    "TZNAME:CEST",
    "DTSTART:19700329T020000",
    "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
    "END:DAYLIGHT",
    "BEGIN:STANDARD",
    "TZOFFSETFROM:+0200",
    "TZOFFSETTO:+0100",
    "TZNAME:CET",
    "DTSTART:19701025T030000",
    "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
    "END:STANDARD",
    "END:VTIMEZONE",
]


def fixture(name):
    """Read one of the .ics files in fixtures/, whole, as `items.create`
    wants it."""
    with open(os.path.join(FIXTURES, name), encoding="utf-8") as f:
        return f.read()


def event(slug, summary=None, lines=(), timezone=False):
    body = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//eas-test//EN"]
    if timezone:
        body += BERLIN
    body += [
        "BEGIN:VEVENT",
        f"UID:{slug}@eas-test.invalid",
        "DTSTAMP:20260801T120000Z",
        f"SUMMARY:{MARKER} {summary or slug}",
    ]
    body += list(lines) + ["END:VEVENT", "END:VCALENDAR"]
    return "\r\n".join(body) + "\r\n"


def task(slug, rrule=None, summary=None, lines=()):
    """A task anchored on 1 Sep 2026 - a Tuesday, and the 1st of month 9.

    Deliberate: every field EAS fills in from DTSTART then has a value you can
    check by eye (DayOfWeek 4, DayOfMonth 1, MonthOfYear 9).
    """
    body = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//eas-test//EN",
        "BEGIN:VTODO",
        f"UID:{slug}@eas-test.invalid",
        "DTSTAMP:20260801T120000Z",
        f"SUMMARY:{MARKER} {summary or slug}",
        "DTSTART:20260901T080000Z",
        "DUE:20260901T090000Z",
    ]
    if rrule:
        body.append(f"RRULE:{rrule}")
    body += list(lines) + ["END:VTODO", "END:VCALENDAR"]
    return "\r\n".join(body) + "\r\n"


def card(slug, extra=()):
    """A contact with enough shape to notice a codec regression: structured
    name, three emails (the EAS ceiling), typed phones, two addresses, a
    birthday, nickname, categories and a note with non-ASCII in it.

    The email is the anchor - `slug@probe.invalid` survives a round trip
    unchanged, while FN can be rebuilt from name parts and the UID is minted
    fresh on a clean pull.
    """
    lines = [
        "BEGIN:VCARD",
        "VERSION:4.0",
        f"FN:{MARKER} {slug}",
        f"N:{slug};{MARKER};Q;Dr.;jun.",
        "NICKNAME:Probey",
        "ORG:Beispiel GmbH;Entwicklung",
        "TITLE:Protokolltester",
        f"EMAIL;TYPE=work:{slug}@probe.invalid",
        f"EMAIL;TYPE=home:{slug}-home@probe.invalid",
        f"EMAIL:{slug}-third@probe.invalid",
        "TEL;TYPE=work:+49 228 1234567",
        "TEL;TYPE=cell:+49 170 7654321",
        "TEL;TYPE=home:+49 228 7654321",
        "ADR;TYPE=work:;;Musterweg 4;Bonn;NRW;53111;Deutschland",
        "ADR;TYPE=home:;;Heimstr. 2;Bonn;NRW;53111;Deutschland",
        "BDAY:19800229",
        "URL:https://example.invalid/probe",
        "CATEGORIES:Probes,Tests",
        "NOTE:Angelegt vom Test. Umlaute: äöü ÄÖÜ ß.",
    ]
    lines += list(extra) + ["END:VCARD", ""]
    return "\r\n".join(lines)


def email_of(slug):
    """The anchor a test matches on - stable across a round trip."""
    return f"{slug}@probe.invalid"


def reset(s, kinds=None, report=True):
    """Evaluate the current state and clear anything a test left behind.

    Every section calls this first. Two reasons it is not merely tidiness:

      - A crashed or throttled run leaves fixtures on the server. The next
        run then imports a second copy, and assertions about "the series"
        start matching whichever one comes first - which is how a duplicate
        TZ6 series survived three runs unnoticed.
      - A section has to be runnable on its own. Clearing here is what makes
        `npm test -- 5` a complete statement rather than something that only
        works after section 3.

    Matches on the SUMMARY marker, never the UID: a clean pull mints fresh
    UIDs, so a UID-keyed sweep silently stops recognising its own litter.

    Returns the summaries removed, so the caller can say what it found.
    """
    import re
    from bridge import rpc

    # Default to what this run actually selected, never to every kind: a
    # section that needs only events has had tasks deselected, and querying
    # an unbound resource fails with "not bound to a calendar yet" - a
    # cleanup step reporting a failure about a resource the section never
    # wanted.
    if kinds is None:
        kinds = s.active or ("events",)

    removed = []
    for kind in kinds:
        if kind == "contacts":
            for card in s.cards():
                vcard = (card.get("properties") or {}).get("vCard") or ""
                if "probe.invalid" not in vcard:
                    continue
                removed.append(card["id"])
                rpc("contacts.remove", id=card["id"])
            continue
        type_ = "task" if kind == "tasks" else "event"
        args = {"resource": "tasks"} if kind == "tasks" else {}
        for item in s.items(kind, type_):
            body = item.get("item") or ""
            if not re.search(r"^SUMMARY:(TZ\d|" + MARKER + ")", body, re.M):
                continue
            m = re.search(r"^SUMMARY:([^\r\n]*)", body, re.M)
            removed.append(m.group(1)[:60] if m else item["id"])
            rpc("items.remove", id=item["id"], **args)
    if removed:
        s.sync()
        for kind in kinds:
            s.settle(kind)
        if report:
            print(f"       (cleared {len(removed)} leftover fixture(s))")
    return removed


def split_calendar(text):
    """Split a multi-item .ics into one VCALENDAR per UID.

    `items.create` takes a single item and refuses a file holding several -
    "Cannot parse more than one parent item". Grouping by UID rather than by
    VEVENT keeps a master and its overrides together, which is what makes
    them one item rather than several.

    Every VTIMEZONE in the source is carried into each result, since the
    events reference them by TZID.
    """
    import re

    head = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//eas-test//EN"]
    zones = re.findall(r"BEGIN:VTIMEZONE.*?END:VTIMEZONE", text, re.S)
    by_uid = {}
    for block in re.findall(r"BEGIN:VEVENT.*?END:VEVENT", text, re.S):
        uid = re.search(r"^UID:(.*)$", block, re.M)
        by_uid.setdefault(uid.group(1).strip() if uid else block[:40], []).append(block)
    out = []
    for uid, blocks in by_uid.items():
        body = head + zones + blocks + ["END:VCALENDAR"]
        out.append((uid, "\r\n".join(body) + "\r\n"))
    return out


def vevent_lines(body, prefix):
    """Lines starting with `prefix`, from the VEVENT components only.

    A pulled item carries the full historical VTIMEZONE - America/New_York
    alone contributes eighteen DTSTART lines, none of them the event's. A
    naive scan of the whole blob therefore compares timezone internals and
    reports a stable event as having moved.
    """
    out, depth = [], 0
    for raw in body.splitlines():
        if raw.startswith("BEGIN:VTIMEZONE"):
            depth += 1
        elif raw.startswith("END:VTIMEZONE"):
            depth -= 1
        elif depth == 0 and raw.startswith(prefix):
            out.append(raw)
    return out


def instants(body):
    """Every VEVENT DTSTART as a UTC instant, so two spellings of the same
    moment compare equal.

    A server may hand an override back in a different zone from the one it
    was sent in - `America/New_York 14:00` returns as `Europe/Berlin 20:00`,
    which is the same instant and not a defect. Comparing the literal lines
    calls that a shift; comparing instants does not.

    Date-only values (all-day) stay as dates: they are floating by
    definition, and converting them through a zone is precisely the bug an
    all-day test is looking for.
    """
    import re
    from datetime import datetime

    out = []
    for line in vevent_lines(body, "DTSTART"):
        value = line.split(":", 1)[1].strip() if ":" in line else ""
        if "VALUE=DATE" in line:
            out.append(("date", value))
            continue
        tzid = None
        m = re.search(r"TZID=([^:;]+)", line)
        if m:
            tzid = m.group(1)
        try:
            naive = datetime.strptime(value.rstrip("Z"), "%Y%m%dT%H%M%S")
        except ValueError:
            out.append(("raw", line))
            continue
        if value.endswith("Z"):
            out.append(("utc", naive.isoformat()))
            continue
        if tzid:
            try:
                from zoneinfo import ZoneInfo

                aware = naive.replace(tzinfo=ZoneInfo(tzid))
                out.append(("utc", aware.astimezone(tz=None).utcoffset() is not None
                            and aware.timestamp()))
                continue
            except Exception:
                pass
        out.append(("floating", naive.isoformat()))
    return out
