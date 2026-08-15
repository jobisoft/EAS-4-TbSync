"""17. All-day exceptions bind by VALUE TYPE - every protocol version.

Split out of section 3, where it sat behind five 16.x-gated tests and a
chained fixture: on a 14.1 account it was the section's only running test,
and on 16.x it never ran at all whenever the move before it failed.

It is version-agnostic, which is the point - <=14.x carries exceptions
embedded in the master's payload and 16.x sends them as per-instance
commands, and the stored form must come out the same either way. So it wants
to be a section of its own rather than the tail of a gated one.

Self-contained: it builds its own all-day series.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

NEEDS = ("events",)


@test("17.1", "all-day exceptions bind - DATE-valued on both wire forms")
def t_17_1(s):
    # Not version-gated: <=14.x carries these embedded in the master's
    # payload, 16.x as per-instance commands, and the stored form must be
    # the same either way. RFC 5545 §3.8.4.4 matches RECURRENCE-ID/EXDATE
    # to an occurrence by VALUE TYPE, so an all-day master - whose DTSTART
    # is a DATE - binds only DATE-valued exceptions. Written as DATE-TIME
    # (the pre-fix form) the override floats beside the series and a
    # cancelled day keeps rendering.
    #
    # Two-step on purpose: at least one 14.1 server (ekir, running Z-Push)
    # accepts an all-day <Add> whose payload embeds <Exceptions> and then
    # quietly drops the recurrence, series and all. The same exceptions
    # sent as a <Change> on the following sync survive. The same server
    # also re-anchors an all-day series one day early, so every assertion
    # below is relative to the DTSTART that comes back, not to the dates
    # that went in.
    ical = "\r\n".join([
        "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//eas-test//EN",
        "BEGIN:VEVENT",
        "UID:allday-bind@eas-test.invalid",
        "DTSTAMP:20260801T120000Z",
        f"SUMMARY:{probes.MARKER} allday-bind",
        "DTSTART;VALUE=DATE:20261012",
        "DTEND;VALUE=DATE:20261013",
        "RRULE:FREQ=DAILY;COUNT=3",
        "END:VEVENT", "END:VCALENDAR", "",
    ])
    ok("items.create", type="event", ical=ical)
    s.sync()
    item = s.find("events", f"{probes.MARKER} allday-bind", "event")
    harness.true(item is not None, "the all-day series did not reach the calendar")

    body = item["item"]
    upd = body.replace(
        "RRULE:FREQ=DAILY;COUNT=3",
        "RRULE:FREQ=DAILY;COUNT=3\r\nEXDATE;VALUE=DATE:20261014",
    ).replace("END:VCALENDAR", "\r\n".join([
        "BEGIN:VEVENT",
        "UID:allday-bind@eas-test.invalid",
        "RECURRENCE-ID;VALUE=DATE:20261013",
        "DTSTAMP:20260801T120000Z",
        f"SUMMARY:{probes.MARKER} allday-bind OVERRIDDEN",
        "DTSTART;VALUE=DATE:20261013",
        "DTEND;VALUE=DATE:20261014",
        "END:VEVENT", "END:VCALENDAR",
    ]))
    ok("items.update", id=item["id"], ical=upd)
    s.sync()
    s.rebind("events")

    import datetime

    pulled = s.find("events", f"{probes.MARKER} allday-bind", "event")
    harness.true(pulled is not None, "the series did not survive the clean pull")
    body = pulled["item"]
    harness.contains(body, "RRULE", "the recurrence must survive the round trip")

    m = re.search(r"^DTSTART;VALUE=DATE:(\d{8})", body, re.M)
    harness.true(m is not None, "the master must come back as a DATE")
    day1 = datetime.datetime.strptime(m.group(1), "%Y%m%d").date()
    grid = lambda n: (day1 + datetime.timedelta(days=n)).strftime("%Y%m%d")

    exdates = [l for l in body.splitlines() if l.startswith("EXDATE")]
    harness.eq(
        exdates,
        [f"EXDATE;VALUE=DATE:{grid(2)}"],
        "the cancelled day must be a DATE on the occurrence grid - as a "
        "DATE-TIME it excludes nothing and the day keeps rendering",
    )
    rids = [l for l in body.splitlines() if l.startswith("RECURRENCE-ID")]
    harness.eq(
        rids,
        [f"RECURRENCE-ID;VALUE=DATE:{grid(1)}"],
        "the override must anchor as a DATE on the occurrence grid - as a "
        "DATE-TIME it binds nothing and floats beside the series",
    )
    harness.contains(body, "allday-bind OVERRIDDEN", "the override's content")
    probes.reset(s)
