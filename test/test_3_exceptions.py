"""3. Recurrence exceptions - AS 16.x only, except 3.6.

One series, one cancelled occurrence and one moved one, from
`fixtures/tz-test-exdate.ics`. On 16.x each exception travels as its own
top-level command carrying an `InstanceId`, where <=14.x embeds
`<Exceptions>` inside the master's payload - so this section is gated, and
the same behaviour on 14.1 is correct while looking completely different on
the wire. 3.6 is the exception to the gate: the all-day DATE form must hold
on both wire shapes, so it runs everywhere.

3.3 is the regression test for the re-assertion bug: before the exception
fingerprint landed, touching the master re-sent every occurrence and Exchange
rejected each one it already had.

Self-contained - 3.1 clears and imports - so `npm test -- 3` is a complete
run. The steps within it deliberately chain, because re-importing the series
per step would mean four more full syncs against a server that throttles.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("events",)

# Matched on the SUMMARY, never the UID: 3.5's clean pull mints fresh UIDs,
# so a UID-keyed lookup reports an intact series as missing.
MARK = "SUMMARY:TZ6 weekly"
VERSIONS = ("16",)


def _series(s):
    for item in s.items("events", "event"):
        if MARK in (item.get("item") or ""):
            return item
    return None


@test("3.1", "import - Delete 20260916T130000Z and Change 20260909T130000Z", VERSIONS)
def t_3_1(s):
    s.mark()
    ok("items.create", type="event", ical=probes.fixture("tz-test-exdate.ics"))
    s.sync()

    cmds = s.instance_commands()
    harness.true(cmds, "no instance commands were sent for the exceptions")
    harness.contains(cmds, ("Delete", "20260916T130000Z"), "cancelled occurrence")
    harness.contains(cmds, ("Change", "20260909T130000Z"), "moved occurrence")
    harness.eq(s.changelog("events"), [], "changelog drained")


@test("3.2", "sync again with no edit - no instance commands", VERSIONS)
def t_3_2(s):
    s.mark()
    s.sync()
    harness.eq(s.instance_commands(), [], "an unchanged series must send nothing")


@test("3.3", "edit only the master's title - no instance commands", VERSIONS)
def t_3_3(s):
    item = _series(s)
    harness.true(item is not None, "3.1 must have left the series in place")
    s.mark()
    ok(
        "items.update",
        id=item["id"],
        ical=item["item"].replace("SUMMARY:TZ6 weekly", "SUMMARY:TZ6 weekly edited"),
    )
    s.sync()
    harness.eq(
        s.instance_commands(),
        [],
        "editing the master re-asserted its exceptions - each one the server "
        "already holds, and rejects",
    )


@test("3.4", "move the override - exactly one <Change>, no <Delete>", VERSIONS)
def t_3_4(s):
    item = _series(s)
    body = item["item"]
    # The override's DTSTART is representation-fragile: the fixture wrote
    # America/New_York 13:00, but as soon as any server echo rebuilds the
    # item (Exchange re-sends the master with its Exceptions shortly after
    # a push, and the post-push pull may pick that up), the codec renders
    # the same instant in the default timezone - Europe/Berlin 19:00. So
    # find the override COMPONENT by its RECURRENCE-ID and move whatever
    # DTSTART it carries one hour later; both forms land on the same UTC
    # instant, which is what 3.5 verifies after the clean pull.
    block_re = re.compile(
        r"BEGIN:VEVENT(?:(?!BEGIN:VEVENT)[\s\S])*?"
        r"RECURRENCE-ID[^\r\n]*20260909T130000Z[\s\S]*?END:VEVENT"
    )
    m = block_re.search(body)
    harness.true(
        m is not None,
        "the 9 Sep override is not in the local item - it never survived "
        "the 3.1 import (check 3.1's instance <Change> for a rejection)",
    )
    block = m.group(0)
    dt = re.search(r"DTSTART;TZID=([^:;\r\n]+):20260909T(\d{2})(\d{4})", block)
    harness.true(dt is not None, "the override carries no TZID DTSTART to move")
    hour = int(dt.group(2)) + 1
    moved_block = block.replace(
        dt.group(0), f"DTSTART;TZID={dt.group(1)}:20260909T{hour:02d}{dt.group(3)}"
    )
    s.mark()
    ok("items.update", id=item["id"], ical=body.replace(block, moved_block))
    s.sync()
    cmds = s.instance_commands()
    harness.eq([c[0] for c in cmds], ["Change"], f"instance commands sent: {cmds}")


@test("3.5", "clean resync - one EXDATE, the move and the cancellation intact", VERSIONS)
def t_3_5(s):
    s.rebind("events")
    item = _series(s)
    harness.true(item is not None, "the series did not survive the clean pull")
    body = item["item"]

    # Count EXDATE *lines*. A substring count also matches the fixture's own
    # description - "16 Sep is cancelled by the EXDATE and must not appear at
    # all" - and reports one cancellation as two. That false positive
    # reproduced on two different 16.x servers before it was spotted, which
    # is exactly how a test bug earns a reputation as a product bug.
    exdates = [l for l in body.splitlines() if l.startswith("EXDATE")]
    harness.eq(len(exdates), 1, f"expected one EXDATE property, got {exdates}")
    harness.contains(exdates[0], "20260916", "the cancelled date")
    harness.contains(body, "RECURRENCE-ID", "the moved occurrence must come back")
    probes.reset(s)


@test("3.6", "all-day exceptions bind - DATE-valued on both wire forms")
def t_3_6(s):
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
