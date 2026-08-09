"""11. Regressions from the v4 issue tracker.

Edge cases users actually hit, each naming its issue. Chosen for being
provokable through the bridge against a live server; the attendee-crash
class (#263) is deliberately absent - creating events with attendees makes
Exchange send real invitation mail, which a test must never do.

Runs against 14.1 and 16.x alike; the all-day and timezone cases are exactly
where those generations have diverged before.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches.
NEEDS = ("events", "tasks")

DAY = "20260916"
NEXT_DAY = "20260917"


def _event_body(s, marker):
    item = s.find("events", marker, "event")
    harness.true(item is not None, f"event {marker!r} not found")
    return item


@test("11.1", "all-day events stay all-day and one day long (#269, #280, #318)")
def t_11_1(s):
    # v4 sent a time-of-day for all-day events; depending on the server's
    # timezone handling the event then showed as two days in other clients,
    # or shifted a day entirely (the September 30th report). The two server
    # generations also encode the boundaries differently, which is where
    # the 14.1 divergence was found in v5.
    ical = "\r\n".join([
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "BEGIN:VEVENT",
        f"UID:allday-{DAY}@eas-test.invalid",
        "DTSTAMP:20260801T120000Z",
        f"DTSTART;VALUE=DATE:{DAY}",
        f"DTEND;VALUE=DATE:{NEXT_DAY}",
        f"SUMMARY:{probes.MARKER} allday-regression",
        "END:VEVENT",
        "END:VCALENDAR",
        "",
    ])
    ok("items.create", type="event", ical=ical)
    s.sync()
    s.rebind("events")
    body = _event_body(s, f"{probes.MARKER} allday-regression")["item"]

    m = re.search(r"^DTSTART[^:\r\n]*:(\S+)", body, re.M)
    harness.true(m is not None, "no DTSTART came back")
    start = m.group(1)
    harness.true(
        start.startswith(DAY),
        f"the all-day event moved: DTSTART came back as {start}, "
        f"expected the {DAY} boundary",
    )
    harness.true(
        "VALUE=DATE" in (re.search(r"^DTSTART[^\r\n]*", body, re.M).group(0))
        or start.endswith("T000000") or start.endswith("T000000Z"),
        f"the all-day event came back timed: {start}",
    )
    m_end = re.search(r"^DTEND[^:\r\n]*:(\S+)", body, re.M)
    harness.true(
        m_end is not None and m_end.group(1).startswith(NEXT_DAY),
        f"the event is no longer exactly one day: DTEND "
        f"{m_end.group(1) if m_end else None}, expected {NEXT_DAY}",
    )


@test("11.2", "multiline descriptions keep their lines and lose their CRs (#262)")
def t_11_2(s):
    # Outlook bodies arrive with \r\n line endings; v4 stored them verbatim
    # and every exported .ics then carried stray CRs that broke other
    # consumers. The codec must normalize to bare newlines.
    lines = ["Erste Zeile", "Zweite; mit Semikolon", "Dritte Zeile: äöü"]
    ical = probes.event(
        "crlf-regression",
        lines=[
            "DTSTART:20260916T100000Z",
            "DTEND:20260916T110000Z",
            "DESCRIPTION:" + "\\n".join(lines),
        ],
    )
    ok("items.create", type="event", ical=ical)
    s.sync()
    s.rebind("events")
    body = _event_body(s, f"{probes.MARKER} crlf-regression")["item"]

    m = re.search(r"^DESCRIPTION[^:\r\n]*:((?:[^\r\n]|\r\n[ \t])*)", body, re.M)
    harness.true(m is not None, "the description vanished in the round trip")
    value = m.group(1).replace("\r\n ", "").replace("\r\n\t", "")
    for line in lines:
        harness.contains(
            value.replace("\\;", ";"),
            line,
            "a description line was lost in the round trip",
        )
    harness.true(
        "\\r" not in value,
        f"carriage returns survived into the stored description - every "
        f".ics export of this event is now broken: {value[:120]!r}",
    )


@test("11.3", "deleting a completed task syncs cleanly (#217)")
def t_11_3(s):
    # v4 choked on the delete of a completed task - the server answered
    # with a status the sync treated as fatal, and the task then blocked
    # every following sync until it was deleted server-side by hand.
    ical = probes.task("done-task-regression")
    ok("items.create", type="task", ical=ical, resource="tasks")
    s.sync()
    item = s.find("tasks", f"{probes.MARKER} done-task-regression", "task")
    harness.true(item is not None, "the task was not created")

    completed = item["item"].replace(
        "END:VTODO",
        "STATUS:COMPLETED\r\nPERCENT-COMPLETE:100\r\n"
        "COMPLETED:20260916T120000Z\r\nEND:VTODO",
    )
    ok("items.update", id=item["id"], ical=completed, resource="tasks")
    s.sync()

    item2 = s.find("tasks", f"{probes.MARKER} done-task-regression", "task")
    harness.true(item2 is not None, "the completed task vanished")
    ok("items.remove", id=item2["id"], resource="tasks")
    s.sync()
    harness.eq(s.changelog("tasks"), [], "changelog drained after the delete")
    harness.true(
        s.find("tasks", f"{probes.MARKER} done-task-regression", "task") is None,
        "the completed task is still there after its delete",
    )
    # The follow-up sync is the regression: v4 kept failing here.
    s.sync()
    harness.eq(s.changelog("tasks"), [], "the delete keeps blocking the sync")
    harness.true(
        s.find("tasks", f"{probes.MARKER} done-task-regression", "task") is None,
        "the deleted task came back",
    )


@test("11.4", "round-tripped timezone ids stay clean (#168)")
def t_11_4(s):
    # v4 could build TZIDs with embedded zero bytes when pushing, and every
    # export of the pulled event then carried the garbage. The pulled TZID
    # must be a printable zone name and the hour must not shift.
    ical = probes.event(
        "tzid-regression",
        lines=[
            "DTSTART;TZID=Europe/Berlin:20260916T140000",
            "DTEND;TZID=Europe/Berlin:20260916T150000",
        ],
        timezone=True,
    )
    ok("items.create", type="event", ical=ical)
    s.sync()
    s.rebind("events")
    body = _event_body(s, f"{probes.MARKER} tzid-regression")["item"]

    m = re.search(r"^DTSTART;TZID=([^:;\r\n]+):(\S+)", body, re.M)
    harness.true(m is not None, f"no TZID came back; DTSTART lines:\n{body[:400]}")
    tzid, start = m.group(1), m.group(2)
    harness.true(
        re.fullmatch(r"[\x20-\x7EäöüÄÖÜß]+", tzid) and "\x00" not in tzid,
        f"the TZID carries non-printable garbage: {tzid!r}",
    )
    harness.true(
        start.endswith("T140000"),
        f"the event shifted: started 14:00 Berlin, came back {start} ({tzid})",
    )
