"""12. Client-side rejections - items EAS cannot represent.

An item whose meaning the wire cannot carry is HELD, not mangled: the
push warns with the item and the reason, the folder shows the localized
"server did not accept N elements" warning, and the entry stays queued so
it retries every sync until the user changes or removes the item.
Retry-forever is deliberate - the same visibility decision as the
task-recurrence rejection policy: a lie on the wire is worse than a
nagging warning.

Three capability gaps drive the section: EAS has no recurrence frequency
below daily (an hourly event would silently sync as daily); a recurring
task needs an anchor - DTSTART or, since this feature, DUE; and EAS
always needs an EndTime, so an event may express its end as DURATION
(derived, representable) but an event with no expressed end - or one
not ending after its start - is held rather than pushed with an
invented end.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches.
NEEDS = ("events", "tasks")


def _event_hourly(slug):
    return probes.event(
        slug,
        lines=[
            "DTSTART:20261005T100000Z",
            "DTEND:20261005T103000Z",
            "RRULE:FREQ=HOURLY;COUNT=5",
        ],
    )


def _vtodo(slug, lines):
    """A bare VTODO. `probes.task` always carries DTSTART and DUE, and this
    section's whole point is controlling which anchors exist."""
    body = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//eas-test//EN",
        "BEGIN:VTODO",
        f"UID:{slug}@eas-test.invalid",
        "DTSTAMP:20260801T120000Z",
        f"SUMMARY:{probes.MARKER} {slug}",
    ]
    body += list(lines) + ["END:VTODO", "END:VCALENDAR"]
    return "\r\n".join(body) + "\r\n"


@test("12.1", "an hourly event is held, loudly - not silently sent as daily")
def t_12_1(s):
    s.mark()
    ok("items.create", type="event", ical=_event_hourly("subdaily"))
    s.sync(allow_errors=True)

    # Nothing went on the wire for it.
    harness.true(
        "SEND Add" not in s.wire(),
        "the hourly event was pushed - it will exist as a DAILY series "
        "server-side, the exact lie this feature removes",
    )
    # The folder says so, in the localized aggregate warning...
    harness.eq(s.status("events"), "warning", "folder status")
    # ...and the event log names the item and the reason.
    reasons = s.warnings("below daily")
    harness.true(reasons, "no warning names the reason")
    harness.contains(
        reasons[0], "subdaily", "the warning does not identify the item"
    )
    # The item survives locally, and its entry stays queued.
    harness.true(
        s.find("events", f"{probes.MARKER} subdaily", "event") is not None,
        "the local item vanished",
    )
    harness.true(s.changelog("events"), "the entry was dropped instead of held")


@test("12.2", "held means retried - and fixing the item releases it")
def t_12_2(s):
    # The nag is the contract: a second sync with no edit warns again.
    s.mark()
    s.sync(allow_errors=True)
    harness.eq(s.status("events"), "warning", "still warning")
    harness.true(s.warnings("below daily"), "the reason line repeats")
    harness.true(s.changelog("events"), "still queued")

    # Fix the rule - the very next sync pushes and everything goes green.
    item = s.find("events", f"{probes.MARKER} subdaily", "event")
    fixed = item["item"].replace("RRULE:FREQ=HOURLY;COUNT=5", "RRULE:FREQ=DAILY;COUNT=5")
    harness.true("FREQ=DAILY" in fixed, "the rule to fix was present")
    s.mark()
    ok("items.update", id=item["id"], ical=fixed)
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the fixed item must be pushed")
    harness.eq(s.status("events"), "success", "folder green after the fix")
    harness.eq(s.changelog("events"), [], "changelog drained")
    # Cleanup.
    item = s.find("events", f"{probes.MARKER} subdaily", "event")
    ok("items.remove", id=item["id"])
    s.sync()


@test("12.3", "a minutely task is held with the same mechanics")
def t_12_3(s):
    ical = _vtodo(
        "subdaily-task",
        ["DTSTART:20261006T090000Z", "RRULE:FREQ=MINUTELY;COUNT=10"],
    )
    s.mark()
    ok("items.create", type="task", ical=ical, resource="tasks")
    s.sync(allow_errors=True)
    harness.eq(s.status("tasks"), "warning", "folder status")
    harness.true(s.warnings("below daily"), "the reason line")
    harness.true(s.changelog("tasks"), "held, not dropped")
    # Cleanup: removing the offender must dig the folder out.
    item = s.find("tasks", f"{probes.MARKER} subdaily-task", "task")
    ok("items.remove", id=item["id"], resource="tasks")
    s.sync()
    harness.eq(s.changelog("tasks"), [], "changelog drained after removal")
    harness.eq(s.status("tasks"), "success", "folder green again")


@test("12.4", "a task with RRULE + DUE but no DTSTART syncs WITH its rule")
def t_12_4(s):
    # The representable half of the gap: DUE anchors the series, the way
    # Exchange itself models recurring tasks. Before this feature the rule
    # was silently dropped.
    ical = _vtodo(
        "due-anchored",
        ["DUE:20261007T170000Z", "RRULE:FREQ=WEEKLY;COUNT=4"],
    )
    s.mark()
    ok("items.create", type="task", ical=ical, resource="tasks")
    s.sync()
    harness.eq(s.status("tasks"), "success", "no warning for a DUE anchor")
    harness.eq(s.changelog("tasks"), [], "changelog drained")
    s.rebind("tasks")
    item = s.find("tasks", f"{probes.MARKER} due-anchored", "task")
    harness.true(item is not None, "the task did not survive the clean pull")
    harness.true(
        re.search(r"^RRULE", item["item"], re.M) is not None,
        "the recurrence was dropped - the DUE anchor did not reach the "
        "server and back",
    )
    ok("items.remove", id=item["id"], resource="tasks")
    s.sync()


@test("12.5", "a recurring task with neither start nor due is held")
def t_12_5(s):
    ical = _vtodo("anchorless", ["RRULE:FREQ=WEEKLY;COUNT=4"])
    s.mark()
    ok("items.create", type="task", ical=ical, resource="tasks")
    s.sync(allow_errors=True)
    harness.eq(s.status("tasks"), "warning", "folder status")
    harness.true(
        s.warnings("start or a due date"), "the anchor reason line"
    )
    harness.true(s.changelog("tasks"), "held, not dropped")
    # Cleanup.
    item = s.find("tasks", f"{probes.MARKER} anchorless", "task")
    ok("items.remove", id=item["id"], resource="tasks")
    s.sync()
    harness.eq(s.changelog("tasks"), [], "changelog drained after removal")


def _sent_xml(s):
    """Every decoded <Commands> payload we SENT since the last mark()."""
    out = []
    for e in s.log():
        if "send" not in (e.get("message") or "").lower():
            continue
        details = re.sub(r"\s+", " ", e.get("details") or "")
        if "<Commands>" in details:
            out.append(details)
    return " ".join(out)


@test("12.6", "a DURATION event is representable - the derived EndTime goes out")
def t_12_6(s):
    # RFC 5545 lets an event express its end as DURATION instead of DTEND
    # (imports do; Thunderbird's own editor writes DTEND). EAS always
    # needs an EndTime, so the push derives DTSTART+DURATION. Asserted on
    # the wire, where it matters - however the platform stores the blob.
    s.mark()
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            "duration-derived",
            lines=["DTSTART:20261110T140000Z", "DURATION:PT1H30M"],
        ),
    )
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the event must be pushed")
    harness.contains(
        _sent_xml(s),
        "20261110T153000Z",
        "the sent EndTime is DTSTART+PT1H30M, not an invented one",
    )
    harness.eq(s.status("events"), "success", "folder green")
    harness.eq(s.changelog("events"), [], "changelog drained")
    # Cleanup.
    item = s.find("events", f"{probes.MARKER} duration-derived", "event")
    ok("items.remove", id=item["id"])
    s.sync()


@test("12.7", "a DATE start with DURATION:P1D is an all-day event on the wire")
def t_12_7(s):
    # The quieter half of the gap: all-day used to be detected off DTEND
    # alone, so a DURATION all-day event went out as a timed one.
    s.mark()
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            "duration-allday",
            lines=["DTSTART;VALUE=DATE:20261111", "DURATION:P1D"],
        ),
    )
    s.sync()
    sent = _sent_xml(s)
    harness.contains(s.wire(), "SEND Add", "the event must be pushed")
    harness.contains(sent, ">1</AllDayEvent", "AllDayEvent flag")
    harness.contains(sent, "20261111T000000Z", "date-shaped StartTime")
    harness.contains(sent, "20261112T000000Z", "date-shaped derived EndTime")
    harness.eq(s.status("events"), "success", "folder green")
    # Cleanup.
    item = s.find("events", f"{probes.MARKER} duration-allday", "event")
    ok("items.remove", id=item["id"])
    s.sync()


@test("12.8", "an event with no derivable end is held - and fixing it releases it")
def t_12_8(s):
    # No DTEND, no DURATION: nothing may be invented, so the item is held
    # with the same mechanics as 12.1. The platform may normalise the
    # missing end to a zero-length DTEND when it stores the blob; both
    # shapes are invalid for us and both reasons name the end.
    s.mark()
    ok(
        "items.create",
        type="event",
        ical=probes.event("endless", lines=["DTSTART:20261112T090000Z"]),
    )
    s.sync(allow_errors=True)
    harness.true(
        "SEND Add" not in s.wire(),
        "the endless event was pushed - with an invented end",
    )
    harness.eq(s.status("events"), "warning", "folder status")
    reasons = [w for w in s.warnings("end") if "endless" in w]
    harness.true(reasons, "no warning names the item and the end problem")
    harness.true(s.changelog("events"), "held, not dropped")

    # Fix it - give the event a real end; the very next sync pushes.
    item = s.find("events", f"{probes.MARKER} endless", "event")
    harness.true(item is not None, "the local item survived the hold")
    fixed = item["item"].replace(
        "DTSTART:20261112T090000Z",
        "DTSTART:20261112T090000Z\r\nDTEND:20261112T093000Z",
    )
    s.mark()
    ok("items.update", id=item["id"], ical=fixed)
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the fixed item must be pushed")
    harness.eq(s.status("events"), "success", "folder green after the fix")
    harness.eq(s.changelog("events"), [], "changelog drained")
    # Cleanup.
    item = s.find("events", f"{probes.MARKER} endless", "event")
    ok("items.remove", id=item["id"])
    s.sync()
