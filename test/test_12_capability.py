"""12. Client-side rejections - items EAS cannot represent.

An item whose meaning the wire cannot carry is HELD, not mangled: the
push warns with the item and the reason, the folder shows the localized
"server did not accept N elements" warning, and the entry stays queued so
it retries every sync until the user changes or removes the item.
Retry-forever is deliberate - the same visibility decision as the
task-recurrence rejection policy: a lie on the wire is worse than a
nagging warning.

Two capability gaps drive the section: EAS has no recurrence frequency
below daily (an hourly event would silently sync as daily), and a
recurring task needs an anchor - DTSTART or, since this feature, DUE;
only a rule with neither is held.
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
