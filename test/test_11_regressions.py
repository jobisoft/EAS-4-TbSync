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
# 11.5 pushes a yearly rule, so the account has to sync recurrence.
NEEDS_RECURRENCE = True

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

    def complete(body):
        # Idempotent, as `edit` requires: a body that already carries the
        # completion is returned untouched rather than collecting a second
        # set of lines when the edit is made again.
        if "STATUS:COMPLETED" in body:
            return body
        return body.replace(
            "END:VTODO",
            "STATUS:COMPLETED\r\nPERCENT-COMPLETE:100\r\n"
            "COMPLETED:20260916T120000Z\r\nEND:VTODO",
        )

    s.edit(
        lambda: s.find("tasks", f"{probes.MARKER} done-task-regression", "task"),
        complete,
        resource="tasks",
        missing="the task vanished before it was completed",
    )

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


@test("11.5", "a rule sends no week start it was never given")
def t_11_5(s):
    """ical.js gives every parsed rule a `wkst`, Monday when the text says
    nothing, so reading the parsed value put `FirstDayOfWeek` on every rule
    pushed from here - a preference nobody expressed, at frequencies where a
    week start means nothing at all.

    The wire is the assertion, not the outcome: every server tried accepts
    the element without complaint, so a test that only watched the folder
    turn green would pass while the bug was still there. The rule is that
    what we send has to be what somebody asked for.

    A yearly rule on a fixed date because that is the case where the
    element is most obviously meaningless - but the same held for every
    frequency.
    """
    s.mark()
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            "wkst-regression",
            lines=[
                "DTSTART;VALUE=DATE:20260831",
                "DTEND;VALUE=DATE:20260901",
                # No WKST: the rule states nothing about the week start, so
                # nothing about one may leave here.
                "RRULE:FREQ=YEARLY;BYMONTHDAY=31;BYMONTH=8",
            ],
        ),
    )
    s.sync()

    pushed = [
        re.sub(r"\s+", " ", e.get("details") or "")
        for e in s.log()
        if "send" in (e.get("message") or "").lower()
        and "<Recurrence" in str(e.get("details") or "")
    ]
    harness.true(
        pushed,
        "no recurrence was pushed at all - the rule never reached the wire",
    )
    offenders = [p for p in pushed if "FirstDayOfWeek" in p]
    harness.eq(
        offenders,
        [],
        "a week start nobody asked for went out with the rule",
    )

    # And the folder survived it. On Z-Push the element is not merely
    # untidy - the push is answered with HTTP 500 and the folder goes red.
    harness.eq(s.status("events"), "success", "events folder status")
    s.settle("events")


@test("11.6", "acknowledging a reminder is not an edit (TbSync #816)")
def t_11_6(s):
    """Thunderbird keeps whether a reminder has been snoozed or dismissed
    inside the item, and saves it through the ordinary write path - on the
    parent, so dealing with one occurrence's alarm rewrites the series.
    Queued as a user edit it leaves as a <Change>, and a server that
    schedules for us then mails every attendee to say the meeting changed.

    Nothing on the wire carries either property, so that <Change> announced
    a change that was not there. The assertion is therefore the queue and
    not the server: an item whose only difference is the acknowledgement
    must not be queued at all.

    The second half is the one that matters most. A filter this close to the
    user's save is one wrong predicate away from swallowing real edits in
    silence, so a title change through the same path has to still arrive.
    """
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            "alarm-ack",
            lines=[
                "DTSTART:20260916T090000Z",
                "DTEND:20260916T093000Z",
                "BEGIN:VALARM",
                "ACTION:DISPLAY",
                "TRIGGER:-PT15M",
                "DESCRIPTION:Reminder",
                "END:VALARM",
            ],
        ),
    )
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained before the test")

    # What `dismissAlarm` writes: the acknowledgement, and nothing else.
    item = _event_body(s, "alarm-ack")
    dismissed = item["item"].replace(
        "END:VALARM", "END:VALARM\r\nX-MOZ-LASTACK:20260916T084500Z"
    )
    harness.true(dismissed != item["item"], "the fixture did not gain the stamp")
    ok("items.update", id=item["id"], ical=dismissed)
    harness.eq(
        s.changelog("events"),
        [],
        "a dismissed reminder was queued as a user edit",
    )

    # ...and a real edit through the very same path still is.
    item = _event_body(s, "alarm-ack")
    ok(
        "items.update",
        id=item["id"],
        ical=item["item"].replace("alarm-ack", "alarm-ack renamed"),
    )
    harness.eq(
        len(s.changelog("events")),
        1,
        "a title change was swallowed with the alarm bookkeeping",
    )
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained")

    ok("items.remove", id=_event_body(s, "alarm-ack renamed")["id"])
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained after cleanup")


@test("11.7", "a reminder answered stays answered across a resync (TbSync #816)")
def t_11_7(s):
    """The other half of 11.6. `<Reminder>` in the server's answer is its
    authority over the alarm, so the merge rebuilds the VALARM - and used to
    drop the user's dismissal with it. Exchange re-sends that element
    whenever anything else about the meeting changes, so an organiser moving
    a room resurrected every attendee's dismissed reminder.

    Driven by a resync rather than a real server-side edit, because that is
    the same code: an inbound <Add> for an item we already hold by ServerId
    is routed into `applyChangeFromAd` with the prior blob, so the merge runs
    exactly as it would on a Change. Nothing else in the suite can make a
    server send a Change carrying an unchanged Reminder on demand.
    """
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            "ack-resync",
            lines=[
                "DTSTART:20260916T110000Z",
                "DTEND:20260916T113000Z",
                "BEGIN:VALARM",
                "ACTION:DISPLAY",
                "TRIGGER:-PT15M",
                "DESCRIPTION:Reminder",
                "END:VALARM",
            ],
        ),
    )
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained before the test")

    # Answer the reminder, the way a dismissal reaches the item. 11.6 proves
    # this is not queued, so the server never hears about it - which is
    # exactly why the local copy is the only place it exists.
    item = _event_body(s, "ack-resync")
    ok(
        "items.update",
        id=item["id"],
        ical=item["item"].replace(
            "END:VALARM", "END:VALARM\r\nX-MOZ-LASTACK:20260916T104500Z"
        ),
    )
    harness.true(
        "X-MOZ-LASTACK" in _event_body(s, "ack-resync")["item"],
        "the dismissal did not reach the stored item",
    )

    # A sync key nobody issued: the server refuses it and re-sends the folder,
    # each item as an <Add> that finds its existing twin. Written straight
    # into host storage the way 19.2 does - there is no verb for this.
    snap = ok("storage.snapshot")
    folders = snap["tbsync.folders"]
    folders[s.account_id][s.folders["events"]]["custom"]["synckey"] = "999999999"
    ok("storage.restore", data={"tbsync.folders": folders})
    s.sync()

    harness.true(
        "X-MOZ-LASTACK" in _event_body(s, "ack-resync")["item"],
        "the resync cleared the dismissal - the reminder will ring again",
    )

    ok("items.remove", id=_event_body(s, "ack-resync")["id"])
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained after cleanup")
