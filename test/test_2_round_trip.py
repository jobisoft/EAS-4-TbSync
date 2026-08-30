"""2. Item round trip - create, modify, delete.

Each step is edit -> sync -> assert, settling with a second sync before the
next edit. The wire assertions are the point: a local store that looks right
proves only that Thunderbird accepted the edit, not that the server heard
about it.

2.4 is the one that has actually caught something. When the pull still ran
before the push, the delete sync's pull received the server's echo of the
2.2 modification and re-created the item the user just deleted. The sync
pushes first now - 2.5 pins that order, 2.6 the <Conflict> preference that
puts genuine two-writer conflicts in the server's hands - and 2.4 remains
as the regression net for the old failure.
"""

import re

import harness
import probes
from bridge import ok, rpc
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("events",)

SLUG = "round-trip"
MARK = f"SUMMARY:{probes.MARKER} {SLUG}"


def _one(s):
    return s.find("events", MARK, "event")


@test("2.1", "items.create, sync - one <Add>; item present")
def t_2_1(s):
    before = len(s.items("events", "event"))
    s.mark()
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            SLUG,
            lines=[
                "DTSTART;TZID=Europe/Berlin:20261110T093000",
                "DTEND;TZID=Europe/Berlin:20261110T104500",
                "LOCATION:Room 1",
            ],
            timezone=True,
        ),
    )
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the create must reach the server")
    harness.true(_one(s) is not None, "item present after the push")
    harness.eq(len(s.items("events", "event")), before + 1, "item count")
    harness.eq(s.changelog("events"), [], "changelog drained")


@test("2.2", "items.update, sync - one <Change>; the local item shows the change")
def t_2_2(s):
    s.edit(
        lambda: _one(s),
        lambda body: body.replace("LOCATION:Room 1", "LOCATION:Room 2"),
        missing="2.1 must have left an item to modify",
    )
    harness.contains(s.wire(), "SEND Change", "the edit must reach the server")
    harness.contains(_one(s)["item"], "Room 2", "the local item")
    harness.eq(s.changelog("events"), [], "changelog drained")


@test("2.3", "items.remove, sync - one <Delete>; item gone locally")
def t_2_3(s):
    item = _one(s)
    harness.true(item is not None, "2.2 must have left an item to delete")
    s.mark()
    ok("items.remove", id=item["id"])
    s.sync()
    harness.contains(s.wire(), "SEND Delete", "the delete must reach the server")
    harness.true(_one(s) is None, "item gone locally")
    harness.eq(s.changelog("events"), [], "changelog drained")


@test("2.4", "sync again - count back to baseline; the echo must not re-create it")
def t_2_4(s):
    # No edit here. The only thing this sync does is read what the server
    # says, which still includes its echo of the 2.2 change for an item we
    # have since deleted.
    s.sync()
    harness.true(
        _one(s) is None,
        "the deleted item came back - the pull re-created it from the "
        "server's echo of an edit we had already deleted",
    )
    s.sync()
    harness.true(_one(s) is None, "still gone after a second settling sync")


def _sync_requests(s):
    """The Sync requests of the marked window, in order sent, classified:
    'push' carries <Commands>, 'pull' asks for changes.

    Classified on <Commands> first, because a push now carries a GetChanges
    element too - `<GetChanges>0</GetChanges>`, which is the client saying it
    wants none. Reading the element's presence alone, as this once did, calls
    every push a pull and quietly inverts the ordering 2.5 asserts.
    """
    out = []
    for entry in s.log():
        if "send" not in (entry.get("message") or "").lower():
            continue
        details = entry.get("details") or ""
        if "<Sync" not in details:
            continue
        flat = re.sub(r"\s+", "", details)
        if "<Commands>" in flat:
            out.append(("push", details))
        elif "<GetChanges" in flat and "<GetChanges>0</GetChanges>" not in flat:
            out.append(("pull", details))
    return out


@test("2.5", "a pending edit is pushed BEFORE the pull runs")
def t_2_5(s):
    # The order is the pull-clobber fix: pushing first means a pending
    # local edit reaches the server before the pull could overwrite it
    # with a server copy, and a genuine two-writer conflict is decided by
    # the server rather than by our own pull.
    s.mark()
    ok("items.create", type="event", ical=probes.event(
        f"{SLUG}-order",
        lines=["DTSTART:20261111T100000Z", "DTEND:20261111T110000Z"],
    ))
    s.sync()
    kinds = [kind for kind, _ in _sync_requests(s)]
    harness.contains(kinds, "push", "no Commands request was sent")
    harness.contains(kinds, "pull", "no GetChanges request was sent")
    harness.true(
        kinds.index("push") < kinds.index("pull"),
        f"the pull ran before the push: request order was {kinds}",
    )


@test("2.6", "every request states the conflict policy - server wins by default")
def t_2_6(s):
    # <Conflict>1</Conflict> makes the implicit server default an explicit
    # contract: on a two-writer conflict the server keeps its copy, answers
    # the losing <Change> with Status 7, and the same sync's pull delivers
    # the winning version - the losing edit is replaced, never silently
    # dropped.
    # Touch ONLY the summary. A slug-wide replace also rewrites the UID line,
    # which re-keys the item - the later delete then cannot resolve a ServerId
    # and is (correctly) dropped, stranding the probe on the server. Found the
    # hard way: section 7's re-pull resurrected it.
    def bump(body):
        # Idempotent: "-order" is a prefix of "-order v2", so a second pass
        # would otherwise append the suffix twice.
        if "order v2" in body:
            return body
        edited = body.replace(
            f"SUMMARY:{probes.MARKER} {SLUG}-order",
            f"SUMMARY:{probes.MARKER} {SLUG}-order v2",
        )
        harness.true("order v2" in edited, "the summary to edit was present")
        return edited

    s.edit(
        lambda: s.find("events", f"{probes.MARKER} {SLUG}-order", "event"),
        bump,
        missing="2.5 must have left an item to modify",
    )
    reqs = _sync_requests(s)
    missing = [kind for kind, details in reqs if "<Conflict>1</Conflict>" not in details.replace(" ", "")]
    harness.eq(
        missing,
        [],
        "requests without the <Conflict> preference - the server is left "
        "to its undeclared default",
    )
    # Cleanup: this slug's last test. The drained changelog is the proof
    # the delete actually resolved and went out - a dropped delete leaves
    # the probe on the server for a later section to trip over.
    item = s.find("events", f"{probes.MARKER} {SLUG}-order", "event")
    ok("items.remove", id=item["id"])
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained after cleanup")


@test("2.7", "a pushed item survives a deselect/reselect - it comes back from the server")
def t_2_7(s):
    """The only test that proves an Add reached the server rather than
    merely leaving the queue.

    `rebind` deletes the local calendar and pulls a fresh one, so anything
    still visible afterwards came from the server. Everything else in this
    section could pass on a purely local copy: the changelog drains when the
    push is *sent*, and a server that quietly dropped the item would look
    identical until the next clean pull.
    """
    slug = f"{SLUG}-teardown"
    s.mark()
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            slug,
            lines=[
                "DTSTART:20261201T090000Z",
                "DTEND:20261201T100000Z",
                "LOCATION:Room 7",
            ],
        ),
    )
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the create must reach the server")
    harness.eq(s.changelog("events"), [], "changelog drained")

    # Throw the local copy away and pull it back down.
    s.rebind("events")

    item = s.find("events", f"{probes.MARKER} {slug}", "event")
    harness.true(
        item is not None,
        "the item did not come back after the local calendar was rebuilt - "
        "it never actually reached the server",
    )
    harness.contains(item["item"], "Room 7", "the item came back incomplete")
    # Cleanup.
    ok("items.remove", id=item["id"])
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained after cleanup")


@test("2.8", "an item created without a UID is still queued, pushed and stored")
def t_2_8(s):
    """The shape Thunderbird's own dialog produces.

    Core decides an edit is an *addition* by the absence of an id, and the
    id is minted after our hook runs - so a UI-created item reaches the
    provider with none. Every fixture in this suite carries a UID, which
    is exactly why this went unnoticed until a user lost events to it: the
    hook dropped the item and nothing was ever queued.

    `probes.event` always writes one, so the fixture is built here.
    """
    ical = "\r\n".join(
        [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//eas-test//EN",
            "BEGIN:VEVENT",
            "DTSTAMP:20260801T120000Z",
            f"SUMMARY:{probes.MARKER} {SLUG}-nouid",
            "DTSTART:20261203T090000Z",
            "DTEND:20261203T100000Z",
            "END:VEVENT",
            "END:VCALENDAR",
            "",
        ]
    )
    harness.true("UID:" not in ical, "the fixture must not carry a UID")

    s.mark()
    ok("items.create", type="event", ical=ical)
    harness.true(
        s.changelog("events"),
        "nothing was queued - the edit was dropped before it could be "
        "recorded, and a teardown would take it with no trace",
    )
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the create must reach the server")
    harness.eq(s.changelog("events"), [], "changelog drained")

    # And it really is on the server, not just locally.
    s.rebind("events")
    item = s.find("events", f"{probes.MARKER} {SLUG}-nouid", "event")
    harness.true(item is not None, "the item did not survive the rebuild")
    # Cleanup.
    ok("items.remove", id=item["id"])
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained after cleanup")
