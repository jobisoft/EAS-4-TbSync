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
    item = _one(s)
    harness.true(item is not None, "2.1 must have left an item to modify")
    s.mark()
    ok(
        "items.update",
        id=item["id"],
        ical=item["item"].replace("LOCATION:Room 1", "LOCATION:Room 2"),
    )
    s.sync()
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
    'push' carries <Commands> without <GetChanges>, 'pull' carries
    <GetChanges>."""
    out = []
    for entry in s.log():
        if "send" not in (entry.get("message") or "").lower():
            continue
        details = entry.get("details") or ""
        if "<Sync" not in details:
            continue
        if "<GetChanges" in details:
            out.append(("pull", details))
        elif "<Commands>" in details:
            out.append(("push", details))
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
    s.mark()
    item = s.find("events", f"{probes.MARKER} {SLUG}-order", "event")
    harness.true(item is not None, "2.5 must have left an item to modify")
    # Touch ONLY the summary. A slug-wide replace also rewrites the UID
    # line, which re-keys the item - the later delete then cannot resolve
    # a ServerId and is (correctly) dropped, stranding the probe on the
    # server. Found the hard way: section 6's re-pull resurrected it.
    edited = item["item"].replace(
        f"SUMMARY:{probes.MARKER} {SLUG}-order",
        f"SUMMARY:{probes.MARKER} {SLUG}-order v2",
    )
    harness.true("order v2" in edited, "the summary to edit was present")
    ok("items.update", id=item["id"], ical=edited)
    s.sync()
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
