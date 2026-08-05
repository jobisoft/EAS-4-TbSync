"""2. Item round trip - create, modify, delete.

Each step is edit -> sync -> assert, settling with a second sync before the
next edit. The wire assertions are the point: a local store that looks right
proves only that Thunderbird accepted the edit, not that the server heard
about it.

2.4 is the one that has actually caught something. The delete sync also
receives the server's echo of the 2.2 modification, and a pull that runs
before the push cannot tell that echo from a genuine server-side add - so it
re-creates the item the user just deleted.
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
