"""15. What a task carries that nobody can see.

Every other section asserts that a sync *succeeded* and that visible fields
survive a round trip. None of them looks at the bookkeeping stored on the
item itself - and that blindness hid a real bug for an afternoon.

`pinEasStamps` is handed every write the user makes. It strips `X-EAS-*`
from VEVENT and VTODO alike, then puts back whatever the previous version of
the item carried. The restore half looked the master up as a VEVENT, so for
a task it found nothing to restore onto: **every edit of a task deleted all
its stamps, `X-EAS-SERVERID` included.** Section 12 edits tasks on every
run and never noticed, because the push takes its ServerId from the
changelog rather than from the blob - a working fallback masking a data
loss. That is the shape of bug this section exists to catch, so it asserts
on the item and not on the outcome.

The recurrence tests are the second half of the same idea. `Regenerate`,
`DeadOccur`, `CalendarType` and `IsLeapMonth` have no iCalendar form and no
control in Thunderbird, but a push rebuilds `<Recurrence>` from the RRULE
and the server replaces the block wholesale - measured on 14.1 and 16.1 by
sending a weekly rule, changing it to daily, and finding the omitted
`DayOfWeek` gone from the server's own copy. Anything not handed back is
therefore destroyed by an edit that had nothing to do with recurrence.

16.x states `Regenerate` and `DeadOccur` on every task recurrence it sends,
which is what 15.3 uses: no hand-made regenerating task is needed, the
server volunteers the values and the test checks they come home again. On
14.x the server states neither, so that test is gated.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches.
NEEDS = ("tasks",)
# 15.3 and 15.4 are about recurrence, and with the account option off no
# rule reaches the wire at all.
NEEDS_RECURRENCE = True

FOLDERS_KEY = "tbsync.folders"

SLUG = "identity"
RECURRING = "identity-recurring"


def _blob(s, slug):
    item = s.find("tasks", slug, type_="task")
    harness.true(item is not None, f"the task {slug} is not in the local list")
    return item["item"] or ""


def _stamps(text):
    return sorted(l.split(":")[0] for l in text.splitlines() if l.startswith("X-EAS-"))


def _sent_recurrence(s, needle):
    """The `<Recurrence>` of a Change we sent naming `needle`, if any."""
    for entry in s.log():
        details = re.sub(r"\s+", " ", entry.get("details") or "")
        if "send" not in (entry.get("message") or "").lower():
            continue
        for change in re.finditer(r"<Change>.*?</Change>", details):
            if needle not in change.group(0):
                continue
            found = re.search(r"<Recurrence.*?</Recurrence>", change.group(0))
            if found:
                return re.sub(r" xmlns='[^']*'", "", found.group(0))
    return None


def _reread(s):
    """Make the next sync a full re-read, so the server re-states every item
    and we see what it actually holds. Section 13 does the same to a
    different end."""
    snap = ok("storage.snapshot")
    folders = snap[FOLDERS_KEY]
    row = folders[s.account_id][s.folders["tasks"]]["custom"]
    row["synckey"] = "0"
    row["indexMap"] = []
    ok("storage.restore", data={FOLDERS_KEY: folders})
    s.mark()
    s.sync()


@test("15.1", "a synced task carries the server's identity in its own blob")
def t_15_1(s):
    s.mark()
    ok("items.create", resource="tasks", type="task", ical=probes.task(SLUG))
    s.sync()
    harness.eq(s.changelog("tasks"), [], "the task was not accepted")
    harness.contains(
        _blob(s, SLUG),
        "X-EAS-SERVERID",
        "a task that has synced must carry the ServerId the server gave it",
    )


@test("15.2", "an edit that has nothing to do with it leaves the stamp alone")
def t_15_2(s):
    # The regression: the guard stripped every X-EAS-* from a VTODO and then
    # restored none of them. Nothing downstream complained, because the push
    # reads its ServerId from the changelog - which is exactly why this
    # asserts on the item.
    def rename(body):
        renamed = re.sub(
            r"SUMMARY:.*", f"SUMMARY:{probes.MARKER} {SLUG} renamed", body
        )
        # Sent without any stamp of its own, as Thunderbird's own UI would.
        return "\r\n".join(
            l for l in renamed.splitlines() if not l.startswith("X-EAS-")
        ) + "\r\n"

    # Asserted between the write and the sync on purpose: the stamp has to be
    # back because the guard restored it on the local write, not because a
    # sync put it back afterwards.
    def stamp_is_back():
        harness.contains(
            _blob(s, "renamed"),
            "X-EAS-SERVERID",
            "the edit stripped the ServerId stamp - a task loses its identity "
            "on every save, and only the changelog fallback hides it",
        )

    s.edit(
        lambda: s.find("tasks", SLUG, type_="task"),
        rename,
        resource="tasks",
        after_write=stamp_is_back,
        missing="15.1 did not leave its task behind",
    )
    harness.eq(s.changelog("tasks"), [], "the renamed task was not accepted")


@test(
    "15.3",
    "what the server says about a recurrence comes back on the next push",
    versions=("16",),
)
def t_15_3(s):
    s.mark()
    ok(
        "items.create",
        resource="tasks",
        type="task",
        ical=probes.task(RECURRING, "FREQ=WEEKLY;BYDAY=MO,WE;COUNT=4"),
    )
    s.sync()
    harness.eq(s.changelog("tasks"), [], "the recurring task was not accepted")

    # Ask the server what it now holds. 16.x states Regenerate and DeadOccur
    # on every task recurrence, so the values come from it, not from us.
    _reread(s)
    stored = _blob(s, RECURRING)
    harness.contains(
        stored,
        "X-EAS-REGENERATE",
        "the server states Regenerate on 16.x and it was not kept",
    )
    harness.contains(
        stored,
        "X-EAS-DEADOCCUR",
        "the server states DeadOccur on 16.x and it was not kept",
    )

    # Now the edit that used to destroy them.
    def rename(body):
        renamed = re.sub(
            r"SUMMARY:.*", f"SUMMARY:{probes.MARKER} {RECURRING} renamed", body
        )
        return "\r\n".join(
            l for l in renamed.splitlines() if not l.startswith("X-EAS-")
        ) + "\r\n"

    s.edit(
        lambda: s.find("tasks", RECURRING, type_="task"),
        rename,
        resource="tasks",
        missing="the recurring task is not there to rename",
    )

    sent = _sent_recurrence(s, "renamed")
    harness.true(sent is not None, "no <Recurrence> was sent for the renamed task")
    harness.contains(sent, "<Regenerate>", "the push dropped Regenerate")
    harness.contains(sent, "<DeadOccur>", "the push dropped DeadOccur")
    harness.true(
        sent.index("<Regenerate>") < sent.index("<Type>"),
        f"the unmapped elements must precede <Type>, as the server sends them: {sent}",
    )
    harness.eq(
        s.changelog("tasks"),
        [],
        "the server refused the block - the element order is wrong, which is "
        "the one detail no document settled",
    )


@test("15.4", "clean up - the section leaves the folder as it found it")
def t_15_4(s):
    s.mark()
    for slug in (SLUG, RECURRING):
        item = s.find("tasks", slug, type_="task")
        if item:
            ok("items.remove", resource="tasks", id=item["id"])
    s.sync()
    harness.eq(s.changelog("tasks"), [], "a deletion did not reach the server")
    for slug in (SLUG, RECURRING):
        harness.true(
            s.find("tasks", slug, type_="task") is None,
            f"{slug} is still in the folder",
        )
