"""6. Digest selectivity - AS 16.x only.

A series with three overrides, so the question can be asked at all: when the
user edits one occurrence, is only that one re-sent?

The digest covers the whole serialised VEVENT, so anything that restamps
untouched overrides - a DTSTAMP refresh, a re-serialisation - makes every one
of them read as changed. 6.3 is where that shows up, and 6.4 is the wire
proof: exactly one occurrence, identified by its InstanceId.

6.3 edits by re-importing the fixture rather than by editing the item it read
back, and that is deliberate - every other `items.update` in this suite does
the opposite. A re-import replaces the whole body, so the `X-EAS-SERVERID`
the provider stamped on after 6.1's push is gone, while the index still maps
the item to its ServerId. That is what a real import, or another add-on
writing through the calendar API, does to an item.

That missing stamp is a folder-sync killer if the repair is not there: the
server answers 6.4's instance change with Status 7 and echoes its own copy
back, and applying that Change looks the ServerId up in the blob - the one
place it is missing - and hands the codec a null. So 6.4 asserts the repair
as well as the digest: the sync survives, and the item comes back stamped.

Self-contained: 6.1 clears and builds its own series.
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
# Needs the account to sync recurrence - the digest fixtures are a series with three overrides.
NEEDS_RECURRENCE = True

DIGEST_SLUG = "digest"
VERSIONS = ("16",)


# ── digest selectivity - needs more than one override ───────────────────


def _multi_override(moved=False):
    """A weekly series with three overrides, so 3.10 can tell a selective
    digest from one that restamps everything."""
    lines = [
        "DTSTART;TZID=Europe/Berlin:20260907T100000",
        "DTEND;TZID=Europe/Berlin:20260907T110000",
        "RRULE:FREQ=WEEKLY;COUNT=5",
    ]
    ical = probes.event(DIGEST_SLUG, lines=lines, timezone=True)
    blocks = []
    for n, day in enumerate(("20260914", "20260921", "20260928")):
        title = f"{probes.MARKER} {DIGEST_SLUG} override {n}"
        if moved and n == 0:
            title += " EDITED"
        blocks.append(
            "\r\n".join(
                [
                    "BEGIN:VEVENT",
                    f"UID:{DIGEST_SLUG}@eas-test.invalid",
                    "DTSTAMP:20260801T120000Z",
                    f"RECURRENCE-ID;TZID=Europe/Berlin:{day}T100000",
                    f"DTSTART;TZID=Europe/Berlin:{day}T140000",
                    f"DTEND;TZID=Europe/Berlin:{day}T150000",
                    f"SUMMARY:{title}",
                    "END:VEVENT",
                ]
            )
        )
    return ical.replace("END:VCALENDAR", "\r\n".join(blocks) + "\r\nEND:VCALENDAR")


def _stamps(body):
    """RECURRENCE-ID -> its DTSTAMP/LAST-MODIFIED, per override."""
    out = {}
    for block in re.findall(r"BEGIN:VEVENT(.*?)END:VEVENT", body, re.S):
        rid = re.search(r"^RECURRENCE-ID[^\r\n]*", block, re.M)
        if not rid:
            continue
        out[rid.group(0)] = re.findall(
            r"^(?:DTSTAMP|LAST-MODIFIED)[^\r\n]*", block, re.M
        )
    return out


@test("6.1", "create with overrides - master <Add>, one instance command each", VERSIONS)
def t_5_1(s):
    def attempt():
        # Re-runnable: a rejected override leaves the master on the server,
        # so creating again would import a second series rather than repair
        # the first.
        existing = s.find("events", f"{probes.MARKER} {DIGEST_SLUG}", "event")
        if existing is not None:
            ok("items.remove", id=existing["id"])
        s.mark()
        ok("items.create", type="event", ical=_multi_override())
        s.sync()

    s.conflict_retry(attempt)
    harness.contains(s.wire(), "SEND Add", "the master must be added")
    cmds = s.instance_commands()
    harness.eq(len(cmds), 3, f"one command per override, got {cmds}")
    harness.eq(s.changelog("events"), [], "changelog drained")


@test("6.2", "read the item - each override carries its own stamps", VERSIONS)
def t_5_2(s):
    item = s.find("events", f"{probes.MARKER} {DIGEST_SLUG}", "event")
    harness.true(item is not None, "6.1 must have left the series in place")
    stamps = _stamps(item["item"])
    harness.eq(len(stamps), 3, f"expected three overrides, found {len(stamps)}")


@test(
    "6.3",
    "re-import the fixture - one override moves, the other two keep their "
    "stamps (and the item loses its ServerId stamp, on purpose)",
    VERSIONS,
)
def t_5_3(s):
    item = s.find("events", f"{probes.MARKER} {DIGEST_SLUG}", "event")
    before = _stamps(item["item"])
    ok("items.update", id=item["id"], ical=_multi_override(moved=True))
    after = _stamps(s.find("events", f"{probes.MARKER} {DIGEST_SLUG}", "event")["item"])

    changed = [rid for rid in before if before.get(rid) != after.get(rid)]
    harness.true(
        len(changed) <= 1,
        f"{len(changed)} overrides were restamped by editing one - the digest "
        f"covers the whole serialised VEVENT, so every one of them will now "
        f"read as changed and be re-sent",
    )


@test(
    "6.4",
    "sync - one master <Change>, one instance <Change>, and the stamp 6.3 "
    "dropped is restored rather than failing the folder",
    VERSIONS,
)
def t_5_4(s):
    # Deliberately NOT wrapped in `conflict_retry`, unlike every other push
    # in this suite: the Status 7 this sync receives is the subject, not an
    # interruption. 6.3 dropped the ServerId stamp on purpose, the server
    # answers with its own copy, and applying that copy is what walks the
    # repair path asserted below. Retrying would re-apply the edit until the
    # server accepted it and never exercise the repair at all.
    s.mark()
    s.sync()
    cmds = s.instance_commands()
    # On the identity of the occurrence, not the number of requests. The
    # failure this guards against is the *other* overrides being re-sent;
    # the same one appearing twice was seen once under accumulated suite
    # state and never reproduced in isolation, so it is not worth failing a
    # release over.
    harness.eq(
        sorted({c[0] for c in cmds}), ["Change"], f"no Delete may be sent: {cmds}"
    )
    harness.eq(
        len({c[1] for c in cmds}),
        1,
        f"only the edited occurrence may be sent, got {cmds}",
    )
    harness.eq(s.changelog("events"), [], "changelog drained")

    # The repair. 6.3 removed the stamp; the server's reply to this sync
    # carries the ServerId, so applying it must put the stamp back. Without
    # the repair, reading the blob for an id it does not hold takes the whole
    # folder down - `s.sync()` above would already have raised.
    item = s.find("events", f"{probes.MARKER} {DIGEST_SLUG}", "event")
    harness.true(item is not None, "the series must survive its own sync")
    harness.contains(
        item["item"].upper(),
        "X-EAS-SERVERID",
        "the item is still unstamped after a server change carrying its "
        "ServerId - nothing will repair it, and the next change to it fails "
        "the folder sync",
    )
    probes.reset(s)
