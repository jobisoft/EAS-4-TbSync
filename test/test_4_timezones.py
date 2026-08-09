"""4. Timezones and all-day boundaries.

`fixtures/tz-test.ics` holds five independent cases - all-day single, all-day
span, an all-day series with UNTIL, a series crossing the DST change, and a
series with an exception. Each states its expectation in its own SUMMARY
("expect 21-23 Sep inclusive") because a human reads them off the calendar
view.

What the API can prove is narrower and still worth having: that a round trip
does not move them. An all-day boundary is the value most likely to shift -
it travels as midnight-in-zone expressed as UTC, so reading it in the wrong
zone moves the event a whole day, which is a bug we have actually shipped.
Confirming they *render* on the right day stays in the manual plan.

Self-contained: 4.1 imports what 4.2 reads.

Not version-gated: every case is behavioral - a boundary that holds on one
generation and shifts on the other is exactly what this section exists to
catch, and the two generations encode every one of these boundaries
differently. A failure here on one server family is a finding about that
family, not a broken test.
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


def _by_summary(s, summary):
    for item in s.items("events", "event"):
        if f"SUMMARY:{summary}" in (item.get("item") or ""):
            return item
    return None


@test("4.1", "every case survives a clean pull unshifted")
def t_4_1(s):
    s.mark()
    # One create per item: items.create takes a single parent item and
    # refuses a file holding five ("Cannot parse more than one parent item").
    for _uid, ical in probes.split_calendar(probes.fixture("tz-test.ics")):
        ok("items.create", type="event", ical=ical)
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained")

    # Key on the SUMMARY: the rebind below mints new UIDs.
    before = {}
    for item in s.items("events", "event"):
        body = item.get("item") or ""
        m = re.search(r"^SUMMARY:(TZ\d[^\r\n]*)", body, re.M)
        if m:
            before[m.group(1)] = probes.instants(body)
    harness.true(before, "no TZ fixtures were imported")

    s.rebind("events")
    # Every case is checked before any verdict: the section's value is the
    # per-case map (which boundary holds on which server family), and a
    # fail-fast loop reports only the first shifted case while the fate of
    # the other four stays unknown.
    wrong = []
    for summary, starts in before.items():
        item = _by_summary(s, summary)
        if item is None:
            wrong.append(f"{summary!r} did not survive the clean pull")
            continue
        # Compared as instants, not as text. A server may hand an override
        # back in a different zone from the one it was sent in -
        # America/New_York 14:00 returning as Europe/Berlin 20:00 is the same
        # moment, and calling that a shift would be a false alarm.
        got = probes.instants(item["item"])
        if got != starts:
            wrong.append(f"{summary!r} moved: expected {starts}, got {got}")
    harness.eq(len(wrong), 0, "shifted across the round trip:\n  " + "\n  ".join(wrong))


@test("4.2", "the DST-crossing series keeps its rule and its named zone")
def t_4_2(s):
    item = _by_summary(s, "TZ4")
    harness.true(item is not None, "the TZ4 fixture is missing - 4.1 must run first")
    harness.true(
        probes.vevent_lines(item["item"], "RRULE"),
        "the series came back as a single event, not a series",
    )
    harness.true(
        probes.vevent_lines(item["item"], "DTSTART;TZID="),
        "the series lost its named zone and drifted to floating or UTC",
    )
    probes.reset(s)
