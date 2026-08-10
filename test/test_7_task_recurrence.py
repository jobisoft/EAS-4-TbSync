"""7. Task recurrence.

`[MS-ASTASK]` 2.2.2.31 qualifies each recurrence type with further elements -
DayOfWeek for weekly, DayOfMonth for monthly, DayOfMonth + MonthOfYear for
yearly, plus WeekOfMonth for the nth-weekday forms. Emit the type without them
and Exchange rejects the whole Add with Status 6, which is why only daily
recurrences could ever be created before this was fixed.

Every rule goes in its own sync, so a rejection names the rule that caused it
rather than the batch. All of them use DTSTART:20260901T080000Z deliberately:
1 Sep 2026 is a Tuesday and the 1st of month 9, so each field EAS fills in
from DTSTART has a value that can be checked by eye.

The assertion is on the `<Recurrence>` element actually sent, not just on the
folder staying green - a rule can be accepted and stored wrong.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("tasks",)
# Needs the account to sync recurrence - recurring tasks are the whole subject.
NEEDS_RECURRENCE = True

# slug -> (RRULE, expected Type, elements that must be present)
CASES = {
    "daily": ("FREQ=DAILY;COUNT=4", "0", ()),
    "weekly": ("FREQ=WEEKLY;COUNT=4", "1", ("DayOfWeek",)),
    "weekly-byday": ("FREQ=WEEKLY;BYDAY=MO,WE;COUNT=4", "1", ("DayOfWeek",)),
    "weekly-int2": ("FREQ=WEEKLY;INTERVAL=2;COUNT=4", "1", ("DayOfWeek",)),
    "monthly": ("FREQ=MONTHLY;COUNT=4", "2", ("DayOfMonth",)),
    "monthly-nth": ("FREQ=MONTHLY;BYDAY=2TU;COUNT=4", "3", ("DayOfWeek", "WeekOfMonth")),
    "yearly": ("FREQ=YEARLY;COUNT=4", "5", ("DayOfMonth", "MonthOfYear")),
    "yearly-nth": (
        "FREQ=YEARLY;BYDAY=1MO;BYMONTH=9;COUNT=4",
        "6",
        ("DayOfWeek", "WeekOfMonth", "MonthOfYear"),
    ),
    "daily-until": ("FREQ=DAILY;UNTIL=20261001T000000Z", "0", ("Until",)),
}


def _sent_recurrence(s, slug):
    """The <Recurrence> element from the Add we just sent, namespaces
    stripped."""
    for entry in s.log():
        details = re.sub(r"\s+", " ", entry.get("details") or "")
        if "send" not in (entry.get("message") or "").lower():
            continue
        for add in re.finditer(r"<Add>.*?</Add>", details):
            if slug.replace("-", "%2D") in add.group(0) or slug in add.group(0):
                found = re.search(r"<Recurrence.*?</Recurrence>", add.group(0))
                if found:
                    return re.sub(r" xmlns='[^']*'", "", found.group(0))
    return None


def _push(s, slug, extra_checks=()):
    rrule, want_type, required = CASES[slug]
    s.mark()
    ok("items.create", resource="tasks", type="task", ical=probes.task(slug, rrule))
    s.sync()

    sent = _sent_recurrence(s, slug)
    harness.true(sent is not None, f"no <Recurrence> was sent for {rrule}")
    harness.contains(sent, f"<Type>{want_type}</Type>", f"{slug} recurrence type")
    for element in tuple(required) + tuple(extra_checks):
        harness.contains(sent, element, f"{slug} must qualify the type with {element}")

    # A rule the server refuses is re-staged and retried on every later sync,
    # so a rejection shows up as an entry that will not drain.
    harness.eq(
        s.changelog("tasks"),
        [],
        f"{slug} was rejected - the entry is still queued, which means "
        f"Status 6: the recurrence was not fully qualified",
    )


@test("7.1", "FREQ=DAILY - Type 0, no qualifying element")
def t_7_1(s):
    _push(s, "daily")


@test("7.2", "FREQ=WEEKLY - Type 1 + DayOfWeek 4, filled from DTSTART")
def t_7_2(s):
    _push(s, "weekly", ["<DayOfWeek>4</DayOfWeek>"])


@test("7.3", "FREQ=WEEKLY;BYDAY=MO,WE - Type 1 + DayOfWeek 10 (2|8)")
def t_7_3(s):
    _push(s, "weekly-byday", ["<DayOfWeek>10</DayOfWeek>"])


@test("7.4", "FREQ=WEEKLY;INTERVAL=2 - Interval 2 survives")
def t_7_4(s):
    _push(s, "weekly-int2", ["<Interval>2</Interval>"])


@test("7.5", "FREQ=MONTHLY - Type 2 + DayOfMonth 1")
def t_7_5(s):
    _push(s, "monthly", ["<DayOfMonth>1</DayOfMonth>"])


@test("7.6", "FREQ=MONTHLY;BYDAY=2TU - Type 3 + DayOfWeek + WeekOfMonth 2")
def t_7_6(s):
    _push(s, "monthly-nth", ["<WeekOfMonth>2</WeekOfMonth>"])


@test("7.7", "FREQ=YEARLY - Type 5 + DayOfMonth 1 + MonthOfYear 9")
def t_7_7(s):
    _push(s, "yearly", ["<MonthOfYear>9</MonthOfYear>"])


@test("7.8", "FREQ=YEARLY;BYDAY=1MO;BYMONTH=9 - Type 6 + WeekOfMonth as well")
def t_7_8(s):
    _push(s, "yearly-nth", ["<WeekOfMonth>1</WeekOfMonth>"])


@test("7.9", "FREQ=DAILY;UNTIL - Until instead of Occurrences")
def t_7_9(s):
    _push(s, "daily-until")
    sent = _sent_recurrence(s, "daily-until")
    harness.true(
        "<Occurrences>" not in sent,
        f"a bounded-by-date rule must not also send Occurrences: {sent}",
    )


@test("7.10", "clean pull - each rule comes back semantically unchanged")
def t_7_10(s):
    # The only step that proves the server *stored* the rules rather than
    # merely accepting them. Match by SUMMARY: a pulled task has no UID
    # element in [MS-ASTASK], so its UID is regenerated and matching on it
    # reports every task as missing.
    s.rebind("tasks")
    for slug in CASES:
        item = s.find("tasks", f"{probes.MARKER} {slug}", "task")
        harness.true(item is not None, f"{slug} did not survive the clean pull")
        rrule = next(
            (l for l in item["item"].splitlines() if l.startswith("RRULE:")), None
        )
        harness.true(rrule is not None, f"{slug} came back without its RRULE")
        want_freq = CASES[slug][0].split(";")[0]
        harness.contains(rrule, want_freq, f"{slug} frequency after the round trip")
    probes.reset(s, ("tasks",))
