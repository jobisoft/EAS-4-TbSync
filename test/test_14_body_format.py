"""14. Note formats: what the server holds vs what it hands us.

A note lives in two places on the item: the DESCRIPTION value, which the
tooltip and list views read, and its ALTREP parameter, which the editor
reads. Keeping those two honest is the whole of this section.

The protocol fact everything here rests on: a server answers in the `Type`
the client asked for, and reports what it ACTUALLY holds separately, in
`NativeBodyType` ([MS-ASAIRS] 2.2.2.32 - 1 plain, 2 HTML, 3 RTF; the two
agree unless the server converted the body to satisfy the request). So we
ask for plain text everywhere and read the truth off NativeBodyType: items
that say 2 are re-fetched as HTML - one request per window, a <Fetch> per
item - and only then does an ALTREP get stored.

Asking for HTML instead is what caused #347's regression: Exchange answers
a plain note by generating an HTML document around it, we stored that, and
pushed it back as a real body edit - which bumped the item's version and
drew Status 7 on the instance command behind it. Kopano generates a wrapper
of its own, on plain notes and on empty ones. Neither wrapper can reach us
now, because neither server is ever asked for HTML in a Sync.

Three events, deliberately typed plain / rich / plain: one rich item among
plain ones is what proves the fetch is selective rather than blanket.

Runs on every version - AirSyncBase Body exists from 12.0, and 2.5 has no
BodyPreference at all, which the codec handles separately.
"""

import xml.etree.ElementTree as ET
from urllib.parse import unquote

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches.
NEEDS = ("events",)

# Multi-line on purpose. Servers separate the lines of a note with CRLF,
# which cannot live in an iCalendar TEXT value, so the note is normalised on
# the way in and the server's own line endings are restored on the way out.
# Every fixture here used to be a single line, which is why nothing caught a
# note being rewritten on every push.
PLAIN_ONE = "plain note one\nsecond line\nthird line"
# The same note as an iCalendar TEXT value: RFC 5545 §3.3.11 escapes a line
# break, so a fixture that embedded the raw newline would just be a broken
# property. What comes back out of the store is the unescaped form above.
PLAIN_ONE_ICAL = PLAIN_ONE.replace("\n", "\\n")
PLAIN_TWO = "plain note two"
RICH_HTML = "<b>bold</b> and <i>italic</i>"
RICH_TEXT = "bold and italic"
PROMOTED_HTML = "<u>promoted</u> again"
PROMOTED_TEXT = "promoted again"

SLUGS = ("body-plain-a", "body-rich-b", "body-plain-c")

# Which items the server said it holds as HTML, recorded by 14.2 rather
# than assumed. Storing a note as HTML is a server capability, not a
# protocol guarantee: Kerio Connect converts one to text on the way in and
# then reports NativeBodyType 1, which is the honest answer for what it
# holds. Everything below 14.2 needs a genuinely rich note to judge, so it
# is gated on this rather than failing where there is nothing to test.
_RICH = None


def _needs_rich():
    if not _RICH:
        raise harness.Skip(
            "the server stores notes as plain text - it reported "
            "NativeBodyType 1 for a body sent as HTML"
        )


def _needs_plain_only():
    if _RICH:
        raise harness.Skip(
            "the server keeps the HTML it is given, so there is no "
            "downgrade to judge"
        )


# ── wire reading ──────────────────────────────────────────────────────────
# Parsed, never pattern-matched: `<Body>` and `<BodyPreference>` share a
# prefix, and the decoder writes a namespace attribute on every element, so
# a text search finds the wrong element about as often as the right one.


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _child(el, name):
    for c in list(el):
        if _local(c.tag) == name:
            return c
    return None


def _descendants(el, name):
    return [e for e in el.iter() if _local(e.tag) == name]


def _docs(s, direction, command):
    """Parsed request/response documents for one command and direction."""
    out = []
    for entry in s.log():
        message = entry.get("message") or ""
        details = entry.get("details") or ""
        if "[eas:net]" not in message or direction not in message:
            continue
        if command not in message or not details.startswith("<?xml"):
            continue
        try:
            out.append(ET.fromstring(details))
        except ET.ParseError:
            continue
    return out


def _description_line(item):
    """The item's DESCRIPTION property as one line, or None.

    RFC 5545 folds a long property across several lines, and a note with an
    ALTREP is always long enough to be folded, so the raw lines have to be
    joined before anything can be read off them.
    """
    unfolded, current = [], None
    for line in (item.get("item") or "").splitlines():
        if line.startswith((" ", "\t")):
            if current is not None:
                current += line[1:]
            continue
        if current is not None:
            unfolded.append(current)
            current = None
        if line.startswith("DESCRIPTION"):
            current = line
    if current is not None:
        unfolded.append(current)
    return unfolded[0] if unfolded else None


def _split_property(line):
    """(parameters, value) - the colon that ends the parameters is the first
    one OUTSIDE a quoted parameter value. `ALTREP="data:text/html,…"` holds
    colons of its own, so splitting on the first one reads the URI as the
    note."""
    quoted = False
    for i, ch in enumerate(line):
        if ch == '"':
            quoted = not quoted
        elif ch == ":" and not quoted:
            return line[:i], line[i + 1 :]
    return line, ""


HTML_ALTREP_PREFIX = "data:text/html,"


def _altrep(item):
    """The HTML in the DESCRIPTION's ALTREP parameter, or None.

    Stored as a `data:` URI - the shape Thunderbird's own descriptionHTML
    setter writes - so the scheme comes off and the rest is percent-decoded
    to get back the markup the user actually authored.
    """
    line = _description_line(item)
    if line is None or "ALTREP=" not in line:
        return None
    params, _ = _split_property(line)
    start = params.index("ALTREP=") + len('ALTREP="')
    uri = unquote(params[start : params.index('"', start)])
    return uri[len(HTML_ALTREP_PREFIX):] if uri.startswith(HTML_ALTREP_PREFIX) else uri


def _unescape_text(value):
    """An iCalendar TEXT value as the user sees it.

    RFC 5545 §3.3.11 escapes a line break as \\n and protects a backslash,
    comma and semicolon, so the raw property text is not the note - reading
    it without unescaping compares the wire form against the real one.
    """
    out, i = [], 0
    while i < len(value):
        ch = value[i]
        if ch == "\\" and i + 1 < len(value):
            nxt = value[i + 1]
            out.append({"n": "\n", "N": "\n"}.get(nxt, nxt))
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _description(item):
    """The DESCRIPTION value, or None when the item carries no note."""
    line = _description_line(item)
    if line is None:
        return None
    return _unescape_text(_split_property(line)[1])


def _bodies(s, marker):
    """[(Type, NativeBodyType)] the Sync responses carried for one item."""
    found = []
    for doc in _docs(s, "receive", "Sync"):
        for ad in _descendants(doc, "ApplicationData"):
            subject = _child(ad, "Subject")
            if subject is None or marker not in unquote(subject.text or ""):
                continue
            body = _child(ad, "Body")
            native = _child(ad, "NativeBodyType")
            type_el = _child(body, "Type") if body is not None else None
            found.append(
                (
                    type_el.text if type_el is not None else None,
                    native.text if native is not None else None,
                )
            )
    return found


def _preferences(s):
    """The BodyPreference Type in every Options block we sent."""
    asked = []
    for doc in _docs(s, "send", "Sync"):
        for options in _descendants(doc, "Options"):
            preference = _child(options, "BodyPreference")
            type_el = _child(preference, "Type") if preference is not None else None
            asked.append(type_el.text if type_el is not None else "ABSENT")
    return asked


def _fetches(s):
    """(requests, fetch_blocks, [Type asked]) for the ItemOperations traffic.

    Requests and Fetch blocks are counted separately because the whole point
    of the batching is that they differ: N rich notes in one window must cost
    one request carrying N <Fetch> elements, not N requests.
    """
    types = []
    requests = 0
    fetch_blocks = 0
    for doc in _docs(s, "send", "ItemOperations"):
        requests += 1
        fetch_blocks += len(_descendants(doc, "Fetch"))
        for preference in _descendants(doc, "BodyPreference"):
            type_el = _child(preference, "Type")
            types.append(type_el.text if type_el is not None else None)
    return requests, fetch_blocks, types


def _pushed_body(s, marker):
    """(Type, Data) of the Body on an outgoing Add/Change for one item."""
    for doc in _docs(s, "send", "Sync"):
        for verb in ("Add", "Change"):
            for command in _descendants(doc, verb):
                ad = _child(command, "ApplicationData")
                body = _child(ad, "Body") if ad is not None else None
                if body is None:
                    continue
                subject = _child(ad, "Subject")
                if subject is not None and marker not in unquote(subject.text or ""):
                    continue
                type_el, data = _child(body, "Type"), _child(body, "Data")
                return (
                    type_el.text if type_el is not None else None,
                    unquote(data.text or "") if data is not None else None,
                )
    return (None, None)


def _rich_ical(slug, summary, html, text, day):
    """An event whose note carries an ALTREP - what Thunderbird's own
    descriptionHTML setter writes when a user formats a note."""
    altrep = "data:text/html," + "".join(
        {"<": "%3C", ">": "%3E", "/": "%2F", " ": "%20", '"': "%22"}.get(c, c)
        for c in html
    )
    return probes.event(
        slug,
        summary=summary,
        lines=[
            f"DTSTART:2026{day}T100000Z",
            f"DTEND:2026{day}T110000Z",
            f'DESCRIPTION;ALTREP="{altrep}":{text}',
        ],
    )


def _find(s, slug):
    item = s.find("events", f"{probes.MARKER} {slug}", "event")
    harness.true(item is not None, f"event {slug!r} is missing locally")
    return item


# ── the tests ─────────────────────────────────────────────────────────────


@test("14.1", "seed three notes - plain, rich, plain - and push them")
def t_14_1(s):
    # Setup, and an assertion in its own right: the editor field decides the
    # wire format. A note carrying an ALTREP is HTML and says so; one
    # without is text. Everything after this reads the items back from the
    # server rather than trusting what we just wrote locally.
    ok("items.create", ical=probes.event(
        SLUGS[0], summary=SLUGS[0],
        lines=["DTSTART:20260921T100000Z", "DTEND:20260921T110000Z",
               f"DESCRIPTION:{PLAIN_ONE_ICAL}"]))
    ok("items.create", ical=_rich_ical(SLUGS[1], SLUGS[1], RICH_HTML, RICH_TEXT, "0922"))
    ok("items.create", ical=probes.event(
        SLUGS[2], summary=SLUGS[2],
        lines=["DTSTART:20260923T100000Z", "DTEND:20260923T110000Z",
               f"DESCRIPTION:{PLAIN_TWO}"]))
    s.sync()

    harness.eq(_pushed_body(s, SLUGS[0])[0], "1", "the plain note went out as Type 1")
    rich_type, rich_data = _pushed_body(s, SLUGS[1])
    harness.eq(rich_type, "2", "the rich note went out as Type 2")
    harness.eq(rich_data, RICH_HTML, "and carried the HTML the user authored")


@test("14.2", "the fresh pull asks for plain everywhere and fetches only the rich item")
def t_14_2(s):
    # The first pull that takes the three items back from the server rather
    # than from the local copies 14.1 created - which is the only state the
    # rest of the section is allowed to judge.
    s.mark()
    s.rebind("events")
    s.settle("events")

    asked = set(_preferences(s))
    harness.true(
        asked <= {"1"} and asked,
        f"every Options block states BodyPreference Type 1, got {sorted(asked)}",
    )

    global _RICH
    natives = {}
    for slug in SLUGS:
        seen = _bodies(s, slug)
        harness.true(seen, f"{slug} came back in the pull")
        harness.eq(seen[0][0], "1", f"{slug} arrived as Type 1, as asked")
        natives[slug] = seen[0][1]
    # Recorded, not asserted: which notes are held as HTML is the server's
    # to say. What must hold everywhere is that we believe it - one fetch
    # per note it calls rich, and none at all when it calls none.
    _RICH = [slug for slug, native in natives.items() if native == "2"]

    requests, blocks, types = _fetches(s)
    harness.eq(requests, 1 if _RICH else 0, "one ItemOperations request per window with rich notes")
    harness.eq(blocks, len(_RICH), "carrying one Fetch per rich item, and no others")
    harness.eq(types, ["2"] * len(_RICH), "and each asked for HTML")


@test("14.3", "each note lands in the right field, and an empty one invents nothing")
def t_14_3(s):
    _needs_rich()
    plain_a, rich, plain_c = (_find(s, slug) for slug in SLUGS)

    harness.eq(_description(plain_a), PLAIN_ONE, "plain note one is the value")
    harness.true(_altrep(plain_a) is None, "a plain note carries no ALTREP")
    harness.eq(_description(plain_c), PLAIN_TWO, "plain note two is the value")
    harness.true(_altrep(plain_c) is None, "the second plain note carries none either")

    harness.eq(_description(rich), RICH_TEXT, "the rich note's value is the plain text")
    # Containment, not equality: Exchange normalises HTML it is given into a
    # complete document (<html><head><meta charset></head><body>…), while
    # Kopano returns the fragment untouched. Both are the item's native form,
    # which is what the ALTREP is for; demanding the author's exact bytes
    # would pin one server's habit rather than the contract. That the
    # normalisation settles instead of accumulating is 14.4's job.
    harness.contains(
        _altrep(rich) or "", RICH_HTML, "the ALTREP carries the authored markup"
    )


@test("14.4", "an edit elsewhere leaves the note untouched - on all three")
def t_14_4(s):
    _needs_rich()
    # The trap this pins: a body we derived rather than received reads as an
    # edit to the server, so an unrelated change would rewrite the note and
    # bump the item's version. What goes back must be what we were given -
    # for a plain note as much as a rich one, and the plain ones are the
    # majority of any real folder.
    before = {}
    for slug in SLUGS:
        item = _find(s, slug)
        before[slug] = (item["id"], _description(item), _altrep(item))

    # One window, three items, one sync - so the whole batch is the unit
    # that repeats if the server overrules any of them. Idempotent: each
    # SUMMARY is replaced with a constant, so re-running changes nothing.
    def rename_all():
        s.mark()
        for slug in SLUGS:
            item = _find(s, slug)
            ical = "\r\n".join(
                f"SUMMARY:{probes.MARKER} {slug} renamed" if line.startswith("SUMMARY")
                else line
                for line in (item["item"] or "").splitlines()
            ) + "\r\n"
            ok("items.update", id=item["id"], ical=ical)
        s.sync()

    s.conflict_retry(rename_all)

    # Each note goes back in the form we hold it, carrying exactly the data
    # we were given. Omitting the body is not an option on 14.x, where an
    # absent field means "clear it".
    for slug, expected_type in ((SLUGS[0], "1"), (SLUGS[1], "2"), (SLUGS[2], "1")):
        _, value, html = before[slug]
        type_, data = _pushed_body(s, slug)
        harness.true(type_ is not None, f"{slug}: the push carried a body at all")
        harness.eq(type_, expected_type, f"{slug} still travels as Type {expected_type}")
        harness.eq(
            data,
            html if expected_type == "2" else value,
            f"{slug} carried back exactly what we were given",
        )

    s.mark()
    s.rebind("events")
    s.settle("events")

    for slug, native in ((SLUGS[0], "1"), (SLUGS[1], "2"), (SLUGS[2], "1")):
        _, value, html = before[slug]
        after = _find(s, slug)
        harness.eq(_description(after), value, f"{slug}: the note's text is unchanged")
        harness.eq(_altrep(after), html, f"{slug}: its formatting is unchanged")
        harness.eq(
            [n for _, n in _bodies(s, slug)][:1],
            [native],
            f"{slug}: the server still holds it as {native}",
        )

    # And the rich item is still the only one worth a second request.
    harness.eq(_fetches(s)[:2], (1, 1), "one request, one Fetch - the rich item alone")

    # Idempotence, which is what makes storing a server-normalised body safe
    # at all. Exchange rewrites HTML it is handed into a complete document,
    # so the note we now hold is the server's wording, not the author's. Push
    # that back and the server must recognise its own work and leave it
    # alone: if it wrapped again, the note would grow on every edit and every
    # push would read as a body change - the churn #347 caused, returning by
    # another route.
    settled = {slug: (_description(_find(s, slug)), _altrep(_find(s, slug)))
               for slug in SLUGS}
    def rename_back():
        s.mark()
        for slug in SLUGS:
            item = _find(s, slug)
            ical = "\r\n".join(
                f"SUMMARY:{probes.MARKER} {slug}" if line.startswith("SUMMARY") else line
                for line in (item["item"] or "").splitlines()
            ) + "\r\n"
            ok("items.update", id=item["id"], ical=ical)
        s.sync()

    s.conflict_retry(rename_back)
    s.rebind("events")
    s.settle("events")

    for slug, native in ((SLUGS[0], "1"), (SLUGS[1], "2"), (SLUGS[2], "1")):
        value, html = settled[slug]
        after = _find(s, slug)
        harness.eq(_description(after), value, f"{slug}: the note settled, text")
        harness.eq(_altrep(after), html, f"{slug}: the note settled, formatting")
        harness.eq(
            [n for _, n in _bodies(s, slug)][:1],
            [native],
            f"{slug}: and the server still holds it as {native}",
        )


@test("14.5", "demoting the rich note to plain sticks, on both sides")
def t_14_5(s):
    _needs_rich()
    # Idempotent, as `edit` requires: every DESCRIPTION and its folded
    # continuations are dropped first, then one constant line is put back,
    # so applying this to its own output yields the same body.
    def demote(body):
        kept = []
        dropping = False
        for line in (body or "").splitlines():
            if line.startswith((" ", "\t")):
                if dropping:
                    continue
            else:
                dropping = line.startswith("DESCRIPTION")
            if not dropping:
                kept.append(line)
        kept.insert(kept.index("END:VEVENT"), f"DESCRIPTION:{PLAIN_ONE_ICAL}")
        return "\r\n".join(kept) + "\r\n"

    s.edit(
        lambda: _find(s, SLUGS[1]),
        demote,
        missing="the rich note is not there to demote",
    )
    harness.eq(_pushed_body(s, SLUGS[1])[0], "1", "a note with no ALTREP goes out as text")

    s.mark()
    s.rebind("events")
    s.settle("events")
    harness.eq(
        [n for _, n in _bodies(s, SLUGS[1])][:1],
        ["1"],
        "the server now holds it as plain text",
    )
    harness.eq(_fetches(s)[:2], (0, 0), "nothing is rich any more, so nothing is re-fetched")
    after = _find(s, SLUGS[1])
    harness.eq(_description(after), PLAIN_ONE, "the demoted note is the plain text")
    harness.true(_altrep(after) is None, "and the old HTML is gone from the editor field")


@test("14.6", "promoting it back to HTML restores the editor's copy")
def t_14_6(s):
    _needs_rich()
    altrep = "data:text/html," + "".join(
        {"<": "%3C", ">": "%3E", "/": "%2F", " ": "%20"}.get(c, c)
        for c in PROMOTED_HTML
    )

    # Idempotent, as `edit` requires: every DESCRIPTION is dropped and one
    # constant line put back, so applying it to its own output is a no-op.
    def promote(body):
        kept = [
            line for line in (body or "").splitlines()
            if not line.startswith("DESCRIPTION")
        ]
        kept.insert(
            kept.index("END:VEVENT"), f'DESCRIPTION;ALTREP="{altrep}":{PROMOTED_TEXT}'
        )
        return "\r\n".join(kept) + "\r\n"

    s.edit(
        lambda: _find(s, SLUGS[1]),
        promote,
        missing="the plain note is not there to promote",
    )
    type_, data = _pushed_body(s, SLUGS[1])
    harness.eq(type_, "2", "the promoted note goes out as HTML")
    harness.eq(data, PROMOTED_HTML, "carrying what the user authored")

    s.mark()
    s.rebind("events")
    s.settle("events")
    harness.eq(
        [n for _, n in _bodies(s, SLUGS[1])][:1],
        ["2"],
        "the server holds it as HTML again",
    )
    harness.eq(_fetches(s)[:2], (1, 1), "so the item is fetched as HTML once more")
    after = _find(s, SLUGS[1])
    harness.eq(_description(after), PROMOTED_TEXT, "the tooltip field is the plain text")
    # Containment, for the reason 14.3 gives: Exchange returns the markup
    # inside a complete document of its own, Kopano returns it as authored.
    harness.contains(
        _altrep(after) or "", PROMOTED_HTML, "the editor field carries the markup"
    )


@test("14.7", "several rich notes cost one request, not one each")
def t_14_7(s):
    _needs_rich()
    # The reason the fetch is batched at all: everything authored in
    # OWA/Outlook is natively HTML even when its text is plain, so a first
    # pull of a real calendar could otherwise approach one request per item
    # - against servers that answer bursts with 503.
    extra = ("body-rich-x", "body-rich-y")
    for i, slug in enumerate(extra):
        ok("items.create", ical=_rich_ical(
            slug, slug, f"<b>batch {i}</b>", f"batch {i}", f"092{4 + i}"))
    s.sync()

    s.mark()
    s.rebind("events")
    s.settle("events")

    requests, blocks, types = _fetches(s)
    harness.eq(requests, 1, "one ItemOperations request for the whole window")
    harness.eq(blocks, 3, "carrying a Fetch per rich note - the two new and 14.6's promotion")
    harness.true(set(types) == {"2"}, "all asking for HTML")

    for i, slug in enumerate(extra):
        item = _find(s, slug)
        harness.eq(_description(item), f"batch {i}", f"{slug}: tooltip text")
        harness.contains(_altrep(item) or "", f"<b>batch {i}</b>", f"{slug}: editor copy")

    for slug in extra:
        ok("items.remove", id=_find(s, slug)["id"])
    s.sync()
    s.settle("events")


@test("14.8", "a downgraded note keeps its text, invents no ALTREP, and is not pushed back")
def t_14_8(s):
    # The mirror of everything above: a server that converts the HTML it is
    # given to plain text. The words must survive the conversion, the editor
    # must not be handed markup the server no longer holds, and - the part
    # that would actually cost something - the converted body must not read
    # as a local edit. That is #347's mechanism seen from the other side: we
    # stored what the server generated and pushed it straight back, which
    # bumped the item's version on every sync.
    _needs_plain_only()
    rich = _find(s, SLUGS[1])

    note = _description(rich) or ""
    harness.true("bold" in note and "italic" in note, f"the note lost its text: {note!r}")
    harness.true(
        "<b>" not in note and "</b>" not in note,
        f"the markup was stored as text rather than converted: {note!r}",
    )
    harness.true(
        _altrep(rich) is None,
        "an ALTREP was invented for a note the server holds as plain text",
    )

    # Nothing changed locally, so the next sync must be silent about it.
    s.mark()
    s.sync()
    harness.true(
        not [w for w in s.wire() if w.startswith("SEND Change")],
        "the server's own conversion was pushed back as an edit",
    )
    requests, _, _ = _fetches(s)
    harness.eq(requests, 0, "HTML was fetched for an item the server holds as plain")


@test("14.9", "clean up - the section leaves the calendar as it found it")
def t_14_8(s):
    # Sections chain, and a leftover fixture is worse than a failure: the
    # next run matches whichever copy it finds first. Removing them through
    # a sync also proves the deletions reached the server, not just the
    # local store.
    for slug in SLUGS:
        item = s.find("events", f"{probes.MARKER} {slug}", "event")
        if item is not None:
            ok("items.remove", id=item["id"])
    s.sync()
    s.settle("events")

    for slug in SLUGS:
        harness.true(
            s.find("events", f"{probes.MARKER} {slug}", "event") is None,
            f"{slug} is gone locally",
        )

    s.mark()
    s.rebind("events")
    s.settle("events")
    for slug in SLUGS:
        harness.true(
            s.find("events", f"{probes.MARKER} {slug}", "event") is None,
            f"{slug} stayed deleted - the server has it gone too",
        )
