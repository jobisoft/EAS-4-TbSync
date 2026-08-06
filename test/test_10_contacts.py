"""10. Contact round trip.

The suite covered events and tasks in depth and contacts not at all - the
contact codec (names, three emails, typed phones, two addresses, birthday,
nickname, categories, notes) had no automated coverage. These sections drive
the same create/modify/delete shape as section 2 does for events, plus a
field-fidelity check across a clean re-pull, which is the only view of what
the server actually stored.

Anchoring rule, learned the hard way in the google suite: match a contact by
its probe email, never by FN (rebuilt from name parts) and never by UID (a
clean pull mints fresh ones).
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches.
NEEDS = ("contacts",)

SLUG = "eas-contact"
ANCHOR = probes.email_of(SLUG)

# Fields that must survive a full round trip to the server and back. Each is
# (label, regex against the unfolded vCard). Kept to what the EAS contact
# codec maps on both directions.
ROUND_TRIP_FIELDS = [
    ("given name", r"^N:[^;\r\n]*;PROBE"),
    ("nickname", r"^NICKNAME:Probey"),
    ("organization", r"^ORG:Beispiel GmbH"),
    # 10.2 edits the title before this runs, so the *edited* value is what
    # must survive - which also proves the Change round-tripped.
    ("title (as edited in 10.2)", r"^TITLE:Cheftester"),
    ("work email", rf"^EMAIL[^:\r\n]*:{re.escape(ANCHOR)}"),
    ("home email", rf"^EMAIL[^:\r\n]*:{re.escape(SLUG)}-home@probe\.invalid"),
    ("third email", rf"^EMAIL[^:\r\n]*:{re.escape(SLUG)}-third@probe\.invalid"),
    ("work phone", r"^TEL[^:\r\n]*:\+49 228 1234567"),
    ("cell phone", r"^TEL[^:\r\n]*:\+49 170 7654321"),
    ("work street", r"^ADR[^:\r\n]*:[^\r\n]*Musterweg 4"),
    ("home street", r"^ADR[^:\r\n]*:[^\r\n]*Heimstr\. 2"),
    ("birthday", r"^BDAY[^:\r\n]*:1980-?02-?29"),
    ("note umlauts", r"^NOTE[^:\r\n]*:[^\r\n]*(äöü|\\u00e4)"),
]


def _unfold(text):
    out = []
    for raw in (text or "").splitlines():
        if raw.startswith((" ", "\t")) and out:
            out[-1] += raw[1:]
        else:
            out.append(raw)
    return "\n".join(out)


def _vcard(card):
    return (card.get("properties") or {}).get("vCard") or ""


def _probe_card(s):
    return s.find_card(ANCHOR)


def _server_id(card):
    for line in _unfold(_vcard(card)).splitlines():
        if line.upper().startswith("X-EAS-SERVERID"):
            return line.split(":", 1)[1]
    return None


@test("10.1", "contacts.create, sync - one <Add>; the card is stamped")
def t_10_1(s):
    before = len(s.cards())
    s.mark()
    ok("contacts.create", vCard=probes.card(SLUG))
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the create must reach the server")

    card = _probe_card(s)
    harness.true(card is not None, "the card vanished after the push")
    harness.true(
        _server_id(card) is not None,
        "no X-EAS-SERVERID - the card was saved locally but the server's "
        "identity for it never came back",
    )
    harness.eq(len(s.cards()), before + 1, "card count")
    harness.eq(s.changelog("contacts"), [], "changelog drained")
    harness.eq(s.status("contacts"), "success", "folder status")


@test("10.2", "contacts.update, sync - one <Change>; the edit sticks")
def t_10_2(s):
    card = _probe_card(s)
    harness.true(card is not None, "10.1 must have left a card to modify")
    s.mark()
    # Edit the read-back vCard, never a rebuilt fixture: replacing the body
    # wholesale is a different scenario (the EAS event suite keeps one on
    # purpose in 5.3), and here it would silently drop the stamp.
    edited = _vcard(card).replace("TITLE:Protokolltester", "TITLE:Cheftester")
    ok("contacts.update", id=card["id"], vCard=edited)
    s.sync()
    harness.contains(s.wire(), "SEND Change", "the edit must reach the server")
    harness.contains(_vcard(_probe_card(s)), "Cheftester", "the local card")
    harness.eq(s.changelog("contacts"), [], "changelog drained")


@test("10.3", "clean re-pull - every mapped field comes back from the server")
def t_10_3(s):
    s.rebind("contacts")
    card = _probe_card(s)
    harness.true(
        card is not None,
        "the card did not survive a clean re-pull, so it never truly "
        "reached the server",
    )
    body = _unfold(_vcard(card))
    missing = [
        label
        for label, pattern in ROUND_TRIP_FIELDS
        if not re.search(pattern, body, re.M | re.I)
    ]
    harness.eq(
        missing,
        [],
        f"fields lost in the server round trip; the pulled card was:\n{body}",
    )



@test("10.4", "contacts.remove, sync - one <Delete>; gone and staying gone")
def t_10_4(s):
    card = _probe_card(s)
    harness.true(card is not None, "10.3 must have left the card in place")
    before = len(s.cards())
    s.mark()
    ok("contacts.remove", id=card["id"])
    s.sync()
    harness.contains(s.wire(), "SEND Delete", "the delete must reach the server")
    harness.true(_probe_card(s) is None, "the card is gone locally")
    harness.eq(len(s.cards()), before - 1, "card count")
    # The echo must not re-create it.
    s.sync()
    harness.true(_probe_card(s) is None, "the echo re-created the card")
    harness.eq(s.changelog("contacts"), [], "changelog drained")
