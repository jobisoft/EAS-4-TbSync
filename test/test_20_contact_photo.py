"""20. The contact photo.

Split out of section 10, which ran the photo behind the field-fidelity
tests: a server that mishandles one contact field abandoned the section
before the photo was ever pushed, so on Kerio Connect - the one server we
have that v4 needed a photo-specific workaround for, `Picture` arriving as
a container rather than a value - these two had never run at all.

They share nothing with section 10 but the folder: the card is built here,
found by its own anchor, and removed at the end.

Unlike Google, ActiveSync carries the bytes inline (<Picture>, base64 in
ApplicationData), so a photo is just another field of the Sync payload -
but Exchange re-encodes it, so presence is asserted, never byte identity.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches.
NEEDS = ("contacts",)

# A 10x10 JPEG (same fixture as the google suite's photo section). Tiny on
# purpose: Exchange caps <Picture> via policy, and the test is about the
# round trip, not the payload.
JPEG_B64 = (
    "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRof"
    "Hh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwh"
    "MjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAAR"
    "CAAKAAoDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAj/xAAUEAEAAAAAAAAAAAAA"
    "AAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oA"
    "DAMBAAIRAxEAPwCdABmX/9k="
)
PHOTO_SLUG = "eas-foto"
PHOTO_ANCHOR = probes.email_of(PHOTO_SLUG)


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


@test("20.1", "photo round trip - the Picture survives a clean re-pull")
def t_20_1(s):
    s.mark()
    ok(
        "contacts.create",
        vCard=probes.card(
            PHOTO_SLUG,
            extra=(f"PHOTO;VALUE=URI:data:image/jpeg;base64,{JPEG_B64}",),
        ),
    )
    s.sync()
    harness.contains(s.wire(), "SEND Add", "the create must reach the server")
    s.rebind("contacts")
    card = s.find_card(PHOTO_ANCHOR)
    harness.true(card is not None, "the card did not survive the re-pull")
    body = _unfold(_vcard(card))
    harness.true(
        re.search(r"^PHOTO[^:\r\n]*:data:image/", body, re.M | re.I),
        f"no PHOTO came back from the server; the pulled card was:\n{body}",
    )


@test("20.2", "photo removal - stripping the PHOTO reaches the server")
def t_20_2(s):
    card = s.find_card(PHOTO_ANCHOR)
    harness.true(card is not None, "20.1 must have left the card in place")
    kept = [
        line
        for line in _unfold(_vcard(card)).splitlines()
        if not line.upper().startswith("PHOTO")
    ]
    s.mark()
    ok("contacts.update", id=card["id"], vCard="\r\n".join(kept) + "\r\n")
    s.sync()
    harness.contains(s.wire(), "SEND Change", "the edit must reach the server")
    s.rebind("contacts")
    card2 = s.find_card(PHOTO_ANCHOR)
    harness.true(card2 is not None, "the card did not survive the re-pull")
    harness.true(
        not re.search(r"^PHOTO", _unfold(_vcard(card2)), re.M | re.I),
        "the photo came back after removal - the empty <Picture> never "
        "reached the server",
    )
    ok("contacts.remove", id=card2["id"])
    s.sync()
