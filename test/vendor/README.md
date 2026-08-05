# Vendored from TbSync

`bridge.py` and `harness.py` are copies of `TbSync/test-harness/*`, verbatim.
The bridge is TbSync's, and every provider's suite runs the same loop, so both
files are maintained there and copied here.

**Do not edit them in this repo.** Change `TbSync/test-harness/`, then:

    cp ../TbSync/test-harness/bridge.py  test/vendor/bridge.py
    cp ../TbSync/test-harness/harness.py test/vendor/harness.py

`diff -q` against the source says whether this copy has fallen behind.

Everything provider-specific — preflight, resource selection, probes, the
tests themselves — lives one directory up.
