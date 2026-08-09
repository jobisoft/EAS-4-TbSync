# Vendored test harness

`bridge.py` and `harness.py` are verbatim copies of `TbSync/common/test-harness/*`.

**Do not edit them here.** Change the file in `TbSync/common/test-harness/`,
then run `TbSync/common/vendor.sh`, which refreshes every consumer and
verifies each copy is byte-identical (`--check` verifies only).

See `TbSync/common/README.md`.
