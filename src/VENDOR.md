# Vendored Files

This file lists files that were not created by this project and are maintained upstream elsewhere.

---

## calendar Experiments API

- **Files** : `/experiments/calendar/*`
- **Source** : https://github.com/thunderbird/webext-experiments, path
  `calendar/experiments/calendar`, from pull request #75
  (https://github.com/thunderbird/webext-experiments/pull/75, not merged yet),
  branch `compatibility-for-160-and-other-fixes` at commit
  `78b584e4250d294f1554f5b27e524a583763c57b` (26 commits on top of upstream
  `b7f7cb3e76807903a785a03784d6e7df7b213f21`). Once that pull request is
  merged, this is the merged upstream commit.
- **Updating** : never edit these copies. They are an exact 1:1 copy of the
  `webext-experiments` checkout beside the TbSync repository, written by
  `TbSync/common/vendor.sh` (`vendor.sh --check` verifies only). See
  `TbSync/common/README.md`.
- **License** : MPL 2.0

---

## ical.min.js

- **File** : `/vendor/ical.min.js`
- **Source** : https://cdn.jsdelivr.net/npm/ical.js@2.2.1/dist/ical.min.js
- **Version** : v2.2.1
- **License** : MPL 2.0 (see header of [ical.min.js](./ical.min.js))

--- 

## i18n.mjs

- **File** : `/vendor/i18n/i18n.mjs`
- **Source** : https://raw.githubusercontent.com/thunderbird/webext-support/6bbbf8ac2105d04c1b59083e8bd52e0046448ec7/modules/i18n/i18n.mjs
- **License** : MIT

---

## tbsync protocol library

- **Files** : `/vendor/tbsync/*`
- **Source** : `TbSync/common/protocol/` — the single source of truth, in the
  [TbSync repository](https://github.com/jobisoft/TbSync)
- **Updating** : never edit these copies. Change the file in `TbSync/common/protocol/`
  and run `TbSync/common/vendor.sh`, which refreshes every consumer and
  verifies each copy is byte-identical (`vendor.sh --check` verifies only).
  See `TbSync/common/README.md`.
