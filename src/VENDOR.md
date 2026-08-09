# Vendored Files

This file lists files that were not created by this project and are maintained upstream elsewhere.

---

## calendar Experiments API

- **Files** : `/experiments/calendar/*`
- **Source** : https://download-directory.github.io/?url=https%3A%2F%2Fgithub.com%2Fthunderbird%2Fwebext-experiments%2Ftree%2Fb7f7cb3e76807903a785a03784d6e7df7b213f21%2Fcalendar%2Fexperiments%2Fcalendar
- **License** : MPL 2.0

---

## ical.min.js

- **File** : `/vendor/ical.min.js`
- **Source** : https://github.com/kewisch/ical.js/releases/download/v2.2.1/ical.min.js
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
- **Source** : `TbSync/protocol/` — the single source of truth, in the
  [TbSync repository](https://github.com/jobisoft/TbSync)
- **Updating** : never edit these copies. Change the file in `TbSync/protocol/`
  and run `TbSync/protocol/vendor.sh`, which refreshes every consumer and
  verifies each copy is byte-identical (`vendor.sh --check` verifies only).
  See `TbSync/protocol/README.md`.
