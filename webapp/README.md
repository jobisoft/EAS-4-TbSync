# Outlook Web Helper

Lets local scripts (or Claude) read and drive the logged-in **Outlook on the web
calendar** (`https://outlook.live.com/`) in Firefox. It reads the DOM of a tab,
clicks, types, presses keys and loads pages of that one site. It was made to set
up calendar invitations for testing a calendar sync, which no API call of the
sync add-on can do.

It is a copy of the Amazon Web Scanner add-on, narrowed to `outlook.live.com`.
It is a test tool: `build.js` zips `src/` only, so nothing in `webapp/` ships in the
EAS-4-TbSync add-on.

| Item | Value |
| --- | --- |
| Add-on id | `outlook-web-helper@jobisoft.de` |
| Native host | `outlook_web_helper_host` |
| Config folder | `outlook-web-helper` |
| Default port | 8767 |

## How it works

1. The Firefox add-on (folder `webapp/addon/`) runs in the browser. Its content script
   reads and drives `outlook.live.com` tabs.
2. Firefox starts the Python helper app
   (`webapp/addon/native-messaging-app/outlook_web_helper_host.py`) through native
   messaging as soon as the add-on loads.
3. The helper app runs a small HTTP server on `127.0.0.1`. A script sends a
   request there, the helper forwards it to the add-on and returns the answer as
   JSON. Every request needs a secret token.

The helper app uses Python 3 only, without extra packages.

## Setup

1. Register the helper app once: `bash webapp/addon/native-messaging-app/install.sh`
   (`--dev` runs it in place from this folder). Restart Firefox.
2. Open `about:debugging#/runtime/this-firefox`, choose "Load Temporary Add-on"
   and pick `webapp/addon/manifest.json`. It has to be loaded again after each Firefox
   restart.
3. The options page of the add-on shows whether the helper app is connected, the
   address of the API and the token. Log in to Outlook in a tab.
4. Test: `curl -H "Authorization: Bearer <token>" http://127.0.0.1:8767/status`

To remove the registration run `webapp/addon/native-messaging-app/uninstall.sh`.

## HTTP API

Same as the Amazon Web Scanner: `GET /status`, `GET /tabs`, `POST /dom`,
`POST /act` (`click`, `type`, `press`), `POST /navigate`, `POST /back`,
`POST /reload`, and one more, `POST /focus`. All need
`Authorization: Bearer <adminToken>`; calls on a tab need a whole number `tabId`
from `/tabs`. `/navigate` only accepts `https://outlook.live.com/` addresses.
There is no endpoint that runs arbitrary JavaScript.

What differs from the Amazon Web Scanner:

- `POST /focus` brings the tab and its window to the front.
- `GET /tabs` also says for every tab whether the page is `visibility: "visible"`
  and whether it `hasFocus`.
- `POST /act` with `"action": "type"` also types into rich-text boxes
  (contenteditable), such as the attendee field of an event.
- `POST /act` with `"action": "click"` accepts `"slow": true`, which lets the
  press last a moment before the release.

## Focus and clicking

Outlook only reacts to a page that has focus. A hidden tab ignores clicks, and so
does a visible tab in a window that is not the focused one. `/focus` asks Firefox
to bring the window forward, but a desktop may refuse to let a program take focus
from another window (GNOME does). Check `hasFocus` in `GET /tabs` before driving
the page; if it is `false`, click once into the page content yourself and leave
the window alone while a script runs.

Clicks on an entry in the calendar are not reliable:

- The same click may open the entry's popup or start renaming the entry. A click on
  an entry that is already selected starts the rename; press Escape in its title
  field to leave it.
- `"slow": true` made the popup open in the checks this was written for. Use it
  for entries.
- The popup's buttons carry an icon character after their text ("Bearbeiten" is
  read as "Bearbeiten" plus a symbol), so match on the start of the text.
- Selectors that match many elements (`button.fui-Button...` matches dozens) click
  whatever comes first. Match on `aria-label`, and check with `POST /dom` that the
  selector matches exactly one element before clicking.
- The page is in the account's language. The labels used while this was written
  were German.

## Safety notes

- The server listens on `127.0.0.1` only. The token is the only protection
  against other programs and web pages on this computer: never print it, log it
  or commit it.
- The add-on only touches `outlook.live.com` tabs.
- The DOM of a logged-in Outlook page holds mail and calendar data. Read only what
  a task needs, keep it on this computer, and do not copy it into logs or chats.
- Intended use is calendar test events whose title starts with `TbSync test`.
  Nothing in the add-on enforces that: it is a rule for whoever drives it.

## Tests

`python3 webapp/addon/test_host.py` starts the helper app on a free port with a
temporary config folder and plays the role of Firefox. The content script and the
options page are not covered by tests.
