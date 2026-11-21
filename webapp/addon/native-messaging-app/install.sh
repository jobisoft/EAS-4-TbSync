#!/usr/bin/env bash
# Install the outlook_web_helper_host native messaging host for Firefox on Linux/macOS.
# Copies the runtime files into the user-level Mozilla native-messaging-hosts dir,
# so the source folder can be deleted afterwards.

# Developer mode:  bash install.sh --dev
# Does not copy the helper app. The manifest points to the file in this folder,
# so a change to it takes effect when the add-on is reloaded. Keep the folder.

set -euo pipefail

DEV=0
if [[ "${1:-}" == "--dev" ]]; then DEV=1; fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY_SRC="$SCRIPT_DIR/outlook_web_helper_host.py"
MANIFEST_SRC="$SCRIPT_DIR/outlook_web_helper_host.json"

if [[ "$OSTYPE" == darwin* ]]; then
  HOSTS_DIRS=("$HOME/Library/Application Support/Mozilla/NativeMessagingHosts")
else
  HOSTS_DIRS=("$HOME/.mozilla/native-messaging-hosts")
  # Firefox installed as a snap or flatpak keeps its own profile tree.
  if [[ -d "$HOME/snap/firefox/common" ]]; then
    HOSTS_DIRS+=("$HOME/snap/firefox/common/.mozilla/native-messaging-hosts")
  fi
  if [[ -d "$HOME/.var/app/org.mozilla.firefox" ]]; then
    HOSTS_DIRS+=("$HOME/.var/app/org.mozilla.firefox/.mozilla/native-messaging-hosts")
  fi
fi

PYTHON_PRESENT=0
if command -v python3 >/dev/null 2>&1 && python3 -c "import sys; sys.exit(0)" >/dev/null 2>&1; then
  PYTHON_PRESENT=1
fi

echo
echo "This will install the helper app for the Firefox add-on"
echo "\"Outlook Web Helper\"."
echo
echo "The following will happen:"
if [[ $PYTHON_PRESENT -eq 1 ]]; then
  echo "  - Python 3: already installed, no action needed"
else
  echo "  - Python 3: not detected (please install via your system's package manager)"
fi
for HOSTS_DIR in "${HOSTS_DIRS[@]}"; do
  if [[ $DEV -eq 1 ]]; then
    echo "  - Developer mode: use the helper app in place, no copy:"
    echo "      $PY_SRC"
  else
    echo "  - Copy the helper app into:"
    echo "      $HOSTS_DIR/outlook_web_helper_host_helper"
  fi
  echo "  - Install the manifest at:"
  echo "      $HOSTS_DIR/outlook_web_helper_host.json"
done
echo
echo "The helper app listens on 127.0.0.1 only. Requests need the secret token,"
echo "see the options page of the add-on."
echo

read -p "Proceed with installation? [y/n] " -n 1 -r
echo
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
  echo "Installation cancelled."
  read -n1 -s -r -p "Press any key to exit..."
  echo
  echo
  exit 1
fi

if [[ $PYTHON_PRESENT -eq 0 ]]; then
  echo "Please ensure a working Python environment, otherwise the helper app"
  echo "will not work."
  echo
fi

for HOSTS_DIR in "${HOSTS_DIRS[@]}"; do
  INSTALL_DIR="$HOSTS_DIR/outlook_web_helper_host_helper"
  MANIFEST_DEST="$HOSTS_DIR/outlook_web_helper_host.json"
  if [[ $DEV -eq 1 ]]; then
    PY_DEST="$PY_SRC"
    chmod +x "$PY_DEST"
  else
    PY_DEST="$INSTALL_DIR/outlook_web_helper_host.py"
    mkdir -p "$INSTALL_DIR"
    cp "$PY_SRC" "$PY_DEST"
    chmod +x "$PY_DEST"
  fi
  sed "s|/path/to/native-messaging-app/outlook_web_helper_host.py|$PY_DEST|" "$MANIFEST_SRC" > "$MANIFEST_DEST"
  echo "Helper app:       $PY_DEST"
  echo "Registered at:    $MANIFEST_DEST"
done
echo
echo "Restart Firefox to apply the changes."
if [[ $DEV -eq 0 ]]; then echo "The downloaded files can now be safely removed."; fi
echo
read -n1 -s -r -p "Press any key to exit..."
echo
echo
