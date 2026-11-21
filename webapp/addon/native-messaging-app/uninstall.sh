#!/usr/bin/env bash
# Uninstall the outlook_web_helper_host native messaging host for Firefox on Linux/macOS.

set -euo pipefail

if [[ "$OSTYPE" == darwin* ]]; then
  HOSTS_DIRS=("$HOME/Library/Application Support/Mozilla/NativeMessagingHosts")
  CONFIG_DIR="~/Library/Application Support/outlook-web-helper"
else
  HOSTS_DIRS=("$HOME/.mozilla/native-messaging-hosts"
              "$HOME/snap/firefox/common/.mozilla/native-messaging-hosts"
              "$HOME/.var/app/org.mozilla.firefox/.mozilla/native-messaging-hosts")
  CONFIG_DIR="~/.config/outlook-web-helper"
fi

echo
echo "This will uninstall the helper app of the Firefox add-on"
echo "\"Outlook Web Helper\"."
echo
echo "The configuration (port and token) stays in:"
echo "      $CONFIG_DIR"
echo

read -p "Proceed with uninstallation? [y/n] " -n 1 -r
echo
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
  echo "Uninstallation cancelled."
  read -n1 -s -r -p "Press any key to exit..."
  echo
  echo
  exit 1
fi

for HOSTS_DIR in "${HOSTS_DIRS[@]}"; do
  MANIFEST_DEST="$HOSTS_DIR/outlook_web_helper_host.json"
  INSTALL_DIR="$HOSTS_DIR/outlook_web_helper_host_helper"
  if [[ -f "$MANIFEST_DEST" ]]; then
    rm "$MANIFEST_DEST"
    echo "Registration file removed: $MANIFEST_DEST"
  fi
  if [[ -d "$INSTALL_DIR" ]]; then
    rm -r "$INSTALL_DIR"
    echo "Helper app removed:        $INSTALL_DIR"
  fi
done
echo
echo "Restart Firefox to apply the changes."
echo
read -n1 -s -r -p "Press any key to exit..."
echo
echo
