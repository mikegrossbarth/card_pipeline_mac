#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

./create_macos_app.sh

DEST="${HOME}/Desktop/LUCAS.app"

if [[ -e "$DEST" ]]; then
  rm -rf "$DEST"
fi

cp -R "LUCAS.app" "$DEST"
xattr -dr com.apple.quarantine "$DEST" 2>/dev/null || true
touch "$DEST"

echo "Installed ${DEST}."
echo "If the old shortcut is pinned in the Dock, remove it and drag this Desktop copy back in."
