#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

PROFILE="${1:-team}"
case "$PROFILE" in
  michael|personal)
    export LUCAS_APP_NAME="Michael LUCAS"
    export LUCAS_APP_DISPLAY_NAME="Michael LUCAS"
    export LUCAS_SETTINGS_PATH="$PWD/lucas_settings.michael.json"
    export LUCAS_ASSIGNMENT_CONFIG_PATH="$PWD/assignment_companies.michael.json"
    DEST="${HOME}/Desktop/Michael LUCAS.app"
    ;;
  team|lucas|"")
    export LUCAS_APP_NAME="LUCAS"
    export LUCAS_APP_DISPLAY_NAME="L.U.C.A.S"
    export LUCAS_SETTINGS_PATH="$PWD/lucas_settings.json"
    export LUCAS_ASSIGNMENT_CONFIG_PATH="$PWD/assignment_companies.json"
    DEST="${HOME}/Desktop/LUCAS.app"
    ;;
  *)
    echo "Usage: ./install_macos_shortcut.sh [team|michael]"
    exit 2
    ;;
esac

./create_macos_app.sh

if [[ -e "$DEST" ]]; then
  rm -rf "$DEST"
fi

cp -R "${LUCAS_APP_NAME}.app" "$DEST"
xattr -dr com.apple.quarantine "$DEST" 2>/dev/null || true
touch "$DEST"

echo "Installed ${DEST}."
echo "If the old shortcut is pinned in the Dock, remove it and drag this Desktop copy back in."
