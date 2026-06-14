#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

APP_NAME="LUCAS"
APP_DIR="${APP_NAME}.app"
MACOS_DIR="${APP_DIR}/Contents/MacOS"
RESOURCES_DIR="${APP_DIR}/Contents/Resources"

mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"

cat > "${APP_DIR}/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleExecutable</key>
  <string>LUCAS</string>
  <key>CFBundleIdentifier</key>
  <string>com.cardpipeline.lucas</string>
  <key>CFBundleName</key>
  <string>L.U.C.A.S</string>
  <key>CFBundleDisplayName</key>
  <string>L.U.C.A.S</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleVersion</key>
  <string>1.0</string>
  <key>CFBundleShortVersionString</key>
  <string>1.0</string>
</dict>
</plist>
PLIST

cat > "${MACOS_DIR}/LUCAS" <<'LAUNCHER'
#!/usr/bin/env bash
set -euo pipefail

APP_ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$APP_ROOT"
exec ./run_card_pipeline.sh
LAUNCHER

chmod +x "${MACOS_DIR}/LUCAS"

echo "Created ${APP_DIR}."
echo "Keep ${APP_DIR} beside the project folder; it launches the local app and .venv."
