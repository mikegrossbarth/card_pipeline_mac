#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

APP_NAME="LUCAS"
APP_DIR="${APP_NAME}.app"
MACOS_DIR="${APP_DIR}/Contents/MacOS"
RESOURCES_DIR="${APP_DIR}/Contents/Resources"
PROJECT_ROOT="$(pwd)"

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

cat > "${MACOS_DIR}/LUCAS" <<LAUNCHER
#!/usr/bin/env bash
set -euo pipefail

APP_ROOT="${PROJECT_ROOT}"
LOG_FILE="\${HOME}/Desktop/LUCAS-launch.log"

{
  echo "[\$(date)] Starting L.U.C.A.S"
  echo "App root: \${APP_ROOT}"
  cd "\${APP_ROOT}"

  if [[ -x ".venv/bin/python" ]]; then
    exec ".venv/bin/python" app.py
  fi

  if command -v python3 >/dev/null 2>&1; then
    exec python3 app.py
  fi

  echo "L.U.C.A.S could not find Python."
  echo "Run ./install_dependencies.sh after installing Python 3.11 or newer."
  exit 1
} >>"\${LOG_FILE}" 2>&1
LAUNCHER

chmod +x "${MACOS_DIR}/LUCAS"

echo "Created ${APP_DIR}."
echo "You can keep ${APP_DIR} here or copy it to your Desktop; it launches this project path and local .venv."
