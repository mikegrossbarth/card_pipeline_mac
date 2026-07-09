#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

APP_NAME="${LUCAS_APP_NAME:-LUCAS}"
APP_DISPLAY_NAME="${LUCAS_APP_DISPLAY_NAME:-L.U.C.A.S}"
APP_DIR="${APP_NAME}.app"
MACOS_DIR="${APP_DIR}/Contents/MacOS"
RESOURCES_DIR="${APP_DIR}/Contents/Resources"
PROJECT_ROOT="$(pwd)"
APP_ICON_SOURCE="${LUCAS_APP_ICON_PATH:-}"
APP_ICON_FILE=""
APP_NAME_LC="$(printf "%s %s" "$APP_NAME" "$APP_DISPLAY_NAME" | tr '[:upper:]' '[:lower:]')"
if [[ -z "$APP_ICON_SOURCE" && ( "$APP_NAME_LC" == *"michael"* || "$APP_NAME_LC" == *"personal"* ) && -f "assets/mikeys_cards_logo.icns" ]]; then
  APP_ICON_SOURCE="assets/mikeys_cards_logo.icns"
fi
if [[ -n "$APP_ICON_SOURCE" && -f "$APP_ICON_SOURCE" ]]; then
  APP_ICON_FILE="LUCAS.icns"
fi
BUNDLE_SUFFIX="$(printf "%s:%s" "$PROJECT_ROOT" "$APP_NAME" | shasum | awk '{print substr($1, 1, 10)}')"
BUNDLE_ID="com.cardpipeline.lucas.${BUNDLE_SUFFIX}"
PROJECT_ROOT_C="${PROJECT_ROOT//\\/\\\\}"
PROJECT_ROOT_C="${PROJECT_ROOT_C//\"/\\\"}"
LUCAS_SETTINGS_PATH_C="${LUCAS_SETTINGS_PATH:-}"
LUCAS_SETTINGS_PATH_C="${LUCAS_SETTINGS_PATH_C//\\/\\\\}"
LUCAS_SETTINGS_PATH_C="${LUCAS_SETTINGS_PATH_C//\"/\\\"}"
LUCAS_ASSIGNMENT_CONFIG_PATH_C="${LUCAS_ASSIGNMENT_CONFIG_PATH:-}"
LUCAS_ASSIGNMENT_CONFIG_PATH_C="${LUCAS_ASSIGNMENT_CONFIG_PATH_C//\\/\\\\}"
LUCAS_ASSIGNMENT_CONFIG_PATH_C="${LUCAS_ASSIGNMENT_CONFIG_PATH_C//\"/\\\"}"
LUCAS_PIPELINE_DIR_C="${LUCAS_PIPELINE_DIR:-}"
LUCAS_PIPELINE_DIR_C="${LUCAS_PIPELINE_DIR_C//\\/\\\\}"
LUCAS_PIPELINE_DIR_C="${LUCAS_PIPELINE_DIR_C//\"/\\\"}"

mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"
if [[ -n "$APP_ICON_FILE" ]]; then
  cp "$APP_ICON_SOURCE" "${RESOURCES_DIR}/${APP_ICON_FILE}"
fi

cat > "${APP_DIR}/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleExecutable</key>
  <string>LUCAS</string>
  <key>CFBundleIdentifier</key>
  <string>${BUNDLE_ID}</string>
  <key>CFBundleName</key>
  <string>${APP_DISPLAY_NAME}</string>
  <key>CFBundleDisplayName</key>
  <string>${APP_DISPLAY_NAME}</string>
PLIST
if [[ -n "$APP_ICON_FILE" ]]; then
  cat >> "${APP_DIR}/Contents/Info.plist" <<PLIST
  <key>CFBundleIconFile</key>
  <string>${APP_ICON_FILE}</string>
PLIST
fi
cat >> "${APP_DIR}/Contents/Info.plist" <<PLIST
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleVersion</key>
  <string>1.0</string>
  <key>CFBundleShortVersionString</key>
  <string>1.0</string>
</dict>
</plist>
PLIST

cat > "${RESOURCES_DIR}/launcher.c" <<C_LAUNCHER
#include <errno.h>
#include <fcntl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/stat.h>
#include <time.h>
#include <unistd.h>

#define APP_ROOT "${PROJECT_ROOT_C}"
#define PROFILE_SETTINGS_PATH "${LUCAS_SETTINGS_PATH_C}"
#define PROFILE_ASSIGNMENT_CONFIG_PATH "${LUCAS_ASSIGNMENT_CONFIG_PATH_C}"
#define PROFILE_PIPELINE_DIR "${LUCAS_PIPELINE_DIR_C}"

static void write_log_line(int fd, const char *message) {
  if (fd < 0) return;
  dprintf(fd, "%s\\n", message);
}

int main(void) {
  char log_path[4096];
  char work_dir[4096];
  snprintf(work_dir, sizeof(work_dir), "%s/work", APP_ROOT);
  mkdir(work_dir, 0755);
  snprintf(log_path, sizeof(log_path), "%s/LUCAS-launch.log", work_dir);

  int fd = open(log_path, O_WRONLY | O_CREAT | O_APPEND, 0644);
  if (fd >= 0) {
    time_t now = time(NULL);
    char stamp[128];
    struct tm local_time;
    localtime_r(&now, &local_time);
    strftime(stamp, sizeof(stamp), "[%a %b %d %H:%M:%S %Z %Y] Starting L.U.C.A.S", &local_time);
    write_log_line(fd, stamp);
    dprintf(fd, "App root: %s\\n", APP_ROOT);
    dup2(fd, STDOUT_FILENO);
    dup2(fd, STDERR_FILENO);
  }

  if (chdir(APP_ROOT) != 0) {
    dprintf(fd, "Could not open app root: %s\\n", strerror(errno));
    return 1;
  }

  if (PROFILE_SETTINGS_PATH[0] != '\\0') {
    setenv("LUCAS_SETTINGS_PATH", PROFILE_SETTINGS_PATH, 1);
    dprintf(fd, "Settings path: %s\\n", PROFILE_SETTINGS_PATH);
  }
  if (PROFILE_ASSIGNMENT_CONFIG_PATH[0] != '\\0') {
    setenv("LUCAS_ASSIGNMENT_CONFIG_PATH", PROFILE_ASSIGNMENT_CONFIG_PATH, 1);
    dprintf(fd, "Assignment config path: %s\\n", PROFILE_ASSIGNMENT_CONFIG_PATH);
  }
  if (PROFILE_PIPELINE_DIR[0] != '\\0') {
    setenv("LUCAS_PIPELINE_DIR", PROFILE_PIPELINE_DIR, 1);
    dprintf(fd, "Pipeline dir: %s\\n", PROFILE_PIPELINE_DIR);
  }

  if (access(".venv/bin/python", X_OK) == 0) {
    execl(".venv/bin/python", ".venv/bin/python", "app.py", (char *)NULL);
    dprintf(fd, "Could not launch .venv/bin/python: %s\\n", strerror(errno));
    return 1;
  }

  execlp("python3", "python3", "app.py", (char *)NULL);
  dprintf(fd, "L.U.C.A.S could not find Python. Run ./install_dependencies.sh after installing Python 3.11 or newer.\\n");
  return 1;
}
C_LAUNCHER

if command -v cc >/dev/null 2>&1; then
  cc "${RESOURCES_DIR}/launcher.c" -o "${MACOS_DIR}/LUCAS"
else
  cat > "${MACOS_DIR}/LUCAS" <<LAUNCHER
#!/usr/bin/env bash
set -euo pipefail

APP_ROOT="${PROJECT_ROOT}"
PROFILE_SETTINGS_PATH="${LUCAS_SETTINGS_PATH:-}"
PROFILE_ASSIGNMENT_CONFIG_PATH="${LUCAS_ASSIGNMENT_CONFIG_PATH:-}"
PROFILE_PIPELINE_DIR="${LUCAS_PIPELINE_DIR:-}"
LOG_FILE="\${APP_ROOT}/work/LUCAS-launch.log"
mkdir -p "\${APP_ROOT}/work"

{
  echo "[\$(date)] Starting L.U.C.A.S"
  echo "App root: \${APP_ROOT}"
  cd "\${APP_ROOT}"
  if [[ -n "\${PROFILE_SETTINGS_PATH}" ]]; then
    export LUCAS_SETTINGS_PATH="\${PROFILE_SETTINGS_PATH}"
    echo "Settings path: \${PROFILE_SETTINGS_PATH}"
  fi
  if [[ -n "\${PROFILE_ASSIGNMENT_CONFIG_PATH}" ]]; then
    export LUCAS_ASSIGNMENT_CONFIG_PATH="\${PROFILE_ASSIGNMENT_CONFIG_PATH}"
    echo "Assignment config path: \${PROFILE_ASSIGNMENT_CONFIG_PATH}"
  fi
  if [[ -n "\${PROFILE_PIPELINE_DIR}" ]]; then
    export LUCAS_PIPELINE_DIR="\${PROFILE_PIPELINE_DIR}"
    echo "Pipeline dir: \${PROFILE_PIPELINE_DIR}"
  fi

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
fi

chmod +x "${MACOS_DIR}/LUCAS"

echo "Created ${APP_DIR}."
echo "Bundle ID: ${BUNDLE_ID}"
if file "${MACOS_DIR}/LUCAS" | grep -q "Mach-O"; then
  echo "Launcher type: native"
else
  echo "Launcher type: shell fallback"
fi
echo "You can keep ${APP_DIR} here or copy it to your Desktop; it launches this project path and local .venv."
