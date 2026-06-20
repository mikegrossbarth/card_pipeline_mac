#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

APP_NAME="LUCAS"
APP_DIR="${APP_NAME}.app"
MACOS_DIR="${APP_DIR}/Contents/MacOS"
RESOURCES_DIR="${APP_DIR}/Contents/Resources"
PROJECT_ROOT="$(pwd)"
BUNDLE_SUFFIX="$(printf "%s" "$PROJECT_ROOT" | shasum | awk '{print substr($1, 1, 10)}')"
BUNDLE_ID="com.cardpipeline.lucas.${BUNDLE_SUFFIX}"
PROJECT_ROOT_C="${PROJECT_ROOT//\\/\\\\}"
PROJECT_ROOT_C="${PROJECT_ROOT_C//\"/\\\"}"

mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"

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

cat > "${RESOURCES_DIR}/launcher.c" <<C_LAUNCHER
#include <errno.h>
#include <fcntl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <unistd.h>

#define APP_ROOT "${PROJECT_ROOT_C}"

static void write_log_line(int fd, const char *message) {
  if (fd < 0) return;
  dprintf(fd, "%s\\n", message);
}

int main(void) {
  const char *home = getenv("HOME");
  char log_path[4096];
  if (home && *home) {
    snprintf(log_path, sizeof(log_path), "%s/Desktop/LUCAS-launch.log", home);
  } else {
    snprintf(log_path, sizeof(log_path), "/tmp/LUCAS-launch.log");
  }

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
