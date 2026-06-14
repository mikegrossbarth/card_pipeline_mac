#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

find_python() {
  for candidate in python3.13 python3.12 python3.11 python3; do
    if command -v "$candidate" >/dev/null 2>&1; then
      if "$candidate" - <<'PY' >/dev/null 2>&1
import sys
raise SystemExit(0 if sys.version_info >= (3, 11) else 1)
PY
      then
        printf '%s\n' "$candidate"
        return 0
      fi
    fi
  done
  return 1
}

PYTHON_CMD="$(find_python || true)"
if [[ -z "${PYTHON_CMD}" ]]; then
  echo "L.U.C.A.S needs Python 3.11 or newer."
  echo
  echo "Install Python from https://www.python.org/downloads/macos/ or with Homebrew:"
  echo "  brew install python"
  exit 1
fi

"$PYTHON_CMD" -m venv .venv
".venv/bin/python" -m pip install --upgrade pip
".venv/bin/python" -m pip install -r requirements.txt

if ! ".venv/bin/python" - <<'PY' >/dev/null 2>&1
import tkinter
PY
then
  echo "Python installed, but Tkinter is not available."
  echo
  echo "Use the python.org macOS installer, or install tkinter support for your Python."
  echo "With Homebrew Python this is usually:"
  echo "  brew install python-tk"
  exit 1
fi

if [[ ! -f .env ]]; then
  cp .env.example .env
fi

chmod +x run_card_pipeline.sh "Run Card Pipeline.command" create_macos_app.sh 2>/dev/null || true

echo
echo "Dependencies installed."
echo "Open .env and add GOOGLE_API_KEY, Google Sheets OAuth credentials, and your WORKING SHEETS path."
echo "Launch with: ./run_card_pipeline.sh"
