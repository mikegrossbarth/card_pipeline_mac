#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

log() {
  printf '\n==> %s\n' "$1"
}

warn() {
  printf '\nWARNING: %s\n' "$1" >&2
}

load_homebrew_env() {
  if command -v brew >/dev/null 2>&1; then
    return 0
  fi
  for candidate in /opt/homebrew/bin/brew /usr/local/bin/brew; do
    if [[ -x "$candidate" ]]; then
      eval "$("$candidate" shellenv)"
      return 0
    fi
  done
  return 1
}

ensure_homebrew() {
  if load_homebrew_env; then
    return 0
  fi

  log "Installing Homebrew"
  if ! command -v curl >/dev/null 2>&1; then
    echo "curl is required to install Homebrew. Install macOS Command Line Tools, then rerun this script."
    echo "  xcode-select --install"
    exit 1
  fi

  NONINTERACTIVE=1 /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
  if ! load_homebrew_env; then
    echo "Homebrew installed, but brew was not found on PATH."
    echo "Open a new Terminal window, then rerun ./install_dependencies.sh."
    exit 1
  fi
}

brew_install_formula() {
  local package="$1"
  if brew list --formula "$package" >/dev/null 2>&1; then
    return 0
  fi
  log "Installing $package"
  brew install "$package"
}

brew_install_cask_best_effort() {
  local package="$1"
  local label="$2"
  if brew list --cask "$package" >/dev/null 2>&1; then
    return 0
  fi
  log "Installing $label"
  if ! brew install --cask "$package"; then
    warn "Could not install $label automatically. Install it manually if this Mac needs it."
  fi
}

brew_install_python_tk() {
  local preferred_version=""
  if command -v python3 >/dev/null 2>&1; then
    preferred_version="$(python3 - <<'PY' 2>/dev/null || true
import sys
print(f"{sys.version_info.major}.{sys.version_info.minor}")
PY
)"
  fi

  local candidates=()
  if [[ -n "$preferred_version" ]]; then
    candidates+=("python-tk@${preferred_version}")
  fi
  candidates+=("python-tk@3.14" "python-tk@3.13" "python-tk@3.12" "python-tk@3.11")

  local package
  for package in "${candidates[@]}"; do
    if brew info --formula "$package" >/dev/null 2>&1; then
      brew_install_formula "$package"
      return 0
    fi
  done

  warn "Could not find a Homebrew python-tk formula. Tkinter may still work if Python came from python.org."
}

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

find_python_with_tkinter() {
  for candidate in python3.13 python3.12 python3.11 python3; do
    if command -v "$candidate" >/dev/null 2>&1; then
      if "$candidate" - <<'PY' >/dev/null 2>&1
import sys
import tkinter
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

log "Preparing macOS system dependencies"
ensure_homebrew
brew_install_formula python
brew_install_python_tk
brew_install_formula cliclick
brew_install_formula tesseract
brew_install_cask_best_effort google-chrome "Google Chrome"
brew_install_cask_best_effort google-drive "Google Drive for desktop"

PYTHON_CMD="$(find_python || true)"
if [[ -z "${PYTHON_CMD}" ]]; then
  echo "L.U.C.A.S needs Python 3.11 or newer, but no usable Python was found after installing Homebrew Python."
  exit 1
fi

PYTHON_WITH_TK="$(find_python_with_tkinter || true)"
if [[ -n "${PYTHON_WITH_TK}" ]]; then
  PYTHON_CMD="${PYTHON_WITH_TK}"
fi

log "Creating Python virtual environment with ${PYTHON_CMD}"
"$PYTHON_CMD" -m venv .venv
".venv/bin/python" -m pip install --upgrade pip
".venv/bin/python" -m pip install -r requirements.txt

if ! ".venv/bin/python" - <<'PY' >/dev/null 2>&1
import tkinter
PY
then
  echo "Python installed, but Tkinter is not available."
  echo
  echo "The installer already tried Homebrew python-tk. Install the python.org macOS Python build, then rerun this script:"
  echo "  https://www.python.org/downloads/macos/"
  exit 1
fi

if [[ ! -f .env ]]; then
  cp .env.example .env
fi

chmod +x run_card_pipeline.sh "Run Card Pipeline.command" create_macos_app.sh install_macos_shortcut.sh 2>/dev/null || true
chmod +x scripts/macos/cgscroll 2>/dev/null || true

echo
echo "Dependencies installed."
echo "Installed/verified: Homebrew, Python/Tkinter, cliclick, tesseract, and the CY scroll helper."
echo "Chrome and Google Drive were installed when Homebrew cask installation was available."
echo "CourtYard/CYCardScanner must still be installed, opened, logged in, and granted macOS permissions."
echo "Open .env and add GOOGLE_API_KEY, Google Sheets OAuth credentials, and your WORKING SHEETS path."
echo "Launch with: ./run_card_pipeline.sh"
