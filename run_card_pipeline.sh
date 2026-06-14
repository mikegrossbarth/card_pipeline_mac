#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ -x ".venv/bin/python" ]]; then
  exec ".venv/bin/python" app.py
fi

if command -v python3 >/dev/null 2>&1; then
  exec python3 app.py
fi

echo "L.U.C.A.S could not find Python."
echo "Run ./install_dependencies.sh after installing Python 3.11 or newer."
exit 1
