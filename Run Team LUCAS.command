#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

export LUCAS_SETTINGS_PATH="$PWD/lucas_settings.json"
export LUCAS_ASSIGNMENT_CONFIG_PATH="$PWD/assignment_companies.json"

exec ./run_card_pipeline.sh
