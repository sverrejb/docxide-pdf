#!/usr/bin/env bash
# Download N new corpus cases (content-deduped, Claude-named) into tests/fixtures/new/,
# then prompt before running the Word PDF conversion automation.
# Thin wrapper over add_new_cases.py. Pass --skip-convert to download+name only.
# Usage: ./tools/download-new.sh 10
exec uv run "$(dirname "$0")/add_new_cases.py" "$@"
