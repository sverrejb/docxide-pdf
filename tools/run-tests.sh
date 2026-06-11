#!/usr/bin/env bash
# Compact test runner — runs cargo test and reports only what changed.
# Produces ~4-10 lines instead of ~666 lines of raw cargo test output.
#
# Usage:
#   ./tools/run-tests.sh                        # run all tests
#   ./tools/run-tests.sh --test visual_comparison  # one test suite
#   ./tools/run-tests.sh --case case5           # one fixture
#   ./tools/run-tests.sh --libreoffice          # include LibreOffice comparison (opt-in)
#   ./tools/run-tests.sh --verbose              # show full cargo output

set -euo pipefail
cd "$(dirname "$0")/.."

VERBOSE=0
TEST_NAME=""
CASE_FILTER=""
GROUP_FILTER=""
LO_COMPARE=0

while [[ $# -gt 0 ]]; do
    case "$1" in
        --verbose|-v) VERBOSE=1; shift ;;
        --test) TEST_NAME="$2"; shift 2 ;;
        --case) CASE_FILTER="$2"; shift 2 ;;
        --group) GROUP_FILTER="$2"; shift 2 ;;
        --libreoffice|--lo) LO_COMPARE=1; shift ;;
        *) echo "Unknown arg: $1"; exit 1 ;;
    esac
done

# Build cargo command
CARGO_ARGS=("test")
if [[ -n "$TEST_NAME" ]]; then
    CARGO_ARGS+=("--test" "$TEST_NAME")
fi
CARGO_ARGS+=("--" "--nocapture")

# Set env filters
[[ -n "$CASE_FILTER" ]] && export DOCXIDE_CASE="$CASE_FILTER"
[[ -n "$GROUP_FILTER" ]] && export DOCXSIDE_GROUP="$GROUP_FILTER"
[[ "$LO_COMPARE" -eq 1 ]] && export DOCXSIDE_LO_COMPARE=1

# Verbose mode: just pass through
if [[ "$VERBOSE" -eq 1 ]]; then
    cargo "${CARGO_ARGS[@]}"
    exit $?
fi

# Clean slate for scores and hashes
rm -f tests/output/latest_scores.json
rm -f tests/output/latest_hashes.json

# Run tests, capture output
TMPFILE=$(mktemp)
trap 'rm -f "$TMPFILE"' EXIT

CARGO_EXIT=0
cargo "${CARGO_ARGS[@]}" > "$TMPFILE" 2>&1 || CARGO_EXIT=$?

# Extract compilation errors (lines with "error[E" or "error:" after Compiling)
COMPILE_ERRORS=$(grep -E '^\s*error(\[E[0-9]+\]|:)' "$TMPFILE" 2>/dev/null || true)

# Extract test result lines (skip trivial "0 passed; 0 failed" and doc tests)
TEST_RESULTS=$(grep '^test result:' "$TMPFILE" 2>/dev/null | grep -v '0 passed; 0 failed; 0 ignored' || true)

# Extract panic messages
PANICS=$(grep 'thread.*panicked' "$TMPFILE" 2>/dev/null || true)

# Print compact report
if [[ -n "$COMPILE_ERRORS" ]]; then
    echo "Compilation errors:"
    # Show errors with 2 lines of context for location info
    grep -E -B1 '^\s*error(\[E[0-9]+\]|:)' "$TMPFILE" 2>/dev/null | head -30
    echo ""
fi

if [[ -n "$PANICS" ]]; then
    echo "Panics:"
    echo "$PANICS"
    echo ""
fi

# Summarize test results into one line
if [[ -n "$TEST_RESULTS" ]]; then
    TOTAL_PASS=0
    TOTAL_FAIL=0
    while IFS= read -r line; do
        p=$(echo "$line" | grep -oE '[0-9]+ passed' | grep -oE '[0-9]+')
        f=$(echo "$line" | grep -oE '[0-9]+ failed' | grep -oE '[0-9]+')
        TOTAL_PASS=$((TOTAL_PASS + ${p:-0}))
        TOTAL_FAIL=$((TOTAL_FAIL + ${f:-0}))
    done <<< "$TEST_RESULTS"
    if [[ $TOTAL_FAIL -gt 0 ]]; then
        echo "Tests: $TOTAL_PASS passed, $TOTAL_FAIL failed"
    else
        echo "Tests: $TOTAL_PASS passed"
    fi
fi

# Run compact diff report if scores were produced
if [[ -f tests/output/latest_scores.json ]]; then
    python3 tools/compact_report.py
    REPORT_EXIT=$?
    # Use the worse exit code
    if [[ $REPORT_EXIT -ne 0 && $CARGO_EXIT -eq 0 ]]; then
        CARGO_EXIT=$REPORT_EXIT
    fi
elif [[ $CARGO_EXIT -ne 0 && -z "$COMPILE_ERRORS" ]]; then
    echo "Tests failed but no scores produced. Run with --verbose for details."
fi

exit $CARGO_EXIT
