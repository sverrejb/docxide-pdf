#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi

iterations="$1"
marge="./ralph/marge_progress.md"
progress="./ralph/maggie_progress.md"

if [ ! -f "$marge" ]; then
  echo "Missing $marge — run marge.sh first to produce the gap audit."
  exit 1
fi

if [ ! -f "$progress" ]; then
  cat > "$progress" <<'EOF'
# Progress for Maggie — fixtures from Marge's coverage audit

Maggie reads Marge's gap audit (`ralph/marge_progress.md`), picks the next
high-value gap not yet handled here, checks whether any existing fixture already
exercises it, and — only if not — creates a new test fixture. It NEVER produces
`reference.pdf`; that is the user's MS Word export job. New fixtures are parked
in `tests/fixtures/SKIPLIST` until the user supplies the reference.

## Ledger

One row per gap Maggie has processed. Status values:
- **COVERED** — an existing fixture already exercises this element; no new fixture.
- **FIXTURE-PENDING-REF** — new fixture created, but its `reference.pdf` does not exist yet (not ready for Homer).
- **READY** — new fixture created AND its `reference.pdf` exists; ready for Homer to implement. (Flip a row from FIXTURE-PENDING-REF to READY once the user has converted its reference.)
- **SKIPPED** — no print impact / not worth a fixture (say why).

| Gap (spec ref + element) | Priority | Status | Fixture / note |
|---|---|---|---|
(none yet)
EOF
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================="
    echo "Iteration $i / $iterations — $(date '+%Y-%m-%d %H:%M:%S')"
    echo "====================="

    result=$(claude --model claude-opus-4-8 --permission-mode acceptEdits -p "@${marge} @${progress} \
  You turn Marge's OOXML coverage gaps into test fixtures, ONE gap per run. Steps:
1. Read ${progress} (your ledger) and the 'Coverage gaps found' section of ${marge}. \
2. Pick the NEXT gap to handle: the highest-priority gap (🔴/HIGH first, then MEDIUM, then LOW) whose spec ref is NOT already a row in your ledger. Work on ONE and ONLY ONE gap. If every gap in ${marge} already has a ledger row, output <promise>COMPLETE</promise> and stop. \
3. Check whether an existing fixture already exercises it: search all fixtures for the relevant XML element/attribute with \`./tools/target/debug/analyze-fixtures --grep \"w:elementName\"\` (build tools first if needed: cd tools && cargo build). If a fixture already contains it, add a COVERED ledger row naming that fixture and stop — do NOT create a new fixture. \
4. If NOT exercised, create a minimal new fixture that isolates this feature: \
   - pick the next free \`tests/fixtures/cases/caseN\` (N = max existing + 1); \
   - write \`generate.py\` using python-docx + ZIP post-processing (follow the pattern in tests/fixtures/cases/case43/generate.py — inject the raw XML the feature needs); the doc should be SMALL and focused on just this element; \
   - run it to produce \`input.docx\` (use the project's python env, e.g. \`uv run python tests/fixtures/cases/caseN/generate.py\`) and confirm the docx opens (\`./tools/target/debug/docx-inspect tests/fixtures/cases/caseN/input.docx\`); \
   - add the new case name to \`tests/fixtures/SKIPLIST\` under a comment '# Awaiting reference.pdf (maggie)' so the test suite stays green until the user exports the reference. \
5. Do NOT create reference.pdf (user's job) and do NOT commit anything. \
6. Add one ledger row in ${progress} for the gap: spec ref + element, priority, status (COVERED / FIXTURE-PENDING-REF / SKIPPED), and the fixture path or note. This file plus the new fixture are the only deliverables. \
7. If and ONLY IF every gap in ${marge} now has a ledger row, output <promise>COMPLETE</promise>. Otherwise end the run without that marker.")

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo ""
    echo "########################################################"
    echo "# NOTHING LEFT TO PICK — every gap in $marge"
    echo "# now has a ledger row. Maggie is done (after $i iterations)."
    echo "########################################################"
    exit 0
  fi
done

echo ""
echo "Ran out of iterations ($iterations) with gaps still un-triaged."
echo "Re-run ./maggie.sh with more iterations to continue."
