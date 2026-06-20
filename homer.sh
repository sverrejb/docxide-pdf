#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi

iterations="$1"
marge="./ralph/marge_progress.md"
maggie="./ralph/maggie_progress.md"
progress="./ralph/homer_progress.md"

if [ ! -f "$maggie" ]; then
  echo "Missing $maggie — run maggie.sh first to produce the case ledger."
  exit 1
fi

if [ ! -f "$progress" ]; then
  cat > "$progress" <<'EOF'
# Progress for Homer — implement features to make Maggie's cases render right

Homer reads Marge's gap detail (`ralph/marge_progress.md`) and Maggie's case
ledger (`ralph/maggie_progress.md`), then works ONE case per iteration: implement
the rendering feature the case isolates, score it against its Word `reference.pdf`,
and commit on improvement. It works only on Maggie's FIXTURE-PENDING-REF cases.

## Cases

Status values:
- **DONE** — crosses thresholds (Jaccard ≥20% / SSIM ≥75%), committed, removed from SKIPLIST.
- **IMPROVED** — score went up but still below threshold (committed, left in SKIPLIST).
- **BLOCKED** — can't fix now; note why and what was learned.
- **IN-PROGRESS** — started, not finished.

| Case | Gap / feature | Status | Score (Jaccard / SSIM) | Notes |
|---|---|---|---|---|
(none yet)
EOF
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================="
    echo "Iteration $i / $iterations — $(date '+%Y-%m-%d %H:%M:%S')"
    echo "====================="

    result=$(claude --model claude-opus-4-8 --permission-mode acceptEdits -p "@${progress} @${maggie} @${marge} \
  You implement the rendering feature behind ONE of Maggie's test cases per run. Steps:
1. Read ${progress} (your case table), ${maggie} (which case isolates which gap), and ${marge} (the gap detail: spec ref, what is missing vs. what exists, the likely owning src/ file). \
2. Select the case to work on: the first Maggie case marked 'READY' (fixture + reference exist) whose row in ${progress} is not yet DONE or BLOCKED. A case is workable as long as \`tests/fixtures/cases/<caseN>/reference.pdf\` exists on disk — verify that file directly and do NOT skip a case based on stale 'awaiting reference' / 'in SKIPLIST' wording in the ledger. (Cases still marked 'FIXTURE-PENDING-REF' genuinely have no reference yet — skip those.) Work on ONE and ONLY ONE case. If every READY case is DONE or BLOCKED, output <promise>COMPLETE</promise> and stop. \
3. Establish the baseline FIRST: run \`./tools/run-tests.sh --case <caseN>\` and record the current Jaccard/SSIM. Inspect the case with \`./tools/target/debug/docx-inspect tests/fixtures/cases/<caseN>/input.docx\` and look at its reference.pdf; consult the spec via mcp__local-rag__query_documents (\`/Users/sverrejb/specs/office.pdf\`) for the exact element semantics. \
4. Implement the feature in the owning src/ module(s) Marge identified. Read ALL the relevant XML properties up front and implement them together — do not fix one attribute at a time, and do not defer or reason yourself out of doing it now. \
5. Re-score: \`./tools/run-tests.sh --case <caseN>\` (these cases already have a Word reference.pdf and run in the full suite). Then run the FULL \`./tools/run-tests.sh\` to confirm no regressions on other cases (suite fails on >2% score drops). If you caused a regression, fix forward — do not revert. \
6. Update its row in ${progress}: gap/feature, status (DONE / IMPROVED / BLOCKED / IN-PROGRESS), before→after score, and notes (for BLOCKED, what you learned). \
7. If any meaningful improvement was achieved with no substantial regressions, commit the code change as me (not as Claude) with a 1-2 sentence summary naming the case. Do NOT run accept-baselines — baseline acceptance needs the user's approval. \
8. If and ONLY IF every READY Maggie case is now DONE or BLOCKED, output <promise>COMPLETE</promise>. Otherwise end the run without that marker.")

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo ""
    echo "########################################################"
    echo "# NO CASES LEFT — every Maggie case is DONE or BLOCKED."
    echo "# Homer is finished (after $i iterations)."
    echo "########################################################"
    exit 0
  fi
done

echo ""
echo "Ran out of iterations ($iterations) with cases still open."
echo "Re-run ./homer.sh with more iterations to continue."
