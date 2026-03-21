#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi


iterations="$1"
progress="./ralph/annotations_progress.md"

if [ ! -f "$progress" ]; then
  echo "# Progress for Lisa" > "$progress"
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================================="
    echo "Iteration $i / $iterations — $(date '+%Y-%m-%d %H:%M:%S')"
    echo "====================================="

result=$(claude --permission-mode acceptEdits -p "@${progress} \
  You are to do the following steps:
0. Read ./ralph/lisa_progress.md \
1. Select from the top the first unsolved problem from tests/output/annotations.json. \
2. Analyze the annotation, case input, output and reference pdf. Use the coordinates in the annotation to pinpoint the problem. Compare results and reference. Consult specs via local-rag mcp. Plan how to fix the problem and implement the plan. \
3. If any meaningful improvement is achieved, and no substansial test regressions occur, commit the fix and write to the ./ralph/annotations_progress.md file, with an timestamp and reference to the problem that was fixed.
4. If you believe the problem is fixed, set the fixed boolean in annotations.json for that case to true.
5. Then output <promise>COMPLETE</promise>.")

  echo "$result"

done