#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi


iterations="$1"
progress="./ralph/lisa_progress.md"

if [ ! -f "$progress" ]; then
  echo "# Progress for Lisa" > "$progress"
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================="
    echo "Iteration $i starting"
    echo "====================="

    result=$(claude --permission-mode acceptEdits -p "@${progress} \
  You are to do the following steps:
0. Read ./ralph/lisa_progress.md \
1. Select a case to improve, either a new one or continue on the last worked on. \
2. Analyze input, output. Compare results and reference. Consult specs via local-rag mcp. Plan how to improve the scores and implement the plan. \
3. If any meaningful improvement is achieved, and no substansial test regressions occur, commit the fix and write to the ./ralph/lisa_progress.md file
4. Update the progress file with processed files and details of what was done.
5. Then output <promise>COMPLETE</promise>.")

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo "Plan complete after $i iterations."
    exit 0
  fi
done