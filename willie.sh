#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi


iterations="$1"
progress="./ralph/willie_progress.md"

if [ ! -f "$progress" ]; then
  echo "# Progress for Willie}" > "$progress"
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================="
    echo "Iteration $i starting"
    echo "====================="

    result=$(claude --permission-mode acceptEdits -p "@${progress} \
  You are to do the following steps:
1. Read the progress file. \
2. Select a not-yet processed file from @src \
3. Simplify that file with code-simplifier:code-simplifier. \
4. Update the progress file with processed files and details of what was done.
ONLY DO ONE FILE AT A TIME. \
If all files have been processed: \
Report any score improvements or regressions.\
Make a commit. The message should have a 1-2 sentence summary. Then output <promise>COMPLETE</promise>.")

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo "Plan complete after $i iterations."
    exit 0
  fi
done