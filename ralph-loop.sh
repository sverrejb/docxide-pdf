#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi

for ((i=1; i<=$1; i++)); do
  result=claude --permission-mode acceptEdits "@plan.md @progress.txt \
1. Read the plan.md and progress file. \
2. Find the next incomplete task and implement it. \
3. Update progress.txt with what you did. \
4. Mark task in plan.md with completed when done.
ONLY DO ONE TASK AT A TIME. You are not to do more than the topmost uncompleted task. \
If the plan item list is complete, output <promise>COMPLETE</promise>."

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo "Plan complete after $i iterations."
    exit 0
  fi
done