#!/bin/bash
set -e

if [ -z "$1" ] || [ -z "$2" ]; then
  echo "Usage: $0 <plan_file> <iterations>"
  exit 1
fi

plan="$1"
iterations="$2"
progress="${plan%.md}_progress.md"

if [ ! -f "$progress" ]; then
  echo "# Progress for ${plan}" > "$progress"
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================="
    echo "Iteration $i starting"
    echo "====================="

    result=$(claude --permission-mode acceptEdits -p "@${plan} @${progress} \
    The plan file has instructions for a feature to be implemented. You are to follow the plan by doing these steps:
1. Read the ${plan} and progress file. \
2. Find the next incomplete task and implement it. \
3. Update ${progress} with what you did. \
4. Mark task in ${plan} with completed when done.
ONLY DO ONE TASK AT A TIME. You are not to do more than the topmost uncompleted task. \
If the plan item list is complete, output <promise>COMPLETE</promise>.")

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo "Plan complete after $i iterations."
    exit 0
  fi
done