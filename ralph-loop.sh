#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi

plan="plan_cleanup.md"

# TODO: Take plan file as input param,
# Create ralph-folder with plans and logs
# Create progress file programmatically

for ((i=1; i<=$1; i++)); do
    echo "====================="
    echo "Iteration $i starting"
    echo "====================="

    result=$(claude --permission-mode acceptEdits -p "@${plan} @progress.txt \
    The plan_cleanup file has the report detailing codebase improvements that can be done. You are to follow the plan by doing these steps:
1. Read the ${plan} and progress file. \
2. Find the next incomplete task and implement it. \
3. Update progress.txt with what you did. \
4. Mark task in plan.md with completed when done.
ONLY DO ONE TASK AT A TIME. You are not to do more than the topmost uncompleted task. \
If the plan item list is complete, output <promise>COMPLETE</promise>.")

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo "Plan complete after $i iterations."
    exit 0
  fi
done