#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi


iterations="$1"
progress="./ralph/bart_progress.md"

# ponytail: keep the mac awake for the whole run; caffeinate dies with this script ($$)
caffeinate -i -w $$ &

if [ ! -f "$progress" ]; then
  echo "# Progress for Bart" > "$progress"
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================================="
    echo "Iteration $i / $iterations — $(date '+%Y-%m-%d %H:%M:%S')"
    echo "====================================="

  result=$(claude --permission-mode acceptEdits --model opus -p "@${progress} \
  Use xhigh reasoning effort throughout. Stay on the main branch — never create or switch branches. \
  You are to do the following steps:
0. Read ./ralph/bart_progress.md \
1. Select from the top the first unsolved problem from tests/output/annotations.json. You are to work on ONE and ONLY ONE annotation. Update the progress file with the selected annotation when selected. \
2. Analyze the annotation, case input, output and reference pdf. Use the coordinates in the annotation to pinpoint the problem. Compare results and reference. Consult specs via local-rag mcp. Plan how to fix the problem and implement the plan. Do not defer work, or reason yourself out of doing it now. The only exception is systemic vertical text spacing or line length causing vertical shift. \
3. Run the tests. We must NOT have severe regressions. A small regression is only acceptable if it is plainly explained by an actual improvement that randomly happens to cause a random regression elsewhere — if you cannot explain it that way, treat it as a real regression and fix forward or back out only that change. \
4. If any meaningful improvement is achieved, and no unexplained regressions occur, mark the annotation fixed: set the fixed boolean in annotations.json for that case to true. DO NOT forget to set it to fixed. THEN commit the fix and write to ./ralph/bart_progress.md with a timestamp and reference to the problem that was fixed. \
5. If a problem is unfixable, make a note of what you have learned in the progress file. \
6. Then output <promise>COMPLETE</promise>.")

  echo "$result"

done
