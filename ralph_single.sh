#!/bin/bash

claude --permission-mode acceptEdits "@PRD.md @progress.txt \
1. Read the plan.md and progress file. \
2. Find the next incomplete task and implement it. \
3. Update progress.txt with what you did. \
4. Mark task in plan.md with completed when done.
ONLY DO ONE TASK AT A TIME."