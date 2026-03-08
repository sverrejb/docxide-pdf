#!/bin/bash
STATE="$HOME/.claude/.claude.json"

# Restore ~/.claude.json from the persisted volume
if [ -f "$STATE" ]; then
    cp "$STATE" "$HOME/.claude.json"
fi

# Run the command (no exec — bash must stay alive for the trap)
"$@"
EXIT_CODE=$?

# Persist ~/.claude.json back into the volume
cp "$HOME/.claude.json" "$STATE" 2>/dev/null
exit $EXIT_CODE
