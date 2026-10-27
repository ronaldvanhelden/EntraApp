#!/usr/bin/env bash
# PostToolUse hook: run Prettier on the file Claude just edited/wrote.
file=$(jq -r '.tool_input.file_path // empty')
[ -n "$file" ] && [ -f "$file" ] || exit 0
cd "$CLAUDE_PROJECT_DIR" || exit 0
npx --no-install prettier --write --ignore-unknown --log-level warn "$file" >&2 || true
