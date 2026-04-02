#!/usr/bin/env python3
"""Hook: auto-backup config.yaml before any edit.

Called by Claude Code PreToolUse hook when Edit targets a config.yaml file.
Reads the file_path from the CLAUDE_TOOL_INPUT env var (JSON with file_path key).

Usage in .claude/settings.json:
  {"hooks": {"PreToolUse": [{"matcher": "Edit", "command": "python3 scripts/hooks/auto_backup_config.py"}]}}
"""

import json
import os
import shutil
import sys
from datetime import datetime


def main():
    # Read tool input from stdin (Claude Code pipes JSON)
    try:
        tool_input = json.load(sys.stdin)
    except (json.JSONDecodeError, EOFError):
        return  # Not a valid hook call, skip silently

    file_path = tool_input.get("file_path", "")

    # Only act on config.yaml files inside projects/
    if not file_path.endswith("config.yaml") or "/projects/" not in file_path:
        return

    if not os.path.exists(file_path):
        return

    # Determine backup dir (sibling config_history/ folder)
    project_dir = os.path.dirname(file_path)
    history_dir = os.path.join(project_dir, "config_history")
    os.makedirs(history_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(history_dir, f"config_{timestamp}.yaml")

    # Don't backup if identical to most recent backup
    existing = sorted(
        [f for f in os.listdir(history_dir) if f.startswith("config_") and f.endswith(".yaml")]
    )
    if existing:
        latest = os.path.join(history_dir, existing[-1])
        if os.path.exists(latest):
            with open(latest, "rb") as a, open(file_path, "rb") as b:
                if a.read() == b.read():
                    return  # No changes since last backup

    shutil.copy2(file_path, backup_path)
    print(f"Config backed up → {backup_path}")


if __name__ == "__main__":
    main()
