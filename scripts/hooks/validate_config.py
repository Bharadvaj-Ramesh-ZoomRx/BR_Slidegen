#!/usr/bin/env python3
"""Hook: auto-validate config.yaml after any write.

Called by Claude Code PostToolUse hook when Write/Edit targets a config.yaml file.
Runs config.validate() and prints errors (shown as hook feedback to Claude).

Usage in .claude/settings.json:
  {"hooks": {"PostToolUse": [{"matcher": "Edit", "command": "python3 scripts/hooks/validate_config.py"}]}}
"""

import json
import sys
import os


def main():
    # Read tool input from stdin
    try:
        tool_input = json.load(sys.stdin)
    except (json.JSONDecodeError, EOFError):
        return

    file_path = tool_input.get("file_path", "")

    # Only act on config.yaml files inside projects/
    if not file_path.endswith("config.yaml") or "/projects/" not in file_path:
        return

    if not os.path.exists(file_path):
        return

    # Add project root to path so slidegen imports work
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)

    try:
        from slidegen.pipeline.project_config import load_project_config

        config = load_project_config(file_path)
        errors = config.validate()

        if errors:
            print(f"[WARN] Config validation found {len(errors)} error(s):")
            for err in errors:
                print(f"  - {err}")
            print("\nFix these before running generate_deck() — slides may fail or render incorrectly.")
        else:
            print("Config valid")
    except Exception as e:
        print(f"[ERROR] Config validation failed: {e}")


if __name__ == "__main__":
    main()
