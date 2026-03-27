"""
stage_runner.py — Claude API calls for each pipeline stage.

Provides:
  run_stage_stream()    — initial run, yields text chunks
  refine_stage_stream() — refinement turn, yields text chunks
  confirm_stage()       — write confirmed output to disk + checkpoint
"""

from __future__ import annotations

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Generator

import anthropic

from web.skill_loader import build_stage_prompt, resolve_output_path

# ── Model ──────────────────────────────────────────────────────────────────────

DEFAULT_MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 16000  # enough for full hypothesis banks and slide plans


# ── Streaming helpers ─────────────────────────────────────────────────────────

def run_stage_stream(
    stage_num: int,
    project_dir: Path,
    wave: str,
    api_key: str,
    message_history: list[dict],
) -> Generator[str, None, None]:
    """
    Run a pipeline stage and stream Claude's response.

    Yields text chunks as they arrive. The caller is responsible for
    accumulating chunks and storing the final result.

    message_history should be the current history for this stage
    (empty list for first run).
    """
    client = anthropic.Anthropic(api_key=api_key)
    system, user_message = build_stage_prompt(stage_num, project_dir, wave)

    # Build messages: history + new user turn
    messages = list(message_history) + [{"role": "user", "content": user_message}]

    with client.messages.stream(
        model=DEFAULT_MODEL,
        max_tokens=MAX_TOKENS,
        system=system,
        messages=messages,
    ) as stream:
        for chunk in stream.text_stream:
            yield chunk


def refine_stage_stream(
    instruction: str,
    stage_num: int,
    project_dir: Path,
    wave: str,
    api_key: str,
    message_history: list[dict],
) -> Generator[str, None, None]:
    """
    Send a refinement instruction as a follow-up turn.

    message_history must include the prior assistant turn (the output to refine).
    Yields text chunks from the revised output.
    """
    client = anthropic.Anthropic(api_key=api_key)
    system, _ = build_stage_prompt(stage_num, project_dir, wave)

    # Append the refinement instruction to existing history
    messages = list(message_history) + [
        {"role": "user", "content": instruction}
    ]

    with client.messages.stream(
        model=DEFAULT_MODEL,
        max_tokens=MAX_TOKENS,
        system=system,
        messages=messages,
    ) as stream:
        for chunk in stream.text_stream:
            yield chunk


def confirm_stage(
    stage_num: int,
    content: str,
    project_dir: Path,
    wave: str,
) -> Path:
    """
    Write confirmed stage output to its canonical path and create a checkpoint.

    Returns the path where the file was written.
    """
    output_path = resolve_output_path(stage_num, project_dir, wave)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Write canonical output
    output_path.write_text(content, encoding="utf-8")

    # Write checkpoint copy with timestamp
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    checkpoint_dir = project_dir / f"context/{wave}/checkpoints"
    checkpoint_dir.mkdir(parents=True, exist_ok=True)
    checkpoint_name = f"{output_path.stem}_checkpoint_{ts}{output_path.suffix}"
    shutil.copy2(output_path, checkpoint_dir / checkpoint_name)

    return output_path


def load_existing_output(stage_num: int, project_dir: Path, wave: str) -> str | None:
    """
    Load existing output for a stage if it already exists on disk.
    Returns None if the file doesn't exist.
    """
    output_path = resolve_output_path(stage_num, project_dir, wave)
    if output_path.exists():
        return output_path.read_text(encoding="utf-8")
    return None


def list_checkpoints(stage_num: int, project_dir: Path, wave: str) -> list[Path]:
    """Return list of checkpoint files for a stage, newest first."""
    output_path = resolve_output_path(stage_num, project_dir, wave)
    checkpoint_dir = project_dir / f"context/{wave}/checkpoints"
    if not checkpoint_dir.exists():
        return []
    pattern = f"{output_path.stem}_checkpoint_*{output_path.suffix}"
    checkpoints = sorted(checkpoint_dir.glob(pattern), reverse=True)
    return checkpoints


def load_checkpoint(checkpoint_path: Path) -> str:
    """Load content from a checkpoint file."""
    return checkpoint_path.read_text(encoding="utf-8")
