"""
project_manager.py — Project and wave discovery, stage completion checks.

Scans the projects/ directory for available projects and waves.
Checks which pipeline stages are already complete (output files exist).
"""

from __future__ import annotations

from pathlib import Path

import yaml

from web.skill_loader import STAGE_DEFS, resolve_output_path

REPO_ROOT = Path(__file__).parent.parent
PROJECTS_DIR = REPO_ROOT / "projects"


# ── Project discovery ─────────────────────────────────────────────────────────

def list_projects() -> list[str]:
    """Return list of project folder names under projects/."""
    if not PROJECTS_DIR.exists():
        return []
    return sorted(
        d.name
        for d in PROJECTS_DIR.iterdir()
        if d.is_dir() and not d.name.startswith(".")
    )


def get_project_dir(project_name: str) -> Path:
    return PROJECTS_DIR / project_name


def list_waves(project_name: str) -> list[str]:
    """Return wave names from input/wave/ for a project, newest first."""
    wave_dir = PROJECTS_DIR / project_name / "input" / "wave"
    if not wave_dir.exists():
        return []
    return sorted(
        [d.name for d in wave_dir.iterdir() if d.is_dir()],
        reverse=True,
    )


def get_project_meta(project_name: str) -> dict:
    """Load top-level metadata from config.yaml (project name, client, wave)."""
    config_path = PROJECTS_DIR / project_name / "config.yaml"
    if not config_path.exists():
        return {"name": project_name, "client": "", "wave": ""}
    try:
        with open(config_path, encoding="utf-8") as f:
            cfg = yaml.safe_load(f)
        proj = cfg.get("project", {})
        return {
            "name": proj.get("name", project_name),
            "client": proj.get("client", ""),
            "wave": proj.get("wave", ""),
        }
    except Exception:
        return {"name": project_name, "client": "", "wave": ""}


# ── Stage status ──────────────────────────────────────────────────────────────

def get_stage_status(project_name: str, wave: str) -> dict[int, str]:
    """
    Return completion status for all pipeline stages.

    Status values:
      "done"    — output file exists on disk
      "pending" — output file does not exist
    """
    project_dir = PROJECTS_DIR / project_name
    status = {}

    # Stage 0: source_data.json exists
    json_path = project_dir / f"context/{wave}/source_data.json"
    status[0] = "done" if json_path.exists() else "pending"

    # Stages 1–4: check output files
    for stage_num in [1, 2, 3, 4]:
        output_path = resolve_output_path(stage_num, project_dir, wave)
        status[stage_num] = "done" if output_path.exists() else "pending"

    # Stage 5: deck.pptx exists
    deck_path = project_dir / f"output/{wave}/deck.pptx"
    status[5] = "done" if deck_path.exists() else "pending"

    return status


def get_deck_path(project_name: str, wave: str) -> Path | None:
    """Return path to deck.pptx if it exists, else None."""
    p = PROJECTS_DIR / project_name / f"output/{wave}/deck.pptx"
    return p if p.exists() else None


def get_input_file_status(project_name: str, wave: str) -> dict[str, bool]:
    """
    Check which Stage 1 input files are present.
    Returns {filename: exists_bool} for all expected inputs.
    """
    project_dir = PROJECTS_DIR / project_name
    stage1_inputs = STAGE_DEFS[1]["inputs"]
    result = {}
    for label, rel_template in stage1_inputs.items():
        rel = rel_template.replace("{wave}", wave)
        result[label] = (project_dir / rel).exists()
    return result


def list_output_files(project_name: str, wave: str) -> list[Path]:
    """Return all generated files in output/{wave}/."""
    out_dir = PROJECTS_DIR / project_name / f"output/{wave}"
    if not out_dir.exists():
        return []
    return [f for f in out_dir.iterdir() if f.is_file()]
