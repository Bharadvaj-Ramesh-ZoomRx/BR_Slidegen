"""
skill_loader.py — Reads SKILL.md files and assembles Claude API prompts.

Extracts all input files for each pipeline stage and builds the
system prompt + user message payload ready for the Claude API.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────────

REPO_ROOT = Path(__file__).parent.parent
SKILLS_DIR = REPO_ROOT / ".claude" / "skills"

# ── Stage definitions ─────────────────────────────────────────────────────────
#   Each stage maps to:
#     skill      — subdirectory name under .claude/skills/ (None = custom prompt)
#     label      — human-readable name
#     description — one-liner for UI
#     inputs     — {label: path_template} where {wave} is substituted at runtime
#     output     — output file path template (relative to project_dir)

STAGE_DEFS: dict[int, dict] = {
    1: {
        "skill": "build-project-context",
        "label": "Build Project Context",
        "description": "Synthesize study docs → project_context.md",
        "inputs": {
            "Call Notes": "input/wave/{wave}/call_notes.docx",
            "Methodology & KBQs (ODT)": "input/wave/{wave}/pet_project_kbq.odt",
            "KBQs": "input/wave/{wave}/kbqs.md",
            "Market Context": "input/wave/{wave}/market_context.md",
            "Prior Wave ES": "input/wave/{wave}/prior_wave_es.md",
        },
        "output": "context/{wave}/project_context.md",
    },
    2: {
        "skill": "hypotheses",
        "label": "Generate Hypothesis Bank",
        "description": "Context files → hypothesis_bank.md",
        "inputs": {
            "Market Context": "input/wave/{wave}/market_context.md",
            "Project Context": "context/{wave}/project_context.md",
            "KBQs": "input/wave/{wave}/kbqs.md",
            "Survey Context": "input/wave/{wave}/survey_context.md",
        },
        "output": "context/{wave}/hypothesis_bank.md",
    },
    3: {
        "skill": "slide-plan",
        "label": "Build Slide Plan",
        "description": "Hypothesis bank → slide_plan.md",
        "inputs": {
            "Hypothesis Bank": "context/{wave}/hypothesis_bank.md",
            "KBQs": "input/wave/{wave}/kbqs.md",
            "Survey Context": "input/wave/{wave}/survey_context.md",
        },
        "output": "context/{wave}/slide_plan.md",
    },
    4: {
        "skill": None,  # uses custom prompt below
        "label": "Generate Config YAML",
        "description": "Slide plan → config.yaml (review required)",
        "inputs": {
            "Slide Plan": "context/{wave}/slide_plan.md",
        },
        "output": "config.yaml",
    },
}


# ── Skill loading ──────────────────────────────────────────────────────────────

def load_skill_md(skill_name: str) -> str:
    """Load SKILL.md content and strip YAML frontmatter."""
    skill_path = SKILLS_DIR / skill_name / "SKILL.md"
    if not skill_path.exists():
        raise FileNotFoundError(f"Skill not found: {skill_path}")
    content = skill_path.read_text(encoding="utf-8")
    # Strip --- frontmatter ---
    if content.startswith("---"):
        parts = content.split("---", 2)
        if len(parts) >= 3:
            content = parts[2].strip()
    return content


# ── File content extraction ────────────────────────────────────────────────────

def extract_file_content(file_path: Path) -> str:
    """Extract readable text from .md, .txt, .docx, or .odt files."""
    suffix = file_path.suffix.lower()

    if suffix in (".md", ".txt"):
        return file_path.read_text(encoding="utf-8", errors="replace")

    elif suffix == ".docx":
        return _extract_docx(file_path)

    elif suffix == ".odt":
        return _extract_odt(file_path)

    elif suffix == ".json":
        return file_path.read_text(encoding="utf-8", errors="replace")

    elif suffix == ".yaml" or suffix == ".yml":
        return file_path.read_text(encoding="utf-8", errors="replace")

    else:
        return f"[Unsupported file type: {suffix} — manual review needed]"


def _extract_docx(path: Path) -> str:
    """Extract plain text from .docx via zipfile (no external deps)."""
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("word/document.xml") as f:
                xml = f.read().decode("utf-8")
        # Insert newlines at paragraph boundaries before stripping tags
        xml = re.sub(r"<w:p[ />]", "\n<w:p>", xml)
        xml = re.sub(r"</w:p>", "\n", xml)
        text = re.sub(r"<[^>]+>", "", xml)
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()
    except Exception as e:
        return f"[Failed to extract .docx: {e}]"


def _extract_odt(path: Path) -> str:
    """Extract plain text from .odt via zipfile."""
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("content.xml") as f:
                xml = f.read().decode("utf-8")
        # Insert newlines at text paragraph elements
        xml = re.sub(r"<text:p[ />]", "\n", xml)
        xml = re.sub(r"<text:line-break/>", "\n", xml)
        text = re.sub(r"<[^>]+>", "", xml)
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()
    except Exception as e:
        return f"[Failed to extract .odt: {e}]"


# ── Prompt assembly ────────────────────────────────────────────────────────────

def build_stage_prompt(
    stage_num: int,
    project_dir: Path,
    wave: str,
) -> tuple[str, str]:
    """
    Build (system_prompt, user_message) for a pipeline stage.

    The system prompt is the SKILL.md content (or a custom prompt for Stage 4),
    with a web-UI preamble that tells Claude to skip file-path gathering steps.

    The user message contains all input file contents injected inline.
    """
    stage = STAGE_DEFS[stage_num]

    # ── System prompt ──
    if stage["skill"]:
        skill_body = load_skill_md(stage["skill"])
        system = (
            "IMPORTANT — WEB UI CONTEXT: You are running inside a web-based pipeline UI. "
            "All required input files have already been gathered and are provided inline "
            "in the user message below. Skip any steps that ask the user for file paths "
            "or confirmation of paths — all files are already included. "
            "Do not ask clarifying questions about file locations. "
            "Proceed directly to analysis and generate the complete output document "
            "without truncation.\n\n"
            + skill_body
        )
    else:
        system = _stage4_system_prompt()

    # ── User message: inject file contents ──
    parts = [
        f"Please **{stage['label']}** for wave **{wave}**.\n\n"
        f"Here are all the input files:\n"
    ]

    missing = []
    for label, rel_template in stage["inputs"].items():
        rel_path = rel_template.replace("{wave}", wave)
        file_path = project_dir / rel_path
        if file_path.exists():
            content = extract_file_content(file_path)
            parts.append(
                f"\n---\n### {label}\n"
                f"**File:** `{rel_path}`\n\n"
                f"{content}\n"
            )
        else:
            missing.append(rel_path)
            parts.append(
                f"\n---\n### {label}\n"
                f"**File:** `{rel_path}`\n\n"
                f"⚠️ FILE NOT FOUND — skip this input.\n"
            )

    # Stage 4 extras: source_data.json _sheets + current config template
    if stage_num == 4:
        json_path = project_dir / f"context/{wave}/source_data.json"
        if json_path.exists():
            with open(json_path, encoding="utf-8") as f:
                data = json.load(f)
            sheets = data.get("_sheets", {})
            # Truncate to avoid token overflow — _sheets can be large
            sheets_text = json.dumps(sheets, indent=2)
            if len(sheets_text) > 30_000:
                sheets_text = sheets_text[:30_000] + "\n... [truncated]"
            parts.append(
                f"\n---\n### Source Data Index (_sheets)\n"
                f"Search this to find valid question codes for config.yaml extractions.\n\n"
                f"```json\n{sheets_text}\n```\n"
            )

        config_path = project_dir / "config.yaml"
        if config_path.exists():
            parts.append(
                f"\n---\n### Current config.yaml (structure reference)\n\n"
                f"```yaml\n{config_path.read_text(encoding='utf-8')}\n```\n"
            )

    if missing:
        parts.append(
            f"\n---\n⚠️ **Missing files:** {', '.join(missing)}\n"
            f"Generate output based on available files only.\n"
        )

    parts.append(
        f"\n---\n\nGenerate the complete **{stage['label']}** output now. "
        f"Write the full document without truncation."
    )

    user_message = "\n".join(parts)
    return system, user_message


def _stage4_system_prompt() -> str:
    return """You are a senior data pipeline engineer. Your job is to update a config.yaml file for the SlideGen YAML-driven PowerPoint pipeline based on a slide plan.

## Your task

Read the Slide Plan and Source Data Index provided. Generate a complete, updated config.yaml.

## Rules

1. For every question code in the slide plan, search the Source Data Index (_sheets) to verify it exists before using it.
2. If a code is NOT found in _sheets, add a comment: `# ⚠️ NOT FOUND IN SOURCE DATA — verify manually`
3. Preserve all top-level fields from the current config.yaml (project, brands, sections, template_path, etc.).
4. Only update the `extractions` and `asks` sections to reflect the new slide plan.
5. Use the exact `slide_type` values from the slide plan — these must match valid renderer keys.
6. Each slide in the plan becomes one `ask` entry. Each unique data source becomes one `extraction` entry.
7. If a slide_type is not in the slide plan, derive it from the chart description.

## Output

Output ONLY the complete updated config.yaml content.
- No markdown code fencing
- No explanation before or after
- Raw YAML only, ready to save as config.yaml
"""


# ── Helpers ────────────────────────────────────────────────────────────────────

def get_stage_def(stage_num: int) -> dict:
    return STAGE_DEFS[stage_num]


def resolve_output_path(stage_num: int, project_dir: Path, wave: str) -> Path:
    """Resolve the output file path for a stage."""
    template = STAGE_DEFS[stage_num]["output"]
    rel = template.replace("{wave}", wave)
    return project_dir / rel


def check_missing_inputs(stage_num: int, project_dir: Path, wave: str) -> list[str]:
    """Return list of missing input file paths for a stage."""
    stage = STAGE_DEFS[stage_num]
    missing = []
    for label, rel_template in stage["inputs"].items():
        rel = rel_template.replace("{wave}", wave)
        if not (project_dir / rel).exists():
            missing.append(f"{label} ({rel})")
    return missing
