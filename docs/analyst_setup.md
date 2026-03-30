# Analyst Setup Guide — SlideGen

## Prerequisites

| Requirement | Details |
|-------------|---------|
| OS | Windows 10/11 (or macOS for development) |
| PowerPoint | Desktop version (Microsoft 365 or standalone) |
| Python | 3.10+ |
| Git | For cloning the repo |
| Claude Code | Installed and authenticated |
| OneDrive | Configured with access to the shared SlideGen projects folder |

## Overview

SlideGen uses a **git + OneDrive** split:

- **Git repo** provides the `slidegen/` Python package, `.claude/skills/`, and docs — updated via `git pull`
- **OneDrive shared folder** is where `projects/` lives — each analyst creates/manages projects here directly. Project data, configs, templates, and output all live in OneDrive. The `projects/` folder is gitignored.

## Step 1 — Clone the Repo

```bash
git clone <repo-url> SlideGen
cd SlideGen
```

This gives you:
- `slidegen/` — Python package (pipeline + pptx_utils)
- `.claude/skills/` — Claude Code skills (auto-discovered)
- `docs/` — Documentation

## Step 2 — Install Python Dependencies

```bash
pip install pandas openpyxl python-pptx lxml pyyaml requests pywin32
```

## Step 3 — Set Up OneDrive Projects Folder

Your IT admin will share a OneDrive/SharePoint folder for SlideGen projects. Once shared, it auto-syncs to your machine. The typical path is:

```
C:\Users\<you>\OneDrive - ZoomRx\Shared\SlideGen\projects\
```

Symlink this into your git clone so the pipeline can find it:

**Windows (PowerShell as Administrator):**
```powershell
cd SlideGen
New-Item -ItemType SymbolicLink -Path "projects" -Target "C:\Users\<you>\OneDrive - ZoomRx\Shared\SlideGen\projects"
```

**macOS/Linux:**
```bash
cd SlideGen
ln -s "/path/to/OneDrive/Shared/SlideGen/projects" projects
```

The `projects/` folder is gitignored, so this symlink won't cause git conflicts.

## Step 4 — Verify Setup

```bash
# Check slidegen package is importable
python -c "from slidegen.pipeline import generate_deck; print('OK')"

# Check skills are visible (from git)
ls .claude/skills/slidegen/

# Check projects folder (from OneDrive)
ls projects/
```

In Claude Code:
```
> What skills do you have for slidegen?
```

## Step 5 — Create a New Project

Create a project folder in the OneDrive projects directory (either directly or via your symlink):

```
projects/my_project/
├── config.yaml              # Copy from another project and edit
├── input/wave/<wave_name>/  # Drop source_data.xlsx + reference docs here
├── templates/               # Drop template.pptx here
├── context/                 # System-generated (auto-created)
└── output/                  # Generated decks (auto-created)
```

Since `projects/` lives on OneDrive, new projects are visible to all team members automatically.

## Step 6 — Generate a Deck

In the Claude Code terminal:

```
Create slides for projects/my_project
```

Claude will run the full pipeline:
1. Index Excel data
2. Build project context (you confirm)
3. Generate hypothesis bank (you confirm)
4. Build slide plan (you confirm)
5. Generate config.yaml
6. Build the deck → `output/<wave>/deck.pptx`

## Step 7 — Edit Slides

After initial generation, use natural language commands:

```
Edit Slide 5 — change headline to "Updated Message Recall"
Edit Slide 3 — sort bars descending by current value
Refresh this deck                    # Re-extract all data and rebuild
Regenerate all slides                # Rebuild from current config
Regenerate with fresh data           # Force re-extraction even if cache is valid
```

## Synapse API (Optional)

If your project uses Synapse for data, set the API key:

```powershell
$env:SYNAPSE_API_KEY = "your-api-key-here"
```

Without the key, the pipeline uses the Excel file in `input/wave/` — no Synapse calls are made.

## Updates

| What | How to Update |
|------|--------------|
| `slidegen/` package, `.claude/skills/`, docs | `git pull` in your clone |
| Project configs, data, output | Managed directly in OneDrive — auto-syncs across team |
| Python dependencies | Re-run `pip install pandas openpyxl python-pptx lxml pyyaml requests pywin32` |

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `ModuleNotFoundError: slidegen` | Run from inside the git clone directory (or add it to PYTHONPATH) |
| `FileNotFoundError: template.pptx` | Place template in `projects/<name>/templates/` |
| `PermissionError` on save | Close the deck in PowerPoint, then retry |
| Stale data warning | Run with `force_fresh=True` or delete `source_data.json` |
| Skills not found | Run `git pull` to get latest skills from the repo |
| Projects folder missing | Re-create symlink: `ln -s /path/to/OneDrive/projects projects` |
| OneDrive not syncing | Check OneDrive app is running and signed in |
