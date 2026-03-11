# SlideGen — Product Requirements Document

**Author:** Sivakumar K (VP of Technology)
**Date:** 2026-03-10
**Version:** 2.0
**Status:** Active Specification
**Tech Lead:** Rajesh
**Rubric Provider:** Vinoth (consulting analyst)

---

## 1. Product Overview

### 1.1 Problem Statement

ZoomRx consulting analysts manually build PowerPoint slide decks for client deliverables. Every engagement repeats the same work — extracting survey data, building charts, aligning layout grids, applying brand formatting — with no institutional accumulation of the effort.

**This system solves three problems:**

1. **Creation** — Generate slides programmatically from structured survey data, at client-delivery quality, without manually placing shapes and charts in PowerPoint.
2. **Editing** — Allow analysts to make natural language surgical edits to generated slides, with changes reflected live in PowerPoint's own canvas.
3. **Knowledge retention** — Capture all hard-won OOXML knowledge, layout patterns, and client brand specs in a reusable skill + resource library that accumulates value across engagements.

### 1.2 Solution Summary

SlideGen is an AI-assisted PowerPoint generation and editing system. Claude Code (terminal) is the analyst's interface. Python scripts use `python-pptx` to create slides from scratch and `win32com` to edit them live in a running PowerPoint instance. A shared skill library — utility functions, brand definitions, layout presets, and workflow instructions — ensures knowledge is captured once and reused across all engagements.

### 1.3 Target User

Consulting analysts at ZoomRx who build client deliverable slide decks.

- All analysts are on **Windows** with **desktop PowerPoint**
- Analysts are not developers but can use terminal-based tools with guidance
- Analysts may make manual edits in PowerPoint between AI-assisted sessions

### 1.4 V1 Success Criteria

1. **Analyst describes a slide → gets client-delivery-quality pptx** — native charts, correct brand formatting, precise layout — without touching PowerPoint during creation.
2. **Analyst types edits in Claude Code → sees changes on PowerPoint canvas live** within seconds.
3. **Manual PowerPoint edits don't break anything.** The next AI session picks up from the actual current state.
4. **New client = add a brand definition + project context**, no code changes to the core system.
5. **Data flows from Synapse to slides** without manual Excel wrangling.
6. **Knowledge accumulates.** Every OOXML workaround, layout pattern, or chart type solved during any engagement is captured in the utility library and available to all future engagements.

---

## 2. System Architecture

### 2.1 Two-Tool Model

SlideGen uses two complementary Python libraries:

| Tool | Phase | Why |
|---|---|---|
| **python-pptx** (+ lxml helpers) | **Creation** — building slides from scratch | Works with files on disk. Doesn't require PowerPoint to be running. Clean Pythonic API for construction. Proven. |
| **win32com** (pywin32) | **Editing** — surgical changes to open slides | Connects to the running PowerPoint process via Windows COM. Changes appear on the analyst's screen immediately. No file lock conflicts. Full PowerPoint object model access. |

The two tools share the same OOXML file format and the same shape naming convention (`zrx_` prefix). Shapes created by `python-pptx` are accessible by name via `win32com`.

### 2.2 Component Map

```
COMPANY SKILLS (shared OneDrive/SharePoint folder, read-only by projects)
├── SKILL.md                 — Workflow rules for slide generation & editing
├── pptx_utils/              — Pre-solved lxml helpers, shape primitives, chart builders
│   ├── brand.py             — Client brand definitions (BRAND{})
│   ├── layout.py            — Standard grid presets (LAYOUTS{})
│   ├── shapes.py            — Layout primitives (textbox, solidrect, horiz_line, etc.)
│   ├── text.py              — Text formatting helpers
│   ├── charts.py            — Chart builders + chart pattern definitions (CHART_PATTERNS{})
│   ├── tables.py            — Table builders (delta tables, data tables)
│   ├── registry.py          — Shape registry management + reconciliation
│   ├── images.py            — Image/logo placement
│   ├── deck.py              — Template handling, slide scaffolding, save behavior
│   └── lxml_helpers.py      — Raw OOXML XML manipulation (the hard-won stuff)
├── SLIDE_RULES.md           — Accumulated formatting rules & pitfalls
├── synapse_fetch.py         — API client for Synapse data ingestion
└── CHART_PATTERNS.md        — Named chart pattern documentation

PROJECT DIRECTORY (per-engagement, analyst's local machine or OneDrive)
├── CLAUDE.md                — Thin pointer: "read project.connect/project_context.md"
├── *.pptx                   — The slide files (each embeds URL to project.connect/)
└── project.connect/         — All project state lives here
    ├── project_context.md   — Client, brand, survey IDs, wave config, template path
    ├── analyst_notes.md     — Accumulated observations, client feedback (optional)
    ├── data/
    │   └── banner_plan.pkl  — Full banner plan snapshot (shared across all decks)
    ├── decks/
    │   ├── PET_Q4_2025/
    │   │   ├── registry.json    — Shape state for this deck (live-reconciled)
    │   │   ├── edit_log.json    — Audit trail
    │   │   └── scripts/         — Generation scripts for this deck
    │   └── PET_Q3_2025/
    │       └── ...
    └── backups/
        └── PET_Q4_2025_2026-03-20T1430.pptx

RUNTIME (analyst's Windows machine)
├── Claude Code              — Front-end interface (terminal)
├── Python + win32com        — Execution engine, COM bridge to PowerPoint
└── PowerPoint.exe           — Live canvas, source of truth

CLOUD
└── Claude API               — Natural language → edit script translation
```

### 2.3 Tokens = Latency

Under a Claude Code subscription, token cost is not the concern — but **every token Claude generates is time the analyst waits**. This is an architecture-level design principle that shapes every component:

- **Utility library:** Claude writes ~30-line composition scripts instead of ~200-line scripts rediscovering lxml from scratch. Directly reduces latency per operation.
- **Brand/layout definitions:** Pre-computed constants mean Claude never guesses coordinates or colors.
- **Skill instructions:** Worked examples let Claude match patterns instead of reasoning from first principles.
- **Registry:** Pre-reconciled state means Claude doesn't need to explore the slide to understand what's on it.

Every design decision that increases the library's coverage reduces wall-clock time for the analyst.

### 2.4 project.connect/ Folder

Every project has a single `project.connect/` folder that holds all state — registries, generation scripts, data snapshots, edit logs, and backups. The `.pptx` files live alongside it in the project directory.

**Key rules:**

1. **Per-project, not per-deck.** Multiple decks in the same engagement share one `project.connect/`. Data and context are shared; per-deck state (registry, scripts, edit log) lives in `project.connect/decks/<deckname>/`.

2. **`.pptx` files embed a URL to `project.connect/`.** Each generated `.pptx` includes a custom document property (`zrx_connect_url`) pointing to the `project.connect/` folder's OneDrive/SharePoint location. Anyone opening a deck can find its companion data.

3. **`CLAUDE.md` at the project root is a thin pointer.** It tells Claude Code where to find the real context:
   ```markdown
   # Project: PET Q4 2025 — Johnson & Johnson
   Project context and all artifacts live in `project.connect/`.
   Read `project.connect/project_context.md` for full context before any operation.
   ```

4. **Context as `.md` files.** Project context, analyst notes, deck briefs, and client feedback are stored as markdown files inside `project.connect/`. Markdown is Claude's most natural input format.

5. **V1 constraint:** Single analyst per project at a time. Multi-analyst concurrent editing is out of scope.

### 2.5 Skill Distribution via OneDrive

Company skills are distributed via a shared OneDrive/SharePoint folder that auto-syncs to each analyst's machine as a local filesystem path.

- **Analysts never interact with git.** OneDrive mounts transparently as a filesystem — updated skills sync automatically.
- **The developer** building skills maintains a git repo locally for version control, then copies updated files to the shared OneDrive folder.
- **V1.5:** CI-based sync (git push → automated copy to OneDrive) when scaling beyond a handful of users.
- **Risk accepted:** No automated versioning or rollback on the OneDrive side. Acceptable for V1's small user base and single developer.

Skills install globally (in `~/.claude/`) or per-project (in project `.claude/`).

---

## 3. Data Layer — Synapse Integration

### 3.1 What Synapse Is

Synapse is ZoomRx's internal ETL'd database system. It stores survey response data with response codes mapped to actual values. It provides APIs for pulling cross-tab survey data — the raw material for consulting slide decks.

- **Auth:** API key/token, stored in environment variable `SYNAPSE_API_KEY`
- **Input:** Survey IDs, wave/quarter identifiers, metric types, segment definitions — provided via `project.connect/project_context.md`

### 3.2 Raw Question Data

Per-question data is the atomic unit. Each question in a survey can be fetched with segmentation breakdowns (by HCP type, region, specialty, etc.). This produces cross-tab results: categories × series × values per wave/quarter.

### 3.3 Virtual Questions

Virtual questions combine responses from multiple raw questions via rules or regex matching. They enable derived metrics (e.g., "aided awareness of any message" = union of individual message recall questions). Virtual questions are defined in Synapse's configuration layer and fetched through the same API as raw questions.

### 3.4 Banner Plans

Banner plans are the standard export format from Synapse — cross-tab survey analysis across all segments.

- **Single-question analysis:** One question × all segments. ~13-14 KB for a small survey.
- **Multi-question analysis (project analysis):** Multiple questions analyzed together. Previously excluded from some implementations for no technical reason — included in SlideGen.
- **Design principle:** Pull the full banner plan in one batch up front. Only do incremental pulls if the system discovers it needs additional data mid-session. Minimize API calls → minimize latency.

### 3.5 Non-Synapsible Data

Not everything is "synapsible." Some analysis requires raw respondent-level data or custom statistical processing that Synapse's standard API doesn't cover.

**In SlideGen, synapsibility is NOT a precondition.** The key design principle is: store the **procedure** (script) that created a data object so it can be re-run with fresh data. Whether the procedure calls the Synapse API, runs a SQL query, or executes a custom Python analysis — the refresh mechanism is the same.

These "stored procedures" live in `project.connect/decks/<deckname>/scripts/` alongside the generation scripts.

### 3.6 Data Cache

```
Synapse API → synapse_fetch.py → banner_plan.pkl → generation scripts → .pptx
               (one batch pull)    (full banner plan)
```

- **pkl as intermediate format.** Binary, fast serialization, preserves Python data types exactly. Zero-config.
- **Timestamped snapshot.** The pkl is the audit artifact — you can verify exactly what data produced each version of the deck.
- **Offline resilience.** Slide generation works from pkl, not live API calls. If the API is unavailable, generation proceeds from cached data and the analyst is informed.
- **pkl schema:** Standardized dict structure — `{slide_id: {metric: {category: {wave: value}}}}`. Must accommodate both single-question and multi-question analysis. Exact schema to be finalized during implementation.

### 3.7 Extensibility

Synapse's capabilities evolve. New API endpoints, new analysis types, new segmentation dimensions can all be supported by extending `synapse_fetch.py` without changing the rest of the system. The pkl schema is the stable interface between data ingestion and slide generation.

---

## 4. Utils Library Specification

### 4.1 Design Rules

1. **All lxml XML manipulation lives in the library and nowhere else.** Claude never writes inline lxml in generated scripts — it calls library functions.
2. **Every function that manipulates raw XML includes a docstring naming the OOXML element it targets** and why `python-pptx` doesn't cover it.
3. **New formatting needs discovered during engagements are added as named functions**, not left inline in project scripts.
4. **If the needed operation isn't in the library, Claude may write inline lxml but must flag it** for later extraction into the library.
5. **Start thin, grow deliberately.** Begin with the ~10 most-used lxml helpers. Let Claude handle the long tail via inline lxml + skill instructions. Expand the library based on usage patterns, not speculation.

### 4.2 Module Organization

The utility library is organized as a Python package `pptx_utils/` with semantically named modules:

| Module | Responsibility | Examples |
|---|---|---|
| `brand.py` | Client brand definitions | `BRAND{}` dict, `get_brand()`, `get_color()` |
| `layout.py` | Grid presets and coordinate systems | `LAYOUTS{}` dict, layout computation helpers |
| `shapes.py` | Layout primitives | `textbox()`, `solidrect()`, `horiz_line()`, `add_delta_col()` |
| `text.py` | Text formatting helpers | Run-level formatting, paragraph alignment |
| `charts.py` | Chart builders + pattern definitions | `make_clustered_bar()`, `make_stacked_bar()`, `make_waterfall()`, `CHART_PATTERNS{}` |
| `tables.py` | Table builders | Delta tables, data summary tables |
| `registry.py` | Shape registry management | `reconcile_registry()`, `register_shape()`, `update_registry()` |
| `images.py` | Image and logo placement | Logo insertion, image sizing |
| `deck.py` | Template handling, slide scaffolding | `open_template()`, `add_blank_slide()`, `save_deck()` |
| `lxml_helpers.py` | Raw OOXML XML manipulation | All `set_*()`, `invert_*()`, `hide_*()` functions |

**When to split from single file:** V1 may start as a single `pptx_utils.py` if the function count is < 30. Split into a package when it exceeds ~50 functions or when multiple developers need to work on it concurrently.

### 4.3 BRAND{} Schema

Client brand definitions. Keyed by client identifier. Derived from client brand guides.

```python
BRAND = {
    "JJ": {
        # Core palette
        "primary":        RGBColor(0xFF, 0x00, 0x00),
        "q_current":      RGBColor(0xF7, 0x58, 0x24),
        "q_previous":     RGBColor(0xFF, 0xC1, 0x99),
        "positive":       RGBColor(0x00, 0xB0, 0x50),
        "negative":       RGBColor(0xFF, 0x00, 0x00),
        "neutral_grey":   RGBColor(0x50, 0x50, 0x50),

        # Typography
        "font_heading":   "Calibri",
        "font_body":      "Calibri",

        # Template
        "template_path":  "templates/jj_q4_2025.pptx",

        # Margins and spacing (inches)
        "slide_margin_left":   0.20,
        "slide_margin_right":  0.20,
        "slide_margin_top":    1.47,  # below header
        "slide_margin_bottom": 0.20,

        # Logo
        "logo_path":      "assets/jj_logo.png",
        "logo_position":  {"left": 0.1, "top": 0.1, "width": 0.8},
    },
    # Additional clients added as engagements onboard
}
```

**Growth model:** New clients are added to `BRAND{}` as their first engagement uses the system. Existing entries are updated when brand guides change.

### 4.4 LAYOUTS{} Specification

Pre-computed layout coordinates for common slide compositions. Values are in inches.

```python
LAYOUTS = {
    "two_chart_with_delta": {
        "chart_top": 1.47, "chart_h": 5.05,
        "left_chart_left": 0.20, "left_chart_w": 7.30,
        "delta_w": 0.62, "delta_gap": 0.04, "chart_gap": 0.10,
    },
    "single_chart_full": { ... },
    "three_column_comparison": { ... },
    "table_only": { ... },
}
```

These are starter examples. The exact set of layouts will be informed by the rubric exercise with the consulting team (Vinoth) and by reverse-engineering existing client decks to identify recurring patterns.

**Open question:** Should layouts be organized by metric type, layout type, or both? See Section 13 (Open Questions).

### 4.5 CHART_PATTERNS{}

Named chart patterns with expected data shapes and formatting defaults. Each pattern encodes the chart type, axis configuration, series formatting, and lxml workarounds needed.

```python
CHART_PATTERNS = {
    "clustered_bar_horizontal": {
        "chart_type": XL_CHART_TYPE.BAR_CLUSTERED,
        "axis_inverted": True,
        "gap_width": 80,
        "overlap": 0,
        "data_label_pos": "outside_end",
        "series_no_border": True,
        "data_shape": "categories × series × values",
    },
    "stacked_bar_100pct": { ... },
    "waterfall": { ... },
    # Patterns grow as new chart types are encountered
}
```

Each pattern maps directly to a combination of `python-pptx` API calls and `lxml_helpers` functions. Claude reads the pattern name from a skill instruction and composes the right sequence of calls.

### 4.6 SKILL.md Format

The skill file that Claude Code reads at session start. Required fields:

```markdown
# SlideGen Skill

## Cardinal Rules
- Never write raw lxml in generation scripts — call pptx_utils functions
- Never hardcode brand colors inline — use BRAND{}
- Always open a template — never create a blank presentation
- Always reconcile the registry before any edit — no exceptions
- Flag any inline lxml for later extraction into pptx_utils

## Standard Workflow
### Creation
1. Read project_context.md for client context
2. Confirm required inputs (data source, metrics, slide type, etc.)
3. Fetch data if needed (synapse_fetch.py → pkl)
4. Write generation script using pptx_utils composition
5. Execute, confirm, tell analyst to open in PowerPoint

### Editing
1. Reconcile registry (always, before any edit)
2. Read reconciled registry
3. Identify target shape from analyst's description
4. Write minimal edit script using pptx_utils + win32com
5. Execute → live canvas update
6. Update registry + edit log
7. Confirm in plain English

## Required Inputs Checklist
- [ ] Data source (pkl path or Synapse query)
- [ ] Client/brand
- [ ] Slide type / layout
- [ ] Metrics to visualize
- [ ] Time periods / waves
- [ ] Sort order
- [ ] Template path
- [ ] Output path

## Decision Rules
- Recreate vs. edit in-place: [criteria]
- Ask for clarification vs. act immediately: [criteria]
- Use COM vs. file-based python-pptx: [criteria]

## Worked Patterns
[Code examples for common slide types showing composition from pptx_utils]

## Known Limitations
[What python-pptx can/cannot do, with pointers to pptx_utils functions]
```

---

## 5. Skills Architecture

### 5.1 What a Skill Is

A skill is a self-contained instruction set that teaches Claude how to produce a specific kind of output. In SlideGen, skills encode:

- **What** the output should look like (layout, chart types, data mapping)
- **How** to build it (which `pptx_utils` functions to compose, in what order)
- **What data** it needs (pkl keys, Synapse queries, input parameters)

A skill is NOT executable code. It is a set of instructions that Claude reads and follows to write executable code. The executable code is the generation/edit script that Claude produces each time.

### 5.2 Company Skills vs Project Artifacts

| | Company Skills | Project Artifacts |
|---|---|---|
| **Scope** | Reusable across all engagements | Specific to one engagement |
| **Location** | Shared OneDrive/SharePoint folder | `project.connect/` |
| **Examples** | `pptx_utils/`, `BRAND{}`, `LAYOUTS{}`, `SKILL.md`, `synapse_fetch.py` | `registry.json`, `edit_log.json`, `banner_plan.pkl`, generation scripts |
| **Ownership** | Developer (Rajesh) maintains; read-only for analysts | Created per-project; disposable after delivery |
| **Growth model** | Accumulates value over time — never loses old capabilities | Lives and dies with the engagement |

**Key separation principle:** Company skills are *how to build slides* (reusable, shared). Project artifacts are *what was built* (per-engagement, disposable). Mixing them causes ownership confusion and makes the skill library fragile.

### 5.3 Skill Hierarchy — Open Question

How should skills be organized? There is a design tension:

- **By metric type** (MR slide, ME slide, RFSOV slide) — gives consistent output for the same kind of analysis
- **By layout type** (left-right chart, single chart full, three-column comparison) — separates data from presentation
- **Both** — metric skills that reference layout skills

**Sriram's concern:** Unconstrained Claude gives different layouts every time. Skills should provide a "track to follow."
**Siva's pushback:** Over-constraining makes it dumb. Claude should decide the best visualization given the data.

**Finding the right balance between giving Claude a track to follow vs. letting it decide freely is the art of this system.** This question will be informed by the rubric exercise with Vinoth — what instruction does the analyst give, what output do they expect?

### 5.4 Global vs Per-Project Installation

- **Global skills** (`~/.claude/` or OneDrive sync target): `pptx_utils/`, `BRAND{}`, `SKILL.md`, `synapse_fetch.py` — shared across all projects
- **Per-project skills** (project `.claude/`): Project-specific overrides, custom layouts for a particular client's deck style, engagement-specific analysis scripts

Per-project skills can extend or override global skills. For example, a project might define a custom chart pattern that only applies to one client's preferred visualization style.

### 5.5 Skill Authoring Guidelines

When creating or extending skills:

1. **Start from a real deck.** Reverse-engineer an existing client-delivery slide to extract the pattern.
2. **Capture the hard parts.** The value of a skill is encoding the non-obvious — lxml workarounds, precise coordinates, formatting tricks that took hours to figure out.
3. **Include worked examples.** Show complete generation scripts that use `pptx_utils` composition.
4. **Specify data shape.** Document exactly what pkl structure the skill expects.
5. **Test against PowerPoint's repair checker.** Generated files must open without repair prompts.

---

## 6. Shape Registry

### 6.1 Purpose

The registry is Claude's cache of what's on the slide. It maps opaque shape IDs to semantic labels, tracks property values, and records data lineage for refresh capability.

**Core principle:** PowerPoint is always the source of truth. The registry is never trusted from a previous session. Before every edit, the system reads live state from PowerPoint and updates the registry.

### 6.2 Registry JSON Schema

```json
{
  "deck": "PET_Q4_2025",
  "last_reconciled": "2026-03-15T14:30:00",
  "slide_index": 0,
  "shapes": {
    "zrx_001": {
      "label": "MR clustered bar chart",
      "type": "chart",
      "metric": "message_recall",
      "position": "left",
      "properties": {
        "left": 0.20,
        "top": 1.47,
        "width": 7.30,
        "height": 5.05
      },
      "data_source": {
        "pkl_file": "banner_plan.pkl",
        "pkl_key": "slide_01.message_recall",
        "synapse_query": "survey=PET_Q4_2025&question=Q12&segments=all"
      },
      "generation_script": "scripts/gen_mr_chart.py"
    },
    "zrx_002": {
      "label": "Slide headline",
      "type": "textbox",
      "role": "headline",
      "properties": {
        "left": 0.20,
        "top": 0.10,
        "width": 12.0,
        "height": 0.50,
        "text": "Message Recall — Rybrevant",
        "font_size": 12,
        "font_color": "505050"
      }
    }
  }
}
```

**Key fields:**
- `label` — Human-readable description, used by Claude to match analyst's natural language to the right shape
- `type` — chart, textbox, rectangle, image, table, group
- `data_source` — Lineage for data-driven shapes: which pkl slice, which Synapse query, which generation script. Enables refresh.
- `generation_script` — Path to the script that created this shape. Enables regeneration.
- `properties` — Live values, updated on every reconciliation

### 6.3 Reconciliation Algorithm

Runs before every edit, no exceptions.

```
Input:  Running PowerPoint instance + existing registry.json
Output: Updated registry.json with live values

Steps:
1. Connect to PowerPoint via win32com
2. Build name → shape(s) map for ALL shapes on the active slide
3. DUPLICATE DETECTION: If any zrx_* name appears more than once → HARD STOP
   - Print which names are duplicated
   - Exit with error — analyst must resolve in PowerPoint first
   - Why: two shapes sharing a name means any edit silently targets
     whichever COM enumerates first. Wrong results with no error.
4. For each registered shape: read live left, top, width, height,
   font properties, text content from PowerPoint
5. FLAG shapes in registry but missing from slide (deleted manually)
6. FLAG unregistered zrx_* shapes on slide (unexpected — copy-paste?)
7. IGNORE non-zrx_* shapes entirely (manually created, not system-managed)
8. Write updated registry to disk
```

### 6.4 zrx_ Naming Convention

All system-created shapes get a `zrx_` prefixed opaque ID (e.g., `zrx_001`, `zrx_002`). Semantic meaning lives in the registry, not in the shape name.

**Why:**
- Separates identity (stable, used for lookup) from description (can change)
- The `zrx_` prefix distinguishes managed shapes from manually-created ones (`Rectangle 7`)
- The VBA deduplication hook can operate on IDs without understanding semantics
- Human-readable and easy to debug

**The one rule for analysts:** Don't rename `zrx_*` shapes. Everything else is fair game.

---

## 7. Creation Pipeline

### 7.1 New Deck Creation

Seven-step sequence from analyst request to finished file:

1. Claude reads `SKILL.md` and `pptx_utils/` from company skills
2. Claude reads `CLAUDE.md` (thin pointer) → reads `project.connect/project_context.md` for client context (brand, template, survey IDs)
3. Claude confirms the required inputs checklist (data source, metrics, slide type, time periods, sort order, template path, output path)
4. If data not yet fetched: Claude runs `synapse_fetch.py` with survey IDs from project context → produces `project.connect/data/banner_plan.pkl`
5. Claude writes a generation script that:
   - Opens the client template via `open_template()`
   - Loads data from pkl
   - Composes the slide using `pptx_utils` functions (chart builders, layout primitives, text boxes)
   - Names every shape with a `zrx_` prefixed ID
   - Writes `project.connect/decks/<deckname>/registry.json` with shape ID → semantic label mapping (including `data_source` lineage for data-driven shapes)
   - Initializes empty `project.connect/decks/<deckname>/edit_log.json`
   - Saves the generation script to `project.connect/decks/<deckname>/scripts/`
   - Embeds `zrx_connect_url` custom property in the `.pptx`
   - Saves the pptx
6. Claude executes the script via bash
7. Claude confirms what was created and tells the analyst to open the file in PowerPoint

### 7.2 Brand Template Output Spec

For each client, reverse-engineering an existing client template produces:

- **Slide masters and layouts** — which master slides exist, naming convention, placeholder positions
- **Color theme** — mapped to `BRAND{}` keys
- **Font theme** — heading and body fonts
- **Header/footer areas** — reserved regions that generation scripts must not overlap
- **Logo placement** — position, size, which slides include it
- **Standard margins** — safe content area boundaries

This reverse-engineering is a one-time setup per client. The output feeds into `BRAND{}`, `LAYOUTS{}`, and the client template `.pptx` file used by `open_template()`.

### 7.3 Slide Initialization from Layout

When creating a new slide, the generation script:

1. Selects the appropriate slide layout from the template's master
2. Reads the corresponding `LAYOUTS{}` preset for coordinate baselines
3. Places shapes within the defined content area, respecting margins
4. Any shape placed outside the safe area triggers a warning (not a hard stop — analyst may intentionally extend)

---

## 8. Live Editing Pipeline

### 8.1 Edit Session Flow

Eight-step cycle after reconciliation:

1. Claude reads the reconciled registry to understand current state
2. Claude identifies the target shape from the analyst's natural language description using registry labels
3. **For positional changes:** Claude describes the planned change and asks for confirmation before executing
4. **For clearly scoped changes** (text, color, font size): Claude executes immediately
5. Claude writes a minimal edit script using `pptx_utils` functions and `win32com` wrappers
6. Script executes → COM message sent → PowerPoint canvas updates live
7. Claude updates `registry.json` and appends to `edit_log.json`
8. Claude confirms the change in plain English

### 8.2 win32com API Patterns

Concrete code patterns for common edit types:

| Edit type | Example NL | Code pattern |
|---|---|---|
| Text content | "Change the headline to say Q4 results" | `shape.TextFrame.TextRange.Text = "..."` |
| Font size | "Make the footer smaller" | `shape.TextFrame.TextRange.Font.Size = 5.5` |
| Color | "Make the headline dark grey" | `shape.TextFrame.TextRange.Font.Color.RGB = 0x505050` |
| Position | "Move the delta table right a bit" | `shape.Left = Inches(current + 0.15) * 914400` |
| Chart data | "Update with Q1 2026 numbers" | `chart.replace_data(new_chart_data)` via python-pptx (file-based) |
| Bar gap width | "Make the bars fatter" | `chart_obj.ChartGroups(1).GapWidth = 50` via COM |
| Axis inversion | "Flip the bar order" | `chart_obj.Axes(2).ReversePlotOrder = True` via COM |

**Chart data replacement note:** `win32com` can modify chart formatting properties directly. However, for replacing the entire dataset of a chart (adding a new quarter), it may be more reliable to use `python-pptx`'s `chart.replace_data()` on the saved file. This requires save → close → edit → reopen. The skill should document when to use COM (formatting) vs. file-based (structural data changes).

### 8.3 Instruction Decomposition

When the analyst's instruction is ambiguous, Claude must ask before acting:

- "Make the text bigger" → Which text? There may be 12 textboxes.
- "Shift it right a bit" → What is "it"? What is "a bit" (0.1"? 0.25")?
- "The colors look off" → Too vague to act on.

For clearly scoped changes ("make the headline 10pt dark grey"), Claude executes immediately. For positional changes, Claude always describes the plan first.

---

## 9. Connected Slides and Deck Refresh

### 9.1 What Connected Slides Are

Every data-driven shape Claude creates is "code-connected" — tagged in the registry with enough metadata to regenerate it from fresh data without the analyst re-specifying anything. The `data_source` field in the registry records:

- **pkl_file** — which data cache file
- **pkl_key** — which slice within the pkl
- **synapse_query** — the original Synapse API call that produced the data
- **generation_script** — the script that created the shape

This lineage makes each chart a live connection to its data source, not a static snapshot.

### 9.2 Refresh Workflow

**Target experience:** Analyst gives batch instructions → walks away → returns to a first draft of a full deck.

The "refresh this deck" command:

1. Claude reads the deck's `registry.json`
2. For each shape with a `data_source` entry:
   a. Re-runs the Synapse query (or re-reads from updated pkl)
   b. Re-executes the generation script with fresh data
   c. Updates the shape in PowerPoint via win32com (or regenerates via python-pptx if structural changes are needed)
3. Shapes without `data_source` (manually created, or text-only) are left untouched
4. Registry is updated with new property values
5. Claude reports what was refreshed and what changed

**Key scenario — fielding day:** When survey fielding closes, new data arrives. Claude refreshes all data-driven objects. The analyst then only needs to write talking headers and an executive summary on top of the refreshed data.

### 9.3 Lineage and Staleness Tracking

The registry tracks when each shape's data was last refreshed:

```json
{
  "zrx_001": {
    "data_source": {
      "pkl_file": "banner_plan.pkl",
      "pkl_key": "slide_01.message_recall",
      "synapse_query": "survey=PET_Q4_2025&question=Q12&segments=all",
      "last_data_pull": "2026-03-10T09:00:00",
      "last_refreshed": "2026-03-10T09:15:00"
    }
  }
}
```

When the analyst asks to refresh, Claude can report which shapes have stale data (pkl older than a threshold) and which are current.

---

## 10. Resilience and Edit Safety

### 10.1 Edit Log

Audit trail of all AI-assisted edits, enabling undo and session review.

```json
{
  "edits": [
    {
      "id": "edit_001",
      "timestamp": "2026-03-15T14:32:00",
      "instruction": "Make the headline smaller and dark grey",
      "shape": "zrx_002",
      "shape_label": "Slide headline",
      "changes": {
        "font_size": {"before": 12, "after": 10},
        "font_color": {"before": "FF0000", "after": "505050"}
      }
    }
  ]
}
```

### 10.2 Undo Mechanism

Reverting an edit means re-applying the `before` values from the log entry via win32com. Claude can do this when asked:
- "Undo the last change"
- "Revert edit_001"
- "Undo the headline change"

For multi-property edits, undo restores all changed properties atomically.

### 10.3 Backup Strategy

Before each edit session, a timestamped copy of the `.pptx` is saved to `project.connect/backups/`:

```
project.connect/backups/PET_Q4_2025_2026-03-20T1430.pptx
```

This is a safety net more robust than per-property undo — it handles cases where multiple edits interact in unexpected ways. The analyst can always revert to a known-good state.

### 10.4 VBA BeforeSave Hook

A VBA macro embedded in the client template file (`.pptm`). Fires on every `Ctrl+S`.

**Behavior:**
1. Scan all shapes on the current slide
2. For any `zrx_*` name that appears more than once: rename the duplicate to `{original_name}_{counter}` (e.g., `zrx_001_2`)
3. Log any new non-`zrx_*` shapes as unregistered (informational)
4. Allow save to proceed

**Why VBA:** Zero-dependency — works without Python running, fires automatically on every save, travels with the file. No installation required.

**Template format:** Requires `.pptm` (macro-enabled). If client-facing files cannot contain macros, the VBA hook is stripped on final export and the delivered `.pptx` is macro-free.

### 10.5 Error Handling

| Error | Detection | Response |
|---|---|---|
| **COM connection failure** — PowerPoint not running | `win32com.client.Dispatch()` throws | Clear message: "Open PowerPoint first." No silent failure. |
| **Duplicate shape names** | Reconciliation step 3 | Hard stop. Print duplicated names. Analyst must resolve. |
| **Missing registered shape** | Reconciliation step 5 | Warning. Shape flagged as deleted in registry. Edit proceeds on remaining shapes. |
| **Unexpected zrx_* shape** | Reconciliation step 6 | Warning. Likely copy-paste. Analyst informed. |
| **Synapse API failure** | HTTP error / timeout | Log error. Proceed from cached pkl if available. Inform analyst of data staleness. |
| **Skill function error** | Python exception in pptx_utils | Error logged with full traceback. No partial writes to registry. Backup ensures recovery. |
| **PowerPoint repair prompt** | Malformed XML in generated file | All XML manipulation centralized in `pptx_utils`. One fix propagates to all projects. |
| **Registry corruption** | JSON parse failure | Rebuild registry by reconciling from live PowerPoint state. No data loss if PPT is open. |

---

## 11. Technical Requirements

### 11.1 Analyst Machine Prerequisites

| Requirement | Details |
|---|---|
| OS | Windows 10/11 |
| PowerPoint | Desktop version (Microsoft 365 or standalone) |
| Python | 3.10+ |
| Python packages | `pywin32`, `python-pptx` (or `python-pptx-ng`), `lxml`, `openpyxl` |
| Claude Code | Installed and authenticated |
| OneDrive | Configured with access to company skills shared folder |
| Network | Access to Claude API (HTTPS outbound), access to Synapse API (internal) |

### 11.2 File Format

- **Templates:** `.pptm` (macro-enabled) to support the VBA BeforeSave hook. Stripped to `.pptx` for client delivery.
- **EMU precision:** All position/size values stored internally as EMUs (1 inch = 914,400 EMU). The `pptx_utils` API accepts inches and converts internally.
- **OOXML fidelity:** Generated files are indistinguishable from manually-created PowerPoint files. Native charts include embedded Excel workbooks. All content is editable — nothing flattened to images.

---

## 12. V1 Scope

### 12.1 In Scope

- Single analyst workflow (one analyst, one project at a time)
- J&J / PET as the first client and deliverable type
- Claude Code (terminal) as the interface
- Slide creation via python-pptx
- Live editing via win32com
- Shape registry with reconciliation
- Synapse data integration (banner plans, single + multi-question analysis)
- Skill library with brand definitions, layout presets, chart patterns
- Connected slides with data refresh capability
- Edit log with undo
- Backup before edit sessions
- VBA BeforeSave hook for name deduplication
- OneDrive-based skill distribution

### 12.2 Out of Scope

- Multi-user concurrent editing
- Web or Mac support
- Installer or setup automation
- Skill authoring UI (skills are authored by the developer in code)
- Automated visual validation / regression testing
- Deck-level cross-slide consistency enforcement
- CI-based OneDrive sync (deferred to V1.5)

### 12.3 Build Order

Ordered component checklist. Each builds on the previous.

- [ ] **Utils foundation** — `pptx_utils/` package with brand, layout, shapes, charts, lxml_helpers modules
- [ ] **Brand template** — Reverse-engineer J&J template → `BRAND{}` entry, `LAYOUTS{}` presets, template `.pptm`
- [ ] **Creation pipeline** — Generation scripts that compose from pptx_utils → client-delivery-quality slide
- [ ] **Registry** — `registry.json` written at creation time, reconciliation from live PowerPoint state
- [ ] **Live editing** — win32com edit scripts, registry update, edit log
- [ ] **Skill instructions** — `SKILL.md` with cardinal rules, workflows, worked patterns
- [ ] **Synapse integration** — `synapse_fetch.py` → `banner_plan.pkl` → generation scripts consume it
- [ ] **Connected slides + refresh** — `data_source` lineage in registry, "refresh this deck" command
- [ ] **Resilience** — VBA hook, backup strategy, undo mechanism, error handling
- [ ] **Real-world validation** — Recreate a client-delivery slide using the full system. Side-by-side comparison with manual output. Target: equivalent quality in under 30 minutes.
- [ ] **Analyst handoff** — Vinoth uses the system to generate a real deck. His experience bootstraps the skill library.

---

## 13. Open Questions

| # | Question | Context | Next Step |
|---|---|---|---|
| 1 | **Skill hierarchy** — metric type vs. layout type vs. both? | Tension between consistency (structured skills) and flexibility (Claude decides). See Section 5.3. | Rubric exercise with Vinoth: what prompts → what outputs expected. Reverse-engineer existing decks. |
| 2 | **V1 skill list** — which specific skills to build first? | Should be informed by the rubric + the most common J&J/PET slide types. | Vinoth identifies the top 5-10 slide types by frequency. |
| 3 | **Non-synapsible stored procedures** — format for J&J/PET? | Some data requires raw respondent-level analysis. Need a script format that can be re-run for refresh. See Section 3.5. | Design the stored procedure convention during Synapse integration. |

---

## Appendix A: Origin Context

SlideGen originated from an architecture exploration between Sriram (CEO) and Claude in February 2026. That conversation started from a working Python script that generated a single J&J Rybrevant slide with native PPT charts. Over ~12 hours of iteration, it:

1. Explored the `python-pptx` + `lxml` boundary in depth
2. Evaluated 5 architecture options for the editing interface
3. Landed on win32com + Claude Code as the architecture
4. Worked through resilience concerns: manual edits, copy-paste, duplicate names, registry staleness
5. Established the company skills vs. project artifacts separation

The full conversation transcript is preserved in `[3] SlideGen — Architecture Exploration (Sriram × Claude, Feb 2026).md`.

Additional context from implementation planning: `[4] SlideGen — Sriram × Siva Discussion Notes (Mar 4 2026).md`.

---

## Appendix B: Ecosystem Research

Before building `pptx_utils/`, evaluate what already exists. Core finding: **nobody has published a maintained library wrapping python-pptx's lxml gaps.** Everyone doing serious chart formatting rolls their own.

### Libraries to Evaluate

| Library | What it does | Relevance | Action |
|---|---|---|---|
| **python-pptx-ng** ([PyPI](https://pypi.org/project/python-pptx-ng/)) | Fork of python-pptx adding gradient fills, enhanced data labels, group shapes, shadow formatting, OLE embedding | May eliminate some lxml workarounds entirely | Evaluate first. If it covers bar gap width, axis orientation, or label positioning natively, use it instead of writing lxml helpers. |
| **pptxlib** ([PyPI](https://pypi.org/project/pptxlib/)) | win32com wrapper for PowerPoint (beta, Dec 2025, 0 stars) | Could provide higher-level win32com API | Evaluate. If its API is stable, it could reduce raw win32com code. Don't depend on it given immaturity. |
| **python-pptx-fix** ([PyPI](https://pypi.org/project/python-pptx-fix/)) | Patched fork fixing specific python-pptx bugs (date axis, Unicode, bubble chart labels) | Fixes known bugs without changing API | Use if you hit specific bugs it patches. Drop-in replacement. |

### Not Relevant

| Library | Why not |
|---|---|
| **presenton** (4.2k stars) | Full-stack AI presentation generator. Different problem. |
| **slide-deck-ai** (315 stars) | Streamlit app for LLM-generated slides. Content generation, not formatting control. |
| **python_pptx_interface** (44 stars) | Higher-level slide composition. Stale since 2020. |
| **Aspose.Slides** (commercial) | Full PowerPoint API without Office. Overkill — we have Office and need COM integration. |

---

## Appendix C: Pragmatic V1 Tradeoffs

These are deliberate compromises favoring speed-to-delivery over long-term scalability. Known tradeoffs, not oversights.

| Choice | Why it's not optimal | Why it's fine for V1 | When to revisit |
|---|---|---|---|
| Claude Code terminal as UI | Not analyst-friendly | Eliminates custom app. Validates logic first. | When non-technical analysts need the system without hand-holding. |
| Two tools (python-pptx + win32com) | Two mental models | python-pptx proven for creation; win32com for editing. Complementary strengths. | If a unified COM approach proves simpler after production use. |
| JSON file for registry | No concurrency, no schema validation | Single-user, single-session. Simple to debug. | Multi-user scenarios or recurring corruption. |
| Pickle for intermediate data | Not human-readable, Python-version-tied | Fast, zero-config, preserves types exactly. | Non-Python consumers or Python version fragmentation. |
| Brand constants in Python dict | Requires code change to update | Brands change infrequently. Simplest to consume. | When non-developers need to manage brand palettes. |
| No automated visual validation | Relies on analyst's eyes | Visual testing tooling is a project in itself. | High-volume generation where manual checks bottleneck. |
| No deck-level awareness | No cross-slide consistency | Most edits are slide-scoped. | Multi-slide decks needing consistency enforcement. |
| OneDrive for skill distribution | No automated versioning | Analysts never touch git. Transparent sync. | Scaling beyond a handful of users. |

---

## Appendix D: OOXML Properties Requiring Raw lxml

Properties that `python-pptx` does not expose via its API. Pre-solved in `pptx_utils/lxml_helpers.py`.

| Property | OOXML Element | Function |
|---|---|---|
| Bar gap width | `c:gapWidth` | `set_plot_area_gap()` |
| Bar overlap | `c:overlap` | `set_overlap()` |
| Axis orientation (invert) | `c:scaling > c:orientation val="maxMin"` | `invert_cat_axis()` |
| Axis tick label visibility | `c:tickLblPos val="none"` | `hide_cat_labels()` |
| Data label position | `c:dLblPos val="outEnd"` | `set_datalabel_pos_outside_end()` |
| Series border removal | `a:ln > a:noFill` | `set_series_no_border()` |
| Axis number format | `c:numFmt formatCode="0"` | `set_val_axis_number_format()` |

This table grows as new formatting needs are discovered. Each addition follows the same pattern: discover the XML, write a named function, document it here.
