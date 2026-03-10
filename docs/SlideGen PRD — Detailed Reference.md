# PRD: AI-Assisted PowerPoint Slide Generation & Editing System

**Author:** Sivakumar K
**Date:** 2026-03-01
**Status:** Draft
**Origin:** Architecture exploration conversation between Sriram and Claude, Feb 2026

---

## 1. Problem Statement

ZoomRx consulting analysts manually build PowerPoint slide decks for client deliverables. This process is time-consuming, error-prone, and requires specialized knowledge of PowerPoint formatting. Every new engagement repeats the same work — extracting survey data, building charts, aligning layout grids, applying brand formatting — with no institutional accumulation of the effort.

A proof-of-concept conducted in Feb 2026 demonstrated that Claude can generate client-delivery-quality slides programmatically using `python-pptx` and raw `lxml` for OOXML manipulation. However, the POC required ~12 hours of iterative back-and-forth, and the knowledge gained (XML element names, layout constants, formatting workarounds) would be lost in the next session.

**This system solves three problems:**

1. **Creation** — Generate slides programmatically from structured survey data, at client-delivery quality, without manually placing shapes and charts in PowerPoint.
2. **Editing** — Allow analysts to make natural language surgical edits to generated slides, with changes reflected live in PowerPoint's own canvas.
3. **Knowledge retention** — Capture all hard-won OOXML knowledge, layout patterns, and client brand specs in a reusable skill + resource library that accumulates value across engagements.

---

## 2. Users

**Primary user:** Consulting analysts at ZoomRx who build client deliverable slide decks.

- All analysts are on **Windows** with **desktop PowerPoint**
- Analysts are not developers but are comfortable using terminal-based tools with guidance
- Analysts currently spend significant time on slide formatting, chart construction, and layout alignment
- Analysts may make manual edits to generated slides in PowerPoint (drag shapes, change text, adjust colors) between AI-assisted edit sessions

---

## 3. Phased Build Plan

This system should be built like layers of an onion — each phase validates a specific idea before the next phase builds on top. **Do not skip phases.** If a phase reveals that an assumption is wrong, stop and reassess before proceeding. The whole point is to fail fast and cheap on bad assumptions.

### Phase 0: Validate the core premise — Can win32com edit a live PowerPoint from Claude Code?

**What this proves:** The entire architecture rests on one bet — that Python, running from Claude Code's bash tool, can connect to a running PowerPoint instance via COM and make live edits that the analyst sees instantly. If this doesn't work, everything else is moot.

**Deliverable:** A single Python script, run from Claude Code, that:
1. Connects to a running PowerPoint instance (`win32com.client.Dispatch("PowerPoint.Application")`)
2. Gets the active presentation and first slide
3. Lists all shapes on the slide (name, type, position)
4. Modifies one text shape's content and color
5. Modifies one shape's position
6. The analyst sees both changes appear on the PowerPoint canvas without saving/reopening

**What is NOT in this phase:**
- No registry, no skill library, no `pptx_utils.py`
- No shape naming convention
- No Claude Code skill/CLAUDE.md setup
- No data pipeline
- The script is hardcoded and disposable — this is a spike, not production code

**Success criteria:**
- Analyst has PowerPoint open with any slide
- Analyst runs the script from Claude Code
- Changes appear live on the canvas
- No file corruption, no crash, no permission errors

**Key risks to watch:**
- Does `pywin32` install cleanly on the analyst's machine?
- Does COM dispatch work when Claude Code's Python subprocess is the caller?
- Are there UAC or antivirus issues blocking inter-process COM communication?
- Does PowerPoint's COM interface behave the same across Microsoft 365 versions?

**Also evaluate in this phase:**
- Install `pptxlib` (beta win32com wrapper) and test whether its API is usable. If it provides cleaner wrappers than raw win32com, consider using it. If it's too immature, proceed with raw win32com. (See Appendix B for details.)

---

### Phase 1: Round-trip — Create a slide with python-pptx, then edit it live with win32com

**What this proves:** The creation-to-editing handoff works. A slide built by `python-pptx` (file on disk) can be opened in PowerPoint and then surgically edited via `win32com` through Claude Code.

**Deliverable:**
1. A simple generation script that creates a slide with `python-pptx`: one chart, one text box, one rectangle — with each shape explicitly named (e.g., `zrx_001`, `zrx_002`, `zrx_003`)
2. Analyst opens the generated file in PowerPoint
3. An edit script (run from Claude Code) finds shapes by name via COM and modifies them
4. Changes appear live

**What is NOT in this phase:**
- No registry file — shape names are hardcoded in the edit script
- No reconciliation from live state
- No skill library — scripts are standalone

**Success criteria:**
- Shapes created by `python-pptx` are findable by name via `win32com`
- Text, color, position, and chart formatting edits all work via COM
- The generated pptx does not corrupt after COM edits + save

**New assumption being tested:** Do `python-pptx`-assigned shape names survive and remain accessible via COM? (They should — both work with the same underlying OOXML — but this needs to be confirmed.)

**Also evaluate in this phase:**
- Test `python-pptx-ng` (fork with extended features) instead of vanilla `python-pptx` for the creation step. If it natively supports bar gap width, axis orientation, or data label positioning, those operations don't need lxml workarounds at all. (See Appendix B for details.)

---

### Phase 2: Registry and reconciliation — Tolerate manual edits

**What this proves:** The system can maintain awareness of slide state across manual analyst edits. Claude always works from the truth, not from stale assumptions.

**Deliverable:**
1. `slide_registry.json` — written at creation time with shape names, labels, and initial property values
2. `reconcile_registry.py` — connects to PowerPoint, reads live state of all `zrx_*` shapes, updates the registry
3. Duplicate name detection with hard stop
4. Missing/unregistered shape warnings

**Test scenario:**
1. Generate a slide (Phase 1 script)
2. Open in PowerPoint
3. Manually drag a shape to a new position, change a font color
4. Run `reconcile_registry.py` from Claude Code
5. Verify: registry now reflects the manual changes
6. Run an edit script that uses registry values — it should operate on the actual current state, not the original creation state

**Success criteria:**
- Manual edits are absorbed transparently
- Duplicate names block the edit cycle (hard stop works)
- Missing shapes produce a warning but don't crash

---

### Phase 3: Skill library — Claude generates correct scripts without inline lxml

**What this proves:** The reusable skill architecture works. Claude reads `SKILL.md` and `pptx_utils.py`, and generates edit/creation scripts by composition rather than rediscovery.

**Deliverable:**
1. Extract helper functions from the Feb 2026 POC into `pptx_utils.py` (lxml helpers, layout primitives, chart builders)
2. Write `SKILL.md` with workflow rules, cardinal rules, and worked examples
3. Set up `CLAUDE.md` at project level with client context

**Test scenario:**
1. Open Claude Code in the project directory
2. Ask Claude to generate a slide (without giving it any code — it should compose from `pptx_utils.py`)
3. Ask Claude to make a live edit (it should use `pptx_utils` + `win32com`, not write raw lxml)
4. Verify: Claude never writes inline lxml or hardcodes brand colors

**Success criteria:**
- Claude uses skill functions correctly
- Generated scripts are short (< 30 lines of composition, not 200 lines of low-level code)
- The same quality as the 12-hour POC is achievable in minutes because the knowledge is pre-captured

---

### Phase 4: Real-world slide at client-delivery quality

**What this proves:** The system produces output good enough for actual client delivery. Not a toy demo — a real slide with real data, real brand formatting, real chart configurations.

**Deliverable:**
1. Recreate the Rybrevant MR+ME two-chart-with-delta slide from the Feb 2026 POC using the skill library
2. Same level of formatting fidelity: native charts, correct colors, precise alignment, delta tables with green/red formatting
3. Live editing session to fine-tune positioning and formatting

**Success criteria:**
- Side-by-side comparison with the POC output shows equivalent or better quality
- The creation + editing cycle completes in under 30 minutes (vs. ~12 hours for the POC)
- An analyst (not the developer) can drive the editing session

---

### Phase 5: Synapse data integration

**What this proves:** The data pipeline works end-to-end — from Synapse API to finished slide.

**Deliverable:**
1. `synapse_fetch.py` — API client that pulls cross-tab data and writes standardized pkl
2. Integration with generation scripts — Claude fetches data, generates slides, all in one session
3. Offline resilience — generation works from cached pkl when API is unavailable

**Prerequisite:** Synapse API access, test survey data available.

**Success criteria:**
- Analyst says "fetch Q4 data for [client]" and Claude does it
- Data flows to slides without manual Excel wrangling
- Cached pkl enables offline work

---

### Phase 6: Resilience and polish

**What this proves:** The system is robust enough for daily use by multiple analysts.

**Deliverable:**
1. VBA BeforeSave hook for auto-deduplicating shape names
2. Edit log (`edit_log.json`) with undo capability
3. Versioned backups before each edit session
4. Error handling for common failure modes (COM disconnection, missing shapes, API timeouts)

**Success criteria:**
- Copy-paste of shapes doesn't break the system
- Analyst can say "undo the last change" and it works
- No data loss from any realistic analyst workflow

---

### Phase summary

| Phase | Validates | Depends on | Key risk tested |
|---|---|---|---|
| 0 | win32com + Claude Code works at all | Nothing | COM from subprocess, UAC/AV, PPT version compat |
| 1 | python-pptx → win32com handoff | Phase 0 | Shape names survive across tools |
| 2 | Manual edit tolerance | Phase 1 | Reconciliation correctness, duplicate detection |
| 3 | Skill library eliminates rediscovery | Phase 1 | Claude's ability to compose from pre-built functions |
| 4 | Client-delivery quality | Phases 2 + 3 | Formatting fidelity, end-to-end workflow speed |
| 5 | Data pipeline | Phase 4 | Synapse API integration, pkl schema design |
| 6 | Production resilience | Phase 4 | Edge cases, undo, error recovery |

**Critical gate:** If Phase 0 fails, the entire win32com approach is invalidated. Fall back to Option 4 from the original conversation (web app with python-pptx file editing + image preview). The skill library and registry concepts (Phases 2–6) still apply — only the live-editing mechanism changes.

---

## 4. System Architecture Overview

The system has four layers, with a clear separation between company-level reusable assets and per-engagement project artifacts.

```
COMPANY SKILLS (central, version-controlled, read-only by projects)
├── SKILL.md                 — Workflow rules for slide generation & editing
├── pptx_utils.py            — Pre-solved lxml helpers, shape primitives, chart builders
├── BRAND{}                  — Color palettes, fonts per client (from brand guides)
├── LAYOUTS{}                — Standard grid presets for common slide compositions
├── CHART_PATTERNS{}         — Clustered bar, stacked bar, waterfall, abacus, etc.
├── SLIDE_RULES.md           — Accumulated formatting rules & pitfalls
└── synapse_fetch.py         — API client for Synapse data ingestion

PROJECT ARTIFACTS (per-engagement, lives in project folder)
├── CLAUDE.md                — Project-specific context (client, survey IDs, template path)
├── slide_registry.json      — Current shape state (live-reconciled from PowerPoint)
├── edit_log.json            — Audit trail of all edits
├── slide_data.pkl           — Extracted/fetched survey data snapshot
├── *.pptx                   — The slide files
└── generation scripts       — Per-slide generation code

RUNTIME (analyst's local machine)
├── Claude Code              — Front-end interface (terminal)
├── Python + win32com        — Execution engine, COM bridge to PowerPoint
└── PowerPoint.exe           — Live canvas, file owner

CLOUD
└── Claude API               — Natural language → edit script translation
```

**Front end:** Claude Code running in a terminal on the analyst's machine. No custom desktop app, no Flask server, no MCP server. Claude Code provides the agentic loop (plan → execute → verify → confirm) natively.

**Execution:** Claude Code writes targeted Python scripts and runs them via bash. Scripts use `win32com` to talk directly to the running PowerPoint instance through Windows COM (local inter-process communication, no network involved).

**State management:** `slide_registry.json` is the shared state between Claude and PowerPoint. It is always reconciled from PowerPoint's live state before any edit, never trusted from a previous session.

---

## 4. Assumptions & Pragmatic Tradeoffs

This section documents design choices that are deliberately **not optimal** but are **good enough to test the waters**. Each is a conscious tradeoff favoring speed-to-validation over long-term scalability. The developer should understand these are known compromises, not oversights.

### 4.A Claude Code terminal as the UI

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Interface | Claude Code in a terminal window | Purpose-built desktop app with confirm/cancel buttons, visual preview pane, edit history sidebar |
| Why V1 is acceptable | Eliminates building any custom UI. Claude Code provides file access, Python execution, and agentic reasoning natively. Gets us to validation in days, not weeks. |
| When to revisit | When non-technical analysts need to use the system without hand-holding. The terminal will be a barrier at that point. |
| What carries forward | Everything behind the terminal — `pptx_utils.py`, registry pattern, win32com scripts — becomes the backend that a future UI wraps. Nothing is thrown away. |

### 4.B Single `pptx_utils.py` file for all utilities

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Structure | One Python file with all functions | Modular package: `pptx_utils/charts.py`, `pptx_utils/lxml_helpers.py`, `pptx_utils/win32com_wrappers.py`, `pptx_utils/layouts.py` |
| Why V1 is acceptable | One file to import, one file to read into Claude's context, one file to maintain. At the current scale (< 30 functions), a single file is navigable. |
| When to revisit | When the file exceeds ~50 functions or when multiple developers are editing it concurrently and hitting merge conflicts. |

### 4.C JSON file on disk for registry

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Storage | `slide_registry.json` — plain file, no schema validation, no concurrency protection | SQLite database or JSON with schema enforcement (JSON Schema / Pydantic) |
| Why V1 is acceptable | Single-user, single-session usage. No concurrent writes. Claude Code reads and writes it sequentially. Simple to debug — just open the file. |
| Risk accepted | File can be accidentally deleted, corrupted by a partial write, or edited manually in a way that breaks the schema. These are low-probability in a single-user workflow. |
| When to revisit | Multi-user scenarios, or if registry corruption becomes a recurring issue. |

### 4.D Pickle (`.pkl`) for intermediate data

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Format | Python pickle — binary, not human-readable, Python-version-sensitive | Parquet (columnar, language-agnostic) or structured JSON (human-readable, inspectable) |
| Why V1 is acceptable | The Feb 2026 POC already used pkl. Zero-config, fast serialization, preserves Python data types exactly. Works well within a single-Python-version environment. |
| Risk accepted | Cannot inspect data without Python. Pickle files from one Python version may not load on another. Not shareable with non-Python tools. |
| When to revisit | If the data pipeline needs to serve non-Python consumers (dashboards, other tools) or if Python version fragmentation becomes an issue across analyst machines. |

### 4.E VBA macro for BeforeSave hook

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Mechanism | VBA macro embedded in the template file (`.pptm`) | Python-side pre-save validation step in Claude Code's workflow, or a lightweight COM event listener |
| Why V1 is acceptable | Zero-dependency — works without Python running, fires automatically on every save, travels with the file. No installation required. |
| Risk accepted | Requires `.pptm` (macro-enabled format). Some organizations block macros. The VBA code is not version-controlled with the skill library — it lives inside the template file. |
| When to revisit | If client-facing files cannot contain macros, or if the hook needs to do more than name deduplication (e.g., registry sync, Synapse push). |

### 4.F Brand constants hardcoded in Python dict

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Storage | `BRAND = { "JJ": { ... } }` inside `pptx_utils.py` | External config file (JSON/YAML) or database-backed brand management system |
| Why V1 is acceptable | Brand palettes change infrequently. The number of clients is manageable (< 20). A Python dict is the simplest structure to consume in generation scripts. |
| When to revisit | When brand updates need to happen without touching code, or when a non-developer (e.g., design team) needs to manage brand palettes. |

### 4.G Two tools for creation vs. editing (python-pptx + win32com)

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Approach | `python-pptx` for building slides from scratch, `win32com` for live surgical edits | Potentially unify on `win32com` for everything — it can also create slides, and its API coverage is broader |
| Why V1 is acceptable | `python-pptx` is proven for creation (from the Feb 2026 POC). It doesn't require PowerPoint to be running, which simplifies the creation pipeline. Whether `win32com` works equally well for creation is unvalidated — and this PRD is about validating `win32com` for editing first. |
| Risk accepted | Developers must hold two mental models. Some operations (chart data replacement) fall awkwardly between the two tools. |
| When to revisit | After Phase 2 validates that `win32com` works reliably for editing. At that point, evaluate whether creation should also move to COM for a unified approach. |

### 4.H No automated visual validation

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Validation | Analyst opens the slide in PowerPoint and visually confirms it looks right | Automated screenshot comparison against reference images, visual regression testing |
| Why V1 is acceptable | Visual validation tooling is a project in itself. The analyst is looking at the real PowerPoint canvas, so the feedback loop is fast even if manual. |
| When to revisit | When slide generation becomes high-volume enough that manual visual checks are a bottleneck, or when regressions start slipping through. |

### 4.I No multi-slide / deck-level awareness

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Scope | Registry is per-slide. No concept of the deck as a coherent whole. | Deck-level registry tracking all slides, cross-slide consistency (consistent colors, shared legends, sequential numbering), and slide ordering. |
| Why V1 is acceptable | Most edits are slide-scoped. The Feb 2026 POC worked on individual slides. Deck-level concerns (table of contents, slide numbering, cross-references) can be deferred. |
| When to revisit | When analysts start generating multi-slide decks and need consistency enforcement across slides. |

### 4.J Shape naming by convention (`zrx_` prefix) rather than PowerPoint internal IDs

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Identification | Custom `zrx_NNN` names assigned at creation time | PowerPoint's internal shape IDs (unique, immutable, assigned by PowerPoint itself) |
| Why V1 is acceptable | `zrx_` names are human-readable and easy to debug. Internal shape IDs are opaque integers that are harder to work with during development. Convention-based naming is simpler to implement. |
| Risk accepted | If an analyst renames a `zrx_*` shape, the system loses track of it. Internal IDs would not have this problem since they cannot be changed by users. |
| When to revisit | If shape renaming becomes a recurring issue, or when the system needs to track shapes it didn't create. |

### 4.K Static utility library vs. letting Claude re-derive lxml each time

This is the most important assumption to pressure-test. The PRD positions `pptx_utils.py` as "the most critical artifact in the system." But Claude's training data already contains OOXML knowledge — it figured out the raw lxml strings in the Feb 2026 POC without a utility library.

| Aspect | V1 (pragmatic) | Optimal (future) |
|---|---|---|
| Approach | Thin `pptx_utils.py` with the ~10 most-used lxml workarounds + clear `SKILL.md` instructions that teach Claude the *pattern* for discovering new XML | Either: (a) comprehensive library covering all known OOXML gaps, or (b) no library at all — rely entirely on Claude + instructions + verification |
| Why V1 is acceptable | The static library gives you tested, auditable code for the operations that come up on every engagement (bar gap, axis inversion, label position). For rare/one-off formatting needs, Claude can still generate inline lxml guided by the skill instructions. You get the safety of tested code for common operations and the flexibility of Claude for the long tail. |
| Risk accepted | The library will never be exhaustive. New formatting needs will arise that aren't covered. The skill must explicitly tell Claude: "If the needed operation isn't in `pptx_utils.py`, you may write inline lxml, but flag it for later extraction into the library." |
| What the ecosystem tells us | Nobody has published a maintained pptx-utils-style library (see Appendix C). Everyone rolls their own. This is not because the idea is bad — it's because the long tail of formatting needs is enormous and varies by use case. A company-specific library (not a general-purpose one) is the right scope. |
| When to revisit | After Phase 4 (real-world slide), audit which lxml operations Claude used. If it's mostly calling `pptx_utils` functions → the library is earning its keep. If Claude is frequently writing inline lxml for uncovered cases → the library is too thin and needs expansion, or the instructions-only approach may be better. |

**Also evaluate before building:**
- **python-pptx-ng** (fork, Jan 2024) — adds gradient fills, enhanced data labels, group shapes. May eliminate some lxml workarounds entirely. Test whether it covers enough gaps to reduce the utility library's scope.
- **pptxlib** (Dec 2025, beta, 0 stars) — a win32com wrapper for PowerPoint. May be useful for Phase 0/1 as a starting point instead of raw win32com. Evaluate but don't depend on it given its immaturity.

---

## 5. Components

### 5.1 Company Skill Library

**Purpose:** Encode all reusable PowerPoint generation knowledge so it is discovered once and never re-derived.

**5.1.1 `pptx_utils.py` — The Utility Library**

A single Python file containing all pre-solved functions. This is the most critical artifact in the system. It grows over time but never loses old functions.

**Categories of functions:**

| Category | Examples | Notes |
|---|---|---|
| lxml XML manipulation | `invert_cat_axis()`, `hide_cat_labels()`, `set_datalabel_pos_outside_end()`, `set_series_no_border()`, `set_val_axis_number_format()`, `set_plot_area_gap()`, `set_overlap()` | These encode OOXML XML element names and attribute values that `python-pptx` does not expose. The hardest knowledge to acquire — found by reverse-engineering PowerPoint's XML output or reading the ECMA-376 spec. |
| Layout primitives | `textbox()`, `solidrect()`, `horiz_line()`, `add_delta_col()` | Reusable shape placement functions with consistent parameter conventions (inches for position/size). |
| Chart builders | `make_clustered_bar()`, `make_stacked_bar()`, `make_waterfall()` | High-level chart construction that internally handles `ChartData`, series formatting, axis configuration, and all necessary lxml workarounds. |
| Slide scaffolding | `open_template()`, `add_blank_slide()`, `save_slide()` | Template handling, slide stripping, save behavior. |
| win32com wrappers | `connect_to_ppt()`, `get_active_slide()`, `find_shape_by_name()`, `read_shape_properties()` | Wrappers around COM calls for use during the live editing phase. |
| Registry helpers | `reconcile_registry()`, `update_registry()`, `register_shape()`, `unregister_shape()` | Shape inventory management and live state synchronization. |

**Design rules for `pptx_utils.py`:**
- All lxml XML manipulation lives here and nowhere else
- Claude never writes inline lxml in generated scripts — it calls functions from this library
- Every function that manipulates raw XML includes a docstring naming the OOXML element it targets and why `python-pptx` doesn't cover it
- New formatting needs discovered during engagements are added here as named functions, not left inline in project scripts

**5.1.2 `BRAND{}` — Client Brand Definitions**

Structured color palettes and font specs derived from client brand guides. Keyed by client identifier.

```python
BRAND = {
    "JJ": {
        "primary":      RGBColor(0xFF, 0x00, 0x00),
        "q_current":    RGBColor(0xF7, 0x58, 0x24),
        "q_previous":   RGBColor(0xFF, 0xC1, 0x99),
        "positive":     RGBColor(0x00, 0xB0, 0x50),
        "negative":     RGBColor(0xFF, 0x00, 0x00),
        "neutral_grey": RGBColor(0x50, 0x50, 0x50),
        "font_heading": "Calibri",
        "font_body":    "Calibri",
        "template_path": "templates/jj_q4_2025.pptx"
    },
    # Additional clients added as engagements onboard
}
```

**Growth model:** New clients are added to `BRAND{}` as their first engagement uses the system. Existing entries are updated when brand guides change.

**5.1.3 `LAYOUTS{}` — Standard Grid Presets**

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

These are starting points. Individual slides may adjust values, but the layout dictionary provides the baseline so Claude never guesses coordinates from scratch.

**5.1.4 `SKILL.md` — Workflow Instructions**

The skill file that Claude Code reads at session start. Contains:

- **Cardinal rules** — Never write raw lxml, never hardcode brand colors inline, always open a template, always reconcile before editing
- **Standard workflow** — The step-by-step sequence for creation and editing
- **Decision rules** — When to recreate vs. edit in-place, when to ask for clarification vs. act immediately
- **Worked patterns** — Code examples for common slide types showing how to compose from `pptx_utils` functions
- **Required inputs checklist** — What Claude must confirm before generating code (data source, client/brand, slide type, metrics, time periods, sort order, template path, output path)
- **Known limitations** — What `python-pptx` can and cannot do, with pointers to the `pptx_utils` function that solves each gap

**5.1.5 `synapse_fetch.py` — Synapse API Client**

API client for pulling cross-tab survey data from ZoomRx's internal Synapse system.

- **Auth:** API key/token based. Key stored in environment variable or local config file (never committed to version control).
- **Input:** Survey IDs, wave/quarter identifiers, metric types — provided via project-level `CLAUDE.md`
- **Output:** Writes `slide_data.pkl` in a standardized schema that all generation scripts consume
- **Offline resilience:** The pkl is a timestamped snapshot. Slide generation works from pkl, not live API calls, enabling offline work and reproducibility.

**Data flow:**
```
Synapse API → synapse_fetch.py → slide_data.pkl → generation scripts → .pptx
```

The pkl decouples slide generation from API availability. It also serves as an audit artifact — you can verify exactly what data produced each version of the deck.

---

### 5.2 Slide Creation Pipeline

**Trigger:** Analyst opens Claude Code in the project directory and describes the slide(s) needed.

**Sequence:**

1. Claude Code reads `SKILL.md` and `pptx_utils.py` from company skills
2. Claude Code reads project `CLAUDE.md` for client context (brand, template, survey IDs)
3. Claude asks the required inputs checklist (data source, metrics, slide type, etc.)
4. If data not yet fetched: Claude runs `synapse_fetch.py` with survey IDs from `CLAUDE.md`, produces `slide_data.pkl`
5. Claude writes a generation script that:
   - Opens the client template via `open_template()`
   - Loads data from pkl
   - Composes the slide using `pptx_utils` functions (chart builders, layout primitives, text boxes)
   - Names every shape with a `zrx_` prefixed ID (e.g., `zrx_001`, `zrx_002`)
   - Writes `slide_registry.json` with shape ID → semantic label mapping
   - Initializes empty `edit_log.json`
   - Saves the pptx
6. Claude executes the script via bash
7. Claude confirms what was created and tells the analyst to open the file in PowerPoint

**Shape naming convention:**

All system-created shapes get a `zrx_` prefixed opaque ID. Semantic meaning lives in the registry, not in the shape name. This separates identity (stable, used for lookup) from description (can change without breaking anything).

```json
{
  "shapes": {
    "zrx_001": {
      "label": "MR clustered bar chart",
      "type": "chart",
      "metric": "message_recall",
      "position": "left"
    },
    "zrx_002": {
      "label": "Slide headline",
      "type": "textbox",
      "role": "headline"
    }
  }
}
```

The `zrx_` prefix also lets the system distinguish managed shapes from manually-created ones (e.g., `Rectangle 7`). The reconcile step only tracks `zrx_*` shapes; everything else is ignored.

---

### 5.3 Live Editing Pipeline (win32com + Registry Reconciliation)

**Trigger:** Analyst has the slide open in PowerPoint and describes an edit in Claude Code.

**Prerequisite:** PowerPoint is running with the target file open. Claude Code is open in a terminal on the same machine.

**Core design principle:** PowerPoint is always the source of truth. The registry (`slide_registry.json`) is Claude's cache of what's on the slide — but it is never trusted from a previous session. Before every edit, the system reads live state from PowerPoint and updates the registry. This means manual edits the analyst makes in PowerPoint are automatically absorbed.

**Step 0 — Reconcile (runs before every edit, no exceptions):**

Claude runs `reconcile_registry.py` which connects to PowerPoint via `win32com.client.Dispatch("PowerPoint.Application")` and performs a full inventory:

```
Input:  Running PowerPoint instance + existing slide_registry.json
Output: Updated slide_registry.json with live values

Steps:
1. Connect to PowerPoint via win32com
2. Build name → shape(s) map for ALL shapes on the active slide
3. DUPLICATE DETECTION: If any zrx_* name appears more than once → HARD STOP
   - Print which names are duplicated
   - Exit with error — analyst must rename in PowerPoint first
   - Why: two shapes sharing a name means any edit silently targets
     whichever COM enumerates first. Wrong results with no error.
     A hard stop is safer than a wrong edit.
4. For each registered shape: read live left, top, width, height,
   font properties, text content from PowerPoint
5. FLAG shapes in registry but missing from slide (deleted manually)
6. FLAG unregistered zrx_* shapes on slide (unexpected — copy-paste?)
7. IGNORE non-zrx_* shapes entirely (manually created, not system-managed)
8. Write updated registry to disk
```

**Steps 1–8 — The edit cycle (after reconciliation):**

1. Claude reads the reconciled registry to understand current state
2. Claude identifies the target shape from the analyst's natural language description using registry labels
3. **For positional changes:** Claude describes the planned change and asks for confirmation before executing
4. **For clearly scoped changes** (text, color, font size): Claude executes immediately
5. Claude writes a minimal edit script using `pptx_utils` functions and `win32com` wrappers
6. Script executes → COM message sent → PowerPoint canvas updates live on the analyst's screen
7. Claude updates `slide_registry.json` and appends to `edit_log.json`
8. Claude confirms the change in plain English

**Edit types and their code patterns:**

| Edit type | Example NL | Code pattern |
|---|---|---|
| Text content | "Change the headline to say Q4 results" | `shape.TextFrame.TextRange.Text = "..."` |
| Font size | "Make the footer smaller" | `shape.TextFrame.TextRange.Font.Size = 5.5` |
| Color | "Make the headline dark grey" | `shape.TextFrame.TextRange.Font.Color.RGB = 0x505050` |
| Position | "Move the delta table right a bit" | `shape.Left = Inches(current + 0.15) * 914400` |
| Chart data | "Update with Q1 2026 numbers" | `chart.replace_data(new_chart_data)` via `python-pptx` (file-based, not COM — see note below) |
| Bar gap width | "Make the bars fatter" | `chart_obj.ChartGroups(1).GapWidth = 50` via COM |
| Axis inversion | "Flip the bar order" | `chart_obj.Axes(2).ReversePlotOrder = True` via COM |

**Important note on chart data replacement:**

`win32com` can modify chart formatting properties directly on the live chart. However, for replacing the entire data set of a chart (e.g., adding a new quarter), it may be more reliable to use `python-pptx`'s `chart.replace_data()` on the file. This requires the file to be saved and closed first, then reopened. The skill should document when to use COM (live formatting edits) vs. file-based `python-pptx` (structural data changes).

**Ambiguity resolution:**

When the analyst's instruction is ambiguous, Claude must ask before acting:
- "Make the text bigger" → Which text? There may be 12 textboxes.
- "Shift it right a bit" → What is "it"? What is "a bit" (0.1"? 0.25")?
- "The colors look off" → Too vague to act on.

The skill should instruct Claude to confirm before executing positional changes and to execute immediately only for clearly scoped text/color/font changes.

---

### 5.4 VBA BeforeSave Hook

**Purpose:** Automatically fix duplicate shape names on every save, so the analyst never has to think about naming discipline.

**Implementation:** A VBA macro embedded in the client template file (`.pptm` or macro-enabled `.pptx`). Fires on every `Ctrl+S`.

**Behavior:**
1. Scan all shapes on the current slide
2. For any `zrx_*` name that appears more than once: rename the duplicate to `{original_name}_{counter}` (e.g., `zrx_001_2`)
3. Log any new non-`zrx_*` shapes as unregistered (informational, no action)
4. Allow save to proceed

**Why VBA and not Python:** The hook lives inside the template file itself. It works for any analyst who opens the file, requires no Python process running in the background, and has no installation dependency. It is the simplest possible implementation for V1.

**Extension point:** In later versions, this hook could also trigger a registry snapshot or push metadata to a central tracking system.

---

### 5.5 Edit Log

**Purpose:** Audit trail of all AI-assisted edits, enabling undo and session review.

**Schema:**
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

**Undo:** Reverting an edit means re-applying the `before` values from the log entry. Claude can do this when asked ("undo the last change" or "revert edit_001").

**Versioned backups:** Before each edit session, a timestamped copy of the pptx is saved as a backup. This is a safety net — more robust than per-property undo for cases where multiple edits interact.

---

## 6. Data Flow

```
                    ┌──────────────┐
                    │  Synapse API │
                    │  (survey data)│
                    └──────┬───────┘
                           │ API key auth
                           ▼
                    synapse_fetch.py
                    (company skill)
                           │
                           ▼
                    slide_data.pkl
                    (project artifact)
                           │
              ┌────────────┴────────────┐
              │                         │
              ▼                         ▼
    CREATION FLOW                 EDITING FLOW
              │                         │
    generation script             reconcile_registry.py
    (uses pptx_utils)            (reads live PPT state)
              │                         │
              ▼                         ▼
    python-pptx writes            slide_registry.json
    .pptx to disk                 (updated with live values)
              │                         │
              ▼                         ▼
    Analyst opens in PPT          Claude writes edit script
                                  (uses pptx_utils + win32com)
                                        │
                                        ▼
                                  win32com → PowerPoint.exe
                                  (live canvas update via COM)
                                        │
                                        ▼
                                  Registry + edit_log updated
```

---

## 7. Key Design Decisions

### 7.1 Claude Code as the front end (no custom app)

**Decision:** Use Claude Code in a terminal as the analyst-facing interface for V1.

**Rationale:** Eliminates building a custom Electron/Tauri desktop app, a Flask localhost server, or an MCP server. Claude Code provides file reading, Python execution, and the agentic reasoning loop natively. Skills and CLAUDE.md files provide all necessary configuration.

**Tradeoff:** Terminal UX is not analyst-friendly for non-technical users. Acceptable for V1 with guided onboarding; becomes a limitation at scale.

**V2 path:** The same `pptx_utils.py` and registry pattern can later be wrapped with a thin desktop UI. V1 validates the logic; V2 polishes the interface.

### 7.2 win32com over python-pptx for live editing

**Decision:** Use `win32com` (Windows COM automation) for the editing phase rather than file-based `python-pptx`.

**Rationale:** COM connects to the running PowerPoint process and edits the live document. Changes appear on the analyst's screen immediately. No file lock conflicts. No save/close/reopen cycle. No approximate image preview — the analyst sees the real PowerPoint canvas at all times.

**Tradeoff:** Windows-only. Acceptable because the entire analyst team is on Windows.

**Coverage advantage:** COM exposes virtually the entire PowerPoint object model, including properties that `python-pptx` lacks (bar gap width, axis orientation, etc.) — eliminating some of the raw lxml workarounds needed during creation.

### 7.3 python-pptx for creation, win32com for editing

**Decision:** Use `python-pptx` (with lxml helpers) for initial slide creation, and `win32com` for subsequent live edits.

**Rationale:** `python-pptx` is better suited for building slides from scratch — it works with files on disk, doesn't require PowerPoint to be running, and has a clean Pythonic API for construction. `win32com` is better for surgical edits to an already-open document. The two tools complement each other.

### 7.4 Non-semantic shape IDs with semantic registry

**Decision:** Shape names in PowerPoint are opaque IDs (`zrx_001`), not semantic labels (`MR_Chart`). Semantic meaning lives in `slide_registry.json`.

**Rationale:** Separates identity from description. IDs are stable — they survive when a shape is repurposed ("actually make this the delta table now"). The `zrx_` prefix lets the system distinguish managed shapes from manually-created ones. The VBA deduplication hook can operate on IDs without understanding semantics.

### 7.5 Registry reconciliation before every edit

**Decision:** Always read live state from PowerPoint before generating any edit script. Never trust stale registry values.

**Rationale:** Analysts will make manual edits in PowerPoint between AI sessions. The system must tolerate this gracefully. Reconciliation absorbs manual changes automatically — the only analyst discipline required is "don't rename `zrx_*` shapes."

### 7.6 Company skills vs. project artifacts separation

**Decision:** Reusable knowledge (utils, brands, layouts, skills) is company-level and read-only by projects. Per-engagement state (registry, edit log, data, pptx) is project-level.

**Rationale:** Company skills accumulate value across all engagements. They are version-controlled centrally and deliberately updated. Project artifacts live and die with the engagement. Mixing them causes ownership confusion and makes the skill library fragile to project-specific changes.

### 7.7 pkl as intermediate data format

**Decision:** Survey data is always written to a pkl file before slide generation, even when sourced from an API.

**Rationale:** Decouples slide generation from API availability (offline work). Provides a timestamped audit snapshot of exactly what data produced each version. Enables re-running generation scripts without re-fetching. The pkl schema is standardized across all projects.

---

## 8. Technical Requirements

### 8.1 Analyst machine prerequisites

| Requirement | Details |
|---|---|
| OS | Windows 10/11 |
| PowerPoint | Desktop version (Microsoft 365 or standalone) |
| Python | 3.10+ |
| Python packages | `pywin32`, `python-pptx`, `lxml`, `openpyxl` (for data extraction fallback) |
| Claude Code | Installed and authenticated |
| Network | Access to Claude API (HTTPS outbound), access to Synapse API (internal network) |

### 8.2 File format considerations

- **Templates:** Must be `.pptm` (macro-enabled) to support the VBA BeforeSave hook. If macro-enabled templates are not acceptable in client-facing files, the VBA hook can be stripped on final export and the generated `.pptx` delivered without macros.
- **EMU precision:** All position/size values stored internally as EMUs (1 inch = 914,400 EMU). The `pptx_utils` API accepts inches and converts internally.
- **OOXML fidelity:** Generated files are indistinguishable from manually-created PowerPoint files. Native charts include embedded Excel workbooks. All content is editable in PowerPoint — nothing is flattened to images.

### 8.3 Synapse integration

| Aspect | Details |
|---|---|
| Auth | API key/token, stored in environment variable `SYNAPSE_API_KEY` |
| Data format | Cross-tab survey results: categories, series, values per wave/quarter |
| pkl schema | Standardized dict structure: `{slide_id: {metric: {category: {wave: value}}}}` — exact schema to be defined during implementation |
| Error handling | API failures logged and surfaced to analyst; generation proceeds from cached pkl if available |

---

## 9. Risks & Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| **COM connection failure** — PowerPoint not running or not responding to COM | Edit session cannot start | Reconcile script detects this upfront and tells analyst to open PowerPoint first. Clear error message, not a silent failure. |
| **Duplicate shape names after copy-paste** | Wrong shape gets edited silently | VBA BeforeSave hook auto-renames duplicates. Reconcile script hard-stops on duplicates as a second safety net. |
| **Analyst renames a `zrx_*` shape** | Reconcile can't find the shape, registry becomes stale | Document as the one rule: "don't rename `zrx_*` shapes." Reconcile detects missing shapes and warns. Low probability — analysts rarely rename shapes manually. |
| **python-pptx lxml workaround breaks on new PowerPoint version** | Chart formatting incorrect or file corrupted | All lxml manipulation is centralized in `pptx_utils.py`. One fix propagates to all projects. OOXML spec is stable — changes are rare. |
| **Template file corruption from malformed XML** | PowerPoint shows "couldn't read some content" repair dialog | This happened in the POC (the `c:manLayout` corruption). Mitigated by keeping all XML manipulation in tested `pptx_utils` functions, never inline. Generation scripts should be tested against PowerPoint's repair checker. |
| **Synapse API unavailable** | Can't fetch fresh data | pkl caching ensures generation works offline from last-fetched data. Analyst is informed they're working from cached data. |
| **Claude generates incorrect edit script** | Wrong shape modified or formatting broken | Versioned backup before each edit session. Edit log enables undo. Confirm-before-execute for positional changes. |
| **Skill library grows unwieldy** | Hard to maintain, functions conflict | Clear ownership (one maintainer), semantic versioning, documented deprecation process. Each function has a docstring explaining the OOXML element it targets. |

---

## 10. Success Criteria

1. **An analyst can describe a new slide in natural language and get a client-delivery-quality pptx file** — with native charts, correct brand formatting, precise layout — without writing any code or touching PowerPoint during creation.

2. **An analyst can make surgical edits via natural language** ("make the headline smaller", "move the delta table right 0.2 inches") **and see the change appear live on their PowerPoint canvas** within seconds.

3. **Manual edits in PowerPoint do not break the system.** An analyst can freely drag shapes, change colors, edit text in PowerPoint, and the next AI-assisted edit session picks up from the actual current state.

4. **A new client engagement can be onboarded** by adding a `BRAND{}` entry and a project-level `CLAUDE.md` — no code changes to the core system.

5. **Knowledge accumulates.** Every new OOXML workaround, layout pattern, or chart type solved during any engagement is captured in `pptx_utils.py` and available to all future engagements.

6. **Data flows from Synapse to slides** without manual Excel wrangling. The analyst specifies survey IDs, and the system fetches, transforms, and visualizes the data.

---

## 11. Glossary

| Term | Definition |
|---|---|
| **OOXML** | Office Open XML (ECMA-376) — the XML-based file format inside `.pptx`, `.docx`, `.xlsx` files |
| **EMU** | English Metric Unit — the internal measurement unit in OOXML. 1 inch = 914,400 EMU |
| **COM** | Component Object Model — Windows system for inter-process communication. PowerPoint exposes its object model via COM |
| **win32com** | Python library (`pywin32` package) that provides a bridge from Python to Windows COM |
| **python-pptx** | Python library for creating and modifying `.pptx` files by manipulating the XML directly |
| **lxml** | Python library for XML/HTML parsing and manipulation. Used to fill gaps where `python-pptx` has no API |
| **Reconcile** | The process of reading live shape state from PowerPoint and updating the registry to match |
| **Registry** | `slide_registry.json` — maps shape IDs to semantic labels and caches current property values |
| **pkl** | Python pickle file — serialized data snapshot used as the intermediate format between data source and slide generation |
| **Synapse** | ZoomRx internal system whose APIs provide survey cross-tab data |
| **PET** | Pharma Executive Tracker — one category of ZoomRx consulting deliverable (but this system serves all categories) |

---

## Appendix A: OOXML Properties Requiring Raw lxml

These are formatting properties that `python-pptx` does not expose via its API. They are pre-solved in `pptx_utils.py` and should never be written inline in generation scripts.

| Property | OOXML Element | `pptx_utils` Function |
|---|---|---|
| Bar gap width | `c:gapWidth` | `set_plot_area_gap()` |
| Bar overlap | `c:overlap` | `set_overlap()` |
| Axis orientation (invert) | `c:scaling > c:orientation val="maxMin"` | `invert_cat_axis()` |
| Axis tick label visibility | `c:tickLblPos val="none"` | `hide_cat_labels()` |
| Data label position | `c:dLblPos val="outEnd"` | `set_datalabel_pos_outside_end()` |
| Series border removal | `a:ln > a:noFill` | `set_series_no_border()` |
| Axis number format | `c:numFmt formatCode="0"` | `set_val_axis_number_format()` |

This table will grow as new formatting needs are discovered. Each addition follows the same pattern: discover the XML, write a named function, document it here.

---

## Appendix B: Existing Ecosystem (as of March 2026)

Before building `pptx_utils.py`, the developer should evaluate what already exists. This research was conducted to pressure-test whether the utility library needs to be built from scratch.

**Core finding: Nobody has published a maintained library wrapping python-pptx's lxml gaps.** Everyone doing serious chart formatting rolls their own helpers.

### Libraries to evaluate before building

| Library | What it does | Relevance | Action |
|---|---|---|---|
| **python-pptx-ng** ([PyPI](https://pypi.org/project/python-pptx-ng/)) | Fork of python-pptx adding gradient fills, enhanced data labels, group shapes, shadow formatting, OLE embedding | May eliminate some lxml workarounds entirely | **Evaluate in Phase 1.** If it covers bar gap width, axis orientation, or label positioning natively, use it instead of writing lxml helpers for those. |
| **pptxlib** ([PyPI](https://pypi.org/project/pptxlib/)) | win32com wrapper for PowerPoint (beta, Dec 2025, 0 stars) | Could provide higher-level win32com API for the editing phase | **Evaluate in Phase 0.** If its API is stable enough, it could reduce the amount of raw win32com code. But given 0 stars and beta status, do not depend on it. |
| **python-pptx-fix** ([PyPI](https://pypi.org/project/python-pptx-fix/)) | Patched fork fixing specific python-pptx bugs (date axis, Unicode, bubble chart labels) | Fixes known bugs without changing API | **Use if** you hit any of the specific bugs it patches. Drop-in replacement for python-pptx. |

### Libraries that exist but are NOT relevant

| Library | Why not |
|---|---|
| **presenton** (4.2k stars) | Full-stack AI presentation generator. Different problem — generates entire decks, doesn't solve low-level chart formatting. |
| **slide-deck-ai** (315 stars) | Streamlit app for LLM-generated slides. Content generation, not formatting control. |
| **python_pptx_interface** (44 stars) | Higher-level slide composition. Stale since 2020. Doesn't address chart formatting gaps. |
| **pptgen** | YAML-configured template engine. Stale since 2017. |
| **Aspose.Slides** (commercial, from $979) | Full PowerPoint API without Office installed. Overkill — we have Office installed and need COM integration. |

### What this means for the PRD

The `pptx_utils.py` approach is validated as filling a genuine gap. But scope it tightly:
- Start with only the ~10 lxml helpers proven in the Feb 2026 POC
- Let Claude handle the long tail of uncommon formatting via inline lxml + skill instructions
- Audit after Phase 4 to decide whether to expand the library or lean more on Claude's native capabilities
- Check python-pptx-ng first — it may already cover some of the gaps natively

---

## Appendix C: Origin Context

This PRD is derived from a detailed architecture exploration conversation between Sriram (CEO) and Claude in February 2026. That conversation:

1. Started from a working Python script that generated a single J&J Rybrevant slide with native PPT charts
2. Explored the `python-pptx` + `lxml` boundary in depth
3. Evaluated 5 architecture options for the editing interface (outside PPT, Office JS add-in, hybrid, web app, COM automation)
4. Landed on win32com + Claude Code as the V1 architecture
5. Worked through resilience concerns: manual edits, copy-paste, duplicate names, registry staleness
6. Established the company skills vs. project artifacts separation

The full conversation transcript is preserved in `[3] SlideGen — Architecture Exploration (Sriram × Claude, Feb 2026).md` for reference.
