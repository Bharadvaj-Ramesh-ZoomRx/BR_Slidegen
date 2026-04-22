# SlideGen: From Vinoth's Experiment to the Scaled-Up Vision

---

## PART 1 — What Is Vinoth's Experiment? What Did He Actually Build?

### Origin Story

SlideGen started as a practical problem Vinoth — a ZoomRx consulting analyst — was facing every single day: manually building PowerPoint slide decks for pharma clients. Every engagement meant the same repetitive work: export survey data from Synapse, wrangle Excel, draw charts, apply brand colors, align shapes, update for the next quarter wave. None of that institutional effort accumulated anywhere — each analyst started from scratch.

The experiment started from one working Python script that generated a **single J&J Rybrevant slide** with native PowerPoint charts. It then evolved — over roughly 12 hours of iteration with Claude — into a full **YAML-driven pipeline** that could create an entire consulting deck from survey data.

The specific brand Vinoth was working with: **Rybrevant (RYB) + Lazcluze vs. Tagrisso (TAG)** — a J&J PET (Pharma Effectiveness Tracking) study covering message recall, rep effectiveness, prescription intent, call-to-action metrics, and high-impact interactions.

---

### What Vinoth Built — The Full System

Let's go through every component term by term.

---

#### 1. The YAML-Driven Pipeline Concept

The **core insight** of the experiment: instead of writing a new Python script for every slide deck (like the old hardcoded `generate_asks.py` in the archive), define what a deck looks like in a **config.yaml file**, and let the code read that config and build the deck.

This means:
- A consultant writes (or Claude generates) a YAML file that says: "Slide 3 is a `single_bar_with_delta` chart, use question code `Q2_01Z`, show prior wave Q3 vs current wave Q4, sort descending."
- The pipeline reads that YAML and produces the `.pptx` — no code changes needed.
- A new project (different brand, different client) = a new `config.yaml`, not new code.

This is what "YAML-driven" means. YAML (Yet Another Markup Language) is a human-readable config format — like a dictionary — that maps parameters to values. Think of it as a recipe that tells the kitchen (the Python pipeline) exactly what to cook.

---

#### 2. The 6-Stage Create Pipeline (The Intelligence Layer)

The most important piece Vinoth built wasn't just the rendering — it was the **analytical intelligence pipeline** that decides *what to put on the slides*. It has 8 stages when fully counted:

```
Stage 0    → Index the Excel file into a structured JSON
Stage 0.5a → Build Market Context (competitive/clinical landscape)
Stage 0.5b → Extract Prior Wave Context (what last quarter said)
Stage 0.5c → Build Survey Context (parse the survey questionnaire)
Stage 1    → Build Project Context (study design, call notes, KBQs)
Stage 2    → Generate Hypothesis Bank (testable predictions by KBQ)
Stage 3    → Validate Data + Write Narrative Threads (story arcs, headlines, ES, recs)
Stage 4    → Build Slide Plan (which chart for which hypothesis)
Stage 5    → Generate config.yaml (map slide plan to YAML extractions)
Stage 6    → Run generate_deck() → output deck.pptx
```

Let's explain each of these in plain language:

**Stage 0 — index_excel()**
Takes the raw Excel file that Synapse exports (`source_data.xlsx`) and converts it into a structured JSON file called `source_data.json`. This JSON has two key sections:
- `_sheets` — every row in the Excel, row by row, with code, description, and values
- `_codes` — per question-code metadata: how many sub-rows it has, whether values are decimals (0–1) or whole percentages (0–100), sample values

Why does this matter? Because the Excel file has no consistent structure — columns shift between clients and waves. The JSON indexer reads the header rows, figures out which column is Q3 and which is Q4, and stores that so future stages never have to look at the raw Excel again. After Stage 0, the JSON is the data source for everything downstream.

**Stage 0.5a — Market Context (`/market-context`)**
Claude uses its knowledge of the therapy area (e.g., non-small cell lung cancer) and the competitive landscape (Tagrisso's dominance, Rybrevant's positioning, clinical trial results) to write a `market_context.md` file. This is NOT wave-specific — it's shared across all waves of the same project. It runs a 3-round adversarial fact-check on itself before saving.

**Stage 0.5b — Prior Wave Context (`/prior-wave-context`)**
Scans the input folder for prior quarter's deck (the PPTX) or a `prior_wave_es.md` summary. Extracts what the prior report said: which metrics went up, which went down, what recommendations were made, what questions were left unanswered. Saves as `prior_wave_context.md`. This feeds directly into Stage 2 — so the hypothesis bank knows what to TEST this wave (not just generate fresh hypotheses from scratch).

**Stage 0.5c — Survey Context (`/survey-context`)**
If there's a survey draft file (Word/PDF), it parses the full questionnaire: every question code (`Q2_01Z`, `Q1_87Z`, etc.), response scales, message lists, module structure. Saves as `survey_context.md`. This ensures hypothesis generation and slide planning reference the right question codes — not guesses.

**Stage 1 — Build Project Context (`/build-project-context`)**
Reads everything else: client call notes (`.docx`), the study design document (`.odt`), the standing Key Business Questions file (`KBQs.md`). Synthesizes a `project_context.md` covering: who the client is, what the study measures, what questions the client is most focused on, any special instructions from the call notes.

**KBQs (Key Business Questions)** — these are the specific business questions the client needs answered. For a PET study, examples might be: "Is Rybrevant's message recall improving vs. Tagrisso?", "Are high-impact interactions driving prescription intent?". These are hand-written by the research team — they can't be auto-generated.

After Stage 1, there is a **single validation gate**: the user reviews all 4 context files before any hypothesis generation begins. This is by design — if the context is wrong, everything downstream is wrong.

**Stage 2 — Generate Hypothesis Bank (`/hypotheses`)**
This is where Claude becomes an analytical research partner. Reading the 5 context files (market, project, prior wave, survey, KBQs), it generates a structured bank of testable hypotheses organized by KBQ domain. Example:
- *"H1 [Prior Wave Validation]: Rybrevant's T2B message recall for efficacy improved vs. prior wave (Q3: 34% → hypothesis: Q4 ≥37%), consistent with rep guidance shift noted in call notes."*

The bank includes 4 hypothesis types:
- **New** — fresh analytical question this wave
- **Prior Wave Validation** — testing whether last wave's finding holds
- **Action Item** — flags something that needs client attention
- **Methodology Artifact** — acknowledges a data quirk or sample size issue

**Stage 3 — Validate Data + Narrative Threads (`/sfea-insight-writer`)**

This stage runs in 2 phases:

*Phase 0 (automatic):* For every hypothesis, Claude looks up the actual survey data in `source_data.json` and extracts: prior value, current value, delta (change in percentage points). Saves as `validated_analysis.md`. This catches hypotheses that assumed the wrong direction ("we expected improvement but actually declined").

*Phase 1 (with user gate):* Claude synthesizes **story arcs** — 3 to 5 cross-domain narrative threads that connect findings across different KBQ domains into a coherent story. It uses arc patterns:
- **CONVERGENCE** — multiple metrics all pointing the same direction
- **TENSION** — metrics pulling in opposite directions, creating strategic tension
- **DIVERGENCE** — two brands or segments diverging
- **CLOSURE** — a prior wave recommendation was followed and it worked

Then it writes: arc-informed slide headlines, an arc-organized executive summary, and arc-driven recommendations. Everything goes into `narrative_threads.md`. This is the "story" of the deck before a single slide is planned.

**Stage 4 — Build Slide Plan (`/slide-plan`)**
Takes the narrative threads and clusters hypotheses into a specific slide plan. Each slide gets:
- A slide number
- A headline (copied verbatim from narrative_threads.md — not written fresh)
- A chart type (e.g., `clustered_compare`, `abacus`, `single_bar_with_delta`)
- The question codes that drive it
- Which story arc it belongs to and its "role" in that arc

Sequencing is arc-informed: ACT NOW arcs (urgent findings) come first, then MONITOR arcs, then CELEBRATE arcs. Always starts with Cover → Executive Summary → Recommendations → data slides.

**Stage 5 — Generate config.yaml (internal)**
Maps the slide plan to actual YAML configuration. For each slide, it looks up every question code in the `_codes` index to confirm it exists, auto-selects the extraction method, determines whether values are decimals or whole numbers, and writes the YAML ask entry. This is Claude doing data engineering — connecting the analysis plan to the actual data.

**Stage 6 — generate_deck()**
The Python pipeline reads the config.yaml, loads data (from JSON cache or by extracting the Excel), calls the appropriate renderer for each slide type, and saves the final `deck.pptx`.

---

#### 3. The Data Architecture — Four Tracks

Vinoth built a sophisticated data loading system with four independent tracks for getting data into the pipeline:

**Track A — JSON-First (fastest)**
For `synapse_report` extractions: calls Synapse's `/reports/generate` API directly and gets JSON back. Bypasses Excel entirely. Only works when `SYNAPSE_API_KEY` environment variable is set.

**Track B — Excel (default, most robust)**
Downloads the survey cross-tab data as an Excel file via Synapse's banner plan API. This is the standard export — the same Excel a human analyst would get. The pipeline then extracts from it using pandas. If no API key is set, the analyst can just drop the Excel file in the input folder and it works exactly the same.

**Track C — Raw API (legacy)**
For respondent-level analysis — not aggregated averages, but individual survey responses. Downloads the raw Excel with every respondent's answers, auto-discovers Virtual Questions (VQs — derived metrics that combine multiple raw questions), applies segment cuts (e.g., "only Academic HCPs"), and aggregates locally.

**Track D — Raw-data-first (preferred new approach)**
Uses Rajesh's `synapse-cli` tool to download all responses + VQs + segments in 1–2 API calls as NDJSON (newline-delimited JSON — a streaming format). Caches as a columnar JSON structure. Faster and more efficient than Track C.

**Three data tiers based on what's being extracted:**
- **Tier 1** — Aggregated stats from the cross-tab Excel (the main percentages you see on charts)
- **Tier 2** — Respondent-level local analysis (when you need top-2-box or recall % computed from raw responses)
- **Tier 3/4** — Respondent-level API-based (same as Tier 2 but pulling directly from Synapse instead of a local Excel)

---

#### 4. The 22 Slide Type Renderers

Each of the 22 slide types is a Python function that takes a data dictionary and configuration parameters and produces one complete, client-delivery-quality slide. Let's explain the key ones:

**`single_bar_with_delta`** — A horizontal bar chart showing one metric (e.g., Rybrevant message recall by message item), with a separate column showing the quarter-over-quarter change in percentage points (the "delta"). The bars show current values; the delta column shows the direction of change (green for positive, red for negative).

**`clustered_compare`** — Side-by-side bars comparing two groups (e.g., Academic vs. Community HCPs, or High Impact vs. Others) on the same metric. Includes a "gap" column showing the difference between groups.

**`abacus`** — A scatter plot where each HCP rep or segment is plotted as a dot on a horizontal axis. Used for rep performance data — you can see which reps are above/below average, with dots for current and prior period. Accompanied by value tables.

**`dual_bar_with_delta`** — Two separate bar charts side by side for two different brands (Rybrevant and Tagrisso), with delta columns for both. Useful for head-to-head metric comparisons.

**`hii_scorecard`** — "HII" stands for High Impact Interaction. A multi-section clustered column chart grouping multiple metrics into sections (e.g., grouped by type of rep activity), with section headers and callout boxes.

**`trended_scorecard`** — A grid of small line charts showing QoQ trends across multiple metrics at once — like a dashboard view of trends.

**`quadrant_scatter`** — A 2×2 matrix scatter plot. Used for importance mapping: X-axis is stated importance (what HCPs say matters), Y-axis is derived importance (what actually correlates with prescribing behavior). Quadrants: high-high (critical), high-low (overpromised), low-high (hidden gems), low-low (deprioritize).

**`dual_doughnut`** — Two sets of side-by-side doughnut charts comparing patient segments (e.g., new vs. established patients) by brand. Good for patient mix analysis.

The technical challenge of all these renderers: Python's `python-pptx` library doesn't expose all PowerPoint XML properties through its API. Things like bar gap width, axis inversion, data label position — these require direct XML manipulation via `lxml`. Vinoth (and Claude) spent significant effort reverse-engineering the OOXML (the XML format inside `.pptx` files) to build these helpers.

---

#### 5. pptx_utils — The Hard-Won Knowledge Package

This is arguably the most valuable thing that came out of the experiment: a Python package (`pptx_utils/`) that encodes all the non-obvious PowerPoint XML knowledge so it never has to be rediscovered.

**`brand.py`** — The `BRAND{}` dictionary. Stores per-client color and font definitions. J&J's colors: Rybrevant orange `#F75824` (current period), `#FFC199` (prior period), Tagrisso violet `#7030A0`, `#AD88C8`. Fonts: Johnson Display (headers), Johnson Text (body). Why does this matter? Because hardcoding colors in each generation script means if J&J updates their brand guide, you have to find and change 50 places. Here, you change one dict entry.

**`layout.py`** — The `LAYOUTS{}` dictionary. Pre-computed coordinates (in inches) for common slide compositions. Example: "two chart with delta" layout — left chart starts at x=0.20", top=1.47", width=7.30", with a 0.62" delta column. These coordinates were reverse-engineered from real client decks to match exactly what was being delivered. This is critical: PowerPoint positions shapes in EMUs (English Metric Units) — 1 inch = 914,400 EMUs. Getting coordinates wrong by even a fraction of an inch makes slides look amateurish.

**`charts.py`** — The `CHART_PATTERNS{}` dictionary and chart builders. Each pattern encodes: chart type, axis configuration (is the category axis inverted?), gap width between bars, data label position, series formatting. This means Claude doesn't have to figure out "how do I make a horizontal clustered bar chart?" every session — it calls `make_clustered_bar()` with the right pattern.

**`lxml_helpers.py`** — The most critical file. These are the 20+ functions that manipulate raw OOXML XML. Examples:
- `set_plot_area_gap()` — sets bar gap width (controls how fat/thin bars appear). The python-pptx API doesn't expose this, so it reaches directly into the XML `<c:gapWidth>` element.
- `invert_cat_axis()` — flips the category axis so bars read top-to-bottom instead of bottom-to-top. In OOXML: `<c:scaling><c:orientation val="maxMin"/></c:scaling>`
- `hide_cat_labels()` — hides axis tick labels (because the label table next to the chart replaces them visually). In OOXML: `<c:tickLblPos val="none"/>`
- `set_datalabel_pos_outside_end()` — places data labels at the outside end of bars. Critical for readability.

**`registry.py`** — The shape registry system. Every shape on every generated slide gets a unique ID (`zrx_001_001` = slide 1, shape 1). The registry saves a JSON file mapping each shape ID to: what it is (chart, textbox, table), where it is (coordinates), what data drives it (which question code, which wave, which Synapse query). This is the data lineage record.

**`com.py`** — Helpers for `win32com`, which is Microsoft's COM automation interface. This allows Python to connect to a *running PowerPoint process* and make changes that appear live on screen without saving and reopening the file. Critical for the editing workflow.

---

#### 6. Wave Versioning

The concept of "wave" in pharma consulting: surveys are run every quarter (Q3 2025, Q4 2025, Q1 2026, etc.). Each quarter is a "wave." The pipeline handles wave versioning by storing everything in wave-named subfolders:

```
input/wave/Q1 2026/source_data.xlsx    ← new wave data
context/Q1 2026/source_data.json       ← extracted from that Excel
output/Q1 2026/deck.pptx               ← the generated deck for that wave
```

The `config.yaml` uses `{{wave}}` as a template variable — so changing `project.wave: "Q4 2025"` to `project.wave: "Q1 2026"` automatically redirects all paths to the new wave's folders. Old wave outputs are preserved.

---

#### 7. Natural Language Interface

The entire workflow is triggered through **natural language in the Claude Code terminal** — not through code changes or CLI flags. Examples:
- `"Create slides for projects/jnj_rybrevant"` → runs the full 6-stage pipeline
- `"Edit Slide 5 — change headline to 'Updated Message Recall'"` → runs `regenerate_slide()` for slide 5 only
- `"Edit slides with new wave data — Q1 2026"` → updates the config wave and runs `generate_deck()`
- `"Refresh this deck"` → re-extracts all data and rebuilds all data-driven slides

Claude Code translates these utterances into the right Python function calls, with the right parameters, against the right config file.

---

#### 8. Distribution Model

Git + OneDrive split:
- The **git repo** (`galen-consulting/`) contains the code (`slidegen/` package, `.claude/skills/`, docs). Analysts pull updates via `git pull`.
- The **projects folder** (`projects/`) lives on a shared OneDrive folder — gitignored. It has all the client-specific data, configs, templates, and outputs. Analysts symlink their local clone's `projects/` directory to the OneDrive folder.

Why this split? The code is universal; the data is client-confidential. Code should be version-controlled; data should be on a shared file system accessible to the whole team.

---

### What Vinoth's Experiment Was NOT

According to the PRD v1.1, here's what the experiment has as its limitations:

1. **Slide rendering was reverse-engineered from ONE J&J deck** — so it has hardcoded assumptions about J&J's specific data structure. Wrong colors appear for other clients. Zero-data messages get included when they shouldn't. Table data ends up misplaced in edge cases.

2. **The flow is prescriptive and tightly coupled** — if you want to refresh a single slide with new data, you have to pretend to run the full hypothesis pipeline. The skills assume the entire upstream chain (hypothesis bank → narrative threads → slide plan) has already executed. For a wave refresh or a single-slide client follow-up, none of that context may be needed.

3. **Only one project type** — PET (Pharma Effectiveness Tracking) for J&J. ATU (Awareness/Trial/Usage), HCP-Patient research, Digital Tracker, PCA (Patient/Caregiver research) are all unsupported.

4. **Only one workflow** — creating a new deck from scratch. There's no structured system for: refreshing an existing deck with new wave data while preserving the narrative, adding a single slide in response to a client question, auditing a deck before delivery, generating an executive summary from an existing deck.

5. **No way to read existing decks** — if you have a deck that was built manually or in a prior system, you can't feed it into SlideGen to start using the pipeline from there.

---

## PART 2 — What Is the PRD v1.1? How Does It Scale Vinoth's Experiment?

The PRD v1.1 was written by **Vijay Ganesan on April 16, 2026**, after a leadership review by Sriram (CEO) and Siva (VP Engineering). It validated the architecture direction and gave the green light to start moving. It's explicitly based on: the Q2 QBR, conversations with Sriram/Siva/Manoj, a SlideGen walkthrough, Siva's March PRD, and — critically — **a reverse-engineering analysis of 905 real client decks across 96 pharma clients**.

This is not a small experiment anymore. This is a product.

---

### Scaling Area 1: From One Workflow to Eight MECE Workflows

**Vinoth's experiment:** One workflow — create a new deck from scratch using the full hypothesis pipeline.

**PRD v1.1:** Eight workflows covering the **entire deck lifecycle**, designed to be MECE (Mutually Exclusive, Collectively Exhaustive — no overlap, no gaps):

1. **`create-deck-workflow`** — Build a new deck from a brief, raw data, and KBQs. Has two modes: `briefing` (Claude generates hypotheses from scratch) and `hypothesis` (pre-built hypothesis bank from Vinoth's flow). The `hypothesis` mode is exactly Vinoth's existing experiment — now one option among many.

2. **`refresh-deck-workflow`** — The most strategically important new workflow. Takes a prior wave deck and new wave data. Produces the next wave deck with the SAME narrative structure but updated numbers, fresh headlines, and delta callouts. This is what happens every quarter when new data comes in — not creating from scratch, but evolving a living document.

3. **`edit-slide-workflow`** — Three sub-modes: `rebuild` (re-render a slide from its current spec), `data_refresh` (re-pull data from Synapse and re-render), `edit` (change specific elements like headline text, colors, sort order).

4. **`add-slide-workflow`** — Client asks a follow-up question mid-engagement. Analyst says: "Create a slide comparing Academic vs. Community HCPs on rep frequency." This workflow handles that — pulls the relevant data, picks the right chart type, inserts the slide in the right position.

5. **`annotate-slide-workflow`** — Add a callout to an existing slide. A quote from a qualitative respondent. A significance marker. A data annotation. Doesn't create a new slide — modifies an existing one.

6. **`structural-edit-workflow`** — Delete a slide, reorder slides, split one slide into two, merge two adjacent slides into one. No changes to individual slide content — pure deck structure.

7. **`deck-audit-workflow`** — Before sending a deck to a client, run an automated pre-delivery QA. Checks: do the data values match what the source says? Do ES bullets have citations back to specific slides? Are any headlines stale (e.g., says "up 3pp" when data actually shows "down 2pp")? Are there unsupported claims? Returns a report — read-only, never modifies the deck.

8. **`executive-summary-workflow`** — Takes a full deck and KBQs, generates 1–3 ES slides with findings that cite back to specific slide IDs. Every bullet must have a citation or it's rejected.

Each workflow maps to a **verb the user actually says**: create, refresh, edit, add, annotate, restructure, audit, summarize.

---

### Scaling Area 2: From Tightly Coupled Skills to a Composable Skill Library (52 Skills)

**Vinoth's experiment:** ~10 skills, all designed to work in sequence for the PET hypothesis flow. Skills assume the whole upstream chain has run.

**PRD v1.1:** 52 skills organized into 6 categories, each independently callable, composable in any order:

**Context + Data Skills (9)** — skills that produce inputs:
- `project-context-builder`, `market-context-builder`, `survey-context-builder`, `prior-wave-context-builder` — these exist in Vinoth's work, essentially unchanged
- `synapse-read` — wraps Rajesh's CLI as a proper callable tool (currently it's standalone, not integrated)
- `deck-reader` — **NEW, critical** — reads an existing PPTX and extracts what's on every slide
- `hashtag-benchmarks` — **NEW** — pulls industry benchmarks from Hashtag (a ZoomRx data platform)
- `raw-data-aggregator` — exists, unchanged
- `excel-indexer` — exists, unchanged

**Planning Skills (7)** — skills that decide what to put on slides:
- `viz-selector` — **NEW, critical** — picks the right chart type deterministically (see Scaling Area 5)
- `slide-plan-generator-hypothesis` — exists (Vinoth's current slide plan skill)
- `slide-plan-generator-refresh` — **NEW** — diffs a prior wave deck vs. new data to create an edit plan
- `slide-plan-generator-single` — **NEW** — one client question → one slide spec
- `slide-plan-generator-exec-summary` — **NEW** — full deck → ES slide specs with citations
- `spec-validator` — **NEW** — validates a slide spec has everything the renderer needs before rendering
- `multi-element-composer` — **NEW** — combines table + chart + benchmark + callout into a single slide

**Creation Skills (7)** — skills that produce slide output:
- `slide-creator` — exists but needs major refactor to be general (not J&J-specific)
- `slide-updater` — **NEW** — update an existing slide with fresh data, preserve structure
- `slide-editor` — **NEW** — edit specific elements via win32com live editing
- `callout-adder` — **NEW** — add an annotation callout to an existing slide
- `deck-assembler` — needs refactor — assembles slides into final PPTX with correct ordering
- `executive-summary-writer` — **NEW** — polished ES content with citations

**Analysis Skills (6)**:
- `hypothesis-generator` — exists (Stage 2)
- `sfea-insight-writer` — exists (Stage 3, PET-specific)
- `atu-insight-writer` — **NEW** — parallel to sfea-insight-writer but for ATU studies, with different arc patterns (FUNNEL_LEAKAGE, SHARE_MOMENTUM, COMPETITIVE_CONVERGENCE, BARRIER_CLUSTER, etc.)
- `segment-comparator` — **NEW** — statistical comparison between segments with T-test and chi-square significance testing
- `trend-analyzer` — **NEW** — cross-wave trend analysis
- `stat-sig-annotator` — **NEW** — annotate slides with statistical significance markers and low sample size footnotes

**Project-Type Skills (5)** — skills encoding the methodology for each project type:
- `pet-deck` — refactor from Vinoth's implementation
- `atu-deck` — **NEW** — ATU funnel structure
- `pca-deck` — **NEW** — Patient/Caregiver analysis
- `hcp-pt-deck` — **NEW** — HCP-Patient research
- `digital-tracker-deck` — **NEW** — Digital engagement

**Workflow Skills (8)** — the top-level orchestrators:
These are the 8 workflows described above, each implemented as a SKILL.md that composes the lower-level skills.

**The key design principle Sriram articulated:**
> "There are a bunch of building blocks and any specific output is a stringing together of those building blocks. In what order to string the building blocks we hand over that decision to Claude Code."

---

### Scaling Area 3: The Slide Spec as a Contract — `slide_spec/schema.py`

**Vinoth's experiment:** The slide plan (`slide_plan.md`) was a markdown document that Claude read and then used to generate a config.yaml. There was no formal contract between the planning stage and the rendering stage. Claude had to interpret the plan correctly, and if it missed something, the render would fail silently.

**PRD v1.1:** Introduces the `SlideSpec` — a formal Python dataclass that is the **contract between intelligent planning skills and deterministic renderers**.

Every slide in the system is represented by a `SlideSpec` object with these components:
- `slide_id` — unique identifier
- `slide_type` — one of the N chart patterns (e.g., `bar_clustered_horizontal`)
- `headline` — the talking header text
- `data` — full values, labels, sort order, question codes
- `formatting` — colors (as tokens, not hardcoded), fonts, layout positions, period labels
- `extras` — chart-type-specific parameters
- `data_source` — provenance: which question code, which segment, which wave, which Synapse analysis
- `spec_completeness` — one of: `"complete"`, `"layout_complete_data_missing"`, `"partial"`

The `spec-validator` skill checks a spec for completeness before any rendering begins. If the spec is incomplete, it fails loudly and forces the upstream planning skill to fix it — rather than producing a broken slide silently.

**Why this is a fundamental architectural upgrade:**
Multiple different skills can now produce a valid spec from completely different starting points:
- `slide-plan-generator-hypothesis` produces a spec from a hypothesis bank
- `slide-plan-generator-refresh` produces a spec by diffing a prior deck against new data
- `slide-plan-generator-single` produces a spec from a single client question
- `deck-reader` produces a spec by reading an existing PPTX

And regardless of how the spec was produced, `slide-creator` renders it identically. No bespoke glue. No workflow-specific rendering branches.

**Color tokens (not hardcoded hex)**:
Brand colors in the spec are stored as tokens:
- `"#F75824"` — explicit hex (used when you know exactly what you want)
- `"{brand.primary_current}"` — resolved from `BRAND{}` at render time
- `"{context.brand_palette.primary}"` — resolved from the project's context files
- `"{deck.slide_4.series_0.color}"` — pulled from a prior wave PPTX by `deck-reader`

This means the same spec is portable across brands and wave refreshes without rewriting. Swap the brand token resolution and the whole deck re-colors.

---

### Scaling Area 4: The Deck Reader — Reading Existing Decks

**Vinoth's experiment:** No way to read an existing PPTX. If you had a deck that was built manually or in a prior quarter, you had to start from scratch with SlideGen.

**PRD v1.1:** `deck-reader` — a new skill that parses any existing PPTX and produces a `SlideSpec` for every slide. This is what makes every other workflow (refresh, edit, add, annotate, audit, restructure) possible when starting from an existing deck.

`deck-reader` works in two tiers:

**Tier 1 — Galen-PowerPoint Connector Tags (preferred)**
ZoomRx has a PowerPoint add-in called the Galen-PowerPoint Synapse Connector. When analysts use this add-in to insert charts, it stamps a tag on every shape containing the full Synapse data lineage: `ProjectId`, `ReportingPlanId`, `AnalysisIds[]`, `SurveyId`, `SegmentIds[]`, time periods, pivot configuration, mapping configuration.

`deck-reader` reads these tags (via `ReportConfigHash` — a SHA256 key that points to a Custom XML Part containing the full JSON config). If the tag is healthy (5 health checks: parseable, IDs still exist in Synapse, schema compatible, data shape consistent, mapping plausible), it populates the spec's `DataLineage` directly — no inference needed.

**Tier 2 — Structural Inference (fallback)**
For shapes WITHOUT Connector tags, or where the tag fails a health check:
- Parse the headline text, chart type, category labels, table content
- Cross-reference against `source_data.json` to propose candidate question codes
- Assign a confidence score to each candidate
- Mark spec completeness as `"layout_complete_data_missing"` — renderable but not automatically refreshable
- Low-confidence inferences surface for user confirmation; high-confidence ones proceed silently

**The two-pass model:**
Pass 1 (automatic): Extracts everything the PPTX itself tells us — chart data, positions, colors, layout. This always completes. The spec is renderable after Pass 1.
Pass 2 (data lineage): Resolves WHERE the numbers came from. Connector tags = fast and complete. No tags = structural inference with candidate proposals.

**Why this is transformative:** It gives SlideGen **backward compatibility with hundreds of existing client decks**. You don't need to have built the deck with SlideGen to start using the refresh and edit workflows on it. One-time bootstrapping → the deck is "SlideGen-legible" forever after.

---

### Scaling Area 5: Viz Selector — Deterministic Chart Type Selection

**Vinoth's experiment:** The slide plan skill decided chart types through Claude's reasoning — which means different runs could produce different chart choices, and edge cases (metrics Claude hadn't seen before) could produce inconsistent output.

**PRD v1.1:** `viz-selector` — a deterministic hierarchy for picking chart types:

```
Level 1: Metric Tag → Chart Pattern mapping (override layer)
  Message Recall           → Table + Cluster Bar Chart
  Topic Recall             → Table + Cluster Bar + Benchmark Table
  Awareness Trend          → Line Chart
  Likelihood to Prescribe  → 100% Stacked Column
  (36 pre-mapped metric tags from real PET chart analysis)

Level 2: Question Type Default (if no metric tag)
  Likert scale             → bar_clustered_horizontal
  Multi-select             → stacked_bar
  Numeric distribution     → histogram / box plot
  Ranking                  → ranked bar

Level 3: Human-in-the-loop (if neither applies)
  → Prompt user: "Which chart type for X?"
```

The 36 metric-to-chart mappings were derived from **analyzing 4,354 real PET charts across 17 clients**. This is data-driven, not opinion-driven. If real analysts consistently use clustered bars for message recall, that's what the selector picks.

**The key insight from Sriram (Apr 14):** "Use all of the intelligence there itself and say these are all the parameters that you need to specify in this format so that that deterministic program can use all of those parameters to create the perfect slide." Intelligence flows INTO the spec; deterministic code reads the spec.

---

### Scaling Area 6: 905-Deck Corpus Analysis → pptx_utils Grounded in Reality

**Vinoth's experiment:** `pptx_utils` was built by reverse-engineering one J&J deck. Brand colors, layout coordinates, chart patterns — all derived from a single data point.

**PRD v1.1:** Before writing a line of scaled code, the team ran a comprehensive reverse-engineering analysis:

- **32 PET decks across 17 clients** (April 15) — initial analysis
- **8 ATU decks** (April 16) — ATU vs. PET comparison
- **905 full-corpus scan** (April 17) — via `mass_deck_scanner.py` across 276 PET, 143 ATU, 83 HCP-Pt, 50 Qualitative, 27 PCA, 26 Digital Tracker decks

Key findings from 50,072 charts:

**Top 9 chart types cover 99% of all charts:**

| Chart Type | Occurrences | % |
|---|---|---|
| `bar_clustered` (horizontal) | 13,994 | 28% |
| `column_stacked_100` (vertical) | 8,227 | 16% |
| `line_markers` (trended) | 6,857 | 14% |
| `bar_stacked_100` (horizontal) | 5,528 | 11% |
| `xy_scatter` (abacus) | 4,260 | 9% |
| `bar_stacked` (horizontal) | 3,637 | 7% |
| `column_clustered` (vertical) | 2,869 | 6% |
| `column_stacked` (vertical) | 1,801 | 4% |
| `doughnut` | 1,752 | 4% |

**OOXML defaults confirmed at scale:**
- Chart title: present in only **0.6%** of charts — almost never used
- Chart legend: present in only **1.2%** — almost never used (labels are next to bars or in separate tables)
- Data label number format: `"0%"` on **74%** of labels
- Category axis inverted: **29%** of charts — very common for horizontal bar charts
- Most common bar gap width: 100 (default), 50 (fatter bars), 150 (thinner bars)

**The outcome:** `BRAND{}` populated with **102 client entries** and **33 brand entries**. `LAYOUTS{}` with **14 coordinate presets**. `CHART_PATTERNS{}` with **15 patterns**. All grounded in real delivered decks, not hypothetical examples. This is the foundation that ensures `slide-creator` produces output that is indistinguishable from what a senior analyst would manually produce.

---

### Scaling Area 7: Evals Infrastructure — Measuring Quality at Scale

**Vinoth's experiment:** No formal testing for quality. You'd know something was wrong when a slide looked wrong.

**PRD v1.1:** A systematic evals framework, owned by Bharadvaj.

**Approach:** Use prior wave decks as ground truth.
1. Take a prior-wave PPTX where the "correct" output is known (e.g., JJ RYB Q4 '25 deck)
2. Run `deck-reader` on it to generate specs
3. Modify a known set of data points (e.g., +3pp on a few metrics)
4. Run `refresh-deck-workflow` against the modified data
5. Compare output to expected: correct values, correct headlines, no stale text, correct layout

**Eval dimensions:**
- **Data fidelity** — output values match source data exactly
- **Headline freshness** — no stale headlines from prior wave (a headline saying "up 3pp" when new data shows "down 2pp" is a client-delivery risk)
- **Layout stability** — no unexpected position shifts when data shape is unchanged
- **Lineage completeness** — `shape_registry.json` has `last_data_pull` on all refreshed slides
- **Regression** — slides that should NOT change are unchanged

Lives in `tests/evals/` — each eval is a directory with `input/` (prior wave PPTX + modified Excel), `expected/` (expected output metadata — not full PPTX), and a `run_eval.py` that compares actual vs expected.

**Why this matters for scale:** Once 5+ people are contributing skills, you can't manually verify everything. The eval harness is the safety net that lets the team move fast without breaking client-delivery quality.

---

### Scaling Area 8: The Headline Freshness Principle

**Vinoth's experiment:** Headlines were written during Stage 3 (narrative threads) and baked into the config. On a wave refresh, those headlines could become stale if data changed direction.

**PRD v1.1:** An explicit principle: **when a slide's data changes, its headline must regenerate**. Never left stale.

Implementation:
- `slide-updater` calls `headline-writer` by default after every data update
- `refresh-deck-workflow` regenerates `narrative_threads.md` from new data *before* updating any slides — so fresh headlines inherit from a fresh narrative backbone
- `deck-audit-workflow` flags any headline whose text contradicts the rendered data as BLOCKER severity
- The only exception: caller passes `preserve_headline=True` when they've explicitly verified the text is still accurate (e.g., a wording-only edit where data didn't change)

---

### Scaling Area 9: Multi-Project Type Support — ATU, HCP-Pt, PCA, Digital Tracker

**Vinoth's experiment:** PET only. PET (Pharma Effectiveness Tracking) studies measure rep-level effectiveness — message recall, prescription intent, detail ratings. Everything was built assuming that methodology.

**PRD v1.1:** Extends to all major Galen-Consulting project types:

**ATU (Awareness/Trial/Usage)** — Studies that measure where HCPs are in the brand adoption funnel. New arc patterns specific to ATU methodology:
- `FUNNEL_LEAKAGE` — where in the awareness→trial→usage funnel are HCPs falling off?
- `SHARE_MOMENTUM` — is brand share accelerating or decelerating?
- `COMPETITIVE_CONVERGENCE` — are competitors eating into the brand's differentiation?
- `BARRIER_CLUSTER` — what barriers are clustering together to block adoption?
- `ADOPTION_CURVE` — where is the brand on the classic innovation adoption S-curve?

Chart patterns lean toward `column_stacked_100` (for funnel analysis) and brand share tracking.

**PCA (Patient/Caregiver Analysis)** — Built as an extension of ATU methodology.

**HCP-Patient Research** — Studies that analyze the doctor-patient conversation: treatment journey, shared decision-making, patient barriers. More `doughnut` charts (patient segment breakdowns) and `bar_stacked` (patient experience analysis).

**Digital Tracker** — Digital engagement metrics: HCP website visits, email open rates, channel effectiveness, omnichannel coordination. Deferred to Q4.

The architecture scales cleanly: each project type is a thin `project-skill` SKILL.md that applies methodology-specific defaults on top of the universal building blocks.

---

### Scaling Area 10: Spec-as-Config (SlideSpec v1.2) — No Data in the Spec

**An important evolution that came out of the v1.1 development (April 18 addition):**

Earlier versions of the SlideSpec stored actual data values — the chart categories, series values, table cell text. This made the spec a snapshot in time.

SlideSpec v1.2 changes this: the spec stores **INSTRUCTIONS** for fetching and rendering, not data values. New schema types:
- `DataFilter` — how to filter respondents
- `DataTransform` — how to aggregate raw responses
- `SeriesConfig` — which columns map to which series
- `ChartDataMapping` — per-chart: transform + series config + raw Connector configs (pivot logic)
- `SegmentRule` — structured segment definition

Two refresh paths:
1. **Connected slides (Connector tags):** The spec stores the raw pivot config and mapping config from the Connector. At refresh time, `pivot_records_to_chart_data()` exactly replicates the Connector's pivot, column selection, row filtering, transpose, and sort logic. Output matches perfectly.
2. **Non-connected slides:** `infer_data_transform()` matches series names from OOXML against actual Synapse column values to generate a `DataTransform`. Same refresh pipeline from there.

**Why this matters:** The spec is now permanently refresh-capable. Old specs from prior waves can drive new wave refreshes without modification. The Connector becomes a spec-generation accelerator, not a runtime dependency.

---

## PART 3 — Side-by-Side Comparison: Experiment vs. Scale

| Dimension | Vinoth's Experiment | PRD v1.1 Target |
|---|---|---|
| **Workflows supported** | 1 (create new deck) | 8 MECE workflows (create, refresh, edit, add, annotate, restructure, audit, executive summary) |
| **Project types** | 1 (PET / SFEA) | 5 (PET, ATU, PCA, HCP-Pt, Digital Tracker) |
| **Skills** | ~10, tightly coupled | 52, composable independently |
| **Rendering grounding** | Reverse-engineered from 1 J&J deck | 905 decks, 50,072 charts across 96 pharma clients |
| **Data lineage** | Shape registry (which extraction produced what) | Full SlideSpec as contract + Connector tag integration + Tier 1/2 deck reader |
| **Existing deck support** | None | `deck-reader` with dual-mode Connector + inference |
| **Chart type selection** | Claude reasons about it | Deterministic hierarchy: Metric tag → Q-type default → HITL |
| **Headline management** | Written once in Stage 3, static | Regenerates with data by default; audit catches stale headlines |
| **Spec contract** | Implied in slide_plan.md markdown | Formal `SlideSpec` dataclass + `spec-validator` enforcement |
| **Testing** | Manual visual inspection | Formal evals harness: data fidelity, headline freshness, layout stability, regression |
| **Team model** | Vinoth + Rajesh + Claude | Vijay (editor-in-chief), Pradeep (create-deck-workflow), Bharadvaj (refresh + evals), project-team contributors (per project type) |
| **Data architecture** | 4-track (A/B/C/D) with 3-tier caching | Same 4-track, extended with Connector tag integration |
| **Brand coverage** | 1 client (J&J) | 102 client entries, 33 brand definitions in BRAND{} |

---

## PART 4 — How to Build This End-to-End with Claude Code

This section is written for a consulting analyst who has never coded before. Here is exactly how to use Claude Code to execute any workflow in this system.

### Prerequisites

Install these once:
```bash
pip install pandas openpyxl python-pptx lxml pyyaml requests pywin32
```

Clone the repo and symlink your OneDrive projects folder:
```bash
git clone <repo-url> galen-consulting
cd galen-consulting
# On Windows:
mklink /D projects "C:\Users\YourName\OneDrive\GalenConsulting\projects"
```

### Creating a Deck for a New Project

Open the Claude Code terminal in the `galen-consulting` folder. Then say:

```
Create slides for projects/jnj_rybrevant
```

Claude will:
1. Ask if there's a prior wave PPTX in the input folder (Stage 0.5b)
2. Ask if there's a survey questionnaire draft (Stage 0.5c)
3. Build all context files automatically (Stages 0, 0.5a, 0.5b, 0.5c)
4. Show you all 4 context files and ask you to review them (the one human gate)
5. Generate hypotheses — tell you the count by type, ask you to confirm
6. Run Phase 0 (validate against data) automatically, then show you narrative threads to confirm
7. Build and show the slide plan, ask you to confirm
8. Generate the config.yaml automatically
9. Run the deck generation and report back

If something in the context is wrong, you can say: "Edit the project context — the study covers Q3 vs Q4, not Q1 vs Q2." Claude will edit the file and continue.

### Refreshing a Deck for a New Wave

```
Edit slides with new wave data — Q1 2026
```

Claude will:
1. Back up your current config.yaml
2. Update the wave identifier in the config
3. Check that the new wave's Excel exists in the right folder
4. Run `generate_deck()` — extracts the new data, rebuilds all slides
5. Output goes to a new folder (`output/Q1 2026/deck.pptx`) — old wave preserved

### Editing a Single Slide

```
Edit Slide 7 — sort bars descending by current value, not ascending
```

Claude will:
1. Read the config, find the ask for Slide 7
2. Back up the config
3. Change `sort_by: "current"` and `sort_desc: true` in the YAML
4. Call `regenerate_slide()` for slide 7 only — backs up the PPTX first
5. Tell you to refresh the file in PowerPoint

### Adding a Slide for a Client Follow-Up

```
The client asked about Academic vs Community HCPs on rep frequency. Add a slide after Slide 12.
```

Claude will:
1. Find the right question code for rep frequency in `source_data.json`
2. Pick the right chart type (`clustered_compare` — comparing two segments)
3. Write a new `ask` entry in the config
4. Run `regenerate_slide()` to create that one slide and insert it at position 13
5. Regenerate surrounding slides if needed for slide numbering

### Refreshing All Data (New Data Drop)

```
Refresh this deck
```

Claude will:
1. Delete `source_data.json` to force re-extraction from the updated Excel
2. Run `refresh_deck()` — re-extracts all data, regenerates all data-driven slides
3. Back up the PPTX before starting
4. Report what was refreshed, what was skipped (manually-edited slides), any errors

### Understanding What Claude Will NOT Do Without You

The system has exactly **3 human gates** in the full create workflow:

1. **After Stage 1** — Review all context files. You read 4 markdown files and confirm they're accurate before hypothesis generation starts.
2. **After Stage 2** — Confirm hypothesis count and domain coverage looks right.
3. **After Stage 3 Phase 1** — Review narrative threads, headlines, executive summary, and recommendations. This is the story of the deck — confirm you agree with the analytical direction before slide planning.

After Stage 4 (slide plan), one more gate to confirm the plan.

Everything else runs automatically.

---

## PART 5 — The Philosophy, in Plain English

The deepest thing Vinoth's experiment discovered — and the PRD v1.1 encodes — is a **division of labor between intelligence and determinism**:

- **Intelligence belongs in the spec.** Claude reasons about data, hypotheses, narratives, and visualization choices. This reasoning produces a complete, validated specification: what the slide should say, what chart type it should use, what data should drive it, in what order the bars should appear.

- **Determinism belongs in the render.** Once the spec is correct and complete, a Python function runs and produces the slide. No reasoning. No variance. Same spec → same slide, every time. This is what `pptx_utils` provides — a library of deterministic composition functions that produce client-delivery-quality output.

This separation solves the core problem: a consulting deliverable must be *consistent* (same quality, same standards, every time), *auditable* (you can trace every number back to its source), and *refreshable* (next quarter's deck should take hours, not days). You can't get those properties if the same LLM reasoning that writes the narrative also positions every shape on every slide. You need a stable, code-backed rendering layer that the intelligent layer talks to through a formal contract (the SlideSpec).

The other key philosophy: **build blocks, not pipelines**. Vinoth's experiment was a pipeline — Stage 1 feeds Stage 2 feeds Stage 3 → you had to run the whole thing. The PRD scales this into a skill library — each skill is independently callable. An analyst doing a wave refresh doesn't need to re-run the hypothesis pipeline. An analyst adding a single slide doesn't need the full narrative thread engine. Claude Code composes the right skills in the right order based on what the user actually needs.

That's SlideGen, end to end.
