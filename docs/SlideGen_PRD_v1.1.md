# SlideGen PRD — v1.1

**Status:** Post-leadership review — Apr 16, 2026. Architecture direction validated by Sriram/Siva; comfortable to start moving.
**Author:** Vijay Ganesan
**Date:** April 16, 2026
**Based on:** Q2 FY26 QBR (Apr 13) + follow-up calls with Sriram/Siva/Manoj (Apr 13-14) + SlideGen walkthrough with Sriram/Siva (Apr 16) + Original PPT Agent Brief (Sep 2025) + Siva's existing SlideGen PRD (Mar 10, 2026) + code walkthrough with Rajesh + **Deck analysis of 32 PET decks across 17 clients + 8 ATU decks across 7 clients** + Galen-PowerPoint Synapse Connector tag model review + workflow-taxonomy MECE audit.

\---

## The Bottom Line

SlideGen is a **library of composable skills + a robust `pptx\\\_utils` Python package** that Claude Code uses to execute any supported consulting workflow — from new deck creation to wave refresh to single-slide client follow-ups to executive summary synthesis.

**Three-layer architecture:**

1. **Skills (SKILL.md files)** — orchestration knowledge: which patterns to use for which workflows
2. **`pptx\\\_utils` package** — composition primitives: brand definitions, layouts, shape builders, chart builders, lxml helpers. Hard-won OOXML knowledge encoded as callable Python functions.
3. **Generated scripts** — \~30-line glue code Claude writes per slide, composing `pptx\\\_utils` calls per the skill's instructions.

**Without `pptx\\\_utils`, Claude rediscovers lxml gaps every session.** Without skills, Claude doesn't know which utilities to compose. Both are building blocks. Both grow over time as workflows demand.

**Workflows drive the inventory.** We enumerate every workflow Galen-Consulting needs to support, then ask: "what skills and utility functions must exist so Claude Code can compose any of these?"

**Key design insight (Apr 14):** Get the slide specification right with intelligent skills UP FRONT, then let deterministic rendering (via `pptx\\\_utils`) produce perfect slides every time. Intelligence flows INTO the spec; deterministic code reads the spec.

By end of Q3, this should work for all major project types (PET, ATU, PCA, HCP-Pt, Digital Trackers) and demonstrate enough versatility that consulting leadership can engage with FY27 staffing decisions.

**Deck analysis grounding:** Reverse-engineering of **32 real PET decks spanning 17 pharma clients** (4,354 charts, 4,950 tables, 3,569 headlines) plus **8 ATU decks across 7 clients** (1,347 charts, 1,539 tables) concretely quantified what the skills and `pptx\_utils` must support. ATU analysis confirmed the same chart-type vocabulary with different frequency distribution (2.8x more stacked-100 composition charts, half the line/scatter usage). See §6.6 for PET findings; see `experiments/deck\_analysis/outputs/atu\_analysis.md` for ATU findings.

\---

## 1\. Problem Statement/context

### 1.1 What's broken with the current Vinoth experiment

|**What's working**|**What's broken**|
|-|-|
|Slide plan generation (hypothesis bank → narrative threads → validated analysis → slide plan)|Slide rendering — reverse-engineered from a J\&J deck, hardcoded patterns, breaks on edge cases (wrong colors, zero-data messages included, misplaced data in tables)|
|Skill-based intelligence layer up to slide\_plan.md|Plumbing from slide plan → config.yaml → PowerPoint deck — too deterministic, no intelligence in the rendering layer|
|Synapse CLI (built standalone by Rajesh)|CLI not yet integrated as a tool into SlideGen|
|Project context, market context, prior wave context, survey context skills|Skills are too interwoven with Vinoth's hypothesis-driven approach — can't extract pieces for other workflows|

**The deeper problem:** "Everything is so interwoven with one specific way of thinking. It's deeply baked into the way the skills themselves are structured." (Sriram, Apr 14). The slide plan skill assumes hypothesis bank, validated analysis, and narrative threads all exist as inputs — but for a wave refresh or a single-slide follow-up, none of that may be needed.

### 1.2 Why a comprehensive skill library (not a prescriptive flow)

Consulting work is fundamentally varied — different project types, different roles, different intent. A single prescriptive flow cannot serve all of these without becoming a maze of branches and exceptions.

The Sagan agents model offers the alternative: build a library of well-defined composable skills, give Claude Code the orchestration responsibility, and let users describe what they need.

> Sriram (Apr 14): "There are a bunch of building blocks and any specific output is a stringing together of those building blocks. In what order to string the building blocks we hand over that decision to Claude Code."

### 1.3 What we are NOT trying to solve in Q3

* **Client-facing platform** — Q3 is internal-first. Typeform-style UX from the agentic MR PRD comes later.
* **Real-time multi-user collaboration** — defer to Q4+.
* **Custom harness / non-Claude-Code interface** — Claude Code is the harness for now.
* **Token cost optimization** — Sriram: "optimizing too early if we focus on token costs now." Build for correctness first.
* **Embedded-in-PowerPoint paradigm** (from original PPT Agent brief) — Claude Code terminal is the current bet.

\---

## 2\. Solution: Skills Composed by Claude Code

### 2.1 Design Principles

**1. Skills + `pptx\\\_utils` are co-equal building blocks.** Skills are orchestration knowledge (when to do what). `pptx\\\_utils` is composition primitives (how to actually produce a shape, chart, table, etc.). Neither works without the other. Both grow over time.

**2. Workflows drive the inventory.** Enumerate workflows first, derive what skills and utilities must exist. Every claimed workflow must be executable as a composition. No exceptions.

**3. Generate-from-scratch, not canonical templates.** `pptx\\\_utils` produces client-delivery quality slides directly via python-pptx + lxml. No dependency on project teams setting up canonical chart templates in hidden slides — that activation energy kills adoption. Brand/layout/pattern definitions live in Python (`BRAND{}`, `LAYOUTS{}`, `CHART\\\_PATTERNS{}`), not in template decks.

**4. Deterministic where possible, intelligent where necessary.** Slide rendering is deterministic (`pptx\\\_utils` calls, scale, cost, reliability). Slide planning is intelligent (skills that reason about data, hypotheses, narratives). Split is clean: intelligence flows into the spec, deterministic code reads the spec.

**5. Slide-level granularity.** The atomic deliverable is one slide. Narrative arcs, hypothesis frameworks, and section structures are layers above the atomic slide.

**6. Reusable skills over bespoke flows.** A skill that serves only one workflow is suspect. Good skills get reused across many workflows.

**7. Visualization selection is deterministic.** Metric tag → Question type default → Human-in-the-loop prompt. Pick chart type by rule, only ask humans on genuine ambiguity.

**8. Specifications are first-class artifacts.** `slide\\\_plan.md` is a contract between the intelligent layer and deterministic layer. Other skills read, write, reason about it.

**9. Human-in-the-loop at decision points, not every step.** Approval gates exist at meaningful boundaries. Within a skill's execution, the agent runs uninterrupted.

**10. Start thin, grow deliberately.** Begin with the most-used `pptx\\\_utils` helpers. Let Claude handle long tail via inline lxml + skill instructions — flagged for later extraction into `pptx\\\_utils`. Library grows based on real usage, not speculation.

### 2.2 Architecture Overview

**Three co-equal layers, plus tools at the bottom:**

```
┌─────────────────────────────────────────────────────────────────┐
│  SKILL LAYER (what Claude reads — orchestration knowledge)       │
│                                                                  │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │  WORKFLOW SKILLS                                          │  │
│  │  (new-deck, refresh, exec-summary, client-followup, ...)  │  │
│  └────────────────────────┬──────────────────────────────────┘  │
│                            │                                     │
│  ┌────────────────────────▼──────────────────────────────────┐  │
│  │  PROJECT-TYPE SKILLS (PET, ATU, HCP-Pt, Digital Tracker)  │  │
│  └────────────────────────┬──────────────────────────────────┘  │
│                            │                                     │
│        ┌───────────────────┼───────────────────┐                 │
│        ▼                   ▼                   ▼                 │
│  ┌──────────┐       ┌──────────┐         ┌──────────┐            │
│  │ ANALYSIS │       │ PLANNING │         │ CREATION │            │
│  │ SKILLS   │       │ SKILLS   │         │ SKILLS   │            │
│  └────┬─────┘       └────┬─────┘         └────┬─────┘            │
│        │                  │                   │                  │
│        └──────────────────┼───────────────────┘                  │
│                           ▼                                      │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │  CONTEXT + DATA SKILLS (project-context, synapse-read,    │  │
│  │                         deck-reader, ...)                  │  │
│  └───────────────────────────────────────────────────────────┘  │
└──────────────────────────┬──────────────────────────────────────┘
                            │  skills reference utilities by name
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│  GENERATED SCRIPT (what Claude writes per slide, \\\~30 lines)      │
│                                                                  │
│   Composes pptx\\\_utils calls per skill instructions.              │
│   Hand-crafted glue, not a framework. Disposable.                │
└──────────────────────────┬──────────────────────────────────────┘
                            │  calls pptx\\\_utils functions
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│  pptx\\\_utils PACKAGE (composition primitives — Python)            │
│                                                                  │
│   brand.py       — BRAND{} — client color/font definitions       │
│   layout.py      — LAYOUTS{} — coordinate presets                │
│   shapes.py      — textbox(), solidrect(), horiz\\\_line(), ...     │
│   charts.py      — make\\\_clustered\\\_bar(), CHART\\\_PATTERNS{}        │
│   tables.py      — delta\\\_table(), value\\\_table()                  │
│   text.py        — format\\\_run(), delta\\\_format(), ...             │
│   images.py      — add\\\_logo(), insert\\\_image()                    │
│   deck.py        — open\\\_template(), clear\\\_slide(), save\\\_deck()   │
│   lxml\\\_helpers.py— set\\\_plot\\\_area\\\_gap(), invert\\\_cat\\\_axis(), ...   │
│   registry.py    — reconcile(), register\\\_shape(), ...            │
│   com.py         — win32com helpers for live editing             │
└──────────────────────────┬──────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────────┐
│  TOOLS (external systems)                                        │
│                                                                  │
│   Synapse CLI · LimeSurvey API · Hashtag API · PowerPoint (COM)  │
│   python-pptx · lxml · Claude Sonnet 4.6 (LLM)                   │
└─────────────────────────────────────────────────────────────────┘
```

**How Claude Code executes a workflow:**

1. Reads user intent → picks a workflow skill
2. Workflow skill invokes project-type skill (e.g., PET)
3. Project-type skill invokes planning/analysis/creation skills
4. Skills instruct Claude what `pptx\\\_utils` functions to compose
5. Claude writes a \~30-line generated script composing those functions
6. Script executes — `pptx\\\_utils` does the deterministic heavy lifting
7. Output: client-delivery-quality `.pptx`

**Why this works:**

* Skills = *knowledge* about what to do (SKILL.md files, Claude-readable)
* `pptx\\\_utils` = *primitives* for doing it (Python, stable, reusable)
* Generated script = *glue* composing them for one specific slide (Claude writes, disposable)

Without `pptx\\\_utils`, Claude writes 200-line scripts rediscovering lxml every session. Without skills, Claude doesn't know which utilities to call or in what order. Both are necessary.

\---

## 3\. Workflows We Need to Support

Derived from Sriram's Apr 14 "hypothetical workflows" exercise + a MECE audit against actual user utterances. Workflows map 1:1 to verbs users say — create, refresh, edit, add, annotate, restructure, audit, summarize. Eight MECE workflows cover the full deck lifecycle.

### Workflow 1: Create deck — `create-deck-workflow`

**Input:** Project brief, raw data (or Synapse setup), KBQs. Optional `mode="hypothesis"` + pre-built `hypothesis\\\_bank.md` for Vinoth's storyboarding flow.
**Output:** Complete first-cut deck following ZoomRx methodology for the project type.
**Modes:** `briefing` (default — data + KBQs → hypotheses → slides) and `hypothesis` (pre-built hypotheses → narrative arcs → slides; the "storyboarding" path).

### Workflow 2: Refresh deck — `refresh-deck-workflow`

**Input:** Prior period's deck (PPTX), new period's data, any client feedback.
**Output:** Next period's deck that preserves narrative continuity from prior period (same slides with refreshed data + deltas highlighted; new slides where warranted; old slides deleted where data is gone). For PET/tracker projects this is the wave refresh; for others it's any multi-slide data refresh.

### Workflow 3: Edit slide — `edit-slide-workflow`

**Input:** Deck, slide ID, action (`rebuild` / `data\\\_refresh` / `edit`). For edit mode, a list of slide-editor instructions.
**Output:** Same deck with the slide regenerated. Rebuild re-renders from current spec; data\_refresh re-pulls from Synapse/Excel first; edit applies whitelisted structural/style changes.

### Workflow 4: Add slide — `add-slide-workflow`

**Input:** Deck, question (or segment-comparison ask), optional reference slide, optional insertion position.
**Output:** One new slide inserted into the existing deck. Handles ad-hoc client questions AND segment comparisons (same output form; different analysis path).

### Workflow 5: Annotate slide — `annotate-slide-workflow`

**Input:** Deck, target slide, callout type (`quote` / `insight` / `freeform`), callout content source.
**Output:** Same slide with a callout inserted inline (quote box, data annotation, insight tag). Does NOT add a new slide — for that, see Workflow 4.

### Workflow 6: Restructure deck — `structural-edit-workflow`

**Input:** Deck, action (`delete` / `reorder` / `split` / `merge`), action-specific params.
**Output:** Same deck with slides removed, reordered, split, or merged. No changes to individual slide content.

### Workflow 7: Audit deck — `deck-audit-workflow`

**Input:** Deck to audit, optional staleness threshold, optional severity filter.
**Output:** Read-only audit report flagging data errors, unsupported claims, stale lineage, missing citations on ES slides, and structural inconsistencies. Report written to `<deck>.audit\\\_<timestamp>.md`; deck is not mutated. Runs as a repeatable pre-delivery QA step (including in CI).

### Workflow 8: Executive summary — `executive-summary-workflow`

**Input:** Full deck, KBQs to answer, optional narrative\_threads.md.
**Output:** 1-3 ES slides with findings citing back to supporting `slide\\\_id`s. Every bullet has `metadata.citations: list\\\[slide\\\_id]`. Unsourced claims rejected.

### Pre-condition: Retroactive Spec Generation

Before any of Workflows 2-8 can run on an *existing* deck (one not created by SlideGen), a one-time bootstrapping step is required: **`deck-reader` reads the existing PPTX and produces slide specs for every slide**, stored in `projects/{name}/context/{wave}/slide\_specs/` alongside a bootstrapped `config.yaml`. This is the "one-time retrofitting" that gives the system backward compatibility with the hundreds of client decks already in use.

How it works:

* For **Connector-tagged shapes**: Tier 1 extraction — reads `ReportConfigHash` → Synapse lineage → populates `DataLineage` with canonical IDs
* For **untagged or unhealthy-tag shapes**: Tier 2 inference — parses headline text, chart pattern, category labels, cross-references against project config
* The resulting `list\[SlideSpec]` + updated `config.yaml` are saved to the project folder and become the canonical inputs for all subsequent workflows

After this one-time step, the deck is "SlideGen-legible" — refresh, edit, audit, and executive summary workflows all operate from the saved specs. Future waves auto-update the specs without re-bootstrapping.

**Priority for existing projects:** Before running Workflow 2 (Refresh) on any live client deck for the first time, run `deck-reader` on the latest wave PPTX. Sriram (Apr 16): *"If we can create a workflow that takes as input an existing PowerPoint deck and creates the associated slide plan, we achieve full backward compatibility through that step."*

### Supported via composition (not separate workflows)

* **Template / brand migration** (swap brand on existing deck) — composable from `edit-slide-workflow` (in edit mode with `set\\\_brand` instructions per slide) + `refresh-deck-workflow` if data changes with brand. Evaluated for Q4 as a dedicated workflow if user demand warrants.
* **Bulk headline revision** (retune all headlines for tighter narrative) — composable by running `edit-slide-workflow` in edit mode across slides in sequence.
* **Variant deck generation** (different client-facing versions) — composable from `edit-slide-workflow` applied in batch.
* **Deck reconcile** (sync registry after manual PowerPoint edits) — already implemented in `slidegen/reconcile.py`; used as a collaboration primitive by multiple workflows.
* **Cross-deck comparison** (diff between waves without refreshing) — composable from `deck-reader` on both decks + ad-hoc Claude reasoning. Evaluate for Q4 if needed.

Synapse platform setup (creating segments, VQs, reporting plans, custom methodologies like RFSOV/MaxDiff) is out of scope for SlideGen — those are `synapse-cli` capabilities. SlideGen consumes the resulting analyses via `synapse-read`.

\---

## 4\. Comprehensive Skill Inventory

Derived from Workflows 1-9. Every workflow must be fully executable as a composition of skills in this list.

Status legend: ✅ Exists · 🔧 Needs refactor · 🆕 New

### 4.1 Context + Data Skills

|Skill|Status|Purpose|
|-|-|-|
|`project-context-builder`|✅|Extract project metadata from brief, KBQs, call notes, methodology docs|
|`market-context-builder`|✅|Build therapy area + competitive landscape context|
|`prior-wave-context-builder`|🔧|Extract story + data from prior wave deck. Needs extension to also extract visualization metadata (chart types, layouts)|
|`survey-context-builder`|✅|Parse survey draft → structured context|
|`synapse-read`|🔧|Pull reports, segments, raw data, banner plans. Needs wrapping around Rajesh's Synapse CLI as tool|
|`hashtag-benchmarks`|🆕|Pull industry benchmarks by metric + therapy area from Hashtag|
|`raw-data-aggregator`|✅|Aggregate respondent-level data (top2box, yes\_pct, etc.)|
|`deck-reader`|🆕|Parse existing PPTX — slides, shapes, data lineage, structure, visualization types. **Dual-mode**: for shapes authored via Galen-PowerPoint Connector, reads `ReportConfigHash` shape tag → Custom XML Part for canonical Synapse lineage (Project/ReportingPlan/Analysis/Segment IDs). For untagged shapes, falls back to structural inference from headline + chart pattern + data shape, with user confirmation. See §6.8.|
|`excel-indexer`|✅|Parse aggregated Excel data (banner plan export) into structured JSON|

### 4.2 Planning Skills

|Skill|Status|Purpose|
|-|-|-|
|`viz-selector`|🆕|Pick chart type via hierarchy: **Metric tag → Question type default → Human-in-the-loop**|
|`slide-plan-generator-hypothesis`|✅|Generate slide plan from hypothesis-driven inputs (Vinoth's flow)|
|`slide-plan-generator-refresh`|🆕|Diff wave N-1 deck vs. new wave data → edit plan|
|`slide-plan-generator-single`|🆕|One ask (e.g., client question) → one slide spec|
|`slide-plan-generator-exec-summary`|🆕|Full deck + KBQs → exec summary slide specs with citations|
|`spec-validator`|🆕|Ensure slide spec has everything renderer needs (catches incomplete specs before rendering)|
|`multi-element-composer`|🆕|Compose table + chart + benchmark + callout into single slide layout|

### 4.3 Creation Skills

|Skill|Status|Purpose|
|-|-|-|
|`slide-creator`|🔧|Render single slide from complete spec via `pptx\\\_utils` composition (THE atomic unit — needs full refactor to be general, not J\&J-specific)|
|`slide-updater`|🆕|Update existing slide with fresh data, preserve structure and formatting|
|`slide-editor`|🆕|Edit specific elements of existing slide (headline, callout, colors, data) — typically via win32com live edit|
|`callout-adder`|🆕|Add data-driven annotation/callout to existing slide|
|`deck-assembler`|🔧|Assemble slides into final PPTX with proper ordering, section breaks, master slide application|
|`executive-summary-writer`|🆕|Generate polished exec summary content with citations back to source slides|

### 4.4 Analysis Skills

|Skill|Status|Purpose|
|-|-|-|
|`hypothesis-generator`|✅|Generate hypothesis bank from KBQs + context|
|`sfea-insight-writer`|✅|PET/SFEA-specific. Two internal phases: Phase 0 validates hypotheses against survey data → `validated\_analysis.md`; Phase 1 synthesizes story arcs (CONVERGENCE/TENSION/DIVERGENCE/CLOSURE) + arc-informed headlines + ES + recommendations → `narrative\_threads.md`. Used by `pet-deck` project skill.|
|`atu-insight-writer`|✅|ATU-specific parallel to sfea-insight-writer. Same two-phase structure but with ATU arc patterns (FUNNEL\_LEAKAGE, SHARE\_MOMENTUM, COMPETITIVE\_CONVERGENCE, BARRIER\_CLUSTER, LOYALTY\_EROSION, SEGMENT\_SPLIT, ADOPTION\_CURVE), share-driven headline framing, and barrier-aware recommendations. Used by `atu-deck` project skill.|
|`segment-comparator`|🆕|Compare segments with significance testing (T-test, chi-square as applicable)|
|`trend-analyzer`|🆕|Trend analysis across waves|
|`stat-sig-annotator`|🆕|Annotate slides with significance markers, low sample footnotes|

### 4.5 Project-Type Skills

Project-type skills orchestrate lower-level skills for a specific methodology.

|Skill|Status|What it encodes|
|-|-|-|
|`pet-deck`|🔧|Build May-Jun. Extract from Vinoth's J\&J-specific implementation. PET deck structure, message testing, prescription intent, rep performance|
|`atu-deck`|🆕|ATU funnel structure, awareness/trial/usage metrics, brand positioning|
|`pca-deck`|🆕|Build as extension of ATU|
|`hcp-pt-deck`|🆕|Stretch (Jul). HCP-Patient research patterns, conversation analysis, treatment journey|
|`digital-tracker-deck`|🆕|Defer to Q4. Digital engagement metrics, channel effectiveness, omnichannel|

### 4.6 Workflow Skills (top-level orchestration)

Each workflow skill maps user intent → composed skill sequence.

|Skill|Status|Triggered by|
|-|-|-|
|`new-deck-workflow`|🔧|"Create a new PET deck for Pfizer Product X"|
|`wave-refresh-workflow`|🆕|"Refresh wave 4 of the Rybrevant PET deck"|
|`single-slide-regen-workflow`|🆕|"Regenerate slide 12 with the fixed data"|
|`slide-update-workflow`|🆕|"Update slide 12 with the latest data"|
|`client-followup-workflow`|🆕|"Client asked about Academic vs Community on slide 35 — create an answer slide"|
|`executive-summary-workflow`|🆕|"Generate an executive summary answering these 3 KBQs"|
|`segment-analysis-workflow`|🆕|"Highlight differences between segments — as callout or new slide"|
|`storyboarding-workflow`|🔧|"Build a hypothesis-driven PET deck" (Vinoth's existing flow, refactored)|

### 4.7 Collaboration \& Persistence Skills

|Skill|Status|Purpose|
|-|-|-|
|`analysis-trace-store`|🆕|Save analysis results within project for team reuse|
|`analysis-trace-retriever`|🆕|Discover + apply colleague's prior analyses without reprocessing|
|`shape-registry-manager`|✅|Track data lineage per slide for audit + refresh|
|`version-manager`|✅|Backup decks, track changes, enable rollback|

### 4.8 Tools (called by skills, not by users directly)

|Tool|Status|Purpose|
|-|-|-|
|Synapse CLI|🔧 (Rajesh built, needs integration)|Programmatic Synapse access (Layer 1 capabilities + Layer 2 endpoints)|
|LimeSurvey API|✅|Read surveys, responses|
|Hashtag API|🆕|Benchmark queries|
|python-pptx|✅|PowerPoint generation|
|Win32COM (PowerPoint)|✅|Live deck editing|
|Claude Sonnet 4.6|✅|LLM for analytical skills|
|Z3 (for spec validation)|🆕|Deterministic spec completeness check|

### 4.9 Totals

* **9 context + data skills** (6 exist, 3 new)
* **7 planning skills** (1 exists, 5 new, 1 refactor)
* **7 creation skills** (2 refactor, 5 new)
* **6 analysis skills** (2 exist, 4 new — includes `atu-insight-writer` parallel to `sfea-insight-writer`)
* **5 project-type skills** (1 refactor, 4 new)
* **8 workflow skills** (1 exists-as-concept, 2 refactor, 5 new)
* **4 collaboration skills** (2 exist, 2 new)
* **7 tools** (4 exist, 1 refactor, 2 new)

**Total: 52 skills + tools.** Of these, \~20 exist in some form today; \~10 need refactoring to be general-purpose; \~22 are new.

\---

## 5\. Workflow → Skill Composition Map

Verifies that every workflow is fully executable as a composition of skills. If a workflow requires a skill not in the inventory, we're missing a building block.

### Workflow 1: Create deck

**Composition:**
`create-deck-workflow` → (`pet-deck` if PET project) → `project-context-builder` + `market-context-builder` + `survey-context-builder` + `prior-wave-context-builder` + `synapse-read` → `hypothesis-generator` (briefing mode) OR use provided `hypothesis\\\_bank.md` (hypothesis mode) → insight-writer (sfea or atu per project type) → `slide-plan-generator-hypothesis` → `viz-selector` + `layout-selector` + `headline-writer` → `spec-validator` → `slide-creator` (×N) → `deck-assembler`

### Workflow 2: Refresh deck

**Composition:**
`refresh-deck-workflow` → `deck-reader` (Tier 1 Connector tags + Tier 2 inference) → `prior-wave-context-builder` → `synapse-read` (new period data) → `slide-plan-generator-refresh` (diff plan) → for updates: `slide-updater`; for adds: `slide-creator` via `viz-selector` + `layout-selector`; for deletes: skip in assembly → optional `trend-analyzer` for cross-period callouts → `deck-assembler`

### Workflow 3: Edit slide

**Composition:**
`edit-slide-workflow` → `deck-reader` (one slide) →

* rebuild mode: `spec-validator` → `slide-creator`
* data\_refresh mode: `synapse-read` (lineage-driven) → `slide-updater` → `spec-validator` → `slide-creator`
* edit mode: `slide-editor` (whitelisted actions) → `spec-validator` → `slide-creator`
→ `deck-assembler` (replace at index)

### Workflow 4: Add slide

**Composition:**
`add-slide-workflow` → `deck-reader` (reference slide for context) → optionally `synapse-read` (new data cuts or segment split) →

* ad-hoc question: `slide-plan-generator-single`
* segment comparison: `segment-comparator` + `stat-sig-annotator` → `slide-plan-generator-single`
→ `viz-selector` + `layout-selector` + `headline-writer` → `spec-validator` → `slide-creator` → `deck-assembler` (insert at position) → optional `analysis-trace-store`

### Workflow 5: Annotate slide

**Composition:**
`annotate-slide-workflow` → `deck-reader` (target slide) →

* quote: `callout-writer` (from `qualitative\\\_data.json`)
* insight: `segment-comparator` + `stat-sig-annotator` → `callout-writer` (data annotation)
* freeform: direct `CalloutComponent` build
→ `slide-editor` (add\_component action) → `spec-validator` → `slide-creator` → `deck-assembler` → optional `analysis-trace-store`

### Workflow 6: Restructure deck

**Composition:**
`structural-edit-workflow` → `deck-reader` (all slides) →

* delete / reorder: rewrite list of `SlideSpec`
* split: split categories or series → 2+ child specs
* merge: combine adjacent specs via `layout-selector` → single merged spec
→ `spec-validator` (on every new/modified spec) → `deck-assembler`

### Workflow 7: Audit deck

**Composition:**
`deck-audit-workflow` → `deck-reader` (all slides) → check categories: data integrity, lineage/audit trail, citation coverage (ES slides), narrative coherence, sample size, structural, chrome/standards → aggregate `list\\\[Finding]` → write `<deck>.audit\\\_<timestamp>.md`
(read-only — does not invoke `slide-creator` or `deck-assembler`)

### Workflow 8: Executive summary

**Composition:**
`executive-summary-workflow` → `deck-reader` (full deck) → optionally `sfea-insight-writer` (if narrative threads missing) → `slide-plan-generator-exec-summary` → `executive-summary-writer` → `spec-validator` → `slide-creator` (×1-3 ES slides) → `deck-assembler` (insert at position)

\---

## 6\. Key Design Decisions

### 6.1 Visualization Selection Hierarchy

Operationalized as the `viz-selector` skill:

```
1. Metric Tag Present?
   YES → Use Metric → Viz mapping (override layer)
   Examples:
     Message Recall → Table + Cluster Bar Chart
     Topic Recall → Table + Cluster Bar + Benchmark Table
     Awareness Trend → Line Chart
     Likelihood to Prescribe → 100% Stacked Column

2. No Metric Tag — Question Type Default?
   YES → Use Question Type → Viz default
   Examples:
     Likert scale → Cluster Bar
     Multi-select → Stacked Bar
     Numeric (distribution) → Histogram / Box Plot
     Ranking → Ranked Bar

3. Neither Available?
   → Human-in-the-loop prompt
```

**Why deterministic-first:** Sriram's principle from Apr 14 — "Use all the intelligence there itself... so the deterministic program can create the perfect slide." Viz selection can be rule-based 90% of the time. Only ask humans for genuine ambiguity.

### 6.2 Rendering: Generate-From-Scratch (not Canonical Templates)

The original PPT Agent brief proposed two approaches: (A) cloning canonical chart templates from hidden slides in the client's template deck, or (B) generating from scratch via python-pptx + rules.

**Decision: Commit fully to Option B (generate-from-scratch).**

Option A was rejected because it requires project teams to maintain canonical template slides per client — a significant setup burden for every new client or wave. Consulting teams won't do it consistently, and the system can't assume they will.

**Implication: `pptx\\\_utils` must be robust enough to produce client-delivery quality without canonicals.** Brand definitions (`BRAND{}`), layout presets (`LAYOUTS{}`), and chart patterns (`CHART\\\_PATTERNS{}`) live in Python, not in template decks. The current system is already Option B — the problem isn't the approach, it's that the renderer was reverse-engineered from one J\&J deck and makes J\&J-specific assumptions. The fix is to make `pptx\\\_utils` general.

**Two rendering paths — both valid:**

* **SlideSpec path (preferred):** Spec → `spec-validator` → `slide-creator` → deterministic PPTX. This is the repeatability path — the same spec produces the same slide every time, reviewable by any team member, reusable across waves.
* **Direct Python path (exploratory):** Claude writes a \~30-line Python script directly assembling `pptx\_utils` calls, without going through a SlideSpec. Valid for one-off slides or exploratory work; does not produce a durable spec for refresh. The generated Python is the auditable artifact in this path. Siva (Apr 16): *"Having that as a flexible layer where you have a lot of these utils reverse engineered, you can put together many of these utils in a quick Python script to create that deck to start with."*

In practice: SlideSpec path for all workflow-driven operations; direct Python path when exploring a new chart type or building outside a tracked workflow.

**Project teams provide minimal assets:**

* Client's template `.pptx` (slide master only — layouts, logos, disclaimers; NO canonical charts required)
* Product → Color Map (JSON/Excel)
* Project Metadata Config (IDs, waves, time periods)

That's it. No hidden slides, no canonicals, no per-client rendering setup.

### 6.3 Specification as Contract

The slide plan is the contract between intelligent skills and deterministic skills. A complete spec includes:

* `slide\\\_id`, `slide\\\_type` (one of N chart types)
* `headline` (talking header)
* `data`: full values, labels, sort order, codes
* `formatting`: colors, fonts, layout positions, period labels
* `extras`: chart-type-specific parameters
* `data\\\_source`: provenance (which question, which segment, which wave)

The `spec-validator` skill enforces completeness before rendering. Renderers fail loudly on incomplete specs rather than inventing fallbacks.

**Specs + config.yaml are the auditability bundle.** Deck, slide specs, and config.yaml should always co-locate in the project's OneDrive folder. They travel together. When you hand off a deck for delivery, the specs + config.yaml travel with it. When a colleague refreshes next wave, they start from the saved specs — not from scratch. Synapse Connector tags are one *way to populate* the spec's data lineage (see §6.8), not a substitute for having specs. Six to twelve months out, every shape-level data connection in the system will flow through the slide spec; Connector tags are a bridge to get there from existing decks. (Sriram, Apr 16)

### 6.4 Claude Code as Harness

Claude Code (Opus 4.6) is the orchestrator. Skills are SKILL.md files in `.claude/skills/`. Tools are programmatic primitives (Python functions, CLI commands, APIs) that skills invoke.

Why Claude Code: it's what we're using today, it's where Anthropic's investment is, and it provides the orchestration we need without building a custom harness.

May evolve to custom harness over time if cost or control demands it. Skills + tools are designed to be harness-portable.

### 6.5 Architecture vs. Ownership

* **Architecture** (this PRD): skills composed by Claude Code
* **Ownership** (Q3 Plan): workflows, not individual skills

  * Pradeep owns the `new-deck-workflow` end-to-end
  * Bharadvaj owns the `wave-refresh-workflow` end-to-end + **evals infrastructure** (see §6.10)
  * Shared skills (like `slide-creator`) are jointly owned with Vijay as tiebreaker
  * **Vijay's role: editor in chief.** Not building everything hands-on. Defining what good looks like, reviewing outputs, unblocking contributors, distributing skill-creation to POs/APOs once the framework is solid. Sriram (Apr 16): *"Your job should become like editor in chief, director in chief, guider in chief. I will become more nervous if the fundamental way this is going to get orchestrated is you sitting with claude code building all of this."*

Workflows are concrete and demonstrable. Skills emerge as artifacts of workflow work.

### 6.6 Deck Analysis Findings (32 PET decks + 8 ATU decks)

A reverse-engineering exercise analyzed every chart, table, headline, and coordinate in 32 real client PET decks. Full methodology + raw data in `experiments/deck\\\_analysis/outputs/ACTIONABLE\\\_FINDINGS.md`. Key findings that drive this PRD:

**The 80/20 on chart types.** Top 6 chart types cover 88% of all client charts:

|Chart type|Occurrences|%|
|-|-|-|
|`bar\\\_clustered` (horizontal)|1,511|35%|
|`xy\\\_scatter` (abacus)|743|17%|
|`line\\\_markers` (trended)|595|14%|
|`column\\\_stacked\\\_100` (vertical)|486|11%|
|`bar\\\_stacked\\\_100` (horizontal)|324|7%|
|`bar\\\_stacked` (horizontal)|305|7%|

**Deterministic-first hypothesis validated.** Charts in real decks have almost no features that would require per-chart intelligence:

* Chart title: present in only **1%** of charts (headline lives in separate text box)
* Chart legend: present in only **1%** (companion table serves as legend)
* Major gridlines: absent on **84%**
* Data label number format: `"0%"` on **96%** of labels
* Category axis inverted (maxMin): **43%** of charts — most common "non-default"
* Category labels hidden (`tickLblPos="none"`): **371 charts** — companion table pattern

This means the slide plan spec can be **complete enough that deterministic rendering always produces correct output**. Renderers don't need to guess.

**`pptx\\\_utils` inventory confirmed.** The OOXML properties that python-pptx does not expose — and that recur in real decks — are enumerable. A fixed list of \~20 `lxml\\\_helpers` functions covers 100% of observed needs. See `pptx\\\_utils/lxml\\\_helpers.py` (branch `vijay-slidegen`) for the current implementation.

**Canonical templates approach confirmed rejected.** No evidence across 32 decks that project teams maintain canonical chart templates in hidden slides. Rendering defaults live in `pptx\\\_utils`; brand-specific differentiation lives in `BRAND{}` alone.

**BRAND{} and LAYOUTS{} populated from observation.** 17 client brand entries with series colors, fonts, and heading colors have been generated into `pptx\\\_utils/brand.py`. 6 layout presets with median coordinates from signature clusters (145 slides for `1\\\_chart\\\_1\\\_table`, 110 for `1\\\_chart\\\_2\\\_table`, etc.) have been generated into `pptx\\\_utils/layout.py`. Both live on branch `vijay-slidegen`.

**Full-corpus regrounding via mass scan.** The 40-deck analysis established the initial baseline. `experiments/deck\_analysis/mass\_deck\_scanner.py` runs the same analysis at 10x scale across 400-500 client decks from the SharePoint archive (organized as `client/project/deck.pptx`). It extracts chart types, layout coordinates, brand colors, headline patterns, and component compositions, then **directly regenerates** `BRAND{}`, `LAYOUTS{}`, and `CHART\_PATTERNS{}` Python source from the full corpus — not a delta report, but a complete re-grounding. The generated files replace the current `pptx\_utils` modules after review. Sriram (Apr 16): *"When you build bottom up, it will get closer to exhaustive."* The scanner is resumable (saves every 25 decks), classifies decks by project type (PET/ATU/HCP-Pt/Digital Tracker/PCA/etc.) for coverage analysis, and the `experiments/deck\_analysis/visual\_regression.py` harness measures fidelity after each update.

### 6.7 Slide Spec as the Contract

The slide spec is a first-class artifact — the interface between intelligent planning skills and deterministic renderers. Implemented in `slidegen/slide\\\_spec/`:

* **`slide\\\_spec/schema.py`** — Python dataclasses defining `SlideSpec` and its components (`ChartComponent`, `LabelTableComponent`, `ValueTableComponent`, `DeltaColumnComponent`, `CalloutComponent`, `ImageComponent`, `TextboxComponent`). JSON serde via `load\\\_spec()` / `dump\\\_spec()`.
* **`slide\\\_spec/validator.py`** — `validate\\\_spec(spec)` returns list of errors; `validate\\\_spec(spec, strict=True)` raises. Checks completeness, layout/pattern/brand existence, color-token syntax, cross-component row-count coherence, position bounds.

**Primary fidelity axes** (in priority order — slide-creator is measured against these):

1. **Layout** — shape positions match observed real-deck clusters (≤0.1" drift). Validated against the 30 coordinate clusters from the Apr 15 deck analysis.
2. **Visualization** — chart pattern matches one of the 10 `CHART\\\_PATTERNS` keys. Top 6 cover 88% of real charts; these are the hard P0 bar.
3. **Data** — categories × series shape matches; values render to the exact numeric labels expected.
4. **Brand colors** — resolvable at render time, *not hardcoded into the schema*. Four resolution sources:

   * Explicit hex: `"#F75824"`
   * `BRAND{}` token: `"{brand.primary\\\_current}"`
   * Context file token: `"{context.brand\\\_palette.primary}"` (resolved from `market\\\_context.md` or `project\\\_context.md`)
   * Deck-reader extraction: `"{deck.slide\\\_4.series\\\_0.color}"` (pulled from a prior wave PPTX by `deck-reader`)

Tokens are resolved by `slide-creator` at render time, not at spec creation time — which means the same spec is portable across brands and refresh cycles without rewriting.

**Why this matters.** Multiple skills can now produce a valid spec from different angles: `slide-plan-generator-hypothesis` from a hypothesis bank, `slide-plan-generator-refresh` from a deck-reader diff, `slide-plan-generator-single` from a single client question. Any valid spec is rendered identically by `slide-creator`. No bespoke glue. No workflow-specific renderer branches.

### 6.8 Data Lineage: SlideSpec as the Long-Term Model

**Long-term vision:** Every deck travels with its slide specs + config.yaml (§6.3). The spec is the authoritative source of data lineage — what analysis, what wave, what segments produced every chart on every slide. This does not depend on the Galen-PowerPoint Connector.

**Short-term bridge:** The existing **Galen-PowerPoint Synapse Connector** already stamps Synapse configuration onto shapes in decks authored through the add-in. SlideGen **leverages these tags as the preferred source of data lineage when bootstrapping specs from existing Connector-tagged decks** — but does not require them, and does NOT trust them blindly. `deck-reader` is two-tier with explicit tag health checks:

**Tier 1 — Trusted tag (preferred).** For shapes carrying a Connector tag that passes all 5 health checks:

* Shape tag `ReportConfigHash` → SHA256 key into a Custom XML Part holding the `ReportConfigDto` JSON
* The config DTO provides canonical identifiers: `ProjectId`, `ReportingPlanId`, `AnalysisIds\\\[]`, `SurveyId`, `SegmentIds\\\[]`, `StaticTimePeriodIds\\\[]` or `DynamicTimePeriod{LatestNDeliverables, IncludeLiveWave}` (deliverables), `AnalysisType`
* Additional shape tags: `DataFrameConfigHashTag` (pivot config), `MappingConfig` (field-to-visual mapping), `LastRefreshTime`, `ColumnKeyLabelMap`, `RefreshErrorMsgTag`
* These fields populate `DataLineage` directly — no inference required

**Tag health check (5 steps)** — any failure routes the shape through Tier 2:

1. Custom XML Part resolves (hash → parseable JSON)
2. Synapse IDs still exist (reporting plan, analyses, segments, deliverables resolve today)
3. DTO schema compatible with current Python dataclass (no legacy incompatibilities)
4. Structural consistency — returned data shape matches chart's current dimensions (±1 row / ±1 column tolerance)
5. Mapping plausibility — `MappingConfig` field names match actual chart data binding

**Tier 2 — Structural inference (fallback).** For untagged shapes AND shapes whose tag failed any health check:

* Parses headline text, chart pattern, category labels, and embedded table content
* Cross-references against project config (`config.yaml`, `source\\\_data.json`) to propose likely extraction method + question codes
* Surfaces inference confidence with the spec; user confirmation only on `confidence="low"` inferences (driven by inference quality, NOT tag status)
* Writes best-effort `DataLineage` into legacy fields (`data\\\_source`, `extraction\\\_method`, `question\\\_codes`, `source\\\_file`)
* Failed-tag content preserved in `metadata.original\\\_tag\\\_lineage` for audit trail (not used for refresh)

**Why two tiers, not three.** An earlier design had a "suspect" middle tier that required user confirmation per unhealthy tag. That's unusable in practice — a 40-slide deck with 10 suspect tags means 10 confirmations before refresh starts. Worse, tag-suspect cases are the ones where the chart has been manually edited — Tier 2 inference from the chart's actual content is a better source than a lineage the chart has drifted from.

**Why blind-trust is dangerous.** Stale tags can point to analyses that have since been deleted, reporting plans that no longer exist, or data shapes that don't match the chart after manual editing. Galen-PowerPoint users have seen refreshes destroy chart layouts when refreshing against such tags. SlideGen must not repeat that failure mode.

**User overrides:**

* `--trust-tags` — skip health checks 2-5 (keep check 1 for parseability); use every parseable tag as Tier 1
* `--ignore-tags` — bypass Tier 1 entirely; every shape goes through Tier 2

**Why not rely on tags alone.** Not every existing deck was produced through the Connector. Older decks, manually-edited slides, and decks from other tools coexist with tagged decks in a typical client workspace. The two-tier design means every refresh workflow works against any deck while still getting the auditability benefits where the tags exist AND pass validation.

**Refresh workflow implications:**

* **Tier 1 path**: deterministic — `deck-reader` → `DataLineage` with `reporting\\\_plan\\\_id`+`analysis\\\_ids` → `synapse-cli` fetches new data → `slide-updater` re-renders → tags updated with new `LastRefreshTime`
* **Tier 2 path**: best-effort — `deck-reader` → inferred `DataLineage` with `question\\\_codes`+`source\\\_file` → Excel extraction via existing `data\\\_loaders` → `slide-updater` re-renders → user shown before/after for confirmation

**Audit trail.** Every refreshed slide (either tier) ends up with `last\\\_data\\\_pull` (and `last\\\_refresh\\\_error` if applicable) in its spec lineage, mirroring the Connector's `LastRefreshTime`/`RefreshErrorMsgTag` pattern. The `shape\\\_registry.json` + spec together form the audit record.

See Galen-PowerPoint repo: `Docs/Export Import Tags - PRD.md` (Connector DTO-to-Excel column mapping), `Constants.cs` (shape tag constants), `Services/ShapeConfigurationServiceBase.cs` (hash-based config read). Reference these when implementing `deck-reader` Tier 1.

### 6.9 Evals Strategy

Sriram identified evals as the first major bottleneck SlideGen will hit (Apr 16): *"The bottlenecks you'll run into are testing, which you solve by solving evals, and the number of skills you need."* Without evals, every change requires manual verification — which doesn't scale once contributors across multiple project teams are landing skills simultaneously.

**Owner: Bharadvaj** (Vijay defines what good looks like; Bharadvaj builds the dataset and harness)

**Eval approach:** Use prior wave decks as ground truth.

1. Take a prior-wave PPTX where the "correct" output is known (e.g., JJ RYB Q4 '25 deck)
2. Run retroactive spec generation (`deck-reader`) on it → spec bundle
3. Modify a known set of data points in the source Excel (e.g., +3pp on a few metrics)
4. Run `refresh-deck-workflow` against the modified data
5. Compare output to expected: correct headlines, correct chart values, correct period labels, no stale data, correct layout

**Eval dimensions:**

* Data fidelity — output values match source data exactly
* Headline freshness — no stale headlines from prior wave (§6.10 principle)
* Layout stability — no unexpected layout shifts when data shape is unchanged
* Lineage completeness — `shape\_registry.json` has `last\_data\_pull` stamped on all refreshed slides
* Regression — slides that should NOT change are unchanged

**Where evals live:** `tests/evals/` — each eval is a directory with `input/` (prior wave PPTX + modified Excel), `expected/` (expected output metadata — not full PPTX), and a `run\_eval.py` that compares actual vs expected.

### 6.10 Headlines must refresh with data

**Principle:** when a slide's data changes, its headline (and period-stamped subheadline) must be regenerated — never left stale. A headline like "Efficacy recall dipped 3pp QoQ" becomes a client-delivery risk when new data says "up 2pp".

**Implementation:**

* `slide-updater` invokes `headline-writer` by default after updating data. Caller passes `preserve\\\_headline=True` only when they've explicitly reconciled the text against new numbers (e.g. a wording-only edit where data is unchanged).
* Subheadline period references ("Q3 '25 vs Q4 '25") are rewritten when period labels shift.
* `refresh-deck-workflow` regenerates `narrative\\\_threads.md` from new data *before* invoking `slide-updater` per slide — so updated headlines inherit from a fresh narrative backbone, not last wave's arcs.
* `deck-audit-workflow` flags any slide whose headline text contradicts the rendered data (BLOCKER severity).

**Why this is explicit:** a previous iteration of `slide-updater` preserved headlines by default. That was wrong. Data + stale headline is the most visible kind of error in a client-ready deck and the hardest to spot in QA because the chart looks fine.

\---

## 7\. What We Keep vs. What We Discard

### From the Vinoth experiment, KEEP:

* Slide plan generation skills (`slide-plan-generator-hypothesis`) through narrative\_threads, validated\_analysis, slide\_plan — Vinoth is happy with these
* Context-building skills (project, market, prior wave, survey) — foundational and reusable
* Skills framework pattern (CLAUDE.md, .claude/skills/, project folder structure)
* Synapse CLI (Rajesh's work — wrap as tool)
* Data extraction patterns (banner plan, raw data, JSON-first)
* The 22 chart renderers as a starting point (but each needs to be made general)

### From the Vinoth experiment, DISCARD:

* The mini-harness within Claude Code that prescribes the five-step flow
* Hardcoded slide renderer patterns (reverse-engineered from one deck)
* Assumption that hypothesis bank + narrative threads are always required inputs
* Tightly-coupled skills that assume the entire upstream chain has executed

### From the original PPT Agent brief, INCORPORATE:

* **Visualization Selection Hierarchy** (Metric > Q-type > HITL)
* **Multi-Element Assembly** pattern (pick layout → generate elements → populate → assemble)
* **Analysis Trace Persistence** for team collaboration
* **Global Defaults by Question Type** for viz selection fallback
* 5 additional workflows (slide update, callout, segment slide, exec summary, custom analysis setup)

### From the original PPT Agent brief, DISCARD / DEFER:

* **Canonical Templates approach (Option A)** — explicitly rejected. Too much setup burden on project teams. Committed to generate-from-scratch (Option B).
* **Embedded-in-PowerPoint paradigm** (Synapse Agent living inside PowerPoint) — Claude Code terminal is current bet
* **Client template deck as primary entry point** (vs. Claude Code terminal)
* **"Add Callout / Add as New Slide / No Action" hard-coded menu** — Claude Code asks naturally in conversation

\---

## 8\. Technical Architecture

### 8.1 Repository Structure (Target)

```
galen-consulting-r3m-report/
├── .claude/
│   ├── CLAUDE.md                    # Orchestration instructions
│   └── skills/
│       ├── workflows/               # Top-level workflow skills
│       │   ├── new-deck/
│       │   ├── wave-refresh/
│       │   ├── client-followup/
│       │   ├── executive-summary/
│       │   └── ...
│       ├── projects/                # Project-type skills
│       │   ├── pet/
│       │   ├── atu/
│       │   ├── pca/
│       │   └── ...
│       ├── planning/                # Planning skills
│       │   ├── viz-selector/
│       │   ├── spec-validator/
│       │   └── ...
│       ├── creation/                # Creation skills
│       │   ├── slide-creator/
│       │   ├── slide-updater/
│       │   └── ...
│       ├── analysis/                # Analysis skills
│       ├── context-data/            # Context + data skills
│       └── collaboration/           # Analysis trace, version mgmt
├── slidegen/
│   ├── pipeline/
│   │   ├── orchestrator.py
│   │   ├── slide\\\_renderers/         # General-purpose renderers (22 types)
│   │   ├── data\\\_loaders.py          # Four-track data access
│   │   ├── config\\\_generator.py      # scaffold\\\_config\\\_from\\\_plan()
│   │   └── ...
│   ├── pptx\\\_utils/                  # ★ Composition primitives (co-equal building block)
│   │   ├── brand.py                 # BRAND{} — per-client color/font defs
│   │   ├── layout.py                # LAYOUTS{} — coordinate presets
│   │   ├── shapes.py                # textbox, solidrect, horiz\\\_line, ...
│   │   ├── charts.py                # CHART\\\_PATTERNS{}, chart builders
│   │   ├── tables.py                # delta\\\_table, value\\\_table
│   │   ├── text.py                  # format\\\_run, delta\\\_format, ...
│   │   ├── images.py                # add\\\_logo, insert\\\_image
│   │   ├── deck.py                  # open\\\_template, clear\\\_slide, save\\\_deck
│   │   ├── lxml\\\_helpers.py          # set\\\_plot\\\_area\\\_gap, invert\\\_cat\\\_axis, ...
│   │   ├── registry.py              # reconcile, register\\\_shape, ...
│   │   └── com.py                   # win32com helpers for live editing
│   ├── slide\\\_spec/                  # ★ the spec contract (schema + validator)
│   │   ├── schema.py                # SlideSpec + components (dataclasses)
│   │   ├── validator.py             # validate\\\_spec() — completeness + semantics
│   │   ├── \\\_\\\_init\\\_\\\_.py              # Re-exports for external use
│   │   └── examples/                # Canonical example specs per chart pattern
│   ├── deck\\\_reader/                 # parse existing decks (dual-mode per §6.8)
│   │   ├── tag\\\_reader.py            # Tier 1 — Connector tag extraction
│   │   └── inference.py             # Tier 2 — structural inference fallback
│   ├── viz\\\_selector/                # Metric > Q-type > HITL (deterministic)
│   ├── create.py                    # SlideBuilder class (python-pptx creation)
│   ├── edit.py                      # LiveEditor class (win32com live editing)
│   └── reconcile.py                 # Registry reconciliation from live PPT
└── projects/
    └── {project-name}/              # User-facing project workspaces
```

### 8.2 Skill Authoring Pattern

Every skill has a SKILL.md with:

* `description`: when to trigger
* `inputs`: required and optional
* `outputs`: what it produces
* `tools`: which tools it calls
* `composed\\\_skills`: which other skills it invokes
* `example\\\_invocations`: sample usage
* `pitfalls`: common failure modes

### 8.3 Project Setup Assets (minimal)

Project teams provide upfront — and only these:

* **Client Slide Master Deck** (required) — client-provided template with layouts, logos, disclaimers. Slide master only — NO canonical chart templates required.
* **Product→Color Map** (required) — JSON/Excel mapping brand names (+ aliases) to RGB/hex colors. Added to `BRAND{}` in `pptx\\\_utils/brand.py`.
* **Project Metadata Config** (required) — Synapse `project\_id`, `reporting\_plan\_id`, `deliverable\_ids` (static period IDs) or dynamic period config (`latest\_n\_deliverables` + `include\_live\_wave`), `analysis\_ids`, industry-average filter IDs. These are the same identifiers the Galen-PowerPoint Connector stamps onto shapes (§6.8) — SlideGen reads them from the config when generating a new deck, and reads them from shape tags when refreshing an existing one.

Assets live in the project workspace. Brand colors become a `BRAND{}` entry in `pptx\\\_utils`; layouts are Python constants in `pptx\\\_utils/layout.py`.

**What project teams do NOT provide:**

* No canonical chart templates in hidden slides
* No chart pattern `.crtx` files
* No theme `.thmx` files
* No per-slide-type rendering hints

All rendering knowledge lives in `pptx\\\_utils` (Python). Project teams provide branding + data, not rendering setup.

### 8.4 Tool Contracts

Skills invoke tools via stable interfaces. Example:

```python
# synapse-read skill calls Synapse CLI tool
synapse\\\_cli.report\\\_pull(analysis\\\_id=99812, quarter="Q1-2026")
```

Tools abstract external systems. Skills don't know about auth, retries, polling — the tool handles it.

\---

## 9\. Q3 Scope

**Target:** All 8 workflows working end-to-end across PET + ATU project types by end of Q3, with HCP-Pt + Digital Tracker + PCA project skills built in parallel by their respective project-team contributors using the same building blocks. Not an MVP — a **fully functional product across all core project types**. Deferring entire workflows to Q4 is not acceptable; sequencing *within* workflows (core paths before edge cases) is.

### 9.1 Building-block-first delivery

Primitives are the investment; workflows are the composition. Once the primitives are solid, each workflow orchestrator is hours of work, not weeks. Three strands of work run in parallel:

**Rendering fidelity** — the atomic unit. Every workflow composes into it.

* `slidegen/slide\_spec/` — spec schema + validator (DONE)
* `slide-creator` Python renderer — takes validated spec, produces one client-ready slide. Zero intelligence. (DONE; line/doughnut Repair bug outstanding)
* Visual regression harness against the 30 coordinate clusters from Apr 15 deck analysis (≤0.1" drift on top-6 chart patterns)
* Top-6 chart patterns (88% of real PET charts) proven to client-delivery parity

**Edit-mode primitives** — unlock every workflow that starts from an existing deck.

* `deck-reader` (dual-mode per §6.8 — Tier 1 Connector-tag-driven + Tier 2 structural inference)
* `slide-updater` (data refresh + headline regen per §6.9)
* `slide-editor` (whitelisted actions)
* `deck-assembler` (insert/replace/reorder/delete)

**Spec producers** — intelligence into the spec.

* `viz-selector` — deterministic Metric → Q-type → HITL hierarchy (§6.1)
* `headline-writer`, `callout-writer`, `layout-selector`
* `slide-plan-generator-refresh` / `-single` / `-exec-summary`
* Context builders (market / prior-wave / survey / project)
* Analysis skills (`segment-comparator`, `stat-sig-annotator`, `trend-analyzer`)

Once these primitives are feature-complete, workflow orchestration is trivial — each workflow SKILL.md is a short composition of already-working parts.

### 9.2 Workflow shipping

The 8 workflow orchestrators ship at \~2/week once primitives are solid. Ship order reflects risk and strategic value:

1. **`edit-slide-workflow`** — end-to-end shake-down of the full stack (rebuild mode first, then data\_refresh + edit modes)
2. **`refresh-deck-workflow`** — primary Q3 demo; exercises deck-reader Tier 1 + Tier 2 + slide-updater + headline-writer
3. **`create-deck-workflow`** — full briefing + hypothesis modes on new primitives
4. **`add-slide-workflow`** — client-followup + segment-new-slide paths
5. **`annotate-slide-workflow`** — callouts on existing slides
6. **`structural-edit-workflow`** — delete / reorder / split / merge
7. **`deck-audit-workflow`** — pre-delivery QA
8. **`executive-summary-workflow`** — uses full-deck deck-reader + narrative-threads + citation-gated bullets

### 9.3 Q3 Milestones

|Date|Milestone|Definition of Done|
|-|-|-|
|**Apr 15 — DONE**|Deck analysis + `pptx\_utils` foundation|32 PET decks + 8 ATU decks analyzed; 18 BRAND entries, 11 LAYOUTS, 10 CHART\_PATTERNS, 9 new `lxml\_helpers` landed|
|**Apr 16 — DONE**|Spec contract + slide-creator Python|`slidegen/slide\_spec/` + `slidegen/slide\_creator.py`; all 10 chart patterns render; 20 canonical example specs covering \~55% of real PET slide compositions|
|**Apr 16 — DONE**|Edit-mode + spec-producer primitives|`deck-reader` (dual-mode), `slide-updater`, `slide-editor`, `deck-assembler`, `viz-selector`, `layout-selector`, `headline-writer`, `callout-writer`, all `slide-plan-generator-\*`, all 3 analysis skills (`segment-comparator`, `stat-sig-annotator`, `trend-analyzer`)|
|**End of Apr**|Rendering fidelity complete + evals bootstrapped|Visual regression passing on top-6 chart patterns × representative brands. Line/doughnut Repair bug closed. Connector-tag integration tested on real tagged PET deck. **Evals harness started** (`tests/evals/`) with ≥1 refresh eval using prior JJ RYB deck.|
|**Mid-May**|First workflow end-to-end|`edit-slide-workflow` (rebuild mode) produces a refreshed slide on a real PET deck, audit trail intact.|
|**End of May**|Primary demo: refresh deck (PET)|Full `refresh-deck-workflow` on a real PET brand, both Connector-tagged and untagged refresh paths demonstrable. Headline regen working (§6.10). Evals covering refresh + headline freshness passing.|
|**Mid-Jun**|5 of 8 workflows live for PET|Create, Refresh, Edit-slide, Add-slide, Annotate.|
|**End of Jun**|All 8 workflows live for PET + ATU|Restructure, Audit, Executive Summary added. `pet-deck` + `atu-deck` project skills both working. Dogfooded on ≥2 live PET projects + 1 ATU.|
|**End of Q3 (Jul)**|Broader project coverage + polish|HCP-Pt + Digital Tracker + PCA project skills contributed in parallel by project-team owners using the established building blocks. Edge-case sweep. Consulting leadership demo across all 5 project types.|

### 9.4 Out of Scope for Q3

* **Canonical templates approach** — explicitly rejected (see §6.2).
* **Synapse platform capability building** (new segments / VQs / reporting plans / RFSOV / MaxDiff endpoints inside SlideGen) — owned by `synapse-cli` team. SlideGen *consumes* what synapse-cli produces via the `synapse-read` skill; integration with synapse-cli itself is IN scope for Q3 (core data-access dependency).
* **Client-facing UI** — SlideGen runs through Claude Code for Q3.
* **Token cost optimization at scale** — Sriram flagged as later concern.

### 9.5 Parallel project-skill delivery (HCP-Pt, Digital Tracker, PCA)

HCP-Pt, Digital Tracker, and PCA are **target Q3 deliverables** — built in parallel by the respective project-team contributors using the same building blocks as `pet-deck` and `atu-deck`. SlideGen's job is to prove the framework is solid enough that each project type is a thin project-skill (a `projects/<type>-deck/` SKILL.md that applies methodology-specific defaults), not a separate codebase. If the framework requires non-trivial changes to support HCP-Pt or DT, that's a framework bug we want surfaced during Q3, not after.

**Dependency:** this hinges on project-team contributors being identified, made available, and onboarded to the building blocks in the first half of Q3. Specifically:

* **Contributor identification** — who owns the project skill for each of HCP-Pt / DT / PCA? Not yet assigned. Needs Sriram/Siva alignment before mid-May.
* **Onboarding ramp** — a contributor who's never written a SlideGen skill before needs \~1 week of hands-on time with `pet-deck` to understand the pattern + building blocks.
* **Core team support** — Vijay/Bharadvaj/Pradeep need bandwidth to unblock contributors when they hit framework gaps. Estimated 20-30% of core-team time in weeks 8-13.

**If contributors aren't available in time:** Q3 target falls back to "PET + ATU first-party complete, with HCP-Pt / DT / PCA skills scaffolded and ready for project-team landing in early Q4." That's still a better bar than the original "PET-only Q3" framing — but not the full 5-project Q3 target. The PRD names the full target explicitly so the resourcing dependency is visible early.

\---

## 10\. Risks \& Open Questions

### 10.1 Risks

|Risk|Mitigation|
|-|-|
|Skill inventory keeps growing as we discover new needs|Tie every skill to a specific workflow. If no workflow needs it, don't build it.|
|Skills too abstract / no users|Workflows are concrete. Build skills to satisfy workflows, not for hypothetical reuse.|
|Slide creator refactor takes longer than expected|Prioritize 5-6 most common chart types first. Long tail follows.|
|Wave refresh harder than new deck (backward compat)|Run parallel tracks — don't block new deck on refresh learnings.|
|`viz-selector` rules get complex as edge cases emerge|Start with 10-15 most common metric/q-type mappings. Extend incrementally.|
|Token cost at scale|Sriram said later concern. May switch models by skill type (Sonnet for rendering, Opus for planning).|
|Skills become tightly coupled (repeat the prior-experiment mistake)|Explicit review: every skill declares inputs, outputs, composed\_skills in SKILL.md. No hidden assumptions. spec-validator enforces the contract at every boundary.|

### 10.2 Open Questions

1. **How do we version skills + `pptx\\\_utils`?** Git-based, but how do users opt into new versions? OneDrive-sync model from existing docs?
2. **Skill registry mechanism:** does Claude Code auto-discover skills from directory structure, or explicit manifest?
3. **How does `viz-selector` remember per-project overrides?** If a user overrides viz-selector's default for a given metric on one slide, should that override apply project-wide (persist as a local override to `METRIC\_TAG\_MAP`) or only to that slide? Currently per-slide-only via `edit-slide-workflow`. Project-level persistence is a possible future enhancement.
4. **What's the contract for `deck-reader` on client-branded decks?** Does it round-trip perfectly, or lose fidelity?
5. **Should `analysis-trace-store` persist in the deck itself (shape metadata) or in a separate project-level store?**
6. **For `executive-summary-writer`: how do we constrain hallucination?** Guardrails on citations, confidence indicators.
7. **When does Claude write inline lxml vs. call `pptx\\\_utils`?** Siva's "start thin, grow deliberately" principle — at what threshold does a repeated inline pattern get extracted into the library?

### 10.3 Resolved questions

Concretely answered during PRD iteration — captured here so the rationale isn't lost:

* ✅ **When does `slide-creator` render from scratch vs. use canonical templates?** → Fully from-scratch. The Apr 15 deck analysis found no evidence of hidden canonical templates in any of the 32 decks. Brand differentiation is achievable via `BRAND{}` + `LAYOUTS{}` Python constants alone.
* ✅ **How does `viz-selector` resolve ambiguity?** → Deterministic rule: (1) 36 pre-mapped pharma metric tags → chart pattern (from Apr 15 analysis of 4,354 real PET charts — full table in §6.1); (2) question type default for unrecognized metrics (likert → bar\_clustered\_horizontal, etc.); (3) `"UNRESOLVED"` sentinel when neither path applies, which `spec-validator` rejects, forcing the planner to surface an HITL prompt. **Users can always override** viz-selector's default via `edit-slide-workflow` edit mode with the `set\_chart\_pattern` action — the selector provides a default, not a prescription.
* ✅ **Test strategy for `pptx\\\_utils` regression?** → Use the 30 coordinate clusters + per-chart OOXML inventory as a golden reference. For each renderer, compare against the corresponding real-deck signature cluster. Harness proposed in `experiments/deck\\\_analysis/outputs/ACTIONABLE\\\_FINDINGS.md`.
* ✅ **How are brand colors sourced?** → Resolvable at render time from 4 sources: explicit hex, `BRAND{}` token, context-file reference (`market\\\_context.md` / `project\\\_context.md`), or deck-reader extraction from prior wave PPTX. `BRAND{}` is a convenience default, not a requirement. The spec carries color tokens; `slide-creator` resolves them. See §6.7.
* ✅ **What is the spec-as-contract implementation?** → `slidegen/slide\\\_spec/` subpackage with `schema.py` (dataclasses) + `validator.py`. See §6.7.
* ✅ **How do we extract data lineage from existing decks for refresh workflows?** → Dual-mode `deck-reader`. Tier 1 reads Galen-PowerPoint Connector tags (`ReportConfigHash` → Custom XML Part) for canonical Synapse lineage. Tier 2 falls back to structural inference with user confirmation for untagged shapes. Not every deck is Connector-authored, so Tier 2 is required. See §6.8.
* ✅ **What's the priority order for fidelity?** → Layout, then visualization, then data, then brand colors. Brand colors are parameterized inputs; the first three axes are where `slide-creator` must deliver pixel/numeric parity with real decks. See §6.7.
* ✅ **Q3 target — PET-only or broader?** → PET + ATU land first-party (weeks 1-9); HCP-Pt + Digital Tracker + PCA project skills land in parallel via project-team contributors (weeks 9-13), using the same building blocks. All 8 workflows × all 5 project types by end of Q3, with dogfooding on ≥2 live PET projects + 1 ATU. See §9.3 + §9.5.
* ✅ **Are the workflows MECE?** → Yes, after Apr 16 audit. Eight workflows map 1:1 to user verbs (create / refresh / edit / add / annotate / restructure / audit / summarize). Earlier drafts had redundancy (two create paths, two edit paths, two add paths) that collapsed into the current taxonomy. Structural-edit and deck-audit were added to close gaps the earlier lists missed.

\---

## 11\. Appendix

### 11.1 Relationship to Other Docs

* **Q3 Plan** — execution plan and team structure. This PRD is the product definition for SlideGen.
* **Survey Design PRD** — parallel doc for the survey workstream. Both follow the same skills-based building blocks approach.
* **Synapse CLI Proposal** — defines the data access layer SlideGen depends on. `synapse-read` wraps this CLI as a tool. Platform setup tasks (create segments, VQs, reporting plans, methodology configs) are the CLI's own responsibility, invoked directly by users outside SlideGen.
* **Galen-PowerPoint Synapse Connector** — the existing PowerPoint add-in that stamps ReportConfig/PivotConfig/MappingConfig tags onto shapes. SlideGen's `deck-reader` leverages these tags as the Tier 1 (preferred) source of data lineage for refresh workflows. See §6.8. Key references: `Docs/Export Import Tags - PRD.md`, `Constants.cs`, `Services/ShapeConfigurationServiceBase.cs`.
* **Original PPT Agent Brief (Sep 2025)** — source of several workflows (edit, add, segment analysis, executive summary) and architectural concepts (viz hierarchy, multi-element assembly, analysis traces). Superseded in modality (embedded-in-PPT → Claude Code terminal) but many ideas survive.
* **Siva's SlideGen PRD (Mar 10, 2026)** — at `docs/SlideGen\_PRD.md`. The implementation blueprint (four-track data layer, win32com live editing, shape registry reconciliation, audit chain). This PRD (v1.1) is strategic direction; Siva's is the technical blueprint. Complementary, not contradictory.
* **Existing SlideGen Workflow** — at `galen-consulting-r3m-report/docs/slidegen\\\_workflow.md`. Current 8-stage pipeline documentation.
* **Deck Analysis** — at `experiments/deck\_analysis/` on branch `vijay-slidegen`. Two rounds: (1) 32 PET decks across 17 clients (Apr 15) — produced ACTIONABLE\_FINDINGS.md, deep\_report.md, layout\_clusters.md, plus JSON outputs. Source of the populated pptx\_utils. (2) 8 ATU decks across 7 clients (Apr 16) — produced atu\_analysis.md with ATU-vs-PET structural comparison. Source of the atu-deck project skill and atu-insight-writer.
* **Agentic MR PRD** — broader client-facing platform vision. SlideGen is an internal capability toward that vision; this PRD scopes it to internal use for Q3.

### 11.2 Glossary

* **Skill:** A SKILL.md file with domain knowledge and orchestration guidance. Invoked by Claude Code.
* **Tool:** A programmatic primitive (Python function, CLI command, API) callable by skills.
* **Workflow skill:** Top-level orchestration skill mapping user intent to a composed sequence.
* **Project-type skill:** Skill encoding methodology for one project type (PET, ATU, etc.).
* **Atomic skill:** Single-purpose skill (slide-creator, viz-selector, etc.).
* **`pptx\\\_utils`:** Python package of composition primitives (brand, layouts, shape builders, chart builders, lxml helpers). Co-equal building block with skills.
* **Generated script:** \~30-line glue code Claude writes per slide, composing `pptx\\\_utils` calls per skill instructions. Disposable.
* **Spec:** Slide plan output. Contract between intelligent and deterministic skills.
* **Spec contract:** The required parameters a renderer needs. Enforced by spec-validator.
* **BRAND{} / LAYOUTS{} / CHART\_PATTERNS{}:** Python dicts in `pptx\\\_utils` encoding per-client colors/fonts, coordinate presets, and named chart patterns respectively.

### 11.3 Key Quotes

> Sriram (Apr 14): "There are a bunch of building blocks and any specific output is a stringing together of those building blocks. In what order to string the building blocks we hand over that decision to Claude Code."

> Sriram (Apr 14): "Use all of the intelligence there itself and say these are all the parameters that you need to specify in this format so that that deterministic program can use all of those parameters to create the perfect slide."

> Sriram (Apr 14): "I'm very comfortable if we take the rest of this week to just think through this stuff... We'll have the PRD ready for the new version at the end of this week."

> Original PPT Agent Brief: "Hierarchy: Metric Tag > Question Type Default > Human-in-Loop prompt."

### 11.4 Workflow Coverage Matrix

All 8 workflows are in Q3 scope. Ship order reflects §9.2 — smallest-risk first, primary demo (refresh deck) early, analysis-heavy workflows later.

|#|Workflow|Skills Used|Ship Order|
|-|-|-|-|
|3|Edit slide|edit-slide-workflow, deck-reader, spec-validator, slide-updater, slide-editor, headline-writer, slide-creator, deck-assembler|1 (shake-down — rebuild mode first, then data\_refresh and edit modes)|
|2|Refresh deck|refresh-deck-workflow, pet-deck, deck-reader (Tier 1+2), prior-wave-context-builder, synapse-read, slide-plan-generator-refresh, slide-updater, slide-creator, trend-analyzer, headline-writer, deck-assembler|2 (primary demo)|
|1|Create deck|create-deck-workflow, pet-deck, context-builders, hypothesis-generator, insight-writer (sfea or atu per project type), slide-plan-generator-hypothesis, viz-selector, layout-selector, headline-writer, slide-creator, deck-assembler|3|
|4|Add slide|add-slide-workflow, deck-reader, synapse-read, slide-plan-generator-single, viz-selector, layout-selector, headline-writer, slide-creator, (segment-comparator + stat-sig-annotator for segment mode), deck-assembler|4|
|5|Annotate slide|annotate-slide-workflow, deck-reader, callout-writer, slide-editor, (segment-comparator + stat-sig-annotator for insight mode), slide-creator, deck-assembler|5|
|6|Restructure deck|structural-edit-workflow, deck-reader, spec-validator, layout-selector, deck-assembler|6|
|7|Audit deck|deck-audit-workflow, deck-reader, spec-validator (read-only)|7|
|8|Executive summary|executive-summary-workflow, deck-reader (full deck), insight-writer (sfea or atu per project type), slide-plan-generator-exec-summary, executive-summary-writer, slide-creator, deck-assembler|8|



