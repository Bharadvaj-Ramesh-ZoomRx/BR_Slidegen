---
name: pet-deck
effort: high
paths: []
description: "Project-type skill for Promotional Effectiveness Tracking (PET) decks. Invoked by create-deck-workflow and refresh-deck-workflow when the project type is PET. Encodes PET-specific methodology defaults, universal slide structure, KBQ framing, and wave-over-wave comparison logic for any pharma client and brand."
---

## Auto-Detected Context
!`python3 -c "
import glob, os
try:
    import yaml
except ImportError:
    yaml = None
configs = sorted(glob.glob('projects/*/config.yaml'), key=os.path.getmtime, reverse=True) if yaml else []
if configs:
    try:
        with open(configs[0]) as f:
            cfg = yaml.safe_load(f)
    except Exception:
        cfg = {}
    if not isinstance(cfg, dict): cfg = {}
    proj = os.path.dirname(configs[0])
    wave = cfg.get('project',{}).get('wave','')
    brand = cfg.get('project',{}).get('brand','')
    client = cfg.get('project',{}).get('client','')
    ta = cfg.get('project',{}).get('therapy_area','')
    print(f'**Active project:** \`{proj}\`')
    print(f'**Brand:** \`{brand}\` | **Client:** \`{client}\` | **TA:** \`{ta}\`')
    print(f'**Active wave:** \`{wave}\`')
    ctx = f'{proj}/context/{wave}' if wave else f'{proj}/context'
    for md in sorted(glob.glob(f'{ctx}/*.md')):
        print(f'  - \`{os.path.basename(md)}\` ({os.path.getsize(md)//1024}KB)')
else:
    print('**No active project detected** — user must specify project folder')
"
`

# PET Deck — Project-Type Skill

You are a project-type skill for **Promotional Effectiveness Tracking (PET)** studies. Workflow skills (`create-deck-workflow`, `refresh-deck-workflow`) invoke you to apply PET-specific methodology defaults before delegating to lower-level planning, analysis, and creation skills.

This skill is **client-agnostic and brand-agnostic**. It works for any pharmaceutical brand that has a PET study — the brand, client, therapy area, competitor set, and visual identity are all resolved at runtime from `config.yaml` and `BRAND{}`.

---

## Cardinal Rules

1. **Never hardcode a brand, client, or therapy area.** All identity comes from `config.yaml` via `project.brand`, `project.client`, `project.therapy_area`. Visual identity (colors, fonts, heading color, template) comes from `BRAND{}` / `CLIENT{}` in `slidegen/pptx_utils/brand.py`.
2. **Wave-over-wave comparison is the core PET value proposition.** Every data slide must show current vs. prior wave with deltas. Trends across multiple waves (trended scorecards) are expected for studies with 3+ waves.
3. **Competitive positioning is always present.** PET studies track the focal brand against named competitors. All message, rep performance, and intent slides must include competitive context.
4. **The slide plan drives the slide list — not this skill.** This skill provides PET methodology defaults and section structure expectations. The actual slide sequence is determined by `slide-plan-generator-hypothesis` (create flow) or `slide-plan-generator-refresh` (refresh flow).
5. **Respect the brand's visual identity.** Use `get_brand(config.project.brand)` for colors, `CLIENT[brand.client]` for fonts and heading color. Never apply one client's typography or palette to another.

---

## Inputs

| Input | Source | Required |
|---|---|---|
| `config.yaml` | `projects/{name}/config.yaml` | Yes |
| `project.brand` | config — `project.brand` (key into `BRAND{}`) | Yes |
| `project.client` | config — `project.client` (key into `CLIENT{}`) | Yes |
| `project.wave` | config — `project.wave` (e.g., "Q1 2026") | Yes |
| `project.therapy_area` | config — `project.therapy_area` | Yes |
| Slide plan | `context/{wave}/slide_plan.md` | Yes (generated upstream) |
| Narrative threads | `context/{wave}/narrative_threads.md` | Yes (generated upstream) |
| `BRAND{}` / `CLIENT{}` | `slidegen/pptx_utils/brand.py` | Yes (runtime lookup) |

---

## What a PET Deck Universally Contains

Based on analysis of **276 PET decks** across the full client portfolio (905-deck corpus, 13,949 slides, 21,918 charts). Every PET deck follows a consistent section structure. The specific slides within each section vary by study scope, but the sections themselves are universal.

**PET chart pattern distribution (276 decks):** `bar_clustered_horizontal` 31.0% | `line_markers_trended` 16.5% | `xy_scatter_abacus` 12.5% | `column_stacked_100_vertical` 10.8% | `bar_stacked_horizontal` 7.3% | `bar_stacked_100_horizontal` 7.2% | `column_clustered_vertical` 6.0% | `column_stacked_vertical` 3.6% | `doughnut_default` 3.6%

### Universal Sections (present in all PET decks)

| Section | Content | Typical Chart Pattern |
|---|---|---|
| **Cover** | Title slide — brand, wave, audience, date | `cover` |
| **Executive Summary** | Key findings synthesized across all modules | `executive_summary` |
| **Recommendations** | Action-oriented bullets tied to findings | `executive_summary` |
| **Message Recall** | Aided/unaided recall rates by message, current vs. prior wave | `bar_clustered_horizontal` / `single_bar_with_delta` |
| **Message Effectiveness (M-B-D)** | Motivation, Believability, Differentiation breakdown per message | `xy_scatter_abacus` / `message_mbd` |
| **Rep Performance / Call Quality** | Overall call quality, attribute ratings vs. competitors | `abacus` / `clustered_compare` |
| **Prescribing Intent (LTIP)** | Likelihood to increase prescribing, by segment/brand | `column_stacked_100_vertical` / `stacked_order` |

### Common Sections (present in most PET decks)

| Section | Content | Typical Chart Pattern |
|---|---|---|
| **Trended Scorecards** | Multi-wave trend panels for key KPIs (3+ wave studies) | `line_markers_trended` / `trended_scorecard` |
| **Activity / Share of Voice** | Reach, frequency, in-person %, visual aid usage | `trended_activity` / `dual_bar_qoq` |
| **Call-to-Action / Branded Close** | Closing behavior, follow-up actions | `bar_stacked_100_horizontal` / `stacked_order` |
| **HII Scorecard** | High Impact Interaction breakdown by segment | `column_clustered_vertical` / `hii_scorecard` |
| **Segment Comparisons** | Academic vs. Community, High Impact vs. Others | `dual_abacus` / `dual_bar_compare` |

### Optional Sections (study-dependent)

| Section | When Present |
|---|---|
| **Non-Personal Promotion (NPP)** | When digital/NPP channels are tracked |
| **KAM / MSL / FRM Performance** | When non-rep roles are in scope |
| **Qualitative Themes** | When open-ended responses are captured |
| **Quadrant Analysis** | When stated vs. derived importance is measured |
| **Heatmap Tables** | When multi-attribute cross-brand comparisons are needed |

---

## PET Methodology Knowledge

### What is PET?

**Promotional Effectiveness Tracking** is a recurring (wave-based) market research study that measures how effectively a pharmaceutical brand's field force communicates key messages to healthcare professionals (HCPs). PET studies are conducted quarterly or semi-annually for brands with active promotional campaigns.

### Standard PET Modules

1. **Activity & Share of Voice** — reach, frequency, in-person vs. virtual, visual aid usage, proactive vs. reactive engagement
2. **Messaging** — message recall (aided/unaided), message effectiveness (M-B-D), optimal message combinations, under-recalled opportunities
3. **Rep Performance** — overall call quality (1-7 scale), attribute-level ratings, closing rates, follow-up actions
4. **Impact & Outcomes** — Likelihood to Increase Prescribing (LTIP), product perception shift, High Impact Interactions (HII), drivers of high-impact calls
5. **Trends** — wave-over-wave scorecards showing trajectory of key metrics

### PET KBQ Framing

PET Key Business Questions are about:
- **Message cut-through:** Which messages are being recalled? Which are effective but under-delivered?
- **Rep quality:** How do reps compare to competitors on call quality and key attributes?
- **Competitive positioning:** Where does the brand lead, trail, or match competitors?
- **Prescription intent shift:** Is promotional effort translating to prescribing behavior change?
- **Wave-over-wave trajectory:** Are key metrics improving, stable, or declining?

### Wave-Over-Wave Comparison

This is the defining feature of PET reporting:
- Every metric is shown as current wave vs. prior wave with delta (percentage point change)
- Deltas use universal color coding: green for positive change, red for negative
- Period labels are dynamic: `{{period_current}}` vs. `{{period_prior}}` from config
- Multi-wave studies add trended scorecards showing 3+ waves as line charts
- Narrative threads frame changes as CONVERGENCE, TENSION, DIVERGENCE, or CLOSURE arcs

### Competitive Context

PET decks always position the focal brand against its competitive set:
- The primary brand's colors come from `BRAND[config.project.brand]`
- Competitor colors come from `BRAND[competitor_brand]` (if defined) or `BRAND[config.project.brand].secondary`
- Dual-brand slides (e.g., `dual_bar_compare`, `dual_abacus`) show the brand and its key competitor side by side
- Competitive language: "led competitors," "trailed [Competitor]," "on par with," "closed the gap," "extended its lead"

---

## PET Defaults Applied to Downstream Skills

When `pet-deck` is invoked by a workflow, it applies these methodology defaults before delegating:

### To `hypothesis-generator`
- Ensure hypotheses cover all 5 PET modules (Activity, Messaging, Rep Performance, Impact, Trends)
- Require PRIOR WAVE VALIDATION hypotheses for every domain finding and recommendation carried forward
- KBQ mapping must span message cut-through, rep quality, competitive positioning, and intent shift

### To `sfea-insight-writer`
- Narrative arcs should reflect wave-over-wave trajectory (improving, declining, stable, mixed)
- Executive Summary must cover: Activity, Messaging, Rep Performance, Impact — each with wave comparison
- Recommendations use PET action verbs: CONTINUE, LEVERAGE, REINFORCE, ANCHOR, ELEVATE, PRIORITIZE, ADDRESS, ENSURE, DRIVE, INCREASE, STRENGTHEN

### To `slide-plan-generator-hypothesis`
- Section ordering: Cover, Executive Summary, Recommendations, Activity/SOV, Messaging, Rep Performance, Impact/LTIP, Trends, Appendix
- Every data slide must specify prior/current wave columns and delta computation
- Trended slides are included for studies with 3+ waves of data
- Arc sequencing: ACT NOW arcs first, then MONITOR, then CELEBRATE

### To `viz-selector` and `layout-selector`
- Message Recall: `single_bar_with_delta` or `bar_clustered_horizontal`
- M-B-D: `message_mbd` (multi-column abacus)
- Rep Attributes: `abacus` or `clustered_compare`
- LTIP: `stacked_order` or `column_stacked_100_vertical`
- Trends: `trended_scorecard` (multi-panel line grid)
- Activity: `trended_activity` (line + stacked column)
- HII: `hii_scorecard` (clustered column with section headers)
- Segment comparison: `dual_abacus` or `dual_bar_compare`

### To `slide-creator` and `deck-assembler`
- Brand colors from `get_brand(config.project.brand)`
- Client fonts and heading color from `CLIENT[brand.client]`
- Template path from brand config or `config.yaml` `template_path`
- Shape naming: `zrx_{slide:03d}_{shape:03d}` convention
- Speaker notes: include question codes and full question text on every data slide

---

## Executive Summary & Recommendations

PET Executive Summaries follow a consistent structure regardless of client. See `pet-es-builder` skill for the full 8-format menu (Formats A through H). Key conventions:

### Language & Tone
- **Data-anchored:** every claim needs a metric, comparison, or directional ("increased," "declined," "led," "trailed")
- **Wave-over-wave language:** "improved vs. {{period_prior}}," "declined QoQ," "stable wave over wave," "at an all-time high"
- **Competitive language:** "led competitors," "trailed {{competitor}}," "on par with," "closed the gap"

### Recommendation Formatting
- Start with a strong **action verb in CAPS**: CONTINUE, LEVERAGE, REINFORCE, ANCHOR, ELEVATE, PRIORITIZE, ADDRESS, ENSURE, DRIVE, INCREASE, STRENGTHEN, REMIND, ENCOURAGE, UTILIZE
- Tie each recommendation to a specific finding
- Frame as opportunity, not failure: "Opportunity exists to..." not "Reps are failing to..."
- 4-6 recommendations is the sweet spot

---

## Orchestration

When invoked by `create-deck-workflow`:

```
1. Read config.yaml — resolve brand, client, therapy_area, wave, competitor set
2. Validate brand exists in BRAND{} (or project provides brand_palette override)
3. Apply PET methodology defaults to all downstream skill invocations
4. Delegate to the standard workflow stages:
     hypothesis-generator (with PET module coverage requirement)
     sfea-insight-writer (with PET narrative conventions)
     slide-plan-generator-hypothesis (with PET section ordering)
     viz-selector + layout-selector (with PET chart pattern defaults)
     slide-creator + deck-assembler (with brand visual identity)
```

When invoked by `refresh-deck-workflow`:

```
1. Read config.yaml — resolve brand, client, wave, new period labels
2. Apply PET defaults to refresh-specific skills:
     deck-reader (PET slide type recognition)
     slide-plan-generator-refresh (PET section structure preservation)
     slide-updater (wave label rewrite, headline regeneration, delta recomputation)
     deck-assembler (brand visual identity)
```

---

## Decision Rules

| Situation | Response |
|---|---|
| Brand not in `BRAND{}` | Halt — user must add brand entry or provide `brand_palette` in config |
| No competitor defined | Proceed — single-brand slides only; skip dual-brand comparisons |
| Study has only 1-2 waves | Skip trended scorecards; use QoQ bars and deltas only |
| Study has 3+ waves | Include trended scorecards for key KPIs |
| Non-rep roles in scope (KAM/MSL/FRM) | Add separate section; use Format F (Strengths & Opportunities) for ES |
| NPP/digital channels tracked | Add NPP section after Rep Performance |
| Qualitative data available | Add `qual_theme_analysis` slides; enable `qual_callout` flags on relevant data slides |
| Multi-audience study (e.g., Onc + Uro) | Use segment comparisons; consider Format H (audience columns) for ES |
| Prior wave deck available (refresh) | `deck-reader` extracts specs; preserve slide structure, regenerate headlines and data |

---

## New PET Project Onboarding

When setting up a PET project for a new brand, the project team must provide:

### Required

1. **Brand entry in `BRAND{}`** — or a `brand_palette` override in `config.yaml`. Each brand needs: `client` key, `therapy_area`, `primary_current`, `primary_prior`, `secondary` colors, and optionally `competitor_brand`. 85 CLIENT entries and 33 BRAND entries are populated from the 551-deck grounding exercise. If the brand is missing, extract colors from the client's slide master or brand guidelines.

2. **Client slide master deck** — a clean `.pptx` template with the client's layout masters, logos, and footer elements. Stored at `projects/{name}/templates/template.pptx`. If unavailable, the pipeline generates from a blank template using `BRAND{}` / `CLIENT{}` definitions.

3. **Synapse project configuration** — `project_id`, `reporting_plan_id`, and `analysis_ids` for the study. Enables Track A (JSON-first) and Track D (raw-data-first) data fetching. Alternatively, provide `source_data.xlsx` for Track B (Excel) extraction.

4. **KBQs.md** — hand-written Key Business Questions for the study. Organized by domain (Activity, Messaging, Rep Performance, Impact). This file has no auto-generation path; the research team authors it.

### Strongly Recommended

5. **Prior wave deck** — a `.pptx` from the most recent delivered wave. Enables `deck-reader` to bootstrap slide structure and `prior-wave-context-builder` to extract findings. Place in `input/wave/{wave}/`.

6. **Survey context** — survey instrument or questionnaire with question codes, message lists, and response scales. Enables `survey-context-builder` to map codes to extractions.

7. **Call notes / methodology doc** — client call notes, study design `.odt`, or methodology brief. Feeds `project-context-builder` for field intel and wave expectations.

### Optional

8. **Competitor brand entry in `BRAND{}`** — if the competitor is tracked in PET and has its own color definition. Set `competitor_brand` in the primary brand's `BRAND{}` entry for automatic color resolution.

9. **Label shortcuts** — `label_shortcuts` mapping in `config.yaml` for long message labels (e.g., mapping 80-character message texts to clean 40-character display labels). Set `use_label_shortcuts: true` in extraction params.

---

## References

- `slidegen/pptx_utils/brand.py` — `BRAND{}` (33 entries, 17 clients), `CLIENT{}`, `get_brand()`, `get_competitor()`
- `.claude/skills/workflows/create-deck-workflow/SKILL.md` — full create orchestration
- `.claude/skills/workflows/refresh-deck-workflow/SKILL.md` — full refresh orchestration
- `.claude/skills/analysis/hypothesis-generator/SKILL.md` — hypothesis generation
- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` — narrative threads + ES + recs
- `.claude/skills/planning/slide-plan-generator-hypothesis/SKILL.md` — slide plan from narrative
- `.claude/skills/planning/viz-selector/SKILL.md`, `layout-selector/SKILL.md` — chart/layout selection
- `.claude/skills/creation/slide-creator/SKILL.md`, `deck-assembler/SKILL.md` — rendering + assembly
- PRD §4.5 — project-type skills definition
