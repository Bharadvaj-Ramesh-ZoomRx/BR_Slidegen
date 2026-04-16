---
name: slide-plan
effort: high
description: "Use when building a structured slide plan from validated analysis and a Hypothesis Bank. Trigger when: user says 'build slide plan', 'create ask document', 'plan slides', or Stage 3 (sfea-insight-writer) outputs have been confirmed and Stage 4 is ready. Reads validated_analysis.md, slide_headlines.md, exec_summary.md, hypothesis_bank.md, KBQs.md, and survey_context.md. Always opens with Cover (Slide 1), ES (Slide 2), Recs (Slide 3), then data slides. Per-slide insights are drawn from validated_analysis.md and matched to slide_headlines.md — not written from scratch."
---

## Auto-Detected Context
!`python3 -c "
import glob, json, os
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
    print(f'**Active project:** \`{proj}\`')
    print(f'**Active wave:** \`{wave}\`')
    ctx = f'{proj}/context/{wave}' if wave else f'{proj}/context'
    for name, req in [('validated_analysis.md','Required'), ('slide_headlines.md','Required'), ('exec_summary.md','Required'), ('hypothesis_bank.md','Required'), ('survey_context.md','Recommended')]:
        path = f'{ctx}/{name}'
        if os.path.exists(path):
            print(f'  ✓ {name} ({os.path.getsize(path)//1024}KB) [{req}]')
        else:
            print(f'  ✗ {name} — MISSING [{req}]')
    for case in ['Wave', 'wave']:
        kbq = f'{proj}/input/{case}/{wave}/KBQs.md'
        if os.path.exists(kbq):
            print(f'  ✓ KBQs.md ({os.path.getsize(kbq)//1024}KB) [Required]')
            break
    else:
        print(f'  ✗ KBQs.md — MISSING [Required]')
else:
    print('**No active project detected** — user must specify project folder')
"
`

# Slide Plan Builder

You are a senior market research analyst. Your job is to read the narrative threads and validated analysis, then build a structured slide plan for a PET SFEA report. Each data slide answers one analytical question. Headlines per slide come from narrative_threads.md — not written from scratch. The deck always opens with Cover → ES → Recs before any data slides.

---

## STEP 1 — Confirm project and resolve file paths

Use the auto-detected project, wave, and file status shown above. If correct, proceed. If not, ask the user to specify.

The required files are:

| File | Path | Required? |
|------|------|-----------|
| **Narrative Threads** | `{project}/context/{wave}/narrative_threads.md` | Required |
| **Validated Analysis** | `{project}/context/{wave}/validated_analysis.md` | Required |
| **Hypothesis Bank** | `{project}/context/{wave}/hypothesis_bank.md` | Required |
| **KBQs** | `{project}/input/Wave/{wave}/KBQs.md` | Required |
| **Survey Context** | `{project}/context/{wave}/survey_context.md` | Strongly recommended |

If any required file is missing (check auto-detected status above), stop and tell the user which is absent.

---

## STEP 2 — Read all files in full

Read every available file completely before doing anything else.

- **Narrative Threads** — Stage 3 Phase 1 output: story arcs (with pattern type and urgency), per-slide headlines (with hypothesis mappings), executive summary (arc-organized), and recommendations. **This is the editorial brief — headlines are copied verbatim, not rewritten.**
- **Validated Analysis** — Phase 0 output: per-hypothesis validation status, extracted data values, segment data, methodology flags
- **Hypothesis Bank** — all hypotheses with "Test with:" question codes, segment cuts, METHODOLOGY ARTIFACT and [ACTION ITEM] flags
- **KBQs** — domain structure; used for section grouping within the narrative sequence
- **Survey Context** — question codes + question text; used to write precise chart descriptions

---

## STEP 3 — Tag each hypothesis

Before clustering, tag every hypothesis in the bank with:

| Tag | What to extract |
|-----|----------------|
| **Primary Q** | The main question code(s) driving the measurement — from "Test with:" line only |
| **Story theme** | A 3-5 word label for what this hypothesis is about |
| **Segment cuts** | Which splits are needed — read ONLY from the "Test with:" line, not from rationale prose |
| **Validation status** | CONFIRMED / PARTIALLY CONFIRMED / NOT CONFIRMED / INSUFFICIENT DATA — from validated_analysis.md |
| **Arc** | Which narrative thread this hypothesis belongs to — from narrative_threads.md |
| **Flags** | METHODOLOGY ARTIFACT and/or [ACTION ITEM] if present |

**Critical tagging rule — rationale is not instruction:** The rationale paragraph explains *why* the hypothesis is directional. Segment cuts, question codes, and chart design come exclusively from the **"Test with:"** line. If Practice Setting, HII, or any other segment is mentioned only in the rationale prose and not in the "Test with:" line, do NOT add it as a cut on the slide.

Do not display the full tag table — use it internally to drive clustering.

---

## STEP 4 — Fix the first three slides

Before clustering any data slides, lock in the opening:

| Slide | Type | Content source |
|-------|------|---------------|
| **Slide 1** | `cover` | Brand name, wave label, study type — from project_context.md |
| **Slide 2** | `executive_summary` | ES content from narrative_threads.md — Executive Summary section |
| **Slide 3** | `executive_summary` | Recommendations from narrative_threads.md — Recommendations section |

These are fixed. Data slides begin at Slide 4.

---

## STEP 5 — Cluster into data slides (arc-informed)

**Primary grouping rule:** Hypotheses that share the same primary question code AND the same story theme belong on the same slide.

**Arc-informed composition:** When clustering, consult narrative_threads.md to understand what story each hypothesis serves. This may influence composition:

- **Hypotheses in the same arc that share data** should be on the same slide (standard clustering)
- **Hypotheses in different arcs that share a question code** — consider whether the slide serves one arc primarily, or whether a competitive comparison slide better serves the narrative (e.g., a side-by-side RYB vs TAG recall slide may serve the "AZ gaining ground" arc better than separate brand slides)
- **Each slide should serve primarily one arc.** If a slide contains hypotheses from two arcs, it will have a confused headline. Split it or assign it to the dominant arc.

For each cluster:

- Assign a **slide title** — a short, declarative topic label (not a question, not a finding)
- Identify the **driving question** — the single analytical question this slide must answer
- List the **hypotheses tested** (e.g., H6, H11)
- **Arc** — which narrative thread this slide serves, from narrative_threads.md
- **Role in arc** — one sentence: what this slide contributes to the arc's argument (e.g., "Shows the messaging dimension of the competitive convergence")
- Pull the **headline** — find the matching headline from narrative_threads.md; copy it verbatim. Do not rewrite it.
- Note the **validation status** — the dominant status across hypotheses on this slide
- Assign the **slide_type** — use the exact value from the Chart Type → slide_type Mapping table below
- Write a **chart description** — specific layout details: sorted by what, delta column, segment splits, etc.
- List **question codes with text** — from Survey Context
- List **segment cuts** needed
- Flag any **methodology artifacts** or **action items**

If a hypothesis cluster is too large for one slide, split it. If two clusters tell the same story with the same chart, merge them.

### Chart Type → `slide_type` Mapping

| `slide_type` | When to use | Data pattern |
|---|---|---|
| `cover` | Title slide | N/A |
| `executive_summary` | Bullet-list insights / recommendations | Text only |
| `single_bar_with_delta` | One brand, one metric, ranked list with QoQ delta | Single question code, prior + current values per row |
| `dual_bar_with_delta` | Two metrics side-by-side (e.g., MR left + ME right) with deltas | Two question codes for the same brand, same category list |
| `dual_bar_qoq` | Two side-by-side Q4-vs-Q3 clustered bars (e.g., MR for two brands) | Same question across two brands, each with prior + current |
| `clustered_compare` | Two groups compared on same metric (brand vs brand, segment vs segment) with gap/delta | One question code, two brands or segments as separate series |
| `dual_bar_compare` | Two separate brand bar charts side-by-side with shared category column | Same question, two brands, need visual separation (not overlaid) |
| `qoq_bar_with_delta` | Single chart showing Q4 vs Q3 clustered bars + delta column | One question, one brand, prior + current as clustered pair |
| `two_section_bar` | Two vertically stacked bar sections on one slide (e.g., RYB top + TAG bottom) | Two brands or segments shown in separate chart areas, same metric |
| `stacked_order` | Stacked bar with ordinal breakdown (1st/2nd/3rd recall) + total column | One question with ordinal sub-rows (nested recall order) |
| `abacus` | XY scatter abacus with prior/current dots, value columns, QoQ delta | Attribute ratings with prior + current, need precise comparison |
| `dual_abacus` | Two side-by-side abacus panels (e.g. Acad vs Comm by brand) | Same attributes, two brands, each with two segments |
| `followup_rep` | Follow-up rep abacus with dual brand dots + per-brand QoQ delta columns | Two brands on same axis with separate delta columns |
| `hii_scorecard` | Multi-section clustered column chart with section headers + callout boxes | Multiple metrics grouped into sections (e.g. HII drivers) |
| `dual_doughnut` | Side-by-side doughnut pairs comparing patient segments by brand | Two segments, each with two brand doughnuts showing QoQ rings |
| `message_mbd` | Multi-column abacus for MBD (Motivation, Believability, Differentiation) | ME question with sub-dimensions (M/B/D), plus composite effectiveness |
| `trended_scorecard` | Multi-panel mini line chart grid showing QoQ trends | Multiple metrics, each with 3+ time periods as series |
| `trended_activity` | Side-by-side line + stacked column panels (reach/SOV/frequency) | Activity metrics with time trends + volume breakdown |
| `quadrant_scatter` | 2×2 quadrant scatter chart (stated vs derived importance) | Two-dimensional attribute scores (x=stated, y=derived) |
| `heatmap_table` | Heatmap table with green gradient fills + QoQ delta columns | Matrix data: rows × columns with values + deltas per cell |
| `qual_theme_analysis` | Theme frequency bars + representative quote boxes | Verbatim responses from qualitative_data.json, theme-coded in validated_analysis.md |

**Selection decision tree:**
```
Single brand, single metric, ranked list?
  → single_bar_with_delta ← SAFEST DEFAULT for {desc, prior, current} data

Same metric, two brands or segments compared?
  → clustered_compare — ONLY if data has primary_current + comp_current fields
  → dual_bar_compare — ONLY if two separate data keys merged
  → two_section_bar — ONLY if extra.top + extra.bottom configured
  ⚠ If data only has {desc, prior, current}: use single_bar_with_delta instead

Two related metrics for same brand (e.g., MR + ME)?
  → dual_bar_with_delta — ONLY if data has {prefix}_current fields per side
  ⚠ If data only has {desc, prior, current}: use single_bar_with_delta instead

QoQ comparison (current vs prior period)?
  → qoq_bar_with_delta (single brand)
  → dual_bar_qoq (two brands, side-by-side)

Recall order breakdown (1st/2nd/3rd)?
  → stacked_order — ONLY with nested_ordinal extraction
  ⚠ Regular row_range data won't stack correctly

Attribute ratings / rep performance?
  → abacus (with value columns and delta) ← works with {desc, prior, current}

Message effectiveness with M/B/D sub-dimensions?
  → message_mbd

Scorecard / multi-section metric table?
  → hii_scorecard — ONLY if extra.sections configured
  → heatmap_table — ONLY if extra.columns configured
  ⚠ Without extra config: use executive_summary with insights list instead

Text-only insights / qualitative findings?
  → executive_summary — populate ask.extra.insights as list of strings

Qualitative theme analysis (from QUALITATIVE hypotheses)?
  → qual_theme_analysis — native PPT bar chart + label TABLE + 4 grouped quote boxes
  Config fields: ask.extra.themes (must sum to 100%), ask.extra.quotes (4 quotes, ≥15 words each),
    ask.extra.qual_subtitle, ask.extra.qual_sample, ask.extra.qual_source
  ⚠ Only for dedicated qual slides (HQ hypotheses). For callouts on quant slides, use qual_callout config instead.
```

**⚠ CRITICAL — Data compatibility check before assigning slide_type:**
Most extractions produce simple `{desc, prior, current}` rows. Only these renderers work with that format:
- `single_bar_with_delta` ✓
- `abacus` ✓
- `executive_summary` ✓ (text only, needs `extra.insights`)

All other renderers need specific field prefixes, extra config, or multi-key data merges. Do NOT assign them unless the extraction is specifically configured to produce matching data.

---

## STEP 5.5 — Deduplicate across slides

Before sequencing, audit the full cluster list for question-level redundancy. Apply the following tests:

**Test 1 — Same question, same cut, different section?**
If the same question code appears with the same segment cut on two slides in different sections, merge them into one slide placed where it is most analytically central.

**Test 2 — Same question, different cut, same section?**
If the same question code appears on two slides in the same section with different cuts (e.g., overall on one slide, by Practice Setting on another), merge into one slide with an overall panel and a segment-split panel — unless the driving questions are genuinely distinct and the segment finding warrants its own narrative.

**Test 3 — Same outcome variable across multiple driver slides?**
If the same outcome question (e.g., LTIP, derived HII) is the dependent variable on three or more slides in the same section, consolidate into one driver-analysis slide — unless each predictor is a distinct client action item requiring standalone presentation.

**After applying these tests:**
- List any merges made and which hypotheses were consolidated
- Confirm every hypothesis still appears on at least one slide after merging
- Proceed to sequencing only once deduplication is complete

---

## STEP 6 — Sequence data slides (arc-informed)

Data slides begin at Slide 4. Sequencing uses **two organizing principles**:

### Primary: Arc urgency drives section order

The lead arc (highest urgency) gets its slides presented first. Within the deck, arrange by:

1. **ACT NOW arcs** — slides for these arcs come first (after opening)
2. **MONITOR arcs** — middle of the deck
3. **CELEBRATE / CLOSURE arcs** — later in the deck (positive closure validates prior work)

### Secondary: Domain grouping within arcs

Within each arc's slides, group by KBQ domain to maintain analytical coherence:

1. Sales Activity
2. Messaging — Recall & Delivery
3. Messaging — Effectiveness
4. Interactions — Setup & Tactics
5. Rep Performance
6. Prescription Intent / LTIP
7. Branded Close / CTA
8. Drivers of High Impact
9. NPP / Omnichannel

### Handling cross-arc slides

Some arcs share domain slides (e.g., a LTIP slide may serve both a CONVERGENCE arc and a DIVERGENCE arc). Place the slide under its **primary arc** — the one listed in its Arc field.

### Orphan slides

Slides for orphan findings (from narrative_threads.md → Orphan Findings section) are placed at the end of the most relevant domain section.

Within each section, order slides from most structural (overall comparison, QoQ) to most diagnostic (segment splits, driver analysis).

Place methodology artifact slides adjacent to the substantive slide they affect — not in an appendix.

---

## OUTPUT FORMAT

```
# Slide Plan — [Wave]
**Generated from:** narrative_threads.md + validated_analysis.md + hypothesis_bank.md + KBQs.md + survey_context.md
**Date:** [today's date]
**Total slides:** [N] (Cover + ES + Recs + [N-3] data slides)
**Narrative arcs:** [list arc titles with urgency]

---

## METHODOLOGY FOOTNOTES (apply to adjacent slides)
- **MF[N] — [title]:** [description] *Adjacent to: Slides [X, Y, Z].*

---

## Slide 1 — Cover
**slide_type:** `cover`
**Content:** [Brand name], [Wave label], [Study type]

---

## Slide 2 — Executive Summary
**slide_type:** `executive_summary`
**Content:** From narrative_threads.md — Executive Summary section (Format [A/B/C…])
**Note:** ES is organized by story arcs, not by domains.

---

## Slide 3 — Recommendations
**slide_type:** `executive_summary`
**Content:** From narrative_threads.md — Recommendations section

---

## SECTION: [Arc Title] ([Pattern] — [Urgency])

---

### Slide [N] — [Slide Title]
**Arc:** [Thread N — Arc title]
**Role in arc:** [One sentence: what this slide contributes to the arc's argument]
**Driving question:** [The single analytical question this slide must answer]
**Hypotheses tested:** H[x], H[y], H[z]
**Validation status:** CONFIRMED / PARTIALLY CONFIRMED / NOT CONFIRMED
**Headline:** [Verbatim from narrative_threads.md — do not rewrite]
**slide_type:** [exact value]
**Chart description:** [Specific layout: sorted by what, delta column, segment splits, etc.]
**Primary questions:**
- [Q code] — "[Question text from Survey Context]"
- [Q code] — "[Question text from Survey Context]"
**Cuts needed:** [QoQ / Practice Setting / HII vs Others / etc.]
**qual_callout:** [verbatim Q code] — "[quote]" ([theme], [%]) ← only if narrative_threads.md specifies a qual_callout for this slide
METHODOLOGY ARTIFACT — [brief description] [only if applicable]
[ACTION ITEM] — [brief description] [only if applicable]

### Slide [N] — [Qualitative Slide Title]  ← for QUALITATIVE hypothesis slides only
**Arc:** [Thread N — Arc title]
**Role in arc:** [One sentence: what this slide contributes to the arc's argument]
**Driving question:** [The single analytical question this slide must answer]
**Hypotheses tested:** HQ[x]
**Validation status:** CONFIRMED / PARTIALLY CONFIRMED / NOT CONFIRMED
**Headline:** [Verbatim from narrative_threads.md — do not rewrite]
**slide_type:** `qual_theme_analysis`
**Chart description:** Left panel: native PPT bar chart (BAR_CLUSTERED) with theme label TABLE + sub-title + "% of respondents" x-axis. Right panel: 4 grouped quote boxes with attribution above line separator.
**Qual source:** [verbatim Q code] (n=[count], from [sheet_key])
**Qual subtitle:** [Short question context, e.g., "Prescribing Conversion – Rationale (Unaided)"]
**Qual sample:** [e.g., "n = 47" or "n = 91 of 189" if filtered to actionable responses]
**Themes:** [ALL themes from validated_analysis.md — MUST sum to 100%. Include "Others / non-specific" bucket for uncoded/small themes]
**Quotes:** [4 representative quotes — each ≥15 words, explanatory (not labels), diverse themes, competitive reasoning preferred]
```

Slide numbers do not restart per section.

---

## OUTPUT FILE

Write the completed slide plan to `{project}/context/{wave}/slide_plan.md`. Ask the user to confirm the path before writing.

---

## RULES

1. **Slides 1-3 are always Cover, ES, Recs.** Data slides begin at Slide 4. No exceptions.
2. **Headlines come from narrative_threads.md — copy verbatim.** Do not rewrite, paraphrase, or improve them. They were confirmed in Stage 3.
3. **ES and Recs come from narrative_threads.md.** Copy the Executive Summary and Recommendations sections verbatim for Slides 2-3.
4. **Each slide serves primarily one arc.** The Arc and Role in arc fields are required for every data slide. If a slide serves two arcs, assign the dominant one.
5. **One driving question per slide.** If a slide tries to answer two unrelated questions, split it.
6. **Every hypothesis must appear on at least one slide.** If a hypothesis has no cluster match, create a standalone slide.
7. **Methodology artifact slides stay adjacent** to the substantive slide they affect — not hidden in appendix.
8. **Slide titles are topic labels, not findings.** Findings go in the headline.
9. **slide_type must be a valid renderer key.** Use the exact value from the mapping table. Add specifics in the Chart description field.
10. **Action items surface visibly.** Any slide from a client action item gets the [ACTION ITEM] flag.
11. **No question code on two slides serving the same analytical purpose.** Merge if same question + same cut + same role.
12. **No context stored in this skill.** Everything comes from reading the files at runtime.
13. **Rationale prose does not define chart design.** Segment cuts and question codes come exclusively from the "Test with:" line.
14. **Show all dimensions of a multi-part question.** If a question has sub-dimensions (e.g., ME Q2.20 has A=Differentiation, B=Believability, C=Motivation), include all unless the hypothesis bank explicitly restricts.
15. **Use exact question text.** In Primary questions, always write the full question text from Survey Context alongside the code — never shorthand or paraphrase.
16. **Arc-informed sequencing.** Sections are ordered by arc urgency (ACT NOW → MONITOR → CELEBRATE), not by domain convention. Domain grouping applies within arcs.
17. **Slide plan and config.yaml must stay in sync.** If Stage 5 downgrades a slide_type because data format does not match (e.g. clustered_compare to single_bar_with_delta), update slide_plan.md to reflect the final type. Discrepancies between plan and config cause confusion during review.
