---
name: slide-plan
description: "Use when building a structured slide plan from validated analysis and a Hypothesis Bank. Trigger when: user says 'build slide plan', 'create ask document', 'plan slides', or Stage 3 (sfea-insight-writer) outputs have been confirmed and Stage 4 is ready. Reads validated_analysis.md, slide_headlines.md, exec_summary.md, hypothesis_bank.md, KBQs.md, and survey_context.md. Always opens with Cover (Slide 1), ES (Slide 2), Recs (Slide 3), then data slides. Per-slide insights are drawn from validated_analysis.md and matched to slide_headlines.md — not written from scratch."
---

# Slide Plan Builder

You are a senior market research analyst. Your job is to read the validated analysis and hypothesis bank, then build a structured slide plan for a PET SFEA report. Each data slide answers one analytical question. Insights per slide are sourced from the validated analysis — not written from scratch. The deck always opens with Cover → ES → Recs before any data slides.

---

## STEP 1 — Resolve file paths

Ask the user for the **project folder** and **wave name** only. Derive all paths:

| File | Path | Required? |
|------|------|-----------|
| **Validated Analysis** | `{project}/context/{wave}/validated_analysis.md` | Required |
| **Slide Headlines** | `{project}/context/{wave}/slide_headlines.md` | Required |
| **Executive Summary** | `{project}/context/{wave}/exec_summary.md` | Required |
| **Hypothesis Bank** | `{project}/context/{wave}/hypothesis_bank.md` | Required |
| **KBQs** | `{project}/input/Wave/{wave}/KBQs.md` | Required |
| **Survey Context** | `{project}/context/{wave}/survey_context.md` | Strongly recommended |

Report status before reading:
```
Input files:
  ✓ validated_analysis.md   (data-validated findings per hypothesis)
  ✓ slide_headlines.md      (confirmed talking headlines per domain)
  ✓ exec_summary.md         (ES + Recs — content for Slides 2–3)
  ✓ hypothesis_bank.md      (question codes, segment cuts, flags)
  ✓ KBQs.md                 (domain structure for narrative sequencing)
  ✓ survey_context.md       (question text for chart descriptions)
```

If any required file is missing, stop and tell the user which is absent.

---

## STEP 2 — Read all files in full

Read every available file completely before doing anything else.

- **Validated Analysis** — Phase 0 output from sfea-insight-writer: per-hypothesis validation status (CONFIRMED / PARTIALLY CONFIRMED / NOT CONFIRMED / INSUFFICIENT DATA), extracted prior/current/delta values, segment data, methodology flags
- **Slide Headlines** — Phase 1 output: one confirmed talking headline per domain; these are the slide-level insights to carry into the plan — do not rewrite them
- **Executive Summary** — Phase 2–3 output: the ES and Recommendations content that populates Slides 2 and 3
- **Hypothesis Bank** — all hypotheses with "Test with:" question codes, segment cuts, METHODOLOGY ARTIFACT and [ACTION ITEM] flags
- **KBQs** — domain structure; use to sequence data slides into narrative sections
- **Survey Context** — question codes + question text; use to write precise chart descriptions

---

## STEP 3 — Tag each hypothesis

Before clustering, tag every hypothesis in the bank with:

| Tag | What to extract |
|-----|----------------|
| **Primary Q** | The main question code(s) driving the measurement — from "Test with:" line only |
| **Story theme** | A 3–5 word label for what this hypothesis is about (e.g., "message recall by setting", "close rate vs TAG") |
| **Segment cuts** | Which splits are needed — read ONLY from the "Test with:" line, not from rationale prose |
| **Validation status** | CONFIRMED / PARTIALLY CONFIRMED / NOT CONFIRMED / INSUFFICIENT DATA — from validated_analysis.md |
| **Flags** | METHODOLOGY ARTIFACT and/or [ACTION ITEM] if present |

**Critical tagging rule — rationale is not instruction:** The rationale paragraph explains *why* the hypothesis is directional. Segment cuts, question codes, and chart design come exclusively from the **"Test with:"** line. If Practice Setting, HII, or any other segment is mentioned only in the rationale prose and not in the "Test with:" line, do NOT add it as a cut on the slide.

Do not display the full tag table — use it internally to drive clustering.

---

## STEP 4 — Fix the first three slides

Before clustering any data slides, lock in the opening:

| Slide | Type | Content source |
|-------|------|---------------|
| **Slide 1** | `cover` | Brand name, wave label, study type — from project_context.md |
| **Slide 2** | `executive_summary` | ES content from exec_summary.md |
| **Slide 3** | `executive_summary` | Recommendations from exec_summary.md |

These are fixed. Data slides begin at Slide 4.

---

## STEP 5 — Cluster into data slides

**Grouping rule:** Hypotheses that share the same primary question code AND the same story theme belong on the same slide.

For each cluster:

- Assign a **slide title** — a short, declarative topic label (not a question, not a finding — e.g., "Message Recall by Practice Setting")
- Identify the **driving question** — the single analytical question this slide must answer
- List the **hypotheses tested** (e.g., H6, H11)
- Pull the **insight** — find the matching headline from slide_headlines.md for this domain/topic; copy it verbatim. Do not rewrite it.
- Note the **validation status** — the dominant status across hypotheses on this slide (e.g., "CONFIRMED", "PARTIALLY CONFIRMED — see note")
- Assign the **slide_type** — use the exact value from the Chart Type → slide_type Mapping table below
- List **question codes with text** — from Survey Context
- List **segment cuts** needed
- Write a **narrative arc** — one sentence: what this slide shows given the actual validated data (not the prediction)
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

## STEP 6 — Sequence data slides

Data slides begin at Slide 4. Order into narrative flow using KBQ domain sequence as the backbone:

1. Sales Activity
2. Messaging — Recall & Delivery
3. Messaging — Effectiveness
4. Interactions — Setup & Tactics
5. Rep Performance
6. Prescription Intent / LTIP
7. Branded Close / CTA
8. Drivers of High Impact

Within each section, order slides from most structural (overall comparison, QoQ) to most diagnostic (segment splits, driver analysis).

Place methodology artifact slides adjacent to the substantive slide they affect — not in an appendix. Flag them clearly.

---

## OUTPUT FORMAT

```
# Slide Plan — [Wave]
**Generated from:** validated_analysis.md + slide_headlines.md + exec_summary.md + hypothesis_bank.md + KBQs.md + survey_context.md
**Date:** [today's date]
**Total slides:** [N] (Cover + ES + Recs + [N-3] data slides)

---

## SECTION: Opening

---

### Slide 1 — Cover
**slide_type:** `cover`
**Content:** [Brand name], [Wave label], [Study type]

---

### Slide 2 — Executive Summary
**slide_type:** `executive_summary`
**Content:** From exec_summary.md — ES section (Format [A/B/C…])

---

### Slide 3 — Recommendations
**slide_type:** `executive_summary`
**Content:** From exec_summary.md — Recommendations section

---

## SECTION: [Domain Name]

---

### Slide [N] — [Slide Title]
**Driving question:** [The single analytical question this slide must answer]
**Hypotheses tested:** H[x], H[y], H[z]
**Validation status:** CONFIRMED / PARTIALLY CONFIRMED / NOT CONFIRMED
**Insight:** [Verbatim from slide_headlines.md for this domain/topic]
**slide_type:** [exact value, e.g., `single_bar_with_delta`]
**Chart description:** [Specific layout: sorted by what, delta column, segment splits, etc.]
**Primary questions:**
- [Q code] — "[Question text from Survey Context]"
- [Q code] — "[Question text from Survey Context]"
**Cuts needed:** [QoQ / Practice Setting / HII vs Others / etc.]
**Narrative arc:** [One sentence: what this slide shows given the actual validated data]
METHODOLOGY ARTIFACT — [brief description] [only if applicable]
[ACTION ITEM] — [brief description] [only if applicable]
```

Slide numbers do not restart per section.

---

## OUTPUT FILE

Write the completed slide plan to `{project}/context/{wave}/slide_plan.md`. Ask the user to confirm the path before writing.

---

## RULES

1. **Slides 1–3 are always Cover, ES, Recs.** Data slides begin at Slide 4. No exceptions.
2. **Insights come from slide_headlines.md — copy verbatim.** Do not rewrite, paraphrase, or improve them. They were confirmed in Phase 1.
3. **Narrative arc reflects actual data, not predictions.** Use the validation status and validated values. A NOT CONFIRMED finding still gets a narrative arc — it describes what the data actually showed.
4. **One driving question per slide.** If a slide is trying to answer two unrelated questions, split it.
5. **Every hypothesis must appear on at least one slide.** If a hypothesis has no cluster match, create a standalone slide for it.
6. **Methodology artifact slides stay adjacent** to the substantive slide they affect — not hidden in appendix.
7. **Slide titles are topic labels, not findings.** Findings go in the insight and narrative arc.
8. **slide_type must be a valid renderer key.** Use the exact value from the mapping table. Add specifics in the Chart description field.
9. **Action items surface visibly.** Any slide derived from a client action item gets the [ACTION ITEM] flag.
10. **No question code appears on two slides serving the same analytical purpose.** Merge if same question + same cut + same analytical role.
11. **No context stored in this skill.** Everything comes from reading the files at runtime.
12. **Rationale prose does not define chart design.** Segment cuts and question codes come exclusively from the "Test with:" line.
13. **Show all dimensions of a multi-part question.** If a question has sub-dimensions (e.g., ME Q2.20 has A=Differentiation, B=Believability, C=Motivation), include all in the spec unless the hypothesis bank explicitly restricts to a specific sub-dimension.
14. **Use exact question text.** In Primary questions, always write the full question text from Survey Context alongside the code — never shorthand or paraphrase.
