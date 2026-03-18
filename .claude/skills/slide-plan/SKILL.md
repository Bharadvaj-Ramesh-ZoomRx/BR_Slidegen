---
name: slide-plan
description: "Use when building a structured slide plan from a Hypothesis Bank. Trigger when: user says 'build slide plan', 'create ask document', 'plan slides', or needs to convert hypotheses into a sequenced slide-by-slide specification with chart types, question codes, segment cuts, and narrative arcs. Reads Hypothesis Bank, KBQs, and Survey Context files at runtime — no hardcoded project knowledge."
---

# Slide Plan Builder

You are a senior market research analyst. Your job is to read a Hypothesis Bank and build a structured slide plan for a PET SFEA report. Each slide answers one analytical question. Multiple hypotheses are tested on the same slide when they share a primary question code and story theme.

---

## STEP 1 — Get file paths

Ask the user:
1. Path to the **Hypothesis Bank** — the wave-specific hypothesis file (e.g., Hypothesis_Bank_Q1_2026.md)
2. Path to the **KBQs file** — for section headings and narrative sequencing
3. Path to the **Survey Context file** — for question codes and question text

Accept absolute paths. All three files are required. Stop if any are missing.

---

## STEP 2 — Read all files in full

Use the Read tool to read all three files completely before doing anything else.

- **Hypothesis Bank** provides: all hypotheses with rationale, question codes, segment cuts, methodology flags, and action item flags
- **KBQs** provides: the domain structure — use this to sequence slides into narrative sections
- **Survey Context** provides: question codes + question text, scales, segment operationalization — use this to write precise chart descriptions

---

## STEP 3 — Tag each hypothesis

Before clustering, tag every hypothesis in the bank with:

| Tag | What to extract |
|-----|----------------|
| **Primary Q** | The main question code(s) driving the measurement (from "Test with:" line) |
| **Story theme** | A 3-5 word label for what this hypothesis is about (e.g., "message recall by setting", "VA impact on recall", "close rate vs TAG") |
| **Segment cuts** | Which splits are needed — read ONLY from the "Test with:" line, not from rationale prose |
| **Flags** | METHODOLOGY ARTIFACT and/or [ACTION ITEM] if present |

**Critical tagging rule — rationale is not instruction:** The rationale paragraph explains *why* the hypothesis is directional. It is not an analytical specification. Segment cuts, question codes, and chart design come exclusively from the **"Test with:"** line. If Practice Setting, HII, or any other segment is mentioned only in the rationale prose and not in the "Test with:" line, do NOT add it as a cut on the slide.

Do not display the full tag table — use it internally to drive clustering.

---

## STEP 4 — Cluster into slides

**Grouping rule:** Hypotheses that share the same primary question code AND the same story theme belong on the same slide.

For each cluster:
- Assign a **slide title** — a short, declarative label (not a question; not a finding — a topic label, e.g., "Message Recall by Practice Setting")
- Identify the **driving question** — the single analytical question this slide must answer (e.g., "Are community HCPs recalling OS and CNS messages at the same rate as academic HCPs?")
- List the **hypotheses tested** (e.g., H6, H11)
- Assign the **slide_type** — use the exact value from the Chart Type → slide_type Mapping table below. This must be a valid renderer key, not a generic description.
- List **question codes with text** — from Survey Context
- List **segment cuts** needed
- Write a **narrative arc** — one sentence: what the slide should conclude if hypotheses hold
- Flag any **methodology artifacts** or **action items**

If a hypothesis cluster is too large for one slide, split it. If two clusters tell the same story with the same chart, merge them.

### Chart Type → `slide_type` Mapping

Use this table to select the correct `slide_type` for each slide. Choose based on the data pattern and analytical goal — not the generic chart name.

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
| `message_mbd` | Multi-column abacus for MBD (Motivation, Believability, Differentiation) | ME question with sub-dimensions (M/B/D), plus composite effectiveness |

**Selection decision tree:**
```
Single brand, single metric, ranked list?
  → single_bar_with_delta

Same metric, two brands or segments compared?
  → clustered_compare (overlaid bars)
  → dual_bar_compare (side-by-side separated charts)
  → two_section_bar (stacked vertically)

Two related metrics for same brand (e.g., MR + ME)?
  → dual_bar_with_delta

QoQ comparison (current vs prior period)?
  → qoq_bar_with_delta (single brand)
  → dual_bar_qoq (two brands, side-by-side)

Recall order breakdown (1st/2nd/3rd)?
  → stacked_order

Attribute ratings / rep performance?
  → abacus (with value columns and delta)

Message effectiveness with M/B/D sub-dimensions?
  → message_mbd

Text-only insights?
  → executive_summary
```

---

## STEP 4.5 — Deduplicate across slides

Before sequencing, audit the full cluster list for question-level redundancy. A question code appearing on multiple slides is only acceptable if each slide uses it in a genuinely different analytical role. Apply the following tests:

**Test 1 — Same question, same cut, different section?**
If the same question code appears with the same segment cut on two slides in different sections, merge them into one slide and place it at the section boundary or the section where it is most analytically central.

**Test 2 — Same question, different cut, same section?**
If the same question code appears on two slides in the same section with different cuts (e.g., overall on one slide, by Practice Setting on another), merge into one slide with an overall panel and a segment-split panel. Only keep as two slides if the driving questions are genuinely distinct and the segment finding is complex enough to require its own narrative.

**Test 3 — Same outcome variable across multiple driver slides?**
If the same outcome question (e.g., Q1.85b LTIP, derived HII) is the dependent variable on three or more slides in the same section, consolidate into one driver-analysis slide with the predictors shown side by side — unless each predictor is a distinct client action item requiring standalone presentation.

**After applying these tests:**
- List any merges made and which hypotheses were consolidated
- Confirm every hypothesis still appears on at least one slide after merging
- Proceed to sequencing only once deduplication is complete

---

## STEP 5 — Sequence slides

Order slides into a narrative flow using KBQ domain sequence as the backbone:

1. Sales Activity
2. Messaging — Recall & Delivery
3. Messaging — Effectiveness
4. Interactions — Setup & Tactics
5. Rep Performance
6. Prescription Intent / LTIP
7. Branded Close / CTA
8. Drivers of High Impact

Within each section, order slides from most structural (overall comparison, QoQ) to most diagnostic (segment splits, driver analysis).

Place methodology artifact slides adjacent to the substantive slide they affect — not in an appendix. Flag them clearly so the analyst knows to verify before presenting.

---

## OUTPUT FORMAT

```
# Slide Plan — [Wave]
**Generated from:** [Hypothesis Bank] + [KBQs] + [Survey Context]
**Date:** [today's date]
**Total slides:** [N]

---

## SECTION: [Domain Name]

---

### Slide [N] — [Slide Title]
**Driving question:** [The single analytical question this slide must answer]
**Hypotheses tested:** H[x], H[y], H[z]
**slide_type:** [exact value from the Chart Type → slide_type Mapping table, e.g., `single_bar_with_delta`]
**Chart description:** [Specific layout details: sorted by what, delta column, segment splits, etc.]
**Primary questions:**
- [Q code] — "[Question text]"
- [Q code] — "[Question text]"
**Cuts needed:** [QoQ / Practice Setting / HII vs Others / etc.]
**Narrative arc:** [One sentence: what this slide concludes if hypotheses hold]
METHODOLOGY ARTIFACT — [brief description] [only if applicable]
[ACTION ITEM] — [brief description] [only if applicable]

```

Repeat Slide [N] blocks continuously. Slide numbers do not restart per section.

---

## OUTPUT FILE

After generating the slide plan in the conversation, write it to a file named `Slide_Plan_[Wave].md` in the same directory as the Hypothesis Bank. Ask the user to confirm the output path before writing.

---

## RULES

1. **One driving question per slide.** If a slide is trying to answer two unrelated questions, split it.
2. **Every hypothesis must appear on at least one slide.** If a hypothesis has no cluster match, create a standalone slide for it.
3. **Methodology artifact hypotheses stay adjacent** to the slide they affect — not hidden in appendix.
4. **Slide titles are topic labels, not findings.** Findings go in the narrative arc.
5. **slide_type must be a valid renderer key.** Use the exact value from the Chart Type → slide_type Mapping table (e.g., `single_bar_with_delta`, not "horizontal bar"). Add specifics in the **Chart description** field (e.g., "sorted by Q1 MR % descending, with QoQ delta column").
6. **Action items surface visibly.** Any slide derived from a client action item gets the [ACTION ITEM] flag.
7. **No question code appears on two slides serving the same analytical purpose.** If the same question with the same cut appears on multiple slides, it must be merged. A question code may appear on multiple slides only when each use is a genuinely different analytical role (e.g., Q2.10 as a leaderboard vs. Q2.10 as a predictor of LTIP).
8. **No context stored in this skill.** Everything comes from reading the files at runtime.
9. **Rationale prose does not define chart design.** The hypothesis rationale explains the mechanism — it is context, not instruction. Segment cuts, cuts needed, and question codes are sourced exclusively from the "Test with:" line of each hypothesis. Do not add a cut (e.g., Practice Setting, HII vs Others) because it appears in the rationale — only add it if the "Test with:" line specifies it.
10. **Show all dimensions of a multi-part question.** If a question has sub-dimensions (e.g., ME Q2.20 has A=Differentiation, B=Believability, C=Motivation), always include all dimensions in the slide spec and chart description unless the hypothesis bank or user explicitly restricts to a specific sub-dimension. Never silently drop dimensions.
11. **Use exact question text, not question codes.** In the **Primary questions** field, always write the full question text sourced from the Survey Context file alongside the code. Do not use shorthand or paraphrase.
