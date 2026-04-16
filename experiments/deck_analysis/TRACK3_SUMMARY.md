# Track 3 — Spec Producers (Summary)

**Date:** April 16, 2026
**Status:** Python primitives + SKILL.mds + skill reorg landed. Tests pending.

## What was built

### Python primitives (2 files)

| File | Purpose |
|---|---|
| `slidegen/viz_selector.py` | Deterministic chart-pattern selector. 36 metric tags mapped from Apr 15 deck-analysis frequencies. Priority hierarchy: Metric tag → Q-type → `UNRESOLVED` (HITL). Pure function — no LLM. |
| `slidegen/layout_selector.py` | Deterministic layout selector. Component-type list → LAYOUTS key. Rules from real-deck cluster counts. `delta_column` treated as structural (not counted as table). |

### SKILL.md files (7 new)

| Skill | Path |
|---|---|
| `viz-selector` | `.claude/skills/planning/viz-selector/` |
| `layout-selector` | `.claude/skills/planning/layout-selector/` |
| `headline-writer` | `.claude/skills/creation/headline-writer/` |
| `callout-writer` | `.claude/skills/creation/callout-writer/` |
| `slide-plan-generator-refresh` | `.claude/skills/planning/slide-plan-generator-refresh/` (stub — forward contract for Workflow 2) |
| `slide-plan-generator-single` | `.claude/skills/planning/slide-plan-generator-single/` (stub — forward contract for Workflows 3/7/8) |
| `slide-plan-generator-exec-summary` | `.claude/skills/planning/slide-plan-generator-exec-summary/` (stub — forward contract for Workflow 9) |

**Not split: `sfea-insight-writer`.** Earlier iteration briefly created wrapper stubs at `validated-analysis-writer/` and `narrative-threads-builder/` to reflect the PRD's two-phase split. Those wrappers were removed — they added cognitive load without providing real separation, since both pointed to the unchanged `sfea-insight-writer`. The skill stays as a single unit with two internal phases until a concrete workflow needs one phase alone. See feedback memory `feedback_no_speculative_splits.md`.

### Skill reorganization (8 moves)

| From (flat) | To (PRD §8.1 structure) |
|---|---|
| `.claude/skills/build-project-context/` | `.claude/skills/context-data/project-context-builder/` |
| `.claude/skills/market-context/` | `.claude/skills/context-data/market-context-builder/` |
| `.claude/skills/prior-wave-context/` | `.claude/skills/context-data/prior-wave-context-builder/` |
| `.claude/skills/survey-context/` | `.claude/skills/context-data/survey-context-builder/` |
| `.claude/skills/hypotheses/` | `.claude/skills/analysis/hypothesis-generator/` |
| `.claude/skills/slide-plan/` | `.claude/skills/planning/slide-plan-generator-hypothesis/` |
| `.claude/skills/pet-es-builder/` | `.claude/skills/projects/pet-deck/` (rename reflects broader role; J&J-specific content still to be generalized) |
| `.claude/skills/sfea-insight-writer/` | `.claude/skills/analysis/sfea-insight-writer/` (intact — kept as single skill with internal Phase 0 + Phase 1) |
| `.claude/skills/slide-creator/` | `.claude/skills/creation/slide-creator/` |
| `.claude/skills/spec-validator/` | `.claude/skills/planning/spec-validator/` |

Left in place: `.claude/skills/slidegen/` (legacy general skill), `.claude/skills/pptx/` (general PPTX primitive skill).

## viz-selector metric → pattern mapping

36 metric tags pre-mapped from Apr 15 deck-analysis of 4,354 real PET charts:

| Metric family | Chart pattern | Real-deck observed |
|---|---|---|
| message_recall, aided_awareness, unaided_awareness, brand_awareness, message_recognition, recall, recognition, top_of_mind, topic_recall | `bar_clustered_horizontal` | 35% |
| message_effectiveness, motivation, believability, differentiation, mbd | `xy_scatter_abacus` | 17% |
| rep_performance, call_quality, rep_attributes, attribute_rating | `xy_scatter_abacus` | 17% |
| awareness_trend, recall_trend, reach, share_of_voice, frequency, trended_scorecard, activity_trend | `line_markers_trended` | 14% |
| likelihood_to_prescribe, prescription_intent, intent_distribution, patient_allocation, share_of_mind | `column_stacked_100_vertical` | 11% |
| call_to_action, branded_close, interaction_format | `bar_stacked_100_horizontal` | 7% |
| segment_comparison, hii_scorecard, brand_comparison | `column_clustered_vertical` | 7% |

**Total: 36 mapped metrics covering 88%+ of observed PET chart production.**

Q-type fallback:
- likert → bar_clustered_horizontal
- ranking → xy_scatter_abacus
- time_series → line_markers_trended
- categorical_composition → column_stacked_100_vertical
- ordinal / multi_select / single_select → bar_clustered_horizontal
- numeric_scale → xy_scatter_abacus
- stacked_composition → bar_stacked_100_horizontal

## What works end-to-end

- `viz_selector.select_chart_pattern("Message Recall")` → `"bar_clustered_horizontal"` ✓
- `viz_selector.select_chart_pattern("awareness_trend")` → `"line_markers_trended"` ✓
- `viz_selector.select_chart_pattern(q_type="likert")` → `"bar_clustered_horizontal"` ✓
- `viz_selector.select_chart_pattern(metric="random_unknown")` → `"UNRESOLVED"` ✓
- `layout_selector.select_layout(["chart", "label_table", "delta_column"])` → `"observed_1chart_1table"` ✓ (after delta_column structural fix)
- `layout_selector.select_layout(["chart", "label_table", "value_table"])` → `"observed_1chart_2table"` ✓
- `layout_selector.select_layout(["chart", "chart", "label_table", "label_table"])` → `"observed_2chart_2table"` ✓
- All 10 existing skills relocated into PRD §8.1 folder structure ✓

## Fix applied during integration

`layout_selector.py` originally counted `delta_column` as a table component, which caused `chart + label_table + delta_column` (the canonical bar_clustered_horizontal pattern matching our example spec) to return `observed_1chart_2table` instead of `observed_1chart_1table`. Fixed by treating `delta_column` as structural (pairs with chart as part of the layout preset). Verified against the canonical example spec.

## Known gaps

1. **sfea-insight-writer intentionally kept whole.** The 897-line skill implements both Phase 0 (hypothesis validation against data) and Phase 1 (narrative synthesis + arcs + headlines + ES) in one unit, which is how Vinoth authored it. The two phases share heavy context so file-based decoupling would lose fidelity. We will only split if a workflow emerges that needs one phase alone. Earlier wrapper stubs were removed — see feedback memory.
2. **pet-deck (née pet-es-builder) still has J&J-specific content.** Rename done; content generalization pending.
3. **3 slide-plan-generator SKILL.mds are contracts only.** Implementations awaiting convergence with Track 2 (deck-reader) and headline-writer.
4. **Test scripts under `tests/track3/` not yet written.** Agent hit permissions wall.
5. **headline-writer and callout-writer have no Python backing** by design (LLM skills). Will need to be invoked from higher-level orchestration.

## Contract for upstream/downstream integration

- Every producer's output passes `validate_spec(strict=True)` where the output is a full SlideSpec.
- `viz-selector` and `layout-selector` are pure functions — zero side effects, safe to call anywhere.
- `headline-writer` and `callout-writer` emit field values (string, CalloutComponent), not full specs.
- Deterministic-first everywhere — no LLM calls in viz-selector / layout-selector / slide-plan-generator-single's orchestration logic (only in headline/callout sub-steps).
