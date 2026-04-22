# Refresh Workflow + Evals Tracker

**Owner:** Bharadvaj · **Last updated:** 2026-04-21

This single file replaces `refresh_workflow_tracker.xlsx` and `evals_coverage_tracker.xlsx`. It tracks the wave-refresh workflow, evals coverage, and maps who (Vijay / Bharadvaj / pending) did what across the End-of-April PRD milestones.

**Status legend:** ✅ DONE · 🟡 PARTIAL / IN PROGRESS · ⏳ PENDING · 🚫 BLOCKED · ❌ FAIL · ℹ️ INFO

---

## 1. Quick Status — where are we right now?

- **End-of-April PRD targets:** 1 of 4 fully done; 1 partial; 1 in progress (Bharadvaj Item B executed today, verdict pending); 1 skipped (not Bharadvaj lane)
- **Step B (run Vijay refresh pipeline on real PET deck):** Stages 1-3 executed on ATU Q1'26. Findings: pipeline works mechanically but data fidelity is broken (21% category match). Three distinct bug categories surfaced — all in Vijay's mapper.
- **Step A (port to evals):** Next up. Will port Stage 3 into `tests/evals/end_to_end/` with mutation-test validation.
- **Evals coverage:** 1 of 40 covered (deck-reader tag counts regression — positive + mutation test passing).
- **Blockers:** None operationally. Bug fixes for the 3 mapper bugs are Vijay's work.

---

## 2. End-of-April Milestone — PRD §9.3 breakdown

Target text from PRD §9.3: *"Rendering fidelity complete + evals bootstrapped. Visual regression passing on top-6 chart patterns × representative brands. Line/doughnut Repair bug closed. Connector-tag integration tested on real tagged PET deck. Evals harness started (`tests/evals/`) with ≥1 refresh eval using prior JJ RYB deck."*

### T1 — Visual regression on top-6 chart patterns × representative brands 🟡

| Sub-step | Owner | Status | Done date | Source / Notes |
|---|---|---|---|---|
| 905-deck grounding analysis | Vijay | ✅ | 2026-04-17 | Commits `0757f03`, `f9a743a`, `a44dc9a`. 905 decks analyzed. |
| pptx_utils regrounding | Vijay | ✅ | 2026-04-17 | Commit `f9a743a`. |
| Visual regression harness in `tests/` | Pradeep | ⏳ | — | Not Bharadvaj lane. Biggest lift of the 4 targets. |
| Top-6 pattern × brand goldens | Pradeep | ⏳ | — | Depends on harness. |

### T2 — Line/doughnut Repair bug closed ✅

| Sub-step | Owner | Status | Done date | Source / Notes |
|---|---|---|---|---|
| Fix PowerPoint Repair bugs | Vijay | ✅ | 2026-04-16 | Commit `59e1412`. |

### T3 — Connector-tag integration tested on real tagged PET deck 🟡

| Sub-step | Owner | Status | Done date | Source / Notes |
|---|---|---|---|---|
| Fix Tier 1 tag extraction from real decks | Vijay | ✅ | 2026-04-16 | Commit `b56f441`. |
| Build 3-stage refresh test pipeline | Vijay | ✅ | 2026-04-17 | Commit `56bb2b4` (`test_spec_refresh_pipeline.py`). |
| SlideSpec v1.2 spec-as-config + mapper rewrite | Vijay | ✅ | 2026-04-20 | Commit `973d843`. 25/27 verified on UAT deck. |
| Run pipeline on ATU Q1'26 deck (Item B) | Bharadvaj | ✅ | 2026-04-21 | Stages 1-3 executed. See §6 for results. |
| Analyze pass rate + failures | Bharadvaj | ✅ | 2026-04-21 | 3 bug categories identified. See §6. |
| Write verdict + handoff to Vijay | Bharadvaj | ⏳ | — | Structured report for Vijay (B14). |

### T4 — Evals harness + ≥1 refresh eval 🟡

| Sub-step | Owner | Status | Done date | Source / Notes |
|---|---|---|---|---|
| `tests/evals/` folder scaffolded | Bharadvaj | ✅ | 2026-04-21 | `fixtures.py`, `generate_golden.py`, `__init__` files. |
| Eval #1 — tag counts (regression) | Bharadvaj | ✅ | 2026-04-21 | Positive + mutation test passing. File: `tests/evals/deck_reader/test_tag_counts.py`. |
| Refresh eval — end-to-end on real deck | Bharadvaj | ⏳ | — | Item A. Ports Stage 3 into `tests/evals/end_to_end/`. Needs to DEEPEN check beyond categories. |
| Mutation-test the refresh eval | Bharadvaj | ⏳ | — | Same workflow as Eval #1. |

---

## 3. Component Dashboard — pipeline parts at a glance

| Component | Status | Owner | Last tested | Notes |
|---|---|---|---|---|
| deck-reader — Tier 1 tag extraction | ✅ DONE | Vijay + Bharadvaj | 2026-04-21 | ATU: 55/55 tagged slides have complete lineage. |
| deck-reader — Tier 2 inference | 🟡 PARTIAL | Vijay | 2026-04-20 | Prototype in `data_inference.py` (commit `973d843`). Tested on Repatha Slide 6. Not production-ready. |
| Synapse token / API auth | ✅ DONE | Vijay + Bharadvaj | 2026-04-21 | Vijay built `synapse_auth.py` (commit `2b88c92`). Bharadvaj installed `sk_` API key; verified HTTP 200 on `/api/projects`. |
| synapse-read (fetch via Synapse API) | ✅ DONE | Vijay | 2026-04-17 | `fetch_synapse_records` working. Tested in UAT + ATU pipelines. |
| synapse_chart_mapper (records → chart data) | 🟡 PARTIAL | Vijay | 2026-04-21 | Full pivot-faithful rewrite (446 lines). **3 bugs surfaced on ATU run** — see §6. |
| slide-updater (in-place chart refresh) | 🟡 PARTIAL | Vijay | 2026-04-21 | `replace_data()` works. 114/148 charts refresh successfully; only 42/197 match source categories. |
| headline-writer (regenerate on data change) | ⏳ PENDING | (unassigned) | — | Not built yet. May milestone. Vijay currently restores original headlines. |
| deck-assembler (final PPTX output) | ✅ DONE | Vijay | 2026-04-20 | In-place save works. No separate assembler needed for refresh path. |
| End-to-end refresh — UAT deck | ✅ DONE | Vijay | 2026-04-20 | 25/27 slides visually correct. |
| End-to-end refresh — real PET client deck | 🟡 PARTIAL | Bharadvaj | 2026-04-21 | ATU: 42/197 category match. Data fidelity broken by wave-filter bug. |
| Evals harness — `tests/evals/` scaffolded | ✅ DONE | Bharadvaj | 2026-04-21 | Folder + fixtures + Eval #1 + golden + mutation test. |
| Evals — Eval #1 tag counts | ✅ DONE | Bharadvaj | 2026-04-21 | Short-circuit bug caught and fixed during mutation test. |
| Evals — refresh eval per PRD §6.9 | ⏳ PENDING | Bharadvaj | — | The PRD-named end-of-Apr deliverable. Item A. |
| Visual regression harness | ⏳ PENDING | Pradeep / Vijay | — | PRD target #1. Not Bharadvaj lane. |

---

## 4. Activity Log — chronological record

### Vijay contributions (from git)

| Date | Activity | Result | Source |
|---|---|---|---|
| 2026-04-16 | Fix PowerPoint Repair bugs (line + doughnut) | ✅ Hits PRD T2 | `59e1412` |
| 2026-04-16 | Fix Tier 1 tag extraction from real Synapse decks | ✅ Foundation for refresh | `b56f441` |
| 2026-04-17 | Add human-readable name fields to DataLineage | ✅ | `ea93fc8` |
| 2026-04-17 | Build Synapse chart mapper (first version) | ✅ | `2c9067c` |
| 2026-04-17 | Build 3-stage refresh test pipeline scaffold | ✅ | `56bb2b4` |
| 2026-04-17 | MSAL + .env token auth with auto-expiry | ✅ Auth infra complete | `2b88c92` |
| 2026-04-20 | SlideSpec v1.2 — Connector-faithful refresh | ✅ 25/27 on UAT | `973d843` |

### Bharadvaj contributions

| Date | Step | Activity | Deck / target | Result | Notes |
|---|---|---|---|---|---|
| 2026-04-20 | B1 | Run deck-reader CLI on ATU Q1'26 | ATU (99 slides) | ℹ️ | Generated 99 SlideSpec JSONs. CLI path does not populate data_mapping. |
| 2026-04-20 | B2 | Tier 1 spec quality audit | ATU | ✅ 55/55 complete | All have project_id, reporting_plan_id, analysis_ids, survey_id. |
| 2026-04-21 | B3 | Re-run via `generate_config_specs()` API | ATU | ✅ | CLI path vs library path produce different output — flagged for Vijay. |
| 2026-04-21 | B4 | Scaffold `tests/evals/` + fixtures | evals infrastructure | ✅ | Folder tree, `fixtures.py`, `generate_golden.py`. |
| 2026-04-21 | B5 | Build Eval #1 — tag count regression | ATU golden | ✅ positive | Located at `tests/evals/deck_reader/test_tag_counts.py`. |
| 2026-04-21 | B6 | Validate Eval #1 via mutation test | ATU golden (mutated) | ✅ | Found + fixed short-circuit bug; all 3 mutations reported clearly. |
| 2026-04-21 | B7 | Build evals coverage tracker | local xlsx | ✅ | 40 evals catalogued across 9 pipeline stages. Now migrated to this markdown. |
| 2026-04-21 | B8 | Install Synapse token in `.env` | synapse_auth | ✅ | `sk_` API key works as Bearer token. Vijay's JWT-decode exception handler passes it through unchanged. |
| 2026-04-21 | B9 | Parameterize test for ATU deck | `test_spec_refresh_atu.py` | ✅ | Option 1: copied Vijay's test, edited 4 paths. |
| 2026-04-21 | B10 | Run Stage 1 — dummy deck creation | ATU | 🟡 143/170 charts | 27 scatter/XY charts failed with `CategoryWorkbookWriter.x_values_ref`. Bug #1. |
| 2026-04-21 | B11 | Run Stage 2 — Synapse fetch + refresh | ATU | 🟡 114/148 charts | 26 Synapse API calls. 34 charts failed with field-name mismatches. Bug #2. |
| 2026-04-21 | B12 | Run Stage 3 — verify vs source | ATU | 🟡 42/197 categories match | Revealed wave-filter + label-transformation bugs on 155 charts. Bug #3. |
| 2026-04-21 | B13 | Analyze failures | ATU | ✅ | 3 distinct bug categories documented in §6. |
| 2026-04-21 | B-migrate | Migrate trackers to markdown | local | ✅ | `.xlsx` files being retired; this file is the new source of truth. |

### Pending Bharadvaj steps

| Step | Activity | Notes |
|---|---|---|
| B14 | Write structured verdict / findings report for Vijay | One document bundling the 3 bug categories with repro steps. |
| A1 | Port Stage 3 into `tests/evals/end_to_end/` | Wrap as pytest eval with golden file for pass rate. |
| A2 | Deepen the comparison beyond categories | Add series + value comparison. Stage 3 currently misses silent-corruption bugs. |
| A3 | Mutation-test the refresh eval | Same workflow as Eval #1: corrupt inputs, verify eval catches. |
| A4 | Mark Eval #36/#37 COVERED in tracker | Hits PRD T4. End-of-Apr deliverable for Bharadvaj. |

---

## 5. Evals Coverage

### Coverage Dashboard

| Pipeline Stage | Total Evals | Covered | Coverage % | Priority |
|---|---|---|---|---|
| Stage 0 — deck-reader (tag extraction, SlideSpec) | 6 | **1** | 17% | P0 |
| Stage 1 — prior-wave-context-builder | 3 | 0 | 0% | P2 |
| Stage 2 — synapse-read (fetch via API or banner plan) | 4 | 0 | 0% | P0 (was blocked on token) |
| Stage 3 — slide-plan-gen-refresh (DIFF UPDATE/ADD/DELETE) | 5 | 0 | 0% | P1 (component not built yet) |
| Stage 4 — slide-updater (in-place refresh) | 6 | 0 | 0% | P0 |
| Stage 4 — headline-writer | 4 | 0 | 0% | P1 (component not built yet) |
| Stage 5 — spec-validator | 3 | 0 | 0% | P2 |
| Stage 5 — slide-creator / deck-assembler | 4 | 0 | 0% | P1 (shared with new-deck) |
| End-to-End — full refresh run | 5 | 0 | 0% | P0 |
| **TOTAL** | **40** | **1** | **2.5%** | |

### Catalog — each eval, in priority order

**P0 evals (must have):**

| # | Stage | Eval | Status | What it checks |
|---|---|---|---|---|
| 1 | deck-reader | Tag count per slide | ✅ COVERED | Runs deck-reader, diffs tagged/untagged counts vs committed golden. |
| 2 | deck-reader | ReportConfig/PivotConfig/MappingConfig resolution | ⏳ | Every tagged shape has all 3 configs resolved. Current ATU: 197/199 pivot, 199/199 mapping. |
| 3 | deck-reader | DataLineage field completeness | ⏳ | All tagged slides: project_id, reporting_plan_id, analysis_ids, survey_id populated. |
| 4 | deck-reader | SlideSpec v1.2 schema compliance | ⏳ | Every spec parses via `load_spec()` without schema errors. |
| 5 | deck-reader | Golden snapshot regression | ⏳ | Full spec JSON diff against committed golden. |
| 10 | synapse-read | Data fetch for known analysis ID | ⏳ | Confirms fetch_synapse_records returns expected schema. **NOW UNBLOCKED** (token installed). |
| 11 | synapse-read | Wave filtering correctness | ⏳ | `dynamic_latest_n=2` returns exactly 2 waves. **CRITICAL** — Stage 3 finding shows this is broken today. |
| 12 | synapse-read | Segment filter application | ⏳ | When lineage has segment_ids, records filtered accordingly. |
| 19 | slide-updater | Same-wave fidelity (round-trip) | ⏳ | Vijay UAT baseline (25/27). Today's ATU baseline: 42/197. |
| 20 | slide-updater | Cross-wave update (Q1→Q2) | ⏳ | Q1 deck + Q2 data produces Q2 values. Deferred — needs headline-writer first. |
| 21 | slide-updater | Layout preservation | ⏳ | Chart types, positions, fonts unchanged after refresh. |
| 22 | slide-updater | Tables refresh correctly | ⏳ | Label + value tables updated. |
| 24 | slide-updater | Formatting restoration | ⏳ | Per-element formatCode (%, $, #) preserved. Vijay's commit notes this as a fix. |
| 36 | end-to-end | Full refresh on ATU deck | ⏳ | **ITEM A — port Stage 3 + deepen.** End-of-Apr deliverable. |
| 37 | end-to-end | Full refresh on UAT deck (regression) | ⏳ | Lock in Vijay 25/27 as regression baseline. |

**P1 evals (should have):**

| # | Stage | Eval | Status | What it checks |
|---|---|---|---|---|
| 13 | synapse-read | Error handling — expired token | ⏳ | Expired token fails with clear error. |
| 14-18 | slide-plan-gen-refresh | UPDATE / ADD / DELETE / no-op / idempotency | ⏳ | Component not yet built. |
| 23 | slide-updater | Split visualization handling | ⏳ | splitGroupID / splitOrder / rowsPerObject correct. |
| 25-28 | headline-writer | Freshness / factual grounding / length / arc consistency | ⏳ | Component not yet built. |
| 32-35 | slide-creator / deck-assembler | Renderer mapping / brand / slide order / registry | ⏳ | Shared with new-deck lane. |
| 38 | end-to-end | Idempotency | ⏳ | Same inputs produce identical output. |
| 40 | end-to-end | Graceful failure modes | ⏳ | Synapse down, token expired, etc. clear errors, no partial output. |

**P2 evals (nice to have):** #6 (Tier 2 inference confidence), #7–9 (prior-wave-context), #27–28 (headline length/arc), #29–31 (spec-validator), #39 (performance baseline).

---

## 6. Stage 3 Findings — ATU Deck (2026-04-21)

**Context:** Ran Vijay's full 3-stage refresh pipeline (`test_spec_refresh_pipeline.py` logic, ATU paths) against the real ATU Q1'26 client deck. First validation of the pipeline outside Vijay's internal UAT fixture.

**Deck:** `projects/J&J Rybrevant PET/Template/ZoomRx_UC_ATU_Report_Q1_'26.pptx` — 99 slides, 197 charts, 209 tables, 199 tagged shapes.

### Headline numbers

| Metric | Count | Interpretation |
|---|---|---|
| Stage 1 — charts dummied | 143 / 170 | 27 scatter/XY charts failed — Bug #1 |
| Stage 2 — charts refreshed without error | 114 / 148 | 34 field-name mismatches — Bug #2 |
| Stage 3 — charts with matching categories | 42 / 197 | Wave-filter + label bugs — Bug #3 |
| Stage 3 — headlines match | 56 / 99 | Expected (55 Tier 1 + untouched untagged) |

**IMPORTANT:** Stage 3's 21% match rate is a **categories-only** check (row labels only). It does NOT verify series names or cell values. The 42 matches are charts where categories are non-wave dimensions (drug names, attributes); their series values may still be corrupt. True data-correctness could be lower than 21%.

### Bug #1 — Scatter/XY chart dummying fails in Stage 1

**Error:** `'CategoryWorkbookWriter' object has no attribute 'x_values_ref'`
**Location:** `tests/test_spec_refresh_pipeline.py` — Stage 1 `stage1_create_dummy_deck()`
**Root cause:** Stage 1 uses `CategoryChartData` to wipe chart values. This works for bar/column/line charts but not for XY scatter charts, which need `XyChartData`.
**Severity:** Medium. Impact is that 27 scatter charts retain their source data in the dummy deck. Stage 2 should still overwrite them with real data — but it's untested, and if Stage 2 has the same gap, those charts fall back to source silently.
**Affected slides:** 22, 23, 24, 25, 30, 33, 34, 35, 44, 45, 46, 74, 86, 90 (15 slides, 27 charts).
**Fix needed:** Stage 1 must detect XY/scatter charts and use `XyChartData` (or skip dummying and trust Stage 2 to overwrite).

### Bug #2 — Unknown field names in PivotConfig

**Error pattern:** `RowFields '['X']' not in records` / `ValueField '['X']' not in records`
**Location:** `slidegen/synapse_chart_mapper.py` — `pivot_records_to_chart_data()` or `apply_spec_transform()`
**Root cause:** SlideSpec's `raw_pivot_config` specifies field names (`decimal`, `options`, `y_label`) that aren't present in the Synapse API response for those analyses. The mapper assumes a specific field-name schema that doesn't hold for all analysis types.
**Severity:** High. Affects segmented analyses and some chart patterns on real decks.
**Affected slides and fields:**

| Slide | Failing field | Analysis |
|---|---|---|
| 6 | ValueField `['decimal']` | 412687 |
| 36, 37, 38 | RowFields `['options']` | 411829 with segments=(4779,) |
| 47, 48, 49 | RowFields `['options']` | 411829 / 680412 with segments=(4745,) |
| 56 | RowFields `['y_label']` | 613146 |

**Fix needed:** Mapper should handle a broader set of Synapse field names, or fall back gracefully when an expected field is missing (e.g., check for aliases `options` ↔ `value` ↔ `title`).

### Bug #3 — Wave filtering + label transformation broken (most critical)

**Error:** No error raised — charts refresh "successfully" but with wrong data.
**Location:** `slidegen/synapse_chart_mapper.py`
**Root cause:** Two separate sub-bugs that compound:
1. **No `dynamic_latest_n` filtering.** SlideSpec says `dynamic_latest_n=2` but mapper returns all waves from Synapse.
2. **No label transformation.** Synapse returns `"Project Wave 11"`, `"Project Wave 12"`. Source deck shows `"Wave 11"`, `"Wave 12"`. The source deck applies a label-shortening rule at Connector time; the mapper doesn't replicate it.

**Severity:** CRITICAL. Produces silent data corruption — charts look plausible but show wrong time periods with wrong labels.

**Representative sample (from Stage 3 output):**

| Slide | Chart position | Source categories | Refreshed categories |
|---|---|---|---|
| 4 | (8.36, 4.51) | `['Wave 11', 'Wave 12']` | `['Project Wave 1', 'Project Wave 10', 'Project Wave 12']` |
| 4 | (2.81, 4.5) | `['Wave 11', 'Wave 12']` | `['Project Wave 1', 'Project Wave 10', 'Project Wave 12']` |
| 5 | (10.85, 2.11) | `['Wave 11', 'Wave 12']` | `['Project Wave 1', 'Project Wave 10', 'Project Wave 12']` |

**Affected:** At least 155 of 197 charts (77%).

**Fix needed:** Mapper must (a) filter returned records to last N waves per SlideSpec's `dynamic_latest_n`, (b) apply a configurable label-transformation rule (or read it from Connector metadata if present).

### Verdict

**Vijay's refresh pipeline is mechanically functional but NOT yet production-ready on real client decks.** The jump from UAT (25/27 visually correct) to ATU (42/197 category match) is large and reveals three real bugs — all in `synapse_chart_mapper.py`. None are architectural problems; all are missing features or field-name gaps.

**Stage 3's own shallowness is also a finding:** category-only comparison misses silent corruption in series/values. The refresh eval we build for Item A must strengthen this.

**Recommendation to Vijay:** prioritize Bug #3 (wave filter + label transform), then Bug #2 (field name gaps), then Bug #1 (XY dummying). Bug #3 alone unlocks the bulk of the 155 mismatches.

---

## 7. How to update this file

**Every time Vijay commits relevant work:**
1. `git log upstream/vijay-slidegen --since="2 days ago" --oneline`
2. Add a row to §4 "Vijay contributions" with commit + result
3. If the commit closes an End-of-April milestone sub-step, flip `⏳` → `✅` in §2
4. If the commit changes a component's state, update §3 Dashboard

**Every time Bharadvaj does work:**
1. Add a row to §4 "Bharadvaj contributions" (Date / Step / Activity / Deck / Result / Notes)
2. Update §2 and §3 if state changes
3. If it's an eval, update §5 (both dashboard and catalog)

**When something gets blocked:**
- Mark `🚫 BLOCKED` on the row, with the blocker in Notes
- Still leave the row visible so the blocker is tracked

**When a finding emerges (like §6):**
- Add a new numbered section under the main structure
- Cross-reference from the Activity Log row that produced it

**Git:** This file is currently untracked. Add to git only when you choose to.
