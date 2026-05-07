# Connector Path — Session Memory

Date-wise log of decisions, fixes, and pending items on the connected-refresh
pipeline. New sessions should read the most recent date heading first to pick
up context, then review prior days for continuity.

---

## 2026-05-07

**Context.** The user reviewed Repatha HCP ATU v10 slide-by-slide and
called out concrete failures across slides 11, 13, 15, 18, 20, 49, 65-67,
and 73. Frustration was high — earlier sessions had reported "ok"
statuses on tables that had silently never been written, so the user
asked for fixes that work universally (not deck-specific) and a clean
explanation of why each thing was failing.

### Issues raised + status after today

| Slide | Issue | Root cause | Status |
|---|---|---|---|
| 11 (tables) | Tables show source values, never refreshed | Spec didn't carry `cell_values`; refresh-time table path required them. "ok" status only meant the connector tag was processed. | **FIXED** — `_auto_compute_cell_values` builds a 2D grid from pivot output (1×N vertical, N×1 horizontal, M×N matrix). Validated byte-level: source `74%, 73%, 62%, 56%` → v9 `70%, 71%, 66%, 65%`. |
| 11 (n cards / n L / n VH/h / n M) | 1×1 cells with embedded count `(n = 120)` not refreshed | Templated cells need regex substitution, not direct write. | **FIXED** — `_refresh_templated_count_with_records` substitutes `(n = X)` from post-filter records. |
| 13 (headlines) | Slide says "no narrative — skipped" but should write fresh | Detector required existing narrative-verb match; LLM only updated; never wrote fresh. | **FIXED** — detector picks topmost wide title regardless of verb match; prompt switches to "write fresh narrative" when source isn't narrative-shaped. v10 jumped from 14 → 33 headlines rewritten. |
| 15 | Badge says "some refreshed, some not" — *which*? | Badge text was vague; per-component detail only in JSON sidecar. | **FIXED** — every slide's speaker notes now have an appended divider + per-component status block (existing notes preserved via lxml direct paragraph append, idempotent on re-run). |
| 18 (Overall Reach) | `alignment_failed: refreshed values are entirely None` | Source categories were full classifier paths (`"Specialty (C/PCP Segments) + Tier Groups - CARD"`); API returned short tail (`"CARD"`). Exact-match alignment found zero hits → all values nulled → safety net fired → preserve source. | **FIXED** — category alignment now uses 4-tier resolution (exact, progressive-suffix, tail-substring, fuzzy ≥ 0.85). |
| 20 (abacus / scatter) | Series labelled `"88.31"` instead of `"CARD - Believer"`; values render as `"8800%"` | Two bugs: (a) X-axis values scaled `*100` because half the chart's formatCodes lacked `%`. (b) Series naming was an artifact of yesterday's mapper, fixed implicitly by formula-matching + value-field measure-fallback. | **FIXED** — for scatter charts, if ANY formatCode contains `%`, disable `_scale_to_whole_percent`. Series names now correct. |
| 49 (cards / nonwriters / dabblers / reservers) | Series order swapped — Wave 7 (latest) source-first ended up second after refresh | Pivot extracts columns alphabetically (`Wave 6` before `Wave 7`); existing source-canonical reorder was gated by `not is_dynamic` so dynamic charts skipped. | **FIXED** — added a positional reorder for dynamic charts when source-series-name count == pivot-series count and each maps 1:1 (using the same 4-tier resolution). |
| 65-67 | "Is new data even being brought in?" | False alarm — refresh IS working. v10 values differ from source by ~0.001 (precision). | **CONFIRMED OK** — no fix needed. |
| 73 (headline) | Section header rewritten as headliner instead of the actual top-of-slide narrative | Old detector picked the first Title-named shape with a narrative verb; section header had `"Trended"` so it won. Real headliner above used `"exhibit a growing trend"` (no verb in `_NARRATIVE_VERBS`). | **FIXED** — detector picks topmost wide title shape regardless of verb. v10 slide 73 now correctly rewrites the top headliner with new narrative. |
| 73 (chart data) | All 4 charts (Chart 139, 22, 23, 24) got identical refreshed values | All 4 charts have BYTE-IDENTICAL connector tags. Tag spec cannot differentiate them — without per-chart filter or split_order, the same pivot output writes to every chart. Spec ambiguity, not a pipeline bug. | **GUARD ADDED** — when ≥ 2 charts on a slide share fingerprint AND source had different values per chart, mark `ambiguous_tag` and preserve source. User must add per-chart filter to enable refresh. |

### Universal fixes shipped (commits, in order)

| Commit | Date | What |
|---|---|---|
| `670160a` | 2026-05-06/07 boundary | Table auto-write (`_auto_compute_cell_values`), filter compound-suffix resolver for `;`-joined values, latest-wave-only when waves aren't a column dim, formula-column suffix matching, headline detector picks topmost wide title. |
| `285a2fc` | 2026-05-07 | `walk_shapes_recursive` — group-nested charts/tables included in refresh shape pools. |
| `c2c7caf` (and `5697fdb`/`d676a15` from prior day) | 2026-05-06 | Shape-id matching (replaces ±0.3" position tolerance), value-field measure-column fallback, formula columns evaluator, per-component static-pin resolution, honor live-wave intent on mixed-pin. |
| `947c4e0` | 2026-05-07 | Templated `(n = X)` cell substitution, slide 73 ambiguous-tag guard, slide 49 dynamic series reorder, slide 20 scatter scaling, per-component speaker-notes report. |
| `e1cf42b` | 2026-05-07 | Scatter scaling guard: don't `*100` when any X-axis formatCode is `%`. |
| `32fa930` | 2026-05-07 | Category alignment progressive-suffix matching (slide 18). |
| `30f8227` | 2026-05-07 | Fuzzy similarity (`SequenceMatcher.ratio() ≥ 0.85`) as 4th-tier fallback for category + series alignment. |

### Architectural decisions made today

1. **Shape identity = XML id, not position.** Position with ±0.3" tolerance was ambiguous on dense slides (Repatha ATU 11/12 had n-row tables 0.13" below value-row tables). Replaced with `_claim_shape(shape_id=..., pos=..., name=...)` — id wins, position+name fall back.
2. **Tables write through `cell_values` always.** The chart-equivalent `replace_data` path for tables didn't exist. Now `_auto_compute_cell_values` synthesizes a 2D grid from pivot output every time the spec doesn't carry pre-computed values.
3. **Resolution chain, applied universally:** exact → progressive suffix → tail substring → fuzzy ≥ 0.85. Same logic everywhere a label-to-label match happens (categories, series names, formula column refs).
4. **Ambiguous tag is user-side, not pipeline-side.** When 2+ charts share fingerprint and had different source values, preserve source and surface the ambiguity. Pipeline can't auto-disambiguate from the spec alone.
5. **Speaker notes preserved.** Per-component status appended below a divider via lxml direct paragraph append — don't use `tf.text =` (clobbers formatting on existing paragraphs). Idempotent on re-runs.
6. **"ok" status validated byte-level, not via JSON.** Hot lesson: status JSON saying "ok" is necessary but not sufficient. Always diff source-vs-output cell content before claiming a refresh worked.

### Honest answers to recurring questions

**"Why so many failures? Connected refresh is simple."**
Each new deck shape has revealed a different code path's edge case. Most common patterns:
- Source labels differ from API labels by Connector's default-alias rule (`"Specialty - X - Y"` → `"Y"`). Fixed with progressive-suffix.
- Source format is `0%` (decimal) but the mapper aggregated wrong column or didn't filter `measure="Mean"`, producing nonsense values. Fixed with measure-column fallback.
- Tables silently never wrote because the spec didn't carry `cell_values`. Fixed with auto-compute.
- Multiple charts share the same tag config — can't auto-differentiate. Now guards instead of corrupting.

**"How and why did scatter fixes break?"**
They didn't. The April 28 commit (`859f101`) fixed None-safe rounding for crashes; that's still in place. The two scatter bugs surfaced today (value scaling for "0%" axis, series naming) were pre-existing issues not addressed by that fix. Today's commits address them.

**"Is the connected path good now?"**
Connected path handles correctly: sparse + dense decks, mixed static+dynamic on same slide, mixed-pin charts, segment-suffixed chronology, formula columns, group-nested shapes, ambiguous tags (preserved + surfaced), category/series default-alias renames + minor drift, period banners adjacent to charts, table auto-refresh.
Known unfixed: field-date stamps (`Q1'26: 01/01/2026 – 25/02/2026`), 1×1 templated cells with non-`(n = X)` patterns, scatter charts where the connector tag's selectedColumns refers to bare value columns without row labels.

### Pending / deferred from today

- **Field-date stamps** — `Q1'26: 01/01/2026 – 25/02/2026` style. Label-sync rewrites the wave token but not the date range. Needs `time_period_start`/`time_period_end` in the shift map.
- **Slide 73 ambiguous-tag root cause** is on the user side. Each of the 4 charts needs its own filter (e.g. `intent_type = "increase"`) or `split_order` to differentiate.
- **Templated cells beyond `(n = X)`** — `"Sample size: 120"`, `"based on 51 respondents"`, etc. need explicit pattern handling.
- **Badge wording for `alignment_failed`** — phrase is technically correct but jargon-heavy. Could be clearer.

### Files created today

- `slidegen/label_sync.py` (prior day, but referenced heavily today)
- `slidegen/refresh_pipeline.py` (prior day, extended today)
- `CONNECTOR_PATH_README.md` (today)
- `CONNECTOR_PATH_MEMORY.md` (this file)

### Late-day fixes (after the v10 review pass)

User asked: can we close all known limitations? Worked through them:

- **Templated 1×1 cells beyond `(n = X)`** — extended `_COUNT_PATTERNS`
  to a list of 7 universal patterns: `(n = X)`, `(N = X)`, `n=X`,
  `Sample size: X`, `Sample Size = X`, `Base: X`, `Total: X`,
  `based on X respondents`, `X HCPs/patients/subjects`. All 12 unit
  tests pass. New patterns can be added in seconds.
- **Scatter series naming fallback** — when mapper produces a purely-
  numeric series name (rare edge case where selectedColumns picks
  bare value columns), writer falls back to: source's series name at
  same index → `columnDefinitions[].Alias` / first non-`<blank:>`
  Name. Universal across decks.
- **Positional alignment 5th-tier fallback** — when categories have
  same count source vs API and the 4 named-match tiers (exact,
  suffix, substring, fuzzy) all miss, fall back to positional
  alignment (source[i] → pivot[i]). Reading values, not consuming
  pivot rows, so multiple sources can resolve to the same pivot
  index without conflict. Reduces alignment_failed cases substantially.
- **Field-date stamps** — investigated, blocked Synapse-side. The
  API records don't include `time_period_start` / `time_period_end`
  columns, so we have no source of truth for the new date range
  when a wave shifts. Documented as a Synapse-API-side requirement.
- **Ambiguous tag** — confirmed user-side. The 4 charts on slide 73
  all carry byte-identical connector tags; the pipeline can't
  auto-disambiguate without risk of wrong assignment. Source values
  preserved + flagged. Action: add per-chart filter or split_order.

### Doc updates (late day)

- README expanded with Prerequisites table + TL;DR walkthrough so a
  fresh reader (Vijay or anyone else) can clone, install, configure
  creds, and refresh without prior context.
- Added an honest "Will refresh this deck just work?" section that
  separates "fully auto-handled" (~90% of slides) from "flagged +
  preserved" (the 5 review buckets) from "not auto-handled" (field-
  date stamps, exotic templated cells, ambiguous tags).
- Known Limitations rewritten to mark which were closed today vs which
  remain (with reasons).

### Late-day commits

| Commit | What |
|---|---|
| `dd69544` | README: Prerequisites + 'will it just work?' section. |
| `e5bc4d6` | Templated patterns extended + scatter series-name fallback + positional alignment 5th tier + README/memory updates. |
| (next)   | Connector unit-test suite (4 files, 55 tests) + README evals section rewrite. |

### Slide-wide period sync (intelligence layer added end-of-day)

User asked: when chart cats shift Oct'25..Mar'26 -> Nov'25..Apr'26,
the slide header / chart header / footnote prose still say the old
window. Can Claude have an intelligent layer that rewrites period
tokens *anywhere* on the slide that references the old wave window?

Extended `slidegen/label_sync.py` with a Pass 2 (slide-wide). The
existing Pass 1 (adjacency-strict, label-shape guard) still runs and
remains the primary path for period-banner cells. Pass 2 then walks
every other text frame on the slide, plus chart titles + axis titles
(which live in chart XML — separate walk via `_rewrite_chart_titles`).
Sentence-level rewriter `_apply_shift_to_sentence` rewrites period
tokens in place without the label-shape filter, preserving the
surrounding sentence intact.

Key design decisions:
  - Two passes, not one. Pass 1 stays strict so period-banner cells
    don't get over-rewritten; Pass 2 is permissive for prose.
  - Pass 2 skips runs already edited by Pass 1 (tracked via
    (id(text_frame), para_idx, run_idx) keys) so we never
    double-rewrite.
  - Single-pass regex alternation for the shift map — no chained
    substitution where Oct -> Nov -> Dec.
  - Curly quotes normalized at match time (' / ' / `).
  - `slide_wide=True` is the default; pass False to revert to
    strict-adjacency-only.

9 unit tests in tests/connector/test_label_sync_slide_wide.py cover
sentence-level rewriter behavior, single-pass guarantee, curly-quote
match, partial shift maps, empty-input handling.

Smoke-tested on Repatha ATU v10: Pass 2 catches `Q1'26 -> Wave 7`
substitutions in tables that the adjacency pass missed (e.g. Table
27/28 on slides 73/75 column headers).

### Per-deck quality eval (added late-late-day)

User pushed back: unit tests are great but they don't tell you "did
this specific refresh of this specific deck pass." Built
`slidegen/refresh_eval.py` — reads the `*_refresh_status.json` sidecar
and outputs a verdict: passed / preserved / review / error counts +
pass-rate %. Wired into the pipeline so every refresh prints this
verdict at the end. Also runnable standalone:

    python -m slidegen.refresh_eval <status.json>
    python -m slidegen.refresh_eval <deck.pptx>           # auto-locates sidecar
    python -m slidegen.refresh_eval <status> --threshold 80  # CI gate

Note-level downgrade — a component with `status=ok` but carrying a
`tag_mismatch` or `selectedColumns_drift` note is classified REVIEW,
not PASSED. Without that, Repatha ATU v10 would have inflated to
~85%; the honest score is 74.5% PASS / 13.7% REVIEW / 9.5% ERROR.

11 unit tests in `tests/connector/test_refresh_eval.py` cover the
classification taxonomy, note downgrade, rate calculations, empty
data, and the format_report renderer.

### Eval coverage closed today

User asked: are evals extensive enough? Walked through gaps and added
4 unit-test files under `tests/connector/`:

| File | Tests | Covers |
|---|---|---|
| `test_formula_evaluator.py` | 21 | Tokenizer + evaluator + full `_build_formula_columns` with default-alias renames. |
| `test_resolution_chain.py` | 7 | 5-tier resolver (exact, suffix, substring, fuzzy, positional) for cats + series. |
| `test_table_auto_write.py` | 21 | `_format_cell_value`, `_refresh_templated_count_with_records` (8 patterns), `_auto_compute_cell_values` (4 layouts). |
| `test_filter_resolver.py` | 6 | `;`-joined Filters with compound-suffix, broken-filter fallback, measure filter, latest-wave-only. |

All 55 pass in ~1 sec. The table-auto-write suite caught a real bug
along the way: `_format_cell_value("0.0%")` was producing 0 decimals
instead of 1 (the format detection code took the wrong branch). Fixed
in the same commit. README's Evals section rewritten to point at
this new structure plus the existing end-to-end + live-API tests.

### Validation deck on disk

- `output_testing/deck_output/Repatha_HCP_ATU_Q2_26_Skeleton_Deck1_2026-05-06_v10.pptx` — last full validation run with table auto-write + headline detector + scatter scaling fix; user instructed not to run v11 today, has other tasks.

---

<!--
## YYYY-MM-DD (template for next session)

**Context.** ...

### Issues raised + status

| Slide | Issue | Root cause | Status |

### Commits

### Decisions

### Pending
-->
