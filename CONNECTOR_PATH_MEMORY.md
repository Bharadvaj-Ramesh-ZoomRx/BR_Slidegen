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

### Testing-Deck assembly (open at end of day; resume next session)

**The ask.** Build a single deck named `Testing Deck.pptx` (in
`output_testing/deck_output/`) by pulling **only the dynamic
connected slides** from each of 9 source decks. Static-config slides
should NOT be included — the goal is a sample deck where every slide
is genuinely refreshable, so user can hit "refresh" and see the full
pipeline exercised.

**Refinement made just before session ended:** original ask was "all
connected slides per deck" (~369 slides total — too big). Refined to
"dynamic slides only" — exclude truly-static-pinned components.

**Definition of dynamic (apply per slide).** A slide qualifies as
"dynamic connected" if AT LEAST ONE of its connected components has
either:
  - `dynamic_latest_n > 0`, OR
  - `include_live_wave == True`

A component that has `static_time_period_ids` set AND no
`dynamic_latest_n` AND `include_live_wave == False` is truly static
and is the kind that gets `static_pinned_skipped` at refresh. Slides
where ALL connected components are truly static should be EXCLUDED
from the Testing Deck.

A slide with mixed pin (static + dynamic_latest_n) DOES refresh, so
it stays IN.

**The 9 source decks (confirmed by user):**

  1. AML & MDS PET Q2FY26 Full Report 27MAR2026 - sandbox migration.pptx
  2. AVEO Wave 5 PET Report v1.0.pptx
  3. AZN LOKELMA PET Quarterly Report Q1 2026.pptx
  4. CREON Share of Voice Study - W33 updated source deck.pptx
  5. DATROWAY EGFRm NSCLC Promotional Effectiveness Tracking (PET)
     Q1 '26 PP and NPP Final Report_v1 (4).pptx
  6. Repatha HCP ATU - Q2'26 Skeleton Deck1 (1).pptx
  7. [ZoomRx] Abilify LAI PET - Full Report (1).pptx
  8. [ZoomRx] ILAI Q1 '26 - Topline Report.pptx
  9. projects/J&J Rybrevant PET/Template/ZoomRx_UC_ATU_Report_Q1_'26.pptx

Connected-slide counts BEFORE the dynamic-only filter (from
`generate_config_specs`):

  AML & MDS                88 total / 53 connected
  AVEO Wave 5              58 total / 36 connected
  AZN LOKELMA              67 total / 33 connected
  CREON W33               104 total / 47 connected
  DATROWAY EGFRm           53 total / 30 connected
  Repatha HCP ATU         109 total / 69 connected
  Abilify LAI              54 total / 32 connected
  ILAI Q1'26               33 total / 13 connected
  UC ATU                   99 total / 56 connected
                                ----
                                369 total connected (pre-dynamic-filter)

The dynamic-only filter is expected to drop a meaningful chunk —
e.g. on Repatha ATU, slides 11/12/etc are mixed-pin and stay; pure
static-history slides drop. Final Testing Deck size likely 200-300
slides.

**Cleanup already done.** Eighteen result decks (CREON v2-v8,
Repatha ATU v1-v10 minus v4, AVEO v1) plus the PowerPoint lock file
were deleted from `output_testing/deck_output/` per user request.
`output_testing/deck_output/` now contains ONLY the 9 source decks
listed above (the J&J one lives in `projects/...`).

**Approach drafted but not executed.** `scripts/build_testing_deck.py`
exists in the working tree (uncommitted) using PowerPoint COM via
pywin32 to copy slides cross-deck. This preserves connector tags +
embedded charts + relationships exactly as a manual PowerPoint
copy/paste would. Was about to run, then user paused for the
dynamic-filter clarification AND to switch sessions.

`pywin32==311` was installed today (`pip install pywin32`), so the
new session has it ready. PowerPoint 16.0 detected on this machine.

**What the next session needs to do.**

  1. Update `scripts/build_testing_deck.py` so the per-deck index
     filter applies the dynamic-only rule (above) instead of the
     blanket `tagged > 0` rule. Helper logic to add:

         from slidegen.deck_reader.tag_reader import generate_config_specs
         from slidegen.slide_spec.schema import dump_spec
         specs, summary = generate_config_specs(str(deck_path))
         dynamic_indices_1based = []
         for spec in specs:
             d = json.loads(dump_spec(spec))
             si = d["slide_index"]
             dynamic = False
             for comp in d.get("components", []):
                 if comp.get("type") not in ("chart", "value_table", "label_table"):
                     continue
                 dm = comp.get("data_mapping", {}) or {}
                 lin = dm.get("raw_data_lineage", {}) or {}
                 if (lin.get("dynamic_latest_n") or 0) > 0:
                     dynamic = True; break
                 if lin.get("include_live_wave"):
                     dynamic = True; break
             if dynamic:
                 dynamic_indices_1based.append(si + 1)

  2. Run the script. Expect ~5-10 minutes runtime (COM open/copy/paste
     across 9 decks). PowerPoint must be installable + able to launch
     headless on the machine; user already had a Repatha lock file
     hanging around from a prior session — close any open decks before
     running.

  3. Verify Testing Deck.pptx exists with the expected slide count.
     Print per-deck breakdown.

  4. Optional follow-up: run `python -m slidegen.refresh_pipeline
     "output_testing/deck_output/Testing Deck.pptx"` to refresh and
     produce the eval verdict.

**Memory notes for the new session:**
  - Do NOT touch the J&J Rybrevant Template folder's content; the
    UC ATU deck there is a source of truth.
  - The pywin32 + PowerPoint COM approach is necessary because
    python-pptx alone can't reliably cross-deck-copy shapes that
    have connector custom XML parts + embedded chart OLE objects.
  - All commits so far today are pushed to `bharadvaj-slidegen`.
    Latest HEAD = `e3b2a1c`. Working tree has uncommitted file
    `scripts/build_testing_deck.py` (this session's draft).

<!--
## YYYY-MM-DD (template for next session)

**Context.** ...

### Issues raised + status

| Slide | Issue | Root cause | Status |

### Commits

### Decisions

### Pending
-->
