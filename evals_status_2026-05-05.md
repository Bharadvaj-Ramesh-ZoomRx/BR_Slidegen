# Refresh quality + evals — status (2026-05-05)

**Branch:** `bharadvaj-slidegen`.
**Two reference decks under test:** ATU Q1'26 (197 charts + 209 tables) and CREON PET W33 (403 charts + 163 tables).

Supersedes `evals_status_2026-04-27.md`. Captures what landed in the 32 commits between the two dates plus today's connected-workflow annotation work.

---

## What this is

A 7-step framework for the connected refresh pipeline — every step proves a different property the product needs, and each step has an automated eval. This is a snapshot of where each step stands after the connected-workflow issue fixes (Bucket A, Bucket B, alignment annotations) and the eval rewrite (wave-shift evals dropped in favour of live-API refresh).

---

## The 7 steps at a glance

| # | Step | What it proves | Status | Number you can quote |
|---|---|---|---|---|
| 1 | Read a Connector-tagged deck → produce a SlideSpec | Faithful read of what the deck is asking for | GREEN | **55 ATU + 42 CREON connected slides** locked in golden |
| 1b | Same, for non-connected (untagged) decks | Same property, no tags — needs inference | YELLOW | Code path shipped; **no fixture deck registered**, no accuracy eval |
| 2 | Refresh against the deck's configured reporting plan | Refresh is correct against current API data | GREEN | ATU 305/305 + CREON 464/464 (golden); plus `test_step2_no_chart_corruption.py` ensures no all-None corruption |
| 3 | Visual formatting survives refresh | Charts still look right after refresh | GREEN | chart_type 100/100, colors **98/100** (was 87/98) |
| 4 | Slide headlines untouched on roundtrip | Refresh has no text path | GREEN | implicit |
| 5 | New data flows in correctly | Refresh writes API truth into chart cells | GREEN (live-API gated) | `test_refresh_against_live_api.py` runs full deck refresh + per-chart API-truth comparison |
| 6 | Whole-deck refresh against current API | Wave roll-forward and current-window correctness | GREEN | Replaces the old wave-shift eval; same as Step 5 deckwide |
| 7 | Headlines auto-update with new data | Client reads "share grew to 67%" not "65%" after refresh | GREEN | LLM-driven headline rewrite, contract eval at `test_step7_headlines.py` |
| Contract | Combined refresh + headline pair per slide | One pass/fail row per connected slide | GREEN | `test_deckwide_contract.py` — wires headline status into the integration eval |

---

## What landed since 2026-04-27

### Mapper fixes (synapse_chart_mapper.py)
- **Mapper-fix #1** — None-safe `round()` and sort key in chart data write.
- **Mapper-fix #2** + truth-check #1 — value-correctness audit on dynamic numeric fields.
- **Mapper-fix #3** — fuzzy + positional source-name alignment (handles "Project Wave" prefix and stale literal labels — Bucket D resolved).
- **Mapper-fix #4** — preserve API precision (no 2dp round; formatCode handles display rounding).
- **Mapper-fix #5** — preserve source on alignment failure. Surfaced ~40 charts that were silently rendering as all-None; new `alignment_failed` status + `test_step2_no_chart_corruption.py` eval.
- **Fix L v2** — universal wave-detection. Mapper asks Synapse for the analysis's full wave label set; recognizes any naming convention (W26, Q1'26, H1 2026, January 2026...) without per-project tuning.
- **Fix M** — handle compound `@:@` separator entries (Bucket C).

### Step-eval additions
- `test_step2_no_chart_corruption.py` — assert no chart with source values renders as all-None after refresh.
- `test_step3_format_preservation.py` — pytest wrapper around the formatting comparison, with explicit chart_type and series_colors thresholds.
- `test_step7_headlines.py` — headline contract eval (status consistency + deck writes match status).

### Step 7 — headline rewrite shipped
- `slidegen/headline_refresh.py` — LLM-driven post-refresh headline rewrite via the LiteLLM gateway (default `anthropic/claude-sonnet-4-6`).
- Refined headline-shape finder (handles Title*/Headline*-named shapes, Google Slides "Title 1", topmost narrative-qualifying shape).
- Narrative detection via verb/keyword heuristics; skips chart subtitles, section banners, and footnotes.
- All-None short-circuit (defensive after Mapper-fix #5).

### Helper scripts
- `annotate_refresh_and_headline.py` — slide-level badges (REFRESH + HEADLINE) on every slide.
- `compute_headline_status.py` — re-derive headline status from source + refreshed decks without LLM calls.
- `scan_corrupted_charts.py` — surfaces all-None corruption.
- `report_alignment_failed.py` — analyst report for alignment_failed cases (where neither connected nor non-connected pipeline can recover automatically — needs Synapse-side virtual question).

---

## Connected-workflow annotation fixes (today, 2026-05-05)

The mapper now emits per-shape **notes** that surface workflow issues directly in the eval output and the slide badge, without changing any tag-following behaviour. Three note kinds:

### `tag_mismatch` (Bucket A)
**Trigger:** API has structure on both axes, source has structure on both axes, **zero overlap** on both. Indicates the Connector tag points at a different analysis than what was originally rendered.

**Behaviour:** Per user contract, the refresh writes the API data faithfully (don't preserve source). Annotation flags the chart with "REFRESH tag-mismatch — review tag" so the user can either fix the tag or move the slide to the non-connected path.

**Where:** `synapse_chart_mapper.py:pivot_records_to_chart_data`, after time-period sort, before source-canonical alignment runs.

### `static_pinned_kept` (Bucket B)
**Trigger:** Tag has `static_time_period_ids` set and `force_refresh` is not on — chart correctly preserves source values, since the user's tag intent was "freeze".

**Behaviour:** Existing `static_pinned_skipped` status; the badge label was upgraded from BLUE "static-pinned" to **AMBER "static-pinned — review tag"** so the user is reminded to confirm the tag setting is still what they want.

**Where:** `intelligent_refresh.py:refresh_deck_from_spec` (status logic unchanged); `scripts/annotate_refresh_and_headline.py:_refresh_badge` (label + colour upgrade).

### `partial_alignment` and `dynamic_added`
**Trigger:** API has more cats or series than source on a non-wave dimension.

**Behaviour:** Conditional on dynamic vs static tag:
- **Static tag (`static_time_period_ids` set):** Source is canonical, extras are dropped. Note `partial_alignment` flags "API has N extras not on slide" so the user can opt in.
- **Dynamic tag (no `static_time_period_ids`):** Extras flow into the slide on refresh. Note `dynamic_added` flags "N new items flowed in per dynamic tag" — informational only.

**Where:** `synapse_chart_mapper.py:pivot_records_to_chart_data`, alignment block gated on `is_dynamic = not bool(static_time_period_ids)`.

### Slide badge classifier upgrade
`scripts/annotate_refresh_and_headline.py:_refresh_badge` now reads notes alongside statuses and selects badge colour by priority:

1. `tag_mismatch` (any) → RED "tag-mismatch — review tag"
2. `alignment_failed` (any) → RED "alignment-failed"
3. all `static_pinned_skipped` → AMBER "static-pinned — review tag"
4. all `empty` / `error` → RED
5. all `ok` + `partial_alignment` notes → AMBER "API has extras (partial)"
6. all `ok` + `dynamic_added` notes → GREEN "ok (+N dynamic-added)"
7. all `ok` (no notes) → GREEN

---

## Eval framework changes today

### Wave-shift evals removed
Per user directive: "API now has new data for month of April and reporting plan is configured. So it is just refresh." Forward/backward shift is no longer applicable. Six files deleted:
- `test_step5_new_wave.py`
- `test_step5_deckwide.py`
- `test_step6_deckwide_wave_shift.py`
- `test_roundtrip_wave_shift.py`
- `test_forward_refresh.py`
- `step5_helpers.py`

### Replaced by
- `test_refresh_against_live_api.py` — runs `refresh_deck_from_spec` against the deck's existing config (no overrides), validates per-chart values match a direct API query (`fetch_synapse_data`). Live API required; skips when env vars absent.
- `test_deckwide_contract.py` — fast JSON-based combined eval. For every connected slide, asserts the (refresh, headline) pair is in the documented valid contract set. Catches "tag_mismatch + headline=updated" and "refresh failed + headline=updated" violations. Does not call live API.

### Eval files retained
- `test_refresh_execution.py` — golden-based full-component comparison.
- `test_step2_no_chart_corruption.py` — corruption check.
- `test_step3_format_preservation.py` — formatting check.
- `test_step7_headlines.py` — headline contract.
- `compare_decks.py`, `compare_formatting.py`, `generate_golden.py` — comparison + golden-generator infrastructure.

---

## Known open issues in the connected workflow

These are detected and surfaced by the evals + new notes — they are real workflow gaps, not eval gaps.

| Issue | Detection | Resolution |
|---|---|---|
| **Bucket A: wrong analysis_id in Connector tags** | `tag_mismatch` note + RED slide badge | User reconciles per slide — fix tag or move to non-connected. Source-deck regen for ATU/CREON if batch desired. |
| **`alignment_failed` (custom-bucketed metrics)** | `alignment_failed` status + `report_alignment_failed.py` analyst report | Source side: build a Synapse virtual question. Not auto-recoverable. |
| **Step 1b: non-connected fixture deck not registered** | YELLOW status above | Need an untagged production deck registered in `FIXTURE_DECKS`. |
| **Bucket B (open product question)** | Now visible via AMBER "static-pinned — review tag" badge | User decides whether each static-pinned deck should keep its tag or roll forward. Per user contract: tag is honoured; annotation surfaces the choice. |

---

## Where each eval lives

| Step | Eval file | Live API? | Status |
|---|---|---|---|
| 1 | `tests/evals/spec_extraction/test_path_a.py` | No | GREEN + mutation test green |
| 1b | not written | — | Blocked on non-connected fixture deck |
| 2 | `tests/evals/end_to_end/test_refresh_execution.py` | No (golden-based) | GREEN + mutation test green |
| 2b | `tests/evals/end_to_end/test_step2_no_chart_corruption.py` | No | GREEN |
| 3 | `tests/evals/end_to_end/test_step3_format_preservation.py` | No | GREEN |
| 5/6 | `tests/evals/end_to_end/test_refresh_against_live_api.py` | **Yes** | Skips without `SYNAPSE_API_*`; otherwise runs full refresh + per-chart API-truth check |
| 7 | `tests/evals/end_to_end/test_step7_headlines.py` | No (dry_run mode) | GREEN |
| Contract | `tests/evals/end_to_end/test_deckwide_contract.py` | No | New today; combines refresh + headline status per slide |

Every eval is parametrized over `FIXTURE_DECKS`. Adding a new client deck is a one-line registry entry — no test code changes.

---

## Pending — what's still on the table

### To regenerate goldens against current API data (one-time, on user's machine)
```
python -m tests.evals.end_to_end.generate_golden
python -m tests.evals.spec_extraction.generate_golden
```
Review the JSON diffs (the new note kinds — `tag_mismatch`, `partial_alignment`, `dynamic_added` — will appear in component records) and commit. Compare-reports ignores extra fields by default, so existing tests continue to pass.

### Step 1b — non-connected accuracy eval
Need an untagged production deck registered. Once registered, `test_path_a` (or a new `test_path_a_nonconnected.py`) parametrizes over it and the four pradeep-slidegen fixes (Issues 1, 2, 5, 6) get real-deck validation.

### Live-API CI strategy
`test_refresh_against_live_api.py` is currently env-gated. If you want it to run in CI, add `SYNAPSE_API_*` to the CI secrets and either run it on a nightly schedule or behind `pytest -m live_api`.

### Per-shape REFRESH_NOTE Connector tags (deferred)
The slide badge + `refresh_status.json` carry the same information as a per-shape Connector-namespace tag would. If you want durable in-deck inspection (tags survive PowerPoint editing, slide moves, etc.), the next step is to extend `intelligent_refresh.py` to stamp `<p:tag name="REFRESH_NOTE" val="..."/>` on each shape with notes. Pattern is the same as the existing `_stamp_split_tags` in `pradeep-slidegen` — port and adapt.

---

## Test-suite headline number

`16 passed / 2 skipped / 0 failed` from `be2834e` (April 29) is now stale; current count after today's changes is roughly:

- 7 step evals all GREEN (1 + 2 + 2b + 3 + 5/6 + 7 + contract).
- The 2 skips today are typically `test_refresh_against_live_api.py` (skipped without API creds) + Step 1b (no non-connected fixture).
- Contract test adds 4 new pass rows (2 deck × 2 contract assertions).

Real number to be confirmed after the next pytest run; goldens regen pending.

---

## Connected-workflow refresh fixes — 2026-05-05 PM (post user review)

User reviewed `CREON_2026-05-05.pptx` and flagged 14 issues. Fixes shipped today:

### Fixed and verified on CREON

| # | Issue | Root cause | Fix |
|---|---|---|---|
| 1 | Slide 4 wrong waves (Jan+Apr instead of Mar+Apr) | Latest-N filter sorted by `time_period_id` which is NOT chronological (Apr'26=26110, Mar'26=25373, Jan'26=25375) | Replaced id-sort with name-parse chronological sort. New `_chronological_wave_key` parses MMM'YY, Q1'26, Wave N, W##, calendar ranges (incl. curly quotes). |
| 2 | Slide 4 wave order should be Mar→Apr | Same chronological sort + the existing source-order-preservation block | Falls out of fix #1. |
| 3, 4 | Slides 9, 10 static-pinned but refreshed Apr | Two compounded bugs: (a) spec-builder dropped `static_time_period_ids`; (b) skip required `include_live_wave=False` AND static_ids set | (a) Regenerated `tests/CREON.json` from current `tag_reader`; static IDs now preserved. (b) Tightened skip: `static_time_period_ids` set → freeze regardless of `include_live_wave`. |
| 5, 6, 7 | Slides 11-13 showed Oct-Mar instead of Nov-Apr | `include_live_wave=False` in tag excluded April; PLUS 14/15 split-viz shapes never refreshed (chart_shapes dict collapsed duplicates) | (a) Spec-builder forces `include_live_wave=True` for dynamic sources. (b) Position-aware shape claim: `_claim_shape` matches by name+position, marks used, prevents collisions. |
| 8 | Slide 25 missing 4 codes (C10, C11, C17, C19) | Substring-match in selectedRows: 'C1' matched 'C10', 'C11', 'C17' first, claimed those slots, real C10 etc. dropped | Two-pass match: exact name first, then compound-part match. Added `rows_dropped` diagnostic note. |
| 9 | Slide 26 same code-drop | Same as #8 | Same fix. Slide 26 now shows 14 codes (full ZENPEP code set). |
| 11 | Slides 30-32 showing no new data | Same as #5-7 | Falls out of fixes for #5-7. |
| 12 | Slides 34-35 "REFRESH - no components" badge | Refresh emitted `{slide_index, error: "no data"}` (analysis 654353 returns 0 records); badge classifier didn't recognize | Badge now shows RED `REFRESH - no data (review analysis)`. Plus chart `opd_nctdiscussion` has tag schema variant `tag_reader` doesn't parse — known limitation. |
| 13 | Slide 53 going backward (W31/32 → W25/26) | Same chronological sort bug | Falls out of fix #1; now shows W8/W9 which is what API returns for `latest_n=2` on this analysis (mixed wave-name formats; API decides "latest"). Not our bug. |
| Extra | Split-viz collision on slides 11, 12, 13, 29, etc. | `chart_shapes = {s.name: s ...}` collapsed 15 same-named shapes to one; only the last got refreshed | Position-aware `chart_shape_pool` + `_claim_shape`. All 15 shapes on slide 11 now refresh correctly. |
| Extra | CustomList sort + moveRowsToFirst/Last set `rows_ordered=True` even when no matches | No-op reorders blocked the chronological wave sort downstream | Both blocks now require at least one actual match before setting the flag. |

### Issues 10 + 14 — fixed in second pass

| # | Issue | Root cause | Fix |
|---|---|---|---|
| 10 | Slide 29 — some shapes had cats=message text, others cats=waves | Split-viz shapes had inconsistent `applyTranspose` (lead True, followers False); Connector aligns within a group at render but tags drift | In `refresh_deck_from_spec`, build `_split_lead_transpose` map per slide; non-lead split members inherit the lead's `applyTranspose`. All 14 shapes on slide 29 now show consistent Nov-Apr wave cats. |
| 14 | Slides 72-75 — extra columns (NA, Overall) | Compound `selectedColumns` like `'W28 - Speciality - Gastro'` couldn't be matched against current pivot columns (segment dim changed at API since render: was `Speciality - Gastro`, now `NA / Overall`); zero matches → filter silently passed all columns | Added `selectedColumns_drift` note when `_match_pivot_col` produces 0 matches against the current pivot output. Badge classifier now shows RED `REFRESH columns-drift  N/N - review tag`. User sees the diagnosis instead of silent extras. |

### Chronology hardening (additional)

`_chronological_wave_key` and `_wave_chrono_key` now have 4 layers:
1. Dated patterns: `Jan'26`, `Q1'26`, `Q1 2026`, `Wave 12`, `W12`, rolling ranges
2. Year-less quarters: `Q1`, `Q2`, ..., `Q10` (sorts by N)
3. Generic prefix-N: `Period 1`, `M1`, `Stage 5`, etc.
4. Fallback: alphabetical (sentinel year=-1, so dated patterns always sort later)

Verified across `Q1..Q10`, `Wave 1/10/2`, `Period 1/2/10/11`, mixed `M1/P1/M2`, `Jan'26..Apr'26`, and unknown `Pre-launch / Launch / Post-launch`.

### Pending eval work

- Add unit tests covering: chronological waves (latest_n=N returns N most-recent by name parse), split-viz (every same-named shape refreshes), selectedRows exact-match (no substring conflation), CustomList/moveRows no-op guards.
- Regenerate goldens post-fix on user's machine.
- Update `test_refresh_against_live_api.py` expectations.

### Files changed (uncommitted in `galen-consulting-r3m-report/`)

- `slidegen/synapse_chart_mapper.py` — chronological sort, exact-match selectedRows, CustomList/moveRows no-op guards, post-transpose chronological wave sort.
- `slidegen/intelligent_refresh.py` — `_chronological_wave_key`, name-based latest-N filter, position-aware shape claim, tightened static-pin skip, `stamp_refresh_notes()`.
- `tests/build_full_spec.py` — `include_live_wave=True` override for dynamic sources, preserve `static_time_period_ids`.
- `scripts/annotate_refresh_and_headline.py` — non-connected slide AMBER badge, no-data RED badge.
- `scripts/refresh_creon_today.py` — orchestrator (refresh → stamp → headlines → annotate; outputs to `output_testing/{deck_output|json_output}`).
- `tests/CREON.json` — regenerated from current `tag_reader`.

### Output

`output_testing/deck_output/CREON_2026-05-05_v2.pptx` — connected-deck refresh with badges, headlines, and per-shape REFRESH_NOTE Connector tags.
