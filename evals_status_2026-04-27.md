# Refresh quality + evals — status for Vijay (2026-04-27)

**Branch:** `bharadvaj-slidegen` (tip `c2ccc44`).
**Two reference decks under test:** ATU Q1'26 (197 charts + 209 tables) and CREON PET W33 (403 charts + 163 tables).

---

## What this is

I'm walking the refresh pipeline through a 7-step framework — each step proves a different property the product needs. For every step there's an eval that runs in CI-style, plus a real number we can quote to clients. This note is the snapshot of where each step stands today.

---

## The 7 steps at a glance

| # | Step | What it proves | Status | Number you can quote |
|---|---|---|---|---|
| 1 | Read a Connector-tagged deck → produce a SlideSpec | We can faithfully read what the deck is asking for | ✅ Done | **55 ATU + 42 CREON connected slides** locked in golden (non-connected slides excluded — covered by Step 1b separately) |
| 1b | Same thing but for non-connected (untagged) decks | Same property, no tags — needs inference | 🟡 Code path shipped (Pradeep v0); **no fixture deck registered**, no accuracy eval | 0 (no eval written yet) |
| 2 | Refresh against the SAME wave the deck was rendered with | Refresh is a clean roundtrip on day-one data | 🟡 Mostly working, failures remain | **ATU 157/305 connected components match values (51%). CREON 423/464 (91%).** Non-connected slides excluded from score. |
| 3 | Visual formatting survives refresh | Charts still look right after refresh | ✅ Done | chart_type 100% / 100%. Colors 87% ATU, 98% CREON |
| 4 | Slide headlines untouched on roundtrip | We don't accidentally rewrite text | ✅ Done (implicit — refresh has no text path) | n/a |
| 5 | Refresh against a DIFFERENT wave for a single chart | New data flows in correctly for one well-behaved chart | ✅ Done | 1 reference chart per deck passes 3 assertions |
| 6 | Refresh the WHOLE deck against a different wave window | The April-end claim — deck rolls forward without anything breaking | 🟡 Eval written, **failing** today; fixes in flight right now | ATU: 40 refreshed / 51 drift. CREON: 28 / 30. Phase 1+2 fixes will improve sharply |
| 7 | Headlines auto-update with new data | Client reads "share grew to 67%" not "65%" after refresh | ❌ Not built | 0 |

---

## Step 2 — where the remaining failures actually live

This is the most-watched number because it's the same-wave roundtrip — it's saying "if we refresh today's deck against today's data, do we get today's deck back?" Today, **connected slides only**: **ATU 157/305 = 51%, CREON 423/464 = 91%** of components matching values. (Non-connected slides are now excluded from the score — they were never refreshed and trivially inflated the count.) The gap is what we're chasing.

The remaining failures sort into 5 buckets. Three of them need **Connector-side or source-deck work** (your lane). Two are **mapper code fixes** (my lane).

### Bucket A — ~37 charts have a wrong analysis_id in the Connector tag (your lane)

The chart visually shows data from analysis X, but the tag points at analysis Y. We refresh against Y because that's what the tag says, and the result diverges from the source. This is a tag/render mismatch baked into the source decks themselves — not a refresh bug.

**Concrete:** CREON slides 71/72/73 (split-viz family). Per-shape tags claim analysis 117414; the source was rendered from slide-level analysis 654352.

**To fix:** the source deck needs a regen with corrected tags, or Connector needs a tag-validation step at render time.

### Bucket B — ~15 charts are dynamic and Eval #3 is asking the wrong question of them (eval design issue, not a bug)

These charts have `dynamic_latest_n: 2` — meaning "always show the latest 2 waves." The source was rendered when latest-2 was W11+W12. Today, latest-2 is W12+W13 (W13 has since landed). The refresh fetches W12+W13 — exactly what the tag asks for. **That's correct product behavior**, not a failure.

The "failure" is in Eval #3, which measures `source values == refreshed values`. Identity is a meaningful property only for charts with frozen `static_time_period_ids` in the tag. For dynamic charts, drifting is the feature, not the bug.

**Not a Connector fix. Not a mapper fix.** This is an eval-design refinement: split Eval #3's totals into two columns — "static charts: identity vs source" (strict) and "dynamic charts: did the refresh execute and return latest-N waves" (looser, the right question for that mode). Bharadvaj's lane.

**Open business question:** what's the default Connector behavior we want when a client hits refresh on an old deck? Frozen-at-render (deck = moment in time, must be explicitly rolled forward) or dynamic (deck always reflects latest-N, rolls forward automatically). Both are valid product modes and the schema already supports both via `static_time_period_ids` vs `dynamic_latest_n`. This is a default-policy decision, not a bug.

### Bucket C — ~8-9 charts use compound `selectedColumns` with wave labels embedded (my lane, fixing right now)

The mapper's existing fix (Fix L) handles entries like `"Feb'26"` standalone. It doesn't handle compound entries like `"Feb'26 @:@ Motivating to Prescribe @:@ CREON_ME"` — wave label fused into a multi-axis path. When the wave shifts, the compound entry no longer matches anything and the chart degrades.

**Concrete:** CREON slides 24, 25, 47.

**Fix:** Fix M, in flight today. Detects wave-prefixes in compound entries, replaces them while keeping the rest. Universal — works on any project's wave naming conventions because detection is value-driven (we ask the API for the analysis's known wave labels and treat any string in that set as a wave).

### Bucket D — "Project Wave" prefix in series names (my lane, fix ready to write)

**Updated 2026-04-27 after investigation.** The Synapse API returns series names like `"Project Wave 12"`, but the source deck shows `"Wave 12"`. Fix D strips the `"Project "` prefix in the mapper to reconcile.

**Exact slides:** 8, 9, 25, 26, 29, 30, 33, 40, 49, 50, 53, 55, 58, 60, 61, 62, 63, 64, 72, 79, 80, 81, 86 — **45 charts across 23 slides.**

This is a one-liner in the mapper's series-name output. Will fix alongside Bucket E.

### Bucket E — ~65 charts hit miscellaneous mapper edge cases (my lane, incremental)

Series count mismatches (3v4, 5v16, 11v16, etc.), value offsets that look like aggregation/unit divergence. Each pattern needs individual repro. Each fix is a small chart-count gain. Diminishing returns; I'll keep grinding through these.

### Bucket F — "Dummy Series" placeholder names (my lane, new — discovered 2026-04-27)

**New bucket, not previously captured.** 83 charts across 33 slides have `"Dummy Series 1"` / `"Dummy Series 2"` as their refreshed series names instead of real labels like `"SKYRIZI Current"` / `"RINVOQ Current"`. This happens on doughnut and stacked charts where the mapper emits a placeholder instead of resolving the actual option value from the API response.

**Exact slides:** 7, 16, 18, 19, 20, 21, 22, 26, 31, 34, 35, 36, 37, 38, 39, 45, 48, 49, 50, 57, 65, 66, 67, 77, 78, 83, 91, 94, 95, 96, 97, 98, 99 — **83 charts across 33 slides.**

This is a mapper bug in how it handles doughnut/stacked chart types — series names come from the chart template rather than being filled from the API data. Needs investigation of the pivot path for those chart types.

### Why series_match is 2/197 in Eval #2 (not just one thing)

The 195/197 mismatch is three separate problems, not one:

| Pattern | Charts | Slides | Root cause |
|---|---|---|---|
| "Project Wave" prefix (Bucket D) | 45 | 23 | Mapper not stripping API prefix |
| "Dummy Series" placeholder (Bucket F) | 83 | 33 | Mapper emitting template name for doughnut/stacked |
| Wrong series entirely (Bucket A spillover) | 67 | 30 | Wrong analysis_id in tag → completely different data |

Only 2 charts escape all three problems (slides 6 and 80, both using non-wave-label series like patient segment descriptions).

### Bucket summary

| Bucket | Charts affected | Owner | Fix type |
|---|---|---|---|
| A — Wrong analysis_id in tag | ~37 | Vijay | Connector-side / source regen |
| B — Dynamic charts mis-categorized as failures | ~15 | Bharadvaj (eval), Vijay (policy decision) | Split Eval #3 totals so dynamic charts use a "did refresh fetch latest-N?" check instead of strict identity |
| C — Compound `selectedColumns` waves | ~8-9 | Bharadvaj | **Mapper fix (Fix M, in flight today)** |
| D — "Project Wave" prefix in series names | 45 (23 slides) | Bharadvaj | One-liner in mapper series-name output |
| E — Mapper edge cases | ~65 | Bharadvaj | Incremental code fixes |
| F — "Dummy Series" placeholder names | 83 (33 slides) | Bharadvaj | Mapper pivot path fix for doughnut/stacked charts |

**Synapse-CLI dependency note:** for Steps 1, 2, 3, 4, 5, 6 (the connected path) we hit the Synapse API directly via `requests` in `fetch_synapse_data()`. `synapse-cli` is *not* in this code path. It only matters for Pradeep's non-connected raw-data extraction.

---

## Where each eval lives

| Step | Eval file | Runtime | Status |
|---|---|---|---|
| 1 | `tests/evals/spec_extraction/test_path_a.py` | ~5s | Green + mutation test green |
| 1b | not written | — | Blocked on (a) a non-connected fixture deck, (b) inference v0 stabilizing |
| 2 | `tests/evals/end_to_end/test_refresh_execution.py` | ~7s | Green vs latest golden + mutation test green (connected slides only: ATU 157/305, CREON 423/464) |
| 3 | `tests/evals/end_to_end/compare_formatting.py` | ~5s | Green |
| 4 | implicit in Step 2 — could add explicit `headline_match` cheaply | — | — |
| 5 | `tests/evals/end_to_end/test_step5_new_wave.py` | ~12 min | Green on 1 chart per deck |
| 6 | `tests/evals/end_to_end/test_step6_deckwide_wave_shift.py` | ~9 min | Failing today; Phase 1+2 fixes running |
| 7 | not written | — | Skill not built |

**Worth knowing:** every eval is parametrized over a `FIXTURE_DECKS` registry. Adding an AstraZeneca deck or any other client deck is a one-line registry entry — no test code changes. Wave shifts, analyses, and waves are all resolved at runtime via API queries; nothing about ATU or CREON is hardcoded.

---

## Step 1b — non-connected path, more precisely

I had this listed too vaguely earlier. Three things are simultaneously true:

1. **The code path is shipped.** Pradeep's `c2359fa` (`scripts/infer_nonconnected.py`) and `b0ea727` (the integrated non-connected → connected pipeline) read an untagged deck, infer the configs, and stamp Connector-format tags. From there, all of Steps 2-7 run identically. This is real.
2. **No non-connected fixture deck is registered.** ATU and CREON are both already-tagged decks; running inference on them would be redundant. We need an actual untagged client deck in `FIXTURE_DECKS` to exercise the inference end-to-end.
3. **Inference is v0.** The accuracy across diverse chart patterns is unknown. Locking a golden today means every inference improvement breaks the eval. The agreed posture: let v0 → v1 stabilize, then lock.

**What I need from you:** is there a production non-connected deck we can register as a fixture? That unblocks (2). The team's call on (3) — when is inference "stable enough" — is yours to make.

---

## What I'm working on right now

I'm in the middle of two phases that both apply universally (not specific to ATU/CREON):

**Phase 1 — eval polish (cleaning up what the test measures):**

1. **Tables were appearing as "didn't change" no matter what.** Each table in the spec carries pre-computed cell values from when the deck was last rendered. When the wave-shift test runs, it tells the pipeline "fetch different waves" — but it doesn't tell tables to recompute their cells. So tables keep their old cached values and always look frozen. With 209 ATU + 163 CREON tables, that's 372 components giving us no signal at all.
   *Fix:* when building the variant spec for the eval, wipe the cached cell values on every table. Forces the refresh to re-pivot from fresh API data.

2. **The "welded" number was lumping two very different things together.** Some charts can't be wave-shifted at all because their source already spans the oldest available wave (CREON slide 8 charts span all 6 waves of analysis 689321 — there's nothing older). My code correctly *skips* those data sources, but the eval was counting their unchanged output the same as charts that DID try to shift and silently fell back. Two different stories, one bucket.
   *Fix:* classify each chart by whether its data source actually got shifted. Report two columns: "shifted ds, didn't change" (the real welded class — the one to drive down) vs "ds couldn't be shifted" (correctly skipped — informational only).

3. **One-wave charts came out with two waves after shift.** A few ATU charts on slide 14 use `dynamic_latest_n=1` (show the single latest wave). The shift should preserve K=1 — produce 1 wave. Some are producing 2 instead, showing up as `1 → 2` cat drift. Likely an off-by-one in the slice math when K=1.
   *Fix:* walk the K=1 path in `compute_one_wave_back_shift` and tighten the boundary.

**Phase 2 — mapper polish (the actual data fix, universal):**

1. **Fix L v2 — make wave detection work for any naming convention.**
   *Today's gap:* Fix L recognizes a wave label by matching it against a hand-coded pattern list (`Wave N`, `Q1'26`, `Jan'26`, etc.). Works for ATU and CREON. Breaks the moment AstraZeneca uses `H1 2026`, `2026 Q1`, `January 2026`, or any format we haven't catalogued — those entries quietly stay welded.
   *Upgrade:* before processing a chart, the refresh pipeline asks the Synapse API for the full set of wave labels that exist for this analysis (cached, one query per analysis per run). The mapper then treats any string in `selectedColumns` that's in that set as a wave label, regardless of format. The hand-coded pattern list stays as a fallback for callers that don't supply the API set. Net result: no per-project tuning needed for new clients.

2. **Fix M — handle compound entries (Bucket C).**
   *Today's gap:* Fix L only rewrites standalone entries like `"Feb'26"`. Compound entries with a wave label fused into a longer path — `"Feb'26 @:@ Motivating to Prescribe @:@ CREON_ME"` — are left untouched. When the wave shifts away from `Feb'26`, that whole compound entry no longer matches any pivot column, the chart loses a series, and we get `series-count: 2 → 1` drift on CREON slides 24/25/47.
   *Upgrade:* the same rewrite helper now splits compound entries on the `@:@` separator, finds the wave-label part, builds a template (everything else preserved), and emits one rewritten entry per fetched wave. So `"Feb'26 @:@ X"` and `"Mar'26 @:@ X"` collapse to one template `"<wave> @:@ X"` and expand to one entry per actual fetched wave.

**Both fixes are universal** — they live in the mapper code, not the eval, so any deck that goes through the refresh pipeline benefits. No deck-specific configuration anywhere.

**Status as of 2026-04-27 end of session:**
- Phase 1+2 code is **written and unit-tested** but not yet committed. Fix L v2 + Fix M unit tests are green, including a synthetic AstraZeneca `H1 2026` case.
- Step 6 eval framework is written and running. Numbers still show drift; Phase 1+2 fixes should reduce it significantly on next full run.
- Bucket D and F were **fully diagnosed today** with exact slide lists (see above).

**Next session — in order:**
1. Commit Phase 1+2 code, push to `bharadvaj-slidegen`, update goldens.
2. Re-run Step 6 eval and record new drift numbers.
3. Fix Bucket D (one-liner: strip `"Project "` from mapper series names).
4. Investigate Bucket F (Dummy Series on doughnut/stacked charts) — understand the mapper pivot path for those chart types, then fix.
5. Forward-shift eval once Step 6 numbers are strong — use backward-shifted output decks as synthesized "older sources."
6. Step 7 (headline-writer) after Step 6 is solid.

---

## Open questions for you

1. **Bucket A (wrong analysis_id):** do we need a coordinated source-deck regen, or are these one-offs to chase per chart? If batched, when?
2. **Bucket B (dynamic vs static default):** what's the intended Connector default — should `Refresh` on a deck rendered last quarter give clients the same numbers or auto-roll-forward to the latest waves? This is a product-default question, not a code fix. Your call shapes how I split the Eval #3 success criteria.
3. **Step 1b — non-connected fixture deck:** any untagged production deck I can register?
4. **Step 7 (`headline-writer`):** ETA from your side, and would my Step 6 work be useful input?
