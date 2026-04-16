---
name: deck-audit-workflow
effort: medium
paths: []
description: "Triggered by: 'Review this deck before delivery' / 'Audit the deck for issues' / 'Find data inconsistencies in the deck' / 'Check every claim has a citation'. Produces a pre-delivery audit report flagging data errors, unsupported claims, stale lineage, missing citations, and structural inconsistencies. Does NOT modify the deck — read-only workflow that produces a review document."
---

# deck-audit-workflow

Pre-delivery QA audit. Read-only. Produces a report of issues to fix; does not mutate the deck.

## Trigger phrases

- "Review this deck before client delivery"
- "Audit the Rybrevant Q1 deck for issues"
- "Find data inconsistencies"
- "Check every claim has a citation on the ES slide"
- "Run a pre-delivery QA on this deck"

## Cardinal Rules

1. **Read-only.** Never modifies the input deck. Audit output is a separate markdown report.
2. **Every finding is actionable.** Each flagged issue points to a specific slide + shape + recommended fix.
3. **Severity levels are explicit.** `BLOCKER` (wrong numbers, missing citations on ES) / `WARN` (stale data, n<30 claims) / `NIT` (stylistic inconsistencies).
4. **Cross-slide coherence matters.** If slide 5 says "up 8pp" and slide 12 cites the same metric as "up 6pp", flag as inconsistent.
5. **Audit scope is the SPEC, not visual rendering.** Can't catch "this chart looks ugly" — catches structural/data/narrative issues.

## Audit checks (the full set)

**Data integrity:**
- Every chart series has `len(values) == len(categories)` (also validator-enforced, but double-check at deck level)
- Delta column values = current - prior per row (if both series present)
- No NaN / None / "—" in cells that should have data
- Percentages are in `[0, 1]` range (not 45 vs 0.45 mixed)

**Lineage / audit trail:**
- Every data-driven slide has `data_lineage.last_data_pull` within staleness threshold (default: 7 days)
- Tier 1 slides (Connector-authored) have valid `reporting_plan_id + analysis_ids`
- Tier 2 slides (inferred) have `confidence != "low"` OR a user-override note
- `last_refresh_error` is empty on all slides (non-empty = last refresh failed)

**Citation coverage (for ES slides):**
- Every ES bullet has `metadata.citations: list[slide_id]` with ≥1 entry
- Every cited `slide_id` exists in the deck
- No "open" claims without citation

**Narrative coherence:**
- Every slide tagged to an arc in metadata (if narrative_threads.md exists)
- No two slides claim opposite deltas on the same metric
- Headline numbers match the slide's data labels

**Sample size / statistical sanity:**
- If a slide cites a segment, its n ≥ 30 (else flag with low-base warning)
- Any delta claim of <2pp is flagged as "within margin of error" (soft warn)

**Structural:**
- Spec validates cleanly (`validate_spec` per slide)
- No duplicate `slide_id`s
- No overlapping component positions (warn; rendering still works)
- Brand resolution succeeds (every `{brand.X}` token resolves)

**Chrome / standards:**
- Every slide has a headline + footer
- Cover slide + ES slide + Recs slide present at expected positions
- Page numbers consistent (if used)

## Inputs

- `deck_path`: the deck to audit
- Optional `staleness_threshold_hours`: default 168 (7 days) for "last_data_pull" warnings
- Optional `checks`: subset of check categories to run (default: all)
- Optional `severity_filter`: only report findings at or above this level

## Outputs

- `<deck_path>.audit_<timestamp>.md` — the audit report (in same directory as deck)
- Exit code: 0 if no BLOCKERs, 1 if BLOCKERs present (for CI integration)
- Terminal summary: counts per severity

## Orchestration

```
1. deck-reader(deck_path, all_slides=True)
     → list[SlideSpec]
     → deck-level metadata (counts, brands, period)

2. Run each check category on the deck:
   - data_integrity_checks(specs)
   - lineage_checks(specs, staleness_threshold_hours)
   - citation_checks(specs)  # focuses on ES slides
   - narrative_coherence_checks(specs, narrative_threads_md if present)
   - sample_size_checks(specs)
   - structural_checks(specs)
   - chrome_checks(specs)
   — each returns list[Finding(severity, slide_id, shape, description, fix_suggestion)]

3. Aggregate findings into a report:
   - Summary: n_blockers / n_warns / n_nits
   - By slide: expandable section per slide
   - By severity: grouped lists
   - Recommended fix actions with links to relevant workflows (edit-slide, refresh-deck, etc.)

4. Write report to <deck_path>.audit_<timestamp>.md

5. Print terminal summary + exit with code 0 or 1
```

## Decision Rules

| Situation | Response |
|---|---|
| Deck has 0 slides | Report "empty deck — nothing to audit"; exit 0 |
| Deck is untagged (all Tier 2) | Lineage checks are weaker; warn but don't block |
| `narrative_threads.md` missing | Skip narrative coherence checks; note skip in report |
| Data is all 0 or all None | Flag as BLOCKER (likely broken extraction) |
| BLOCKER found in a cover / divider slide | Weird — usually these are structural-only; investigate |

## Integration

- **Recommended post-deck generation step** for `create-deck-workflow` and `refresh-deck-workflow`: run audit before marking the deck ready for delivery
- Can be run standalone for any existing deck (including decks SlideGen didn't produce — as long as deck-reader Tier 2 can parse them)

## References

- PRD §3 Workflow 7, §5 composition map
- `.claude/skills/context-data/deck-reader/SKILL.md`
- `.claude/skills/planning/spec-validator/SKILL.md`
- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` (PET) / `atu-insight-writer/SKILL.md` (ATU) (narrative backbone reference)
- `slidegen/pptx_utils/registry.py` (shape registry for audit trail)
