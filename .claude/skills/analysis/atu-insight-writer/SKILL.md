---
name: atu-insight-writer
effort: high
paths: []
description: "ATU-specific hypothesis validation + narrative arc synthesis. Parallel to sfea-insight-writer but with ATU methodology: funnel leakage arcs (awareness-to-trial conversion), competitive share shift arcs, loyalty erosion patterns, barrier cluster analysis, and segment adoption curves. Produces validated_analysis.md + narrative_threads.md with ATU-appropriate framing. Trigger when: atu-deck project skill needs validated findings + narrative threads for an ATU study."
---

# atu-insight-writer

ATU-specific hypothesis validation + narrative synthesis. The ATU equivalent of `sfea-insight-writer` (which is PET/SFEA-specific).

## Why a separate skill from sfea-insight-writer

SFEA (Sales Force Effectiveness Analysis) and ATU (Awareness, Trial, Usage) are different analytical methodologies with different:
- **Hypothesis vocabularies** — SFEA hypothesizes about message recall, rep quality, prescription intent; ATU hypothesizes about funnel conversion, brand positioning, trial barriers, loyalty drivers
- **Narrative arc patterns** — SFEA arcs are about message-driven behavior change; ATU arcs are about market-position shifts and funnel dynamics
- **Headline framing** — SFEA headlines are delta-pp-driven ("Efficacy recall dipped 3pp QoQ"); ATU headlines are share-driven ("Rybrevant unaided awareness reaches 62%, closing the gap with Tagrisso")
- **Recommendation style** — SFEA recs target rep strategy; ATU recs target market access + brand positioning

Reusing sfea-insight-writer for ATU produces PET-flavored narratives for funnel data — wrong framing.

## Cardinal Rules

1. **Two phases, same as SFEA.** Phase 0 validates every hypothesis against data. Phase 1 synthesizes arcs + headlines + ES + recommendations. Both phases run before slide-plan-generator takes over.
2. **Funnel logic is the backbone.** Every ATU narrative maps to a position in the brand funnel (awareness → consideration → trial → usage → loyalty → advocacy). Findings that don't map to the funnel still get analyzed but are tagged as "cross-funnel" or "methodology."
3. **Share language, not delta language.** ATU speaks in shares ("62% unaided awareness", "Rybrevant holds 30% share of new Rx") rather than percentage-point deltas. Deltas are secondary framing, not primary.
4. **Competitive framing is mandatory.** Every ATU finding is positioned against the competitive set. "Rybrevant awareness at 62%" means nothing without "vs Tagrisso at 78%."
5. **Barrier analysis produces actionable recommendations.** ATU's unique contribution is diagnosing WHY trial/usage lags despite awareness — barriers (access, formulary, efficacy concerns, safety perception). Recommendations must name barriers and proposed mitigations.

## Phase 0 — Hypothesis Validation

Same mechanics as sfea-insight-writer Phase 0 (validate each hypothesis against survey data), but with ATU-specific hypothesis types:

### ATU Hypothesis Types

| Type | Example | Validation method |
|---|---|---|
| **FUNNEL_POSITION** | "Unaided awareness will trail competitors in Academic segment" | Compare awareness metrics across brands × segments |
| **CONVERSION_RATE** | "Awareness-to-trial conversion improved vs prior wave" | Compute (trial% / awareness%) and compare to prior wave |
| **SHARE_SHIFT** | "New Rx share gained 3pp from competitor switching" | Compare current share vs prior; decompose by source of switch |
| **BARRIER_HYPOTHESIS** | "Trial barriers are primarily access-related, not efficacy" | Cross-tabulate barrier responses; rank by frequency |
| **LOYALTY_DRIVER** | "Formulary position is the strongest loyalty driver" | Correlate loyalty metrics with satisfaction / access / efficacy attributes |
| **SEGMENT_DIVERGENCE** | "Community HCPs have significantly lower awareness than Academic" | Segment comparison with stat-sig testing |
| **PRIOR_WAVE_VALIDATION** | "Prior wave recommended awareness campaign; check if awareness moved" | Compare prior wave awareness to current; cite the recommendation being evaluated |

### Output: `validated_analysis.md`

For each hypothesis:
- **Verdict:** SUPPORTED / REFUTED / MIXED / INSUFFICIENT_DATA
- **Evidence:** Prior and current values, competitive context, segment breakdown
- **Funnel position:** Where in the brand funnel this finding sits
- **Statistical notes:** Sample size, significance markers

## Phase 1 — Narrative Arc Synthesis

### ATU Arc Patterns (distinct from SFEA's CONVERGENCE/TENSION/DIVERGENCE/CLOSURE)

| Arc Pattern | Use When | Example |
|---|---|---|
| **FUNNEL_LEAKAGE** | High awareness but low trial (or high trial but low loyalty) — the funnel narrows faster than expected | "62% awareness but only 18% trial — conversion bottleneck at the consideration stage" |
| **SHARE_MOMENTUM** | Brand share is moving meaningfully (up or down) wave over wave | "New Rx share climbed from 22% to 28%, driven by Academic segment adoption" |
| **COMPETITIVE_CONVERGENCE** | Gap between brand and a competitor is closing | "Awareness gap with Tagrisso narrowed from 20pp to 8pp over 3 waves" |
| **BARRIER_CLUSTER** | Multiple barriers cluster around a theme (access, safety, efficacy) | "Three of the top 5 trial barriers relate to formulary restrictions — access is the primary obstacle" |
| **LOYALTY_EROSION** | Usage or loyalty metrics declining despite stable awareness/trial | "Loyalty score dropped 8pp among Community HCPs while awareness held steady — suggests post-trial disappointment" |
| **SEGMENT_SPLIT** | One segment behaves fundamentally differently from another | "Academic HCPs at 72% awareness vs Community at 48% — a 24pp gap that widened from 15pp last wave" |
| **ADOPTION_CURVE** | New product showing classic adoption pattern (innovators → early majority) | "Trial concentrated in Academic centers; Community adoption just beginning" |

### Headline Style for ATU (differs from PET)

| PET headline pattern | ATU headline pattern |
|---|---|
| "Efficacy recall dipped 3pp QoQ, driven by Safety" | "Rybrevant unaided awareness reaches 62%, narrowing the gap with Tagrisso" |
| "Message believability up across all segments" | "Trial barriers cluster around formulary access, not efficacy concerns" |
| "Rep quality drives 2x the message recall" | "Awareness-to-trial conversion improved 4pp, driven by Academic adopters" |

**ATU headline rules:**
- Lead with share/level, not delta
- Always name the competitive context
- Funnel position in the headline when relevant ("awareness-to-trial conversion")
- Barrier headlines name the barrier theme, not just the magnitude
- Segment headlines name the segments, not just "segments differ"

### Executive Summary for ATU

ATU ES bullets answer funnel-centric KBQs:
- "Where are we in the funnel?" (awareness → trial → usage share)
- "Where is the funnel leaking?" (which stage has the biggest drop-off)
- "What's driving the leak?" (barriers, competitive pressure, segment gaps)
- "What should we do?" (barrier mitigation, segment targeting, positioning shift)

Every ES bullet cites specific slide_ids (same as sfea-insight-writer).

### Recommendations for ATU

ATU recommendations target market access + brand positioning (not rep strategy):
- **ACCELERATE** — awareness is growing, trial conversion is healthy; invest in the current trajectory
- **UNBLOCK** — awareness is high but trial is low; the named barrier needs a specific intervention
- **DEFEND** — competitor is gaining share; competitive positioning needs sharpening
- **EXPAND** — one segment is adopting; extend the playbook to lagging segments
- **INVESTIGATE** — loyalty is eroding with no clear driver; qualitative deep-dive needed

## Inputs

Same as sfea-insight-writer:
- `hypothesis_bank.md` (from `hypothesis-generator`, framed with ATU hypothesis types)
- `source_data.json` (from `excel-indexer` or `synapse-read`)
- `project_context.md`, `market_context.md`, `prior_wave_context.md` (from context builders)

## Outputs

- `context/{wave}/validated_analysis.md` — every hypothesis validated with funnel-position tagging
- `context/{wave}/narrative_threads.md` — ATU arcs + ATU-style headlines + funnel-centric ES + barrier-aware recommendations

## Orchestration

Invoked by `atu-deck` project skill within `create-deck-workflow`:

```
hypothesis-generator (with ATU hypothesis types)
  → atu-insight-writer Phase 0 → validated_analysis.md
  → atu-insight-writer Phase 1 → narrative_threads.md
    → slide-plan-generator-hypothesis (reads ATU arcs)
      → slide-creator renders ATU slides
```

For `refresh-deck-workflow`:
```
deck-reader → prior specs
new data → atu-insight-writer re-validates hypotheses against fresh data
  → slide-plan-generator-refresh diffs against prior validated_analysis
  → slide-updater refreshes each slide + headline-writer regenerates with ATU framing
```

## Decision Rules

| Situation | Response |
|---|---|
| Hypothesis bank uses PET/SFEA terminology | Flag mismatch; suggest re-framing with ATU hypothesis types (or switch to sfea-insight-writer if this is actually a PET study) |
| Data doesn't have funnel metrics (no awareness / trial / usage) | Proceed with whatever metrics exist but warn that funnel arcs can't be built; fall back to generic CONVERGENCE/TENSION patterns |
| No competitor data available | Build arcs from brand's own trajectory (share momentum, segment split) rather than competitive convergence; note gap in ES |
| Barrier questions not in the survey | Skip BARRIER_CLUSTER arcs; note in recommendations as "barrier analysis not available — recommend adding to next wave" |
| Very few waves of data (wave 1 or 2) | Can't build ADOPTION_CURVE or trend arcs; focus on FUNNEL_LEAKAGE and cross-sectional SEGMENT_SPLIT |

## Relationship to other skills

| Skill | Relationship |
|---|---|
| `sfea-insight-writer` | Parallel — PET/SFEA equivalent. Same two-phase structure, different methodology framing. |
| `hypothesis-generator` | Upstream — generates hypothesis bank. When invoked from `atu-deck`, should produce ATU hypothesis types. |
| `headline-writer` | Downstream — regenerates headlines on data refresh. ATU headlines use share language (this skill's Phase 1 output sets the initial headline patterns). |
| `slide-plan-generator-hypothesis` | Downstream — reads narrative_threads.md to build slide plan with ATU arc assignments. |
| `atu-deck` | Parent project skill — invokes atu-insight-writer in both create + refresh flows. |

## References

- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` — PET equivalent (structural template)
- `.claude/skills/projects/atu-deck/SKILL.md` — ATU project skill (invokes this)
- `experiments/deck_analysis/outputs/atu_analysis.md` — ATU deck analysis findings (8 decks, 7 clients)
- PRD §4.4 (analysis skills), §3 Workflows 1-2 (create + refresh)
