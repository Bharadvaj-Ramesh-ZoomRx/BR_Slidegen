---
name: create-deck-workflow
effort: high
paths: []
description: "Triggered by: 'Create a new PET deck for <brand>' / 'Build a deck from scratch' / 'Storyboard a hypothesis-driven deck'. Produces a complete first-cut deck from project brief + KBQs + data. Supports two modes: briefing-driven (default — data + KBQs → slides) and hypothesis-driven (storyboarding path — explicit hypothesis bank → narrative arcs → slides)."
---

# create-deck-workflow

Full-stack deck creation from scratch. Consolidates new-deck + storyboarding into one skill with a mode flag.

## Trigger phrases

**Either mode (briefing-driven by default):**
- "Create a new PET deck for Pfizer Product X"
- "Build the Q1 deck for Rybrevant from scratch"
- "Set up a new deck in projects/jnj_new_brand"

**Hypothesis-driven (explicit):**
- "Build a hypothesis-driven PET deck"
- "Storyboard a new deck from these hypotheses"
- "Use the narrative-threads flow to build this"

## Modes

**`mode="briefing"` (default):** data + KBQs + project context → hypotheses generated → narrative → slides. Typical new-deck path.

**`mode="hypothesis"` (Vinoth's storyboarding):** user brings a pre-built `hypothesis_bank.md`; skips hypothesis-generator; goes straight to validation + narrative threads. Deck's every slide maps to a pre-defined arc.

**Detection:** if user utterance contains "hypothesis" / "storyboard", or `hypothesis_bank.md` exists and is non-stub → hypothesis mode. Else briefing mode.

## Cardinal Rules

1. **Project-type-aware.** For PET projects invoke `pet-deck` wrapper; for ATU, `atu-deck`. Project-type skills apply methodology defaults.
2. **Context-first.** Always build context (market + project + prior wave + survey) before hypotheses. Skipping context produces shallow output.
3. **Main content gate after narrative threads.** Per PRD §7 stage model — the user reviews the narrative backbone before any rendering.
4. **Every spec passes `spec-validator`** before `slide-creator` runs.
5. **Reuse existing context** if already generated for this wave — don't rebuild unless user says "regenerate".

## Inputs

- `config_path`: project config.yaml
- `mode` (optional): `"briefing"` | `"hypothesis"` (default: auto-detect)
- Files in `projects/{name}/input/wave/{wave}/`:
  - `KBQs.md` (required)
  - `hypothesis_bank.md` (required if `mode="hypothesis"`)
  - call_notes, methodology odt, survey draft, prior wave files (optional)

## Outputs

- Full deck at `projects/{name}/output/{wave}/deck.pptx`
- Context files at `projects/{name}/context/{wave}/`
- Narrative threads, slide plan, shape registry

## Orchestration

```
Stage 0 — auto-index, no gate
1. index_excel / synapse-read → source_data.json
2. index_qualitative → qualitative_data.json (optional)

Stage 0.5 — parallel context (subagents)
3. market-context-builder → context/market_context.md (if missing)
4. prior-wave-context-builder → context/{wave}/prior_wave_context.md (if prior files)
5. survey-context-builder → context/{wave}/survey_context.md (if survey draft)

Stage 1
6. project-context-builder → context/{wave}/project_context.md

GATE: user reviews all 4 context files

Stage 2 — mode-dependent
7. if mode="briefing":
     hypothesis-generator → context/{wave}/hypothesis_bank.md
   if mode="hypothesis":
     (skip — user-provided hypothesis_bank.md is used)

GATE: user confirms hypothesis count + domain coverage

Stage 3
8. insight-writer (sfea-insight-writer for PET, atu-insight-writer for ATU — selected by project-type skill)
     Phase 0: validated_analysis.md
     Phase 1: narrative_threads.md

GATE: user reviews narrative threads (the main content gate)

Stage 4
9. slide-plan-generator-hypothesis → slide_plan.md (every slide tagged to an arc)

GATE: user confirms slide count + arc distribution

Stage 5 — internal, no gate
10. for each slide in plan:
      viz-selector + layout-selector
      headline-writer (or reuse from narrative_threads)
      assemble SlideSpec + metadata.arc + metadata.role_in_arc
      spec-validator(spec, strict=True)
      slide-creator(spec)

11. deck-assembler(specs, output_path)
```

## Decision Rules

| Situation | Response |
|---|---|
| `KBQs.md` missing | Halt at Stage 1 — user must author it |
| `mode="hypothesis"` but `hypothesis_bank.md` missing or stub | Halt — offer to switch to briefing mode |
| Prior wave files present | Invoke prior-wave-context-builder; else skip |
| Survey draft present | Invoke survey-context-builder; else code-less mode |
| `market_context.md` already exists | Skip regeneration unless user says "regenerate market context" |
| Narrative has arcs but some slides don't map | Allowed; metadata notes "new in this wave" |
| User wants PET-specific structure | Orchestrator is invoked via `pet-deck` which applies PET defaults |

## References

- PRD §3 Workflow 1, §5 composition map, §7 stage model
- `.claude/skills/context-data/*` — context builders
- `.claude/skills/analysis/hypothesis-generator/SKILL.md`
- `.claude/skills/analysis/sfea-insight-writer/SKILL.md` (PET), `.claude/skills/analysis/atu-insight-writer/SKILL.md` (ATU)
- `.claude/skills/planning/slide-plan-generator-hypothesis/SKILL.md`
- `.claude/skills/planning/viz-selector/SKILL.md`, `layout-selector/SKILL.md`
- `.claude/skills/creation/slide-creator/SKILL.md`, `deck-assembler/SKILL.md`
- `.claude/skills/projects/pet-deck/SKILL.md` (PET-specific wrapper)
