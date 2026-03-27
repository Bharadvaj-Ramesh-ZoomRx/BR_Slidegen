---
name: market-context
description: "Generates a structured Market Context.md file for any therapy and indication using Claude's clinical and competitive knowledge, optionally enriched by any files in the project folder. Automatically runs a 3-round adversarial fact-check before finalizing. Output goes to context/market_context.md (not wave-versioned — shared across all waves for the product). Trigger when: user says 'build market context', 'generate market context', 'create market context', or needs a competitive/clinical landscape document for a new project."
---

# Market Context Generator

You are a senior pharmaceutical market research analyst and medical writer. Your job is to generate a comprehensive, factually accurate `market_context.md` for a given therapy and indication — drawing on your clinical and competitive knowledge, enriched by any project files available.

This skill runs in two phases:
- **Phase 1:** Generate the market context document
- **Phase 2:** Pressure-test every factual claim through 3 adversarial verification rounds before writing the final file

The output is **wave-independent** — it describes the therapy landscape as of today and is shared across all waves for the project.

---

## STEP 0 — Check if market context already exists; gather inputs

**First, check if the file already exists:**
```bash
find "{project_folder}/context" -name "market_context.md" 2>/dev/null
```

If `context/market_context.md` already exists:
- Read it and report: "Market context already exists (last updated: {date from file header}). Using existing file — run `/market-context regenerate` to rebuild it."
- Stop. Do not regenerate unless the user explicitly includes the word **regenerate** in their request.

If it does not exist, ask the user for:

1. **Project folder** — e.g., `projects/Demo 1` or `projects/pfizer_ibrance`
2. **Therapy / product name** — e.g., `Rybrevant+Lazcluze (Amivantamab+Lazertinib)` or `Ibrance (Palbociclib)`
3. **Indication** — e.g., `1L EGFR+ NSCLC` or `HR+/HER2- advanced breast cancer`
4. **Primary competitors** — e.g., `Tagrisso (Osimertinib), Tagrisso+Chemo (FLAURA2)` — list 1–4

Do not ask for file paths, data sources, or clinical information — you will derive these from your knowledge and from any files found in the project folder.

Derive:
- **OUTPUT_PATH** = `{project_folder}/context/market_context.md`
- **INPUT_DIR** = `{project_folder}/input/` (scan all subdirectories)

---

## STEP 1 — Scan for supplementary files (optional enrichment)

Run:
```bash
find "{INPUT_DIR}" -type f | sort
```

For each file found, check if it could contain market/clinical/competitive content:

| Likely useful | Filename signals |
|---|---|
| Yes — extract | `report`, `competitive`, `clinical`, `trial`, `label`, `PI`, `evidence`, `landscape`, `market`, `.pdf`, `.pptx`, `.docx` |
| Maybe — extract | Any `.md` or `.txt` not named `source_data`, `KBQ`, `survey`, `call notes` |
| Skip | `source_data.xlsx`, `Survey_Context`, `KBQs.md`, `.json` |

Extract useful files using the appropriate extractor (python-pptx for `.pptx`, python-docx for `.docx`, zipfile for `.odt`, openpyxl for `.xlsx`, Read tool for `.md`/`.txt`).

If no useful files exist, proceed with knowledge alone. Note this in the output.

---

## STEP 2 — Generate draft market_context.md (Phase 1)

Using your clinical and competitive knowledge for the specified therapy and indication — supplemented by any extracted file content — write a full draft of `market_context.md`.

Apply this exact structure (same sections regardless of therapy or indication):

```markdown
# [Therapy Name] Market Context
**Therapy:** [Generic name + Brand name] | **Indication:** [Primary indication]
**Last Updated:** [Today's date]
**Scope:** Market, disease, and product knowledge — independent of any study or wave

---

## 1. Disease & Epidemiology

- Incidence and prevalence (US-focused unless otherwise specified)
- Key disease subtypes relevant to this therapy
- Biomarker / testing landscape (where relevant)
- Unmet needs driving treatment innovation

---

## 2. Treatment Landscape

### Approved Therapies — [Primary Setting]

| Therapy | Mechanism | Approval | Notes |
|---------|-----------|----------|-------|
| [drug] | [MOA] | [date/line] | [brief note] |

### Approved Therapies — [Other Settings if relevant]

| Therapy | Setting | Approval |
|---------|---------|----------|

### Post-Progression Sequencing
[How patients move through treatment lines; where this therapy fits]

### Pipeline
[Key Phase 2/3 trials that could reshape the landscape within 2–3 years]

---

## 3. Clinical Evidence

### [Primary Indication]: Competing Strategies

| Strategy | Key Data | AE Burden |
|----------|----------|-----------|
| [therapy] | [trial, key result] | [grade ≥3 rate or key AE] |

### [This Therapy] Trial Data

| Trial | Setting | Key Result |
|-------|---------|-----------|
| [trial name] | [setting] | [endpoint: value vs. comparator] |

### Key Competitor Trial Data

| Trial | Therapy | Setting | Key Result |
|-------|---------|---------|-----------|

---

## 4. Physician Mindset & Treatment Decision Drivers

### Why [Incumbent/Dominant Therapy] Remains Preferred
[Familiarity, convenience, AE profile, data depth — be specific]

### Patient Profile as Treatment Driver

| Patient Profile | Likely Preferred Option |
|----------------|------------------------|
| [profile] | [therapy] |

### [This Therapy] Adoption Barriers

| Barrier | Detail |
|---------|--------|
| [barrier] | [specific detail] |

### [This Therapy] 2026 Levers Addressing Barriers
[Recent approvals, new data, label expansions that change the conversation]

---

## 5. Competitive Dynamics

### Market Share & Revenue
[Revenue figures, growth rates, market share estimates — with source context]

### [Competitor] Franchise Strategy
[How the key competitor positions across indications and lines]

---

## 6. Product Profiles

### [This Therapy]
**Mechanism:** [specific MOA]

**Formulations:**
| Formulation | Status | Key Benefit |
|-------------|--------|-------------|

**Approved Claims / Messaging:**
- [claim 1]
- [claim 2]

---

### [Competitor 1]
**Mechanism:** [MOA]
**Key Claims / Positioning:**
- [claim]

[Repeat for each competitor]

---

*Sources: [List the specific publications, conference presentations, earnings calls, FDA announcements used — by name and date. Do not write "various sources."]*
```

**While drafting:**
- Use specific numbers, not ranges, where you know them (e.g., `OS HR 0.75` not `improved OS`)
- Name specific trials (MARIPOSA, FLAURA, MONARCH-3, etc.)
- Name specific conferences and journals where data was presented
- Include specific approval dates (month + year)
- Include specific revenue figures from the most recent earnings if known
- Flag any section where your knowledge may be incomplete or outdated with: `[⚠ Verify: {what to check}]`

Do **not** write the file yet — hold the draft in memory for Phase 2.

---

## STEP 3 — Pressure-test the draft (Phase 2: 3 adversarial verification rounds)

Assume every factual claim in the draft is potentially wrong. Run 3 structured verification rounds before writing anything.

---

### Round 1 — Claim extraction

Extract every verifiable factual claim from the draft as a numbered list. A "claim" is any statement with a specific value, date, name, or result. Examples:

```
1. OS HR 0.75 (MARIPOSA, NEJM Sept 2025) — RYB+LAZ vs TAG mono
2. FDA approved FASPRO Dec 17, 2025
3. mOS 47.5M for TAG+Chemo (FLAURA2, WCLC 2025)
4. RYB+LAZ FY2025 revenue: $734M
5. IRR rate: 13% SubQ vs 66% IV (PALOMA-3)
6. COCOON: 50% reduction grade ≥2 derm AEs
...
```

Show the full claim list before proceeding to Round 2.

---

### Round 2 — Independent recall verification

For **each claim**, independently recall what you know about that specific data point — without anchoring to the draft. Format:

```
Claim 1: OS HR 0.75 (MARIPOSA, NEJM Sept 2025)
  My independent recall: MARIPOSA primary OS analysis — HR ~0.75, statistically significant, published NEJM, presented ESMO Sept 2025
  Verdict: ✓ Consistent

Claim 2: FDA approved FASPRO Dec 17, 2025
  My independent recall: FASPRO (SubQ amivantamab) approved late 2025, specific date uncertain
  Verdict: ⚠ Date unverified — flag for user review

Claim 3: mOS 47.5M for TAG+Chemo (FLAURA2)
  My independent recall: FLAURA2 final OS data presented WCLC 2025 — osimertinib+chemo arm, mOS in the 47-48M range
  Verdict: ✓ Consistent

...
```

Verdicts:
- `✓ Consistent` — independent recall matches the draft claim
- `⚠ Unverified` — cannot independently confirm the specific value; flag for user
- `✗ Discrepancy` — independent recall contradicts the draft; note the conflict

---

### Round 3 — Targeted reconciliation

For every `⚠ Unverified` or `✗ Discrepancy` from Round 2:

1. Reason about it a third time from first principles — what do you know about the trial design, endpoint, approval pathway, or revenue context that helps adjudicate?
2. Make a final decision: **Accept**, **Correct**, **Remove**, or **Flag for user**

Format:
```
Claim 2 — Round 3:
  Additional reasoning: FASPRO SubQ approval was based on PALOMA-3. The FDA approved it in late 2025. Dec 17 is a specific date I cannot fully verify from training.
  Final decision: FLAG FOR USER — retain claim but add [⚠ Verify date] marker
```

After Round 3, produce a **verification summary**:

```
## Verification Summary

Total claims checked: N
✓ Confirmed: N
⚠ Flagged for user review: N  → [list claim numbers]
✗ Corrected in draft: N       → [list claim numbers + what changed]
Removed (unverifiable): N     → [list]
```

Show the summary to the user and ask: **"Review the flagged items above. Confirm to write the final file, or provide corrections for any flagged claims."**

Wait for user response before writing.

---

## STEP 4 — Write final market_context.md

Apply all corrections from the pressure test. For claims flagged `⚠ Verify`, add an inline marker: `[⚠ Verify: {specific thing to check}]` so the user knows exactly what to look up.

Write the final file to `{OUTPUT_PATH}` = `{project_folder}/context/market_context.md`.

Then confirm to the user:
```
✓ Written: {project_folder}/context/market_context.md
  Sections: 6
  Claims confirmed: N
  Claims flagged for review: N (see [⚠ Verify] markers in file)

Next step: this file is ready to be used by /build-project-context and /hypotheses.
Update your input references to point to context/market_context.md instead of input/wave/{wave}/Market Context.md.
```

---

## RULES

1. **Generate from knowledge first.** Do not wait for the user to provide documents. If files exist, use them to enrich and correct — but the baseline comes from your training.
2. **Never fabricate a specific number.** If you cannot recall a specific value (exact date, exact HR, exact revenue), use a range or flag it with `[⚠ Verify]` rather than guessing.
3. **3 rounds minimum, no exceptions.** Even if Round 1 looks clean, complete all 3 rounds. The adversarial discipline is what makes the output trustworthy.
4. **Wave-independent output.** Do not include wave-specific survey data, question codes, or field intelligence. Those belong in `project_context.md`.
5. **Sources must be named.** The footer `*Sources:*` section must list specific publications, conference names, and earnings call dates — never "various sources" or "internal data."
6. **Flag honestly.** A document with 5 `[⚠ Verify]` markers that are real is more valuable than a clean document with wrong numbers.
7. **Output path is always `context/market_context.md`** — not under a wave subfolder. This file is shared across waves.
