---
name: build-project-context
description: "Use when building or updating a Project_Context.md file from source documents (ODT methodology files, Excel hypothesis sheets, client call notes). Trigger when: user says 'build project context', 'update project context', or needs to synthesize study design, KBQs, wave hypotheses, and methodology notes into a structured context file for downstream skills like /hypotheses."
---

# Build Project Context

You are a ZoomRx senior analyst. Your job is to read the PET study methodology, KBQ list, and wave-specific hypothesis documents and synthesize them into a clean, structured `Project_Context.md` file that the `/hypotheses` skill can use at runtime.

---

## STEP 0 — Gather file paths from the user

Ask the user for:

1. **Project folder** — e.g., `projects/jnj_rybrevant`
2. **Study methodology + KBQs file** — an `.odt` file (e.g., `reference/context/pet_project_kbq.odt`)
3. **Wave hypotheses + client questions file** — an `.xlsx` file with a sheet containing market/client context and business hypotheses (e.g., `reference/context/rybrevant_market_context_hypotheses.xlsx`). Also ask which **sheet name** to read.
4. **Client call notes file** — a `.docx` or `.txt` file (e.g., `reference/context/rybrevant_call_notes.docx`)
5. **Output path** — where to write `Project_Context.md` (default: `{project_folder}/reference/context/Project_Context.md`)

All paths are relative to the repo root. If the user provides partial info, infer reasonable defaults from the project folder structure.

Store these as variables for the steps below:
- `PROJECT_DIR` — project folder path
- `ODT_PATH` — methodology/KBQ file path
- `XLSX_PATH` — hypotheses Excel file path
- `XLSX_SHEET` — sheet name to read
- `NOTES_PATH` — call notes file path
- `OUTPUT_PATH` — output file path

---

## STEP-BY-STEP PROCESS

### Step 1 — Extract last 6 months of client call notes

Read the call notes file. If it's a `.docx`, use `python -m markitdown` or `docx2txt` to extract text first. If `.txt`, read directly.

Then filter to last 6 months:

```bash
python3 -c "
import re
from datetime import datetime, timedelta

cutoff = datetime.today() - timedelta(days=180)

with open('{NOTES_PATH}', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

entries = re.split(r'(?=Date:\s*\d{1,2}/\d{1,2}/202\d)', content)
recent = []
for e in entries:
    m = re.search(r'Date:\s*(\d{1,2})/(\d{1,2})/(202\d)', e)
    if m:
        try:
            d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            if d >= cutoff:
                recent.append(e)
        except:
            pass

print(f'Entries found in last 6 months: {len(recent)}')
for r in recent:
    print(r[:3000])
    print('---')
"
```

Use only this extracted content to inform Sections 3 (Wave Hypotheses) and the Open Action Items in the output. Do not reference older call notes content.

---

### Step 2a — Extract ODT content

Run this Bash command to extract all text from the ODT file:

```bash
python3 -c "
import zipfile, re
with zipfile.ZipFile('{ODT_PATH}') as z:
    with z.open('content.xml') as f:
        content = f.read().decode('utf-8')
        text = re.sub(r'<[^>]+>', ' ', content)
        text = re.sub(r'\s+', ' ', text)
        print(text)
"
```

### Step 2b — Extract Excel sheet content

Run this Bash command to extract all content from the specified sheet:

```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('{XLSX_PATH}')
ws = wb['{XLSX_SHEET}']
for row in ws.iter_rows(values_only=True):
    non_empty = [str(c).strip() for c in row if c]
    if non_empty:
        print(' | '.join(non_empty))
"
```

### Step 3 — Ask for wave-specific additions

Ask the user:

1. **Current wave** — e.g., Q1 2026
2. **Any new or updated KBQs** this wave not already in the source files?
3. **Any message changes this wave** — new message codes introduced, messages retired, textual modifications?
4. **Any methodology changes** — survey changes, new questions, sample changes, screener updates?
5. **Any open client action items** from the most recent readout to carry forward?

If the user says "none" or "same as before," proceed with source file content only.

### Step 4 — Synthesize and write Project_Context.md

Using all extracted content and user inputs, write `Project_Context.md` to `{OUTPUT_PATH}` with the structure below. Be specific and non-generic — use real trial names, real message domains, real segment names, real sample sizes.

---

## OUTPUT STRUCTURE

Write the file with exactly these sections:

```
# Project Context — [BRAND/STUDY NAME]
**Wave:** [current wave]
**Last Updated:** [today's date]

---

## 1. Study Design

### Methodology
[Describe touchpoints — what is captured at each, who qualifies, what the base unit is (interactions vs. unique HCPs)]

### Sample & Cadence
- [Research cadence — quarterly, monthly, etc.]
- [Sample targets per brand/product]
- [Any wave-specific sample changes]

### HCP Segments Tracked
- [Segment 1]: [definition]
- [Segment 2]: [definition]
- [Practice setting segments if applicable]

---

## 2. Key Business Questions (KBQs)

### [Domain 1 — e.g., Sales Activity]
[List all KBQs from source, verbatim or lightly edited for clarity]

### [Domain 2 — e.g., Messaging]
[List all KBQs]

### [Domain 3 — e.g., Rep Performance]
[List all KBQs]

[Add/remove domain sections based on what the source files contain]

---

## 3. Wave-Specific Hypotheses & Client Questions

For each topic area from the Excel sheet, capture:
- **Secondary research context** (what is known from public/external sources)
- **Client-shared intelligence** (what the client has told us)
- **Hypothesis going into this wave** (what we expect to see)
- **Questions to answer** (the specific analytical questions the client wants answered)

Format each block as:

### [Topic Area]
**Context:** [secondary research + client intel combined, deduplicated]
**Hypothesis:** [what we expect to find this wave]
**Questions to answer:**
- [question 1]
- [question 2]

---

## 4. Analytical Priorities

List the top 5-7 analytical priorities for this wave, in order of client importance, based on:
- Open action items from prior readouts
- Highest-priority KBQs
- Metric areas showing movement or concern

---

## 5. Methodology Notes

Document any survey or methodology changes this wave that could affect metric comparability:
- [e.g., SOV screener change, new question format, benchmark changes, new message codes]
- Flag which metrics are NOT comparable wave-over-wave due to changes

---

## 6. Message Reference

List all active message codes/names being tracked this wave, with:
- Message domain (Efficacy / Safety / Convenience / etc.)
- New this wave? (Y/N)
- Baseline available? (Y/N — N if new this wave)
```

---

## RULES

1. **Always run all extraction steps before writing.** Do not write the context file from memory alone.
2. **Be specific.** Never write generic placeholders like "message recall metrics will be tracked." Name the actual messages, actual segments, actual trial names.
3. **Wave hypothesis section is the most critical.** This is what feeds the `/hypotheses` skill. Every topic area from the Excel sheet must appear here, even if the hypothesis is brief.
4. **Methodology notes must be exhaustive.** Any change that could create a false wave-over-wave signal must be documented. This prevents the team from presenting a methodology artifact as a real finding.
5. **Do not duplicate Market Context.** This file covers project/study design and wave-specific expectations only. Clinical data, competitive landscape, and physician mindset live in a separate Market Context file.
6. **Confirm before writing.** After Step 3, show the user a brief outline of what will be written (section headers + 1-line summary each) and ask for confirmation before generating the full file.
7. **No hardcoded paths.** All file paths come from user input in Step 0. This skill works for any project.
