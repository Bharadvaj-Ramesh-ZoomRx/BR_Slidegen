---
name: survey-context
effort: medium
description: "Pre-Stage 1 sub-skill. Converts any survey draft file (Word, PDF, Excel, text) into a structured survey_context.md file. Parses question codes, question text, response scales, message lists, module architecture, and segment definitions. Output goes to context/{wave}/survey_context.md and is consumed by /hypotheses ('Test with:' lines) and /slide-plan (question code mapping). Trigger when: user says 'build survey context', 'convert survey draft', 'extract survey questions', or a survey questionnaire file exists in the project input folder."
---

## Auto-Detected Context
!`python3 -c "
import glob, json, os
try:
    import yaml
except ImportError:
    yaml = None
configs = sorted(glob.glob('projects/*/config.yaml'), key=os.path.getmtime, reverse=True) if yaml else []
if configs:
    try:
        with open(configs[0]) as f:
            cfg = yaml.safe_load(f)
    except Exception:
        cfg = {}
    if not isinstance(cfg, dict): cfg = {}
    proj = os.path.dirname(configs[0])
    wave = cfg.get('project',{}).get('wave','')
    print(f'**Active project:** \`{proj}\`')
    print(f'**Active wave:** \`{wave}\`')
    ctx = f'{proj}/context/{wave}' if wave else f'{proj}/context'
    for md in sorted(glob.glob(f'{ctx}/*.md')):
        print(f'  - \`{os.path.basename(md)}\` ({os.path.getsize(md)//1024}KB)')
    sj = f'{ctx}/source_data.json'
    if os.path.exists(sj):
        with open(sj) as f:
            idx = json.load(f)
        codes = idx.get('_codes',{})
        print(f'**source_data.json:** {sum(len(v) for v in codes.values())} codes indexed')
    for case in ['Wave', 'wave']:
        inp = f'{proj}/input/{case}/{wave}'
        if os.path.isdir(inp):
            files = [f for f in os.listdir(inp) if not f.startswith('.')]
            print(f'**Input files ({len(files)}):** ' + ', '.join(sorted(files)[:10]))
            break
else:
    print('**No active project detected** — user must specify project folder')
"
`

# Survey Context Extractor

You are a ZoomRx senior methodologist. Your job is to read any survey draft file — regardless of format or study type — and convert it into a structured `survey_context.md` that downstream skills can reliably use to anchor hypotheses to real question codes and generate accurate slide plans.

This skill is format-agnostic and study-type-agnostic. It works for PET/SFEA surveys, awareness/attitude studies, patient chart audits, or any quantitative questionnaire.

---

## STEP 0 — Confirm project and wave

Use the auto-detected project and wave shown above. If correct, proceed. If not, ask the user to specify.

Also ask for:
- **Study/product name** — e.g., `RYBREVANT+LAZCLUZE SFEA` (used in the output header)

Derive:
- **INPUT_DIR** = `{project_folder}/input/Wave/{wave}` (check both `Wave` and `wave` casing)
- **CONTEXT_DIR** = `{project_folder}/context/{wave}/`
- **OUTPUT_PATH** = `{project_folder}/context/{wave}/survey_context.md`
- **SOURCE_JSON** = `{project_folder}/context/{wave}/source_data.json` (use for code cross-check if it exists)

---

## STEP 1 — Detect survey files in the input folder

Run:
```bash
find "{INPUT_DIR}" -type f | sort
```

Classify each file using these signals:

| Role | Filename signals | Extensions |
|------|-----------------|------------|
| **Survey draft** (primary) | `survey`, `questionnaire`, `draft`, `instrument`, `screener`, `v1`, `v2`, `final` | `.docx`, `.pdf`, `.txt`, `.xlsx` |
| **Message exhibit / aid** | `message`, `aid`, `exhibit`, `show card`, `list`, `stimulus` | `.docx`, `.pdf`, `.xlsx` |
| **Prior survey context** | `survey_context`, `Survey_Context` | `.md` — use to carry forward unchanged questions |
| **Source data** | `source_data` | `.xlsx`, `.json` — use only for code cross-check, not extraction |
| **Skip** | `call_notes`, `market_context`, `KBQ`, `template`, `prior_wave` | any |

Show the user the detected files and ask: **"These files appear to contain survey content — confirm to proceed, or add/remove files."**

Wait for confirmation before extracting.

---

## STEP 2 — Extract raw content from all confirmed files

### `.docx` — structured extraction preserving paragraph order

```bash
python -c "
from docx import Document
import re

doc = Document(r'{FILE_PATH}')
print('=== PARAGRAPHS ===')
for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if text:
        style = para.style.name if para.style else ''
        print(f'[{i}|{style}] {text}')

print()
print('=== TABLES ===')
for t_idx, table in enumerate(doc.tables):
    print(f'-- Table {t_idx+1} --')
    for row in table.rows:
        cells = [c.text.strip() for c in row.cells]
        if any(cells):
            print(' | '.join(cells))
"
```

### `.pdf` — page-by-page extraction

```bash
python -c "
import pdfplumber
with pdfplumber.open(r'{FILE_PATH}') as pdf:
    for i, page in enumerate(pdf.pages, 1):
        text = page.extract_text()
        if text and text.strip():
            print(f'=== Page {i} ===')
            print(text)
"
```

If pdfplumber unavailable:
```bash
python -c "
import fitz
doc = fitz.open(r'{FILE_PATH}')
for i, page in enumerate(doc, 1):
    print(f'=== Page {i} ===')
    print(page.get_text())
"
```

### `.xlsx` — all sheets

```bash
python -c "
import openpyxl
wb = openpyxl.load_workbook(r'{FILE_PATH}', data_only=True)
for sheet_name in wb.sheetnames:
    print(f'=== Sheet: {sheet_name} ===')
    ws = wb[sheet_name]
    for row in ws.iter_rows(values_only=True):
        cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
        if cells:
            print(' | '.join(cells))
"
```

### `.md` / `.txt` — read directly with Read tool

---

## STEP 3 — Parse the extracted content

After extraction, run the following parsing script against the full extracted text to identify question codes, question text, response scales, and message lists:

```bash
python -c "
import re, sys

text = open(r'{EXTRACTED_TEXT_PATH}', encoding='utf-8').read()

# ── Question code detection ───────────────────────────────────────────────────
# Matches patterns like: Q1.87, Q2.10, NPP Q1.00, ME Q2.20C, Q1.85b,
# Impact Q1.15, Retention Q1.00, S0.65, Q27, Q20
code_pattern = re.compile(
    r'\b((?:NPP|ME|Impact|Retention|Screener|S)\s+)?'
    r'(Q\d+(?:\.\d+)?[a-zA-Z]?)\b'
)

# ── Scale detection ───────────────────────────────────────────────────────────
scale_pattern = re.compile(
    r'(?:1\s*=|scale\s+of\s+1|1-7|select\s+all|yes\s*/\s*no|rank|allocate\s+\d+\s+points)',
    re.IGNORECASE
)

# ── Section header detection ──────────────────────────────────────────────────
section_pattern = re.compile(
    r'^(?:MODULE|SECTION|PART|TOUCHPOINT|TP[12]|SCREENING|SCREENER|'
    r'PERSONAL PROMO|NON-PERSONAL|NPP|RETENTION|MESSAGE EFFECT|IMPACT)',
    re.IGNORECASE | re.MULTILINE
)

lines = text.split('\n')
found_codes = []
for i, line in enumerate(lines):
    codes = code_pattern.findall(line)
    for prefix, code in codes:
        full_code = (prefix.strip() + ' ' + code).strip()
        # Get surrounding context (2 lines before, 5 after)
        context_start = max(0, i - 2)
        context_end = min(len(lines), i + 6)
        context = '\n'.join(lines[context_start:context_end])
        found_codes.append({'code': full_code, 'line': i, 'context': context})

# Deduplicate by code
seen = set()
for item in found_codes:
    if item['code'] not in seen:
        seen.add(item['code'])
        print(f\"CODE: {item['code']}\")
        print(f\"CONTEXT:\\n{item['context']}\")
        print('---')
"
```

Save the extracted text to a temp file first if needed:
```bash
python -c "
from docx import Document
doc = Document(r'{FILE_PATH}')
with open(r'{TEMP_PATH}', 'w', encoding='utf-8') as f:
    for para in doc.paragraphs:
        if para.text.strip():
            f.write(para.text + '\n')
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                f.write(' | '.join(cells) + '\n')
print('Saved to {TEMP_PATH}')
"
```

---

## STEP 3b — Cross-check codes against source_data.json (if available)

If `source_data.json` exists in the context folder, cross-check every extracted question code against the `_sheets` index to verify it appears in the actual data:

```bash
python -c "
import json

with open(r'{SOURCE_JSON}', encoding='utf-8') as f:
    data = json.load(f)

sheets = data.get('_sheets', {})
all_codes = set()
for sheet_name, rows in sheets.items():
    for row in rows:
        code = row.get('code', '').strip()
        if code:
            all_codes.add(code)

survey_codes = {EXTRACTED_CODES_LIST}
for code in sorted(survey_codes):
    base = code.split()[-1]  # strip prefix like 'NPP', 'ME'
    found = base in all_codes or code in all_codes
    status = 'FOUND' if found else 'NOT IN DATA'
    print(f'{status}: {code}')
"
```

Flag any code that is NOT FOUND in source data — these may be new questions, screener-only questions, or codes that don't reach the data file. Note them in the output.

---

## STEP 4 — Identify message lists

Message lists (aided recall exhibits) are typically presented as:
- Numbered lists in a show card / exhibit section of the questionnaire
- Tables in a Word document with columns for message number, label, and full text
- Sections labeled "Message List", "Show Card X", "Exhibit", "Aid"

For each message found, extract:
- **Number / label** (e.g., "1", "OS Message", "NCCN-NSCLC")
- **Full message text** (verbatim)
- **Brand** it belongs to (if multi-brand survey)
- **Domain** (Efficacy / Safety / MoA / NCCN / Convenience / etc.) — infer from content
- **New this wave?** — compare against prior wave's message list if `prior_wave_context.md` or an existing `survey_context.md` exists in context folder
- **Baseline available?** — if new this wave, baseline = No

If message lists appear in a separate exhibit file, extract from that file using the same methods.

---

## STEP 5 — Synthesize survey_context.md

Using all parsed content, write `{OUTPUT_PATH}` with exactly this structure:

```markdown
# Survey Context — {study_name}
**Wave:** {wave}
**Source:** {source filename(s)}
**Last Updated:** {today's date}
**Purpose:** Maps KBQs to question codes and question text, defines response scales, and lists tracked messages — used to anchor "Test with:" lines in hypothesis generation

---

## 1. Survey Architecture

| Module | Touchpoint | Base Unit | Sample Target |
|--------|-----------|-----------|--------------|
| {module name} | {TP1/TP2/etc.} | {interactions/responses/HCPs} | {brand: N} |

**Notes on sample or methodology changes this wave:**
- {any notes found in the survey about sample size changes, new screener criteria, etc.}

---

## 2. Question Map by KBQ Domain

Organize every extracted question into the appropriate KBQ domain. Use these standard domain names where applicable — add new domains as needed for the specific study:

### Sales Activity / SOV
| What it measures | Question | Question text |
|-----------------|---------|--------------|
| {description} | {code} | "{verbatim question text}" |

### Messaging — Recall & Delivery
| What it measures | Question | Question text |
|-----------------|---------|--------------|

### Messaging — Effectiveness
| What it measures | Question | Question text |
|-----------------|---------|--------------|

### Interactions — Setup & Narrative
| What it measures | Question | Question text |
|-----------------|---------|--------------|

### Rep Performance
| What it measures | Question | Question text |
|-----------------|---------|--------------|

### Branded Close / Call to Action
| What it measures | Question | Question text |
|-----------------|---------|--------------|

### Prescription Intent
| What it measures | Question | Scale | Question text |
|-----------------|---------|-------|--------------|

### Product Perception & Usage
| What it measures | Question | Question text |
|-----------------|---------|--------------|

### NPP (if present)
| What it measures | Question | Question text |
|-----------------|---------|--------------|

{Add or remove domain sections based on what the survey actually contains}

**Questions not mapped to a domain:**
| Question | Question text | Reason not mapped |
|---------|--------------|------------------|
| {code} | {text} | {e.g., screener only / admin / routing} |

---

## 3. Key Metric Definitions & Scales

| Metric | Question | Scale | Reporting Cut | Question text (abbreviated) |
|--------|---------|-------|--------------|----------------------------|
| {metric name} | {code} | {e.g., 1–7} | {e.g., Top-2-box (≥6)} | "{abbreviated question text}" |

### Derived Segment / Composite Definitions
{Any metric that is computed from two or more questions — e.g., HII definition}

**{HII or other composite}:**
{Formula — e.g., "Q1.87 overall quality = 7 (top-box) AND Q1.85b LTIP ≥ 6 (top-2-box) — both required"}

---

## 4. {Brand 1} Message List — {wave} (Tracked in {question code})

| # | Label | Message Text | Domain | New this wave? | Baseline? |
|---|-------|-------------|--------|---------------|-----------|
| 1 | {label} | {full verbatim message text} | {domain} | Y/N | Y/N |

**Notes:**
- {Any messages that are tracked differently — e.g., topics tracked via Q1.50 rather than the main MR list}
- {Any messages retired this wave}

---

## 5. {Brand 2} Message List — {wave} (if multi-brand survey)

{Same table structure}

---

## 6. Segment Operationalization

| Segment | How Defined in Survey | Question |
|---------|----------------------|---------|
| {segment name} | {definition — e.g., top-box on Q1.87 AND top-2-box on Q1.85b} | {question code(s)} |

---

## 7. Codes Not Found in Source Data

{Only present if source_data.json was cross-checked}

| Question | Status | Likely reason |
|---------|--------|--------------|
| {code} | Not in data | {e.g., screener only / new this wave / routing — not all respondents see} |
```

---

## STEP 6 — Quality check before writing

Before writing the file, show the user a summary:

```
Proposed survey_context.md:

  Source files:      {N files}
  Questions parsed:  {N total codes}
  Domains covered:   {list}
  Messages found:    {N for Brand 1} + {N for Brand 2}
  Codes in data:     {N confirmed} / {N total} ({N flagged as not found})
  New questions:     {N} (vs. prior wave context if available)

Confirm to write, or flag any issues.
```

Wait for user confirmation.

---

## STEP 7 — Write and confirm

Write the file to `{OUTPUT_PATH}`.

Then confirm:
```
✓ Written: {OUTPUT_PATH}

  {N} question codes mapped across {N} domains
  {N} messages documented ({N} brands)
  {N} codes flagged — not found in source_data.json (see Section 7)

This file is ready for:
  → /hypotheses          (reads it for "Test with:" question codes)
  → /build-project-context  (reads it for Section 6 message reference)
  → /slide-plan          (reads it for question code mapping)

Run /build-project-context next (or /hypotheses if project context already exists).
```

---

## RULES

1. **Verbatim question text.** Copy question text exactly as written in the survey draft. Do not paraphrase, summarize, or clean up. The downstream skills need the exact wording to understand what each question measures.
2. **Every code must have a row.** If a question code appears in the survey but you cannot extract its full text (e.g., page cut off, extraction artifact), add it with a `[⚠ text not extracted]` note rather than omitting it.
3. **Domain assignment is your judgment.** The survey draft does not label domains — you classify each question based on what it measures. When in doubt, add it to "Questions not mapped to a domain" and let the user decide.
4. **Message text is verbatim.** The full text of each message in the aided recall exhibit must be copied exactly — abbreviations, brand names, data claims, and all. These are what reps say and HCPs hear.
5. **New-this-wave detection requires prior context.** Mark a message as "New this wave: Y" only if it does not appear in `prior_wave_context.md` or an existing `survey_context.md` in the context folder. If no prior context exists, mark all as "N (no prior context to compare)".
6. **Cross-check is advisory, not blocking.** If a code is not found in source_data.json, flag it but do not remove it from the output. Screener questions and routing instructions legitimately don't appear in the data.
7. **One confirmation only.** Confirm the file list at Step 1, confirm the summary at Step 6. No other gates.
8. **Output is wave-versioned.** `survey_context.md` goes into `context/{wave}/` — each wave gets its own copy because the message list and question set can change wave-over-wave.
