"""
highlighter.py — Message Highlighter analysis and slide builder.

Standalone utility (Path A): uses slidegen pptx_utils chrome for template
consistency but is NOT part of the automated pipeline. Invoke on demand via
natural language or the /message-highlighter skill.

Usage:
    from slidegen.pptx_utils.highlighter import build_highlighter_deck
    build_highlighter_deck(
        data_path="path/to/Message highlighter data.xlsx",
        template_path="path/to/template.pptx",
        output_path="path/to/output.pptx",
        messages=MESSAGES,       # list of message config dicts
        col_map=COL_OVERALL,     # {code_group: col_index}
    )

Data format: Each Excel cell contains JSON arrays of highlight ranges:
    [{"startIndex": 24, "charLength": 12}, ...]
"""

from __future__ import annotations

import json
import logging
from typing import Optional

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# COLOUR PALETTE
# ══════════════════════════════════════════════════════════════════════════════

JJ_RED      = RGBColor(0xFF, 0x00, 0x00)
RYB_ORANGE  = RGBColor(0xF7, 0x58, 0x24)
BLACK       = RGBColor(0x00, 0x00, 0x00)
GRAY_DARK   = RGBColor(0x50, 0x50, 0x50)
GRAY_MED    = RGBColor(0x7F, 0x7F, 0x7F)
GRAY_MID    = RGBColor(0x70, 0x70, 0x70)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
GREEN_POS   = RGBColor(0x00, 0xB0, 0x50)
GREEN_DARK  = RGBColor(0x00, 0x7A, 0x33)
NAVY        = RGBColor(0x00, 0x46, 0x86)
GOLD        = RGBColor(0xC8, 0x86, 0x0A)
LIGHT_GREEN = RGBColor(0xE2, 0xF4, 0xE8)
LIGHT_RED   = RGBColor(0xFD, 0xE8, 0xE8)
ZEBRA       = RGBColor(0xF2, 0xF2, 0xF2)

BAND_COLOR = {
    '>50%':   RYB_ORANGE,
    '41-50%': GOLD,
    '31-40%': GRAY_MID,
    None:     BLACK,
}
BAND_BOLD = {'>50%': True, '41-50%': True, '31-40%': False, None: False}
BAND_DEFS = [
    ('>50%',   '>50% Selected',         BAND_COLOR['>50%']),
    ('41-50%', '41\u201350% Selected',  BAND_COLOR['41-50%']),
]


# ══════════════════════════════════════════════════════════════════════════════
# DATA ANALYSIS — JSON parsing, coverage, phrase extraction
# ══════════════════════════════════════════════════════════════════════════════

def load_highlights(df: pd.DataFrame, col_idx: int) -> list[list[tuple[int, int]]]:
    """Parse JSON highlight ranges from an Excel column.

    Returns list of highlight-range lists, one per respondent with at least
    one valid range. Each range is (startIndex, charLength).
    """
    result = []
    for val in df.iloc[:, col_idx]:
        if pd.isna(val) or str(val).strip() in ('', '[]', 'nan'):
            continue
        try:
            entries = json.loads(str(val))
            ranges = [(e.get('startIndex', 0), e.get('charLength', 0))
                      for e in entries if e.get('charLength', 0) > 0]
            if ranges:
                result.append(ranges)
        except Exception:
            pass
    return result


def get_band(pct: float) -> Optional[str]:
    """Classify coverage percentage into a frequency band."""
    if pct > 50:  return '>50%'
    if pct > 40:  return '41-50%'
    return None


def compute_coverage(
    respondent_highlights: list[list[tuple[int, int]]],
    msg_text: str,
) -> tuple[list[float], int]:
    """Character-level coverage: % of respondents who highlighted each position.

    Returns (pct_array, n_respondents).
    """
    msg_len = len(msg_text)
    n = len(respondent_highlights)
    if n == 0:
        return [0.0] * msg_len, 0
    counts = [0] * msg_len
    for ranges in respondent_highlights:
        for start, length in ranges:
            for i in range(start, min(start + length, msg_len)):
                counts[i] += 1
    return [c / n * 100 for c in counts], n


def segment_by_band(
    cov_pct: list[float], msg_text: str,
) -> list[tuple[str, Optional[str]]]:
    """Split message text into (fragment, band) segments for multi-color rendering."""
    n = min(len(cov_pct), len(msg_text))
    segs, i = [], 0
    while i < n:
        band = get_band(cov_pct[i])
        j = i + 1
        while j < n and get_band(cov_pct[j]) == band:
            j += 1
        segs.append((msg_text[i:j], band))
        i = j
    if n < len(msg_text):
        segs.append((msg_text[n:], None))
    return segs


# ── Phrase extraction helpers ─────────────────────────────────────────────

STOPWORDS = frozenset({
    'and', 'or', 'the', 'a', 'an', 'in', 'of', 'to', 'with', 'for',
    'is', 'are', 'was', 'be', 'by', 'at', 'on', 'it', 'its', 'as',
    'but', 'from', 'this', 'that', 'have', 'had', 'also', 'not',
})


def _tokenize(msg_text: str) -> list[dict]:
    """Split on spaces; return list of {text, start, end, band}."""
    words, i = [], 0
    while i < len(msg_text):
        while i < len(msg_text) and msg_text[i] == ' ':
            i += 1
        if i >= len(msg_text):
            break
        j = i
        while j < len(msg_text) and msg_text[j] != ' ':
            j += 1
        words.append({'text': msg_text[i:j], 'start': i, 'end': j, 'band': None})
        i = j
    return words


def _meaningful(phrase_words: list[str], min_regular: int = 7) -> bool:
    """True if phrase contains sufficient content words."""
    clean = [t.strip('.,;:()[]®™%') for t in phrase_words]
    content = [t for t in clean if t.lower() not in STOPWORDS and len(t) >= 2]
    if len(content) >= 2:
        return True
    if len(content) == 1:
        t = content[0]
        is_abbrev = (any(c.isdigit() or c == '%' for c in t) or
                     (len(t) > 1 and any(c.isupper() for c in t[1:])))
        return len(t) >= (3 if is_abbrev else min_regular)
    return False


def _trim_edges(words_list: list[str]) -> list[str]:
    """Strip leading/trailing stopword tokens."""
    while words_list and words_list[0].strip('.,;:()[]®™').lower() in STOPWORDS:
        words_list.pop(0)
    while words_list and words_list[-1].strip('.,;:()[]®™').lower() in STOPWORDS:
        words_list.pop()
    return words_list


def extract_band_phrases(
    respondent_highlights: list[list[tuple[int, int]]],
    msg_text: str,
    n_total: int,
) -> dict[str, list[str]]:
    """Word-level phrase extraction grouped by frequency band.

    Returns {'>50%': [...], '41-50%': [...], '31-40%': [...]}.
    """
    words = _tokenize(msg_text)
    if n_total > 0:
        for w in words:
            count = 0
            for ranges in respondent_highlights:
                for start, length in ranges:
                    if start < w['end'] and (start + length) > w['start']:
                        count += 1
                        break
            w['band'] = get_band(count / n_total * 100)

    bands = {'>50%': [], '41-50%': [], '31-40%': []}

    for target_band in ['>50%', '41-50%', '31-40%']:
        i = 0
        while i < len(words):
            if words[i]['band'] != target_band:
                i += 1
                continue
            phrase_words = [words[i]['text']]
            j = i + 1
            while j < len(words):
                w = words[j]
                if w['band'] == target_band:
                    phrase_words.append(w['text'])
                    j += 1
                elif (w['text'].lower().strip('.,;:()[]®™') in STOPWORDS and
                      j + 1 < len(words) and words[j + 1]['band'] == target_band):
                    phrase_words.append(w['text'])
                    j += 1
                else:
                    break

            display_words = _trim_edges(list(phrase_words))
            if not display_words:
                i = j
                continue

            if _meaningful(display_words):
                phrase = ' '.join(display_words).strip(' ,;:.()')
                if phrase not in bands[target_band]:
                    bands[target_band].append(phrase)
            i = j

    # Deduplicate: remove lower-band phrases that are substrings of higher-band ones
    seen: set = set()
    for bk in ['>50%', '41-50%', '31-40%']:
        filtered = []
        for ph in bands[bk]:
            ph_l = ph.lower()
            if not any(ph_l in s.lower() for s in seen):
                filtered.append(ph)
                seen.add(ph)
        bands[bk] = filtered

    return bands


def analyze_message(
    df: pd.DataFrame, msg_text: str, col_idx: int,
) -> dict:
    """Full analysis pipeline for one message+group column.

    Returns {n, cov, segs, bands} ready for rendering.
    """
    highlights = load_highlights(df, col_idx)
    cov, n = compute_coverage(highlights, msg_text)
    bands = extract_band_phrases(highlights, msg_text, n)
    segs = segment_by_band(cov, msg_text)
    return {'n': n, 'cov': cov, 'segs': segs, 'bands': bands}


# ══════════════════════════════════════════════════════════════════════════════
# SHAPE HELPERS — slide primitives
# ══════════════════════════════════════════════════════════════════════════════

def _add_filled_box(slide, left, top, width, height, fill_rgb,
                    border_rgb=None, border_pt=0.5):
    box = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height))
    box.fill.solid()
    box.fill.fore_color.rgb = fill_rgb
    if border_rgb:
        box.line.color.rgb = border_rgb
        box.line.width = Pt(border_pt)
    else:
        box.line.fill.background()
    return box


def _set_para_line_spacing(p, pct=115):
    pPr = p._pPr if p._pPr is not None else p._p.get_or_add_pPr()
    lnSpc = etree.SubElement(pPr, qn("a:lnSpc"))
    spcPct = etree.SubElement(lnSpc, qn("a:spcPct"))
    spcPct.set("val", str(int(pct * 1000)))


def _set_para_space_before(p, pts=0):
    pPr = p._pPr if p._pPr is not None else p._p.get_or_add_pPr()
    spcBef = etree.SubElement(pPr, qn("a:spcBef"))
    spcPts = etree.SubElement(spcBef, qn("a:spcPts"))
    spcPts.set("val", str(int(pts * 100)))


def _add_title(slide, text, font_hdr="Johnson Display", size=17):
    txb = slide.shapes.add_textbox(
        Inches(0.33), Inches(0.08), Inches(12.67), Inches(0.92))
    tf = txb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.0)
    tf.margin_top = Inches(0.0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.name = font_hdr
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = JJ_RED
    return txb


def _add_breadcrumb(slide, text, font_hdr="Johnson Display"):
    txb = slide.shapes.add_textbox(
        Inches(9.02), Inches(0.002), Inches(4.307), Inches(0.25))
    tf = txb.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = text
    run.font.name = font_hdr
    run.font.size = Pt(10)
    run.font.color.rgb = GRAY_DARK
    return txb


def _add_message_box_multicolor(slide, segs, left, top, width, height,
                                 font_body="Johnson Text"):
    """Multi-color text box: different colors per frequency band."""
    txb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.12)
    tf.margin_top = Inches(0.08)
    tf.margin_right = Inches(0.12)
    tf.margin_bottom = Inches(0.04)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _set_para_line_spacing(p, 120)
    for text, band in segs:
        if not text:
            continue
        r = p.add_run()
        r.text = text
        r.font.name = font_body
        r.font.size = Pt(10)
        r.font.bold = BAND_BOLD.get(band, False)
        r.font.color.rgb = BAND_COLOR.get(band, BLACK)
    return txb


def _add_band_legend_line(slide, left, top, width, font_body="Johnson Text"):
    """Colour-swatch legend for the message text highlights."""
    txb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(0.22))
    tf = txb.text_frame
    tf.margin_left = Inches(0.04)
    tf.margin_top = Inches(0.0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT

    def _run(text, color, bold=False, italic=False, size=7.5):
        r = p.add_run()
        r.text = text
        r.font.name = font_body
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color

    _run("Highlighted text = HM group selection frequency:   ", GRAY_MED, italic=True)
    _run("\u25a0", BAND_COLOR['>50%'], bold=True, size=9)
    _run(" >50%     ", GRAY_MED)
    _run("\u25a0", BAND_COLOR['41-50%'], bold=True, size=9)
    _run(" 41\u201350%", GRAY_MED)
    return txb


def _add_column_header(slide, left, top, width, fill_rgb, line1, line2="",
                        font_hdr="Johnson Display", font_body="Johnson Text"):
    h = 1.02 if line2 else 0.40
    _add_filled_box(slide, left, top, width, h, fill_rgb)
    txb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(h))
    tf = txb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.12)
    tf.margin_top = Inches(0.07)
    tf.margin_right = Inches(0.08)
    tf.margin_bottom = Inches(0.04)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.LEFT
    _set_para_line_spacing(p1, 100)
    r1 = p1.add_run()
    r1.text = line1
    r1.font.name = font_hdr
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = WHITE

    if line2:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        _set_para_line_spacing(p2, 115)
        _set_para_space_before(p2, 2)
        r2 = p2.add_run()
        r2.text = line2
        r2.font.name = font_body
        r2.font.size = Pt(8.5)
        r2.font.bold = False
        r2.font.italic = True
        r2.font.color.rgb = WHITE

        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.LEFT
        _set_para_line_spacing(p3, 110)
        _set_para_space_before(p3, 3)
        r3 = p3.add_run()
        r3.text = "\u201c\u201d italic = verbatim HCP quote"
        r3.font.name = font_body
        r3.font.size = Pt(7)
        r3.font.italic = False
        r3.font.color.rgb = (RGBColor(0xFF, 0xDD, 0xDD)
                             if fill_rgb == JJ_RED else RGBColor(0xCC, 0xEE, 0xD5))

    return txb, h


def _add_banded_phrase_block(slide, left, top, width, height, bands,
                              count_color, font_hdr="Johnson Display",
                              font_body="Johnson Text"):
    """Frequency-banded phrase list grouped by band (>50%, 41-50%)."""
    txb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.10)
    tf.margin_top = Inches(0.05)
    tf.margin_right = Inches(0.06)
    tf.margin_bottom = Inches(0.04)

    para_idx = 0
    for band_key, band_label, band_col in BAND_DEFS:
        phrases = [ph for ph in bands.get(band_key, []) if ph.strip()]
        if not phrases:
            continue

        p_hdr = tf.paragraphs[0] if para_idx == 0 else tf.add_paragraph()
        para_idx += 1
        p_hdr.alignment = PP_ALIGN.LEFT
        _set_para_line_spacing(p_hdr, 100)
        if para_idx > 1:
            _set_para_space_before(p_hdr, 7)
        r_hdr = p_hdr.add_run()
        r_hdr.text = band_label.upper()
        r_hdr.font.name = font_hdr
        r_hdr.font.size = Pt(8)
        r_hdr.font.bold = True
        r_hdr.font.color.rgb = band_col

        for phrase in phrases:
            p = tf.add_paragraph()
            para_idx += 1
            p.alignment = PP_ALIGN.LEFT
            _set_para_line_spacing(p, 115)
            _set_para_space_before(p, 2)

            r_sq = p.add_run()
            r_sq.text = "\u25a0  "
            r_sq.font.name = font_body
            r_sq.font.size = Pt(8)
            r_sq.font.color.rgb = band_col

            r_phr = p.add_run()
            r_phr.text = f'"{phrase}"'
            r_phr.font.name = font_body
            r_phr.font.size = Pt(9)
            r_phr.font.color.rgb = BLACK

    if para_idx == 0:
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = "No phrases above threshold"
        r.font.name = font_body
        r.font.size = Pt(9)
        r.font.italic = True
        r.font.color.rgb = GRAY_MED

    return txb


def _add_divider_label(slide, left, top, width, label, label_color,
                        font_hdr="Johnson Display"):
    txb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(0.22))
    tf = txb.text_frame
    tf.margin_left = Inches(0.08)
    tf.margin_top = Inches(0.01)
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = label
    run.font.name = font_hdr
    run.font.size = Pt(8.5)
    run.font.bold = True
    run.font.color.rgb = label_color
    return txb


def _add_bullet_list(slide, left, top, width, height, items,
                      text_color=None, font_size=9,
                      font_body="Johnson Text"):
    if text_color is None:
        text_color = GRAY_DARK
    QUOTE_COLOR = RGBColor(0x40, 0x40, 0x40)
    txb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_top = Inches(0.04)
    tf.margin_right = Inches(0.06)
    tf.margin_bottom = Inches(0.04)

    for i, item in enumerate(items):
        kind, text = item if isinstance(item, tuple) else ("insight", item)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _set_para_line_spacing(p, 118)
        if i > 0:
            _set_para_space_before(p, 5)

        if kind == "quote":
            rb = p.add_run()
            rb.text = "\u201c  "
            rb.font.name = font_body
            rb.font.size = Pt(font_size + 0.5)
            rb.font.italic = True
            rb.font.color.rgb = GRAY_MED

            rt = p.add_run()
            rt.text = f"\u201c{text}\u201d"
            rt.font.name = font_body
            rt.font.size = Pt(font_size)
            rt.font.italic = True
            rt.font.color.rgb = QUOTE_COLOR
        else:
            rb = p.add_run()
            rb.text = "\u2022  "
            rb.font.name = font_body
            rb.font.size = Pt(font_size)
            rb.font.color.rgb = text_color

            rt = p.add_run()
            rt.text = text
            rt.font.name = font_body
            rt.font.size = Pt(font_size)
            rt.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)

    return txb


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

# Layout constants for message detail slides
_HEADER_H      = 1.02
_TOP_GAP       = 0.07
_SEC_LABEL_H   = 0.20
_PHRASE_ITEM_H = 0.24
_BAND_HDR_H    = 0.20
_BAND_GAP      = 0.06
_DIVIDER_H     = 0.09
_BULLET_ITEM_H = 0.55
_BOTTOM_PAD    = 0.10
_MAX_COL_BOTTOM = 6.90


def _band_block_height(bands: dict) -> float:
    """Height needed for a banded phrase block."""
    h = 0.0
    first = True
    for band_key, _, _ in BAND_DEFS:
        phrases = [ph for ph in bands.get(band_key, []) if ph.strip()]
        if not phrases:
            continue
        if not first:
            h += _BAND_GAP
        h += _BAND_HDR_H + len(phrases) * _PHRASE_ITEM_H
        first = False
    return max(h, 0.30)


def _calc_col_h(phrase_block_h: float, n_bullets: int) -> float:
    return (_HEADER_H + _TOP_GAP
            + _SEC_LABEL_H + phrase_block_h
            + _DIVIDER_H
            + _SEC_LABEL_H + n_bullets * _BULLET_ITEM_H
            + _BOTTOM_PAD)


def _build_cover_slide(prs, blank_layout, messages, msg_cov, *,
                        breadcrumb="Message Highlighter",
                        subtitle="", sample_desc="",
                        font_hdr="Johnson Display", font_body="Johnson Text"):
    """Slide 1: Cover + overview table."""
    slide = prs.slides.add_slide(blank_layout)

    # Red banner
    _add_filled_box(slide, 0, 0, 13.33, 3.40, JJ_RED)

    txb = slide.shapes.add_textbox(Inches(0.6), Inches(0.5), Inches(12.0), Inches(2.0))
    tf = txb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.0)

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "Message Highlighter Analysis"
    r.font.name = font_hdr
    r.font.size = Pt(34)
    r.font.bold = True
    r.font.color.rgb = WHITE

    if subtitle:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        _set_para_space_before(p2, 4)
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.name = font_hdr
        r2.font.size = Pt(16)
        r2.font.color.rgb = WHITE

    if sample_desc:
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.LEFT
        _set_para_space_before(p3, 6)
        r3 = p3.add_run()
        r3.text = sample_desc
        r3.font.name = font_body
        r3.font.size = Pt(10)
        r3.font.italic = True
        r3.font.color.rgb = RGBColor(0xFF, 0xCC, 0xCC)

    # Overview table
    COLS = [0.33, 2.40, 8.80, 11.10]
    COL_W = [1.85, 6.20, 2.10, 2.10]
    HDRS = ["Message", "Key Insight", "HM Group", "WM Group"]
    ROW_H = 0.48
    TBL_TOP = 3.48

    for lft, hdr, wid in zip(COLS, HDRS, COL_W):
        _add_filled_box(slide, lft, TBL_TOP, wid, 0.52, NAVY)
        txb = slide.shapes.add_textbox(
            Inches(lft), Inches(TBL_TOP), Inches(wid), Inches(0.52))
        tf = txb.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.06)
        tf.margin_top = Inches(0.04)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = hdr
        run.font.name = font_hdr
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = WHITE

    for ri, msg in enumerate(messages):
        code = msg['code']
        hm_n = msg_cov.get(f'{code}_HM', {}).get('n', 0)
        wm_n = msg_cov.get(f'{code}_WM', {}).get('n', 0)
        label = msg.get('label', f"{msg.get('rank', '')} {code}")
        insight = msg.get('cover_insight', msg.get('title', ''))
        row_data = [label, insight, f"n={hm_n}", f"n={wm_n}"]

        row_top = TBL_TOP + 0.52 + ri * ROW_H
        fill = ZEBRA if ri % 2 == 0 else WHITE
        for ci, (lft, val, wid) in enumerate(zip(COLS, row_data, COL_W)):
            _add_filled_box(slide, lft, row_top, wid, ROW_H, fill,
                            border_rgb=RGBColor(0xD8, 0xD8, 0xD8), border_pt=0.25)
            al = PP_ALIGN.CENTER if ci > 1 else PP_ALIGN.LEFT
            txb = slide.shapes.add_textbox(
                Inches(lft + 0.04), Inches(row_top + 0.04),
                Inches(wid - 0.08), Inches(ROW_H - 0.04))
            tf = txb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.04)
            tf.margin_top = Inches(0.0)
            p = tf.paragraphs[0]
            p.alignment = al
            run = p.add_run()
            run.text = val
            run.font.name = font_hdr if ci == 0 else font_body
            run.font.size = Pt(8.5) if ci == 1 else Pt(10)
            run.font.bold = (ci == 0)
            run.font.color.rgb = JJ_RED if ci == 0 else BLACK

    _add_breadcrumb(slide, breadcrumb, font_hdr)

    # Legend note
    txb = slide.shapes.add_textbox(Inches(0.33), Inches(7.08), Inches(12.67), Inches(0.32))
    tf = txb.text_frame
    tf.margin_left = Inches(0.04)
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = ("n = respondents with valid highlight responses  |  "
                "HM group: rated message as Highly Motivating  |  "
                "WM group: rated as Low/Neutral (Would Motivate More)  |  "
                "Message text coloured by % of HM group who selected each phrase")
    run.font.name = font_body
    run.font.size = Pt(7.5)
    run.font.color.rgb = GRAY_MED

    return slide


def _build_message_slide(prs, blank_layout, msg, msg_cov, *,
                          breadcrumb="Message Highlighter",
                          font_hdr="Johnson Display", font_body="Johnson Text"):
    """One message detail slide: title + message text + two-column HM vs WM."""
    slide = prs.slides.add_slide(blank_layout)
    code = msg['code']
    hm_cov = msg_cov.get(f'{code}_HM', {})
    wm_cov = msg_cov.get(f'{code}_WM', {})
    hm_segs = hm_cov.get('segs', [(msg['text'], None)])
    hm_bands = hm_cov.get('bands', {'>50%': [], '41-50%': [], '31-40%': []})
    wm_bands = wm_cov.get('bands', {'>50%': [], '41-50%': [], '31-40%': []})
    hm_n = hm_cov.get('n', 0)
    wm_n = wm_cov.get('n', 0)
    n_messages = msg.get('n_messages', 4)

    # Title
    _add_title(slide, msg['title'], font_hdr, size=13)
    _add_breadcrumb(slide, breadcrumb, font_hdr)

    # Rank badge
    badge = slide.shapes.add_textbox(Inches(0.33), Inches(0.88), Inches(1.5), Inches(0.18))
    bp = badge.text_frame.paragraphs[0]
    br = bp.add_run()
    br.text = f"Message {msg['rank']} of {n_messages}  |  {code}"
    br.font.name = font_body
    br.font.size = Pt(8)
    br.font.color.rgb = GRAY_MED

    # Red rule
    _add_filled_box(slide, 0.33, 1.04, 12.67, 0.025, JJ_RED)

    # Message text box
    MSG_BOX_H = 0.85
    MSG_TEXT_TOP = 1.065
    _add_filled_box(slide, 0.33, MSG_TEXT_TOP, 12.67, MSG_BOX_H,
                    ZEBRA, border_rgb=RGBColor(0xC0, 0xC0, 0xC0), border_pt=0.5)

    lbl = slide.shapes.add_textbox(
        Inches(0.40), Inches(MSG_TEXT_TOP + 0.03), Inches(1.4), Inches(0.18))
    lr = lbl.text_frame.paragraphs[0].add_run()
    lr.text = "Message tested:"
    lr.font.name = font_hdr
    lr.font.size = Pt(8)
    lr.font.bold = True
    lr.font.color.rgb = GRAY_DARK

    _add_message_box_multicolor(
        slide, hm_segs,
        left=0.40, top=MSG_TEXT_TOP + 0.18, width=12.40, height=MSG_BOX_H - 0.22,
        font_body=font_body)

    # Band legend
    LEGEND_TOP = MSG_TEXT_TOP + MSG_BOX_H + 0.03
    _add_band_legend_line(slide, 0.40, LEGEND_TOP, 12.40, font_body)

    # Metrics bar
    metrics_top = LEGEND_TOP + 0.22 + 0.03
    _add_filled_box(slide, 0.33, metrics_top, 12.67, 0.35,
                    RGBColor(0xF5, 0xF5, 0xF5),
                    border_rgb=RGBColor(0xD0, 0xD0, 0xD0), border_pt=0.25)

    for lft, txt, col in [
        (0.45, f"HM group: n={hm_n} respondents with highlights", GREEN_DARK),
        (6.5, f"WM group: n={wm_n} respondents with highlights", JJ_RED),
    ]:
        t = slide.shapes.add_textbox(
            Inches(lft), Inches(metrics_top + 0.04), Inches(5.8), Inches(0.28))
        p = t.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = txt
        run.font.name = font_body
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = col

    # Two columns
    col_top = metrics_top + 0.35 + 0.07
    col_l_left = 0.33
    col_r_left = 6.87
    col_width = 6.30

    pb_h_l = _band_block_height(hm_bands)
    pb_h_r = _band_block_height(wm_bands)
    col_h_l = _calc_col_h(pb_h_l, len(msg.get('hm_rationale', [])))
    col_h_r = _calc_col_h(pb_h_r, len(msg.get('wm_gaps', [])))
    col_h = min(_MAX_COL_BOTTOM - col_top, max(col_h_l, col_h_r))

    _add_filled_box(slide, col_l_left, col_top, col_width, col_h,
                    LIGHT_GREEN, border_rgb=GREEN_POS, border_pt=0.75)
    _add_filled_box(slide, col_r_left, col_top, col_width, col_h,
                    LIGHT_RED, border_rgb=JJ_RED, border_pt=0.75)

    _, hdr_h = _add_column_header(
        slide, col_l_left, col_top, col_width, GREEN_POS,
        "\u2714  WHAT RESONATES \u2014 Highly Motivating Group",
        msg.get('hm_insight', ''),
        font_hdr=font_hdr, font_body=font_body)
    _add_column_header(
        slide, col_r_left, col_top, col_width, JJ_RED,
        "\u26a0  WHAT\u2019S MISSING \u2014 Would Motivate More Group",
        msg.get('wm_insight', ''),
        font_hdr=font_hdr, font_body=font_body)

    content_top = col_top + hdr_h + _TOP_GAP

    # Left column
    _add_divider_label(slide, col_l_left + 0.08, content_top,
                       col_width - 0.16, "TOP HIGHLIGHTED PHRASES", GREEN_DARK,
                       font_hdr=font_hdr)
    _add_banded_phrase_block(
        slide, col_l_left + 0.08, content_top + _SEC_LABEL_H,
        col_width - 0.16, pb_h_l,
        hm_bands, count_color=GREEN_DARK,
        font_hdr=font_hdr, font_body=font_body)
    div_y_l = content_top + _SEC_LABEL_H + pb_h_l + 0.04
    _add_filled_box(slide, col_l_left + 0.12, div_y_l, col_width - 0.24, 0.02,
                    RGBColor(0xA0, 0xD0, 0xA8))
    rat_top_l = div_y_l + 0.04
    rat_h_l = len(msg.get('hm_rationale', [])) * _BULLET_ITEM_H
    _add_divider_label(slide, col_l_left + 0.08, rat_top_l,
                       col_width - 0.16, "WHY IT RESONATES (RATIONALE)", GREEN_DARK,
                       font_hdr=font_hdr)
    _add_bullet_list(
        slide, col_l_left + 0.08, rat_top_l + _SEC_LABEL_H,
        col_width - 0.16, rat_h_l,
        msg.get('hm_rationale', []), font_size=10, font_body=font_body)

    # Right column
    _add_divider_label(slide, col_r_left + 0.08, content_top,
                       col_width - 0.16, "TOP HIGHLIGHTED PHRASES (WOULD EXPAND)", JJ_RED,
                       font_hdr=font_hdr)
    _add_banded_phrase_block(
        slide, col_r_left + 0.08, content_top + _SEC_LABEL_H,
        col_width - 0.16, pb_h_r,
        wm_bands, count_color=JJ_RED,
        font_hdr=font_hdr, font_body=font_body)
    div_y_r = content_top + _SEC_LABEL_H + pb_h_r + 0.04
    _add_filled_box(slide, col_r_left + 0.12, div_y_r, col_width - 0.24, 0.02,
                    RGBColor(0xF0, 0xA0, 0xA0))
    rat_top_r = div_y_r + 0.04
    rat_h_r = len(msg.get('wm_gaps', [])) * _BULLET_ITEM_H
    _add_divider_label(slide, col_r_left + 0.08, rat_top_r,
                       col_width - 0.16, "WHAT\u2019S MISSING / WHAT THEY WANT MORE OF", JJ_RED,
                       font_hdr=font_hdr)
    _add_bullet_list(
        slide, col_r_left + 0.08, rat_top_r + _SEC_LABEL_H,
        col_width - 0.16, rat_h_r,
        msg.get('wm_gaps', []), font_size=10, font_body=font_body)

    return slide


def _build_implications_slide(prs, blank_layout, implications, *,
                               breadcrumb="Message Highlighter",
                               font_hdr="Johnson Display", font_body="Johnson Text"):
    """Final slide: Key Implications for Field Messaging."""
    slide = prs.slides.add_slide(blank_layout)
    _add_title(slide, "Key Implications for Field Messaging", font_hdr, size=17)
    _add_breadcrumb(slide, breadcrumb, font_hdr)
    _add_filled_box(slide, 0.33, 1.00, 12.67, 0.025, JJ_RED)

    row_top = 1.06
    for impl in implications:
        color = impl.get('color', GRAY_MED)
        if isinstance(color, str):
            color = RGBColor(int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16))
        title = impl['title']
        detail = impl['detail']
        row_h = 0.80

        _add_filled_box(slide, 0.33, row_top, 0.10, row_h, color)
        _add_filled_box(slide, 0.43, row_top, 12.57, row_h,
                        ZEBRA, border_rgb=RGBColor(0xD0, 0xD0, 0xD0), border_pt=0.25)

        txb = slide.shapes.add_textbox(
            Inches(0.55), Inches(row_top + 0.06), Inches(12.2), Inches(0.32))
        run = txb.text_frame.paragraphs[0].add_run()
        run.text = title
        run.font.name = font_hdr
        run.font.size = Pt(10.5)
        run.font.bold = True
        run.font.color.rgb = color

        txb2 = slide.shapes.add_textbox(
            Inches(0.55), Inches(row_top + 0.38), Inches(12.2), Inches(0.38))
        txb2.text_frame.word_wrap = True
        txb2.text_frame.margin_left = Inches(0.0)
        txb2.text_frame.margin_top = Inches(0.0)
        run2 = txb2.text_frame.paragraphs[0].add_run()
        run2.text = detail
        run2.font.name = font_body
        run2.font.size = Pt(9)
        run2.font.color.rgb = GRAY_DARK

        row_top += row_h + 0.025

    return slide


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════════════

def build_highlighter_deck(
    data_path: str,
    template_path: str,
    output_path: str,
    messages: list[dict],
    col_map: dict[str, int],
    *,
    data_start_row: int = 6,
    groups: tuple[str, ...] = ('HM', 'WM'),
    breadcrumb: str = "Message Highlighter",
    subtitle: str = "",
    sample_desc: str = "",
    implications: list[dict] | None = None,
    blank_layout_idx: int = 24,
    font_hdr: str = "Johnson Display",
    font_body: str = "Johnson Text",
) -> str:
    """Build a complete Message Highlighter deck.

    Args:
        data_path:     Path to the highlighter Excel data file.
        template_path: Path to the PPTX template.
        output_path:   Where to save the generated deck.
        messages:      List of message config dicts. Each must have:
                       - code: str (e.g. "R33Z")
                       - rank: str (e.g. "#1")
                       - title: str (headline)
                       - text: str (the full message text tested)
                       - hm_insight: str
                       - wm_insight: str
                       - hm_rationale: list of ("quote", text) tuples
                       - wm_gaps: list of ("quote", text) tuples
                       Optional:
                       - label: str (cover table label)
                       - cover_insight: str (cover table insight)
                       - n_messages: int (total count for "Message X of N")
        col_map:       Dict mapping "CODE_GROUP" → 0-based column index
                       e.g. {"R33Z_HM": 13, "R33Z_WM": 33, ...}
        data_start_row: 0-based row where respondent data begins (default 6).
        groups:        Tuple of group suffixes (default ("HM", "WM")).
        breadcrumb:    Breadcrumb text for all slides.
        subtitle:      Cover slide subtitle line.
        sample_desc:   Cover slide sample description line.
        implications:  List of dicts with {color, title, detail} for
                       the implications slide. If None, skipped.
        blank_layout_idx: Index of the blank slide layout (default 24).
        font_hdr:      Header font name.
        font_body:     Body font name.

    Returns:
        Path to the saved PPTX file.
    """
    # Load data
    logger.info("Loading highlighter data: %s", data_path)
    df_raw = pd.read_excel(data_path, header=None)
    df_data = df_raw.iloc[data_start_row:].reset_index(drop=True)

    # Analyze all messages × groups
    msg_cov = {}
    n_messages = len(messages)
    for msg in messages:
        code = msg['code']
        msg['n_messages'] = n_messages
        for grp in groups:
            key = f"{code}_{grp}"
            col = col_map.get(key)
            if col is None:
                continue
            result = analyze_message(df_data, msg['text'], col)
            msg_cov[key] = result
            active = [(b, phrases) for b, phrases in result['bands'].items() if phrases]
            logger.info("  %s: n=%d  %s", key, result['n'],
                        "  |  ".join(f"{b}: {len(ph)} phrases" for b, ph in active))

    # Build presentation
    logger.info("Building deck from template: %s", template_path)
    prs = Presentation(template_path)
    blank_layout = prs.slide_masters[0].slide_layouts[blank_layout_idx]

    # Remove existing slides
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    sldIdLst = prs.slides._sldIdLst
    rId_list = [el.get(f'{{{NS_R}}}id') for el in list(sldIdLst)]
    for rId in rId_list:
        try:
            prs.part.drop_rel(rId)
        except Exception:
            pass
    for el in list(sldIdLst):
        sldIdLst.remove(el)

    # Slide 1: Cover
    _build_cover_slide(prs, blank_layout, messages, msg_cov,
                       breadcrumb=breadcrumb, subtitle=subtitle,
                       sample_desc=sample_desc,
                       font_hdr=font_hdr, font_body=font_body)

    # Slides 2–N+1: Message detail
    for msg in messages:
        _build_message_slide(prs, blank_layout, msg, msg_cov,
                             breadcrumb=breadcrumb,
                             font_hdr=font_hdr, font_body=font_body)

    # Final slide: Implications (optional)
    if implications:
        _build_implications_slide(prs, blank_layout, implications,
                                  breadcrumb=breadcrumb,
                                  font_hdr=font_hdr, font_body=font_body)

    # Save
    prs.save(output_path)
    logger.info("Saved highlighter deck: %s (%d slides)", output_path, len(prs.slides))
    return output_path
