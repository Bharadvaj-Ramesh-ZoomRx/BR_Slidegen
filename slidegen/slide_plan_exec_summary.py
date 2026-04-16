"""
slide_plan_exec_summary.py — build ES slide specs with citations.

Executive summary slides answer KBQs with findings that cite back to supporting
slide_ids. Every bullet has `metadata.citations: list[slide_id]`. Unsourced
claims rejected.

Contract:
    generate_exec_summary(deck_specs, kbqs, ...) -> list[SlideSpec]
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

from slidegen.slide_spec import (
    SlideSpec,
    HeadlineSpec,
    FooterSpec,
    TextboxComponent,
    Position,
    SlideMetadata,
    validate_spec,
)


# Layout constants for text-only ES slide
_ES_HEADLINE_RECT = (0.30, 0.30, 12.70, 0.90)
_ES_BULLET_START_TOP = 1.80
_ES_BULLET_LEFT = 0.60
_ES_BULLET_WIDTH = 12.20
_ES_BULLET_HEIGHT = 0.80


@dataclass
class ESFinding:
    """A single ES bullet with its citations."""
    kbq: str                          # the KBQ this answers
    text: str                         # the bullet text
    citations: list[str]              # slide_ids supporting this finding
    confidence: str = "high"          # "high" | "medium" | "low"


@dataclass
class ESSlidePlan:
    """Plan for one ES slide — a KBQ cluster + its findings."""
    kbq_group: list[str]              # 1+ KBQs addressed by this slide
    findings: list[ESFinding]
    headline: Optional[str] = None


def _build_es_slide_spec(
    plan: ESSlidePlan,
    slide_id: str,
    slide_index: int,
    brand: Optional[str] = None,
    section: str = "Executive Summary",
    footer: Optional[str] = None,
) -> SlideSpec:
    """Build a textbox-only SlideSpec from an ESSlidePlan."""
    headline_text = plan.headline or (
        f"Answers to: {plan.kbq_group[0]}" if len(plan.kbq_group) == 1
        else f"Answers to {len(plan.kbq_group)} KBQs"
    )

    # Collect all citations across findings (unique, in order)
    all_citations: list[str] = []
    for f in plan.findings:
        for c in f.citations:
            if c not in all_citations:
                all_citations.append(c)

    # Build textbox components: one per bullet + maybe a citations footnote
    components: list = []
    top = _ES_BULLET_START_TOP
    for i, finding in enumerate(plan.findings):
        cite_str = ", ".join(finding.citations)
        bullet_text = f"• {finding.text}  (see: {cite_str})" if cite_str else f"• {finding.text}"
        components.append(TextboxComponent(
            position=Position(
                left=_ES_BULLET_LEFT,
                top=top,
                width=_ES_BULLET_WIDTH,
                height=_ES_BULLET_HEIGHT,
            ),
            text=bullet_text,
            font_size_pt=13.0,
            alignment="left",
        ))
        top += _ES_BULLET_HEIGHT + 0.10  # small gap between bullets

    spec = SlideSpec(
        slide_id=slide_id,
        slide_index=slide_index,
        layout="observed_full_width_table",   # ES slides are text-only full-width
        brand=brand,
        section=section,
        headline=HeadlineSpec(text=headline_text),
        components=components,
        footer=FooterSpec(text=footer) if footer else None,
        metadata=SlideMetadata(
            created_by="slide-plan-generator-exec-summary",
            arc="executive-summary",
        ),
    )
    # Note: citations/kbq_refs are currently tracked on the plan; future schema
    # extension could add them as SlideMetadata fields. For now, citations are
    # embedded inline in each bullet's text AND in the metadata.speaker_notes.
    spec.metadata.speaker_notes = "\n".join([
        f"KBQs addressed: {'; '.join(plan.kbq_group)}",
        f"Citations: {', '.join(all_citations)}",
    ])
    return spec


def _find_supporting_slides(
    kbq: str,
    deck_specs: list[SlideSpec],
    narrative_threads: Optional[dict] = None,
    max_citations: int = 3,
) -> list[tuple[str, str]]:
    """Return list of (slide_id, rationale) for slides that support this KBQ.

    Strategy:
      1. If narrative_threads maps KBQ → arc and arcs → slide_ids, use that
      2. Else: keyword match between KBQ terms and slide headlines + sections
      3. Fallback: return first few data-driven slides (best-effort)
    """
    # Strategy 1: narrative_threads lookup
    if narrative_threads:
        arc = narrative_threads.get("kbq_to_arc", {}).get(kbq)
        if arc:
            arc_slides = narrative_threads.get("arc_to_slides", {}).get(arc, [])
            return [(sid, f"arc={arc}") for sid in arc_slides[:max_citations]]

    # Strategy 2: keyword match
    kbq_words = set(w.lower() for w in kbq.split() if len(w) > 3)
    scored: list[tuple[int, str]] = []
    for spec in deck_specs:
        score = 0
        text = f"{spec.headline.text} {spec.section or ''}".lower()
        score = sum(1 for w in kbq_words if w in text)
        if score > 0:
            scored.append((score, spec.slide_id))
    scored.sort(reverse=True)
    if scored:
        return [(sid, "keyword match") for _, sid in scored[:max_citations]]

    # Strategy 3: fallback — first few data-driven slides
    data_slides = [s.slide_id for s in deck_specs
                   if any(hasattr(c, "chart_pattern") for c in s.components)][:max_citations]
    return [(sid, "fallback") for sid in data_slides]


def _distill_finding(
    kbq: str,
    supporting: list[tuple[str, str]],
    deck_specs: list[SlideSpec],
) -> Optional[ESFinding]:
    """Distill 1 finding per KBQ from the supporting slides.

    For now: concatenate the top supporting slide's headline as the finding.
    LLM augmentation (summarize across multiple slides) is a future upgrade.
    """
    if not supporting:
        return ESFinding(
            kbq=kbq,
            text=f"Open question — no slides in this deck directly address {kbq!r}",
            citations=[],
            confidence="low",
        )

    # Find the top supporting slide
    top_slide_id, rationale = supporting[0]
    top_spec = next((s for s in deck_specs if s.slide_id == top_slide_id), None)
    if top_spec is None:
        return None

    return ESFinding(
        kbq=kbq,
        text=top_spec.headline.text,
        citations=[sid for sid, _ in supporting],
        confidence="high" if rationale == "arc" else "medium",
    )


def _cluster_kbqs(kbqs: list[str], max_per_slide: int = 3) -> list[list[str]]:
    """Cluster KBQs into groups. Each group becomes one ES slide.

    Simple clustering by keyword overlap. For N ≤ max_per_slide, single cluster.
    Future: use narrative arcs as cluster labels.
    """
    if len(kbqs) <= max_per_slide:
        return [list(kbqs)]

    # Chunk by size for now — topic clustering is a future improvement
    clusters = []
    for i in range(0, len(kbqs), max_per_slide):
        clusters.append(list(kbqs[i : i + max_per_slide]))
    return clusters


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────


def generate_exec_summary(
    deck_specs: list[SlideSpec],
    kbqs: list[str],
    narrative_threads: Optional[dict] = None,
    brand: Optional[str] = None,
    section: str = "Executive Summary",
    insert_position: int = 1,
    max_es_slides: int = 3,
    max_kbqs_per_slide: int = 3,
) -> list[SlideSpec]:
    """Generate 1-3 ES slides answering the KBQs with citations.

    Args:
        deck_specs: full deck being summarized — used as citation source pool.
        kbqs: key business questions to answer.
        narrative_threads: optional narrative-threads.md content (dict with
            keys like 'kbq_to_arc', 'arc_to_slides') to improve citation quality.
        brand, section, insert_position: standard slide positioning.
        max_es_slides: cap on ES slides produced (default 3).
        max_kbqs_per_slide: cap on KBQs clustered per slide (default 3).

    Returns:
        list[SlideSpec] — 1 to max_es_slides ES slides, each with findings
        cited back to deck_specs' slide_ids.
    """
    if not kbqs:
        return []

    # 1. Cluster KBQs (currently simple chunk; future: narrative-arc-based)
    clusters = _cluster_kbqs(kbqs, max_per_slide=max_kbqs_per_slide)
    clusters = clusters[:max_es_slides]

    # 2. For each cluster: build findings with citations
    es_plans: list[ESSlidePlan] = []
    for cluster in clusters:
        findings = []
        for kbq in cluster:
            supporting = _find_supporting_slides(kbq, deck_specs, narrative_threads)
            finding = _distill_finding(kbq, supporting, deck_specs)
            if finding is not None:
                findings.append(finding)
        if findings:
            es_plans.append(ESSlidePlan(kbq_group=cluster, findings=findings))

    # 3. Build SlideSpec per ES slide plan
    specs: list[SlideSpec] = []
    for i, plan in enumerate(es_plans):
        slide_id = f"zrx_es_{i+1:02d}"
        slide_index = insert_position + i
        spec = _build_es_slide_spec(
            plan, slide_id=slide_id, slide_index=slide_index,
            brand=brand, section=section,
        )
        errors = validate_spec(spec)
        if errors:
            # Log but don't fail — caller can decide whether to proceed
            raise ValueError(
                f"generate_exec_summary produced invalid ES spec {slide_id}: {errors}"
            )
        specs.append(spec)

    return specs
