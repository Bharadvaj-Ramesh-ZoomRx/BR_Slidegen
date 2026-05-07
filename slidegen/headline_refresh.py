"""Backward-compat shim — module renamed to headliner_full_workflow.

The talking-header rewriter lives in
``slidegen.headliner_full_workflow`` as of 2026-05-08. This module
re-exports the same symbols so existing callers keep working without
changes:

  - ``slidegen.refresh_pipeline`` (Step 3 of the pipeline)
  - ``scripts/refresh_*_today.py`` (per-deck convenience runners)
  - ``scripts/probe_*headline*.py``
  - ``scripts/compute_headline_status.py``
  - ``tests/connector/test_headline_refresh.py``
  - ``tests/evals/end_to_end/test_step7_headlines.py``

New code should import from ``slidegen.headliner_full_workflow``
directly.
"""
from slidegen.headliner_full_workflow import (  # noqa: F401
    HeadlineUpdate,
    _all_data_shapes,
    _all_none,
    _build_prompt,
    _call_claude,
    _find_headline_shape,
    _largest_chart,
    _looks_like_data_narrative,
    _summarize_chart,
    _summarize_table,
    _values_eq,
    refresh_headlines,
)
