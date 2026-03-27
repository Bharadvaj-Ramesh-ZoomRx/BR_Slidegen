"""
app.py — SlideGen Pipeline UI

A Streamlit web app that runs the full 6-stage slide generation pipeline.
Each stage streams Claude's output in real time. Users can edit the output,
send refinement instructions, and confirm before advancing to the next stage.

Run:
    cd galen-consulting-r3m-report
    streamlit run web/app.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import streamlit as st

# Ensure repo root is on the Python path
REPO_ROOT = Path(__file__).parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from web.project_manager import (
    list_projects,
    list_waves,
    get_project_dir,
    get_project_meta,
    get_stage_status,
    get_deck_path,
    get_input_file_status,
)
from web.stage_runner import (
    run_stage_stream,
    refine_stage_stream,
    confirm_stage,
    load_existing_output,
    list_checkpoints,
    load_checkpoint,
)
from web.skill_loader import STAGE_DEFS, check_missing_inputs
from web.deck_previewer import render_thumbnails, render_slide_hires, soffice_available

# pipeline imports (Stage 0 + Stage 5)
from slidegen.pipeline import generate_deck, index_excel


# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SlideGen Pipeline",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

_CSS = """
<style>
  .stage-header { font-size: 1.1rem; font-weight: 600; margin-bottom: 0; }
  .status-done   { color: #00b050; font-weight: 600; }
  .status-review { color: #f75824; font-weight: 600; }
  .status-running { color: #1f77b4; font-weight: 600; }
  .status-pending { color: #888; }
  .file-ok  { color: #00b050; }
  .file-missing { color: #cc0000; }
  .warn-box { background: #fff3cd; border-left: 4px solid #ffc107;
              padding: 0.5rem 0.75rem; border-radius: 4px; margin: 0.5rem 0; }
</style>
"""
st.markdown(_CSS, unsafe_allow_html=True)


# ── Session state init ────────────────────────────────────────────────────────

def _init_state():
    defaults = {
        "project": None,
        "wave": None,
        "api_key": os.environ.get("ANTHROPIC_API_KEY", ""),
        # Per-stage: output text, status, message history
        "stage_outputs": {},       # {stage_num: str}
        "stage_status": {},        # {stage_num: "pending"|"running"|"needs_review"|"confirmed"|"done"}
        "message_history": {1: [], 2: [], 3: [], 4: []},  # multi-turn history per stage
        # Stage 5
        "thumbnails": [],          # list of PNG bytes
        "selected_slide": None,    # int index for hires view
        "deck_generated": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ── Helpers (defined before sidebar so they can be called during render) ──────

def _load_existing_outputs(project: str, wave: str):
    """On project/wave selection, load any already-confirmed outputs from disk."""
    project_dir = get_project_dir(project)
    disk_status = get_stage_status(project, wave)
    for stage_num in [1, 2, 3, 4]:
        if disk_status.get(stage_num) == "done":
            content = load_existing_output(stage_num, project_dir, wave)
            if content:
                st.session_state.stage_outputs[stage_num] = content
                st.session_state.stage_status[stage_num] = "confirmed"
    if disk_status.get(0) == "done":
        st.session_state.stage_status[0] = "confirmed"
    if disk_status.get(5) == "done":
        st.session_state.stage_status[5] = "confirmed"
        st.session_state.deck_generated = True


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("SlideGen Pipeline")
    st.caption("YAML-driven PowerPoint generation")
    st.divider()

    # API key
    api_key_input = st.text_input(
        "Anthropic API Key",
        value=st.session_state.api_key,
        type="password",
        help="Used for Stage 1–4 Claude calls. Set ANTHROPIC_API_KEY env var to pre-fill.",
    )
    if api_key_input:
        st.session_state.api_key = api_key_input

    st.divider()

    # Project picker
    projects = list_projects()
    if not projects:
        st.warning("No projects found in projects/")
        st.stop()

    selected_project = st.selectbox(
        "Project",
        options=projects,
        index=projects.index(st.session_state.project) if st.session_state.project in projects else 0,
    )

    # Reset stage state on project change
    if selected_project != st.session_state.project:
        st.session_state.project = selected_project
        st.session_state.wave = None
        st.session_state.stage_outputs = {}
        st.session_state.stage_status = {}
        st.session_state.message_history = {1: [], 2: [], 3: [], 4: []}
        st.session_state.thumbnails = []
        st.session_state.deck_generated = False

    # Wave picker
    waves = list_waves(selected_project)
    if not waves:
        st.warning(f"No waves found in projects/{selected_project}/input/wave/")
        st.stop()

    selected_wave = st.selectbox(
        "Wave",
        options=waves,
        index=waves.index(st.session_state.wave) if st.session_state.wave in waves else 0,
    )

    # Reset stage state on wave change
    if selected_wave != st.session_state.wave:
        st.session_state.wave = selected_wave
        st.session_state.stage_outputs = {}
        st.session_state.stage_status = {}
        st.session_state.message_history = {1: [], 2: [], 3: [], 4: []}
        st.session_state.thumbnails = []
        st.session_state.deck_generated = False
        # Load any existing disk outputs into session state
        _load_existing_outputs(selected_project, selected_wave)

    st.divider()

    # Project meta
    meta = get_project_meta(selected_project)
    st.caption(f"**{meta['name']}**  \n{meta['client']}")

    # Input file checklist
    st.markdown("**Input files (Stage 1)**")
    file_status = get_input_file_status(selected_project, selected_wave)
    for label, exists in file_status.items():
        icon = "✅" if exists else "❌"
        st.caption(f"{icon} {label}")

    # soffice availability
    st.divider()
    if soffice_available():
        st.caption("✅ LibreOffice — slide preview enabled")
    else:
        st.caption("⚠️ LibreOffice not found — slide preview unavailable")


# ── Main content ──────────────────────────────────────────────────────────────

project = st.session_state.project
wave = st.session_state.wave
project_dir = get_project_dir(project)

st.title(f"📊 {get_project_meta(project)['name']} — {wave}")
st.caption("Complete each stage in order. Confirm output before advancing.")

# Load existing disk state on first render
if not st.session_state.stage_status:
    _load_existing_outputs(project, wave)


# ── Helper: stage status badge ────────────────────────────────────────────────

_STATUS_LABELS = {
    "pending":     ("⬜", "Pending"),
    "running":     ("🔄", "Running…"),
    "needs_review":("🟡", "Needs Review"),
    "confirmed":   ("✅", "Confirmed"),
    "done":        ("✅", "Done"),
    "error":       ("🔴", "Error"),
}

def _status_badge(stage_num: int) -> str:
    status = st.session_state.stage_status.get(stage_num, "pending")
    icon, label = _STATUS_LABELS.get(status, ("⬜", status))
    return f"{icon} {label}"


def _is_confirmed(stage_num: int) -> bool:
    return st.session_state.stage_status.get(stage_num) in ("confirmed", "done")


def _is_locked(stage_num: int) -> bool:
    """Stage N is locked until stage N-1 is confirmed."""
    if stage_num == 0:
        return False
    return not _is_confirmed(stage_num - 1)


# ── STAGE 0: Index Excel ──────────────────────────────────────────────────────

with st.expander(f"**Stage 0 — Index Excel** {_status_badge(0)}", expanded=not _is_confirmed(0)):
    st.caption("Converts source_data.xlsx → source_data.json for downstream stages. Runs automatically.")

    xlsx_path = project_dir / f"input/wave/{wave}/source_data.xlsx"
    json_path = project_dir / f"context/{wave}/source_data.json"

    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown(f"**Source:** `input/wave/{wave}/source_data.xlsx`")
        if xlsx_path.exists():
            st.markdown('<span class="file-ok">✅ File found</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="file-missing">❌ File not found — place source_data.xlsx in the input folder</span>', unsafe_allow_html=True)

        if json_path.exists():
            st.markdown(f'<span class="file-ok">✅ source_data.json already exists</span>', unsafe_allow_html=True)

    with col2:
        if xlsx_path.exists() and not _is_confirmed(0):
            if st.button("▶ Run Stage 0", key="run_0", type="primary"):
                with st.spinner("Indexing Excel…"):
                    try:
                        json_path.parent.mkdir(parents=True, exist_ok=True)
                        index_excel(str(xlsx_path), str(json_path))
                        st.session_state.stage_status[0] = "confirmed"
                        st.success("Done — source_data.json created")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed: {e}")
        elif _is_confirmed(0):
            st.success("Stage 0 complete")


# ── Stage panels 1–3 (AI-driven) ─────────────────────────────────────────────

def render_ai_stage(stage_num: int):
    """Render an AI-driven stage panel (1, 2, or 3)."""
    stage_def = STAGE_DEFS[stage_num]
    label = stage_def["label"]
    desc = stage_def["description"]
    status = st.session_state.stage_status.get(stage_num, "pending")
    locked = _is_locked(stage_num)

    with st.expander(
        f"**Stage {stage_num} — {label}** {_status_badge(stage_num)}",
        expanded=(status in ("needs_review",) or (not locked and status == "pending")),
    ):
        st.caption(desc)

        if locked:
            st.info(f"Complete Stage {stage_num - 1} first.")
            return

        if not st.session_state.api_key:
            st.warning("Enter your Anthropic API key in the sidebar.")
            return

        # Check missing inputs
        missing = check_missing_inputs(stage_num, project_dir, wave)
        if missing:
            st.markdown(
                f'<div class="warn-box">⚠️ Missing inputs: {", ".join(missing)}</div>',
                unsafe_allow_html=True,
            )

        # ── Run button ──
        if status == "pending":
            col_run, col_load = st.columns([1, 1])
            with col_run:
                if st.button(f"▶ Run Stage {stage_num}", key=f"run_{stage_num}", type="primary"):
                    _execute_stage(stage_num)
                    st.rerun()
            with col_load:
                # Option to load from disk if file exists
                output_path = project_dir / stage_def["output"].replace("{wave}", wave)
                if output_path.exists():
                    if st.button(f"📂 Load existing output", key=f"load_{stage_num}"):
                        content = output_path.read_text(encoding="utf-8")
                        st.session_state.stage_outputs[stage_num] = content
                        st.session_state.stage_status[stage_num] = "needs_review"
                        st.rerun()

        # ── Output editor ──
        if status in ("needs_review", "confirmed"):
            output_content = st.session_state.stage_outputs.get(stage_num, "")

            # Two-column layout: editor + stats
            col_editor, col_meta = st.columns([3, 1])

            with col_editor:
                edited = st.text_area(
                    "Output (editable):",
                    value=output_content,
                    height=500,
                    key=f"editor_{stage_num}",
                    help="Edit directly, or use the Refinement box below to ask AI to revise.",
                )
                # Update session state when user edits
                if edited != output_content:
                    st.session_state.stage_outputs[stage_num] = edited

            with col_meta:
                # Word / line count
                lines = output_content.count("\n") + 1
                words = len(output_content.split())
                st.metric("Lines", lines)
                st.metric("Words", words)

                # Checkpoint restore
                checkpoints = list_checkpoints(stage_num, project_dir, wave)
                if checkpoints:
                    st.caption(f"**{len(checkpoints)} checkpoint(s)**")
                    selected_ckpt = st.selectbox(
                        "Restore from checkpoint:",
                        options=["— select —"] + [c.name for c in checkpoints],
                        key=f"ckpt_{stage_num}",
                    )
                    if selected_ckpt != "— select —":
                        ckpt_path = next(c for c in checkpoints if c.name == selected_ckpt)
                        if st.button("Restore", key=f"restore_{stage_num}"):
                            st.session_state.stage_outputs[stage_num] = load_checkpoint(ckpt_path)
                            st.session_state.stage_status[stage_num] = "needs_review"
                            st.rerun()

            # ── Refinement ──
            st.caption("**Refinement** — describe what to change and AI will revise the output:")
            ref_col1, ref_col2 = st.columns([4, 1])
            with ref_col1:
                refinement = st.text_input(
                    label="Refinement instruction",
                    label_visibility="collapsed",
                    placeholder='e.g. "Add more detail to the Methodology Notes section"',
                    key=f"refine_input_{stage_num}",
                )
            with ref_col2:
                if st.button("Refine ↩", key=f"refine_btn_{stage_num}"):
                    if refinement:
                        _execute_refinement(stage_num, refinement)
                        st.rerun()

            # ── Confirm / Re-run ──
            st.divider()
            btn_col1, btn_col2 = st.columns([2, 1])
            with btn_col1:
                if status == "needs_review":
                    if st.button(
                        f"✅ Confirm & Advance to Stage {stage_num + 1}",
                        key=f"confirm_{stage_num}",
                        type="primary",
                    ):
                        final_content = st.session_state.stage_outputs.get(stage_num, "")
                        confirm_stage(stage_num, final_content, project_dir, wave)
                        st.session_state.stage_status[stage_num] = "confirmed"
                        st.success(f"Stage {stage_num} confirmed — output saved to disk.")
                        st.rerun()
                else:
                    st.success(f"Stage {stage_num} confirmed.")
            with btn_col2:
                if st.button(f"↺ Re-run", key=f"rerun_{stage_num}"):
                    st.session_state.stage_status[stage_num] = "pending"
                    st.session_state.message_history[stage_num] = []
                    st.rerun()


def _execute_stage(stage_num: int):
    """Stream Stage N output into session state. Runs synchronously."""
    st.session_state.stage_status[stage_num] = "running"
    placeholder = st.empty()
    full_text = ""
    try:
        history = st.session_state.message_history.get(stage_num, [])
        # Rebuild initial user message for history tracking
        from web.skill_loader import build_stage_prompt
        _, user_msg = build_stage_prompt(stage_num, project_dir, wave)

        for chunk in run_stage_stream(
            stage_num, project_dir, wave,
            st.session_state.api_key,
            history,
        ):
            full_text += chunk
            placeholder.markdown(full_text + "▌")

        placeholder.markdown(full_text)
        st.session_state.stage_outputs[stage_num] = full_text
        st.session_state.stage_status[stage_num] = "needs_review"

        # Update message history
        st.session_state.message_history[stage_num] = [
            {"role": "user", "content": user_msg},
            {"role": "assistant", "content": full_text},
        ]

    except Exception as e:
        placeholder.error(f"Stage {stage_num} failed: {e}")
        st.session_state.stage_status[stage_num] = "error"


def _execute_refinement(stage_num: int, instruction: str):
    """Stream a refinement response into session state."""
    placeholder = st.empty()
    full_text = ""
    try:
        history = st.session_state.message_history.get(stage_num, [])

        for chunk in refine_stage_stream(
            instruction, stage_num, project_dir, wave,
            st.session_state.api_key,
            history,
        ):
            full_text += chunk
            placeholder.markdown(full_text + "▌")

        placeholder.markdown(full_text)
        st.session_state.stage_outputs[stage_num] = full_text
        st.session_state.stage_status[stage_num] = "needs_review"

        # Extend history
        st.session_state.message_history[stage_num] += [
            {"role": "user", "content": instruction},
            {"role": "assistant", "content": full_text},
        ]

    except Exception as e:
        placeholder.error(f"Refinement failed: {e}")


# Render AI stages
for sn in [1, 2, 3]:
    render_ai_stage(sn)


# ── STAGE 4: Generate Config YAML ─────────────────────────────────────────────

with st.expander(
    f"**Stage 4 — Generate Config YAML** {_status_badge(4)}",
    expanded=not _is_confirmed(3) is False and not _is_confirmed(4),
):
    st.caption(STAGE_DEFS[4]["description"])

    if _is_locked(4):
        st.info("Complete Stage 3 first.")
    elif not st.session_state.api_key:
        st.warning("Enter your Anthropic API key in the sidebar.")
    else:
        status4 = st.session_state.stage_status.get(4, "pending")

        if status4 == "pending":
            c1, c2 = st.columns([1, 1])
            with c1:
                if st.button("▶ Generate Config", key="run_4", type="primary"):
                    _execute_stage(4)
                    st.rerun()
            with c2:
                config_path = project_dir / "config.yaml"
                if config_path.exists():
                    if st.button("📂 Load current config.yaml", key="load_4"):
                        content = config_path.read_text(encoding="utf-8")
                        st.session_state.stage_outputs[4] = content
                        st.session_state.stage_status[4] = "needs_review"
                        st.rerun()

        if status4 in ("needs_review", "confirmed"):
            output4 = st.session_state.stage_outputs.get(4, "")
            edited4 = st.text_area(
                "config.yaml (editable):",
                value=output4,
                height=600,
                key="editor_4",
                help="Review and edit the generated YAML before confirming.",
            )
            if edited4 != output4:
                st.session_state.stage_outputs[4] = edited4

            ref4_col1, ref4_col2 = st.columns([4, 1])
            with ref4_col1:
                ref4 = st.text_input(
                    label="Refinement instruction",
                    label_visibility="collapsed",
                    placeholder='e.g. "Add extraction for Q2.20 message effectiveness"',
                    key="refine_input_4",
                )
            with ref4_col2:
                if st.button("Refine ↩", key="refine_btn_4"):
                    if ref4:
                        _execute_refinement(4, ref4)
                        st.rerun()

            st.divider()
            if status4 == "needs_review":
                if st.button("✅ Confirm Config & Advance to Stage 5", key="confirm_4", type="primary"):
                    final4 = st.session_state.stage_outputs.get(4, "")
                    confirm_stage(4, final4, project_dir, wave)
                    st.session_state.stage_status[4] = "confirmed"
                    st.success("config.yaml saved.")
                    st.rerun()
            else:
                st.success("Stage 4 confirmed — config.yaml saved.")
                if st.button("↺ Re-run Config Generation", key="rerun_4"):
                    st.session_state.stage_status[4] = "pending"
                    st.session_state.message_history[4] = []
                    st.rerun()


# ── STAGE 5: Generate Deck ────────────────────────────────────────────────────

with st.expander(
    f"**Stage 5 — Generate Deck** {_status_badge(5)}",
    expanded=_is_confirmed(4) and not _is_confirmed(5),
):
    st.caption("Runs generate_deck() → renders deck.pptx → shows slide previews")

    if _is_locked(5):
        st.info("Complete Stage 4 first.")
    else:
        config_path = project_dir / "config.yaml"
        deck_path = get_deck_path(project, wave)

        top_col1, top_col2 = st.columns([2, 1])

        with top_col1:
            if st.button("▶ Generate Deck", key="run_5", type="primary", disabled=not config_path.exists()):
                with st.spinner("Running generate_deck()… this may take 30–60 seconds"):
                    try:
                        # Run pipeline from the project directory
                        import os as _os
                        _orig_cwd = _os.getcwd()
                        _os.chdir(str(project_dir))
                        try:
                            generate_deck("config.yaml")
                        finally:
                            _os.chdir(_orig_cwd)

                        st.session_state.stage_status[5] = "confirmed"
                        st.session_state.deck_generated = True
                        st.session_state.thumbnails = []  # force re-render
                        st.success("Deck generated!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"generate_deck() failed: {e}")

        with top_col2:
            if deck_path:
                with open(deck_path, "rb") as f:
                    st.download_button(
                        "⬇ Download deck.pptx",
                        data=f,
                        file_name=f"{project}_{wave}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="download_deck",
                    )

        # ── Slide thumbnails ──
        if deck_path:
            if not st.session_state.thumbnails:
                with st.spinner("Rendering slide previews…"):
                    st.session_state.thumbnails = render_thumbnails(str(deck_path))

            thumbnails = st.session_state.thumbnails

            if thumbnails:
                st.caption(
                    f"**{len(thumbnails)} slides** — click a slide number to view full size.  \n"
                    "⚠️ Preview uses system fonts. Open in PowerPoint for correct rendering."
                )

                # Show thumbnails in a grid of 4 columns
                cols_per_row = 4
                for row_start in range(0, len(thumbnails), cols_per_row):
                    cols = st.columns(cols_per_row)
                    for col_idx, col in enumerate(cols):
                        slide_idx = row_start + col_idx
                        if slide_idx >= len(thumbnails):
                            break
                        with col:
                            st.image(thumbnails[slide_idx], use_container_width=True)
                            c1, c2 = st.columns(2)
                            with c1:
                                st.caption(f"Slide {slide_idx + 1}")
                            with c2:
                                if st.button("🔍", key=f"view_{slide_idx}", help="View full size"):
                                    st.session_state.selected_slide = slide_idx
                                    st.rerun()

                # ── Full-size slide viewer ──
                sel = st.session_state.selected_slide
                if sel is not None and deck_path:
                    st.divider()
                    st.subheader(f"Slide {sel + 1} — Full Resolution")
                    hires = render_slide_hires(str(deck_path), sel)
                    st.image(hires, use_container_width=True)

                    rcol1, rcol2, rcol3 = st.columns([1, 1, 2])
                    with rcol1:
                        if st.button("← Prev", key="prev_slide", disabled=(sel == 0)):
                            st.session_state.selected_slide = sel - 1
                            st.rerun()
                    with rcol2:
                        if st.button("Next →", key="next_slide", disabled=(sel >= len(thumbnails) - 1)):
                            st.session_state.selected_slide = sel + 1
                            st.rerun()
                    with rcol3:
                        if st.button("↺ Regenerate this slide", key="regen_slide"):
                            with st.spinner(f"Regenerating slide {sel + 1}…"):
                                try:
                                    from slidegen.pipeline import regenerate_slide
                                    import os as _os
                                    _orig = _os.getcwd()
                                    _os.chdir(str(project_dir))
                                    try:
                                        regenerate_slide("config.yaml", slide_index=sel)
                                    finally:
                                        _os.chdir(_orig)
                                    # Re-render thumbnails
                                    st.session_state.thumbnails = render_thumbnails(str(deck_path))
                                    st.success(f"Slide {sel + 1} regenerated.")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Regeneration failed: {e}")

            else:
                if not soffice_available():
                    st.markdown(
                        '<div class="warn-box">'
                        "Slide preview requires LibreOffice. Install from libreoffice.org and ensure "
                        "<code>soffice</code> is on your PATH, then refresh."
                        "</div>",
                        unsafe_allow_html=True,
                    )
                else:
                    st.info("No slides to preview.")
