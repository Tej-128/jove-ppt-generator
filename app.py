"""
JoVE PPT Generator - Streamlit App V6.
Strict-formatting workflow with MP4-first visuals, approved AI fallback, and QA validation.
"""

import gc
import json
import os
import shutil
import tempfile
from pathlib import Path

import streamlit as st

from pipeline import run_pipeline

st.set_page_config(page_title="JoVE PPT Generator", page_icon="📊", layout="centered")

st.markdown("""
<style>
    .stProgress > div > div > div { background-color: #6D9EEB; }
    .flag-error { background: #fff0f0; border-left: 4px solid #FF465E; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
    .flag-warning { background: #fffbf0; border-left: 4px solid #FF9900; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
    .flag-review { background: #f0f4ff; border-left: 4px solid #6D9EEB; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
    .flag-info { background: #f0fff4; border-left: 4px solid #28a745; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)


def _render_flag(flag: dict):
    level = (flag.get("level") or "INFO").upper()
    cls = {
        "ERROR": "flag-error",
        "WARNING": "flag-warning",
        "REVIEW": "flag-review",
        "FORMATTING_REVIEW": "flag-review",
        "INFO": "flag-info",
    }.get(level, "flag-info")
    st.markdown(f'<div class="{cls}"><strong>{level}</strong>: {flag.get("message", "")}</div>', unsafe_allow_html=True)


def _stream_uploaded_file(uploaded_file, destination_path: str):
    uploaded_file.seek(0)
    with open(destination_path, "wb") as out:
        while True:
            chunk = uploaded_file.read(1024 * 1024)
            if not chunk:
                break
            out.write(chunk)


def _clear_previous_run():
    run_dir = st.session_state.get("run_dir")
    if run_dir and os.path.exists(run_dir):
        shutil.rmtree(run_dir, ignore_errors=True)
    for key in ["pptx_path", "pptx_name", "qa_path", "qa_report", "chapter_number", "run_dir"]:
        st.session_state.pop(key, None)
    gc.collect()


col1, col2 = st.columns([1, 5])
with col1:
    logo_path = os.path.join(os.path.dirname(__file__), "jove_logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=85)
with col2:
    st.title("JoVE PPT Generator")
    st.caption("V6 strict-formatting deck generation with validation.")

with st.expander("Required ZIP structure", expanded=True):
    st.markdown("""
Upload the DOCX files and MP4 files together.

Recommended naming:

```text
10649_Pagetext.docx
10649_Transcript.docx
10649_video.mp4
```

Also supported:

```text
10649_TheScientificMethod_Pagetext.docx
10649_TheScientificMethod_Transcript.docx
10649_TheScientificMethod.mp4
```

Rules:
- The MP4 filename must contain the same lesson ID.
- JoVE video frames are first priority.
- Approved AI fallback is used only if no suitable JoVE frame exists.
- No writer names, placeholders, Google images, or Wikipedia images.
""")

with st.form("jove_generator"):
    st.subheader("Chapter Details")
    zip_file = st.file_uploader("Upload Chapter ZIP", type=["zip"])
    chapter_name = st.text_input("Chapter Name", placeholder="e.g., Macromolecules")
    chapter_number = st.text_input("Chapter Number", placeholder="e.g., 3")

    lesson_order_input = st.text_area(
        "Lesson Order (optional)",
        placeholder="Enter lesson IDs in order, one per line:\n10649\n10650\n10651",
        height=120,
        help="Paste lesson IDs in the order the team wants. Leave blank to sort numerically."
    )

    total_slides_input = st.number_input(
        "Total Slide Count (optional)",
        min_value=0,
        max_value=200,
        value=0,
        step=1,
        help="Target total slides for this chapter. Leave 0 for automatic planning."
    )

    st.subheader("AI Settings")
    model = st.selectbox(
        "Text Model",
        ["gpt-4.1", "gpt-4o", "gpt-4.1-mini", "gpt-4o-mini", "gpt-5.5"],
        index=0,
        help="Used for planning and slide content generation."
    )

    vision_model = st.selectbox(
        "Vision Model",
        ["gpt-4.1", "gpt-4o", "gpt-4.1-mini", "gpt-4o-mini"],
        index=0,
        help="Used to pick the best frame from the lesson MP4."
    )

    submitted = st.form_submit_button("🚀 Generate Presentation", type="primary", use_container_width=True)


if submitted:
    _clear_previous_run()

    errors = []
    if not chapter_name.strip():
        errors.append("Chapter name is required.")
    if not chapter_number.strip():
        errors.append("Chapter number is required.")
    if not zip_file:
        errors.append("ZIP file is required.")

    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        errors.append("OPENAI_API_KEY is missing in Streamlit Secrets.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    run_dir = tempfile.mkdtemp(prefix="jove_streamlit_run_")
    st.session_state["run_dir"] = run_dir
    tmp_zip_path = os.path.join(run_dir, zip_file.name)
    _stream_uploaded_file(zip_file, tmp_zip_path)

    order_ids = [line.strip() for line in lesson_order_input.splitlines() if line.strip()]
    total_slide_budget = int(total_slides_input) if int(total_slides_input) > 0 else None

    progress_bar = st.progress(0)
    status = st.empty()

    def progress_callback(message, pct=None):
        status.info(message)
        if pct is not None:
            progress_bar.progress(min(100, max(0, int(pct))))

    try:
        out_path, qa_report = run_pipeline(
            zip_path=tmp_zip_path,
            chapter_name=chapter_name.strip(),
            chapter_number=chapter_number.strip(),
            openai_api_key=api_key,
            order_ids=order_ids,
            model=model,
            total_slide_budget=total_slide_budget,
            progress_callback=progress_callback,
            vision_model=vision_model
        )

        final_pptx_path = os.path.join(run_dir, Path(out_path).name)
        if os.path.abspath(out_path) != os.path.abspath(final_pptx_path):
            shutil.copy2(out_path, final_pptx_path)

        qa_path = os.path.join(run_dir, f"QA_Report_Chapter{chapter_number.strip()}.json")
        with open(qa_path, "w", encoding="utf-8") as f:
            json.dump(qa_report, f, indent=2)

        st.session_state["pptx_path"] = final_pptx_path
        st.session_state["pptx_name"] = Path(final_pptx_path).name
        st.session_state["qa_path"] = qa_path
        st.session_state["qa_report"] = qa_report
        st.session_state["chapter_number"] = chapter_number.strip()

        progress_bar.progress(100)
        status.success("Presentation generated successfully.")
        gc.collect()

    except Exception as e:
        progress_bar.progress(0)
        status.error("Generation failed.")
        st.error(str(e))
        st.stop()


if st.session_state.get("pptx_path") and os.path.exists(st.session_state["pptx_path"]):
    st.success("Ready to download.")

    with open(st.session_state["pptx_path"], "rb") as f:
        st.download_button(
            "📥 Download PPTX",
            data=f,
            file_name=st.session_state["pptx_name"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

    if st.session_state.get("qa_path") and os.path.exists(st.session_state["qa_path"]):
        with open(st.session_state["qa_path"], "rb") as f:
            st.download_button(
                "📥 Download QA Report",
                data=f,
                file_name=Path(st.session_state["qa_path"]).name,
                mime="application/json",
                use_container_width=True
            )

    qa_report = st.session_state.get("qa_report", {})
    validation = qa_report.get("formatting_validation", {})
    if validation:
        score = validation.get("formatting_score")
        target_met = validation.get("target_met")
        if target_met:
            st.success(f"Formatting validation score: {score}%")
        else:
            st.warning(f"Formatting validation score: {score}% — review QA findings before use.")

    with st.expander("QA Summary", expanded=True):
        st.write({
            "total_slides": qa_report.get("total_slides"),
            "lessons_processed": len(qa_report.get("lessons_processed", [])),
            "video_frames_used": len(qa_report.get("images_used", [])),
            "missing_images": len(qa_report.get("images_missing", [])),
            "formatting_score": validation.get("formatting_score"),
            "formatting_target_met": validation.get("target_met"),
            "vision_model": qa_report.get("vision_model"),
        })
        for flag in qa_report.get("flags", []):
            _render_flag(flag)
        st.json(qa_report)

    if st.button("Clear generated files"):
        _clear_previous_run()
        st.rerun()
