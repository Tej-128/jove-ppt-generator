"""
JoVE PPT Generator — Streamlit App
Team-facing UI: upload ZIP -> configure -> generate -> download.

Updated for strict MP4 frame workflow:
- MP4s must be included in the ZIP.
- Every image-bearing slide uses a frame from that lesson's MP4.
- No web fallback and no placeholders.
"""

import streamlit as st
import json
import os
import tempfile
from pathlib import Path

from pipeline import run_pipeline

st.set_page_config(
    page_title="JoVE PPT Generator",
    page_icon="📊",
    layout="centered"
)

st.markdown("""
<style>
    .stProgress > div > div > div { background-color: #4A86E8; }
    .flag-error { background: #fff0f0; border-left: 4px solid #FF465E; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
    .flag-warning { background: #fffbf0; border-left: 4px solid #F0A500; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
    .flag-review { background: #f0f4ff; border-left: 4px solid #4A86E8; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
    .flag-info { background: #f0fff4; border-left: 4px solid #28a745; padding: 8px 12px; margin: 4px 0; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)


def _render_flag(flag: dict):
    level = (flag.get("level") or "INFO").upper()
    cls = {
        "ERROR": "flag-error",
        "WARNING": "flag-warning",
        "REVIEW": "flag-review",
        "INFO": "flag-info",
    }.get(level, "flag-info")
    st.markdown(
        f'<div class="{cls}"><strong>{level}</strong>: {flag.get("message", "")}</div>',
        unsafe_allow_html=True
    )


# Header
col1, col2 = st.columns([1, 5])
with col1:
    logo_path = os.path.join(os.path.dirname(__file__), "jove_logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=85)
with col2:
    st.title("JoVE PPT Generator")
    st.caption("Transcript-first slide generation with OpenAI Vision MP4 frame selection.")

with st.expander("Required ZIP structure", expanded=True):
    st.markdown("""
Upload the DOCX files and MP4 files together.

Recommended naming:

```text
10649_TheScientificMethod_Transcript.docx
10649_TheScientificMethod_Pagetext.docx
10649_TheScientificMethod.mp4
```

Rules:
- The MP4 filename must contain the same lesson ID.
- MP4s can be at the same level as DOCXs. The app also searches subfolders if needed.
- Every image-bearing lesson slide uses a selected frame from the matching lesson MP4.
- No placeholders.
- No Google/Wikimedia fallback.
""")

with st.form("jove_generator"):
    st.subheader("Chapter Details")
    zip_file = st.file_uploader("Upload Chapter ZIP", type=["zip"])
    chapter_name = st.text_input("Chapter Name", placeholder="e.g., Meiosis")
    chapter_number = st.text_input("Chapter Number", placeholder="e.g., 11")

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
        help="Target total slides for this chapter. Includes cover, summary, glossary, and all lesson slides."
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
    # Clear previous results
    st.session_state.pop("pptx_bytes", None)
    st.session_state.pop("pptx_name", None)
    st.session_state.pop("qa_report", None)
    st.session_state.pop("chapter_number", None)

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

    order_ids = [line.strip() for line in lesson_order_input.splitlines() if line.strip()]
    total_slide_budget = int(total_slides_input) if int(total_slides_input) > 0 else None

    tmp_zip_path = os.path.join(tempfile.gettempdir(), zip_file.name)
    with open(tmp_zip_path, "wb") as f:
        f.write(zip_file.getbuffer())

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

        with open(out_path, "rb") as f:
            pptx_bytes = f.read()

        st.session_state["pptx_bytes"] = pptx_bytes
        st.session_state["pptx_name"] = Path(out_path).name
        st.session_state["qa_report"] = qa_report
        st.session_state["chapter_number"] = chapter_number.strip()

        progress_bar.progress(100)
        status.success("Presentation generated successfully.")

    except Exception as e:
        progress_bar.progress(0)
        status.error("Generation failed.")
        st.error(str(e))
        st.stop()


if "pptx_bytes" in st.session_state:
    st.success("Ready to download.")

    st.download_button(
        "📥 Download PPTX",
        data=st.session_state["pptx_bytes"],
        file_name=st.session_state["pptx_name"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )

    qa_report = st.session_state.get("qa_report", {})
    qa_bytes = json.dumps(qa_report, indent=2).encode("utf-8")

    st.download_button(
        "📥 Download QA Report",
        data=qa_bytes,
        file_name=f"QA_Report_Chapter{st.session_state.get('chapter_number', '')}.json",
        mime="application/json",
        use_container_width=True
    )

    with st.expander("QA Summary", expanded=True):
        st.write({
            "total_slides": qa_report.get("total_slides"),
            "lessons_processed": len(qa_report.get("lessons_processed", [])),
            "video_frames_used": len(qa_report.get("images_used", [])),
            "missing_images": len(qa_report.get("images_missing", [])),
            "vision_model": qa_report.get("vision_model"),
        })

        for flag in qa_report.get("flags", []):
            _render_flag(flag)

        st.json(qa_report)
