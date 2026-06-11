"""
JoVE PPT Generator — Streamlit App
Team-facing UI: upload ZIP → configure → generate → download.
"""

import streamlit as st
import json
import os
import tempfile
import time
from pathlib import Path

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

# ── Header ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 5])
with col1:
    logo_path = os.path.join(os.path.dirname(__file__), "assets", "jove_logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=80)
with col2:
    st.title("JoVE Lecture Slide Generator")
    st.caption("Upload a chapter ZIP → AI generates a pixel-perfect PPTX")

st.divider()

# ── Input Form ────────────────────────────────────────────────────────────────
with st.form("generator_form"):
    st.subheader("Chapter Configuration")

    col_a, col_b = st.columns([3, 1])
    with col_a:
        chapter_name = st.text_input(
            "Chapter Name",
            placeholder="e.g. The Chemistry of Life",
            help="This appears on the cover slide"
        )
    with col_b:
        chapter_number = st.text_input(
            "Chapter Number",
            placeholder="e.g. 21",
            help="e.g. 21"
        )

    st.subheader("Files")
    zip_file = st.file_uploader(
        "Chapter ZIP file",
        type=["zip"],
        help="ZIP containing all _Pagetext.docx and _Transcript.docx files"
    )

    order_input = st.text_area(
        "Lesson Order (optional)",
        placeholder="Enter lesson IDs in order, one per line:\n10655\n10656\n10657\n...",
        height=120,
        help="Paste lesson IDs in the order the team wants. Leave blank to sort numerically."
    )

    st.subheader("AI Settings")
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    model = st.selectbox(
        "Model",
        ["gpt-4.1", "gpt-4.1-mini", "gpt-4o", "gpt-4o-mini"],
        index=0,
        help="gpt-4.1 recommended. Use mini for faster/cheaper testing."
    )

    submitted = st.form_submit_button("🚀 Generate Presentation", type="primary",
                                       use_container_width=True)

# ── Processing ────────────────────────────────────────────────────────────────
if submitted:
    # Validation
    errors = []
    if not chapter_name.strip():
        errors.append("Chapter name is required")
    if not chapter_number.strip():
        errors.append("Chapter number is required")
    if not zip_file:
        errors.append("Please upload a chapter ZIP file")
    if not api_key:
        errors.append("OpenAI API key not configured in Streamlit secrets")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    # Parse order IDs
    order_ids = None
    if order_input.strip():
        order_ids = [line.strip() for line in order_input.strip().split('\n') if line.strip()]

    # Save uploaded ZIP to temp
    with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp:
        tmp.write(zip_file.read())
        tmp_zip_path = tmp.name

    # Progress tracking
    progress_bar = st.progress(0)
    status_text = st.empty()
    log_container = st.expander("Live log", expanded=True)
    log_lines = []

    def progress_callback(message, percent=None):
        log_lines.append(f"{'[' + str(percent) + '%]' if percent else '     '} {message}")
        with log_container:
            st.code('\n'.join(log_lines[-15:]), language=None)
        if percent is not None:
            progress_bar.progress(min(percent, 100) / 100)
        status_text.text(message)

    try:
        from pipeline import run_pipeline

        start_time = time.time()
        pptx_path, qa_report = run_pipeline(
            zip_path=tmp_zip_path,
            chapter_name=chapter_name.strip(),
            chapter_number=chapter_number.strip(),
            openai_api_key=api_key.strip(),
            order_ids=order_ids,
            model=model,
            progress_callback=progress_callback
        )
        elapsed = time.time() - start_time

        progress_bar.progress(1.0)
        status_text.text(f"✅ Done in {elapsed:.0f}s")

        st.success(f"Generated **{qa_report['total_slides']} slides** in {elapsed:.0f} seconds")

        # ── Downloads ─────────────────────────────────────────────────────────
        st.subheader("Downloads")
        col1, col2 = st.columns(2)

        with open(pptx_path, 'rb') as f:
            pptx_bytes = f.read()

        with col1:
            st.download_button(
                "📥 Download PPTX",
                data=pptx_bytes,
                file_name=Path(pptx_path).name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary"
            )

        qa_json = json.dumps(qa_report, indent=2)
        with col2:
            st.download_button(
                "📋 Download QA Report (JSON)",
                data=qa_json,
                file_name=f"QA_Report_Chapter{chapter_number}.json",
                mime="application/json",
                use_container_width=True
            )

        # ── QA Summary ────────────────────────────────────────────────────────
        st.subheader("QA Report Summary")

        # Stats
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        col_s1.metric("Total Slides", qa_report["total_slides"])
        col_s2.metric("Lessons Processed", len(qa_report["lessons_processed"]))
        col_s3.metric("Lessons Skipped", len(qa_report["lessons_skipped"]))
        img_ok = len(qa_report["images_used"])
        img_miss = len(qa_report["images_missing"])
        col_s4.metric("Images Found", f"{img_ok}/{img_ok+img_miss}")

        # Flags
        if qa_report["flags"]:
            st.write("**Flags for human reviewer:**")
            for flag in qa_report["flags"]:
                level = flag["level"].lower()
                css = f"flag-{level}"
                icon = {"error": "🔴", "warning": "🟡", "review": "🔵", "info": "🟢"}.get(level, "⚪")
                st.markdown(
                    f'<div class="{css}">{icon} <strong>{flag["level"]}</strong>: {flag["message"]}</div>',
                    unsafe_allow_html=True
                )
        else:
            st.success("✅ No flags — clean generation")

        # Lesson breakdown
        with st.expander("Lesson breakdown"):
            for lq in qa_report["lessons_processed"]:
                img_status = f"✅ {lq['images_found']} images" if lq['images_missing'] == 0 \
                    else f"⚠️ {lq['images_found']} found, {lq['images_missing']} missing"
                st.write(f"**{lq['name']}** — {lq['slides_built']} slides | {img_status}")

        if qa_report["lessons_skipped"]:
            with st.expander(f"⚠️ {len(qa_report['lessons_skipped'])} skipped lessons"):
                for s in qa_report["lessons_skipped"]:
                    st.write(f"• **{s['name']}** (ID {s['id']}): {s['reason']}")

        if qa_report.get("scientific_names"):
            with st.expander("🔬 Scientific names to verify"):
                for name in sorted(set(qa_report["scientific_names"])):
                    st.write(f"• *{name}*")

        if qa_report["images_missing"]:
            with st.expander(f"📷 {len(qa_report['images_missing'])} slides need images"):
                for img in qa_report["images_missing"]:
                    st.write(f"• **{img['lesson']}** ({img['slide_type']}): `{img['query']}`")

    except Exception as e:
        st.error(f"Generation failed: {str(e)}")
        st.exception(e)

    finally:
        try:
            os.unlink(tmp_zip_path)
        except Exception:
            pass

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("JoVE PPT Generator · Built with OpenAI GPT-4.1 + python-pptx · Images: Wikimedia Commons (CC)")
