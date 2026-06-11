"""
JoVE PPT Pipeline
Orchestrates the full waterfall: ZIP → parse → AI → PPTX + QA report.
"""

import os
import re
import zipfile
import tempfile
import json
from pathlib import Path
from datetime import datetime
from docx import Document as DocxDocument

from ai_generator import generate_slide_content, calculate_slide_budget, search_wikimedia_image
from ppt_builder import (
    create_presentation, build_cover_slide, build_concept_slide,
    build_table_slide, build_discussion_question_slide,
    build_discussion_answer_slide, build_summary_slide, build_glossary_slide
)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "jove_logo.png")


def _camel_to_title(name: str) -> str:
    """'ThePeriodicTableAndOrganismalElements' → 'The Periodic Table And Organismal Elements'"""
    name = name.replace("_", " ")
    spaced = re.sub(r'([A-Z])', r' \1', name).strip()
    # Clean double spaces
    return re.sub(r'\s+', ' ', spaced)


def _read_docx(path: str) -> str:
    """Extract plain text from a .docx file."""
    try:
        doc = DocxDocument(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        return ""


def parse_chapter_zip(zip_path: str, order_ids: list = None) -> list:
    """
    Parse chapter ZIP, pair Pagetext + Transcript by LessonID.
    Returns ordered list of lesson dicts.
    If order_ids provided, sort by that. Otherwise sort numerically.
    """
    lessons = {}

    with zipfile.ZipFile(zip_path, 'r') as z:
        tmpdir = tempfile.mkdtemp()
        z.extractall(tmpdir)

    # Walk extracted files
    for root, dirs, files in os.walk(tmpdir):
        for fname in files:
            if not fname.endswith('.docx'):
                continue
            m = re.match(r'^(\d+)_(.+?)_(Pagetext|Transcript)\.docx$', fname, re.IGNORECASE)
            if not m:
                continue

            lesson_id = m.group(1)
            raw_name = m.group(2)
            doc_type = m.group(3).lower()
            lesson_name = _camel_to_title(raw_name)
            full_path = os.path.join(root, fname)
            content = _read_docx(full_path)

            if lesson_id not in lessons:
                lessons[lesson_id] = {
                    "id": lesson_id,
                    "name": lesson_name,
                    "pagetext": "",
                    "transcript": "",
                    "has_pagetext": False,
                    "has_transcript": False,
                    "is_stub": False
                }

            if doc_type == 'pagetext':
                lessons[lesson_id]['pagetext'] = content
                lessons[lesson_id]['has_pagetext'] = bool(content.strip())
            elif doc_type == 'transcript':
                lessons[lesson_id]['transcript'] = content
                lessons[lesson_id]['has_transcript'] = bool(content.strip())

    # Mark stubs
    for lid, lesson in lessons.items():
        if not lesson['pagetext'].strip() and not lesson['transcript'].strip():
            lesson['is_stub'] = True

    # Sort by order_ids if provided, else numeric
    all_ids = list(lessons.keys())
    if order_ids:
        ordered = [lessons[str(oid)] for oid in order_ids if str(oid) in lessons]
        # Append any IDs not in order list
        listed = set(str(oid) for oid in order_ids)
        ordered += [lessons[lid] for lid in sorted(all_ids, key=int) if lid not in listed]
    else:
        ordered = [lessons[lid] for lid in sorted(all_ids, key=int)]

    return ordered


def run_pipeline(zip_path: str, chapter_name: str, chapter_number: str,
                 openai_api_key: str, order_ids: list = None,
                 model: str = "gpt-4.1",
                 progress_callback=None) -> tuple[str, dict]:
    """
    Full pipeline. Returns (output_pptx_path, qa_report_dict).
    progress_callback(message, percent) for Streamlit UI updates.
    """

    def progress(msg, pct=None):
        print(f"[{pct or '?':>3}%] {msg}")
        if progress_callback:
            progress_callback(msg, pct)

    qa_report = {
        "chapter": chapter_name,
        "chapter_number": chapter_number,
        "generated_at": datetime.now().isoformat(),
        "model": model,
        "total_slides": 0,
        "lessons_processed": [],
        "lessons_skipped": [],
        "images_used": [],
        "images_missing": [],
        "scientific_names": [],
        "flags": []
    }

    progress("Parsing chapter ZIP...", 2)
    lessons = parse_chapter_zip(zip_path, order_ids)
    populated = [l for l in lessons if not l['is_stub']]
    stubs = [l for l in lessons if l['is_stub']]

    progress(f"Found {len(populated)} lessons, {len(stubs)} stubs", 5)

    if stubs:
        for s in stubs:
            qa_report["lessons_skipped"].append({
                "id": s["id"], "name": s["name"], "reason": "Empty stub — no content in either file"
            })
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"{len(stubs)} lessons have empty content files: {', '.join(s['name'] for s in stubs)}"
        })

    # Calculate word counts for adaptive budgeting
    word_counts = [len((l['transcript'] + l['pagetext']).split()) for l in populated]
    avg_words = sum(word_counts) / len(word_counts) if word_counts else 200

    # Init presentation
    prs = create_presentation(LOGO_PATH)
    all_glossary = {}
    slide_count = 0

    # ── Cover slide ──────────────────────────────────────────────────────────
    progress("Building cover slide...", 8)
    build_cover_slide(prs, chapter_name, chapter_number, LOGO_PATH)
    slide_count += 1

    # ── Lesson loop (waterfall) ───────────────────────────────────────────────
    for i, lesson in enumerate(populated):
        pct = 10 + int((i / len(populated)) * 70)
        progress(f"Processing lesson {i+1}/{len(populated)}: {lesson['name']}", pct)

        lesson_word_count = len((lesson['transcript'] + lesson['pagetext']).split())
        budget = calculate_slide_budget(len(populated), lesson_word_count, avg_words)

        lesson_qa = {
            "id": lesson["id"],
            "name": lesson["name"],
            "word_count": lesson_word_count,
            "slide_budget": budget,
            "slides_built": 0,
            "images_found": 0,
            "images_missing": 0
        }

        try:
            # Call AI
            slide_data = generate_slide_content(
                lesson_name=lesson['name'],
                transcript=lesson['transcript'],
                pagetext=lesson['pagetext'],
                slide_budget=budget,
                api_key=openai_api_key,
                model=model
            )
        except Exception as e:
            qa_report["flags"].append({
                "level": "ERROR",
                "message": f"AI generation failed for lesson '{lesson['name']}': {str(e)}"
            })
            qa_report["lessons_skipped"].append({
                "id": lesson["id"], "name": lesson["name"],
                "reason": f"AI generation error: {str(e)}"
            })
            continue

        # Merge glossary
        if "glossary_terms" in slide_data:
            all_glossary.update(slide_data["glossary_terms"])

        # Detect scientific names for QA
        sci_names = re.findall(r'\*\*_([A-Z][a-z]+ [a-z]+)_\*\*',
                               json.dumps(slide_data))
        if sci_names:
            qa_report["scientific_names"].extend(sci_names)

        # ── Build each slide ─────────────────────────────────────────────────
        for slide_def in slide_data.get("slides", []):
            stype = slide_def.get("type", "concept")
            title = slide_def.get("title", lesson["name"])
            notes = slide_def.get("speaker_notes", "")
            img_query = slide_def.get("image_query", lesson["name"] + " biology")

            # Fetch image
            img_url = search_wikimedia_image(img_query)
            if img_url:
                lesson_qa["images_found"] += 1
                qa_report["images_used"].append({
                    "lesson": lesson["name"], "slide_type": stype,
                    "query": img_query, "url": img_url,
                    "license": "Wikimedia Commons (CC / Public Domain)"
                })
            else:
                lesson_qa["images_missing"] += 1
                qa_report["images_missing"].append({
                    "lesson": lesson["name"], "slide_type": stype,
                    "query": img_query
                })

            if stype == "concept":
                build_concept_slide(
                    prs,
                    lesson_name=title,
                    body_text=slide_def.get("body", ""),
                    sub_label=slide_def.get("sub_label"),
                    image_url=img_url,
                    speaker_notes=notes,
                    logo_path=LOGO_PATH
                )

            elif stype == "table":
                build_table_slide(
                    prs,
                    lesson_name=title,
                    headers=slide_def.get("headers", []),
                    rows=slide_def.get("rows", []),
                    sub_title=slide_def.get("sub_title"),
                    image_url=img_url,
                    speaker_notes=notes,
                    logo_path=LOGO_PATH
                )

            elif stype == "discussion_question":
                build_discussion_question_slide(
                    prs,
                    lesson_name=title,
                    question_text=slide_def.get("question", ""),
                    hint_text=slide_def.get("hint"),
                    image_url=img_url,
                    speaker_notes=notes,
                    logo_path=LOGO_PATH
                )

            elif stype == "discussion_answer":
                build_discussion_answer_slide(
                    prs,
                    lesson_name=title,
                    answer_summary=slide_def.get("answer_summary", ""),
                    answer_explanation=slide_def.get("answer_explanation", ""),
                    image_url=img_url,
                    speaker_notes=notes,
                    logo_path=LOGO_PATH
                )

            elif stype == "summary":
                build_summary_slide(
                    prs,
                    summary_statement=slide_def.get("summary_statement", ""),
                    table_headers=slide_def.get("headers", []),
                    table_rows=slide_def.get("rows", []),
                    logo_path=LOGO_PATH,
                    speaker_notes=notes
                )

            slide_count += 1
            lesson_qa["slides_built"] += 1

        qa_report["lessons_processed"].append(lesson_qa)

    # ── Final Chapter Summary ────────────────────────────────────────────────
    progress("Building final summary...", 83)
    lesson_names = [l["name"] for l in populated]
    summary_rows = [[l["name"], f"Covered in {qa_report['lessons_processed'][i]['slides_built']} slides"]
                    for i, l in enumerate(populated)
                    if i < len(qa_report["lessons_processed"])]

    build_summary_slide(
        prs,
        summary_statement=f"{chapter_name} — chapter complete.",
        table_headers=["Lesson", "Coverage"],
        table_rows=summary_rows[:12],  # max 12 rows
        logo_path=LOGO_PATH,
        speaker_notes="This slide summarizes all lessons covered in this chapter."
    )
    slide_count += 1

    # ── Glossary (split into pages of 5 terms each) ─────────────────────────
    progress("Building glossary...", 88)
    TERMS_PER_PAGE = 5
    gloss_items = list(all_glossary.items())
    for page_start in range(0, max(1, len(gloss_items)), TERMS_PER_PAGE):
        page_terms = dict(gloss_items[page_start:page_start + TERMS_PER_PAGE])
        build_glossary_slide(prs, page_terms, logo_path=LOGO_PATH)
        slide_count += 1

    # ── Save PPTX ────────────────────────────────────────────────────────────
    progress("Saving PPTX...", 94)
    safe_name = re.sub(r'[^\w\s-]', '', chapter_name).replace(' ', '_')
    out_path = os.path.join(tempfile.gettempdir(), f"JoVE_Chapter{chapter_number}_{safe_name}.pptx")
    prs.save(out_path)

    qa_report["total_slides"] = slide_count
    qa_report["output_file"] = out_path

    # QA flags
    missing_count = len(qa_report["images_missing"])
    if missing_count > 0:
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"{missing_count} slides have placeholder images. See images_missing list."
        })
    if qa_report["scientific_names"]:
        qa_report["flags"].append({
            "level": "REVIEW",
            "message": f"Scientific names detected — verify spelling: {', '.join(set(qa_report['scientific_names']))}"
        })

    progress(f"Done! {slide_count} slides generated.", 100)
    return out_path, qa_report
