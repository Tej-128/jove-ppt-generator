"""
JoVE PPT Pipeline v3
Architecture:
parse ZIP -> planning pass -> AI slide generation -> transcript-aligned MP4 frame selection ->
presentation build -> QA report.

Strict image rule:
No placeholders and no web fallback. Every image-bearing lesson slide receives a frame
from that lesson's MP4. If an MP4 is missing/unreadable, generation stops.
"""

import os
import re
import zipfile
import tempfile
import json
from datetime import datetime
from typing import Dict, List, Optional

from docx import Document as DocxDocument

from ai_generator import generate_slide_content, generate_chapter_summary
from planner import plan_chapter_slides, default_chapter_budget
from video_sourcing import assign_frames_to_slides
from ppt_builder import (
    create_presentation, build_cover_slide, build_concept_slide,
    build_table_slide, build_discussion_question_slide,
    build_discussion_answer_slide, build_summary_slide, build_glossary_slide
)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "jove_logo.png")


def _camel_to_title(name: str) -> str:
    name = name.replace("_", " ")
    spaced = re.sub(r'([A-Z])', r' \1', name).strip()
    return re.sub(r'\s+', ' ', spaced)


def _norm_for_match(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(text or "").lower())


def _read_docx(path: str) -> str:
    try:
        doc = DocxDocument(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception:
        return ""


def _match_video_by_name(lesson_id: str, lesson_name: str, all_video_paths: List[str]) -> Optional[str]:
    """
    User rule:
    MP4s are expected at the same level as DOCXs, but this safely searches everywhere
    inside the ZIP in case a folder slips in.

    Matching:
    1. filename starts with lesson ID
    2. filename contains lesson ID
    3. filename contains normalized lesson name
    """
    # Starts with ID: 10649.mp4 / 10649_TheScientificMethod.mp4 / 10649 TheScientificMethod.mp4
    for path in all_video_paths:
        fname = os.path.basename(path)
        if re.match(rf"^{re.escape(str(lesson_id))}(?:[_\s-].*)?\.mp4$", fname, re.IGNORECASE):
            return path

    # Contains ID anywhere
    for path in all_video_paths:
        if str(lesson_id) in os.path.basename(path):
            return path

    # Lesson name fallback
    lesson_norm = _norm_for_match(lesson_name)
    for path in all_video_paths:
        fname_norm = _norm_for_match(os.path.basename(path))
        if lesson_norm and lesson_norm in fname_norm:
            return path

    return None


def parse_chapter_zip(zip_path: str, order_ids: list = None) -> list:
    lessons: Dict[str, Dict] = {}
    tmpdir = tempfile.mkdtemp(prefix="jove_chapter_")

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(tmpdir)

    thumbnails = {}
    videos = []

    for root, dirs, files in os.walk(tmpdir):
        for fname in sorted(files):
            lower = fname.lower()
            full_path = os.path.join(root, fname)

            if lower.endswith(('.jpg', '.jpeg', '.png')):
                m_img = re.match(r'^(\d+)(?:[_\s-].*)?\.(jpg|jpeg|png)$', fname, re.IGNORECASE)
                if m_img:
                    thumbnails[m_img.group(1)] = full_path
                continue

            if lower.endswith('.mp4'):
                videos.append(full_path)
                continue

            if not lower.endswith('.docx'):
                continue

            m = re.match(r'^(\d+)_(.+?)_(Pagetext|Transcript)\.docx$', fname, re.IGNORECASE)
            if not m:
                continue

            lesson_id = m.group(1)
            raw_name = m.group(2)
            doc_type = m.group(3).lower()
            lesson_name = _camel_to_title(raw_name)
            content = _read_docx(full_path)

            if lesson_id not in lessons:
                lessons[lesson_id] = {
                    "id": lesson_id,
                    "name": lesson_name,
                    "pagetext": "",
                    "transcript": "",
                    "has_pagetext": False,
                    "has_transcript": False,
                    "is_stub": False,
                    "thumbnail": None,
                    "video_path": None,
                    "tmpdir": tmpdir,
                }

            if doc_type == 'pagetext':
                lessons[lesson_id]['pagetext'] = content
                lessons[lesson_id]['has_pagetext'] = bool(content.strip())
            elif doc_type == 'transcript':
                lessons[lesson_id]['transcript'] = content
                lessons[lesson_id]['has_transcript'] = bool(content.strip())

    for lid, lesson in lessons.items():
        lesson['thumbnail'] = thumbnails.get(lid)
        lesson['video_path'] = _match_video_by_name(lid, lesson['name'], videos)
        if not lesson['pagetext'].strip() and not lesson['transcript'].strip():
            lesson['is_stub'] = True

    all_ids = list(lessons.keys())
    if order_ids:
        ordered = [lessons[str(oid)] for oid in order_ids if str(oid) in lessons]
        listed = set(str(oid) for oid in order_ids)
        ordered += [lessons[lid] for lid in sorted(all_ids, key=int) if lid not in listed]
    else:
        ordered = [lessons[lid] for lid in sorted(all_ids, key=int)]

    return ordered


def _count_image_slides(slide_defs: list) -> int:
    return sum(
        1 for sd in slide_defs
        if sd.get("type") in {"concept", "table", "discussion_question", "discussion_answer"}
        and sd.get("image_required", True)
    )


def run_pipeline(zip_path: str, chapter_name: str, chapter_number: str,
                 openai_api_key: str, order_ids: list = None,
                 model: str = "gpt-4.1",
                 total_slide_budget: int = None,
                 google_api_key: str = "", google_cse_id: str = "",
                 progress_callback=None,
                 vision_model: str = "gpt-4.1") -> tuple:
    """
    google_api_key/google_cse_id are accepted for backward compatibility with old app.py
    but are intentionally ignored because web fallback is disabled.
    """

    def progress(msg, pct=None):
        print(f"[{pct if pct is not None else '?':>3}%] {msg}")
        if progress_callback:
            progress_callback(msg, pct)

    qa_report = {
        "chapter": chapter_name,
        "chapter_number": chapter_number,
        "generated_at": datetime.now().isoformat(),
        "model": model,
        "vision_model": vision_model,
        "total_slides": 0,
        "lessons_processed": [],
        "lessons_skipped": [],
        "images_used": [],
        "images_missing": [],
        "scientific_names": [],
        "flags": [],
        "planning": {},
        "strict_image_rule": "Every image-bearing slide uses a frame from that lesson MP4. No web fallback. No placeholders."
    }

    progress("Parsing chapter ZIP...", 2)
    lessons = parse_chapter_zip(zip_path, order_ids)
    populated = [l for l in lessons if not l['is_stub']]
    stubs = [l for l in lessons if l['is_stub']]

    progress(f"Found {len(populated)} lessons, {len(stubs)} stubs", 4)

    if stubs:
        for s in stubs:
            qa_report["lessons_skipped"].append({
                "id": s["id"], "name": s["name"], "reason": "Empty stub"
            })
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"{len(stubs)} lessons skipped (empty files): {', '.join(s['name'] for s in stubs)}"
        })

    if not populated:
        raise RuntimeError("No valid lessons found in ZIP.")

    missing_videos = [f"{l['id']} - {l['name']}" for l in populated if not l.get("video_path")]
    if missing_videos:
        raise RuntimeError(
            "Missing MP4 file(s). Add lesson MP4s to the ZIP. Missing: "
            + "; ".join(missing_videos)
        )

    # Planning pass
    if total_slide_budget and total_slide_budget > 0:
        chapter_budget = total_slide_budget
    else:
        chapter_budget = default_chapter_budget(len(populated))

    progress(f"Planning slide allocation (target: {chapter_budget} slides)...", 6)

    try:
        plan = plan_chapter_slides(populated, chapter_budget, openai_api_key, model=model)
    except Exception as e:
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"Planning pass failed ({str(e)}); using even allocation fallback."
        })
        fallback_concept = max(1, min(4, round(chapter_budget / len(populated)) - 2))
        plan = {
            "reasoning": "Fallback even allocation (planning pass unavailable).",
            "glossary_pages": 2,
            "allocations": {l["id"]: fallback_concept for l in populated}
        }

    qa_report["planning"] = {
        "target_total": chapter_budget,
        "reasoning": plan.get("reasoning", ""),
        "glossary_pages": plan.get("glossary_pages", 2),
        "allocations": plan.get("allocations", {})
    }
    progress(f"Plan: {plan.get('reasoning','')[:120]}", 8)

    # Generate content and select frames before building slides
    lesson_outputs = []
    all_glossary = {}
    lesson_recaps = []
    cover_image_path = None

    for i, lesson in enumerate(populated):
        pct = 10 + int((i / len(populated)) * 62)
        progress(f"Generating lesson {i+1}/{len(populated)}: {lesson['name']}", pct)

        concept_budget = plan["allocations"].get(lesson["id"], 2)

        slide_data = generate_slide_content(
            lesson_name=lesson['name'],
            transcript=lesson['transcript'],
            pagetext=lesson['pagetext'],
            concept_slide_budget=concept_budget,
            api_key=openai_api_key,
            model=model
        )

        if "glossary_terms" in slide_data:
            all_glossary.update(slide_data["glossary_terms"])

        sci_names = re.findall(r'\*\*_([A-Z][a-z]+ [a-z]+)_\*\*', json.dumps(slide_data))
        if sci_names:
            qa_report["scientific_names"].extend(sci_names)

        key_points = []
        for sd in slide_data.get("slides", []):
            if sd.get("type") == "concept" and sd.get("body"):
                first_line = sd["body"].strip().split("\n")[0]
                clean = re.sub(r'\*\*', '', first_line)
                key_points.append(clean[:150])
        lesson_recaps.append({"name": lesson["name"], "key_points": "; ".join(key_points[:3])})

        image_slide_count = _count_image_slides(slide_data.get("slides", []))
        progress(f"Selecting {image_slide_count} video frame(s) for {lesson['name']}...", pct + 2)

        frame_map = assign_frames_to_slides(
            lesson=lesson,
            slide_defs=slide_data.get("slides", []),
            api_key=openai_api_key,
            vision_model=vision_model,
            progress_callback=progress
        )

        if not cover_image_path and frame_map:
            cover_image_path = next(iter(frame_map.values())).get("path")

        lesson_outputs.append({
            "lesson": lesson,
            "concept_budget": concept_budget,
            "slide_data": slide_data,
            "frame_map": frame_map,
        })

    # Build deck
    prs = create_presentation(LOGO_PATH)
    slide_count = 0

    progress("Building cover slide...", 75)
    build_cover_slide(prs, chapter_name, chapter_number, LOGO_PATH, cover_image_path=cover_image_path)
    slide_count += 1

    for i, bundle in enumerate(lesson_outputs):
        lesson = bundle["lesson"]
        slide_data = bundle["slide_data"]
        frame_map = bundle["frame_map"]
        concept_budget = bundle["concept_budget"]

        lesson_qa = {
            "id": lesson["id"],
            "name": lesson["name"],
            "word_count": len((lesson['transcript'] + lesson['pagetext']).split()),
            "concept_budget": concept_budget,
            "slides_built": 0,
            "images_found": 0,
            "images_missing": 0,
            "video_path": lesson.get("video_path")
        }

        progress(f"Building slides for {lesson['name']}...", 77 + int((i / max(1, len(lesson_outputs))) * 10))

        for slide_index, slide_def in enumerate(slide_data.get("slides", [])):
            stype = slide_def.get("type", "concept")
            title = slide_def.get("title", lesson["name"])
            notes = slide_def.get("speaker_notes", "")

            image_required = (
                stype in {"concept", "table", "discussion_question", "discussion_answer"}
                and slide_def.get("image_required", True)
            )
            frame_info = frame_map.get(slide_index) if image_required else None
            img_path = frame_info.get("path") if frame_info else None

            if image_required and not img_path:
                lesson_qa["images_missing"] += 1
                qa_report["images_missing"].append({
                    "lesson": lesson["name"],
                    "slide_type": stype,
                    "title": title,
                    "reason": "No MP4 frame selected"
                })
                raise RuntimeError(f"No MP4 frame selected for {lesson['name']} / {stype}. No placeholders are allowed.")

            if image_required:
                lesson_qa["images_found"] += 1
                qa_report["images_used"].append({
                    "lesson": lesson["name"],
                    "slide_type": stype,
                    "title": title,
                    "source_video": lesson.get("video_path"),
                    "frame_path": img_path,
                    "timestamp_seconds": frame_info.get("timestamp"),
                    "target_time_seconds": frame_info.get("target_time"),
                    "anchor_text": frame_info.get("anchor_text"),
                    "visual_focus": slide_def.get("visual_focus"),
                    "selection_method": frame_info.get("selection_method"),
                    "selection_reason": frame_info.get("selection_reason"),
                    "vision_confidence": frame_info.get("vision_confidence"),
                    "technical_score": round(frame_info.get("technical_score", 0), 4),
                    "candidate_count": frame_info.get("total_candidates_considered", 0)
                })

            if stype == "concept":
                build_concept_slide(
                    prs, lesson_name=title, body_text=slide_def.get("body", ""),
                    sub_label=slide_def.get("sub_label"),
                    image_path=img_path,
                    speaker_notes=notes, logo_path=LOGO_PATH
                )
            elif stype == "table":
                build_table_slide(
                    prs, lesson_name=title,
                    headers=slide_def.get("headers", []),
                    rows=slide_def.get("rows", []),
                    sub_title=slide_def.get("sub_title"),
                    image_path=img_path,
                    speaker_notes=notes, logo_path=LOGO_PATH
                )
            elif stype == "discussion_question":
                build_discussion_question_slide(
                    prs, lesson_name=title,
                    question_text=slide_def.get("question", ""),
                    hint_text=slide_def.get("hint"),
                    image_path=img_path,
                    speaker_notes=notes, logo_path=LOGO_PATH
                )
            elif stype == "discussion_answer":
                build_discussion_answer_slide(
                    prs, lesson_name=title,
                    answer_summary=slide_def.get("answer_summary", ""),
                    answer_explanation=slide_def.get("answer_explanation", ""),
                    image_path=img_path,
                    speaker_notes=notes, logo_path=LOGO_PATH
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

    # Student-facing chapter summary
    progress("Building chapter summary...", 88)
    try:
        chapter_summary = generate_chapter_summary(
            chapter_name, lesson_recaps, openai_api_key, model=model
        )
        build_summary_slide(
            prs,
            summary_statement=chapter_summary.get("summary_statement", f"{chapter_name} - key takeaways."),
            table_headers=chapter_summary.get("headers", ["Concept", "Definition", "Key Point"]),
            table_rows=chapter_summary.get("rows", [])[:8],
            logo_path=LOGO_PATH,
            speaker_notes="Use this slide to recap the chapter's core concepts with students before moving on."
        )
    except Exception as e:
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"Chapter summary generation failed ({str(e)}); using fallback."
        })
        fallback_rows = [[lr["name"], "Core lesson concept", lr["key_points"][:100]] for lr in lesson_recaps[:8]]
        build_summary_slide(
            prs,
            summary_statement=f"{chapter_name} - key concepts covered in this chapter.",
            table_headers=["Lesson", "Definition", "Key Idea"],
            table_rows=fallback_rows,
            logo_path=LOGO_PATH,
            speaker_notes="Recap the chapter's core concepts with students."
        )
    slide_count += 1

    # Glossary
    progress("Building glossary...", 92)
    TERMS_PER_PAGE = 10
    glossary_pages = qa_report["planning"].get("glossary_pages", 2)
    max_terms = TERMS_PER_PAGE * glossary_pages
    gloss_items = list(all_glossary.items())[:max_terms]
    if not gloss_items:
        gloss_items = [("Glossary pending", "No glossary terms were returned for this chapter.")]

    for page_start in range(0, len(gloss_items), TERMS_PER_PAGE):
        page_terms = dict(gloss_items[page_start:page_start + TERMS_PER_PAGE])
        build_glossary_slide(prs, page_terms, logo_path=LOGO_PATH)
        slide_count += 1

    # Save
    progress("Saving PPTX...", 96)
    safe_name = re.sub(r'[^\w\s-]', '', chapter_name).replace(' ', '_')
    out_path = os.path.join(tempfile.gettempdir(), f"JoVE_Chapter{chapter_number}_{safe_name}.pptx")
    prs.save(out_path)

    qa_report["total_slides"] = slide_count
    qa_report["output_file"] = out_path

    if qa_report["images_missing"]:
        raise RuntimeError(
            f"Strict image rule violated: {len(qa_report['images_missing'])} slide(s) did not receive MP4 frames."
        )

    if qa_report["scientific_names"]:
        qa_report["flags"].append({
            "level": "REVIEW",
            "message": f"Scientific names to verify: {', '.join(set(qa_report['scientific_names']))}"
        })

    progress(f"Done! {slide_count} slides generated (target was {chapter_budget}).", 100)
    return out_path, qa_report
