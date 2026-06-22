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
from video_sourcing import assign_frames_to_slides, select_frame_for_slide
from ppt_builder import (
    create_presentation, build_cover_slide, build_concept_slide,
    build_table_slide, build_discussion_question_slide,
    build_discussion_answer_slide, build_summary_slide, build_glossary_slide
)
from formatting_validator import validate_pptx_formatting

LOGO_PATH = os.path.join(os.path.dirname(__file__), "jove_logo.png")

FORBIDDEN_METADATA_RE = re.compile(r"^\s*(writer|author|reviewer|prepared\s*by|created\s*by)\s*[:\-].*$", re.IGNORECASE)
MARKDOWN_RE = re.compile(r"(\*\*|__|`)")


def _sanitize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("**", "").replace("__", "").replace("`", "")
    lines = []
    for line in text.splitlines():
        if FORBIDDEN_METADATA_RE.match(line):
            continue
        if re.search(r"\[(INSERT IMAGE|TODO|PLACEHOLDER|IMAGE)\]", line, re.IGNORECASE):
            continue
        lines.append(re.sub(r"\s+", " ", line).strip())
    return "\n".join(line for line in lines if line).strip()


def _sanitize_obj(obj):
    if isinstance(obj, dict):
        return {k: _sanitize_obj(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_obj(v) for v in obj]
    if isinstance(obj, str):
        return _sanitize_text(obj)
    return obj


def _normalize_slide_data_for_formatting(slide_data: dict) -> dict:
    slide_data = _sanitize_obj(slide_data or {})
    slides = slide_data.get("slides") or []
    normalized = []
    for sd in slides:
        if not isinstance(sd, dict):
            continue
        stype = sd.get("type", "concept")
        # Hard cap table rows for formatting; extra content belongs in notes/source, not squeezed into slide.
        if stype == "table":
            sd["rows"] = (sd.get("rows") or [])[:4]
        if stype == "summary":
            sd["rows"] = (sd.get("rows") or [])[:3]
            sd["summary_statement"] = _sanitize_text(sd.get("summary_statement", ""))
        if stype == "concept":
            title = _sanitize_text(sd.get("title", ""))
            if title.lower() in {"definition / core idea", "definition and core process", "definition", "core idea"}:
                sd["title"] = f"What is {title}?" if title else "What is this concept?"
            sd["body"] = _sanitize_text(sd.get("body", ""))
        normalized.append(sd)
    slide_data["slides"] = normalized
    glossary = slide_data.get("glossary_terms") or {}
    if isinstance(glossary, dict):
        slide_data["glossary_terms"] = {
            _sanitize_text(k): _sanitize_text(v)
            for k, v in glossary.items()
            if _sanitize_text(k)
        }
    return slide_data


def _sanitize_lessons(lessons: list) -> list:
    for lesson in lessons:
        lesson["name"] = _sanitize_text(lesson.get("name")) or lesson.get("id", "Lesson")
        lesson["pagetext"] = _sanitize_text(lesson.get("pagetext", ""))
        lesson["transcript"] = _sanitize_text(lesson.get("transcript", ""))
    return lessons


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


def _extract_first_id(filename: str) -> Optional[str]:
    m = re.search(r'(?<!\d)(\d{4,8})(?!\d)', filename)
    return m.group(1) if m else None


def _detect_doc_type(filename: str) -> Optional[str]:
    stem = os.path.splitext(os.path.basename(filename))[0]
    norm = _norm_for_match(stem)
    has_pagetext = "pagetext" in norm or ("page" in norm and "text" in norm)
    has_transcript = "transcript" in norm or "transcription" in norm
    if has_pagetext and not has_transcript:
        return "pagetext"
    if has_transcript and not has_pagetext:
        return "transcript"
    return None


def _clean_name_piece(piece: str) -> str:
    piece = os.path.splitext(os.path.basename(piece))[0]
    piece = re.sub(r'(?<!\d)\d{4,8}(?!\d)', ' ', piece)
    piece = re.sub(r'(?i)pagetext|page text|page-text|page_text|transcript|transcription|video|vid|mp4|final|draft|copy', ' ', piece)
    piece = re.sub(r'[_\-.]+', ' ', piece)
    return re.sub(r'\s+', ' ', piece).strip()


def _infer_lesson_name_from_filename(filename: str, lesson_id: str) -> str:
    cleaned = _clean_name_piece(filename)
    if cleaned and _norm_for_match(cleaned) != _norm_for_match(lesson_id):
        return _camel_to_title(cleaned)
    return f"Lesson {lesson_id}"


def _infer_lesson_name_from_content(content: str, lesson_id: str) -> str:
    fallback = f"Lesson {lesson_id}"
    if not content:
        return fallback
    lines = [re.sub(r'\s+', ' ', line).strip() for line in content.splitlines()]
    skip_terms = {"pagetext", "page text", "transcript", "lesson", "copyright"}
    for line in lines[:12]:
        if not line:
            continue
        lower = line.lower()
        if any(term in lower for term in skip_terms):
            continue
        if re.fullmatch(r'[\d\W_]+', line):
            continue
        words = line.split()
        if 1 <= len(words) <= 12 and 3 <= len(line) <= 90:
            return _camel_to_title(line)
    return fallback


def _is_fallback_lesson_name(name: str, lesson_id: str) -> bool:
    return not name or name.strip().lower() == f"lesson {lesson_id}".lower()


def _set_better_lesson_name(lesson: Dict, candidate: str) -> None:
    if candidate and _is_fallback_lesson_name(lesson.get("name", ""), lesson["id"]):
        lesson["name"] = candidate


def _match_video_by_name(lesson_id: str, lesson_name: str, all_video_paths: List[str]) -> Optional[str]:
    lesson_id = str(lesson_id)

    for path in all_video_paths:
        fname = os.path.basename(path)
        if re.match(rf"^{re.escape(lesson_id)}(?:[^\d].*)?\.mp4$", fname, re.IGNORECASE):
            return path

    for path in all_video_paths:
        fname = os.path.basename(path)
        if re.search(rf"(?<!\d){re.escape(lesson_id)}(?!\d)", fname):
            return path

    lesson_norm = _norm_for_match(lesson_name)
    if lesson_norm and not lesson_norm.startswith("lesson"):
        for path in all_video_paths:
            if lesson_norm in _norm_for_match(os.path.basename(path)):
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
            lesson_id = _extract_first_id(fname)

            if lower.endswith(('.jpg', '.jpeg', '.png')):
                if lesson_id:
                    thumbnails[lesson_id] = full_path
                continue

            if lower.endswith('.mp4'):
                videos.append(full_path)
                continue

            if not lower.endswith('.docx'):
                continue

            if not lesson_id:
                continue

            doc_type = _detect_doc_type(fname)
            if not doc_type:
                continue

            content = _read_docx(full_path)
            filename_name = _infer_lesson_name_from_filename(fname, lesson_id)
            content_name = _infer_lesson_name_from_content(content, lesson_id)
            lesson_name = content_name if not _is_fallback_lesson_name(content_name, lesson_id) else filename_name

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
            else:
                _set_better_lesson_name(lessons[lesson_id], lesson_name)

            if doc_type == 'pagetext':
                lessons[lesson_id]['pagetext'] = content
                lessons[lesson_id]['has_pagetext'] = bool(content.strip())
                _set_better_lesson_name(lessons[lesson_id], content_name)
            elif doc_type == 'transcript':
                lessons[lesson_id]['transcript'] = content
                lessons[lesson_id]['has_transcript'] = bool(content.strip())
                _set_better_lesson_name(lessons[lesson_id], content_name)

    for lid, lesson in lessons.items():
        lesson['thumbnail'] = thumbnails.get(lid)
        lesson['video_path'] = _match_video_by_name(lid, lesson['name'], videos)
        if not lesson['pagetext'].strip() and not lesson['transcript'].strip():
            lesson['is_stub'] = True

    all_ids = list(lessons.keys())

    def _sort_key(lid: str):
        try:
            return int(lid)
        except ValueError:
            return lid

    if order_ids:
        ordered = [lessons[str(oid)] for oid in order_ids if str(oid) in lessons]
        listed = set(str(oid) for oid in order_ids)
        ordered += [lessons[lid] for lid in sorted(all_ids, key=_sort_key) if lid not in listed]
    else:
        ordered = [lessons[lid] for lid in sorted(all_ids, key=_sort_key)]

    return ordered



def _rebalance_plan_to_budget(lessons: list, plan: dict, total_slide_budget: int) -> dict:
    """
    Final deterministic budget guard. The planner prompt should already do this,
    but this function prevents drift if the model over-allocates or if the
    fallback plan is used.
    """
    if not lessons or not total_slide_budget:
        return plan

    glossary_pages = max(1, min(3, int(plan.get("glossary_pages", 2))))
    reserves = 1 + 1 + glossary_pages
    mandatory_qa = 2 * len(lessons)
    available_concepts = max(len(lessons), total_slide_budget - reserves - mandatory_qa)

    allocations = {str(k): max(1, min(4, int(v))) for k, v in (plan.get("allocations") or {}).items()}
    for lesson in lessons:
        allocations.setdefault(str(lesson["id"]), 2)

    def density(lesson):
        return len((lesson.get("transcript", "") + " " + lesson.get("pagetext", "")).split())

    current = sum(allocations[str(lesson["id"])] for lesson in lessons)

    while current > available_concepts:
        reducible = [lesson for lesson in lessons if allocations[str(lesson["id"])] > 1]
        if not reducible:
            break
        reducible.sort(key=lambda lesson: (allocations[str(lesson["id"])], density(lesson)), reverse=True)
        chosen = reducible[0]
        allocations[str(chosen["id"])] -= 1
        current -= 1

    while current < available_concepts:
        expandable = [lesson for lesson in lessons if allocations[str(lesson["id"])] < 4]
        if not expandable:
            break
        expandable.sort(key=density, reverse=True)
        chosen = expandable[0]
        allocations[str(chosen["id"])] += 1
        current += 1

    plan["glossary_pages"] = glossary_pages
    plan["allocations"] = allocations
    note = (
        f"Final deterministic budget guard: target={total_slide_budget}, "
        f"reserves={reserves}, mandatory_QA={mandatory_qa}, "
        f"available_concepts={available_concepts}, final_concepts={current}."
    )
    plan["reasoning"] = (str(plan.get("reasoning", "")).strip() + " " + note).strip()
    plan["budget_guard_note"] = note
    return plan



def _count_image_slides(slide_defs: list) -> int:
    """Count slides that require a main visual/frame selection."""
    if not slide_defs:
        return 0
    image_types = {"concept", "table", "discussion_question", "discussion_answer", "summary"}
    return sum(1 for slide in slide_defs if (slide or {}).get("type") in image_types)


def _build_table_row_frame_map(lesson: dict, slide_defs: list, openai_api_key: str,
                               vision_model: str, progress_callback=None) -> dict:
    """
    Select one image per table row so table slides can embed visuals inside rows.
    Uses JoVE MP4 first; video_sourcing may use approved AI fallback if no suitable frame exists.
    """
    result = {}
    video_path = lesson.get("video_path")
    if not video_path:
        return result

    used_paths = set()
    used_timestamps = []
    work_dir = os.path.join(tempfile.gettempdir(), "jove_table_row_frames", lesson["id"])

    for slide_idx, slide_def in enumerate(slide_defs):
        if slide_def.get("type") != "table":
            continue
        rows = (slide_def.get("rows") or [])[:4]
        row_paths = []
        for ri, row in enumerate(rows):
            row_text = " ".join(str(x) for x in (row if isinstance(row, list) else [row]))
            fake_slide = {
                "type": "concept",
                "title": slide_def.get("sub_title") or slide_def.get("title") or lesson["name"],
                "visual_focus": row_text or slide_def.get("visual_focus") or lesson["name"],
                "transcript_anchor_text": row_text or slide_def.get("transcript_anchor_text") or "",
            }
            if progress_callback:
                progress_callback(f"Selecting table row image {ri+1}/{len(rows)} for {lesson['name']}...", None)
            try:
                info = select_frame_for_slide(
                    video_path=video_path,
                    lesson_name=lesson["name"],
                    transcript=lesson.get("transcript", ""),
                    slide_def=fake_slide,
                    total_image_slides=max(3, len(rows)),
                    api_key=openai_api_key,
                    work_dir=work_dir,
                    vision_model=vision_model,
                    used_frame_paths=used_paths,
                    used_timestamps=used_timestamps,
                )
                row_paths.append(info.get("path"))
                used_paths.add(info.get("path"))
                used_timestamps.append(info.get("timestamp"))
            except Exception:
                row_paths.append(None)
        result[slide_idx] = row_paths
    return result




def _build_cover_description(chapter_name: str, lesson_outputs: list) -> str:
    """Create a short cover description from generated slide content, not hallucinated externally."""
    pieces = []
    for bundle in lesson_outputs:
        slide_data = bundle.get("slide_data", {})
        for sd in slide_data.get("slides", []):
            if sd.get("type") in {"concept", "summary"}:
                body = _sanitize_text(sd.get("body") or sd.get("summary_statement") or "")
                if body:
                    pieces.append(body)
                    break
        if len(pieces) >= 2:
            break
    if pieces:
        text = " ".join(pieces)
        words = text.split()
        return " ".join(words[:28]).rstrip(".,;:") + "."
    return f"An overview of key concepts in {chapter_name}."


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
        "image_rule": "JoVE MP4 frame first. Approved AI fallback only when no suitable JoVE frame exists. No web search. No placeholders."
    }

    progress("Parsing chapter ZIP...", 2)
    lessons = parse_chapter_zip(zip_path, order_ids)
    populated = _sanitize_lessons([l for l in lessons if not l['is_stub']])
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

    plan = _rebalance_plan_to_budget(populated, plan, chapter_budget)

    qa_report["planning"] = {
        "target_total": chapter_budget,
        "reasoning": plan.get("reasoning", ""),
        "glossary_pages": plan.get("glossary_pages", 2),
        "allocations": plan.get("allocations", {}),
        "budget_guard_note": plan.get("budget_guard_note", "")
    }
    progress(f"Plan: {plan.get('reasoning','')[:120]}", 8)

    # Generate content and select frames before building slides
    lesson_outputs = []
    all_glossary = {}
    lesson_recaps = []
    cover_image_path = None
    cover_image_paths = []

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
        slide_data = _normalize_slide_data_for_formatting(slide_data)

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

        table_row_frame_map = _build_table_row_frame_map(
            lesson=lesson,
            slide_defs=slide_data.get("slides", []),
            openai_api_key=openai_api_key,
            vision_model=vision_model,
            progress_callback=progress
        )

        if frame_map:
            for info in frame_map.values():
                fp = info.get("path")
                if fp and fp not in cover_image_paths:
                    cover_image_paths.append(fp)
                    if not cover_image_path:
                        cover_image_path = fp
                if len(cover_image_paths) >= 3:
                    break

        lesson_outputs.append({
            "lesson": lesson,
            "concept_budget": concept_budget,
            "slide_data": slide_data,
            "frame_map": frame_map,
            "table_row_frame_map": table_row_frame_map,
        })

    # Build deck
    prs = create_presentation(LOGO_PATH)
    slide_count = 0

    progress("Building cover slide...", 75)
    build_cover_slide(prs, chapter_name, chapter_number, LOGO_PATH, cover_image_path=cover_image_path, cover_image_paths=cover_image_paths, chapter_description=_build_cover_description(chapter_name, lesson_outputs), slide_number=slide_count + 1)
    slide_count += 1

    for i, bundle in enumerate(lesson_outputs):
        lesson = bundle["lesson"]
        slide_data = bundle["slide_data"]
        frame_map = bundle["frame_map"]
        table_row_frame_map = bundle.get("table_row_frame_map", {})
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
                    speaker_notes=notes, logo_path=LOGO_PATH, slide_number=slide_count + 1
                )
            elif stype == "table":
                build_table_slide(
                    prs, lesson_name=title,
                    headers=slide_def.get("headers", []),
                    rows=slide_def.get("rows", []),
                    sub_title=slide_def.get("sub_title"),
                    table_kind=slide_def.get("table_kind"),
                    image_path=img_path,
                    row_image_paths=table_row_frame_map.get(slide_index),
                    speaker_notes=notes, logo_path=LOGO_PATH, slide_number=slide_count + 1
                )
            elif stype == "discussion_question":
                build_discussion_question_slide(
                    prs, lesson_name=title,
                    question_text=slide_def.get("question", ""),
                    hint_text=slide_def.get("hint"),
                    image_path=img_path,
                    speaker_notes=notes, logo_path=LOGO_PATH, slide_number=slide_count + 1
                )
            elif stype == "discussion_answer":
                build_discussion_answer_slide(
                    prs, lesson_name=title,
                    answer_summary=slide_def.get("answer_summary", ""),
                    answer_explanation=slide_def.get("answer_explanation", ""),
                    image_path=img_path,
                    speaker_notes=notes, logo_path=LOGO_PATH, slide_number=slide_count + 1
                )
            elif stype == "summary":
                build_summary_slide(
                    prs,
                    summary_statement=slide_def.get("summary_statement", ""),
                    table_headers=slide_def.get("headers", []),
                    table_rows=slide_def.get("rows", []),
                    logo_path=LOGO_PATH,
                    speaker_notes=notes,
                    slide_number=slide_count + 1
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
            table_rows=chapter_summary.get("rows", [])[:3],
            logo_path=LOGO_PATH,
            speaker_notes="Use this slide to recap the chapter's core concepts with students before moving on.",
            slide_number=slide_count + 1
        )
    except Exception as e:
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"Chapter summary generation failed ({str(e)}); using fallback."
        })
        fallback_rows = [[lr["name"], "Core lesson concept", lr["key_points"][:100]] for lr in lesson_recaps[:3]]
        build_summary_slide(
            prs,
            summary_statement=f"{chapter_name} - key concepts covered in this chapter.",
            table_headers=["Lesson", "Definition", "Key Idea"],
            table_rows=fallback_rows,
            logo_path=LOGO_PATH,
            speaker_notes="Recap the chapter's core concepts with students.",
            slide_number=slide_count + 1
        )
    slide_count += 1

    # Glossary
    progress("Building glossary...", 92)
    TERMS_PER_PAGE = 6
    glossary_pages = qa_report["planning"].get("glossary_pages", 2)
    max_terms = TERMS_PER_PAGE * glossary_pages
    gloss_items = list(all_glossary.items())[:max_terms]
    if not gloss_items:
        gloss_items = [("Glossary pending", "No glossary terms were returned for this chapter.")]

    for page_start in range(0, len(gloss_items), TERMS_PER_PAGE):
        page_terms = dict(gloss_items[page_start:page_start + TERMS_PER_PAGE])
        build_glossary_slide(prs, page_terms, logo_path=LOGO_PATH, slide_number=slide_count + 1)
        slide_count += 1

    # Save
    progress("Saving PPTX...", 96)
    safe_name = re.sub(r'[^\w\s-]', '', chapter_name).replace(' ', '_')
    out_path = os.path.join(tempfile.gettempdir(), f"JoVE_Chapter{chapter_number}_{safe_name}.pptx")
    prs.save(out_path)

    qa_report["total_slides"] = slide_count
    qa_report["output_file"] = out_path

    try:
        formatting_validation = validate_pptx_formatting(out_path)
        qa_report["formatting_validation"] = formatting_validation
        if not formatting_validation.get("target_met"):
            qa_report["flags"].append({
                "level": "FORMATTING_REVIEW",
                "message": f"Formatting score {formatting_validation.get('formatting_score')}% is below 95%. Review validator findings."
            })
    except Exception as e:
        qa_report["flags"].append({"level": "WARNING", "message": f"Formatting validator failed: {str(e)}"})

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
