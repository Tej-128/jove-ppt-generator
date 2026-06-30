"""
JoVE PPT Pipeline v3
Architecture:
parse ZIP -> planning pass -> AI slide generation -> transcript-aligned MP4 frame selection ->
presentation build -> QA report.

Strict visual rule:
No placeholders and no web fallback. Every image-bearing lesson slide first receives a
frame from that lesson's MP4. Selected high-impact frames are upgraded through
the OpenAI image API into polished, presentation-grade educational illustrations.
The AI visual step is budgeted and timeout-protected so a full chapter finishes;
when the budget is reached or one upgrade fails, the selected JoVE frame is kept
and locally presentation-enhanced rather than blocking the entire PPT.
"""

import os
import re
import zipfile
import tempfile
import json
import base64
import hashlib
import time
import traceback
from datetime import datetime
from typing import Dict, List, Optional

from docx import Document as DocxDocument

from ai_generator import generate_slide_content, generate_chapter_summary, generate_chapter_overview
from planner import plan_chapter_slides, default_chapter_budget
from video_sourcing import assign_frames_to_slides, select_frame_for_slide
from ppt_builder import (
    create_presentation, build_cover_slide, build_concept_slide,
    build_table_slide, build_discussion_question_slide,
    build_discussion_answer_slide, build_summary_slide, build_glossary_slide
)
from formatting_validator import validate_pptx_formatting

LOGO_PATH = os.path.join(os.path.dirname(__file__), "jove_logo.png")
PIPELINE_VERSION = "v7.3.15_chapter_intro_discussion_captions"


def _safe_int_env(name: str, default: int, minimum: int = None, maximum: int = None) -> int:
    """Read an integer setting safely from environment variables."""
    raw = os.getenv(name)
    try:
        value = int(raw) if raw not in {None, ""} else int(default)
    except Exception:
        value = int(default)
    if minimum is not None:
        value = max(minimum, value)
    if maximum is not None:
        value = min(maximum, value)
    return value


def _resolve_secret_value(*names: str) -> str:
    """Read from env first, then Streamlit secrets when available.

    This keeps local runs and Streamlit Cloud runs aligned without requiring
    app.py changes. Missing Streamlit is ignored silently.
    """
    for name in names:
        val = os.getenv(name)
        if val:
            return val
    try:
        import streamlit as st  # type: ignore
        for name in names:
            try:
                val = st.secrets.get(name)
                if val:
                    return str(val)
            except Exception:
                pass
            # Also support nested [openai] api_key style secrets.
            try:
                if name.upper() == "OPENAI_API_KEY":
                    val = st.secrets.get("openai", {}).get("api_key")
                    if val:
                        return str(val)
            except Exception:
                pass
    except Exception:
        pass
    return ""


def _resolve_openai_api_key(api_key: str = None) -> str:
    return (api_key or _resolve_secret_value("OPENAI_API_KEY", "openai_api_key", "OPENAI_KEY") or "").strip()


def _runtime_diag_dir() -> str:
    path = os.path.join(tempfile.gettempdir(), "jove_runtime_diagnostics")
    os.makedirs(path, exist_ok=True)
    return path


def _new_runtime_diag_path(chapter_name: str, chapter_number: str) -> str:
    safe_ch = re.sub(r"[^A-Za-z0-9_-]+", "_", str(chapter_name or "chapter")).strip("_")[:60]
    safe_no = re.sub(r"[^A-Za-z0-9_-]+", "_", str(chapter_number or "x")).strip("_")[:20]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(_runtime_diag_dir(), f"jove_ch{safe_no}_{safe_ch}_{stamp}.jsonl")


def _append_runtime_event(qa_report: dict, event: str, message: str = "", **extra) -> None:
    """Append a crash-safe JSONL runtime trace.

    This file is written during the run, so if Streamlit or an API call cuts off
    midway, the last successful stage is still visible from app logs / temp path.
    """
    if not isinstance(qa_report, dict):
        return
    path = qa_report.get("runtime_diagnostics_path")
    if not path:
        return
    payload = {
        "time": datetime.now().isoformat(timespec="seconds"),
        "event": event,
        "message": str(message or "")[:1000],
        **extra,
    }
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        pass


def _ai_visual_budget_settings() -> dict:
    """Production-safe defaults.

    The first full-chapter run must finish. These settings upgrade a limited
    number of high-impact images, avoid table-row image upgrades by default, and
    keep a clear diagnostic trail when image work is skipped or fails.
    """
    return {
        "max_per_chapter": _safe_int_env("JOVE_AI_VISUALS_MAX_PER_CHAPTER", 8, minimum=0, maximum=200),
        "max_seconds": _safe_int_env("JOVE_AI_VISUALS_MAX_SECONDS", 300, minimum=0, maximum=7200),
        "timeout_seconds": _safe_int_env("JOVE_AI_VISUAL_TIMEOUT_SECONDS", 60, minimum=10, maximum=300),
        "max_per_lesson": _safe_int_env("JOVE_AI_VISUALS_MAX_PER_LESSON", 1, minimum=0, maximum=20),
        "table_rows_enabled": str(os.getenv("JOVE_AI_VISUALS_TABLE_ROWS", "0")).strip().lower() in {"1", "true", "yes", "on"},
        "fail_on_error": str(os.getenv("JOVE_AI_VISUALS_FAIL_ON_ERROR", "0")).strip().lower() in {"1", "true", "yes", "on"},
    }

FORBIDDEN_METADATA_RE = re.compile(r"^\s*(writer|author|reviewer|prepared\s*by|created\s*by|presenter|date|file\s*name|source\s*file|pagetext|page\s*text|script|transcript|transcription)\s*[:\-].*$", re.IGNORECASE)
MARKDOWN_RE = re.compile(r"(\*\*|__|`)")

COMMON_TEXT_FIXES = {
    "Funtion": "Function",
    "funtion": "function",
}

GENERIC_GLOSSARY_TERMS = {
    "living organisms", "single-celled organisms", "multicellular organisms",
    "plants", "animals", "cell structure", "cell function", "cell shape",
    "lesson", "writer", "page text", "transcript", "definition", "example"
}


def _apply_common_text_fixes(value: str) -> str:
    for bad, good in COMMON_TEXT_FIXES.items():
        value = value.replace(bad, good)
    return value


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
    value = "\n".join(line for line in lines if line).strip()
    return _apply_common_text_fixes(value)


def _sanitize_obj(obj):
    if isinstance(obj, dict):
        return {k: _sanitize_obj(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_obj(v) for v in obj]
    if isinstance(obj, str):
        return _sanitize_text(obj)
    return obj


def _word_count(text: str) -> int:
    return len(re.findall(r"[A-Za-z0-9]+", str(text or "")))

def _sentence_list(text: str) -> list:
    text = _sanitize_text(text)
    if not text:
        return []
    return [p.strip() for p in re.split(r"(?<=[.!?])\s+", text) if p.strip()]


def _is_complete_sentence(text: str) -> bool:
    text = _sanitize_text(text)
    return bool(text and re.search(r"[.!?]$", text) and not re.search(r"\b(and|or|for|with|to|of|the|their|a|an|relative|growing|associated)\.$", text, re.I))


def _complete_sentence_text(text: str, max_words: int = 18) -> str:
    text = _sanitize_text(text)
    for sent in _sentence_list(text):
        if _is_complete_sentence(sent) and len(sent.split()) <= max_words:
            return sent
    words = text.split()[:max_words]
    return (" ".join(words).rstrip(".,;:") + ".") if words else ""


def _is_blank_cell(value) -> bool:
    text = _sanitize_text(value)
    return not text or text.lower() in {"n/a", "na", "none", "null", "-", "—"}


def _is_image_header(header: str) -> bool:
    h = _sanitize_text(header).lower()
    return h in {"image", "visual", "figure", "picture"} or "image" in h or "visual" in h


def _make_cell_more_useful(text: str, header: str, row_label: str) -> str:
    """Light deterministic cleanup only; avoid adding outside scientific facts."""
    text = _sanitize_text(text)
    if _is_blank_cell(text):
        return ""
    if _word_count(text) >= 5:
        return text
    h = _sanitize_text(header).lower()
    if "example" in h or "application" in h or "key" in h:
        return f"Key point: {text}."
    if "definition" in h or "meaning" in h:
        return f"{row_label}: {text}."
    return text


def _drop_unusable_text_columns(headers: list, rows: list) -> tuple:
    """Remove text columns that would render blank/weak visible columns."""
    headers = [_sanitize_text(h) for h in (headers or [])]
    rows = [[_sanitize_text(x) for x in list(row)] for row in (rows or [])]

    if not headers:
        return ["Concept", "Key Point"], rows or [["Core idea", "Main learning point from this slide."]]

    keep_indices = [i for i, h in enumerate(headers) if not _is_image_header(h)]
    headers = [headers[i] for i in keep_indices]
    rows = [[(row[i] if i < len(row) else "") for i in keep_indices] for row in rows]

    if not headers:
        headers = ["Concept", "Key Point"]
        rows = [["Core idea", "Main learning point from this slide."] for _ in (rows or [None])]

    # Normalize row lengths.
    n_cols = len(headers)
    norm_rows = []
    for row in rows or []:
        row = list(row)
        if len(row) < n_cols:
            row += [""] * (n_cols - len(row))
        norm_rows.append(row[:n_cols])
    rows = norm_rows or [["Core idea"] + [""] * (n_cols - 1)]

    # Drop all-blank non-first columns.
    keep = []
    for ci, h in enumerate(headers):
        col = [row[ci] if ci < len(row) else "" for row in rows]
        if ci == 0 or not all(_is_blank_cell(x) for x in col):
            keep.append(ci)
    headers = [headers[i] for i in keep]
    rows = [[(row[i] if i < len(row) else "") for i in keep] for row in rows]

    # Drop non-first columns with any blanks instead of showing partial blank columns.
    keep = []
    for ci, h in enumerate(headers):
        col = [row[ci] if ci < len(row) else "" for row in rows]
        if ci == 0 or not any(_is_blank_cell(x) for x in col):
            keep.append(ci)
    headers = [headers[i] for i in keep]
    rows = [[(row[i] if i < len(row) else "") for i in keep] for row in rows]

    rows = [row for row in rows if row and not _is_blank_cell(row[0])] or rows

    if len(headers) == 1:
        headers = [headers[0] or "Concept", "Key Point"]
        rows = [[row[0], row[0]] for row in rows]

    repaired = []
    for row in rows:
        label = row[0] if row else "Core idea"
        new_row = []
        for ci, h in enumerate(headers):
            val = row[ci] if ci < len(row) else ""
            if ci == 0:
                new_row.append(_sanitize_text(val))
            else:
                new_row.append(_make_cell_more_useful(val, h, label))
        repaired.append(new_row)

    # Final pass: if any generated repair still leaves a blank non-first column, drop it.
    final_keep = []
    for ci, h in enumerate(headers):
        col = [row[ci] if ci < len(row) else "" for row in repaired]
        if ci == 0 or not any(_is_blank_cell(x) for x in col):
            final_keep.append(ci)
    headers = [headers[i] for i in final_keep]
    repaired = [[(row[i] if i < len(row) else "") for i in final_keep] for row in repaired]

    return headers, repaired


def _repair_table_slide_content(slide_def: dict) -> dict:
    """Contextual table cleanup: no blank columns, no mandatory 3-column forcing."""
    headers = slide_def.get("headers") or []
    rows = (slide_def.get("rows") or [])
    headers, rows = _drop_unusable_text_columns(headers, rows)
    slide_def["headers"] = headers
    slide_def["rows"] = rows
    return slide_def


def _ensure_table_row_images(row_image_paths, rows, fallback_image_path=None):
    """
    Return row images only when every row has a valid, unique visual.

    Earlier versions filled missing row images by repeating the main slide frame.
    That prevented blank Image cells but created redundant/low-quality table visuals.
    Current rule: add an Image column only when it can be fully and meaningfully
    populated. Otherwise, the builder renders a text-only table with no Image column.
    """
    row_count = len(rows or [])
    if row_count <= 0:
        return []

    paths = list(row_image_paths or [])[:row_count]
    fixed = []
    seen = set()
    for p in paths:
        if not p or not os.path.exists(p):
            return []
        key = os.path.abspath(p)
        if key in seen:
            return []
        seen.add(key)
        fixed.append(p)

    if len(fixed) != row_count:
        return []
    return fixed

def _summary_row_images(rows, available_images):
    """Use already-selected JoVE frames for summary rows so summary tables do not get blank Image columns."""
    row_count = len(rows or [])
    imgs = [p for p in (available_images or []) if p and os.path.exists(p)]
    if not imgs:
        return [None] * row_count
    out = []
    for i in range(row_count):
        out.append(imgs[i % len(imgs)])
    return out




def _first_source_sentence(lesson: dict, max_words: int = 42) -> str:
    source = _sanitize_text((lesson.get("transcript", "") or "") + " " + (lesson.get("pagetext", "") or ""))
    for sent in re.split(r"(?<=[.!?])\s+", source):
        sent = _sanitize_text(sent)
        if len(sent.split()) >= 6:
            words = sent.split()
            return " ".join(words[:max_words]).rstrip(".,;:") + "."
    return f"This lesson introduces the core concept of {lesson.get('name', 'this topic')} and connects it to the chapter's larger learning goals."


def _fallback_content_slide_for_lesson(lesson: dict) -> dict:
    """Create a source-grounded fallback content slide if AI omits lesson content."""
    body = _first_source_sentence(lesson, 42)
    anchor = _first_source_sentence(lesson, 20)
    return {
        "type": "concept",
        "title": lesson.get("name") or f"Lesson {lesson.get('id', '')}".strip(),
        "sub_label": "Core Idea",
        "body": body,
        "image_required": True,
        "visual_focus": lesson.get("name") or body,
        "transcript_anchor_text": anchor,
        "speaker_notes": f"Use this fallback slide to ensure the lesson '{lesson.get('name', '')}' is still covered with source-grounded content.",
        "coverage_fallback": True,
    }


COVERAGE_SLIDE_TYPES = {"concept", "table", "discussion_question", "discussion_answer"}
CORE_CONTENT_SLIDE_TYPES = {"concept", "table"}


def _tag_slides_with_lesson(slide_data: dict, lesson: dict) -> dict:
    """Attach parsed lesson ID/name to every generated slide for deterministic coverage checks."""
    for slide in slide_data.get("slides") or []:
        if isinstance(slide, dict):
            slide["lesson_id"] = str(lesson.get("id", ""))
            slide["lesson_name"] = lesson.get("name", "")
    return slide_data


def _ensure_lesson_content_coverage(slide_data: dict, lesson: dict, qa_report: dict = None) -> dict:
    """
    Hard coverage guard. The user-approved rule is that every parsed lesson ID
    must contribute at least one content/table/discussion slide, and summary or
    glossary alone never count. We keep a stricter production guard as well:
    if the AI did not create a core content slide (concept/table), insert a
    source-grounded fallback concept slide so the lesson is visibly taught.
    """
    slide_data = _tag_slides_with_lesson(slide_data, lesson)
    slides = slide_data.get("slides") or []
    has_any_lesson_slide = any((s or {}).get("type") in COVERAGE_SLIDE_TYPES for s in slides)
    has_core_content = any((s or {}).get("type") in CORE_CONTENT_SLIDE_TYPES for s in slides)

    if has_any_lesson_slide and has_core_content:
        return slide_data

    fallback = _fallback_content_slide_for_lesson(lesson)
    fallback["lesson_id"] = str(lesson.get("id", ""))
    fallback["lesson_name"] = lesson.get("name", "")
    slide_data["slides"] = [fallback] + slides
    if qa_report is not None:
        reason = "no lesson coverage slide" if not has_any_lesson_slide else "no core concept/table slide"
        qa_report.setdefault("coverage_guard", []).append({
            "lesson_id": lesson.get("id"),
            "lesson_name": lesson.get("name"),
            "action": f"Inserted fallback concept slide because AI returned {reason} for this lesson."
        })
        qa_report.setdefault("flags", []).append({
            "level": "COVERAGE_GUARD",
            "message": f"Inserted fallback concept slide for {lesson.get('id')} - {lesson.get('name')} because AI returned {reason}."
        })
    return slide_data


def _validate_all_lesson_ids_covered(lesson_outputs: list, lessons: list, qa_report: dict) -> None:
    """Fail fast if any parsed lesson ID did not actually contribute a covered slide."""
    covered = set()
    core_covered = set()
    for bundle in lesson_outputs or []:
        for slide in (bundle.get("slide_data", {}) or {}).get("slides", []) or []:
            lid = str((slide or {}).get("lesson_id") or "")
            stype = (slide or {}).get("type")
            if lid and stype in COVERAGE_SLIDE_TYPES:
                covered.add(lid)
            if lid and stype in CORE_CONTENT_SLIDE_TYPES:
                core_covered.add(lid)
    missing = [l for l in lessons if str(l.get("id")) not in covered]
    missing_core = [l for l in lessons if str(l.get("id")) not in core_covered]
    qa_report["lesson_coverage_validation"] = {
        "parsed_lesson_count": len(lessons),
        "covered_lesson_ids": sorted(covered),
        "core_content_covered_lesson_ids": sorted(core_covered),
        "missing_lesson_ids": [str(l.get("id")) for l in missing],
        "missing_core_content_lesson_ids": [str(l.get("id")) for l in missing_core],
        "rule": "Each parsed lesson ID must have at least one concept/table/discussion slide; each lesson also receives a concept/table fallback if missing core content."
    }
    if missing:
        names = "; ".join(f"{l.get('id')} - {l.get('name')}" for l in missing)
        raise RuntimeError("Coverage validation failed. Missing generated slide coverage for: " + names)

# Table splitting was explicitly rejected in review feedback.
# Tables are kept on a single slide; crowding is controlled by prompt rules,
# no-empty-column cleanup, concise cell limits, and consistent table typography.


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
            sd = _repair_table_slide_content(sd)
        if stype == "summary":
            sd["rows"] = (sd.get("rows") or [])[:3]
            sd["summary_statement"] = _sanitize_text(sd.get("summary_statement", ""))
            sd = _repair_table_slide_content(sd)
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

    # Existing PageText naming plus current chapter naming:
    # 11084_Plant_Cell_Wall_Pagetext_KS.docx
    has_pagetext = (
        "pagetext" in norm
        or ("page" in norm and "text" in norm)
    )

    # Existing transcript naming plus current chapter naming:
    # 11084_Plant_Cell_Wall_Script_KS.docx
    has_transcript = (
        "transcript" in norm
        or "transcription" in norm
        or "script" in norm
    )

    if has_pagetext and not has_transcript:
        return "pagetext"
    if has_transcript and not has_pagetext:
        return "transcript"

    # If both are present in an unusual filename, keep PageText priority only
    # when PageText is explicit. Otherwise leave ambiguous files untouched.
    if has_pagetext and "pagetext" in norm:
        return "pagetext"
    if has_transcript and "script" in norm and not has_pagetext:
        return "transcript"

    return None


def _clean_name_piece(piece: str) -> str:
    piece = os.path.splitext(os.path.basename(piece))[0]
    piece = re.sub(r'(?<!\d)\d{4,8}(?!\d)', ' ', piece)
    piece = re.sub(r'(?i)pagetext|page text|page-text|page_text|transcript|transcription|script|video|vid|mp4|final|draft|copy', ' ', piece)
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



def _extract_lesson_heading_from_pagetext(content: str, lesson_id: str) -> str:
    """Use the real PT Lesson/topic heading for this lesson, without using writer metadata."""
    if not content:
        return ""

    lines = [re.sub(r"\s+", " ", line).strip() for line in content.splitlines()]
    metadata_re = re.compile(r"^(writer|author|reviewer|prepared\s*by|created\s*by)\s*[:\-]", re.I)

    # Highest priority: explicit Lesson: heading in PageText.
    for line in lines[:30]:
        if not line or metadata_re.match(line):
            continue
        m = re.match(r"^lesson\s*[:\-]\s*(.+)$", line, re.I)
        if m:
            candidate = m.group(1).strip()
            if candidate and not re.fullmatch(r"[\d\W_]+", candidate):
                return _camel_to_title(candidate)

    # Fallback: first clean heading-like line after metadata.
    skip_terms = {"pagetext", "page text", "transcript", "copyright"}
    for line in lines[:30]:
        if not line or metadata_re.match(line):
            continue
        lower = line.lower()
        if any(term in lower for term in skip_terms):
            continue
        if re.fullmatch(r"[\d\W_]+", line):
            continue
        words = line.split()
        if 1 <= len(words) <= 14 and 2 <= len(line) <= 110:
            return _camel_to_title(line)

    return ""


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
                    "pt_lesson_heading": "",
                    "tmpdir": tmpdir,
                }
            else:
                _set_better_lesson_name(lessons[lesson_id], lesson_name)

            if doc_type == 'pagetext':
                lessons[lesson_id]['pagetext'] = content
                lessons[lesson_id]['has_pagetext'] = bool(content.strip())
                pt_heading = _extract_lesson_heading_from_pagetext(content, lesson_id)
                if pt_heading:
                    # Source of truth for concept/table slide headings: the lesson heading from this lesson's PT.
                    lessons[lesson_id]['name'] = pt_heading
                    lessons[lesson_id]['pt_lesson_heading'] = pt_heading
                else:
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
    # Glossary pages are generated after the requested slide budget and do not reduce
    # concept/table/discussion coverage requested in the tool.
    discussion_pairs = min(6, max(1, round(len(lessons) / 3)))
    reserves = 1 + 1 + 1  # cover + chapter overview + chapter summary
    discussion_reserve = discussion_pairs * 2
    available_concepts = max(len(lessons), total_slide_budget - reserves - discussion_reserve)

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
        f"reserves_without_glossary={reserves}, glossary_extra_pages_allowed={glossary_pages}, discussion_pairs={discussion_pairs}, "
        f"discussion_reserve={discussion_reserve}, available_concepts={available_concepts}, final_concepts={current}."
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


def _short_figure_legend(value: str, fallback: str = "Figure") -> str:
    words = re.findall(r"[A-Za-z0-9][A-Za-z0-9\-]*", _sanitize_text(value or ""))[:5]
    return " ".join(words).strip() or fallback


def _one_sentence(value: str, max_words: int = 22) -> str:
    text = re.sub(r"\s+", " ", _sanitize_text(value or "")).strip()
    if not text:
        return ""
    sentence = re.split(r"(?<=[.!?])\s+", text)[0].strip()
    words = sentence.split()[:max_words]
    sentence = " ".join(words).strip()
    if sentence and sentence[-1] not in ".!?":
        sentence += "."
    return sentence


def _slide_topic(slide: dict) -> str:
    if not isinstance(slide, dict):
        return ""
    for key in ("sub_label", "sub_title", "title", "summary_statement", "question", "answer_summary"):
        text = _sanitize_text(slide.get(key))
        if text:
            return text[:80]
    return ""


def _ensure_transition_captions(slide_data: dict) -> dict:
    slides = slide_data.get("slides", []) if isinstance(slide_data, dict) else []
    for idx, slide in enumerate(slides):
        if (slide or {}).get("type") != "concept":
            continue
        if _sanitize_text(slide.get("transition_caption")):
            slide["transition_caption"] = _one_sentence(slide.get("transition_caption"), max_words=20)
            continue
        next_slide = None
        for candidate in slides[idx + 1:]:
            if (candidate or {}).get("type") in {"concept", "table"}:
                next_slide = candidate
                break
            if (candidate or {}).get("type") in {"discussion_question", "discussion_answer", "summary"}:
                break
        if not next_slide:
            continue
        current_topic = _slide_topic(slide) or "this concept"
        next_topic = _slide_topic(next_slide) or "the next topic"
        slide["transition_caption"] = _one_sentence(
            f"We learned about {current_topic}; next, we will explore {next_topic}.",
            max_words=20
        )
    return slide_data


def _select_discussion_lesson_ids(lessons: list) -> set:
    if not lessons:
        return set()
    pair_count = min(6, max(1, round(len(lessons) / 3)))
    indexed = list(enumerate(lessons))
    indexed.sort(
        key=lambda item: len((item[1].get("transcript", "") + " " + item[1].get("pagetext", "")).split()),
        reverse=True
    )
    chosen_indexes = sorted(idx for idx, _lesson in indexed[:pair_count])
    return {str(lessons[idx].get("id")) for idx in chosen_indexes}


def _chapter_overview_fallback(chapter_name: str, lesson_recaps: list) -> dict:
    lesson_names = [_sanitize_text(item.get("name")) for item in lesson_recaps if _sanitize_text(item.get("name"))]
    first_topics = ", ".join(lesson_names[:3])
    if first_topics:
        body = f"This chapter introduces the major ideas behind {chapter_name}. Key themes include {first_topics}. Each lesson builds the foundation needed for the next concept."
    else:
        body = f"This chapter introduces the major ideas behind {chapter_name}. Each lesson builds the foundation needed for the next concept."
    return {
        "chapter_definition": _one_sentence(f"{chapter_name} introduces the key concepts and processes students will study in this chapter.", max_words=22),
        "overview_title": "Chapter Overview",
        "overview_body": body,
        "transition_caption": _one_sentence("We have seen the chapter focus; next, we will begin with the first lesson.", max_words=20),
        "figure_legend": _short_figure_legend(chapter_name, fallback="Chapter overview"),
        "speaker_notes": "Introduce the chapter's main themes before starting the first lesson."
    }



def _ai_visuals_enabled(api_key: str = None) -> bool:
    """Return whether polished AI visual generation should run.

    Default is AUTO: enabled whenever an OpenAI API key is available. Reads both
    env vars and Streamlit secrets so a business account key in secrets is used.
    """
    raw = str(os.getenv("JOVE_AI_VISUALS", "auto")).strip().lower()
    if raw in {"0", "false", "no", "off", "disabled"}:
        return False
    if raw in {"1", "true", "yes", "on", "enabled"}:
        return bool(_resolve_openai_api_key(api_key))
    return bool(_resolve_openai_api_key(api_key))


def _ai_visual_fail_on_error() -> bool:
    """Strict debug switch only. Production default favors finishing the deck."""
    return bool(_ai_visual_budget_settings()["fail_on_error"])


def _ai_visual_table_rows_enabled() -> bool:
    """Table-row AI upgrades are expensive; keep disabled unless explicitly needed."""
    return bool(_ai_visual_budget_settings()["table_rows_enabled"])


def _ai_visuals_state(qa_report: dict) -> dict:
    """Shared per-run image-generation budget state."""
    state = qa_report.setdefault("ai_visuals_state", {})
    if "started_at" not in state:
        state["started_at"] = time.time()
    state.setdefault("generated_count", 0)
    state.setdefault("skipped_count", 0)
    state.setdefault("failed_count", 0)
    state.setdefault("per_lesson_generated", {})
    state["settings"] = _ai_visual_budget_settings()
    return state


def _ai_visuals_remaining_budget(qa_report: dict, lesson_id: str = None) -> bool:
    state = _ai_visuals_state(qa_report)
    settings = _ai_visual_budget_settings()
    max_per_chapter = settings["max_per_chapter"]
    max_seconds = settings["max_seconds"]
    timeout_seconds = settings["timeout_seconds"]
    max_per_lesson = settings["max_per_lesson"]

    if max_per_chapter <= 0:
        state["budget_stop_reason"] = "AI visual chapter budget is 0"
        return False
    if state.get("generated_count", 0) >= max_per_chapter:
        state["budget_stop_reason"] = f"max image count reached ({max_per_chapter})"
        return False
    elapsed = time.time() - state.get("started_at", time.time())
    if max_seconds > 0 and elapsed >= max_seconds:
        state["budget_stop_reason"] = f"max image time reached ({max_seconds}s)"
        return False
    # Do not start a new image if there is not enough remaining budget for one timeout window.
    if max_seconds > 0 and elapsed + timeout_seconds > max_seconds:
        state["budget_stop_reason"] = f"not enough AI visual budget left for another image ({int(max_seconds - elapsed)}s remaining)"
        return False
    if lesson_id:
        per_lesson = state.setdefault("per_lesson_generated", {})
        if max_per_lesson >= 0 and per_lesson.get(str(lesson_id), 0) >= max_per_lesson:
            state["budget_stop_reason"] = f"max image count for lesson reached ({max_per_lesson})"
            return False
    return True


def _visual_role_for_slide(slide_def: dict) -> str:
    """Classify image intent so all visuals are generated in the correct style.

    This encodes the Natural Selection feedback: not every slide needs the
    same kind of image. Cover/concept visuals should be premium hero visuals;
    process/table visuals should be clean scientific explanatory illustrations.
    """
    stype = str((slide_def or {}).get("type") or "").strip().lower()
    table_kind = str((slide_def or {}).get("table_kind") or "").strip().lower()
    title = _sanitize_text((slide_def or {}).get("title") or (slide_def or {}).get("sub_title") or "").lower()
    focus = _sanitize_text((slide_def or {}).get("visual_focus") or "").lower()
    text = f"{title} {focus}"
    if stype in {"table_row_visual", "table_cell_visual"}:
        return "table_cell_scientific_example"
    if stype == "table":
        if table_kind in {"process", "timeline"} or any(k in text for k in ["process", "step", "sequence", "cycle", "pathway"]):
            return "clean_process_diagram"
        if table_kind in {"comparison", "definition_example", "cause_effect"}:
            return "compact_table_explanation"
        return "compact_table_explanation"
    if stype in {"discussion_question", "discussion_answer"}:
        return "supporting_discussion_visual"
    if any(k in text for k in ["process", "step", "sequence", "cycle", "pathway", "mechanism"]):
        return "clean_process_diagram"
    if any(k in text for k in ["compare", "comparison", "versus", "vs", "types", "levels", "stages"]):
        return "side_by_side_comparison"
    return "hero_educational_visual"


def _visual_style_block(role: str) -> str:
    role = role or "hero_educational_visual"
    if role == "hero_educational_visual":
        return """Visual style for this slide:
- Premium educational hero visual, like the Natural Selection animal panel examples.
- Strong subject focus with relevant biology context or environment.
- Cinematic, high-resolution, clean composition; visually attractive but scientifically serious.
- Rich but controlled colors; no clutter; no random decorative objects.
- Should feel like a polished textbook/science-explainer hero image, not stock filler."""
    if role == "supporting_discussion_visual":
        return """Visual style for this slide:
- Clean supporting concept visual for a discussion slide.
- Visually interesting but not distracting; leave conceptual room for the question.
- One clear central idea, clean background, professional biology lecture tone."""
    if role == "clean_process_diagram":
        return """Visual style for this slide:
- Clean scientific explanatory process diagram, like the bacteria selection sequence reference.
- Clear step-by-step visual progression, separated stages, simple arrows if needed.
- White/light background, minimal clutter, readable at lecture-slide size.
- Simplified but biologically accurate; polished textbook-quality diagram."""
    if role == "side_by_side_comparison":
        return """Visual style for this slide:
- Clear side-by-side comparison visual.
- Each compared item must be visually distinct and easy to understand.
- Balanced layout, clean background, no tiny text, no unnecessary labels."""
    if role == "table_cell_scientific_example":
        return """Visual style for this table cell:
- Compact scientific example illustration optimized for a small table cell.
- One central subject only; simple background; high contrast and instantly readable.
- Similar clarity to the Natural Selection table example images.
- Do not create a busy hero scene; do not add text labels unless essential."""
    return """Visual style for this table/concept visual:
- Compact polished scientific illustration that supports the table or concept.
- Clear central subject, simple background, easy to read at slide size.
- Presentation-ready and non-redundant."""


def _ai_visual_prompt(lesson: dict, slide_def: dict) -> str:
    title = _sanitize_text(slide_def.get("title") or slide_def.get("sub_title") or lesson.get("name") or "")
    focus = _sanitize_text(slide_def.get("visual_focus") or slide_def.get("transcript_anchor_text") or "")
    role = _visual_role_for_slide(slide_def)
    style_block = _visual_style_block(role)
    return f"""Create an absolutely high-quality, high-resolution, presentation-ready biology lecture visual using the reference video frame as scientific context.

Topic: {lesson.get('name', '')}
Slide title: {title}
Visual focus: {focus}
Image role: {role}

{style_block}

Universal requirements:
- Use the reference frame for scientific subject/context, but do not output a raw screenshot or simple filtered screenshot.
- Convert the reference into a polished educational visual suitable for a JoVE lecture slide.
- The final image must look intentionally designed, not low-effort frame enhancement.
- Remove watermarks, UI artifacts, borders, awkward crops, and tiny unreadable labels.
- Avoid redundant/repeated-looking visuals; vary composition when nearby slides cover related topics.
- Avoid repeated generic protein-ribbon imagery unless the specific slide is about protein structure.
- Do not add decorative clutter or unrelated objects.
- Do not add text labels unless absolutely necessary for scientific clarity.
- Use a clean white or light background when that improves lecture-slide readability.
- Output must be visually strong enough for a professional deck used at very large student scale.
""".strip()


def _generate_ai_visual_from_frame(frame_path: str, lesson: dict, slide_def: dict, api_key: str, qa_report: dict = None) -> str:
    """Budgeted paid visual upgrade: video frame -> presentation-grade illustration.

    The deck must finish. AI visual generation is therefore capped and timeout-
    protected. If the cap is reached or one image fails, the selected JoVE frame
    is kept; ppt_builder still applies local presentation cleanup. Set
    JOVE_AI_VISUALS_FAIL_ON_ERROR=1 only when intentionally debugging image API
    failures.
    """
    if not _ai_visuals_enabled(api_key) or not frame_path or not os.path.exists(frame_path):
        return frame_path
    if qa_report is None:
        qa_report = {}
    state = _ai_visuals_state(qa_report)
    if not _ai_visuals_remaining_budget(qa_report, str(lesson.get("id", ""))):
        state["skipped_count"] = state.get("skipped_count", 0) + 1
        _append_runtime_event(qa_report, "ai_visual_skipped_budget", state.get("budget_stop_reason", "budget exhausted"), lesson_id=lesson.get("id"), slide_type=slide_def.get("type"))
        qa_report.setdefault("ai_visuals", []).append({
            "lesson_id": lesson.get("id"),
            "lesson_name": lesson.get("name"),
            "slide_type": slide_def.get("type"),
            "slide_title": slide_def.get("title"),
            "source_frame": frame_path,
            "status": "skipped_budget",
            "reason": state.get("budget_stop_reason", "budget exhausted"),
        })
        return frame_path
    try:
        from openai import OpenAI
        prompt = _ai_visual_prompt(lesson, slide_def)
        cache_dir = os.path.join(tempfile.gettempdir(), "jove_ai_visuals")
        os.makedirs(cache_dir, exist_ok=True)
        key_raw = (os.path.abspath(frame_path) + "|" + prompt).encode("utf-8", errors="ignore")
        key = hashlib.sha256(key_raw).hexdigest()[:24]
        out_path = os.path.join(cache_dir, f"{key}.png")
        if os.path.exists(out_path):
            _append_runtime_event(qa_report, "ai_visual_cache_hit", f"Using cached AI visual for {lesson.get('name')} / {slide_def.get('type')}", lesson_id=lesson.get("id"), slide_type=slide_def.get("type"))
            return out_path
        timeout_seconds = _ai_visual_budget_settings()["timeout_seconds"]
        _append_runtime_event(qa_report, "ai_visual_start", f"Starting AI visual for {lesson.get('name')} / {slide_def.get('type')}", lesson_id=lesson.get("id"), slide_type=slide_def.get("type"), timeout_seconds=timeout_seconds, settings=_ai_visual_budget_settings())
        client = OpenAI(api_key=api_key, timeout=timeout_seconds)
        model_name = os.getenv("JOVE_IMAGE_MODEL", "gpt-image-2")
        with open(frame_path, "rb") as ref_image:
            result = client.images.edit(
                model=model_name,
                image=[ref_image],
                prompt=prompt,
                size=os.getenv("JOVE_IMAGE_SIZE", "1536x1024"),
                quality=os.getenv("JOVE_IMAGE_QUALITY", "high"),
            )
        image_base64 = result.data[0].b64_json
        with open(out_path, "wb") as f:
            f.write(base64.b64decode(image_base64))
        state["generated_count"] = state.get("generated_count", 0) + 1
        per_lesson = state.setdefault("per_lesson_generated", {})
        lid = str(lesson.get("id", ""))
        per_lesson[lid] = per_lesson.get(lid, 0) + 1
        _append_runtime_event(qa_report, "ai_visual_generated", f"Generated AI visual for {lesson.get('name')} / {slide_def.get('type')}", lesson_id=lesson.get("id"), slide_type=slide_def.get("type"), generated_count=state.get("generated_count"))
        qa_report.setdefault("ai_visuals", []).append({
            "lesson_id": lesson.get("id"),
            "lesson_name": lesson.get("name"),
            "slide_type": slide_def.get("type"),
            "slide_title": slide_def.get("title"),
            "source_frame": frame_path,
            "generated_image": out_path,
            "model": model_name,
            "status": "generated"
        })
        return out_path
    except Exception as e:
        state["failed_count"] = state.get("failed_count", 0) + 1
        _append_runtime_event(qa_report, "ai_visual_failed", f"AI visual failed; kept selected JoVE frame. Error: {str(e)[:250]}", lesson_id=lesson.get("id"), slide_type=slide_def.get("type"))
        qa_report.setdefault("ai_visuals", []).append({
            "lesson_id": lesson.get("id"),
            "lesson_name": lesson.get("name"),
            "slide_type": slide_def.get("type"),
            "slide_title": slide_def.get("title"),
            "source_frame": frame_path,
            "status": "failed_kept_frame",
            "error": str(e)[:500]
        })
        if _ai_visual_fail_on_error():
            raise RuntimeError(
                f"AI visual upgrade failed for {lesson.get('id')} - {lesson.get('name')} / {slide_def.get('type')}. "
                "Set JOVE_AI_VISUALS_FAIL_ON_ERROR=0 to finish with selected JoVE frames when an image upgrade fails."
            ) from e
        return frame_path


def _upgrade_frame_map_with_ai_visuals(frame_map: dict, lesson: dict, slide_defs: list, openai_api_key: str, qa_report: dict, progress_callback=None) -> dict:
    if not _ai_visuals_enabled(openai_api_key):
        return frame_map
    for slide_index, info in list((frame_map or {}).items()):
        if not isinstance(info, dict):
            continue
        slide_def = slide_defs[slide_index] if slide_index < len(slide_defs) else {}
        src = info.get("path")
        if progress_callback:
            state = _ai_visuals_state(qa_report)
            progress_callback(f"AI visual upgrade {state.get('generated_count', 0)+1}/{str(_ai_visual_budget_settings()['max_per_chapter'])} for {lesson['name']}...", None)
        upgraded = _generate_ai_visual_from_frame(src, lesson, slide_def, openai_api_key, qa_report)
        if upgraded != src:
            info["original_frame_path"] = src
            info["path"] = upgraded
            info["visual_upgrade"] = "ai_generated_from_video_frame"
        else:
            info["visual_upgrade"] = info.get("visual_upgrade") or "selected_video_frame_after_ai_budget_or_failure"
    return frame_map


def _upgrade_table_row_images_with_ai_visuals(table_row_frame_map: dict, lesson: dict, slide_defs: list, openai_api_key: str, qa_report: dict, progress_callback=None) -> dict:
    if not _ai_visuals_enabled(openai_api_key) or not _ai_visual_table_rows_enabled():
        if table_row_frame_map:
            qa_report.setdefault("ai_visuals", []).append({
                "lesson_id": lesson.get("id"),
                "lesson_name": lesson.get("name"),
                "status": "table_row_ai_skipped",
                "reason": "JOVE_AI_VISUALS_TABLE_ROWS is off so the deck finishes faster; selected JoVE row frames are still used when valid."
            })
        return table_row_frame_map
    for slide_index, paths in list((table_row_frame_map or {}).items()):
        slide_def = slide_defs[slide_index] if slide_index < len(slide_defs) else {}
        upgraded_paths = []
        for pth in paths or []:
            if progress_callback:
                progress_callback(f"AI table-row visual upgrade for {lesson['name']}...", None)
            upgraded_paths.append(_generate_ai_visual_from_frame(pth, lesson, slide_def, openai_api_key, qa_report))
        table_row_frame_map[slide_index] = upgraded_paths
    return table_row_frame_map


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
                "type": "table_row_visual",
                "title": slide_def.get("sub_title") or slide_def.get("title") or lesson["name"],
                "visual_focus": row_text or slide_def.get("visual_focus") or lesson["name"],
                "transcript_anchor_text": row_text or slide_def.get("transcript_anchor_text") or "",
                "table_kind": slide_def.get("table_kind") or "",
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




def _add_glossary_term(glossary: dict, term: str, definition: str, max_terms: int = None) -> None:
    term = _sanitize_text(term)
    definition = _sanitize_text(definition)
    if not term or not definition:
        return
    if len(term.split()) > 5 or len(term) > 60:
        return
    if term.lower().strip() in GENERIC_GLOSSARY_TERMS:
        return
    # Prefer specific technical/scientific terms; reject broad generic plurals that make weak glossary entries.
    if term.lower().strip() in {"bacteria", "plants", "animals"}:
        return
    definition = _complete_sentence_text(definition, max_words=18) or definition
    key_norm = term.lower()
    if any(existing.lower() == key_norm for existing in glossary.keys()):
        return
    if max_terms is not None and len(glossary) >= max_terms:
        return
    glossary[term] = definition


def _extract_backup_glossary_terms(lessons: list, existing: dict, target_terms: int) -> dict:
    """Source-grounded glossary backup from PT/transcript only."""
    glossary = dict(existing or {})
    if len(glossary) >= target_terms:
        return glossary

    # 1) Lesson headings from PT are valid glossary candidates.
    for lesson in lessons:
        if len(glossary) >= target_terms:
            break
        heading = lesson.get("name", "")
        pt = lesson.get("pagetext", "")
        first_sentence = ""
        for sent in re.split(r"(?<=[.!?])\s+", _sanitize_text(pt)):
            if sent and len(sent.split()) >= 6:
                first_sentence = sent
                break
        if heading and first_sentence:
            _add_glossary_term(glossary, heading, first_sentence[:220], target_terms)

    # 2) Explicit definition patterns from source text.
    patterns = [
        r"\b([A-Z][A-Za-z0-9\-\s]{2,45})\s+are\s+([^.!?]{20,180})[.!?]",
        r"\b([A-Z][A-Za-z0-9\-\s]{2,45})\s+is\s+([^.!?]{20,180})[.!?]",
        r"\b([A-Za-z][A-Za-z0-9\- ]{2,45})\s+are called\s+([^.!?]{10,160})[.!?]",
        r"\b([A-Za-z][A-Za-z0-9\- ]{2,45})\s+is called\s+([^.!?]{10,160})[.!?]",
    ]
    for lesson in lessons:
        if len(glossary) >= target_terms:
            break
        source = _sanitize_text((lesson.get("pagetext", "") or "") + " " + (lesson.get("transcript", "") or ""))
        for pattern in patterns:
            if len(glossary) >= target_terms:
                break
            for m in re.finditer(pattern, source):
                if len(glossary) >= target_terms:
                    break
                term = re.sub(r"^(The|A|An)\s+", "", m.group(1).strip()).strip()
                definition = m.group(2).strip()
                if 1 <= len(term.split()) <= 4:
                    _add_glossary_term(glossary, term, definition, target_terms)

    # 3) Heading-like PT section lines as last source-grounded fallback.
    for lesson in lessons:
        if len(glossary) >= target_terms:
            break
        lines = [re.sub(r"\s+", " ", line).strip() for line in (lesson.get("pagetext", "") or "").splitlines()]
        for idx, line in enumerate(lines[:80]):
            if len(glossary) >= target_terms:
                break
            if not line or re.match(r"^(writer|author|reviewer|lesson)\s*[:\-]", line, re.I):
                continue
            if 1 <= len(line.split()) <= 5 and 3 <= len(line) <= 55 and not line.endswith("."):
                definition = ""
                for next_line in lines[idx + 1: idx + 5]:
                    if len(next_line.split()) >= 6:
                        definition = next_line
                        break
                if definition:
                    _add_glossary_term(glossary, line, definition[:220], target_terms)

    return glossary


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

    chapter_name = _sanitize_text(chapter_name) or "Chapter"
    run_started_at = time.time()

    def progress(msg, pct=None):
        elapsed = int(time.time() - run_started_at)
        print(f"[{pct if pct is not None else '?':>3}%] +{elapsed}s {msg}")
        try:
            _append_runtime_event(qa_report, "progress", msg, pct=pct, elapsed_seconds=elapsed)
        except Exception:
            pass
        if progress_callback:
            progress_callback(msg, pct)

    openai_api_key = _resolve_openai_api_key(openai_api_key)
    diagnostics_path = _new_runtime_diag_path(chapter_name, chapter_number)

    qa_report = {
        "chapter": chapter_name,
        "chapter_number": chapter_number,
        "generated_at": datetime.now().isoformat(),
        "pipeline_version": PIPELINE_VERSION,
        "runtime_diagnostics_path": diagnostics_path,
        "runtime_seconds": None,
        "ai_visual_budget_settings": _ai_visual_budget_settings(),
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
        "coverage_guard": [],
        "table_crowding_guard": [],
        "ai_visuals": [],
        "image_rule": "JoVE MP4 frame first. AI visual upgrade is enabled by default when an OpenAI API key is available, but capped/time-protected so the deck finishes. High-impact frames are upgraded into polished Natural Selection-style educational illustrations. Remaining frames are kept from JoVE video and locally presentation-enhanced. No web search. No placeholders."
    }
    _append_runtime_event(qa_report, "run_started", f"Started {chapter_name} chapter {chapter_number}", pipeline_version=PIPELINE_VERSION, ai_visual_budget_settings=_ai_visual_budget_settings(), has_openai_key=bool(openai_api_key))

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
        fallback_discussion_pairs = min(6, max(1, round(len(populated) / 3)))
        fallback_concept = max(1, min(4, round((chapter_budget - 3 - (fallback_discussion_pairs * 2)) / len(populated))))
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
    discussion_lesson_ids = _select_discussion_lesson_ids(populated)
    qa_report["planning"]["discussion_lesson_ids"] = sorted(discussion_lesson_ids)
    qa_report["planning"]["discussion_pair_cap"] = 6
    qa_report["planning"]["discussion_pair_count"] = len(discussion_lesson_ids)
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

        _append_runtime_event(qa_report, "lesson_ai_content_start", f"Generating slide content for {lesson['name']}", lesson_id=lesson.get("id"))
        slide_data = generate_slide_content(
            lesson_name=lesson['name'],
            transcript=lesson['transcript'],
            pagetext=lesson['pagetext'],
            concept_slide_budget=concept_budget,
            api_key=openai_api_key,
            model=model,
            include_discussion=str(lesson.get("id")) in discussion_lesson_ids
        )
        _append_runtime_event(qa_report, "lesson_ai_content_done", f"Generated slide content for {lesson['name']}", lesson_id=lesson.get("id"), raw_slide_count=len(slide_data.get("slides", [])) if isinstance(slide_data, dict) else None)
        slide_data = _normalize_slide_data_for_formatting(slide_data)
        slide_data = _ensure_lesson_content_coverage(slide_data, lesson, qa_report)
        slide_data = _ensure_transition_captions(slide_data)

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

        _append_runtime_event(qa_report, "frame_selection_start", f"Selecting frames for {lesson['name']}", lesson_id=lesson.get("id"), image_slide_count=image_slide_count)
        frame_map = assign_frames_to_slides(
            lesson=lesson,
            slide_defs=slide_data.get("slides", []),
            api_key=openai_api_key,
            vision_model=vision_model,
            progress_callback=progress
        )

        _append_runtime_event(qa_report, "frame_selection_done", f"Selected main frames for {lesson['name']}", lesson_id=lesson.get("id"), selected_count=len(frame_map or {}))
        table_row_frame_map = _build_table_row_frame_map(
            lesson=lesson,
            slide_defs=slide_data.get("slides", []),
            openai_api_key=openai_api_key,
            vision_model=vision_model,
            progress_callback=progress
        )

        if _ai_visuals_enabled(openai_api_key):
            progress(f"Upgrading selected video frames into polished visuals for {lesson['name']}...", pct + 3)
            _append_runtime_event(qa_report, "ai_visual_lesson_start", f"AI visual pass for {lesson['name']}", lesson_id=lesson.get("id"), settings=_ai_visual_budget_settings())
            frame_map = _upgrade_frame_map_with_ai_visuals(frame_map, lesson, slide_data.get("slides", []), openai_api_key, qa_report, progress_callback=progress)
            table_row_frame_map = _upgrade_table_row_images_with_ai_visuals(table_row_frame_map, lesson, slide_data.get("slides", []), openai_api_key, qa_report, progress_callback=progress)

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

    _validate_all_lesson_ids_covered(lesson_outputs, populated, qa_report)

    try:
        chapter_overview = generate_chapter_overview(chapter_name, lesson_recaps, openai_api_key, model=model)
    except Exception as e:
        chapter_overview = _chapter_overview_fallback(chapter_name, lesson_recaps)
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"Chapter overview generation failed ({str(e)}); using fallback."
        })

    # Build deck
    prs = create_presentation(LOGO_PATH)
    slide_count = 0

    progress("Building cover slide...", 75)
    build_cover_slide(
        prs, chapter_name, chapter_number, LOGO_PATH,
        cover_image_path=cover_image_path, cover_image_paths=cover_image_paths,
        chapter_description=chapter_overview.get("chapter_definition"),
        slide_number=slide_count + 1
    )
    slide_count += 1

    progress("Building chapter overview slide...", 76)
    build_concept_slide(
        prs,
        lesson_name=chapter_overview.get("overview_title") or "Chapter Overview",
        body_text=chapter_overview.get("overview_body") or f"This chapter introduces the major ideas in {chapter_name}.",
        sub_label="Chapter Introduction",
        image_path=cover_image_path,
        figure_legend=chapter_overview.get("figure_legend"),
        transition_caption=chapter_overview.get("transition_caption"),
        speaker_notes=chapter_overview.get("speaker_notes"),
        logo_path=LOGO_PATH,
        slide_number=slide_count + 1,
        allow_no_image=True
    )
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
            "video_path": lesson.get("video_path"),
            "slide_types": [],
            "lesson_id_coverage_met": False,
            "content_coverage_met": False
        }

        progress(f"Building slides for {lesson['name']}...", 77 + int((i / max(1, len(lesson_outputs))) * 10))

        for slide_index, slide_def in enumerate(slide_data.get("slides", [])):
            stype = slide_def.get("type", "concept")
            lesson_qa["slide_types"].append(stype)
            if stype in COVERAGE_SLIDE_TYPES:
                lesson_qa["lesson_id_coverage_met"] = True
            if stype in CORE_CONTENT_SLIDE_TYPES:
                lesson_qa["content_coverage_met"] = True
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
                    figure_legend=slide_def.get("figure_legend"),
                    transition_caption=slide_def.get("transition_caption"),
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
                    row_image_paths=_ensure_table_row_images(table_row_frame_map.get(slide_index), slide_def.get("rows", []), img_path),
                    speaker_notes=notes, logo_path=LOGO_PATH, slide_number=slide_count + 1
                )
            elif stype == "discussion_question":
                build_discussion_question_slide(
                    prs, lesson_name=title,
                    question_text=slide_def.get("question", ""),
                    hint_text=slide_def.get("hint"),
                    image_path=img_path,
                    figure_legend=slide_def.get("figure_legend"),
                    speaker_notes=notes, logo_path=LOGO_PATH, slide_number=slide_count + 1
                )
            elif stype == "discussion_answer":
                build_discussion_answer_slide(
                    prs, lesson_name=title,
                    answer_summary=slide_def.get("answer_summary", ""),
                    answer_explanation=slide_def.get("answer_explanation", ""),
                    image_path=img_path,
                    figure_legend=slide_def.get("figure_legend"),
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
        chapter_summary = _repair_table_slide_content({
            "headers": chapter_summary.get("headers", ["Concept", "Definition", "Key Point"]),
            "rows": (chapter_summary.get("rows", []) or [])[:5],
        })
        summary_rows = chapter_summary.get("rows", [])[:5]
        summary_images = _summary_row_images(summary_rows, cover_image_paths or ([cover_image_path] if cover_image_path else []))
        build_summary_slide(
            prs,
            summary_statement=chapter_summary.get("summary_statement", f"{chapter_name} - key takeaways."),
            table_headers=chapter_summary.get("headers", ["Concept", "Key Point"]),
            table_rows=summary_rows,
            logo_path=LOGO_PATH,
            speaker_notes="Use this slide to recap the chapter's core concepts with students before moving on.",
            slide_number=slide_count + 1
        )
    except Exception as e:
        qa_report["flags"].append({
            "level": "WARNING",
            "message": f"Chapter summary generation failed ({str(e)}); using fallback."
        })
        fallback_rows = [[lr["name"], "Core lesson concept", _complete_sentence_text(lr.get("key_points", ""), 14)] for lr in lesson_recaps[:5]]
        fallback_rows = _repair_table_slide_content({"headers": ["Lesson", "Definition", "Key Idea"], "rows": fallback_rows}).get("rows", fallback_rows)
        fallback_images = _summary_row_images(fallback_rows, cover_image_paths or ([cover_image_path] if cover_image_path else []))
        build_summary_slide(
            prs,
            summary_statement=f"{chapter_name} - key concepts covered in this chapter.",
            table_headers=["Lesson", "Key Idea"],
            table_rows=fallback_rows,
            logo_path=LOGO_PATH,
            speaker_notes="Recap the chapter's core concepts with students.",
            slide_number=slide_count + 1
        )
    slide_count += 1

    # Glossary
    progress("Building glossary...", 92)
    TERMS_PER_PAGE = 9
    planned_glossary_pages = int(qa_report["planning"].get("glossary_pages", 2) or 2)
    glossary_pages = max(1, min(3, planned_glossary_pages + 2))
    max_terms = TERMS_PER_PAGE * glossary_pages

    # Ensure the glossary is not tiny in multi-lesson decks. Use only source-grounded terms.
    if len(populated) >= 7:
        target_terms = min(max_terms, 18)
    elif len(populated) >= 4:
        target_terms = min(max_terms, 12)
    else:
        target_terms = min(max_terms, 9)

    all_glossary = _extract_backup_glossary_terms(populated, all_glossary, target_terms)
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
            f"Strict image rule violated: {len(qa_report['images_missing'])} slide(s) did not receive MP4 frames. AI visual budget fallback is allowed only after a valid JoVE frame exists; missing MP4/frame coverage is still a hard failure."
        )

    if qa_report["scientific_names"]:
        qa_report["flags"].append({
            "level": "REVIEW",
            "message": f"Scientific names to verify: {', '.join(set(qa_report['scientific_names']))}"
        })

    qa_report["runtime_seconds"] = int(time.time() - run_started_at)
    _append_runtime_event(qa_report, "run_completed", f"Done: {slide_count} slides generated", runtime_seconds=qa_report["runtime_seconds"], output_file=out_path)
    progress(f"Done! {slide_count} slides generated (target was {chapter_budget}).", 100)
    return out_path, qa_report
