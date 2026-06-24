"""
JoVE AI Content Generator
Sends lesson content to OpenAI and receives structured slide JSON.

Key update:
- Definition-first sequencing is enforced.
- Tables must include Definition/Meaning + Example/Application columns when introducing key terms/types/stages when those columns can be meaningfully filled.
- If a Definition/Meaning or Example/Application column would be blank or weak for that specific table, the table may use fewer text columns instead of forcing an empty column.
- Each image-bearing slide includes transcript_anchor_text and visual_focus so the video-frame picker can find the matching moment in the lesson MP4 and the image-upgrade step can create a premium visual from that frame.
"""

import json
import re
from typing import Dict, Any
from openai import OpenAI


SYSTEM_PROMPT = """You are an expert educational content designer for JoVE (Journal of Visualized Experiments), a scientific video platform. Your job is to convert lesson transcripts into structured slide content for lecture presentations.

SOURCE PRIORITY — THIS IS CRITICAL:
- The TRANSCRIPT is your PRIMARY and MAIN source. Build slide content, structure, and language from the transcript first.
- The PAGE TEXT is SUPPORTING ONLY — use it to add scientific terminology, precise definitions, or fill small gaps the transcript doesn't cover. Never let page text override or replace transcript content. If the transcript covers a topic, use the transcript's framing even if page text phrases it differently.

CONTENT ORDER — CRITICAL FEEDBACK FIX:
1. Start every lesson with a definition/concept-setting slide before examples.
2. Never jump directly into examples before definitions.
3. Tables are contextual, but key term/type/stage tables should FIRST try to use Definition/Meaning + Example/Application:
   - For key terms, types, stages, steps, methods, variables, and classes, include Definition/Meaning and Example/Application when both columns can be meaningfully filled for every row.
   - Skip/collapse a Definition/Meaning or Example/Application column ONLY when it is not possible to fill that column accurately from the transcript/page text without making it blank, repetitive, or weak.
   - For comparison, cause/effect, input/output, timeline, or simple process tables, use the most natural text columns instead of forcing 3 columns.
   - Do not add an Image column in JSON; the builder adds the Image column and row images separately.
4. Table slide titles/subtitles must describe the table content, not just repeat a generic label.
5. Before explaining an example, add the general "how it works" or "method/process" explanation.
6. Keep wording student-facing and useful.
7. Table text must not be too short: non-heading text cells should usually be 8-24 words.
8. Never leave required table text cells blank. If a table cannot support multiple text columns, use one strong explanatory column instead.
9. Avoid crowded tables. Do NOT assume the renderer will split tables into two slides. Keep each table readable on one slide by using no more than 3-4 rows, concise but explanatory cells, and only columns that can be fully filled. If more material is needed, cover it in concept slides or speaker notes rather than making a crowded table.

STRICT RULES:
0. Never add writer name, author name, reviewer name, prepared-by text, individual credit lines, or any metadata copied from source files. If source text contains such lines, ignore them completely.
1. Each slide covers ONE parent concept only. Related subtypes may appear together only when the parent concept is the slide concept. Never mix unrelated concepts on one slide.
2. All scientific names (genus and species) MUST be formatted as: **_Genus species_** (italic, genus capitalized, species lowercase).
3. Concept slides should use the lesson name from that lesson's PageText/PT unless the lesson name is generic or numeric. Table slides may use descriptive table titles/subtitles that clearly explain the table content; do NOT force table titles to exactly match the PageText lesson name when a descriptive title is clearer. Discussion slides keep the visible "Discussion" layout in the PPT builder.
4. Discussion questions MUST use two separate slides: question on one slide, answer on the next. The answer must never appear on the question slide.
5. Generate speaker notes for every slide using conversational, transcript-style language — as if the presenter is talking through the transcript's explanation.
6. If content has types, conditions, stages, or comparisons (2 or more), generate a table when useful. For key terms/types/stages/steps, use Definition/Meaning + Example/Application when possible, but skip/collapse a column if it would otherwise be blank or weak. Never create a blank Definition/Meaning or Example/Application column.
7. Body text per slide: maximum 4 lines / 3 distinct points. Split across multiple slides if the transcript covers more.
8. Discussion slide JSON titles may include the specific topic for internal context, but the visible PPT heading must remain exactly "Discussion".
9. glossary_terms must include every keyword, scientific term, and defined concept mentioned in the transcript for this lesson — comprehensive but only terms actually discussed.
10. For every slide that requires an image, provide:
    - image_required: true
    - visual_focus: a precise description of the best matching video-frame reference to transform into a premium, polished educational visual
    - transcript_anchor_text: a short exact or near-exact phrase from the transcript that tells where in the video this slide concept occurs
11. Image planning must match the final visual style target from the approved examples:
    - Concept/opening visuals should be premium educational hero visuals: clean composition, strong subject focus, beautiful but scientifically relevant scene, high-resolution, presentation-ready.
    - Process/mechanism visuals should be clean scientific diagrams: simplified, step-by-step, white/light background, clear arrows or progression only when helpful.
    - Table-row visuals should be compact scientific example illustrations: one clear central subject, minimal clutter, readable inside a small table cell.
    - Avoid redundant visuals across nearby slides. Do not ask for the same generic protein ribbon or repeated molecular screenshot unless the exact concept requires it.
    - Treat JoVE video frames as reference/context only; the final visual should be upgraded into a polished educational image.

SLIDE TYPES you can create:
- "concept": Main content slide with body text
- "table": Comparison/definition/stages table
- "discussion_question": Quiz question (NO answer visible)
- "discussion_answer": Answer to the previous quiz question
- "summary": A lesson-level recap slide for STUDENTS (only generate if the lesson genuinely has multiple sub-concepts that benefit from a recap — not every lesson needs one)

OUTPUT: Return ONLY valid JSON, no markdown, no explanation. Use this exact schema:

{
  "lesson_name": "...",
  "slides": [
    {
      "type": "concept",
      "title": "exact lesson name from PT/PageText",
      "sub_label": "Definition / Core Idea / How It Works / Example",
      "body": "Content text from the TRANSCRIPT. Use **bold** for key terms. Max 4 lines.",
      "image_required": true,
      "visual_focus": "specific visual frame to pick from the lesson video",
      "transcript_anchor_text": "short exact or near-exact transcript phrase for timing alignment",
      "speaker_notes": "conversational notes paraphrasing the transcript's explanation"
    },
    {
      "type": "table",
      "title": "descriptive table title or lesson name",
      "sub_title": "descriptive subtitle for what this table shows",
      "table_kind": "definition_example | comparison | process | timeline | cause_effect | other",
      "headers": ["Use the most suitable text columns for this table. Do not include Image here."],
      "rows": [["row values matching the headers; no blank required text cells"]],
      "image_required": true,
      "visual_focus": "specific visual frame to pick from the lesson video",
      "transcript_anchor_text": "short exact or near-exact transcript phrase for timing alignment",
      "speaker_notes": "..."
    },
    {
      "type": "discussion_question",
      "title": "Discussion: <specific topic from this lesson>",
      "question": "The question text",
      "hint": "optional hint",
      "image_required": true,
      "visual_focus": "specific visual frame to pick from the lesson video",
      "transcript_anchor_text": "short exact or near-exact transcript phrase for timing alignment",
      "speaker_notes": "notes about facilitating discussion"
    },
    {
      "type": "discussion_answer",
      "title": "Discussion: <same specific topic, matches question slide>",
      "answer_summary": "short answer headline",
      "answer_explanation": "full explanation paragraph from transcript reasoning",
      "image_required": true,
      "visual_focus": "specific visual frame to pick from the lesson video",
      "transcript_anchor_text": "short exact or near-exact transcript phrase for timing alignment",
      "speaker_notes": "explanation for presenter"
    },
    {
      "type": "summary",
      "title": "exact lesson name",
      "summary_statement": "One sentence capturing this lesson's core takeaway, written FOR STUDENTS as a key learning point.",
      "headers": ["Concept", "Key Point"],
      "rows": [["concept1", "what students should remember"]],
      "image_required": false,
      "speaker_notes": "wrap-up notes"
    }
  ],
  "glossary_terms": {
    "Term": "Definition derived from the transcript"
  }
}"""


IMAGE_TYPES = {"concept", "table", "discussion_question", "discussion_answer"}


def _supports_json_response(model: str) -> bool:
    model = model or ""
    return not any(model.startswith(m) for m in ["o1", "o3", "o4"])


def _uses_completion_tokens(model: str) -> bool:
    model = model or ""
    return any(model.startswith(m) for m in ["gpt-5", "o1", "o3", "o4"])


def _chat_json(client: OpenAI, params: Dict[str, Any]) -> Dict[str, Any]:
    response = client.chat.completions.create(**params)
    raw = response.choices[0].message.content
    return json.loads(raw)


def _first_meaningful_transcript_phrase(transcript: str, max_len: int = 180) -> str:
    clean = re.sub(r"\s+", " ", transcript or "").strip()
    if not clean:
        return ""
    sentences = re.split(r"(?<=[.!?])\s+", clean)
    return (sentences[0] if sentences else clean)[:max_len]


def _table_needs_definition_example(slide: dict) -> bool:
    """
    Decide whether a table should be normalized to include Definition/Meaning
    and Example/Application columns.

    This is intentionally not universal. It only applies to semantic tables
    that introduce terms, steps, types, stages, methods, variables, or key concepts.
    """
    table_kind = str(slide.get("table_kind") or "").lower().strip()
    if table_kind in {"definition_example", "definitions", "terms", "steps", "types", "stages", "method", "variables"}:
        return True
    if table_kind in {"comparison", "timeline", "cause_effect", "pros_cons", "inputs_outputs", "other"}:
        return False

    text_blob = " ".join([
        str(slide.get("sub_title", "")),
        " ".join(str(h) for h in slide.get("headers", []) or []),
        " ".join(" ".join(str(x) for x in row) for row in slide.get("rows", []) or []),
    ]).lower()

    semantic_terms = [
        "term", "definition", "meaning", "step", "stage", "type", "method",
        "variable", "hypothesis", "observation", "concept", "key term"
    ]
    non_definition_structures = [
        "versus", " vs ", "compare", "difference", "timeline", "before", "after",
        "cause", "effect", "positive", "negative", "relationship", "pros", "cons"
    ]

    if any(x in text_blob for x in non_definition_structures) and not any(x in text_blob for x in ["definition", "meaning"]):
        return False
    return any(x in text_blob for x in semantic_terms)



def _clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("**", "").replace("__", "").replace("`", "")
    text = re.sub(r"\[(INSERT IMAGE|TODO|PLACEHOLDER|IMAGE)\]", "", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def _word_count(text: str) -> int:
    return len(re.findall(r"[A-Za-z0-9]+", str(text or "")))


def _is_blank_cell(value) -> bool:
    text = str(value or "").strip()
    return not text or text.lower() in {"n/a", "na", "none", "null", "-", "—"}


def _is_image_header(header: str) -> bool:
    h = str(header or "").strip().lower()
    return h in {"image", "visual", "figure", "picture"} or "image" in h or "visual" in h


def _make_cell_more_useful(text: str, header: str, row_label: str) -> str:
    """Light deterministic cleanup only. Do not invent outside facts."""
    text = str(text or "").strip()
    header_l = str(header or "").lower()
    if _is_blank_cell(text):
        return ""
    wc = _word_count(text)
    if wc >= 5:
        return text
    # Make very short fragments read as useful table cells without adding new scientific claims.
    if "example" in header_l or "application" in header_l or "key" in header_l:
        return f"Key point: {text}."
    if "definition" in header_l or "meaning" in header_l:
        return f"{row_label}: {text}."
    return text


def _drop_unusable_text_columns(headers: list, rows: list) -> tuple:
    """Remove only text columns that would create blank/weak visible table columns.
    This preserves Definition/Meaning + Example/Application for semantic tables whenever both are usable.
    """
    if not headers:
        return ["Key Point"], [[""]]

    # Remove image-like columns from AI JSON. PPT builder owns the Image column.
    keep_indices = [i for i, h in enumerate(headers) if not _is_image_header(h)]
    headers = [headers[i] for i in keep_indices]
    rows = [[(row[i] if i < len(row) else "") for i in keep_indices] for row in rows]

    if not headers:
        return ["Key Point"], [[""] for _ in rows]

    # Drop any non-first column that is blank for every row.
    keep = []
    for ci, h in enumerate(headers):
        col = [str(row[ci] if ci < len(row) else "").strip() for row in rows]
        all_blank = all(_is_blank_cell(x) for x in col)
        if ci == 0 or not all_blank:
            keep.append(ci)

    headers = [headers[i] for i in keep]
    rows = [[(row[i] if i < len(row) else "") for i in keep] for row in rows]

    # If any non-first column has blanks in some rows, drop that column instead of showing blanks.
    keep = []
    for ci, h in enumerate(headers):
        col = [str(row[ci] if ci < len(row) else "").strip() for row in rows]
        any_blank = any(_is_blank_cell(x) for x in col)
        if ci == 0 or not any_blank:
            keep.append(ci)

    headers = [headers[i] for i in keep]
    rows = [[(row[i] if i < len(row) else "") for i in keep] for row in rows]

    # If the first column itself has blanks, remove those rows rather than showing blank terms.
    cleaned_rows = []
    for row in rows:
        if row and not _is_blank_cell(row[0]):
            cleaned_rows.append(row)
    rows = cleaned_rows or rows

    return headers, rows


def _normalize_table_slide(slide: dict) -> dict:
    headers = [_clean_text(h) for h in (slide.get("headers") or [])]
    rows = [[_clean_text(x) for x in list(row)] for row in (slide.get("rows") or [])]

    if not headers:
        headers = ["Concept", "Key Point"]

    if not rows:
        rows = [["Core idea", _clean_text(slide.get("sub_title") or slide.get("title") or "")]]

    # Normalize row lengths first.
    n_cols = len(headers)
    norm_rows = []
    for row in rows:
        row = list(row)
        if len(row) < n_cols:
            row += [""] * (n_cols - len(row))
        norm_rows.append(row[:n_cols])

    headers, rows = _drop_unusable_text_columns(headers, norm_rows)

    # If only a label column remains, add one contextual Key Point column from the best available row text.
    if len(headers) == 1:
        headers = [headers[0] or "Concept", "Key Point"]
        rows = [[row[0], row[0]] for row in rows]

    # Make short non-label cells a little more useful without adding outside facts.
    repaired_rows = []
    for row in rows:
        label = _clean_text(row[0] if row else "")
        repaired = []
        for ci, h in enumerate(headers):
            val = row[ci] if ci < len(row) else ""
            if ci == 0:
                repaired.append(_clean_text(val))
            else:
                repaired.append(_make_cell_more_useful(val, h, label))
        repaired_rows.append(repaired)

    # Final guard: remove any text column that still has blank cells.
    headers, repaired_rows = _drop_unusable_text_columns(headers, repaired_rows)

    # Keep tables contextual: do not force Definition + Example columns if they don't fit.
    slide["headers"] = headers
    slide["rows"] = repaired_rows or [["Core idea", _clean_text(slide.get("sub_title") or slide.get("title") or "")]]
    slide["table_kind"] = slide.get("table_kind") or ("definition_example" if len(headers) >= 3 else "other")
    return slide


def _normalize_slide(slide: dict, lesson_name: str, transcript: str, is_first_content_slide: bool = False) -> dict:
    stype = slide.get("type", "concept")

    if stype == "concept":
        slide["title"] = lesson_name
    elif stype == "table":
        # Preserve descriptive table titles/subtitles. Do not force table titles
        # to the PT lesson heading when the table content is clearer.
        slide["title"] = _clean_text(slide.get("title") or slide.get("sub_title") or lesson_name)
        if not _clean_text(slide.get("sub_title")) and _clean_text(slide.get("title")) != lesson_name:
            slide["sub_title"] = _clean_text(slide.get("title"))

    if stype in IMAGE_TYPES:
        slide["image_required"] = True
        visual = (
            slide.get("visual_focus")
            or slide.get("image_query")
            or slide.get("body")
            or slide.get("question")
            or slide.get("answer_summary")
            or lesson_name
        )
        slide["visual_focus"] = str(visual).strip()[:300]
        anchor = str(slide.get("transcript_anchor_text") or "").strip()
        if not anchor:
            anchor = _first_meaningful_transcript_phrase(transcript) if is_first_content_slide else slide["visual_focus"][:180]
        slide["transcript_anchor_text"] = anchor
    else:
        slide["image_required"] = False

    if stype == "table":
        slide = _normalize_table_slide(slide)

    return slide

def _validate_slide_payload(payload: dict, lesson_name: str, transcript: str, concept_slide_budget: int = None) -> dict:
    slides = payload.get("slides", []) or []

    content_seen = False
    normalized = []
    for slide in slides:
        stype = slide.get("type", "concept")
        is_first_content = stype in {"concept", "table"} and not content_seen
        normalized_slide = _normalize_slide(slide, lesson_name, transcript, is_first_content_slide=is_first_content)
        normalized.append(normalized_slide)
        if stype in {"concept", "table"}:
            content_seen = True

    # Keep discussion Q+A as final pair.
    discussion_q = [s for s in normalized if s.get("type") == "discussion_question"]
    discussion_a = [s for s in normalized if s.get("type") == "discussion_answer"]
    other = [s for s in normalized if s.get("type") not in {"discussion_question", "discussion_answer"}]

    # Hard budget guard: the model is prompted to obey concept_slide_budget, but
    # this code enforces it so total slide count cannot drift because of extra
    # concept/table/summary slides. This is generic and not lesson-specific.
    if concept_slide_budget is not None:
        try:
            max_content = max(1, int(concept_slide_budget))
        except Exception:
            max_content = None
        if max_content is not None and len(other) > max_content:
            # Prefer keeping definition/core idea and table slides first. Summary
            # slides are dropped first when the budget is tight because the
            # chapter summary/glossary still provide recap coverage.
            non_summary = [s for s in other if s.get("type") != "summary"]
            summary = [s for s in other if s.get("type") == "summary"]
            other = (non_summary + summary)[:max_content]

    if not discussion_q:
        discussion_q = [_normalize_slide({
            "type": "discussion_question",
            "title": f"Discussion: {lesson_name}",
            "question": f"What is the core idea from {lesson_name}, and how would you explain it before giving an example?",
            "hint": "Start with the definition, then connect it to the example.",
            "visual_focus": lesson_name,
            "transcript_anchor_text": _first_meaningful_transcript_phrase(transcript),
            "speaker_notes": "Ask students to explain the concept in their own words before applying it."
        }, lesson_name, transcript)]

    if not discussion_a:
        discussion_a = [_normalize_slide({
            "type": "discussion_answer",
            "title": discussion_q[0].get("title", f"Discussion: {lesson_name}"),
            "answer_summary": "Definition first, then application",
            "answer_explanation": f"The strongest answer defines the concept from {lesson_name} first and then applies it to the lesson example.",
            "visual_focus": lesson_name,
            "transcript_anchor_text": _first_meaningful_transcript_phrase(transcript),
            "speaker_notes": "Reinforce the difference between defining a concept and applying it to an example."
        }, lesson_name, transcript)]

    payload["lesson_name"] = lesson_name
    payload["slides"] = other + [discussion_q[0], discussion_a[0]]
    payload["glossary_terms"] = payload.get("glossary_terms", {}) or {}
    return payload


def generate_slide_content(lesson_name: str, transcript: str, pagetext: str,
                            concept_slide_budget: int, api_key: str,
                            model: str = "gpt-4.1") -> dict:
    """
    concept_slide_budget: number of concept/table/summary slides to generate
    (discussion Q+A is ALWAYS added on top - exactly 2 more slides).
    """
    client = OpenAI(api_key=api_key)

    user_prompt = f"""Generate slide content for this lesson.

LESSON NAME FROM PT: {lesson_name}
Use this as the concept-slide title. For table slides, use a descriptive table title/subtitle when that is clearer than repeating the lesson name.

CONCEPT SLIDE BUDGET: {concept_slide_budget} slides (concept/table/summary combined).
This is IN ADDITION to the mandatory discussion_question + discussion_answer pair (2 slides),
which you must ALWAYS include at the end. Do not count Q&A toward this budget.

=== TRANSCRIPT (PRIMARY SOURCE - build content from this) ===
{transcript}

=== PAGE TEXT (SUPPORTING ONLY - use only to fill gaps or add precise terminology) ===
{pagetext if pagetext else "No page text available. Use transcript only."}

REMINDERS:
- First content slide MUST define the lesson's core concept before examples.
- If using a table for key terms/steps/types/stages, first try to include Definition/Meaning and Example/Application columns.
- Skip/collapse Definition/Meaning or Example/Application ONLY when that column cannot be accurately and usefully filled for every row from the transcript/page text.
- Table subtitles must be specific and suitable to the table; do not use generic headings.
- Table cells must be useful and explanatory, not tiny fragments. Most non-heading table cells should be 8-24 words.
- For table slides, do NOT include an Image column in JSON. Row-level images are added by the PPT builder automatically.
- Normal tables should usually be 3-4 rows and must be designed to fit one slide. Do not rely on table splitting. If more material is important, cover it with concise concept-slide text or speaker notes instead of creating a crowded table. Summary tables remain max 3 rows.
- Never include markdown syntax like **bold**, backticks, TODO, placeholders, or bracketed image instructions in any field.
- Concept slide titles should be the PT lesson name unless the PT name is generic/numeric; never use generic titles like "Definition / Core Idea" as the visible title.
- For processes/sequences such as hydrolysis, folding, polymerization, or bond formation, generate content suitable for a flowchart/step process rather than a long paragraph.
- Before explaining an example, include how the method/process works.
- Max 4 lines of body text per concept slide.
- Always end with discussion_question + discussion_answer.
- For video-frame selection, every image-bearing slide must include visual_focus and transcript_anchor_text.
- visual_focus should target a clean, presentation-grade educational visual moment. Avoid frames that would look like watermarked raw screenshots, repeated protein ribbons, tiny labels, awkward crops, or low-quality visuals when a better lesson-relevant frame exists. Prefer a frame that can be converted into a polished Natural Selection-style educational illustration.
- Explicitly plan visuals by intent: hero educational visual for major concept/opening slides, clean process diagram for mechanisms/sequences, compact table-cell illustration for row examples, and side-by-side comparison visual for comparisons.
- Avoid redundant visuals: if several slides discuss related molecular topics, vary the composition and instructional focus rather than repeating similar protein/molecule imagery.
- transcript_anchor_text should be a short exact or near-exact transcript phrase.
- glossary_terms: comprehensive for THIS lesson's transcript content."""

    params = dict(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ]
    )
    if _supports_json_response(model):
        params["response_format"] = {"type": "json_object"}

    if _uses_completion_tokens(model):
        params["max_completion_tokens"] = 7000
    else:
        params["max_tokens"] = 7000
        params["temperature"] = 0.3

    payload = _chat_json(client, params)
    return _validate_slide_payload(payload, lesson_name, transcript, concept_slide_budget=concept_slide_budget)


def generate_chapter_summary(chapter_name: str, lesson_summaries: list,
                              api_key: str, model: str = "gpt-4.1") -> dict:
    """
    Generates a STUDENT-FACING chapter summary — a recap table of the
    key concepts learned, matching the benchmark style, NOT a build report.
    """
    client = OpenAI(api_key=api_key)

    summaries_text = "\n\n".join(
        f"Lesson: {s['name']}\nKey points: {s.get('key_points', 'N/A')}"
        for s in lesson_summaries
    )

    prompt = f"""Create a STUDENT-FACING chapter summary for "{chapter_name}".

This summary will be the LAST content slide students see. It should recap the
core concepts they learned across the chapter — like a study guide, NOT a
list of lesson titles or slide counts.

Lessons covered:
{summaries_text}

Return ONLY valid JSON:
{{
  "summary_statement": "One bold sentence capturing the chapter's overarching takeaway.",
  "headers": ["Concept", "Definition", "Key Point"],
  "rows": [["ConceptName", "Short definition", "What students should remember about it"]]
}}

Include 3-6 rows covering the most important concepts across the whole chapter — grouped thematically, not one row per lesson."""

    params = dict(
        model=model,
        messages=[{"role": "user", "content": prompt}]
    )
    if _supports_json_response(model):
        params["response_format"] = {"type": "json_object"}

    if _uses_completion_tokens(model):
        params["max_completion_tokens"] = 1800
    else:
        params["max_tokens"] = 1800
        params["temperature"] = 0.3

    return _chat_json(client, params)
