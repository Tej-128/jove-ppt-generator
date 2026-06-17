"""
JoVE AI Content Generator
Sends lesson content to OpenAI and receives structured slide JSON.

Key update:
- Definition-first sequencing is enforced.
- Tables must include Definition + Example columns when introducing key terms/types/stages.
- Each image-bearing slide includes transcript_anchor_text and visual_focus so the video-frame picker can find the matching moment in the lesson MP4.
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
3. Tables are NOT always 3 columns. Choose the table structure based on the concept:
   - Use Definition/Meaning and Example/Application columns when introducing terms, steps, types, stages, or methods.
   - Use fewer or different columns for simple comparisons, pros/cons, cause/effect, inputs/outputs, timelines, or other structures where 3 columns would be forced or awkward.
4. Table slide titles/subtitles must describe the table content, not just repeat a generic label.
5. Before explaining an example, add the general "how it works" or "method/process" explanation.
6. Keep wording student-facing and concise.
7. Keep table rows short. Each cell should ideally be under 14 words.
8. Avoid tables with more than 4 rows unless the content absolutely requires it. If there are many steps, group them logically or make the text more compact.

STRICT RULES:
0. Never add writer name, author name, reviewer name, prepared-by text, or individual credit lines anywhere.
1. Each slide covers ONE parent concept only. Related subtypes may appear together only when the parent concept is the slide concept. Never mix unrelated concepts on one slide.
2. All scientific names (genus and species) MUST be formatted as: **_Genus species_** (italic, genus capitalized, species lowercase).
3. Slide titles MUST exactly match the lesson name provided for concept/table/summary slides. Use sub_label or sub_title for the more specific reference-style heading.
4. Discussion questions MUST be split: question on one slide, answer on the next. The answer must never appear on the question slide.
5. Generate speaker notes for every slide using conversational, transcript-style language — as if the presenter is talking through the transcript's explanation.
6. If content has types, conditions, stages, or comparisons (2 or more), generate a table when useful. Add Definition/Meaning and Example/Application columns only when they fit the table purpose.
7. Body text per slide: maximum 4 lines / 3 distinct points. Split across multiple slides if the transcript covers more.
8. Descriptive discussion titles: instead of generic "Discussion", use specific framing like "Discussion: Evolution of Mimicry" — derived from the actual question topic.
9. glossary_terms must include every keyword, scientific term, and defined concept mentioned in the transcript for this lesson — comprehensive but only terms actually discussed.
10. For every slide that requires an image, provide:
    - image_required: true
    - visual_focus: a precise description of the best matching visual frame from the lesson video
    - transcript_anchor_text: a short exact or near-exact phrase from the transcript that tells where in the video this slide concept occurs

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
      "title": "exact lesson name",
      "sub_label": "Definition / Core Idea / How It Works / Example",
      "body": "Content text from the TRANSCRIPT. Use **bold** for key terms. Max 4 lines.",
      "image_required": true,
      "visual_focus": "specific visual frame to pick from the lesson video",
      "transcript_anchor_text": "short exact or near-exact transcript phrase for timing alignment",
      "speaker_notes": "conversational notes paraphrasing the transcript's explanation"
    },
    {
      "type": "table",
      "title": "exact lesson name",
      "sub_title": "descriptive subtitle for what this table shows",
      "table_kind": "definition_example | comparison | process | timeline | cause_effect | other",
      "headers": ["Use the most suitable columns for this table"],
      "rows": [["row values matching the headers"]],
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


def _normalize_table_slide(slide: dict) -> dict:
    headers = [str(h) for h in (slide.get("headers") or [])]
    rows = [list(row) for row in (slide.get("rows") or [])]

    if not headers:
        slide["headers"] = ["Concept", "Key Point"]
        slide["rows"] = rows or [["", ""]]
        slide.setdefault("table_kind", "other")
        return slide

    lower_headers = [h.lower() for h in headers]
    has_definition = any("definition" in h or "meaning" in h for h in lower_headers)
    has_example = any("example" in h or "application" in h for h in lower_headers)

    if _table_needs_definition_example(slide) and (not has_definition or not has_example):
        first_header = headers[0] if headers else "Term/Step/Type"
        slide["headers"] = [first_header, "Definition/Meaning", "Example/Application"]
        new_rows = []
        for row in rows:
            row = [str(x) for x in row]
            if len(row) == 0:
                row = ["", "", ""]
            elif len(row) == 1:
                row = [row[0], "", ""]
            elif len(row) == 2:
                row = [row[0], row[1], ""]
            else:
                # Preserve the first value, put the current explanation in definition,
                # and merge remaining values into example/application.
                row = [row[0], row[1], "; ".join(row[2:])]
            new_rows.append(row)
        slide["rows"] = new_rows
        slide["table_kind"] = slide.get("table_kind") or "definition_example"
    else:
        n_cols = len(headers)
        normalized_rows = []
        for row in rows:
            row = [str(x) for x in row]
            if len(row) < n_cols:
                row += [""] * (n_cols - len(row))
            normalized_rows.append(row[:n_cols])
        slide["headers"] = headers
        slide["rows"] = normalized_rows or [[""] * n_cols]
        slide.setdefault("table_kind", "other")

    return slide


def _normalize_slide(slide: dict, lesson_name: str, transcript: str, is_first_content_slide: bool = False) -> dict:
    stype = slide.get("type", "concept")

    if stype in {"concept", "table", "summary"}:
        slide["title"] = lesson_name

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

LESSON NAME (use EXACTLY as slide title for concept/table/summary slides): {lesson_name}

CONCEPT SLIDE BUDGET: {concept_slide_budget} slides (concept/table/summary combined).
This is IN ADDITION to the mandatory discussion_question + discussion_answer pair (2 slides),
which you must ALWAYS include at the end. Do not count Q&A toward this budget.

=== TRANSCRIPT (PRIMARY SOURCE - build content from this) ===
{transcript}

=== PAGE TEXT (SUPPORTING ONLY - use only to fill gaps or add precise terminology) ===
{pagetext if pagetext else "No page text available. Use transcript only."}

REMINDERS:
- First content slide MUST define the lesson's core concept before examples.
- If using a table for key terms/steps/types, include definition/meaning and example/application columns.
- Table subtitles must be specific and suitable to the table; do not use generic headings.
- Keep table cells compact so tables stay inside the slide margins.
- For table slides, assume row-level images will be embedded in the table, so rows must be short and image-friendly.
- Before explaining an example, include how the method/process works.
- Max 4 lines of body text per concept slide.
- Always end with discussion_question + discussion_answer.
- For video-frame selection, every image-bearing slide must include visual_focus and transcript_anchor_text.
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
