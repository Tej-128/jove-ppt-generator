"""
JoVE AI Content Generator
Generates structured lesson slides from transcript-first lesson content.
"""

import json
from typing import Dict, Any
from openai import OpenAI

SYSTEM_PROMPT = """You are an expert educational content designer for JoVE (Journal of Visualized Experiments), a scientific video platform.

Your job is to convert a lesson transcript into structured lecture slide content.

SOURCE PRIORITY:
- The TRANSCRIPT is the PRIMARY source.
- The PAGE TEXT is SUPPORTING only and may be used to clarify scientific wording.

STRICT RULES:
1. Each slide covers ONE concept only.
2. Every slide title for concept/table/summary slides must use the EXACT lesson name.
3. Discussion slides must have descriptive titles such as "Discussion: ...".
4. Discussion question and answer MUST be the final two slides.
5. Speaker notes must sound natural and conversational.
6. If the lesson contains 2+ types/stages/comparisons, at least one table slide should be used when appropriate.
7. Body text on concept slides must be short: max 4 lines / 3 compact points.
8. glossary_terms must include all important terms actually used in the lesson.
9. For every slide that needs an image, add:
   - image_required: true
   - visual_focus: what should ideally be visible in the frame
   - transcript_anchor_text: a short exact or near-exact snippet from the transcript representing the concept on that slide. This is critical for aligning slide content to the lesson video timeline.
10. Summary and glossary-style recap slides do not need an image.

OUTPUT JSON SCHEMA:
{
  "lesson_name": "...",
  "slides": [
    {
      "type": "concept",
      "title": "exact lesson name",
      "sub_label": "optional",
      "body": "short body text",
      "image_required": true,
      "visual_focus": "what should be visible in the matching video frame",
      "transcript_anchor_text": "exact or near-exact short snippet from transcript",
      "speaker_notes": "conversational presenter notes"
    },
    {
      "type": "table",
      "title": "exact lesson name",
      "sub_title": "descriptive subtitle",
      "headers": ["Column 1", "Column 2"],
      "rows": [["A", "B"]],
      "image_required": true,
      "visual_focus": "what should be visible in the matching video frame",
      "transcript_anchor_text": "exact or near-exact short snippet from transcript",
      "speaker_notes": "notes"
    },
    {
      "type": "discussion_question",
      "title": "Discussion: ...",
      "question": "question text",
      "hint": "optional hint",
      "image_required": true,
      "visual_focus": "what should be visible in the matching video frame",
      "transcript_anchor_text": "exact or near-exact short snippet from transcript",
      "speaker_notes": "notes"
    },
    {
      "type": "discussion_answer",
      "title": "Discussion: ...",
      "answer_summary": "short answer summary",
      "answer_explanation": "explanation",
      "image_required": true,
      "visual_focus": "what should be visible in the matching video frame",
      "transcript_anchor_text": "exact or near-exact short snippet from transcript",
      "speaker_notes": "notes"
    },
    {
      "type": "summary",
      "title": "exact lesson name",
      "summary_statement": "key takeaway",
      "headers": ["Concept", "Key Point"],
      "rows": [["A", "B"]],
      "image_required": false,
      "speaker_notes": "notes"
    }
  ],
  "glossary_terms": {
    "Term": "Definition"
  }
}
"""


def _json_request(client: OpenAI, params: Dict[str, Any]) -> Dict[str, Any]:
    response = client.chat.completions.create(**params)
    content = response.choices[0].message.content
    return json.loads(content)


def _normalize_slide(slide: Dict[str, Any], lesson_name: str) -> Dict[str, Any]:
    stype = slide.get("type", "concept")

    if stype in {"concept", "table", "discussion_question", "discussion_answer"}:
        slide["image_required"] = True
        slide["visual_focus"] = (slide.get("visual_focus") or slide.get("body") or slide.get("question") or slide.get("answer_summary") or lesson_name).strip()
        anchor = (slide.get("transcript_anchor_text") or "").strip()
        if not anchor:
            anchor = slide["visual_focus"][:180]
        slide["transcript_anchor_text"] = anchor
    else:
        slide["image_required"] = False

    if stype in {"concept", "table", "summary"}:
        slide["title"] = lesson_name

    return slide


def _validate_slide_payload(payload: Dict[str, Any], lesson_name: str) -> Dict[str, Any]:
    slides = payload.get("slides", []) or []
    normalized = [_normalize_slide(s, lesson_name) for s in slides]

    # Force discussion Q+A to be last two slides if present.
    discussion_q = [s for s in normalized if s.get("type") == "discussion_question"]
    discussion_a = [s for s in normalized if s.get("type") == "discussion_answer"]
    other = [s for s in normalized if s.get("type") not in {"discussion_question", "discussion_answer"}]

    if not discussion_q:
        discussion_q = [{
            "type": "discussion_question",
            "title": f"Discussion: {lesson_name}",
            "question": f"What is the most important concept students should remember from {lesson_name}?",
            "hint": "Use the main idea from the lesson.",
            "image_required": True,
            "visual_focus": lesson_name,
            "transcript_anchor_text": lesson_name,
            "speaker_notes": "Invite students to summarize the core lesson idea in their own words."
        }]
    if not discussion_a:
        discussion_a = [{
            "type": "discussion_answer",
            "title": discussion_q[0]["title"],
            "answer_summary": f"Key takeaway from {lesson_name}",
            "answer_explanation": f"The answer should reinforce the main scientific idea presented in {lesson_name}.",
            "image_required": True,
            "visual_focus": lesson_name,
            "transcript_anchor_text": lesson_name,
            "speaker_notes": "Connect the answer back to the transcript's key explanation."
        }]

    payload["lesson_name"] = lesson_name
    payload["slides"] = other + [discussion_q[0], discussion_a[0]]
    payload["glossary_terms"] = payload.get("glossary_terms", {}) or {}
    return payload


def generate_slide_content(lesson_name: str, transcript: str, pagetext: str,
                           concept_slide_budget: int, api_key: str,
                           model: str = "gpt-4.1") -> dict:
    """
    Generates structured slide content for one lesson.
    concept_slide_budget counts concept/table/summary slides only.
    A discussion question + answer pair is always appended.
    """
    client = OpenAI(api_key=api_key)

    user_prompt = f"""Generate slide content for this lesson.

LESSON NAME: {lesson_name}
CONCEPT/TABLE/SUMMARY SLIDE BUDGET: {concept_slide_budget}

The mandatory discussion_question and discussion_answer pair must be included at the end and do NOT count toward the budget.

=== TRANSCRIPT (PRIMARY SOURCE) ===
{transcript}

=== PAGE TEXT (SUPPORTING ONLY) ===
{pagetext if pagetext else 'No page text available.'}

Important alignment rule:
For each image-bearing slide, transcript_anchor_text must help align the slide with the correct moment in the lesson video. Use a short exact or near-exact snippet from the transcript.
"""

    params = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0.3,
        "max_tokens": 7000,
    }

    payload = _json_request(client, params)
    return _validate_slide_payload(payload, lesson_name)


def generate_chapter_summary(chapter_name: str, lesson_summaries: list,
                             api_key: str, model: str = "gpt-4.1") -> dict:
    client = OpenAI(api_key=api_key)

    lessons_text = "\n\n".join(
        f"Lesson: {item['name']}\nKey points: {item.get('key_points', 'N/A')}"
        for item in lesson_summaries
    )

    prompt = f"""Create a concise STUDENT-FACING summary slide for the chapter "{chapter_name}".

Lessons covered:
{lessons_text}

Return only valid JSON with this schema:
{{
  "summary_statement": "one strong recap sentence",
  "headers": ["Concept", "Key Point"],
  "rows": [["Concept", "Key takeaway"]]
}}

Use 3-6 rows. Group thematically where possible.
"""

    params = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "response_format": {"type": "json_object"},
        "temperature": 0.3,
        "max_tokens": 1800,
    }
    return _json_request(client, params)
