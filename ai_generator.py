"""
JoVE AI Content Generator
Sends lesson content to OpenAI and receives structured slide JSON.
Transcript is the PRIMARY source; Pagetext is SUPPORTING only.
"""

import json
import re
from typing import Optional
from openai import OpenAI

SYSTEM_PROMPT = """You are an expert educational content designer for JoVE (Journal of Visualized Experiments), a scientific video platform. Your job is to convert lesson transcripts into structured slide content for lecture presentations.

SOURCE PRIORITY — THIS IS CRITICAL:
- The TRANSCRIPT is your PRIMARY and MAIN source. Build slide content, structure, and language from the transcript first.
- The PAGE TEXT is SUPPORTING ONLY — use it to add scientific terminology, precise definitions, or fill small gaps the transcript doesn't cover. Never let page text override or replace transcript content. If the transcript covers a topic, use the transcript's framing even if page text phrases it differently.

STRICT RULES:
1. Each slide covers ONE concept only. Never put two concepts on one slide.
2. All scientific names (genus and species) MUST be formatted as: **_Genus species_** (italic, genus capitalized, species lowercase).
3. Slide titles MUST exactly match the lesson name provided (for video reference) — EXCEPT for discussion slides, which get their own descriptive title.
4. Discussion questions MUST be split: question on one slide, answer on the next. The answer must never appear on the question slide.
5. Generate speaker notes for every slide using conversational, transcript-style language — as if the presenter is talking through the transcript's explanation.
6. If content has types, conditions, stages, or comparisons (2 or more), you MUST generate a 'table' slide for them — a concept slide is not sufficient.
7. Image search queries must be scientifically specific and visual — describe what should be SEEN (a diagram, a labeled illustration, a process, an organism), not abstract concepts. E.g. "DNA double helix structure diagram labeled" not "genetics importance".
8. Body text per slide: maximum 4 lines / 3 distinct points. Split across multiple slides if the transcript covers more.
9. Descriptive discussion titles: instead of generic "Discussion", use specific framing like "Discussion: Evolution of Mimicry" — derived from the actual question topic.
10. glossary_terms must include every keyword, scientific term, and defined concept mentioned in the transcript for this lesson — comprehensive but only terms actually discussed.

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
      "sub_label": "optional section label like 'The Mechanism'",
      "body": "Content text from the TRANSCRIPT. Use **bold** for key terms. Max 4 lines.",
      "image_query": "specific visual description for image search - diagram/illustration/organism/process",
      "speaker_notes": "conversational notes paraphrasing the transcript's explanation"
    },
    {
      "type": "table",
      "title": "exact lesson name",
      "sub_title": "descriptive subtitle for what this table shows",
      "headers": ["Column 1", "Column 2"],
      "rows": [["row1col1", "row1col2"]],
      "image_query": "specific visual description",
      "speaker_notes": "..."
    },
    {
      "type": "discussion_question",
      "title": "Discussion: <specific topic from this lesson>",
      "question": "The question text",
      "hint": "optional hint",
      "image_query": "specific visual description",
      "speaker_notes": "notes about facilitating discussion"
    },
    {
      "type": "discussion_answer",
      "title": "Discussion: <same specific topic, matches question slide>",
      "answer_summary": "short answer headline",
      "answer_explanation": "full explanation paragraph from transcript reasoning",
      "image_query": "specific visual description",
      "speaker_notes": "explanation for presenter"
    },
    {
      "type": "summary",
      "summary_statement": "One sentence capturing this lesson's core takeaway, written FOR STUDENTS as a key learning point.",
      "headers": ["Concept", "Key Point"],
      "rows": [["concept1", "what students should remember"]],
      "speaker_notes": "wrap-up notes"
    }
  ],
  "glossary_terms": {
    "Term": "Definition derived from the transcript"
  }
}"""


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
- Max 4 lines of body text per concept slide.
- Always end with discussion_question + discussion_answer (descriptive titles, not generic "Discussion").
- If the transcript describes 2+ types/stages/comparisons, generate a table slide for them.
- Only generate a "summary" slide if genuinely useful for this lesson's recap - it's optional.
- glossary_terms: comprehensive for THIS lesson's transcript content."""

    new_models = ["gpt-5", "o1", "o3", "o4"]
    use_new = any(model.startswith(m) for m in new_models)

    params = dict(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ]
    )

    if not any(model.startswith(m) for m in ["o1", "o3", "o4"]):
        params["response_format"] = {"type": "json_object"}

    if use_new:
        params["max_completion_tokens"] = 6000
    else:
        params["max_tokens"] = 6000
        params["temperature"] = 0.4

    response = client.chat.completions.create(**params)
    raw = response.choices[0].message.content
    return json.loads(raw)


def generate_chapter_summary(chapter_name: str, lesson_summaries: list,
                              api_key: str, model: str = "gpt-4.1") -> dict:
    """
    Generates a STUDENT-FACING chapter summary - a recap table of the
    key concepts learned, matching the benchmark style (Mechanism / Types /
    Frequency etc.), NOT a build report of slide counts.
    """
    client = OpenAI(api_key=api_key)

    summaries_text = "\n\n".join(
        f"Lesson: {s['name']}\nKey points: {s.get('key_points', 'N/A')}"
        for s in lesson_summaries
    )

    prompt = f"""Create a STUDENT-FACING chapter summary for "{chapter_name}".

This summary will be the LAST content slide students see. It should recap the
core concepts they learned across the chapter - like a study guide, NOT a
list of lesson titles or slide counts.

Lessons covered:
{summaries_text}

Return ONLY valid JSON:
{{
  "summary_statement": "One bold sentence capturing the chapter's overarching takeaway.",
  "headers": ["Concept", "Key Point"],
  "rows": [["ConceptName", "What students should remember about it"]]
}}

Include 3-6 rows covering the most important concepts across the whole chapter - grouped thematically, not one row per lesson."""

    new_models = ["gpt-5", "o1", "o3", "o4"]
    use_new = any(model.startswith(m) for m in new_models)

    params = dict(
        model=model,
        messages=[{"role": "user", "content": prompt}]
    )
    if not any(model.startswith(m) for m in ["o1", "o3", "o4"]):
        params["response_format"] = {"type": "json_object"}
    if use_new:
        params["max_completion_tokens"] = 1500
    else:
        params["max_tokens"] = 1500
        params["temperature"] = 0.4

    response = client.chat.completions.create(**params)
    return json.loads(response.choices[0].message.content)
