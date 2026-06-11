"""
JoVE AI Content Generator
Sends lesson content to OpenAI and receives structured slide JSON.
"""

import json
import re
from typing import Optional
from openai import OpenAI

SYSTEM_PROMPT = """You are an expert educational content designer for JoVE (Journal of Visualized Experiments), a scientific video platform. Your job is to convert lesson transcripts and page text into structured slide content for lecture presentations.

STRICT RULES:
1. The Transcript is PRIMARY — use it for tone, pacing, and main content.
2. The Pagetext is REFERENCE — use it for additional depth, definitions, and scientific accuracy.
3. Each slide covers ONE concept only. Never put two concepts on one slide.
4. All scientific names (genus species) MUST be formatted as: **_Genus species_** (italic, genus capitalized, species lowercase).
5. Slide titles MUST exactly match the lesson name provided (for video reference).
6. Discussion questions MUST be split: question on one slide, answer on the next.
7. Generate speaker notes for every slide using conversational transcript language.
8. MANDATORY: If the lesson mentions 2 or more types, stages, conditions, or comparisons you MUST generate a 'table' slide for them. This is not optional — a concept slide is not sufficient when types or conditions are present.
9. Image search queries must be scientifically specific, not decorative. E.g. "hemoglobin iron oxygen transport red blood cell diagram" not "biology".
10. Decide slide count based on content volume: simple lesson = 2-3 slides, complex = 4-5. Never exceed budget.
11. Summary slides must list ONLY the key points covered in that lesson — one concise bullet per concept, nothing else.
12. glossary_terms must include EVERY keyword, scientific term, and defined concept from the entire lesson — comprehensive, no omissions.
13. Body text per slide must be SHORT — maximum 4 lines, 3 points. If there is more content, split across multiple slides.

SLIDE TYPES you can create:
- "concept": Main content slide with body text
- "table": Comparison/definition table
- "discussion_question": Quiz question (NO answer on this slide)
- "discussion_answer": Answer to the previous quiz question
- "summary": Summary table at end of multi-concept lesson

OUTPUT: Return ONLY valid JSON, no markdown, no explanation. Use this exact schema:

{
  "lesson_name": "...",
  "slides": [
    {
      "type": "concept",
      "title": "exact lesson name",
      "sub_label": "optional section label like 'The Mechanism'",
      "body": "Content text. Use **bold** for key terms. Max 4 lines. Use newlines for separate points.",
      "image_query": "specific scientific image search query for Wikimedia Commons",
      "speaker_notes": "conversational notes for the presenter"
    },
    {
      "type": "table",
      "title": "exact lesson name",
      "sub_title": "descriptive subtitle e.g. 'Essential Elements in Living Systems'",
      "headers": ["Column 1", "Column 2"],
      "rows": [["row1col1", "row1col2"], ["row2col1", "row2col2"]],
      "image_query": "specific scientific image search query",
      "speaker_notes": "..."
    },
    {
      "type": "discussion_question",
      "title": "exact lesson name",
      "question": "The question text",
      "hint": "optional hint",
      "image_query": "specific scientific image query",
      "speaker_notes": "notes about facilitating discussion"
    },
    {
      "type": "discussion_answer",
      "title": "exact lesson name",
      "answer_summary": "short answer headline",
      "answer_explanation": "full explanation paragraph",
      "image_query": "specific scientific image query",
      "speaker_notes": "explanation for presenter"
    },
    {
      "type": "summary",
      "summary_statement": "One sentence that captures the lesson core idea.",
      "headers": ["Topic", "Key Point"],
      "rows": [["topic1", "point1"], ["topic2", "point2"]],
      "speaker_notes": "wrap-up notes"
    }
  ],
  "glossary_terms": {
    "Term": "Definition",
    "Another Term": "Its definition"
  }
}"""


def generate_slide_content(lesson_name: str, transcript: str, pagetext: str,
                            slide_budget: int, api_key: str,
                            model: str = "gpt-4.1") -> dict:
    client = OpenAI(api_key=api_key)

    user_prompt = f"""Generate slide content for this lesson.

LESSON NAME (use EXACTLY as slide title): {lesson_name}
CONCEPT SLIDE BUDGET: {slide_budget} concept/table/summary slides. The discussion Q+A pair (2 slides) will ALWAYS be added automatically on top of this budget. Do NOT count Q+A in your budget.

=== TRANSCRIPT (PRIMARY SOURCE) ===
{transcript}

=== PAGE TEXT (REFERENCE/DEPTH) ===
{pagetext if pagetext else "No page text available. Use transcript only."}

CRITICAL RULES FOR THIS GENERATION:
- Max 4 lines of body text per concept slide. Split content across multiple slides if needed.
- Always end with one discussion_question + one discussion_answer slide (these are IN ADDITION to your concept budget, not counted within it).
- glossary_terms must capture EVERY defined term in the lesson — be comprehensive.
- image_query must be a Wikimedia Commons compatible scientific search query."""

    # Detect model family for correct parameter usage
    new_models = ["gpt-5", "o1", "o3", "o4"]
    use_new = any(model.startswith(m) for m in new_models)

    params = dict(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ]
    )

    # response_format not supported by o-series reasoning models
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


def calculate_slide_budget(num_lessons: int, lesson_word_count: int,
                           avg_word_count: float) -> int:
    if num_lessons == 0:
        return 3
    chapter_budget = max(20, min(60, round(20 + ((num_lessons - 5) / 20) * 40)))
    lesson_budget = chapter_budget - 4
    base = max(2, round(lesson_budget / num_lessons))
    if avg_word_count > 0:
        ratio = lesson_word_count / avg_word_count
        if ratio > 1.3:
            base = min(5, base + 1)
        elif ratio < 0.6:
            base = max(2, base - 1)
    return min(5, max(2, base))


def search_wikimedia_image(query: str) -> Optional[str]:
    """
    Search Wikimedia Commons for a scientifically relevant image.
    Uses pageimages API for reliable thumbnail retrieval.
    Returns direct image URL or None.
    """
    import requests

    headers = {"User-Agent": "JoVE-PPT-Generator/1.0 (educational use)"}

    try:
        # Step 1: Search Wikipedia for the most relevant article
        search_url = "https://en.wikipedia.org/w/api.php"
        search_params = {
            "action": "query",
            "list": "search",
            "srsearch": query,
            "srlimit": "5",
            "format": "json"
        }
        r = requests.get(search_url, params=search_params, timeout=10, headers=headers)
        results = r.json().get("query", {}).get("search", [])

        if not results:
            return None

        # Step 2: Get main image from the top articles
        titles = "|".join([res["title"] for res in results[:3]])
        img_params = {
            "action": "query",
            "prop": "pageimages",
            "titles": titles,
            "pithumbsize": "800",
            "pilimit": "3",
            "format": "json"
        }
        r2 = requests.get(search_url, params=img_params, timeout=10, headers=headers)
        pages = r2.json().get("query", {}).get("pages", {})

        for page in pages.values():
            thumb = page.get("thumbnail", {})
            src = thumb.get("source", "")
            if src and not any(x in src.lower() for x in ["icon", "flag", "logo", "symbol"]):
                return src

    except Exception:
        pass

    # Step 3: Fallback to Wikimedia Commons generator search
    try:
        commons_url = "https://commons.wikimedia.org/w/api.php"
        commons_params = {
            "action": "query",
            "generator": "search",
            "gsrsearch": f"filetype:bitmap {query}",
            "gsrnamespace": "6",
            "gsrlimit": "10",
            "prop": "imageinfo",
            "iiprop": "url|mime|width|height",
            "format": "json"
        }
        r3 = requests.get(commons_url, params=commons_params, timeout=10, headers=headers)
        pages = r3.json().get("query", {}).get("pages", {})
        for page in pages.values():
            info = page.get("imageinfo", [])
            if not info:
                continue
            img_url = info[0].get("url", "")
            mime = info[0].get("mime", "")
            w = info[0].get("width", 0)
            h = info[0].get("height", 0)
            if (mime.startswith("image/") and
                    "svg" not in img_url.lower() and
                    w > 300 and h > 200 and
                    not any(x in img_url.lower() for x in ["icon", "logo", "flag"])):
                return img_url
    except Exception:
        pass

    return None
