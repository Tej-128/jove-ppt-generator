"""
JoVE AI Content Generator
Sends lesson content to OpenAI GPT-4.1 and receives structured slide JSON.
"""

import json
import re
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
8. If content has types, conditions, or comparisons — use a table or structured format.
9. Image search queries must be scientifically specific, not decorative. E.g. "hemoglobin iron oxygen transport red blood cell diagram" not "biology".
10. Decide slide count based on content volume: simple lesson = 2-3 slides, complex = 4-5. Never exceed budget.

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
      "body": "Content text. Use **bold** for key terms. Use newlines for separate points.",
      "image_query": "specific scientific image search query",
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
      "summary_statement": "One sentence that captures the lesson's core idea.",
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
    """
    Call OpenAI to generate structured slide content for one lesson.
    Returns parsed JSON dict.
    """
    client = OpenAI(api_key=api_key)

    user_prompt = f"""Generate slide content for this lesson.

LESSON NAME (use EXACTLY as slide title): {lesson_name}
SLIDE BUDGET: {slide_budget} slides maximum (including discussion Q&A pair which counts as 2)

=== TRANSCRIPT (PRIMARY SOURCE) ===
{transcript}

=== PAGE TEXT (REFERENCE/DEPTH) ===
{pagetext if pagetext else "No page text available. Use transcript only."}

Generate {slide_budget} slides maximum. Always end with one discussion question + one discussion answer slide.
Extract key terms for the glossary.
Remember: image_query must be scientifically specific for educational use."""

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.4,
        max_tokens=4000,
        response_format={"type": "json_object"}
    )

    raw = response.choices[0].message.content
    return json.loads(raw)


def calculate_slide_budget(num_lessons: int, lesson_word_count: int,
                           avg_word_count: float) -> int:
    """
    Adaptively assign slide budget per lesson.
    Base: proportional from chapter budget.
    Adjust: +1 for heavy lessons (>120% of avg), -1 for light (<70% of avg).
    Min 2, Max 5 per lesson (excluding cover, final summary, glossary).
    """
    if num_lessons == 0:
        return 3

    # Chapter budget formula: 20 + ((lessons-5)/20)*40, clamped 20-60
    chapter_budget = max(20, min(60, round(20 + ((num_lessons - 5) / 20) * 40)))
    # Reserve 4 slides for cover + final summary + 2 glossary pages
    lesson_budget = chapter_budget - 4
    base = max(2, round(lesson_budget / num_lessons))

    # Adjust for content volume
    if avg_word_count > 0:
        ratio = lesson_word_count / avg_word_count
        if ratio > 1.3:
            base = min(5, base + 1)
        elif ratio < 0.6:
            base = max(2, base - 1)

    return min(5, max(2, base))


def search_wikimedia_image(query: str) -> str | None:
    """
    Search Wikimedia Commons for a scientifically relevant image.
    Returns direct image URL or None.
    """
    import requests
    try:
        # Search for images
        search_url = "https://en.wikipedia.org/w/api.php"
        params = {
            "action": "query",
            "list": "search",
            "srsearch": f"{query} filetype:bitmap",
            "srnamespace": "6",  # File namespace
            "srlimit": "5",
            "format": "json"
        }
        r = requests.get(search_url, params=params, timeout=8)
        data = r.json()

        results = data.get("query", {}).get("search", [])
        if not results:
            # Fallback: try commons API
            commons_url = "https://commons.wikimedia.org/w/api.php"
            params_c = {
                "action": "query",
                "generator": "search",
                "gsrsearch": f"filetype:bitmap {query}",
                "gsrnamespace": "6",
                "gsrlimit": "3",
                "prop": "imageinfo",
                "iiprop": "url|mime",
                "format": "json"
            }
            r2 = requests.get(commons_url, params=params_c, timeout=8)
            data2 = r2.json()
            pages = data2.get("query", {}).get("pages", {})
            for page in pages.values():
                info = page.get("imageinfo", [])
                if info and info[0].get("mime", "").startswith("image/"):
                    url = info[0]["url"]
                    if not any(x in url.lower() for x in ["svg", "icon", "logo", "flag"]):
                        return url
            return None

        # Get image URL from first result
        file_title = results[0]["title"]
        img_params = {
            "action": "query",
            "titles": file_title,
            "prop": "imageinfo",
            "iiprop": "url|mime",
            "format": "json"
        }
        r3 = requests.get(search_url, params=img_params, timeout=8)
        data3 = r3.json()
        pages = data3.get("query", {}).get("pages", {})
        for page in pages.values():
            info = page.get("imageinfo", [])
            if info:
                mime = info[0].get("mime", "")
                url = info[0].get("url", "")
                if mime.startswith("image/") and "svg" not in url.lower():
                    return url
    except Exception:
        pass
    return None
