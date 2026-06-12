"""
JoVE Planning Pass
Before any slide generation begins, this module reads a brief summary of
every lesson in the chapter and decides how many concept slides each
lesson deserves, based on actual content density and pedagogical weight —
not a mechanical division.
"""

import json
from openai import OpenAI


PLANNING_PROMPT = """You are a curriculum planning expert for JoVE educational content.

You will be given a list of lessons (each with a word count and a short content preview)
and a TOTAL SLIDE BUDGET for the entire chapter.

Your job: allocate a CONCEPT SLIDE COUNT to each lesson, based on:
- Content density (longer/more complex lessons need more slides)
- Pedagogical importance (foundational concepts that later lessons depend on deserve more room)
- Whether the lesson has multiple sub-topics, types, or comparisons that need their own slides

RULES:
1. The TOTAL chapter budget INCLUDES: 1 cover slide + chapter summary slide + glossary pages + all lesson slides.
   - Reserve 1 slide for cover.
   - Reserve 1 slide for chapter summary.
   - Reserve glossary pages based on expected term count: 1 page per ~10 terms, max 3 pages. Assume 2 pages unless content suggests otherwise.
2. Each lesson ALWAYS gets +2 slides on top of your concept allocation for the mandatory discussion Question + Answer pair. Do NOT include Q&A in your concept count.
3. Concept slide count per lesson should typically be 1-4. Do not mechanically divide — use judgment. A short, simple lesson might need only 1 concept slide; a dense lesson with multiple comparisons might need 4.
4. The SUM of (concept_slides + 2) across all lessons, PLUS reserves (cover + summary + glossary), should be as close as possible to the TOTAL SLIDE BUDGET, without exceeding it by more than 2.
5. If the total budget is very tight (e.g. fewer slides than 3 per lesson average), prioritize foundational/complex lessons with more slides and give simpler lessons the minimum of 1.
6. If the total budget is generous, give content-dense lessons up to 4, but don't pad simple lessons unnecessarily — it's fine to have leftover budget; just don't be wasteful or mechanical.

OUTPUT: Return ONLY valid JSON, no markdown, in this exact schema:
{
  "reasoning": "Brief 2-3 sentence explanation of your overall allocation strategy for this chapter.",
  "glossary_pages": 2,
  "allocations": [
    {"lesson_id": "10649", "lesson_name": "The Scientific Method", "concept_slides": 3, "reason": "Foundational lesson introducing core methodology; multiple sub-concepts (hypothesis, prediction, control)."},
    {"lesson_id": "10650", "lesson_name": "Inductive Reasoning", "concept_slides": 2, "reason": "Builds on prior lesson; single clear concept with one example."}
  ]
}"""


def plan_chapter_slides(lessons: list, total_slide_budget: int,
                         api_key: str, model: str = "gpt-5.5") -> dict:
    """
    lessons: list of dicts with keys 'id', 'name', 'transcript', 'pagetext'
    total_slide_budget: the team's target total slide count for the chapter
    Returns: {
        "reasoning": str,
        "glossary_pages": int,
        "allocations": {lesson_id: concept_slides}
    }
    """
    client = OpenAI(api_key=api_key)

    # Build a compact preview of each lesson for the planner
    lesson_previews = []
    for l in lessons:
        word_count = len((l.get('transcript', '') + l.get('pagetext', '')).split())
        # Use transcript as primary preview source
        preview_source = l.get('transcript', '') or l.get('pagetext', '')
        preview = preview_source[:400].strip()
        lesson_previews.append({
            "lesson_id": l["id"],
            "lesson_name": l["name"],
            "word_count": word_count,
            "preview": preview
        })

    user_prompt = f"""TOTAL SLIDE BUDGET FOR THIS CHAPTER: {total_slide_budget}

NUMBER OF LESSONS: {len(lessons)}

LESSONS (in chapter order):
{json.dumps(lesson_previews, indent=2)}

Allocate concept slide counts per lesson following the rules. Remember:
- Reserve 1 for cover, 1 for chapter summary, and glossary_pages (you decide, 1-3) for glossary.
- Each lesson's actual slide count will be concept_slides + 2 (mandatory Q&A).
- Sum should land close to {total_slide_budget} total."""

    new_models = ["gpt-5", "o1", "o3", "o4"]
    use_new = any(model.startswith(m) for m in new_models)

    params = dict(
        model=model,
        messages=[
            {"role": "system", "content": PLANNING_PROMPT},
            {"role": "user", "content": user_prompt}
        ]
    )
    if not any(model.startswith(m) for m in ["o1", "o3", "o4"]):
        params["response_format"] = {"type": "json_object"}

    if use_new:
        params["max_completion_tokens"] = 3000
    else:
        params["max_tokens"] = 3000
        params["temperature"] = 0.3

    response = client.chat.completions.create(**params)
    raw = response.choices[0].message.content
    plan = json.loads(raw)

    # Normalize into a simple lookup dict, with fallback safety
    allocation_map = {}
    for item in plan.get("allocations", []):
        lid = str(item.get("lesson_id", ""))
        concept_slides = item.get("concept_slides", 2)
        # Safety clamp
        concept_slides = max(1, min(4, int(concept_slides)))
        allocation_map[lid] = concept_slides

    # Ensure every lesson has an allocation (fallback to 2 if planner missed one)
    for l in lessons:
        if l["id"] not in allocation_map:
            allocation_map[l["id"]] = 2

    glossary_pages = max(1, min(3, int(plan.get("glossary_pages", 2))))

    return {
        "reasoning": plan.get("reasoning", ""),
        "glossary_pages": glossary_pages,
        "allocations": allocation_map,
        "raw_plan": plan
    }


def default_chapter_budget(num_lessons: int) -> int:
    """
    Fallback formula when the team doesn't specify a total slide count.
    Anchored to spec: 5 lessons -> 20 slides, 25 lessons -> 60 slides.
    """
    if num_lessons <= 0:
        return 12
    slides = 20 + (num_lessons - 5) * 2
    return max(12, min(60, round(slides)))
