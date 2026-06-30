"""
JoVE Planning Pass
Before any slide generation begins, this module reads a brief summary of
every lesson in the chapter and decides how many concept slides each
lesson deserves, based on actual content density and pedagogical weight —
not a mechanical division.

This version preserves the original planning rules and adds stronger
definition-first sequencing guidance for the downstream slide generator.
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
- Whether the first slide of a lesson needs a clear definition before examples begin

RULES:
1. The TOTAL chapter budget INCLUDES: 1 cover slide + 1 chapter overview slide + chapter summary slide + glossary pages + all lesson slides.
   - Reserve 1 slide for cover.
   - Reserve 1 slide for the chapter-level overview immediately after the cover.
   - Reserve 1 slide for chapter summary.
   - Reserve glossary pages based on expected term count: 1 page per ~10 terms, max 3 pages. Assume 2 pages unless content suggests otherwise.
2. Discussion Question + Answer pairs are NOT mandatory for every lesson. They are capped at chapter level and should be reserved for strong teaching moments only.
3. Concept slide count per lesson should typically be 1-4. Do not mechanically divide — use judgment. A short, simple lesson might need only 1 concept slide; a dense lesson with multiple comparisons might need 4.
4. The SUM of concept_slides across all lessons, PLUS reserves (cover + chapter overview + summary + glossary) and the chapter-level discussion reserve, should be as close as possible to the TOTAL SLIDE BUDGET, without exceeding it by more than 2.
5. If the total budget is very tight, prioritize foundational/complex lessons with more concept explanation slides and give simpler lessons the minimum of 1.
6. If the total budget is generous, give content-dense lessons up to 4, but don't pad simple lessons unnecessarily — it's fine to have leftover budget; just don't be wasteful or mechanical.
7. For foundational lessons such as "What is...", "Introduction to...", or "Scientific Method", ensure the allocation supports a definition-first opening before examples.
8. For lessons with key terms/types/stages, favor enough room for a table that includes definition + example instead of jumping directly to examples.

OUTPUT: Return ONLY valid JSON, no markdown, in this exact schema:
{
  "reasoning": "Brief 2-3 sentence explanation of your overall allocation strategy for this chapter.",
  "glossary_pages": 2,
  "allocations": [
    {"lesson_id": "10649", "lesson_name": "The Scientific Method", "concept_slides": 3, "reason": "Foundational lesson introducing core methodology; multiple sub-concepts (hypothesis, prediction, control)."},
    {"lesson_id": "10650", "lesson_name": "Inductive Reasoning", "concept_slides": 2, "reason": "Builds on prior lesson; single clear concept with one example."}
  ]
}"""


def _supports_json_response(model: str) -> bool:
    model = model or ""
    return not any(model.startswith(m) for m in ["o1", "o3", "o4"])


def _uses_completion_tokens(model: str) -> bool:
    model = model or ""
    return any(model.startswith(m) for m in ["gpt-5", "o1", "o3", "o4"])



def _rebalance_allocations_to_budget(lessons: list, allocations: dict,
                                     glossary_pages: int, total_slide_budget: int) -> tuple:
    """
    Enforce the requested total budget after the LLM planning pass.

    This fixes cases where the planner says target_total=20 but allocates
    too many concept slides, causing the actual PPT to exceed the requested count.
    """
    if not lessons or not total_slide_budget:
        return allocations, "No budget rebalance needed."

    discussion_pairs = min(6, max(1, round(len(lessons) / 3)))
    reserves = 1 + 1 + 1 + glossary_pages  # cover + chapter overview + chapter summary + glossary
    discussion_reserve = discussion_pairs * 2
    available_concepts = total_slide_budget - reserves - discussion_reserve

    # Every non-stub lesson should get at least 1 concept slide.
    min_concepts = len(lessons)
    if available_concepts < min_concepts:
        available_concepts = min_concepts

    # Normalize all lesson IDs.
    normalized = {}
    for lesson in lessons:
        lid = str(lesson["id"])
        normalized[lid] = max(1, min(4, int(allocations.get(lid, 2))))

    current = sum(normalized.values())

    # Reduce from lessons with the largest allocations first, preserving minimum 1.
    while current > available_concepts:
        reducible = [
            lesson for lesson in lessons
            if normalized[str(lesson["id"])] > 1
        ]
        if not reducible:
            break

        reducible.sort(
            key=lambda lesson: (
                normalized[str(lesson["id"])],
                len((lesson.get("transcript", "") + " " + lesson.get("pagetext", "")).split())
            ),
            reverse=True
        )

        chosen = reducible[0]
        normalized[str(chosen["id"])] -= 1
        current -= 1

    # If budget allows more and the planner under-allocated, add to dense lessons up to 4.
    while current < available_concepts:
        expandable = [
            lesson for lesson in lessons
            if normalized[str(lesson["id"])] < 4
        ]
        if not expandable:
            break

        expandable.sort(
            key=lambda lesson: len((lesson.get("transcript", "") + " " + lesson.get("pagetext", "")).split()),
            reverse=True
        )

        chosen = expandable[0]
        normalized[str(chosen["id"])] += 1
        current += 1

    note = (
        f"Budget rebalance applied: reserves={reserves}, discussion_pairs={discussion_pairs}, "
        f"discussion_reserve={discussion_reserve}, available_concepts={available_concepts}, final_concepts={sum(normalized.values())}."
    )
    return normalized, note


def plan_chapter_slides(lessons: list, total_slide_budget: int,
                         api_key: str, model: str = "gpt-4.1") -> dict:
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

    lesson_previews = []
    for l in lessons:
        word_count = len((l.get('transcript', '') + " " + l.get('pagetext', '')).split())
        preview_source = l.get('transcript', '') or l.get('pagetext', '')
        preview = preview_source[:500].strip()
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
- Reserve 1 for cover, 1 for the chapter overview slide, 1 for chapter summary, and glossary_pages (you decide, 1-3) for glossary.
- Discussion Q&A is capped at chapter level, not per lesson; prioritize concept explanation slides over frequent discussion slides.
- Sum should land close to {total_slide_budget} total.
- Definition-first lessons need enough room to define the concept before examples.
- Tables should include definition and example columns when key terms/types/stages are introduced."""

    params = dict(
        model=model,
        messages=[
            {"role": "system", "content": PLANNING_PROMPT},
            {"role": "user", "content": user_prompt}
        ]
    )

    if _supports_json_response(model):
        params["response_format"] = {"type": "json_object"}

    if _uses_completion_tokens(model):
        params["max_completion_tokens"] = 3000
    else:
        params["max_tokens"] = 3000
        params["temperature"] = 0.3

    response = client.chat.completions.create(**params)
    raw = response.choices[0].message.content
    plan = json.loads(raw)

    allocation_map = {}
    for item in plan.get("allocations", []):
        lid = str(item.get("lesson_id", ""))
        concept_slides = item.get("concept_slides", 2)
        concept_slides = max(1, min(4, int(concept_slides)))
        allocation_map[lid] = concept_slides

    for l in lessons:
        if l["id"] not in allocation_map:
            allocation_map[l["id"]] = 2

    glossary_pages = max(1, min(3, int(plan.get("glossary_pages", 2))))

    allocation_map, rebalance_note = _rebalance_allocations_to_budget(
        lessons=lessons,
        allocations=allocation_map,
        glossary_pages=glossary_pages,
        total_slide_budget=total_slide_budget
    )

    reasoning = plan.get("reasoning", "")
    if rebalance_note:
        reasoning = (reasoning + " " + rebalance_note).strip()

    return {
        "reasoning": reasoning,
        "glossary_pages": glossary_pages,
        "allocations": allocation_map,
        "raw_plan": plan,
        "rebalance_note": rebalance_note
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
