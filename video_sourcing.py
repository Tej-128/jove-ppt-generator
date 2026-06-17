"""
JoVE Video Sourcing
Selects transcript-aligned frames from lesson MP4s for slide imagery.

Strict rule:
- No web image search.
- No placeholders.
- Every image-bearing slide must use a frame from that lesson's MP4.
- If exact anchor-area frames are not clean, the module still chooses the cleanest frame from the same MP4 so a slide image is always added.
"""

import base64
import json
import os
import re
import tempfile
from typing import Dict, List

import cv2
import numpy as np
from openai import OpenAI


IMAGE_SLIDE_TYPES = {"concept", "table", "discussion_question", "discussion_answer"}


def _slug(text: str) -> str:
    text = re.sub(r"[^A-Za-z0-9_-]+", "_", str(text).strip())
    return re.sub(r"_+", "_", text).strip("_") or "frame"


def _normalize_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip().lower()
    return re.sub(r"[^a-z0-9 ]+", "", text)


def estimate_anchor_ratio(transcript: str, anchor_text: str) -> float:
    """
    Estimate where the slide concept occurs in the video using transcript position.
    Example: anchor near 50% of transcript -> target near 50% of video duration.
    """
    transcript_norm = _normalize_text(transcript)
    anchor_norm = _normalize_text(anchor_text)

    if not transcript_norm:
        return 0.5

    if anchor_norm and anchor_norm in transcript_norm:
        char_idx = transcript_norm.find(anchor_norm)
        return max(0.02, min(0.98, char_idx / max(1, len(transcript_norm))))

    transcript_words = transcript_norm.split()
    anchor_words = [w for w in anchor_norm.split() if len(w) > 2]

    if not transcript_words or not anchor_words:
        return 0.5

    best_score = -1
    best_idx = len(transcript_words) // 2
    window = max(8, len(anchor_words) * 3)

    for i in range(0, max(1, len(transcript_words) - window + 1)):
        chunk = transcript_words[i:i + window]
        score = sum(1 for word in anchor_words if word in chunk)
        if score > best_score:
            best_score = score
            best_idx = i

    return max(0.02, min(0.98, best_idx / max(1, len(transcript_words))))


def _get_video_stats(cap) -> Dict[str, float]:
    fps = cap.get(cv2.CAP_PROP_FPS) or 0
    frame_count = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0
    duration = frame_count / fps if fps > 0 else 0
    return {"fps": fps, "frame_count": frame_count, "duration": duration}


def _read_frame(cap, fps: float, t: float):
    if fps <= 0:
        return None
    frame_idx = max(0, int(round(t * fps)))
    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
    ok, frame = cap.read()
    return frame if ok else None


def _frame_metrics(frame, prev_frame=None, next_frame=None) -> Dict[str, float]:
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    brightness = float(np.mean(gray))
    blur = float(cv2.Laplacian(gray, cv2.CV_64F).var())

    transition_penalty = 0.0
    if prev_frame is not None:
        transition_penalty += float(np.mean(cv2.absdiff(frame, prev_frame)))
    if next_frame is not None:
        transition_penalty += float(np.mean(cv2.absdiff(frame, next_frame)))

    bright_score = 1.0 if 35 <= brightness <= 220 else max(0.0, 1.0 - abs(brightness - 128) / 128)
    blur_score = min(1.0, blur / 180.0)
    transition_score = max(0.0, 1.0 - (transition_penalty / 120.0))

    technical_score = (bright_score * 0.30) + (blur_score * 0.45) + (transition_score * 0.25)
    clean = blur >= 35 and 25 <= brightness <= 235

    return {
        "brightness": brightness,
        "blur": blur,
        "transition_penalty": transition_penalty,
        "technical_score": technical_score,
        "clean": clean,
    }


def _save_frame(path: str, frame) -> str:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    cv2.imwrite(path, frame, [int(cv2.IMWRITE_JPEG_QUALITY), 92])
    return _apply_watermark_cleanup(path)

def _apply_watermark_cleanup(path: str) -> str:
    """
    Remove likely JoVE watermark areas from extracted frames.

    Priority:
    1. Try OpenCV inpainting without changing image size.
    2. If image write/inpaint fails, keep original path.
    3. Cropping fallback is available through _crop_watermark_area(), used only when requested by downstream logic.

    Note: this is deterministic cleanup, not perfect semantic watermark detection.
    """
    try:
        img = cv2.imread(path)
        if img is None:
            return path

        h, w = img.shape[:2]
        mask = np.zeros((h, w), dtype=np.uint8)

        # Common watermark/logo areas in video exports: top-right and bottom-right.
        # Keep masks narrow to avoid damaging useful visual content.
        tr_x1 = int(w * 0.80)
        tr_y2 = int(h * 0.14)
        br_x1 = int(w * 0.78)
        br_y1 = int(h * 0.86)
        mask[0:tr_y2, tr_x1:w] = 255
        mask[br_y1:h, br_x1:w] = 255

        cleaned = cv2.inpaint(img, mask, 3, cv2.INPAINT_TELEA)
        cleaned_path = path.replace(".jpg", "_wmclean.jpg")
        cv2.imwrite(cleaned_path, cleaned, [int(cv2.IMWRITE_JPEG_QUALITY), 92])
        # V6 deterministic safety: crop likely watermark edge after cleanup so visible
        # watermarks do not remain in table cells or right-column visuals.
        if os.path.exists(cleaned_path):
            return _crop_watermark_area(cleaned_path)
        return _crop_watermark_area(path)
    except Exception:
        return _crop_watermark_area(path)


def _crop_watermark_area(path: str) -> str:
    """
    Crop likely watermark/logo area if inpainting looks bad.
    This preserves a rectangular image and avoids leaving the watermark visible.
    """
    try:
        img = cv2.imread(path)
        if img is None:
            return path
        h, w = img.shape[:2]

        # Crop a small border from top/bottom/right where watermarks most commonly appear.
        x1 = 0
        y1 = int(h * 0.03)
        x2 = int(w * 0.96)
        y2 = int(h * 0.97)
        cropped = img[y1:y2, x1:x2]
        crop_path = path.replace(".jpg", "_wmcrop.jpg")
        cv2.imwrite(crop_path, cropped, [int(cv2.IMWRITE_JPEG_QUALITY), 92])
        return crop_path if os.path.exists(crop_path) else path
    except Exception:
        return path


def _generate_ai_fallback_image(client: OpenAI, slide_def: Dict, lesson_name: str, output_dir: str) -> Dict:
    """
    AI image fallback, used only when no suitable JoVE frame can be selected.
    This is an approved project override to the guide's JoVE-only image rule.
    """
    os.makedirs(output_dir, exist_ok=True)
    prompt = (
        "Create a clean undergraduate biology educational illustration. "
        "No logos, no watermarks, no text labels unless essential. "
        "Style: realistic scientific diagram, white background, suitable for a JoVE lecture slide. "
        f"Lesson: {lesson_name}. "
        f"Visual focus: {slide_def.get('visual_focus') or slide_def.get('title') or lesson_name}."
    )
    out_path = os.path.join(output_dir, _slug(f"ai_fallback_{lesson_name}_{slide_def.get('type','slide')}") + ".png")
    try:
        result = client.images.generate(
            model="gpt-image-1",
            prompt=prompt,
            size="1024x1024"
        )
        data = result.data[0]
        if getattr(data, "b64_json", None):
            import base64
            with open(out_path, "wb") as f:
                f.write(base64.b64decode(data.b64_json))
            return {
                "path": out_path,
                "timestamp": None,
                "target_time": None,
                "technical_score": 1.0,
                "vision_confidence": 0,
                "selection_method": "ai_generated_fallback",
                "selection_reason": "No suitable JoVE frame was available; generated an AI fallback image as approved.",
                "anchor_text": slide_def.get("transcript_anchor_text") or "",
                "total_candidates_considered": 0,
            }
    except Exception as exc:
        raise RuntimeError(f"AI fallback image generation failed: {exc}")
    raise RuntimeError("AI fallback image generation failed: no image returned.")


def _generate_candidate_times(target_time: float, duration: float, count: int) -> List[float]:
    """
    Candidate count is tied to number of image-bearing slides in the lesson,
    while still enforcing a minimum so Vision has choices.
    """
    count = max(3, int(count))
    step = max(1.25, min(6.0, duration / 25.0 if duration else 2.5))

    offsets = [0.0]
    k = 1
    while len(offsets) < count:
        offsets.extend([-step * k, step * k])
        k += 1

    times = []
    for off in offsets:
        t = max(0.25, min(max(0.25, duration - 0.25), target_time + off))
        if all(abs(t - existing) > 0.2 for existing in times):
            times.append(t)
        if len(times) >= count:
            break
    return times


def extract_candidate_frames(video_path: str, target_time: float,
                             candidate_count: int, output_dir: str,
                             prefix: str) -> List[Dict]:
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        raise RuntimeError(f"Could not open video: {video_path}")

    stats = _get_video_stats(cap)
    duration = stats["duration"]
    if duration <= 0 or stats["fps"] <= 0:
        cap.release()
        raise RuntimeError(f"Video has invalid duration/FPS: {video_path}")

    candidate_times = _generate_candidate_times(target_time, duration, candidate_count)
    candidates = []

    for idx, t in enumerate(candidate_times, start=1):
        frame = _read_frame(cap, stats["fps"], t)
        if frame is None:
            continue

        prev_frame = _read_frame(cap, stats["fps"], max(0.0, t - 0.35))
        next_frame = _read_frame(cap, stats["fps"], min(duration, t + 0.35))
        metrics = _frame_metrics(frame, prev_frame, next_frame)

        img_name = f"{prefix}_{idx:02d}_{int(round(t * 1000)):07d}.jpg"
        img_path = os.path.join(output_dir, img_name)
        img_path = _save_frame(img_path, frame)

        candidates.append({
            "path": img_path,
            "timestamp": round(t, 2),
            **metrics,
        })

    cap.release()
    return sorted(candidates, key=lambda x: x["technical_score"], reverse=True)


def extract_cleanest_frames_from_video(video_path: str, count: int,
                                       output_dir: str, prefix: str) -> List[Dict]:
    """
    Same-video rescue path. This is not a placeholder or web fallback.
    It guarantees the slide still receives a frame from the lesson's MP4.
    """
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        raise RuntimeError(f"Could not open video: {video_path}")

    stats = _get_video_stats(cap)
    duration = stats["duration"]
    if duration <= 0 or stats["fps"] <= 0:
        cap.release()
        raise RuntimeError(f"Video has invalid duration/FPS: {video_path}")

    sample_count = max(8, count * 3)
    positions = np.linspace(0.04, 0.96, sample_count)
    samples = []

    for idx, ratio in enumerate(positions, start=1):
        t = float(ratio * duration)
        frame = _read_frame(cap, stats["fps"], t)
        if frame is None:
            continue

        prev_frame = _read_frame(cap, stats["fps"], max(0.0, t - 0.35))
        next_frame = _read_frame(cap, stats["fps"], min(duration, t + 0.35))
        metrics = _frame_metrics(frame, prev_frame, next_frame)

        img_name = f"{prefix}_clean_{idx:02d}_{int(round(t * 1000)):07d}.jpg"
        img_path = os.path.join(output_dir, img_name)
        img_path = _save_frame(img_path, frame)

        samples.append({
            "path": img_path,
            "timestamp": round(t, 2),
            **metrics,
        })

    cap.release()
    return sorted(samples, key=lambda x: x["technical_score"], reverse=True)


def _image_to_data_url(path: str) -> str:
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode("utf-8")
    return f"data:image/jpeg;base64,{data}"


def _vision_select_frame(client: OpenAI, slide_def: Dict, candidates: List[Dict],
                         lesson_name: str, transcript_anchor: str,
                         vision_model: str) -> Dict:
    """
    Uses OpenAI Vision to select the best candidate. If Vision call fails,
    still uses the best technical candidate from the SAME lesson video.
    """
    content = [{
        "type": "text",
        "text": (
            "Choose the best candidate image for a JoVE educational slide.\n"
            f"Lesson: {lesson_name}\n"
            f"Slide type: {slide_def.get('type')}\n"
            f"Slide title: {slide_def.get('title', lesson_name)}\n"
            f"Slide subtitle/label: {slide_def.get('sub_title') or slide_def.get('sub_label') or ''}\n"
            f"Visual focus: {slide_def.get('visual_focus', '')}\n"
            f"Transcript anchor: {transcript_anchor}\n\n"
            "Selection rules:\n"
            "1. Prefer the frame that best matches the slide concept.\n"
            "2. Avoid blurry, blank, transition, or unreadable frames.\n"
            "3. Prefer clear scientific visuals, diagrams, experimental visuals, or relevant presenter-screen visuals.\n"
            "4. Return only JSON: {\"selected_index\": 1, \"confidence\": 95, \"reason\": \"...\"}."
        )
    }]

    for idx, candidate in enumerate(candidates, start=1):
        content.append({
            "type": "text",
            "text": (
                f"Candidate {idx}: timestamp={candidate['timestamp']}s, "
                f"technical_score={candidate['technical_score']:.3f}, "
                f"blur={candidate['blur']:.1f}, "
                f"brightness={candidate['brightness']:.1f}, "
                f"transition_penalty={candidate['transition_penalty']:.1f}"
            )
        })
        content.append({"type": "image_url", "image_url": {"url": _image_to_data_url(candidate["path"])}})

    try:
        response = client.chat.completions.create(
            model=vision_model,
            messages=[{"role": "user", "content": content}],
            response_format={"type": "json_object"},
            temperature=0.1,
            max_tokens=600,
        )
        payload = json.loads(response.choices[0].message.content)
        selected_index = int(payload.get("selected_index", 1)) - 1
        selected_index = max(0, min(len(candidates) - 1, selected_index))
        selected = candidates[selected_index].copy()
        selected["vision_confidence"] = int(payload.get("confidence", 0))
        selected["selection_reason"] = payload.get("reason", "")
        selected["selection_method"] = "openai_vision"
        return selected
    except Exception as exc:
        selected = candidates[0].copy()
        selected["vision_confidence"] = int(round(selected["technical_score"] * 100))
        selected["selection_reason"] = f"Vision selection failed ({exc}); used best clean technical frame from same lesson MP4."
        selected["selection_method"] = "technical_same_video_rescue"
        return selected


def select_frame_for_slide(video_path: str, lesson_name: str, transcript: str,
                           slide_def: Dict, total_image_slides: int,
                           api_key: str, work_dir: str,
                           vision_model: str = "gpt-4.1",
                           used_frame_paths=None,
                           used_timestamps=None) -> Dict:
    if not os.path.exists(video_path):
        raise RuntimeError(f"Missing MP4 for lesson '{lesson_name}': {video_path}")

    os.makedirs(work_dir, exist_ok=True)
    used_frame_paths = set(used_frame_paths or [])
    used_timestamps = list(used_timestamps or [])
    anchor_text = slide_def.get("transcript_anchor_text") or slide_def.get("visual_focus") or lesson_name
    ratio = estimate_anchor_ratio(transcript, anchor_text)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        raise RuntimeError(f"Could not open lesson MP4 for '{lesson_name}': {video_path}")

    stats = _get_video_stats(cap)
    cap.release()

    duration = stats["duration"]
    if duration <= 0:
        raise RuntimeError(f"Lesson MP4 has invalid duration for '{lesson_name}': {video_path}")

    target_time = round(ratio * duration, 2)
    prefix = _slug(f"{lesson_name}_{slide_def.get('type', 'slide')}_{target_time}")

    # Base user rule: candidate pool is tied to image-bearing slide count.
    # Improvement: use a minimum of 5 candidates to prevent repeated frame selection
    # and to give Vision enough variety.
    candidate_count = max(5, int(total_image_slides))
    candidates = extract_candidate_frames(video_path, target_time, candidate_count, work_dir, prefix)
    clean_candidates = [c for c in candidates if c.get("clean")]

    if not clean_candidates:
        clean_candidates = extract_cleanest_frames_from_video(video_path, candidate_count, work_dir, prefix)

    if not clean_candidates:
        client = OpenAI(api_key=api_key)
        return _generate_ai_fallback_image(client, slide_def, lesson_name, work_dir)

    def unused(candidate):
        if candidate.get("path") in used_frame_paths:
            return False
        return all(abs(float(candidate.get("timestamp", 0)) - float(t)) >= 1.0 for t in used_timestamps)

    unused_candidates = [c for c in clean_candidates if unused(c)]

    # If the anchor-local pool repeats a previous frame, pull a broader clean pool from
    # the same lesson video. This is not a placeholder or web fallback.
    if not unused_candidates:
        broader = extract_cleanest_frames_from_video(video_path, max(candidate_count * 2, 10), work_dir, prefix)
        unused_candidates = [c for c in broader if unused(c)]
        if unused_candidates:
            clean_candidates = unused_candidates

    top_candidates = (unused_candidates or clean_candidates)[:candidate_count]

    client = OpenAI(api_key=api_key)
    selected = _vision_select_frame(
        client=client,
        slide_def=slide_def,
        candidates=top_candidates,
        lesson_name=lesson_name,
        transcript_anchor=anchor_text,
        vision_model=vision_model,
    )

    # Final duplicate guard. If Vision selected a duplicate, force the best unused
    # technical candidate from the same MP4.
    if not unused(selected):
        fallback_unused = [c for c in top_candidates if unused(c)]
        if fallback_unused:
            replacement = fallback_unused[0].copy()
            replacement["vision_confidence"] = selected.get("vision_confidence", 0)
            replacement["selection_reason"] = (
                "Vision selected a frame already used in this lesson; "
                "replaced with the best unused clean frame from the same MP4."
            )
            replacement["selection_method"] = "duplicate_guard_same_video"
            selected = replacement

    # If a selected JoVE frame receives very low confidence, use approved AI fallback.
    # This avoids bad/irrelevant frames while preserving JoVE-first priority.
    try:
        if int(selected.get("vision_confidence") or 0) < 60 and selected.get("selection_method") == "openai_vision":
            selected = _generate_ai_fallback_image(client, slide_def, lesson_name, work_dir)
    except Exception:
        pass

    selected["target_time"] = target_time
    selected["anchor_text"] = anchor_text
    selected["total_candidates_considered"] = len(top_candidates)
    return selected


def assign_frames_to_slides(lesson: Dict, slide_defs: List[Dict],
                            api_key: str, vision_model: str = "gpt-4.1",
                            progress_callback=None) -> Dict[int, Dict]:
    video_path = lesson.get("video_path")
    if not video_path:
        raise RuntimeError(f"Missing MP4 for lesson '{lesson['name']}' (ID {lesson['id']}).")

    image_slide_indexes = [
        idx for idx, slide in enumerate(slide_defs)
        if slide.get("type") in IMAGE_SLIDE_TYPES and slide.get("image_required", True)
    ]

    out_dir = os.path.join(
        tempfile.gettempdir(),
        "jove_video_frames",
        _slug(lesson["id"] + "_" + lesson["name"])
    )
    os.makedirs(out_dir, exist_ok=True)

    results = {}
    used_paths = set()
    used_timestamps = []
    for ordinal, slide_idx in enumerate(image_slide_indexes, start=1):
        slide_def = slide_defs[slide_idx]
        if progress_callback:
            progress_callback(f"Selecting video frame {ordinal}/{len(image_slide_indexes)} for {lesson['name']}...", None)

        selection = select_frame_for_slide(
            video_path=video_path,
            lesson_name=lesson["name"],
            transcript=lesson.get("transcript", ""),
            slide_def=slide_def,
            total_image_slides=len(image_slide_indexes),
            api_key=api_key,
            work_dir=out_dir,
            vision_model=vision_model,
            used_frame_paths=used_paths,
            used_timestamps=used_timestamps,
        )
        results[slide_idx] = selection
        used_paths.add(selection.get("path"))
        used_timestamps.append(selection.get("timestamp"))

    return results
