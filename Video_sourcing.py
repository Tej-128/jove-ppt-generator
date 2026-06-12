"""
JoVE Video Sourcing
Selects transcript-aligned frames from lesson MP4s for slide imagery.
No web fallback. All image-bearing slides must use frames from the lesson video.
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
    text = re.sub(r"[^A-Za-z0-9_-]+", "_", text.strip())
    return re.sub(r"_+", "_", text).strip("_") or "frame"


def _normalize_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip().lower()
    return re.sub(r"[^a-z0-9 ]+", "", text)


def estimate_anchor_ratio(transcript: str, anchor_text: str) -> float:
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

    clean = blur >= 40 and 25 <= brightness <= 235
    return {
        "brightness": brightness,
        "blur": blur,
        "transition_penalty": transition_penalty,
        "technical_score": technical_score,
        "clean": clean,
    }


def _save_frame(path: str, frame) -> None:
    cv2.imwrite(path, frame, [int(cv2.IMWRITE_JPEG_QUALITY), 92])


def _generate_candidate_times(target_time: float, duration: float, count: int) -> List[float]:
    count = max(3, count)
    offsets = [0]
    step = max(1.5, min(6.0, duration / 25.0 if duration else 2.5))
    k = 1
    while len(offsets) < count:
        offsets.extend([-step * k, step * k])
        k += 1
    times = []
    for off in offsets[:count]:
        t = max(0.2, min(max(0.2, duration - 0.2), target_time + off))
        if t not in times:
            times.append(t)
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
        _save_frame(img_path, frame)

        candidates.append({
            "path": img_path,
            "timestamp": round(t, 2),
            **metrics,
        })

    cap.release()
    return sorted(candidates, key=lambda x: x["technical_score"], reverse=True)


def extract_global_fallback_frames(video_path: str, count: int,
                                   output_dir: str, prefix: str) -> List[Dict]:
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        raise RuntimeError(f"Could not open video: {video_path}")
    stats = _get_video_stats(cap)
    duration = stats["duration"]
    if duration <= 0:
        cap.release()
        raise RuntimeError(f"Video has invalid duration: {video_path}")

    samples = []
    positions = np.linspace(0.05, 0.95, max(5, count + 2))
    for idx, ratio in enumerate(positions, start=1):
        t = float(ratio * duration)
        frame = _read_frame(cap, stats["fps"], t)
        if frame is None:
            continue
        prev_frame = _read_frame(cap, stats["fps"], max(0.0, t - 0.35))
        next_frame = _read_frame(cap, stats["fps"], min(duration, t + 0.35))
        metrics = _frame_metrics(frame, prev_frame, next_frame)
        img_name = f"{prefix}_global_{idx:02d}_{int(round(t * 1000)):07d}.jpg"
        img_path = os.path.join(output_dir, img_name)
        _save_frame(img_path, frame)
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
    content = [{
        "type": "text",
        "text": (
            f"Choose the best candidate image for a JoVE slide.\n"
            f"Lesson: {lesson_name}\n"
            f"Slide type: {slide_def.get('type')}\n"
            f"Slide title: {slide_def.get('title', lesson_name)}\n"
            f"Visual focus: {slide_def.get('visual_focus', '')}\n"
            f"Transcript anchor: {transcript_anchor}\n"
            f"Pick the image that best matches the concept while also looking clean and stable.\n"
            f"Prefer images that match the visual focus and avoid transition-like or unreadable frames.\n"
            f"Return only valid JSON like {{\"selected_index\": 1, \"confidence\": 97, \"reason\": \"...\"}}."
        )
    }]

    for idx, candidate in enumerate(candidates, start=1):
        content.append({
            "type": "text",
            "text": (
                f"Candidate {idx}: timestamp={candidate['timestamp']}s, "
                f"technical_score={candidate['technical_score']:.3f}, "
                f"blur={candidate['blur']:.1f}, brightness={candidate['brightness']:.1f}"
            )
        })
        content.append({"type": "image_url", "image_url": {"url": _image_to_data_url(candidate["path"])}})

    try:
        response = client.chat.completions.create(
            model=vision_model,
            messages=[{"role": "user", "content": content}],
            response_format={"type": "json_object"},
            temperature=0.1,
            max_tokens=500,
        )
        payload = json.loads(response.choices[0].message.content)
        selected_index = int(payload.get("selected_index", 1)) - 1
        selected_index = max(0, min(len(candidates) - 1, selected_index))
        selected = candidates[selected_index].copy()
        selected["vision_confidence"] = int(payload.get("confidence", 0))
        selected["selection_reason"] = payload.get("reason", "")
        selected["selection_method"] = "openai_vision"
        return selected
    except Exception:
        selected = candidates[0].copy()
        selected["vision_confidence"] = int(round(selected["technical_score"] * 100))
        selected["selection_reason"] = "Vision selection failed; used best technical candidate from same lesson video."
        selected["selection_method"] = "technical_fallback_same_video"
        return selected


def select_frame_for_slide(video_path: str, lesson_name: str, transcript: str,
                           slide_def: Dict, total_image_slides: int,
                           api_key: str, work_dir: str,
                           vision_model: str = "gpt-4.1") -> Dict:
    if not os.path.exists(video_path):
        raise RuntimeError(f"Missing MP4 for lesson '{lesson_name}': {video_path}")

    os.makedirs(work_dir, exist_ok=True)
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

    candidate_count = max(3, total_image_slides)
    candidates = extract_candidate_frames(video_path, target_time, candidate_count, work_dir, prefix)
    clean_candidates = [c for c in candidates if c.get("clean")]
    if not clean_candidates:
        clean_candidates = extract_global_fallback_frames(video_path, candidate_count, work_dir, prefix)

    if not clean_candidates:
        raise RuntimeError(f"No frame could be extracted from lesson MP4 for '{lesson_name}'.")

    top_candidates = clean_candidates[:candidate_count]
    client = OpenAI(api_key=api_key)
    selected = _vision_select_frame(
        client=client,
        slide_def=slide_def,
        candidates=top_candidates,
        lesson_name=lesson_name,
        transcript_anchor=anchor_text,
        vision_model=vision_model,
    )

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

    out_dir = os.path.join(tempfile.gettempdir(), "jove_video_frames", _slug(lesson["id"] + "_" + lesson["name"]))
    os.makedirs(out_dir, exist_ok=True)

    results = {}
    for ordinal, slide_idx in enumerate(image_slide_indexes, start=1):
        slide_def = slide_defs[slide_idx]
        if progress_callback:
            progress_callback(f"Selecting frame {ordinal}/{len(image_slide_indexes)} for {lesson['name']}...", None)
        selection = select_frame_for_slide(
            video_path=video_path,
            lesson_name=lesson["name"],
            transcript=lesson.get("transcript", ""),
            slide_def=slide_def,
            total_image_slides=len(image_slide_indexes),
            api_key=api_key,
            work_dir=out_dir,
            vision_model=vision_model,
        )
        results[slide_idx] = selection
    return results
