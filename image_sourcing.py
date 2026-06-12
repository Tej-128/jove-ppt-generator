"""
JoVE Image Sourcing

IMPORTANT:
The active workflow now uses lesson MP4 frames through video_sourcing.py.
No web image search is used by pipeline.py.

This file keeps the older Google/Wikimedia helpers for compatibility/manual testing,
but find_image() is intentionally disabled so the app cannot silently fall back to
web images or placeholders.
"""

import os
import requests
from typing import Optional

from video_sourcing import assign_frames_to_slides, select_frame_for_slide


REJECT_TERMS = [
    "logo", "icon", "flag", "symbol", "button", "banner",
    "map of", "coat of arms", "stamp", "coin", "currency",
    "painting by", "portrait of", "statue", "monument",
    "album cover", "movie poster", "book cover",
]

REJECT_EXTENSIONS = [".svg", ".ico", ".gif"]


def _is_relevant(url: str, title: str = "") -> bool:
    """Quick heuristic filter: reject obviously irrelevant images."""
    combined = (url + " " + title).lower()
    if any(ext in combined for ext in REJECT_EXTENSIONS):
        return False
    if any(term in combined for term in REJECT_TERMS):
        return False
    return True


def search_google_images(query: str, api_key: str, cse_id: str) -> Optional[str]:
    """
    Legacy helper only. Not used by the active PPT pipeline.
    """
    if not api_key or not cse_id:
        return None

    try:
        url = "https://www.googleapis.com/customsearch/v1"
        params = {
            "key": api_key,
            "cx": cse_id,
            "q": query,
            "searchType": "image",
            "num": 8,
            "safe": "active",
            "imgType": "photo",
        }
        r = requests.get(url, params=params, timeout=10)
        if r.status_code != 200:
            return None

        data = r.json()
        items = data.get("items", [])

        for item in items:
            link = item.get("link", "")
            title = item.get("title", "")
            if _is_relevant(link, title):
                return link

    except Exception:
        pass

    return None


def search_wikimedia_image(query: str) -> Optional[str]:
    """
    Legacy helper only. Not used by the active PPT pipeline.
    """
    headers = {"User-Agent": "JoVE-PPT-Generator/1.0 (educational use)"}

    try:
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

        titles = "|".join([res["title"] for res in results[:5]])
        img_params = {
            "action": "query",
            "prop": "pageimages",
            "titles": titles,
            "pithumbsize": "800",
            "pilimit": "5",
            "format": "json"
        }
        r2 = requests.get(search_url, params=img_params, timeout=10, headers=headers)
        pages = r2.json().get("query", {}).get("pages", {})

        for page in pages.values():
            thumb = page.get("thumbnail", {})
            src = thumb.get("source", "")
            title = page.get("title", "")
            if src and _is_relevant(src, title):
                return src

    except Exception:
        pass

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
            title = page.get("title", "")
            if (mime.startswith("image/") and w > 300 and h > 200
                    and _is_relevant(img_url, title)):
                return img_url
    except Exception:
        pass

    return None


def find_image(query: str, google_api_key: str = "", google_cse_id: str = "") -> Optional[str]:
    """
    Disabled by design.

    The new production requirement is:
    every image-bearing slide must use a frame from that lesson's MP4.
    """
    raise RuntimeError(
        "Web image fallback is disabled. Use video_sourcing.assign_frames_to_slides() so every image comes from the lesson MP4."
    )
