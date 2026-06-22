"""
JoVE PPT Builder V6 - strict formatting templates.

AI provides content only. This builder owns formatting decisions.
The goal is deterministic adherence to the JoVE presentation guide and the
approved project overrides.
"""

import os
import re
from typing import Iterable, List, Tuple

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

try:
    from PIL import Image
except Exception:  # Pillow should be installed by requirements.
    Image = None

try:
    from style_guide import (
        FOOTER_TEXT, FONT_PRIMARY,
        PRIMARY_DARK, BRAND_BLUE, BRAND_BLUE_LIGHT, WHITE, MID_GRAY,
        BLACK, DARK_GRAY, LIGHT_GRAY, SLIDE_W, SLIDE_H, MARGIN, RIGHT_X, RIGHT_W
    )
except Exception:
    FOOTER_TEXT = "Copyright © 2026 MyJoVE Corporation. All rights reserved"
    FONT_PRIMARY = "Roboto"
    PRIMARY_DARK = RGBColor(0x24, 0x29, 0x2F)
    BRAND_BLUE = RGBColor(0x6D, 0x9E, 0xEB)
    BRAND_BLUE_LIGHT = RGBColor(0xA4, 0xC2, 0xF4)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    MID_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
    BLACK = RGBColor(0x00, 0x00, 0x00)
    DARK_GRAY = RGBColor(0x4B, 0x55, 0x63)
    LIGHT_GRAY = RGBColor(0x85, 0x85, 0x85)
    SLIDE_W = Inches(20)
    SLIDE_H = Inches(11.25)
    MARGIN = Inches(0.75)
    RIGHT_X = Inches(11.0)
    RIGHT_W = Inches(8.25)

C_TEXT_DARK = PRIMARY_DARK
C_SUBTITLE = DARK_GRAY
C_COPYRIGHT = BLACK
C_ACCENT_BLUE = BRAND_BLUE
C_TABLE_HEADER = BRAND_BLUE
C_TABLE_ROW = WHITE  # approved override: all table body rows are white
C_WHITE = WHITE
C_BORDER = MID_GRAY
C_LIGHT_PANEL = RGBColor(0xF3, 0xF6, 0xFB)

FONT = FONT_PRIMARY

# Guide-driven layout constants.
SLIDE_SAFE_TOP = Inches(0.35)
SLIDE_SAFE_BOTTOM = Inches(10.35)
TITLE_SAFE_TOP_TABLE = Inches(2.05)
TITLE_BOX_H = Inches(1.25)
COVER_TITLE_MAX_W = Inches(10.9)
COVER_SINGLE_IMAGE_H = Inches(5.2)
LEFT = Inches(0.75)
TEXT_W = Inches(9.65)
GUTTER = Inches(0.30)
IMG_L = RIGHT_X
IMG_T = Inches(1.55)
IMG_W = RIGHT_W
IMG_H = Inches(8.45)
LOGO_L = Inches(18.444)
LOGO_T = Inches(0.326)
LOGO_W = Inches(1.087)
LOGO_H = Inches(0.551)
FOOTER_T = Inches(10.64)
FOOTER_H = Inches(0.28)

# Required font sizes.
FS_COVER_TITLE = 102
FS_SLIDE_TITLE = 48
FS_BODY = 30
FS_BODY_SECONDARY = 24
FS_TABLE_HEADER = 28
FS_TABLE_BODY = 24
FS_DISCUSSION_BADGE = 20
FS_CAPTION = 14
FS_FOOTER = 11

FORBIDDEN_LINE_RE = re.compile(r"^\s*(writer|author|reviewer|prepared\s*by|created\s*by)\s*[:\-]", re.I)


def _clean_text(text) -> str:
    """Remove markdown and forbidden metadata before writing to PPT."""
    if text is None:
        return ""
    value = str(text)
    value = value.replace("**", "")
    value = value.replace("__", "")
    value = value.replace("`", "")
    value = re.sub(r"\[(INSERT IMAGE|TODO|PLACEHOLDER|IMAGE)\]", "", value, flags=re.I)
    lines = []
    for line in value.splitlines():
        if FORBIDDEN_LINE_RE.search(line):
            continue
        lines.append(line.strip())
    value = "\n".join(line for line in lines if line)
    return re.sub(r"[ \t]+", " ", value).strip()



def _safe_slide_title(title: str, body_text: str = "") -> str:
    """Prevent numeric lesson IDs or metadata from becoming visible slide titles."""
    raw = _clean_text(title or "")
    if not raw or re.fullmatch(r"(lesson\s*)?\d{4,8}", raw.strip(), flags=re.I):
        body = _clean_text(body_text or "")
        m = re.match(r"([A-Z][A-Za-z\- ]{2,35}?)(?:\s+are|\s+is|\s+include|\s+consist|\s+form|\s+have)\b", body)
        if m:
            return m.group(1).strip()
        words = re.findall(r"[A-Za-z][A-Za-z\-]+", body)[:4]
        return " ".join(words).title() if words else "Core Concept"
    return raw


def _fit_title_font(title: str, target=FS_SLIDE_TITLE, min_size=34) -> int:
    title = str(title or "")
    if len(title) <= 58:
        return target
    if len(title) <= 72:
        return max(min_size, target - 6)
    if len(title) <= 88:
        return max(min_size, target - 10)
    return max(min_size, target - 14)


def _fit_cover_font(title: str) -> int:
    title = str(title or "")
    if len(title) <= 14:
        return FS_COVER_TITLE
    if len(title) <= 22:
        return 94
    if len(title) <= 32:
        return 84
    return 76


def _shorten_words(text: str, max_words: int) -> str:
    text = _clean_text(text)
    words = text.split()
    if len(words) <= max_words:
        return text
    return " ".join(words[:max_words]).rstrip(".,;:") + "."


def _font(run, size_pt, bold=False, italic=False, color=None):
    run.font.name = FONT
    try:
        run._r.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii", FONT)
        run._r.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi", FONT)
    except Exception:
        pass
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color


def _tb(slide, left, top, width, height, text, size_pt,
        bold=False, italic=False, color=None, align=PP_ALIGN.LEFT, wrap=True):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = wrap
    tf.vertical_anchor = MSO_ANCHOR.TOP
    try:
        tf.margin_left = Inches(0.0)
        tf.margin_right = Inches(0.0)
        tf.margin_top = Inches(0.0)
        tf.margin_bottom = Inches(0.0)
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _clean_text(text)
    _font(run, size_pt, bold, italic, color or C_TEXT_DARK)
    return box


def _white_bg(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = C_WHITE


def _logo(slide, logo_path):
    """Place slide-level JoVE logo in one exact location on every slide."""
    if logo_path and os.path.exists(logo_path):
        try:
            pic = slide.shapes.add_picture(logo_path, LOGO_L, LOGO_T, width=LOGO_W, height=LOGO_H)
            return pic
        except Exception:
            # Do not create inconsistent fallback logos.
            return None



def _copyright(slide):
    _tb(slide, Inches(5.35), FOOTER_T, Inches(9.3), FOOTER_H,
        FOOTER_TEXT, FS_FOOTER, color=C_COPYRIGHT, align=PP_ALIGN.CENTER)


def _slide_number(slide, number):
    if number is None:
        return
    _tb(slide, Inches(18.65), Inches(10.62), Inches(0.55), Inches(0.28),
        str(number), FS_FOOTER, color=LIGHT_GRAY, align=PP_ALIGN.RIGHT)


def _notes(slide, text):
    text = _clean_text(text)
    if text:
        slide.notes_slide.notes_text_frame.text = text


def _base_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _add_image_contain(slide, image_path, left, top, width, height):
    if not image_path or not os.path.exists(image_path):
        raise ValueError("A valid local image_path is required for every image-bearing slide. No placeholders are allowed.")
    if Image is None:
        slide.shapes.add_picture(image_path, left, top, width=width)
        return
    try:
        with Image.open(image_path) as img:
            iw, ih = img.size
        if iw <= 0 or ih <= 0:
            slide.shapes.add_picture(image_path, left, top, width=width)
            return
        box_ratio = float(width) / float(height)
        img_ratio = float(iw) / float(ih)
        if img_ratio >= box_ratio:
            final_w = width
            final_h = int(width / img_ratio)
        else:
            final_h = height
            final_w = int(height * img_ratio)
        x = left + int((width - final_w) / 2)
        y = top + int((height - final_h) / 2)
        slide.shapes.add_picture(image_path, x, y, width=final_w, height=final_h)
    except Exception:
        slide.shapes.add_picture(image_path, left, top, width=width)


def _image(slide, image_path):
    _add_image_contain(slide, image_path, IMG_L, IMG_T, IMG_W, IMG_H)
    return True


def _body_text(slide, body_text, top, max_words=68):
    body_text = _shorten_words(body_text, max_words)
    box = slide.shapes.add_textbox(LEFT, top, TEXT_W, SLIDE_SAFE_BOTTOM - top)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    try:
        tf.margin_left = Inches(0.0)
        tf.margin_right = Inches(0.0)
        tf.margin_top = Inches(0.0)
        tf.margin_bottom = Inches(0.0)
    except Exception:
        pass

    lines = [ln.strip() for ln in body_text.split("\n") if ln.strip()] or [""]
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        if idx > 0:
            p.space_before = Pt(10)
        p.alignment = PP_ALIGN.LEFT
        segments = re.split(r'\*\*(.+?)\*\*', line)
        for i, seg in enumerate(segments):
            if not seg:
                continue
            run = p.add_run()
            run.text = _clean_text(seg)
            _font(run, FS_BODY, bold=(i % 2 == 1), color=C_TEXT_DARK)


def _normalize_table(headers, rows, table_kind=None):
    headers = [_clean_text(h) for h in (headers or [])]
    rows = rows or []
    if not headers:
        headers = ["Concept", "Definition/Meaning", "Example/Application"]

    n_cols = max(1, len(headers))
    normalized = []
    for row in rows:
        row = [_clean_text(x) for x in list(row)]
        if len(row) < n_cols:
            row += [""] * (n_cols - len(row))
        normalized.append(row[:n_cols])
    if not normalized:
        normalized = [[""] * n_cols]
    return headers, normalized



def _set_cell_border_blue(cell):
    """Apply JoVE-blue borders to each table cell."""
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        for edge in ("lnL", "lnR", "lnT", "lnB"):
            existing = tcPr.find(qn(f"a:{edge}"))
            if existing is not None:
                tcPr.remove(existing)
            ln = OxmlElement(f"a:{edge}")
            ln.set("w", "12700")
            ln.set("cap", "flat")
            ln.set("cmpd", "sng")
            ln.set("algn", "ctr")
            solid = OxmlElement("a:solidFill")
            srgb = OxmlElement("a:srgbClr")
            srgb.set("val", "6D9EEB")
            solid.append(srgb)
            ln.append(solid)
            prst = OxmlElement("a:prstDash")
            prst.set("val", "solid")
            ln.append(prst)
            tcPr.append(ln)
    except Exception:
        pass


def _cell_text(cell, text, size, bold=False, align=PP_ALIGN.CENTER, color=None):
    try:
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    except Exception:
        pass
    tf = cell.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _clean_text(text)
    _font(run, size, bold=bold, color=color or C_TEXT_DARK)



def _table_cell_limit(ci: int, n_cols: int, has_images: bool) -> int:
    # Precise but useful. Avoid long paragraphs, but do not strip meaning.
    if ci == 0:
        return 7
    if has_images and ci == n_cols - 1:
        return 0
    if "Example" and has_images and ci == n_cols - 2:
        return 18
    return 20


def _table_cell_limit(ci: int, n_cols: int, has_images: bool) -> int:
    # Keep the earlier useful table detail: precise, but not vague.
    # Image column should not reduce Example/Application meaning.
    if ci == 0:
        return 8
    if has_images and ci == n_cols - 1:
        return 0
    if has_images and ci == n_cols - 2:
        return 26
    return 28


def _add_table(slide, headers, rows, left, top, width, max_height=Inches(7.95), row_image_paths=None, max_rows=4):
    headers, rows = _normalize_table(headers, rows)
    row_image_paths = list(row_image_paths or [])

    if len(rows) > max_rows:
        rows = rows[:max_rows]
        row_image_paths = row_image_paths[:max_rows]

    # Confirmed rule: every table must have an Image column.
    # If images exist, populate that column. If no image exists for a row, keep the column blank.
    lower_headers = [str(h).strip().lower() for h in headers]
    has_image_col = any(h in {"image", "visual", "figure"} for h in lower_headers)
    rows = [list(row) for row in rows]
    if not has_image_col:
        headers.append("Image")
        rows = [row + [""] for row in rows]
    else:
        for row in rows:
            if len(row) < len(headers):
                row.extend([""] * (len(headers) - len(row)))
            row[-1] = ""

    n_rows = len(rows) + 1
    n_cols = len(headers)
    max_allowed_height = max(0, SLIDE_SAFE_BOTTOM - top)
    height = min(max_height, max_allowed_height)
    row_h = height / max(1, n_rows)

    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table

    # Image column is always present; keep it compact so text columns stay useful.
    image_col_w = int(width * 0.18)
    remaining = int(width) - image_col_w
    if n_cols == 4:
        widths = [int(remaining * 0.23), int(remaining * 0.42), int(remaining * 0.35), image_col_w]
    elif n_cols == 5:
        widths = [int(remaining * 0.18), int(remaining * 0.28), int(remaining * 0.29), int(remaining * 0.25), image_col_w]
    else:
        widths = [int(remaining / max(1, n_cols - 1))] * (n_cols - 1) + [image_col_w]
    drift = int(width) - sum(widths)
    if len(widths) >= 2:
        widths[-2] += drift
    for ci, cw in enumerate(widths[:n_cols]):
        tbl.columns[ci].width = cw

    header_size = FS_TABLE_HEADER if n_cols <= 4 else 24
    body_size = FS_TABLE_BODY if n_cols <= 4 else 22

    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_TABLE_HEADER
        _set_cell_border_blue(cell)
        try:
            cell.margin_left = Inches(0.07)
            cell.margin_right = Inches(0.07)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
        except Exception:
            pass
        _cell_text(cell, _shorten_words(h, 8), header_size, bold=True, align=PP_ALIGN.CENTER, color=C_WHITE)

    for ri, row in enumerate(rows):
        for ci in range(n_cols):
            cell = tbl.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE
            _set_cell_border_blue(cell)
            try:
                cell.margin_left = Inches(0.07)
                cell.margin_right = Inches(0.07)
                cell.margin_top = Inches(0.04)
                cell.margin_bottom = Inches(0.04)
            except Exception:
                pass

            is_image_col = ci == n_cols - 1
            val = row[ci] if ci < len(row) else ""
            if is_image_col:
                _cell_text(cell, "", body_size)
            else:
                limit = _table_cell_limit(ci, n_cols, True)
                _cell_text(
                    cell,
                    _shorten_words(val, limit),
                    body_size,
                    bold=(ci == 0),
                    align=PP_ALIGN.CENTER
                )

    # Place images inside the Image column only. Images never replace useful text.
    if row_image_paths:
        image_col = n_cols - 1
        x = left + sum(tbl.columns[ci].width for ci in range(image_col))
        col_w = tbl.columns[image_col].width
        for ri, img_path in enumerate(row_image_paths[:len(rows)]):
            if not img_path or not os.path.exists(img_path):
                continue
            y = top + row_h * (ri + 1)
            pad = Inches(0.08)
            _add_image_contain(slide, img_path, x + pad, y + pad, col_w - pad * 2, row_h - pad * 2)
    return tbl



def create_presentation(logo_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def build_cover_slide(prs, chapter_name, chapter_number, logo_path, cover_image_path=None, slide_number=None, cover_image_paths=None, chapter_description=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    title = _safe_slide_title(chapter_name, "")
    title_font = _fit_cover_font(title)
    _tb(slide, LEFT, Inches(1.20), COVER_TITLE_MAX_W, Inches(2.7),
        title, title_font, bold=True, color=C_TEXT_DARK, wrap=False)
    desc = _clean_text(chapter_description or "")
    if desc:
        _tb(slide, LEFT, Inches(4.25), Inches(8.9), Inches(0.9),
            _shorten_words(desc, 26), 26, italic=True, color=DARK_GRAY, wrap=True)
    _tb(slide, LEFT, Inches(5.60), Inches(8.8), Inches(0.45),
        f"Chapter {chapter_number}", 24, color=DARK_GRAY)
    _tb(slide, LEFT, Inches(8.55), Inches(6.0), Inches(0.45),
        "Lecture Slides", 24, color=DARK_GRAY)

    # Project override: first chapter/cover slide must use exactly ONE image.
    # No stacked image set and no repeated image.
    selected_cover_image = None
    if cover_image_path and os.path.exists(cover_image_path):
        selected_cover_image = cover_image_path
    elif cover_image_paths:
        for img in cover_image_paths:
            if img and os.path.exists(img):
                selected_cover_image = img
                break

    if selected_cover_image:
        _add_image_contain(slide, selected_cover_image, RIGHT_X, Inches(2.25), RIGHT_W, COVER_SINGLE_IMAGE_H)

    _notes(slide, "Welcome students. Introduce the lesson topic and outline the key concepts.")



def build_concept_slide(prs, lesson_name, body_text, sub_label=None,
                        image_path=None, speaker_notes=None, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    title = _safe_slide_title(lesson_name, body_text)
    if title.lower().startswith("definition") or title.lower() in {"core idea", "definition and core process"}:
        title = "What is the concept?"
    _tb(slide, LEFT, Inches(0.75), TEXT_W, Inches(1.15),
        title, FS_SLIDE_TITLE, bold=True, color=C_TEXT_DARK)
    _body_text(slide, body_text, Inches(2.55), max_words=65)
    _image(slide, image_path)
    _notes(slide, speaker_notes)


def build_table_slide(prs, lesson_name, headers, rows, sub_title=None,
                      table_kind=None, image_path=None, row_image_paths=None, speaker_notes=None, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    seed_text = " ".join(" ".join(map(str, r)) if isinstance(r, list) else str(r) for r in (rows or [])[:2])
    title = _safe_slide_title(sub_title or lesson_name, seed_text)
    _tb(slide, LEFT, Inches(0.42), Inches(17.25), TITLE_BOX_H,
        title, FS_SLIDE_TITLE, bold=True, color=C_TEXT_DARK, wrap=True)

    _add_table(slide, headers, rows, Inches(0.75), TITLE_SAFE_TOP_TABLE, Inches(18.5),
               max_height=Inches(7.8), row_image_paths=row_image_paths, max_rows=4)
    _notes(slide, speaker_notes)


def _discussion_header(slide):
    _tb(slide, LEFT, Inches(0.65), TEXT_W, Inches(0.8),
        "Discussion", FS_SLIDE_TITLE, bold=True, color=C_TEXT_DARK)
    pill = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, LEFT, Inches(1.55), Inches(3.0), Inches(0.42))
    pill.fill.solid()
    pill.fill.fore_color.rgb = C_ACCENT_BLUE
    pill.line.color.rgb = C_ACCENT_BLUE
    _tb(slide, LEFT + Inches(0.18), Inches(1.62), Inches(2.65), Inches(0.25),
        "Discuss with the class", FS_DISCUSSION_BADGE, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)


def build_discussion_question_slide(prs, lesson_name, question_text,
                                     hint_text=None, image_path=None,
                                     speaker_notes=None, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _discussion_header(slide)
    _tb(slide, LEFT, Inches(2.45), TEXT_W, Inches(2.8),
        _shorten_words(question_text, 30), 40, bold=True, color=C_TEXT_DARK)
    if hint_text:
        _tb(slide, LEFT, Inches(5.65), TEXT_W, Inches(1.0),
            "Hint: " + _shorten_words(hint_text, 20), FS_BODY_SECONDARY, italic=True, color=DARK_GRAY)
    _image(slide, image_path)
    _notes(slide, speaker_notes)


def build_discussion_answer_slide(prs, lesson_name, answer_summary,
                                   answer_explanation, image_path=None,
                                   speaker_notes=None, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _discussion_header(slide)
    _tb(slide, LEFT, Inches(2.35), TEXT_W, Inches(1.1),
        "Answer: " + _shorten_words(answer_summary, 8), 32, bold=True, color=C_TEXT_DARK)
    _tb(slide, LEFT, Inches(3.75), TEXT_W, Inches(4.8),
        _shorten_words(answer_explanation, 60), FS_BODY, color=C_TEXT_DARK)
    _image(slide, image_path)
    _notes(slide, speaker_notes)


def build_summary_slide(prs, summary_statement, table_headers=None,
                         table_rows=None, logo_path="", speaker_notes=None, slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _tb(slide, LEFT, Inches(0.55), Inches(17.0), Inches(0.95),
        "Summary", FS_SLIDE_TITLE, bold=True, color=C_TEXT_DARK)

    clean_summary = _shorten_words(_clean_text(summary_statement), 22)
    _tb(slide, LEFT, Inches(1.50), Inches(17.0), Inches(0.75),
        clean_summary, 28, italic=True, color=C_TEXT_DARK)

    rows = (table_rows or [])[:3]
    if table_headers and rows:
        _add_table(slide, table_headers, rows, LEFT, Inches(2.55), Inches(17.0),
                   max_height=Inches(5.75), row_image_paths=None, max_rows=3)
    _notes(slide, speaker_notes)


def build_glossary_slide(prs, terms_dict, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _tb(slide, LEFT, Inches(0.65), Inches(17.0), Inches(0.9),
        "Glossary", FS_SLIDE_TITLE, bold=True, color=C_TEXT_DARK)

    items = list(terms_dict.items())[:6]
    box = slide.shapes.add_textbox(LEFT, Inches(1.75), Inches(17.2), Inches(8.1))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    first = True
    for term, definition in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        if p is not tf.paragraphs[0]:
            p.space_before = Pt(10)
        r1 = p.add_run()
        r1.text = f"{_clean_text(term)}: "
        _font(r1, FS_BODY_SECONDARY, bold=True, color=C_TEXT_DARK)
        r2 = p.add_run()
        r2.text = _shorten_words(definition, 22)
        _font(r2, FS_BODY_SECONDARY, color=C_TEXT_DARK)
    _notes(slide, "Review these key terms with students.")
