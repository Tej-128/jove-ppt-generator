"""
JoVE PPT Builder - video-frame version

Preserves the original JoVE layout intent, but removes image placeholders.
All image-bearing slides require a valid local image_path from the lesson MP4.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
import os
import re

try:
    from style_guide import (
        FOOTER_TEXT, FONT_PRIMARY, FONT_FALLBACK_1,
        PRIMARY_DARK, BRAND_BLUE, BRAND_BLUE_LIGHT, WHITE, MID_GRAY,
        BLACK, DARK_GRAY, LIGHT_GRAY, SLIDE_W, SLIDE_H, MARGIN, RIGHT_X, RIGHT_W
    )
except Exception:
    FOOTER_TEXT = "Copyright © 2026 MyJoVE Corporation. All rights reserved"
    FONT_PRIMARY = "Roboto"
    FONT_FALLBACK_1 = "Helvetica Neue"
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
C_SUBTITLE = LIGHT_GRAY
C_COPYRIGHT = BLACK
C_ACCENT_BLUE = BRAND_BLUE
C_TABLE_HEADER = BRAND_BLUE
C_TABLE_ROW = BRAND_BLUE_LIGHT
C_WHITE = WHITE
C_LIGHT_PANEL = RGBColor(0xF3, 0xF6, 0xFB)

FONT = FONT_PRIMARY

LEFT = Inches(1.042)
TEXT_W = Inches(7.0)
IMG_L = Inches(8.8)
IMG_T = Inches(1.5)
IMG_W = Inches(10.5)
IMG_H = Inches(8.8)
LOGO_L = Inches(18.444)
LOGO_T = Inches(0.326)
LOGO_W = Inches(1.087)
LOGO_H = Inches(0.551)
CPY_L = Inches(7.5)
CPY_T = Inches(10.55)
CPY_W = Inches(5.5)
CPY_H = Inches(0.45)


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
    tf.word_wrap = wrap
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text or "")
    _font(run, size_pt, bold, italic, color or C_TEXT_DARK)
    return box


def _white_bg(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = C_WHITE


def _logo(slide, logo_path):
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, LOGO_L, LOGO_T, LOGO_W, LOGO_H)


def _copyright(slide):
    _tb(slide, Inches(5.4), Inches(10.64), Inches(9.2), Inches(0.28),
        FOOTER_TEXT, 11, color=C_COPYRIGHT, align=PP_ALIGN.CENTER)


def _slide_number(slide, number):
    if number is None:
        return
    _tb(slide, Inches(18.65), Inches(10.62), Inches(0.55), Inches(0.28),
        str(number), 11, color=C_SUBTITLE, align=PP_ALIGN.RIGHT)


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = str(text)


def _base_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _image(slide, image_path):
    if not image_path or not os.path.exists(image_path):
        raise ValueError("A valid local image_path is required for every image-bearing slide. No placeholders are allowed.")
    slide.shapes.add_picture(image_path, IMG_L, IMG_T, IMG_W, IMG_H)
    return True


def _add_cover_panel(slide, image_path=None):
    panel = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, IMG_L, IMG_T, IMG_W, IMG_H)
    panel.fill.solid()
    panel.fill.fore_color.rgb = C_LIGHT_PANEL
    panel.line.color.rgb = C_LIGHT_PANEL

    if image_path and os.path.exists(image_path):
        slide.shapes.add_picture(image_path, IMG_L, IMG_T, IMG_W, IMG_H)
    else:
        _tb(slide, Inches(10.05), Inches(4.25), Inches(8.25), Inches(1.0),
            "JoVE Lecture Deck", 34, bold=True, color=C_ACCENT_BLUE, align=PP_ALIGN.CENTER)
        _tb(slide, Inches(10.05), Inches(5.1), Inches(8.25), Inches(0.8),
            "Video-aligned presentation", 24, color=C_SUBTITLE, align=PP_ALIGN.CENTER)


def _body_text(slide, body_text, top):
    box = slide.shapes.add_textbox(LEFT, top, TEXT_W, Inches(10.25) - top)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP

    lines = [ln.strip() for ln in str(body_text or "").strip().split("\n") if ln.strip()]
    if not lines:
        lines = [""]

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
            run.text = seg
            _font(run, 25, bold=(i % 2 == 1), color=C_TEXT_DARK)


def _table_needs_definition_example(headers, table_kind=None):
    table_kind = str(table_kind or "").lower().strip()
    if table_kind in {"definition_example", "definitions", "terms", "steps", "types", "stages", "method", "variables"}:
        return True
    if table_kind in {"comparison", "timeline", "cause_effect", "pros_cons", "inputs_outputs", "other"}:
        return False

    text_blob = " ".join(str(h).lower() for h in (headers or []))
    semantic_terms = ["term", "definition", "meaning", "step", "stage", "type", "method", "variable", "concept"]
    return any(x in text_blob for x in semantic_terms)


def _normalize_table(headers, rows, table_kind=None):
    """
    Normalize row lengths without forcing every table to 3 columns.

    Only term/step/type/stage/method tables are upgraded to include
    Definition/Meaning and Example/Application columns. Comparison/timeline/
    cause-effect tables keep their natural structure.
    """
    headers = [str(h) for h in (headers or [])]
    rows = rows or []

    if not headers:
        headers = ["Concept", "Key Point"]

    lower = [h.lower() for h in headers]
    has_definition = any("definition" in h or "meaning" in h for h in lower)
    has_example = any("example" in h or "application" in h for h in lower)

    if _table_needs_definition_example(headers, table_kind) and (not has_definition or not has_example):
        first_header = headers[0] if headers else "Term/Step/Type"
        headers = [first_header, "Definition/Meaning", "Example/Application"]
        converted_rows = []
        for row in rows:
            row = [str(x) for x in list(row)]
            if len(row) == 0:
                row = ["", "", ""]
            elif len(row) == 1:
                row = [row[0], "", ""]
            elif len(row) == 2:
                row = [row[0], row[1], ""]
            else:
                row = [row[0], row[1], "; ".join(row[2:])]
            converted_rows.append(row)
        rows = converted_rows

    n_cols = max(1, len(headers))
    normalized_rows = []
    for row in rows:
        row = [str(x) for x in list(row)]
        if len(row) < n_cols:
            row += [""] * (n_cols - len(row))
        normalized_rows.append(row[:n_cols])

    if not normalized_rows:
        normalized_rows = [[""] * n_cols]

    return headers, normalized_rows

def _add_table(slide, headers, rows, left, top, width, max_height=Inches(8.1), row_image_paths=None):
    headers, rows = _normalize_table(headers, rows)
    row_image_paths = list(row_image_paths or [])

    if row_image_paths:
        # If the table already has an Example/Image/Visual column, use the rightmost column.
        # Otherwise add a dedicated Image column so example text is not overwritten.
        lower = [h.lower() for h in headers]
        if not any(("image" in h or "visual" in h or "example" in h or "graph" in h) for h in lower):
            headers.append("Image")
            rows = [list(row) + [""] for row in rows]
        else:
            rows = [list(row) for row in rows]
            # Clear text in rightmost cell if images are being overlaid there.
            for row in rows:
                if len(row) < len(headers):
                    row.extend([""] * (len(headers) - len(row)))
                row[-1] = ""

    n_rows = len(rows) + 1
    n_cols = len(headers)
    row_h = min(Inches(1.45), max(Inches(0.72), max_height / max(1, n_rows)))
    height = min(max_height, row_h * n_rows)

    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table

    # Wider final image column when row images are present.
    if row_image_paths and n_cols >= 3:
        image_col_w = int(width * 0.30)
        remaining = int(width) - image_col_w
        text_col_w = int(remaining / (n_cols - 1))
        for ci in range(n_cols - 1):
            tbl.columns[ci].width = text_col_w
        tbl.columns[n_cols - 1].width = image_col_w
    else:
        for ci in range(n_cols):
            tbl.columns[ci].width = int(width / n_cols)

    header_font = 24 if n_cols <= 4 else 20
    body_font = 22 if n_cols <= 3 and n_rows <= 4 else 18

    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_TABLE_HEADER
        try:
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
        except Exception:
            pass
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(h)
        _font(run, header_font, bold=True, color=C_WHITE)

    for ri, row in enumerate(rows):
        for ci in range(n_cols):
            val = row[ci] if ci < len(row) else ""
            cell = tbl.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE if ri % 2 == 0 else C_TABLE_ROW
            try:
                cell.margin_left = Inches(0.08)
                cell.margin_right = Inches(0.08)
                cell.margin_top = Inches(0.05)
                cell.margin_bottom = Inches(0.05)
            except Exception:
                pass
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if ci == 0 else PP_ALIGN.LEFT
            run = p.add_run()
            run.text = str(val)
            _font(run, body_font, bold=(ci == 0), color=C_TEXT_DARK)

    # Overlay images in the rightmost column, fitting within row boxes and slide bounds.
    if row_image_paths:
        image_col = n_cols - 1
        x = left + sum(tbl.columns[ci].width for ci in range(image_col))
        col_w = tbl.columns[image_col].width
        for ri, img_path in enumerate(row_image_paths[:len(rows)]):
            if not img_path or not os.path.exists(img_path):
                continue
            y = top + row_h * (ri + 1)
            pad = Inches(0.08)
            max_w = col_w - pad * 2
            max_h = row_h - pad * 2
            try:
                slide.shapes.add_picture(img_path, x + pad, y + pad, max_w, max_h)
            except Exception:
                pass

    return tbl

def create_presentation(logo_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def build_cover_slide(prs, chapter_name, chapter_number, logo_path, cover_image_path=None, slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _tb(slide, LEFT, Inches(2.0), Inches(7.8), Inches(5.0),
        chapter_name, 68, bold=True, color=C_TEXT_DARK, wrap=True)
    _tb(slide, LEFT, Inches(7.0), Inches(7.0), Inches(0.8),
        f"Chapter {chapter_number}", 32, bold=True, color=C_SUBTITLE)
    _tb(slide, LEFT, Inches(9.5), Inches(5.0), Inches(0.8),
        "Lecture Slides", 32, bold=True, color=C_SUBTITLE)
    _add_cover_panel(slide, cover_image_path)
    _notes(slide, "Welcome students. Introduce the chapter topic and outline the key lessons they will cover today.")


def build_concept_slide(prs, lesson_name, body_text, sub_label=None,
                        image_path=None, speaker_notes=None, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    if sub_label:
        _tb(slide, LEFT, Inches(0.38), TEXT_W, Inches(0.5),
            sub_label, 22, bold=True, color=C_SUBTITLE)
        title_top = Inches(0.92)
    else:
        title_top = Inches(0.38)

    _tb(slide, LEFT, title_top, TEXT_W, Inches(1.45),
        lesson_name, 34, bold=True, color=C_TEXT_DARK)

    body_top = title_top + Inches(1.65)
    _body_text(slide, body_text, body_top)
    _image(slide, image_path)
    _notes(slide, speaker_notes)


def build_table_slide(prs, lesson_name, headers, rows, sub_title=None,
                      table_kind=None, image_path=None, row_image_paths=None, speaker_notes=None, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    title = sub_title or lesson_name
    if sub_title:
        _tb(slide, LEFT, Inches(0.30), Inches(18.2), Inches(0.36),
            lesson_name, 18, bold=True, color=C_SUBTITLE)
        _tb(slide, LEFT, Inches(0.70), Inches(18.2), Inches(0.78),
            title, 38, bold=True, color=C_TEXT_DARK)
        table_top = Inches(1.65)
    else:
        _tb(slide, LEFT, Inches(0.45), Inches(18.2), Inches(0.9),
            title, 40, bold=True, color=C_TEXT_DARK)
        table_top = Inches(1.65)

    # Guideline: table slides are full-width and images belong inside the table, not beside it.
    _add_table(
        slide,
        headers,
        rows,
        Inches(0.75),
        table_top,
        Inches(18.5),
        max_height=Inches(8.55),
        row_image_paths=row_image_paths
    )
    _notes(slide, speaker_notes)

def build_discussion_question_slide(prs, lesson_name, question_text,
                                     hint_text=None, image_path=None,
                                     speaker_notes=None, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _tb(slide, LEFT, Inches(0.38), TEXT_W + Inches(1.5), Inches(1.2),
        lesson_name, 36, bold=True, color=C_TEXT_DARK)
    _tb(slide, LEFT, Inches(1.75), Inches(5.5), Inches(0.6),
        "Discuss with the class", 26, color=C_ACCENT_BLUE)

    q_box = slide.shapes.add_textbox(LEFT, Inches(2.55), TEXT_W + Inches(0.5), Inches(2.8))
    tf = q_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = str(question_text or "")
    _font(run, 31, bold=True, color=C_TEXT_DARK)

    if hint_text:
        _tb(slide, LEFT, Inches(5.6), TEXT_W + Inches(0.5), Inches(1.5),
            f"Hint: {hint_text}", 23, italic=True, color=C_TEXT_DARK)

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

    _tb(slide, LEFT, Inches(0.38), TEXT_W + Inches(1.5), Inches(1.2),
        lesson_name, 36, bold=True, color=C_TEXT_DARK)
    _tb(slide, LEFT, Inches(1.75), Inches(5.5), Inches(0.6),
        "Discuss with the class", 26, color=C_ACCENT_BLUE)
    _tb(slide, LEFT, Inches(2.55), TEXT_W + Inches(0.5), Inches(1.45),
        f"Answer: {answer_summary}", 29, bold=True, color=C_TEXT_DARK)

    ans_box = slide.shapes.add_textbox(LEFT, Inches(4.15), TEXT_W + Inches(0.5), Inches(4.8))
    tf = ans_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = str(answer_explanation or "")
    _font(run, 24, italic=True, color=C_TEXT_DARK)

    _image(slide, image_path)
    _notes(slide, speaker_notes)


def build_summary_slide(prs, summary_statement, table_headers=None,
                         table_rows=None, logo_path="", speaker_notes=None, slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _tb(slide, LEFT, Inches(0.38), Inches(18.5), Inches(0.5),
        "SUMMARY", 22, bold=True, color=C_TEXT_DARK)

    _tb(slide, LEFT, Inches(0.9), Inches(18.0), Inches(1.8),
        summary_statement, 36, bold=True, color=C_TEXT_DARK)

    if table_headers and table_rows:
        _add_table(slide, table_headers, table_rows, LEFT, Inches(3.05), Inches(17.9))

    _notes(slide, speaker_notes)


def build_glossary_slide(prs, terms_dict, logo_path="", slide_number=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)
    _slide_number(slide, slide_number)

    _tb(slide, LEFT, Inches(0.38), Inches(18.5), Inches(0.8),
        "Glossary", 44, bold=True, color=C_TEXT_DARK)

    box = slide.shapes.add_textbox(LEFT, Inches(1.4), Inches(18.0), Inches(9.0))
    tf = box.text_frame
    tf.word_wrap = True

    first = True
    for term, definition in terms_dict.items():
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        if p is not tf.paragraphs[0]:
            p.space_before = Pt(8)
        p.alignment = PP_ALIGN.LEFT

        r1 = p.add_run()
        r1.text = f"{term}: "
        _font(r1, 24, bold=True, color=C_TEXT_DARK)

        r2 = p.add_run()
        r2.text = str(definition)
        _font(r2, 24, color=C_TEXT_DARK)

    _notes(slide, "Review these key terms with students. Ask them to define each term in their own words before moving on.")
