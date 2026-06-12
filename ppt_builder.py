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

SLIDE_W = Inches(20)
SLIDE_H = Inches(11.25)

C_TEXT_DARK = RGBColor(0x24, 0x29, 0x2F)
C_SUBTITLE = RGBColor(0x85, 0x85, 0x85)
C_COPYRIGHT = RGBColor(0xCC, 0xCC, 0xCC)
C_ACCENT_BLUE = RGBColor(0x4A, 0x86, 0xE8)
C_TABLE_HEADER = RGBColor(0x50, 0x90, 0xEE)
C_TABLE_ROW = RGBColor(0xC9, 0xDA, 0xF8)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT_PANEL = RGBColor(0xF3, 0xF6, 0xFB)

FONT = "Helvetica Neue"

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
    _tb(slide, CPY_L, CPY_T, CPY_W, CPY_H,
        "Copyright © 2026 JoVE", 14, color=C_COPYRIGHT, align=PP_ALIGN.CENTER)


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


def _normalize_table(headers, rows):
    headers = [str(h) for h in (headers or [])]
    rows = rows or []

    if not headers:
        headers = ["Term/Step/Type", "Definition/Meaning", "Example/Application"]

    # Feedback fix: keyterm/stage/type tables should have definition + example columns.
    lower = [h.lower() for h in headers]
    has_definition = any("definition" in h or "meaning" in h for h in lower)
    has_example = any("example" in h or "application" in h for h in lower)

    if len(headers) < 3 or not has_definition or not has_example:
        headers = ["Term/Step/Type", "Definition/Meaning", "Example/Application"]

    n_cols = len(headers)
    normalized_rows = []
    for row in rows:
        row = [str(x) for x in list(row)]
        if len(row) < n_cols:
            row += [""] * (n_cols - len(row))
        normalized_rows.append(row[:n_cols])

    if not normalized_rows:
        normalized_rows = [["", "", ""][:n_cols]]

    return headers, normalized_rows


def _add_table(slide, headers, rows, left, top, width):
    headers, rows = _normalize_table(headers, rows)

    n_rows = len(rows) + 1
    n_cols = len(headers)
    row_h = Inches(0.66 if n_rows > 5 else 0.75)
    height = min(Inches(6.8), row_h * n_rows)

    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table

    for ci in range(n_cols):
        tbl.columns[ci].width = int(width / n_cols)

    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_TABLE_HEADER
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = h
        _font(run, 18 if n_cols >= 3 else 21, bold=True, color=C_WHITE)

    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_TABLE_ROW
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = str(val)
            _font(run, 16 if n_cols >= 3 else 19, bold=(ci == 0), color=C_TEXT_DARK)


def create_presentation(logo_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def build_cover_slide(prs, chapter_name, chapter_number, logo_path, cover_image_path=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

    _tb(slide, LEFT, Inches(2.0), Inches(7.8), Inches(5.0),
        chapter_name, 68, bold=True, color=C_TEXT_DARK, wrap=True)
    _tb(slide, LEFT, Inches(7.0), Inches(7.0), Inches(0.8),
        f"Chapter {chapter_number}", 32, bold=True, color=C_SUBTITLE)
    _tb(slide, LEFT, Inches(9.5), Inches(5.0), Inches(0.8),
        "Lecture Slides", 32, bold=True, color=C_SUBTITLE)
    _add_cover_panel(slide, cover_image_path)
    _notes(slide, "Welcome students. Introduce the chapter topic and outline the key lessons they will cover today.")


def build_concept_slide(prs, lesson_name, body_text, sub_label=None,
                        image_path=None, speaker_notes=None, logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

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
                      image_path=None, speaker_notes=None, logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

    _tb(slide, LEFT, Inches(0.38), TEXT_W, Inches(1.25),
        lesson_name, 34, bold=True, color=C_TEXT_DARK)

    if sub_title:
        _tb(slide, LEFT, Inches(1.68), TEXT_W, Inches(0.75),
            sub_title, 23, bold=True, color=C_TEXT_DARK)
        table_top = Inches(2.55)
    else:
        table_top = Inches(2.05)

    _add_table(slide, headers, rows, LEFT, table_top, Inches(7.25))
    _image(slide, image_path)
    _notes(slide, speaker_notes)


def build_discussion_question_slide(prs, lesson_name, question_text,
                                     hint_text=None, image_path=None,
                                     speaker_notes=None, logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

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
                                   speaker_notes=None, logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

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
                         table_rows=None, logo_path="", speaker_notes=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

    _tb(slide, LEFT, Inches(0.38), Inches(18.5), Inches(0.5),
        "SUMMARY", 22, bold=True, color=C_TEXT_DARK)

    _tb(slide, LEFT, Inches(0.9), Inches(18.0), Inches(1.8),
        summary_statement, 36, bold=True, color=C_TEXT_DARK)

    if table_headers and table_rows:
        _add_table(slide, table_headers, table_rows, LEFT, Inches(3.05), Inches(17.9))

    _notes(slide, speaker_notes)


def build_glossary_slide(prs, terms_dict, logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

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
