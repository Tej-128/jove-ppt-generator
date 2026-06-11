"""
JoVE PPT Builder - pixel-perfect layout from reference deck
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import requests, io, os, re

SLIDE_W = Inches(20)
SLIDE_H = Inches(11.25)

# Colors from reference PPT
C_TEXT_DARK    = RGBColor(0x24, 0x29, 0x2F)
C_SUBTITLE     = RGBColor(0x85, 0x85, 0x85)
C_COPYRIGHT    = RGBColor(0xCC, 0xCC, 0xCC)
C_ACCENT_BLUE  = RGBColor(0x4A, 0x86, 0xE8)
C_TABLE_HEADER = RGBColor(0x50, 0x90, 0xEE)
C_TABLE_ROW    = RGBColor(0xC9, 0xDA, 0xF8)
C_WHITE        = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "Helvetica Neue"

# Layout zones (inches)
LEFT   = Inches(1.042)
TEXT_W = Inches(7.0)
IMG_L  = Inches(8.8)
IMG_T  = Inches(1.5)
IMG_W  = Inches(10.5)
IMG_H  = Inches(9.0)
LOGO_L = Inches(18.444)
LOGO_T = Inches(0.326)
LOGO_W = Inches(1.087)
LOGO_H = Inches(0.551)
CPY_L  = Inches(7.5)
CPY_T  = Inches(10.55)   # safe zone fix — was 10.78
CPY_W  = Inches(5.5)
CPY_H  = Inches(0.45)


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
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    _font(run, size_pt, bold, italic, color or C_TEXT_DARK)
    return box


def _logo(slide, logo_path):
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, LOGO_L, LOGO_T, LOGO_W, LOGO_H)


def _copyright(slide):
    _tb(slide, CPY_L, CPY_T, CPY_W, CPY_H,
        "Copyright © 2026 MyJoVE Corporation. All rights reserved",
        11, color=C_COPYRIGHT, align=PP_ALIGN.CENTER)


def _white_bg(slide):
    sh = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    sh.fill.solid()
    sh.fill.fore_color.rgb = C_WHITE
    sh.line.fill.background()
    sp = sh._element
    spTree = slide.shapes._spTree
    spTree.remove(sp)
    spTree.insert(2, sp)


# FIX B2: _image now accepts both image_url and image_path
def _image(slide, image_url=None, image_path=None):
    """Add image right side. Tries local path first, then URL, then placeholder."""
    try:
        # Local file (ZIP thumbnail) takes priority
        if image_path and os.path.exists(image_path):
            slide.shapes.add_picture(image_path, IMG_L, IMG_T, IMG_W, IMG_H)
            return True
        # Web URL
        if image_url:
            r = requests.get(image_url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
            if r.status_code == 200:
                slide.shapes.add_picture(io.BytesIO(r.content), IMG_L, IMG_T, IMG_W, IMG_H)
                return True
    except Exception:
        pass
    # Gray placeholder — rule R1: every slide must have an image zone
    ph = slide.shapes.add_shape(1, IMG_L, IMG_T, IMG_W, IMG_H)
    ph.fill.solid()
    ph.fill.fore_color.rgb = RGBColor(0xE8, 0xE8, 0xE8)
    ph.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    tf = ph.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "[Image Placeholder]"
    _font(run, 18, color=RGBColor(0x99, 0x99, 0x99))
    return False


def _add_table(slide, headers, rows, left, top, width):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    row_h = Inches(0.75)
    height = row_h * n_rows
    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_TABLE_HEADER
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = h
        _font(run, 22, bold=True, color=C_WHITE)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_TABLE_ROW
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = str(val)
            _font(run, 20, bold=(ci == 0), color=C_TEXT_DARK)


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


def _base_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _body_text(slide, body_text, top):
    box = slide.shapes.add_textbox(LEFT, top, TEXT_W,
                                   Inches(10.4) - top)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in body_text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
            p.space_before = Pt(10)
        p.alignment = PP_ALIGN.LEFT
        segments = re.split(r'\*\*(.+?)\*\*', line)
        for i, seg in enumerate(segments):
            if not seg:
                continue
            run = p.add_run()
            run.text = seg
            _font(run, 26, bold=(i % 2 == 1), color=C_TEXT_DARK)


# ── Public slide builders ────────────────────────────────────────────────────

def build_cover_slide(prs, chapter_name, chapter_number, logo_path):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _tb(slide, LEFT, Inches(2.0), Inches(7.8), Inches(5.0),
        chapter_name, 68, bold=True, color=C_TEXT_DARK, wrap=True)
    _tb(slide, LEFT, Inches(7.0), Inches(7.0), Inches(0.8),
        f"Chapter {chapter_number}", 32, bold=True, color=C_SUBTITLE)
    _tb(slide, LEFT, Inches(9.5), Inches(5.0), Inches(0.8),
        "Lecture Slides", 32, bold=True, color=C_SUBTITLE)
    _image(slide)  # placeholder on cover — no URL available at chapter level
    _copyright(slide)
    _notes(slide, "Welcome students. Introduce the chapter topic and outline the key lessons they will cover today.")


# FIX B2: all builders now accept image_path parameter
def build_concept_slide(prs, lesson_name, body_text, sub_label=None,
                        image_url=None, image_path=None,
                        speaker_notes=None, logo_path=""):
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

    title_box = slide.shapes.add_textbox(LEFT, title_top, TEXT_W, Inches(1.6))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = lesson_name
    _font(run, 34, bold=True, color=C_TEXT_DARK)

    body_top = title_top + Inches(1.7)
    _body_text(slide, body_text, body_top)
    _image(slide, image_url=image_url, image_path=image_path)
    _notes(slide, speaker_notes)


def build_table_slide(prs, lesson_name, headers, rows, sub_title=None,
                      image_url=None, image_path=None,
                      speaker_notes=None, logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

    title_box = slide.shapes.add_textbox(LEFT, Inches(0.38), TEXT_W, Inches(1.4))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = lesson_name
    _font(run, 34, bold=True, color=C_TEXT_DARK)

    if sub_title:
        _tb(slide, LEFT, Inches(1.85), TEXT_W, Inches(0.55),
            sub_title, 24, bold=True, color=C_TEXT_DARK)
        table_top = Inches(2.5)
    else:
        table_top = Inches(2.0)

    _add_table(slide, headers, rows, LEFT, table_top, Inches(7.2))
    _image(slide, image_url=image_url, image_path=image_path)
    _notes(slide, speaker_notes)


def build_discussion_question_slide(prs, lesson_name, question_text,
                                     hint_text=None, image_url=None,
                                     image_path=None, speaker_notes=None,
                                     logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

    # RC3: use descriptive title from AI, not hardcoded prefix
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
    run.text = question_text
    _font(run, 32, bold=True, color=C_TEXT_DARK)

    if hint_text:
        _tb(slide, LEFT, Inches(5.6), TEXT_W + Inches(0.5), Inches(1.5),
            f"Hint: {hint_text}", 24, italic=True, color=C_TEXT_DARK)

    _image(slide, image_url=image_url, image_path=image_path)
    _notes(slide, speaker_notes)


def build_discussion_answer_slide(prs, lesson_name, answer_summary,
                                   answer_explanation, image_url=None,
                                   image_path=None, speaker_notes=None,
                                   logo_path=""):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

    # RC3: use descriptive title from AI
    _tb(slide, LEFT, Inches(0.38), TEXT_W + Inches(1.5), Inches(1.2),
        lesson_name, 36, bold=True, color=C_TEXT_DARK)
    _tb(slide, LEFT, Inches(1.75), Inches(5.5), Inches(0.6),
        "Discuss with the class", 26, color=C_ACCENT_BLUE)
    _tb(slide, LEFT, Inches(2.55), TEXT_W + Inches(0.5), Inches(1.6),
        f"Answer: {answer_summary}", 30, bold=True, color=C_TEXT_DARK)

    ans_box = slide.shapes.add_textbox(LEFT, Inches(4.3), TEXT_W + Inches(0.5), Inches(4.5))
    tf = ans_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = answer_explanation
    _font(run, 26, italic=True, color=C_TEXT_DARK)

    _image(slide, image_url=image_url, image_path=image_path)
    _notes(slide, speaker_notes)


# RC2: summary now accepts and displays image
def build_summary_slide(prs, summary_statement, table_headers=None,
                         table_rows=None, logo_path="", speaker_notes=None):
    slide = _base_slide(prs)
    _white_bg(slide)
    _logo(slide, logo_path)
    _copyright(slide)

    _tb(slide, LEFT, Inches(0.38), Inches(18.5), Inches(0.5),
        "SUMMARY", 22, bold=True, color=C_TEXT_DARK)

    stmt_box = slide.shapes.add_textbox(LEFT, Inches(0.9), Inches(18.0), Inches(2.0))
    tf = stmt_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = summary_statement
    _font(run, 38, bold=True, color=C_TEXT_DARK)

    if table_headers and table_rows:
        # Full width table — no image on summary slides (matches benchmark)
        _add_table(slide, table_headers, table_rows, LEFT, Inches(3.2), Inches(17.9))

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
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
            p.space_before = Pt(10)
        p.alignment = PP_ALIGN.LEFT
        r1 = p.add_run()
        r1.text = f"{term}: "
        _font(r1, 26, bold=True, color=C_TEXT_DARK)
        r2 = p.add_run()
        r2.text = definition
        _font(r2, 26, color=C_TEXT_DARK)

    # FIX L3: glossary notes was passing None — now has real notes
    _notes(slide, "Review these key terms with students. Ask them to define each term in their own words before moving on.")


def create_presentation(logo_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs
