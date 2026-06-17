"""
JoVE style guide constants.
Generated from the locked presentation guidelines.
"""

from pptx.util import Inches
from pptx.dml.color import RGBColor

SLIDE_W_IN = 20.0
SLIDE_H_IN = 11.25

SLIDE_W = Inches(SLIDE_W_IN)
SLIDE_H = Inches(SLIDE_H_IN)

MARGIN = Inches(0.75)
RIGHT_X = Inches(11.0)
RIGHT_W = Inches(8.25)
CONTENT_TOP = Inches(1.35)
CONTENT_BOTTOM = Inches(10.35)
CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP

FONT_PRIMARY = "Roboto"
FONT_FALLBACK_1 = "Helvetica Neue"
FONT_FALLBACK_2 = "Arial"

PRIMARY_DARK = RGBColor(0x24, 0x29, 0x2F)
BRAND_BLUE = RGBColor(0x6D, 0x9E, 0xEB)
BRAND_BLUE_LIGHT = RGBColor(0xA4, 0xC2, 0xF4)
BRAND_BLUE_PALE = RGBColor(0xC9, 0xDA, 0xF8)
BRAND_BLUE_DARK = RGBColor(0x3C, 0x78, 0xD8)
INTERACTIVE_BLUE = RGBColor(0x50, 0x90, 0xEE)
BLACK = RGBColor(0x00, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MID_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
DARK_GRAY = RGBColor(0x4B, 0x55, 0x63)
LIGHT_GRAY = RGBColor(0x85, 0x85, 0x85)
ACCENT_ORANGE = RGBColor(0xFF, 0x99, 0x00)
ACCENT_RED = RGBColor(0xFF, 0x46, 0x5E)
OFF_WHITE = RGBColor(0xE5, 0xE7, 0xEB)

FOOTER_TEXT = "Copyright © 2026 MyJoVE Corporation. All rights reserved"
