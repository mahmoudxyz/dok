"""
dok.constants
~~~~~~~~~~~~~
Shared unit conversions, margin/paper presets, and common values
used by converter and docx_writer.
"""

from __future__ import annotations


# ---------------------------------------------------------------------------
# Unit conversions
# ---------------------------------------------------------------------------

INCH_TO_EMU  = 914400
INCH_TO_TWIP = 1440

def inch_to_emu(i: float) -> int:   return int(i * INCH_TO_EMU)
def inch_to_twip(i: float) -> int:  return int(i * INCH_TO_TWIP)
def twip_to_px(t: int) -> float:    return round(t / 14.4, 2)    # 96 dpi
def twip_to_pt(t: int) -> float:    return round(t / 20, 2)      # 1 twip = 1/20 pt
def emu_to_px(e: int) -> float:     return round(e / 9144, 2)    # 96 dpi


# ---------------------------------------------------------------------------
# Margin presets (twips)
# ---------------------------------------------------------------------------

MARGIN_PRESETS = {
    "normal": {"top": 1440, "right": 1440, "bottom": 1440, "left": 1440},
    "narrow": {"top":  720, "right":  720, "bottom":  720, "left":  720},
    "wide":   {"top": 1440, "right": 2880, "bottom": 1440, "left": 2880},
    "none":   {"top":    0, "right":    0, "bottom":    0, "left":    0},
}


# ---------------------------------------------------------------------------
# Paper sizes (twips: width, height)
# ---------------------------------------------------------------------------

PAPER_SIZES = {
    "a4":     (11906, 16838),
    "letter": (12240, 15840),
    "a3":     (16838, 23811),
}


def content_width_twip(paper: str = "a4", margin: str = "normal",
                       *, margins: dict[str, int] | None = None) -> int:
    """Usable content width after subtracting left+right margins.

    If *margins* dict is given, uses those exact values instead of preset lookup.
    """
    pw, _ = PAPER_SIZES.get(paper, PAPER_SIZES["a4"])
    if margins is not None:
        m = margins
    else:
        m = MARGIN_PRESETS.get(margin, MARGIN_PRESETS["normal"])
    return pw - m["left"] - m["right"]


# ---------------------------------------------------------------------------
# Common colors
# ---------------------------------------------------------------------------

DEFAULT_BORDER_COLOR = "BFBFBF"
HYPERLINK_COLOR = "0563C1"
