"""
dok.colors
~~~~~~~~~~
Color resolution: converts Dok color values to 6-digit uppercase hex.

Used by the converter when emitting fill/stroke/color props into DOCX XML.
"""

from __future__ import annotations
import re


# Named colors → 6-digit hex (no #)
NAMED: dict[str, str] = {
    "red":         "FF0000",
    "orange":      "FFA500",
    "yellow":      "FFFF00",
    "green":       "008000",
    "blue":        "0000FF",
    "navy":        "1F3864",
    "purple":      "7030A0",
    "gray":        "808080",
    "grey":        "808080",
    "black":       "000000",
    "white":       "FFFFFF",
    "gold":        "FFC000",
    "silver":      "C0C0C0",
    "lightblue":   "BDD7EE",
    "lightgreen":  "E2EFDA",
    "lightyellow": "FFFF99",
    "lightgray":   "F2F2F2",
    "lightgrey":   "F2F2F2",
    "lightpink":   "FFD7D7",
}

# DOCX only supports these 16 names for <w:highlight>
HIGHLIGHT_NAMES = {
    "yellow", "green", "cyan", "magenta", "blue", "red",
    "darkBlue", "darkCyan", "darkGreen", "darkYellow",
    "darkMagenta", "darkRed", "darkGray", "lightGray",
    "black", "white",
}


def resolve(value: str) -> str | None:
    """
    Convert any Dok color value to 6-digit uppercase hex (no #).

    Accepts:
      "navy"       → "1F3864"
      "#4472C4"    → "4472C4"
      "#ABC"       → "AABBCC"
      "none"       → None  (transparent / no fill)

    Returns None if the value is "none" or cannot be parsed.
    """
    if not value:
        return None

    v = value.strip().lower()

    if v == "none":
        return None

    # Already a named color
    if v in NAMED:
        return NAMED[v]

    # #rrggbb
    m = re.match(r"^#?([0-9a-f]{6})$", v)
    if m:
        return m.group(1).upper()

    # #rgb → #rrggbb
    m = re.match(r"^#?([0-9a-f]{3})$", v)
    if m:
        r, g, b = m.group(1)
        return (r*2 + g*2 + b*2).upper()

    return None


def nearest_highlight(hex6: str) -> str:
    """
    Map a 6-digit hex color to the nearest DOCX highlight name.
    DOCX only supports 16 named highlight colors.

    Used when the 'highlight' prop is given an arbitrary hex.
    """
    # Direct name lookup first
    for name, h in NAMED.items():
        if h == hex6.upper() and name in HIGHLIGHT_NAMES:
            return name

    # Nearest by Euclidean distance in RGB space
    r1 = int(hex6[0:2], 16)
    g1 = int(hex6[2:4], 16)
    b1 = int(hex6[4:6], 16)

    highlight_hex = {
        "yellow":      "FFFF00",
        "green":       "00FF00",
        "cyan":        "00FFFF",
        "magenta":     "FF00FF",
        "blue":        "0000FF",
        "red":         "FF0000",
        "darkBlue":    "00008B",
        "darkCyan":    "008B8B",
        "darkGreen":   "006400",
        "darkYellow":  "808000",
        "darkMagenta": "8B008B",
        "darkRed":     "8B0000",
        "darkGray":    "A9A9A9",
        "lightGray":   "D3D3D3",
        "black":       "000000",
        "white":       "FFFFFF",
    }

    best, best_dist = "yellow", float("inf")
    for name, h in highlight_hex.items():
        r2, g2, b2 = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        dist = (r1-r2)**2 + (g1-g2)**2 + (b1-b2)**2
        if dist < best_dist:
            best_dist, best = dist, name

    return best