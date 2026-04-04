"""
dok.units
~~~~~~~~~
Production-ready unit parsing and conversion.

All measurements in Dok documents go through this module.
Users can specify values with unit suffixes:
    pt   — points (1/72 inch) — DEFAULT for most properties
    cm   — centimeters
    mm   — millimeters
    in   — inches
    px   — pixels (at 96 DPI)
    emu  — English Metric Units (raw OOXML)
    twip — twips (1/20 point, raw OOXML)

If no unit is given, the default depends on the property context:
    Font sizes:  pt
    Dimensions:  pt (width, height, gaps, padding, margins)
    Spacers:     pt

Conversion targets:
    twips   — for paragraph spacing, margins, table widths, positions
    EMU     — for images and drawing shapes
    half-pt — for font sizes in OOXML (1 half-pt = 0.5 pt)
"""

from __future__ import annotations
import re

# ---------------------------------------------------------------------------
# Core constants
# ---------------------------------------------------------------------------

# Twips per unit
_TWIPS_PER = {
    "pt":   20,          # 1 pt = 20 twips
    "cm":   566.929,     # 1 cm = 566.929 twips (1440 * cm/inch)
    "mm":   56.6929,     # 1 mm = 56.6929 twips
    "in":   1440,        # 1 in = 1440 twips
    "px":   15,          # 1 px = 15 twips (at 96 DPI: 1440/96)
    "emu":  1 / 635,     # 1 EMU = 1/635 twips
    "twip": 1,           # identity
}

# EMU per unit
_EMU_PER = {
    "pt":   12700,       # 1 pt = 12700 EMU
    "cm":   360000,      # 1 cm = 360000 EMU
    "mm":   36000,       # 1 mm = 36000 EMU
    "in":   914400,      # 1 in = 914400 EMU
    "px":   9525,        # 1 px = 9525 EMU (at 96 DPI)
    "emu":  1,           # identity
    "twip": 635,         # 1 twip = 635 EMU
}

# Regex for parsing a value with optional unit suffix
_VALUE_RE = re.compile(
    r'^(-?\d+(?:\.\d+)?)\s*(pt|cm|mm|in|px|emu|twip)?$',
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_to_twips(value, default_unit: str = "pt") -> int:
    """Parse a value (int, float, or string with unit) and return twips.

    Examples:
        parse_to_twips(12)         → 240    (12pt = 240 twips)
        parse_to_twips("12pt")     → 240
        parse_to_twips("2.54cm")   → 1440   (1 inch)
        parse_to_twips("1in")      → 1440
        parse_to_twips("96px")     → 1440
        parse_to_twips("25.4mm")   → 1440
    """
    num, unit = _parse_value(value, default_unit)
    factor = _TWIPS_PER.get(unit)
    if factor is None:
        raise ValueError(f"Unknown unit: {unit!r}")
    return round(num * factor)


def parse_to_emu(value, default_unit: str = "pt") -> int:
    """Parse a value and return EMU (for images and drawing shapes).

    Examples:
        parse_to_emu(72)          → 914400  (72pt = 1 inch)
        parse_to_emu("2.54cm")    → 914400
        parse_to_emu("100px")     → 952500
    """
    num, unit = _parse_value(value, default_unit)
    factor = _EMU_PER.get(unit)
    if factor is None:
        raise ValueError(f"Unknown unit: {unit!r}")
    return round(num * factor)


def parse_to_half_points(value, default_unit: str = "pt") -> int:
    """Parse a value and return half-points (for OOXML font sizes).

    OOXML uses half-points: 22 = 11pt.

    Examples:
        parse_to_half_points(11)      → 22
        parse_to_half_points("14pt")  → 28
    """
    num, unit = _parse_value(value, default_unit)
    # Convert to points first, then to half-points
    twips = num * _TWIPS_PER.get(unit, 20)
    pt = twips / 20.0
    return round(pt * 2)


def parse_to_pt(value, default_unit: str = "pt") -> float:
    """Parse a value and return points.

    Examples:
        parse_to_pt(12)           → 12.0
        parse_to_pt("1in")        → 72.0
        parse_to_pt("2.54cm")     → 72.0
    """
    num, unit = _parse_value(value, default_unit)
    twips = num * _TWIPS_PER.get(unit, 20)
    return round(twips / 20.0, 2)


def pt_to_twips(pt: float) -> int:
    """Convert points to twips."""
    return round(pt * 20)


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU."""
    return round(pt * 12700)


def twips_to_pt(twips: int) -> float:
    """Convert twips to points."""
    return round(twips / 20.0, 2)


def twips_to_emu(twips: int) -> int:
    """Convert twips to EMU."""
    return round(twips * 635)


def emu_to_twips(emu: int) -> int:
    """Convert EMU to twips."""
    return round(emu / 635)


# ---------------------------------------------------------------------------
# Internal
# ---------------------------------------------------------------------------

def _parse_value(value, default_unit: str) -> tuple[float, str]:
    """Parse a value into (number, unit) tuple."""
    if isinstance(value, (int, float)):
        return (float(value), default_unit)

    s = str(value).strip()
    m = _VALUE_RE.match(s)
    if m:
        num = float(m.group(1))
        unit = (m.group(2) or default_unit).lower()
        return (num, unit)

    # Try plain numeric
    try:
        return (float(s), default_unit)
    except ValueError:
        raise ValueError(f"Cannot parse measurement: {value!r}")
