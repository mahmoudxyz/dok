"""
dok.context
~~~~~~~~~~~
ParaCtx and RunCtx — the two context structs the converter
carries as it walks the node tree.

Design rule:
  - Wrapper nodes (layout/style) call ctx.with_*() to produce
    a new context for their subtree. The original is unchanged.
  - Leaf nodes (content) consume the context to produce DOCX XML.
  - No globals. No mutation. Just immutable updates.
"""

from __future__ import annotations
from dataclasses import dataclass, replace
from typing import Literal


Align     = Literal["left", "center", "right", "justify"]
Direction = Literal["ltr", "rtl"]


# ---------------------------------------------------------------------------
# Paragraph context — affects the <w:pPr> element
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ParaCtx:
    """
    Context inherited by paragraph-level elements.
    Frozen so every update produces a new instance.
    """
    align:     Align     = "left"
    direction: Direction = "ltr"
    indent:    int       = 0        # levels, each = 720 twips (0.5 inch)
    style:     str       = "Normal" # DOCX style name

    # Spacing in twips (1pt = 20 twips)
    space_before: int = 0
    space_after:  int = 160   # default Word paragraph spacing

    def with_align(self, align: Align) -> "ParaCtx":
        return replace(self, align=align)

    def with_direction(self, direction: Direction) -> "ParaCtx":
        return replace(self, direction=direction)

    def with_indent(self, levels: int) -> "ParaCtx":
        return replace(self, indent=self.indent + levels)

    def with_style(self, style: str) -> "ParaCtx":
        return replace(self, style=style)

    def as_heading(self, level: int) -> "ParaCtx":
        """Return a context for a heading paragraph."""
        return replace(self, style=f"Heading{level}", space_before=240, space_after=120)

    def as_quote(self) -> "ParaCtx":
        return replace(self, style="BlockText", indent=self.indent + 1)

    def as_code(self) -> "ParaCtx":
        return replace(self, style="SourceCode", space_before=0, space_after=0)

    def indent_twips(self) -> int:
        """Convert indent levels to twips for DOCX <w:ind>."""
        return self.indent * 720

    def jc_val(self) -> str:
        """DOCX <w:jc w:val="..."> value for this alignment."""
        return {"left": "left", "center": "center",
                "right": "right", "justify": "both"}[self.align]


# ---------------------------------------------------------------------------
# Run context — affects the <w:rPr> element
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class RunCtx:
    """
    Context inherited by run-level (inline) elements.
    Frozen so every update produces a new instance.
    """
    bold:      bool       = False
    italic:    bool       = False
    underline: bool       = False
    strike:    bool       = False
    sup:       bool       = False
    sub:       bool       = False

    color:     str | None = None    # 6-digit hex, no #, e.g. "FF0000"
    highlight: str | None = None    # DOCX highlight name e.g. "yellow"
    size_pt:   int | None = None    # points; None = inherit from doc default
    font:      str | None = None    # None = inherit from doc default

    def with_bold(self)      -> "RunCtx": return replace(self, bold=True)
    def with_italic(self)    -> "RunCtx": return replace(self, italic=True)
    def with_underline(self) -> "RunCtx": return replace(self, underline=True)
    def with_strike(self)    -> "RunCtx": return replace(self, strike=True)
    def with_sup(self)       -> "RunCtx": return replace(self, sup=True, sub=False)
    def with_sub(self)       -> "RunCtx": return replace(self, sub=True, sup=False)

    def with_color(self, hex_color: str) -> "RunCtx":
        return replace(self, color=hex_color.lstrip("#").upper())

    def with_highlight(self, name: str) -> "RunCtx":
        return replace(self, highlight=name)

    def with_size(self, pt: int) -> "RunCtx":
        return replace(self, size_pt=pt)

    def with_font(self, name: str) -> "RunCtx":
        return replace(self, font=name)

    def sz_val(self, doc_default_pt: int = 11) -> int:
        """DOCX <w:sz w:val="..."> — half-points."""
        return (self.size_pt or doc_default_pt) * 2

    def vert_align(self) -> str | None:
        """DOCX <w:vertAlign w:val="..."> or None."""
        if self.sup: return "superscript"
        if self.sub: return "subscript"
        return None