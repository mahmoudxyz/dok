"""
dok.models
~~~~~~~~~~
All model dataclasses produced by the Converter and consumed by writers.

These are the intermediate representation between the AST (nodes) and
the final output (DOCX XML or HTML). Each element type has a model.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path


# ---------------------------------------------------------------------------
# Run & Paragraph — the core building blocks
# ---------------------------------------------------------------------------

@dataclass
class RunModel:
    text:      str
    bold:      bool       = False
    italic:    bool       = False
    underline: bool       = False
    strike:    bool       = False
    sup:       bool       = False
    sub:       bool       = False
    color:     str | None = None
    highlight: str | None = None
    size_pt:   int | None = None
    font:      str | None = None
    rtl:       bool       = False
    shading:   str | None = None
    hyperlink_url: str | None = None
    field:     str | None = None   # "PAGE", "NUMPAGES"

    def style_key(self) -> tuple:
        """All style fields as a tuple — used for run merging."""
        return (self.bold, self.italic, self.underline, self.strike,
                self.sup, self.sub, self.color, self.highlight,
                self.size_pt, self.font, self.rtl, self.shading,
                self.hyperlink_url, self.field)


@dataclass
class ParagraphModel:
    runs:         list[RunModel] = field(default_factory=list)
    style:        str            = "Normal"
    align:        str            = "left"
    direction:    str            = "ltr"
    indent_twips: int            = 0
    space_before: int            = 0
    space_after:  int            = 160
    line_spacing: int            = 0      # 0 = inherit from style, >0 = twips (240=single, 360=1.5)
    shading:      str | None     = None
    border_left:  str | None     = None
    border_left_sz: int          = 0
    num_id:       int            = 0      # list numbering ID (0 = none)
    num_ilvl:     int            = 0      # list nesting level


# ---------------------------------------------------------------------------
# Visual elements — print-friendly (native Word formatting)
# ---------------------------------------------------------------------------

@dataclass
class LineModel:
    color: str = "BFBFBF"
    style: str = "single"
    thick: bool = False
    space_before: int = 80
    space_after:  int = 80


@dataclass
class BoxModel:
    content:  list              = field(default_factory=list)
    fill:     str | None        = None
    stroke:   str | None        = "BFBFBF"
    rounded:  bool              = False
    shadow:   bool              = False
    accent:   str | None        = None      # left border color (banner/callout)
    inline:   bool              = False     # inline badge behavior
    text:     str | None        = None      # shorthand text for inline badge
    color:    str | None        = None      # text color override
    align:    str               = "left"    # alignment for inline badge
    width_pct: int              = 0         # 0 = full width, 1-100 = % of content area
    height_pt: int              = 0         # 0 = auto, >0 = fixed height in points


# Backward compat aliases
BannerModel = BoxModel
BadgeModel = BoxModel


# ---------------------------------------------------------------------------
# Drawing shapes — circle, diamond, chevron (OOXML drawing)
# ---------------------------------------------------------------------------

@dataclass
class ShapeModel:
    preset:       str
    fill:         str | None
    stroke:       str | None
    stroke_style: str                   = "solid"
    stroke_thick: bool                  = False
    color:        str | None            = None
    rounded:      bool                  = False
    shadow:       bool                  = False
    inline:       bool                  = True
    float_side:   str | None            = None
    paragraphs:   list[ParagraphModel]  = field(default_factory=list)


@dataclass
class RowModel:
    items:  list                = field(default_factory=list)   # ShapeModel, BoxModel, any
    arrows: list[str | None]    = field(default_factory=list)


# ---------------------------------------------------------------------------
# Layout tables (cols/col)
# ---------------------------------------------------------------------------

@dataclass
class TableModel:
    rows:   list["TableRowModel"] = field(default_factory=list)
    border: bool                  = False


@dataclass
class TableRowModel:
    cells: list["TableCellModel"] = field(default_factory=list)


@dataclass
class TableCellModel:
    content:   list = field(default_factory=list)
    width_pct: int  = 50


# ---------------------------------------------------------------------------
# Data tables (table/tr/td/th)
# ---------------------------------------------------------------------------

@dataclass
class DataTableModel:
    rows:    list["DataTableRowModel"] = field(default_factory=list)
    border:  bool = True
    striped: bool = False


@dataclass
class DataTableRowModel:
    cells:     list["DataTableCellModel"] = field(default_factory=list)
    is_header: bool = False


@dataclass
class DataTableCellModel:
    content: list = field(default_factory=list)
    is_th:   bool = False
    colspan: int  = 1


# ---------------------------------------------------------------------------
# Inline / meta elements
# ---------------------------------------------------------------------------

@dataclass
class ImageModel:
    src:        str
    width_emu:  int
    height_emu: int
    align:      str = "left"


@dataclass
class SpacerModel:
    height_twips: int = 200


@dataclass
class HeaderModel:
    paragraphs: list[ParagraphModel] = field(default_factory=list)


@dataclass
class FooterModel:
    paragraphs: list[ParagraphModel] = field(default_factory=list)


@dataclass
class PageBreakModel:
    pass


# ---------------------------------------------------------------------------
# Section & Document
# ---------------------------------------------------------------------------

@dataclass
class SectionModel:
    margin: str = "normal"
    paper:  str = "a4"
    cols:   int = 1


@dataclass
class DocxModel:
    content:         list              = field(default_factory=list)
    sections:        list[SectionModel] = field(default_factory=list)
    default_font:    str               = "Calibri"
    default_size_pt: int               = 11
    spacing:         str               = "normal"   # compact | tight | normal | relaxed
    header:          HeaderModel | None = None
    footer:          FooterModel | None = None
    base_dir:        Path | None       = None
    has_lists:       bool              = False

    def current_section(self) -> SectionModel:
        """Return the most recent section, or a default if none exist."""
        return self.sections[-1] if self.sections else SectionModel()


# ---------------------------------------------------------------------------
# Spacing presets
# ---------------------------------------------------------------------------

# (para_after_twips, heading_before_scale, line_spacing)
# line_spacing: 240 = single, 276 = 1.15, 360 = 1.5
SPACING_PRESETS: dict[str, tuple[int, float, int]] = {
    "compact": (0,   0.4, 240),
    "tight":   (60,  0.6, 240),
    "normal":  (160, 1.0, 276),
    "relaxed": (200, 1.2, 312),
}
