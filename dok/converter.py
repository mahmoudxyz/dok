"""
dok.converter
~~~~~~~~~~~~~
Walks the node tree and produces a DocxModel.

Print-friendly: box/callout/banner/badge/line use native Word formatting.
Only circle/diamond/chevron use drawing shapes.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from .nodes   import (Node, ElementNode, TextNode, ArrowNode,
                       LAYOUT_NODES, STYLE_NODES, SHAPE_PRESETS,
                       BLOCK_NODES, ALL_KNOWN_NODES)
from .context import ParaCtx, RunCtx
from .colors  import resolve as resolve_color


# ---------------------------------------------------------------------------
# Model objects
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


@dataclass
class ParagraphModel:
    runs:         list[RunModel] = field(default_factory=list)
    style:        str            = "Normal"
    align:        str            = "left"
    direction:    str            = "ltr"
    indent_twips: int            = 0
    space_before: int            = 0
    space_after:  int            = 160
    shading:      str | None     = None
    border_left:  str | None     = None
    border_left_sz: int          = 0
    num_id:       int            = 0      # list numbering ID (0 = none)
    num_ilvl:     int            = 0      # list nesting level


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


@dataclass
class BannerModel:
    paragraphs: list[ParagraphModel] = field(default_factory=list)
    fill:       str | None           = None
    accent:     str | None           = None


@dataclass
class BadgeModel:
    text:  str
    fill:  str | None = None
    color: str | None = None
    align: str        = "left"


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
    shapes: list[ShapeModel]    = field(default_factory=list)
    arrows: list[str | None]    = field(default_factory=list)


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


@dataclass
class DataTableModel:
    """Visible data table with borders."""
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


# Spacing presets: (para_after, heading_before_scale, line_spacing)
# line_spacing is in 240ths of a line (240 = single, 276 = 1.15, 360 = 1.5)
_SPACING_PRESETS: dict[str, tuple[int, float, int]] = {
    "compact": (0,   0.4, 240),   # no para gap, single line, tight headings
    "tight":   (60,  0.6, 240),   # small para gap, single line
    "normal":  (160, 1.0, 276),   # Word default (8pt after, 1.15 line)
    "relaxed": (200, 1.2, 312),   # generous spacing
}


# ---------------------------------------------------------------------------
# Converter
# ---------------------------------------------------------------------------

_INCH_TO_EMU = 914400

class Converter:

    def __init__(self) -> None:
        self._model      = DocxModel()
        self._float_next = None
        self._num_counter = 0   # for unique list numbering IDs

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def convert(self, nodes: list[Node]) -> DocxModel:
        self._model.sections.append(SectionModel())

        if (len(nodes) == 1
                and isinstance(nodes[0], ElementNode)
                and nodes[0].name == "doc"):
            doc = nodes[0]
            self._model.default_font    = str(doc.props.get("font", "Calibri"))
            self._model.default_size_pt = int(doc.props.get("size", 11))
            self._model.spacing         = str(doc.props.get("spacing", "normal"))
            nodes = doc.children

        para = ParaCtx()
        run  = RunCtx()
        for node in nodes:
            self._walk(node, para, run)

        return self._model

    # ------------------------------------------------------------------
    # Walker
    # ------------------------------------------------------------------

    def _walk(self, node: Node, para: ParaCtx, run: RunCtx) -> None:
        if isinstance(node, TextNode):
            p = ParagraphModel(
                runs=[self._make_run(node.text, run, para)],
                style=para.style, align=para.align,
                direction=para.direction,
                indent_twips=para.indent_twips(),
                space_before=para.space_before,
                space_after=para.space_after,
            )
            self._model.content.append(p)
            return

        if isinstance(node, ArrowNode):
            return

        if not isinstance(node, ElementNode):
            return

        name = node.name

        # Document structure
        if name == "doc":
            self._model.default_font    = str(node.props.get("font", "Calibri"))
            self._model.default_size_pt = int(node.props.get("size", 11))
            for child in node.children:
                self._walk(child, para, run)
        elif name == "page":
            self._handle_page(node, para, run)

        # Layout
        elif name == "row":
            self._handle_row(node, para, run)
        elif name == "cols":
            self._handle_cols(node, para, run)
        elif name == "float":
            self._float_next = node.props.get("side", "right")
            for child in node.children:
                self._walk(child, para, run)
            self._float_next = None
        elif name in LAYOUT_NODES:
            self._handle_layout(node, para, run)

        # Style
        elif name in STYLE_NODES:
            self._handle_style(node, para, run)

        # Special
        elif name == "---":
            self._model.content.append(PageBreakModel())

        # Block content
        elif name in BLOCK_NODES:
            self._emit_paragraph(node, para, run)

        # Print-friendly elements
        elif name == "line":    self._emit_line(node)
        elif name == "box":     self._emit_box(node, para, run)
        elif name == "callout": self._emit_callout(node, para, run)
        elif name == "badge":   self._emit_badge(node, para, run)
        elif name == "banner":  self._emit_banner(node, para, run)

        # Lists
        elif name in ("ul", "ol"):
            self._emit_list(node, para, run)

        # Data tables
        elif name == "table":
            self._emit_data_table(node, para, run)

        # Images
        elif name == "img":
            self._emit_image(node, para)

        # Links (inline — usually inside p{})
        elif name == "link":
            self._emit_link_paragraph(node, para, run)

        # Page number (standalone)
        elif name == "page-number":
            self._emit_page_number(para)

        # Spacer
        elif name == "space":
            height = int(node.props.get("size", 10))
            self._model.content.append(SpacerModel(height_twips=height * 20))

        # Header / Footer
        elif name == "header":
            self._emit_header(node, para, run)
        elif name == "footer":
            self._emit_footer(node, para, run)

        # Drawing shapes (circle, diamond, chevron)
        elif name in SHAPE_PRESETS:
            self._emit_drawing_shape(node, para, run)

        else:
            for child in node.children:
                self._walk(child, para, run)

    # ------------------------------------------------------------------
    # Page
    # ------------------------------------------------------------------

    def _handle_page(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        section = SectionModel(
            margin=str(node.props.get("margin", "normal")),
            paper =str(node.props.get("paper",  "a4")),
            cols  =int(node.props.get("cols",   1)),
        )
        if self._model.sections and self._model.sections[-1].margin == "normal":
            self._model.sections[-1] = section
        else:
            self._model.sections.append(section)
        for child in node.children:
            self._walk(child, para, run)

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _handle_layout(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        name = node.name
        if name == "center":    para = para.with_align("center")
        elif name == "right":   para = para.with_align("right")
        elif name == "justify": para = para.with_align("justify")
        elif name == "left":    para = para.with_align("left")
        elif name == "rtl":     para = para.with_direction("rtl")
        elif name == "ltr":     para = para.with_direction("ltr")
        elif name == "indent":
            level = int(node.props.get("level", 1))
            para  = para.with_indent(level)
        for child in node.children:
            self._walk(child, para, run)

    # ------------------------------------------------------------------
    # Style
    # ------------------------------------------------------------------

    def _handle_style(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        name = node.name
        if name == "bold":       run = run.with_bold()
        elif name == "italic":   run = run.with_italic()
        elif name == "underline":run = run.with_underline()
        elif name == "strike":   run = run.with_strike()
        elif name == "sup":      run = run.with_sup()
        elif name == "sub":      run = run.with_sub()
        elif name == "color":
            v = node.props.get("value") or (node.children[0].text
                if node.children and isinstance(node.children[0], TextNode) else None)
            if v:
                h = resolve_color(str(v))
                if h: run = run.with_color(h)
        elif name == "size":
            v = node.props.get("value")
            if v is not None: run = run.with_size(int(v))
        elif name == "font":
            v = node.props.get("value")
            if v: run = run.with_font(str(v))
        elif name == "highlight":
            v = node.props.get("value")
            if v: run = run.with_highlight(str(v))

        if "color" in node.props:
            h = resolve_color(str(node.props["color"]))
            if h: run = run.with_color(h)
        if "size" in node.props:
            run = run.with_size(int(node.props["size"]))
        if "font" in node.props:
            run = run.with_font(str(node.props["font"]))
        if node.props.get("bold"):    run = run.with_bold()
        if node.props.get("italic"):  run = run.with_italic()

        for child in node.children:
            self._walk(child, para, run)

    # ------------------------------------------------------------------
    # Row
    # ------------------------------------------------------------------

    def _handle_row(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        row_model = RowModel()
        for child in node.children:
            if isinstance(child, ArrowNode):
                row_model.arrows.append(child.label)
            elif isinstance(child, ElementNode) and child.name in SHAPE_PRESETS:
                shape = self._build_drawing_shape(child, para, run)
                row_model.shapes.append(shape)
        self._model.content.append(row_model)

    # ------------------------------------------------------------------
    # Cols
    # ------------------------------------------------------------------

    def _handle_cols(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        ratio_str = str(node.props.get("ratio", "1:1"))
        parts     = [int(p) for p in ratio_str.split(":")]
        total     = sum(parts)
        widths    = [round(p / total * 100) for p in parts]

        col_nodes = [c for c in node.children
                    if isinstance(c, ElementNode) and c.name == "col"]

        table = TableModel(border=False)
        row   = TableRowModel()
        for i, col_node in enumerate(col_nodes):
            pct  = widths[i] if i < len(widths) else 100 // len(col_nodes)
            cell = TableCellModel(width_pct=pct)
            sub = Converter()
            sub._model.default_font    = self._model.default_font
            sub._model.default_size_pt = self._model.default_size_pt
            for child in col_node.children:
                sub._walk(child, para, run)
            cell.content = sub._model.content
            row.cells.append(cell)
        table.rows.append(row)
        self._model.content.append(table)

    # ------------------------------------------------------------------
    # Paragraph
    # ------------------------------------------------------------------

    def _emit_paragraph(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        name = node.name
        if name == "h1":    para = para.as_heading(1)
        elif name == "h2":  para = para.as_heading(2)
        elif name == "h3":  para = para.as_heading(3)
        elif name == "h4":  para = para.as_heading(4)
        elif name == "quote": para = para.as_quote()
        elif name == "code":  para = para.as_code()

        if "color" in node.props:
            h = resolve_color(str(node.props["color"]))
            if h: run = run.with_color(h)
        if "size" in node.props:
            run = run.with_size(int(node.props["size"]))

        runs  = self._collect_runs(node.children, run, para.direction)
        runs  = self._merge_runs(runs)

        self._model.content.append(ParagraphModel(
            runs=runs, style=para.style, align=para.align,
            direction=para.direction, indent_twips=para.indent_twips(),
            space_before=para.space_before, space_after=para.space_after,
        ))

    # ------------------------------------------------------------------
    # Line
    # ------------------------------------------------------------------

    def _emit_line(self, node: ElementNode) -> None:
        props = node.props
        stroke = resolve_color(str(props["stroke"])) if "stroke" in props else "BFBFBF"
        style = "dashed" if props.get("dashed") else "single"
        self._model.content.append(LineModel(
            color=stroke or "BFBFBF", style=style,
            thick=bool(props.get("thick", False)),
        ))

    # ------------------------------------------------------------------
    # Box
    # ------------------------------------------------------------------

    def _emit_box(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        fill   = resolve_color(str(props["fill"]))   if "fill"   in props else None
        stroke = resolve_color(str(props["stroke"])) if "stroke" in props else "BFBFBF"
        color  = resolve_color(str(props["color"]))  if "color"  in props else None

        box_run = run
        if color: box_run = box_run.with_color(color)

        sub = Converter()
        sub._model.default_font    = self._model.default_font
        sub._model.default_size_pt = self._model.default_size_pt
        for child in node.children:
            if isinstance(child, TextNode):
                sub._model.content.append(ParagraphModel(
                    runs=[self._make_run(child.text, box_run)], align=para.align,
                ))
            else:
                sub._walk(child, para, box_run)

        self._model.content.append(BoxModel(
            content=sub._model.content, fill=fill, stroke=stroke,
            rounded=bool(props.get("rounded", False)),
            shadow=bool(props.get("shadow", False)),
        ))

    # ------------------------------------------------------------------
    # Callout
    # ------------------------------------------------------------------

    def _emit_callout(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        fill   = resolve_color(str(props["fill"]))   if "fill"   in props else "FFF2CC"
        stroke = resolve_color(str(props["stroke"])) if "stroke" in props else "FFC000"

        sub = Converter()
        sub._model.default_font    = self._model.default_font
        sub._model.default_size_pt = self._model.default_size_pt
        for child in node.children:
            if isinstance(child, TextNode):
                sub._model.content.append(ParagraphModel(
                    runs=[self._make_run(child.text, run)],
                ))
            else:
                sub._walk(child, para, run)

        for item in sub._model.content:
            if isinstance(item, ParagraphModel):
                item.shading = fill
                item.border_left = stroke
                item.border_left_sz = 24
                item.indent_twips = max(item.indent_twips, 180)
            self._model.content.append(item)

    # ------------------------------------------------------------------
    # Badge
    # ------------------------------------------------------------------

    def _emit_badge(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        fill  = resolve_color(str(props["fill"]))  if "fill"  in props else "1F3864"
        color = resolve_color(str(props["color"])) if "color" in props else "FFFFFF"
        text_parts = [c.text for c in node.children if isinstance(c, TextNode)]
        self._model.content.append(BadgeModel(
            text=" ".join(text_parts), fill=fill, color=color, align=para.align,
        ))

    # ------------------------------------------------------------------
    # Banner
    # ------------------------------------------------------------------

    def _emit_banner(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        fill   = resolve_color(str(props["fill"]))   if "fill"   in props else None
        accent = resolve_color(str(props["accent"])) if "accent" in props else None
        color  = resolve_color(str(props["color"]))  if "color"  in props else None

        banner_run = run
        if color: banner_run = banner_run.with_color(color)

        sub = Converter()
        sub._model.default_font    = self._model.default_font
        sub._model.default_size_pt = self._model.default_size_pt
        for child in node.children:
            if isinstance(child, TextNode):
                sub._model.content.append(ParagraphModel(
                    runs=[self._make_run(child.text, banner_run)], align=para.align,
                ))
            else:
                sub._walk(child, para, banner_run)

        paras = []
        for item in sub._model.content:
            if isinstance(item, ParagraphModel):
                item.shading = fill
                if accent:
                    item.border_left = accent
                    item.border_left_sz = 48
                paras.append(item)

        self._model.content.append(BannerModel(
            paragraphs=paras, fill=fill, accent=accent,
        ))

    # ------------------------------------------------------------------
    # Lists
    # ------------------------------------------------------------------

    def _emit_list(self, node: ElementNode, para: ParaCtx, run: RunCtx,
                   ilvl: int = 0) -> None:
        is_ordered = (node.name == "ol")
        # num_id: 1 = bullet, 2 = ordered
        num_id = 2 if is_ordered else 1
        self._model.has_lists = True

        for child in node.children:
            if isinstance(child, ElementNode) and child.name == "li":
                self._emit_list_item(child, para, run, num_id, ilvl)
            elif isinstance(child, ElementNode) and child.name in ("ul", "ol"):
                # Nested list
                self._emit_list(child, para, run, ilvl + 1)

    def _emit_list_item(self, node: ElementNode, para: ParaCtx, run: RunCtx,
                        num_id: int, ilvl: int) -> None:
        runs = self._collect_runs(node.children, run, para.direction)
        runs = self._merge_runs(runs)

        # Check for nested lists among children
        non_list_children = []
        nested_lists = []
        for child in node.children:
            if isinstance(child, ElementNode) and child.name in ("ul", "ol"):
                nested_lists.append(child)
            else:
                non_list_children.append(child)

        if non_list_children:
            runs = self._collect_runs(non_list_children, run, para.direction)
            runs = self._merge_runs(runs)
        else:
            runs = []

        if runs:
            self._model.content.append(ParagraphModel(
                runs=runs, align=para.align, direction=para.direction,
                num_id=num_id, num_ilvl=ilvl,
            ))

        for nested in nested_lists:
            self._emit_list(nested, para, run, ilvl + 1)

    # ------------------------------------------------------------------
    # Data Table
    # ------------------------------------------------------------------

    def _emit_data_table(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        table = DataTableModel(
            border=bool(props.get("border", True)),
            striped=bool(props.get("striped", False)),
        )

        for child in node.children:
            if isinstance(child, ElementNode) and child.name == "tr":
                row = DataTableRowModel()
                has_th = False
                for cell_node in child.children:
                    if isinstance(cell_node, ElementNode) and cell_node.name in ("td", "th"):
                        is_th = (cell_node.name == "th")
                        if is_th: has_th = True
                        colspan = int(cell_node.props.get("colspan", 1))

                        sub = Converter()
                        sub._model.default_font    = self._model.default_font
                        sub._model.default_size_pt = self._model.default_size_pt
                        cell_run = run.with_bold() if is_th else run
                        for cc in cell_node.children:
                            if isinstance(cc, TextNode):
                                sub._model.content.append(ParagraphModel(
                                    runs=[self._make_run(cc.text, cell_run)],
                                    align=para.align,
                                ))
                            else:
                                sub._walk(cc, para, cell_run)

                        row.cells.append(DataTableCellModel(
                            content=sub._model.content, is_th=is_th,
                            colspan=colspan,
                        ))
                row.is_header = has_th
                table.rows.append(row)

        self._model.content.append(table)

    # ------------------------------------------------------------------
    # Image
    # ------------------------------------------------------------------

    def _emit_image(self, node: ElementNode, para: ParaCtx) -> None:
        props = node.props
        src   = str(props.get("src", ""))
        width_in  = int(props.get("width", 4))
        height_in = int(props.get("height", 0))

        if height_in == 0:
            # Try to read dimensions from file
            height_in = self._auto_image_height(src, width_in)

        self._model.content.append(ImageModel(
            src=src,
            width_emu=width_in * _INCH_TO_EMU,
            height_emu=height_in * _INCH_TO_EMU,
            align=para.align,
        ))

    def _auto_image_height(self, src: str, width_in: int) -> int:
        """Try to compute proportional height. Falls back to width."""
        try:
            from .image import image_dimensions
            base = self._model.base_dir
            path = (base / src) if base else Path(src)
            w, h = image_dimensions(path)
            if w > 0 and h > 0:
                return int(width_in * h / w)
        except Exception:
            pass
        return width_in  # square fallback

    # ------------------------------------------------------------------
    # Hyperlink
    # ------------------------------------------------------------------

    def _emit_link_paragraph(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        href = str(node.props.get("href", ""))
        link_run = run.with_color("0563C1").with_underline()

        runs = self._collect_runs(node.children, link_run, para.direction)
        for r in runs:
            r.hyperlink_url = href
        runs = self._merge_runs(runs)

        self._model.content.append(ParagraphModel(
            runs=runs, align=para.align, direction=para.direction,
        ))

    # ------------------------------------------------------------------
    # Page Number
    # ------------------------------------------------------------------

    def _emit_page_number(self, para: ParaCtx) -> None:
        self._model.content.append(ParagraphModel(
            runs=[RunModel(text="", field="PAGE")],
            align=para.align,
        ))

    # ------------------------------------------------------------------
    # Spacer — handled inline in _walk
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # Header / Footer
    # ------------------------------------------------------------------

    def _emit_header(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        sub = Converter()
        sub._model.default_font    = self._model.default_font
        sub._model.default_size_pt = self._model.default_size_pt
        for child in node.children:
            sub._walk(child, para, run)

        paras = [c for c in sub._model.content if isinstance(c, ParagraphModel)]
        self._model.header = HeaderModel(paragraphs=paras)

    def _emit_footer(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        sub = Converter()
        sub._model.default_font    = self._model.default_font
        sub._model.default_size_pt = self._model.default_size_pt
        for child in node.children:
            sub._walk(child, para, run)

        paras = [c for c in sub._model.content if isinstance(c, ParagraphModel)]
        self._model.footer = FooterModel(paragraphs=paras)

    # ------------------------------------------------------------------
    # Drawing shapes (circle, diamond, chevron)
    # ------------------------------------------------------------------

    def _emit_drawing_shape(self, node: ElementNode,
                            para: ParaCtx, run: RunCtx) -> None:
        self._model.content.append(self._build_drawing_shape(node, para, run))

    def _build_drawing_shape(self, node: ElementNode,
                             para: ParaCtx, run: RunCtx) -> ShapeModel:
        props  = node.props
        preset = SHAPE_PRESETS.get(node.name, "rect")
        fill   = resolve_color(str(props["fill"]))   if "fill"   in props else None
        stroke = resolve_color(str(props["stroke"])) if "stroke" in props else "BFBFBF"
        color  = resolve_color(str(props["color"]))  if "color"  in props else None

        stroke_style = "solid"; stroke_thick = False
        if "stroke" in props:
            v = str(props["stroke"])
            if v == "dashed":   stroke = "BFBFBF"; stroke_style = "dashed"
            elif v == "dotted": stroke = "BFBFBF"; stroke_style = "dotted"
            elif v == "thick":  stroke_thick = True
            elif v == "none":   stroke = None

        shape = ShapeModel(
            preset=preset, fill=fill, stroke=stroke,
            stroke_style=stroke_style, stroke_thick=stroke_thick,
            color=color,
            rounded=bool(props.get("rounded", False)),
            shadow=bool(props.get("shadow", False)),
            inline=self._float_next is None,
            float_side=self._float_next,
        )

        if node.children:
            shape_run = run
            if color: shape_run = shape_run.with_color(color)
            sub = Converter()
            sub._model.default_font    = self._model.default_font
            sub._model.default_size_pt = self._model.default_size_pt
            for child in node.children:
                if isinstance(child, TextNode):
                    shape.paragraphs.append(ParagraphModel(
                        runs=[sub._make_run(child.text, shape_run)],
                    ))
                elif isinstance(child, ElementNode):
                    sub._walk(child, para, shape_run)
                    shape.paragraphs.extend(
                        c for c in sub._model.content if isinstance(c, ParagraphModel)
                    )
                    sub._model.content.clear()

        return shape

    # ------------------------------------------------------------------
    # Run helpers
    # ------------------------------------------------------------------

    def _collect_runs(self, nodes: list[Node], run: RunCtx,
                      direction: str = "ltr") -> list[RunModel]:
        result: list[RunModel] = []
        for node in nodes:
            if isinstance(node, TextNode):
                result.append(self._make_run(node.text, run, direction=direction))
            elif isinstance(node, ElementNode):
                if node.name in STYLE_NODES or node.name == "span":
                    child_run = self._apply_style_props(node, run)
                    result.extend(self._collect_runs(node.children, child_run, direction))
                elif node.name == "link":
                    href = str(node.props.get("href", ""))
                    link_run = run.with_color("0563C1").with_underline()
                    link_runs = self._collect_runs(node.children, link_run, direction)
                    for r in link_runs:
                        r.hyperlink_url = href
                    result.extend(link_runs)
                elif node.name == "page-number":
                    result.append(RunModel(text="", field="PAGE"))
                else:
                    for child in node.children:
                        if isinstance(child, TextNode):
                            result.append(self._make_run(child.text, run, direction=direction))
        return result

    def _apply_style_props(self, node: ElementNode, run: RunCtx) -> RunCtx:
        name = node.name
        if name == "bold":       run = run.with_bold()
        elif name == "italic":   run = run.with_italic()
        elif name == "underline":run = run.with_underline()
        elif name == "strike":   run = run.with_strike()
        elif name == "sup":      run = run.with_sup()
        elif name == "sub":      run = run.with_sub()
        elif name == "color":
            v = node.props.get("value")
            if v:
                h = resolve_color(str(v))
                if h: run = run.with_color(h)
        elif name == "size":
            v = node.props.get("value")
            if v is not None: run = run.with_size(int(v))
        elif name == "font":
            v = node.props.get("value")
            if v: run = run.with_font(str(v))
        elif name == "highlight":
            v = node.props.get("value")
            if v: run = run.with_highlight(str(v))

        if "color" in node.props:
            h = resolve_color(str(node.props["color"]))
            if h: run = run.with_color(h)
        if "size" in node.props:  run = run.with_size(int(node.props["size"]))
        if "font" in node.props:  run = run.with_font(str(node.props["font"]))
        if node.props.get("bold"):   run = run.with_bold()
        if node.props.get("italic"): run = run.with_italic()
        return run

    def _make_run(self, text: str, run: RunCtx,
                  para: ParaCtx | None = None,
                  direction: str = "ltr") -> RunModel:
        dir_ = para.direction if para else direction
        return RunModel(
            text=text, bold=run.bold, italic=run.italic,
            underline=run.underline, strike=run.strike,
            sup=run.sup, sub=run.sub,
            color=run.color, highlight=run.highlight,
            size_pt=run.size_pt, font=run.font,
            rtl=(dir_ == "rtl"),
        )

    def _merge_runs(self, runs: list[RunModel]) -> list[RunModel]:
        if not runs:
            return []
        merged: list[RunModel] = [runs[0]]
        for r in runs[1:]:
            prev = merged[-1]
            if (prev.bold == r.bold and prev.italic == r.italic
                    and prev.underline == r.underline and prev.strike == r.strike
                    and prev.sup == r.sup and prev.sub == r.sub
                    and prev.color == r.color and prev.highlight == r.highlight
                    and prev.size_pt == r.size_pt and prev.font == r.font
                    and prev.rtl == r.rtl and prev.shading == r.shading
                    and prev.hyperlink_url == r.hyperlink_url
                    and prev.field is None and r.field is None):
                merged[-1] = RunModel(
                    text=prev.text + r.text,
                    bold=prev.bold, italic=prev.italic,
                    underline=prev.underline, strike=prev.strike,
                    sup=prev.sup, sub=prev.sub,
                    color=prev.color, highlight=prev.highlight,
                    size_pt=prev.size_pt, font=prev.font,
                    rtl=prev.rtl, shading=prev.shading,
                    hyperlink_url=prev.hyperlink_url,
                )
            else:
                merged.append(r)
        return merged
