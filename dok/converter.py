"""
dok.converter
~~~~~~~~~~~~~
Walks the node tree and produces a DocxModel.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from .nodes   import (Node, ElementNode, TextNode, ArrowNode,
                       LAYOUT_NODES, STYLE_NODES, SHAPE_PRESETS,
                       BLOCK_NODES)
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


@dataclass
class ParagraphModel:
    runs:         list[RunModel] = field(default_factory=list)
    style:        str            = "Normal"
    align:        str            = "left"
    direction:    str            = "ltr"
    indent_twips: int            = 0
    space_before: int            = 0
    space_after:  int            = 160


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
    full_width:   bool                  = False
    accent:       str | None            = None
    paragraphs:   list[ParagraphModel]  = field(default_factory=list)


@dataclass
class RowModel:
    shapes: list[ShapeModel]    = field(default_factory=list)
    arrows: list[str | None]    = field(default_factory=list)   # labels between shapes


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
class PageBreakModel:
    pass


@dataclass
class SectionModel:
    margin: str = "normal"
    paper:  str = "a4"
    cols:   int = 1


@dataclass
class DocxModel:
    content:        list              = field(default_factory=list)
    sections:       list[SectionModel] = field(default_factory=list)
    default_font:   str               = "Calibri"
    default_size_pt: int              = 11


# ---------------------------------------------------------------------------
# Converter
# ---------------------------------------------------------------------------

class Converter:

    def __init__(self) -> None:
        self._model        = DocxModel()
        self._float_next   = None   # "right" | "left" | None

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def convert(self, nodes: list[Node]) -> DocxModel:
        # Default section
        self._model.sections.append(SectionModel())

        # If the root is a single doc node, unpack it
        if (len(nodes) == 1
                and isinstance(nodes[0], ElementNode)
                and nodes[0].name == "doc"):
            doc = nodes[0]
            self._model.default_font    = str(doc.props.get("font", "Calibri"))
            self._model.default_size_pt = int(doc.props.get("size", 11))
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
            # Bare text at document level — wrap in a paragraph
            p = ParagraphModel(
                runs=[self._make_run(node.text, run, para)],
                style=para.style,
                align=para.align,
                direction=para.direction,
                indent_twips=para.indent_twips(),
                space_before=para.space_before,
                space_after=para.space_after,
            )
            self._model.content.append(p)
            return

        if isinstance(node, ArrowNode):
            return  # handled inside _handle_row

        if not isinstance(node, ElementNode):
            return

        name = node.name

        if name == "doc":
            self._model.default_font    = str(node.props.get("font", "Calibri"))
            self._model.default_size_pt = int(node.props.get("size", 11))
            for child in node.children:
                self._walk(child, para, run)

        elif name == "page":
            self._handle_page(node, para, run)

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

        elif name in STYLE_NODES:
            self._handle_style(node, para, run)

        elif name == "---":
            self._emit_page_break()

        elif name in BLOCK_NODES:
            self._emit_paragraph(node, para, run)

        elif name in SHAPE_PRESETS:
            self._emit_shape(node, para, run)

        else:
            # Unknown node — recurse children so nothing is silently lost
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
        # Replace default section or append
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

        if name == "center":
            para = para.with_align("center")
        elif name == "right":
            para = para.with_align("right")
        elif name == "justify":
            para = para.with_align("justify")
        elif name == "left":
            para = para.with_align("left")
        elif name == "rtl":
            para = para.with_direction("rtl")
        elif name == "ltr":
            para = para.with_direction("ltr")
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

        if name == "bold":
            run = run.with_bold()
        elif name == "italic":
            run = run.with_italic()
        elif name == "underline":
            run = run.with_underline()
        elif name == "strike":
            run = run.with_strike()
        elif name == "sup":
            run = run.with_sup()
        elif name == "sub":
            run = run.with_sub()
        elif name == "color":
            v = node.props.get("value") or (node.children[0].text
                if node.children and isinstance(node.children[0], TextNode) else None)
            if v:
                hex_color = resolve_color(str(v))
                if hex_color:
                    run = run.with_color(hex_color)
        elif name == "size":
            v = node.props.get("value")
            if v is not None:
                run = run.with_size(int(v))
        elif name == "font":
            v = node.props.get("value")
            if v:
                run = run.with_font(str(v))
        elif name == "highlight":
            v = node.props.get("value")
            if v:
                run = run.with_highlight(str(v))
        elif name == "span":
            # span can carry multiple props directly
            pass

        # Props directly on style nodes: bold(color: red) { }
        if "color" in node.props:
            hex_color = resolve_color(str(node.props["color"]))
            if hex_color:
                run = run.with_color(hex_color)
        if "size" in node.props:
            run = run.with_size(int(node.props["size"]))
        if "font" in node.props:
            run = run.with_font(str(node.props["font"]))
        if node.props.get("bold"):
            run = run.with_bold()
        if node.props.get("italic"):
            run = run.with_italic()

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
                shape = self._build_shape(child, para, run)
                row_model.shapes.append(shape)
            # Text nodes and other elements inside row are ignored

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

    def _emit_paragraph(self, node: ElementNode,
                         para: ParaCtx, run: RunCtx) -> None:
        name = node.name

        if name == "h1":
            para = para.as_heading(1)
        elif name == "h2":
            para = para.as_heading(2)
        elif name == "h3":
            para = para.as_heading(3)
        elif name == "h4":
            para = para.as_heading(4)
        elif name == "quote":
            para = para.as_quote()
        elif name == "code":
            para = para.as_code()

        # Props directly on paragraph nodes
        if "color" in node.props:
            hex_color = resolve_color(str(node.props["color"]))
            if hex_color:
                run = run.with_color(hex_color)
        if "size" in node.props:
            run = run.with_size(int(node.props["size"]))

        runs  = self._collect_runs(node.children, run, para.direction)
        runs  = self._merge_runs(runs)

        model = ParagraphModel(
            runs         = runs,
            style        = para.style,
            align        = para.align,
            direction    = para.direction,
            indent_twips = para.indent_twips(),
            space_before = para.space_before,
            space_after  = para.space_after,
        )
        self._model.content.append(model)

    def _collect_runs(self, nodes: list[Node],
                      run: RunCtx,
                      direction: str = "ltr") -> list[RunModel]:
        result: list[RunModel] = []

        for node in nodes:
            if isinstance(node, TextNode):
                result.append(self._make_run(node.text, run, direction=direction))

            elif isinstance(node, ElementNode):
                if node.name in STYLE_NODES or node.name == "span":
                    # Update run ctx for this subtree
                    child_run = self._apply_style_props(node, run)
                    result.extend(self._collect_runs(node.children, child_run, direction))
                else:
                    # Non-style element inside paragraph (unusual) — extract text
                    for child in node.children:
                        if isinstance(child, TextNode):
                            result.append(self._make_run(child.text, run, direction=direction))

        return result

    def _apply_style_props(self, node: ElementNode, run: RunCtx) -> RunCtx:
        """Apply a style node's name and props to a RunCtx and return the new one."""
        name = node.name

        if name == "bold":      run = run.with_bold()
        elif name == "italic":  run = run.with_italic()
        elif name == "underline": run = run.with_underline()
        elif name == "strike":  run = run.with_strike()
        elif name == "sup":     run = run.with_sup()
        elif name == "sub":     run = run.with_sub()
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

        # Inline props: bold(color: red)
        if "color" in node.props:
            h = resolve_color(str(node.props["color"]))
            if h: run = run.with_color(h)
        if "size" in node.props:
            run = run.with_size(int(node.props["size"]))
        if "font" in node.props:
            run = run.with_font(str(node.props["font"]))
        if node.props.get("bold"):
            run = run.with_bold()
        if node.props.get("italic"):
            run = run.with_italic()

        return run

    def _make_run(self, text: str, run: RunCtx,
                  para: ParaCtx | None = None,
                  direction: str = "ltr") -> RunModel:
        dir_ = para.direction if para else direction
        return RunModel(
            text      = text,
            bold      = run.bold,
            italic    = run.italic,
            underline = run.underline,
            strike    = run.strike,
            sup       = run.sup,
            sub       = run.sub,
            color     = run.color,
            highlight = run.highlight,
            size_pt   = run.size_pt,
            font      = run.font,
            rtl       = (dir_ == "rtl"),
        )

    def _merge_runs(self, runs: list[RunModel]) -> list[RunModel]:
        if not runs:
            return []

        merged: list[RunModel] = [runs[0]]

        for r in runs[1:]:
            prev = merged[-1]
            # Compare all fields except text
            if (prev.bold == r.bold and prev.italic   == r.italic
                    and prev.underline == r.underline  and prev.strike == r.strike
                    and prev.sup       == r.sup         and prev.sub    == r.sub
                    and prev.color     == r.color       and prev.highlight == r.highlight
                    and prev.size_pt   == r.size_pt     and prev.font   == r.font
                    and prev.rtl       == r.rtl):
                merged[-1] = RunModel(
                    text=prev.text + r.text,
                    bold=prev.bold, italic=prev.italic,
                    underline=prev.underline, strike=prev.strike,
                    sup=prev.sup, sub=prev.sub,
                    color=prev.color, highlight=prev.highlight,
                    size_pt=prev.size_pt, font=prev.font,
                    rtl=prev.rtl,
                )
            else:
                merged.append(r)

        return merged

    # ------------------------------------------------------------------
    # Shape
    # ------------------------------------------------------------------

    def _emit_shape(self, node: ElementNode,
                    para: ParaCtx, run: RunCtx) -> None:
        shape = self._build_shape(node, para, run)
        self._model.content.append(shape)

    def _build_shape(self, node: ElementNode,
                     para: ParaCtx, run: RunCtx) -> ShapeModel:
        props  = node.props
        name   = node.name
        preset = SHAPE_PRESETS.get(name, "rect")

        fill   = resolve_color(str(props["fill"]))   if "fill"   in props else None
        stroke = resolve_color(str(props["stroke"])) if "stroke" in props else "BFBFBF"
        color  = resolve_color(str(props["color"]))  if "color"  in props else None
        accent = resolve_color(str(props["accent"])) if "accent" in props else None

        # stroke prop can also be a style keyword
        stroke_style = "solid"
        stroke_thick = False
        if "stroke" in props:
            v = str(props["stroke"])
            if v == "dashed":
                stroke = "BFBFBF"; stroke_style = "dashed"
            elif v == "dotted":
                stroke = "BFBFBF"; stroke_style = "dotted"
            elif v == "thick":
                stroke_thick = True
            elif v == "none":
                stroke = None
            elif v == "thin":
                stroke_thick = False

        float_side = self._float_next
        inline     = float_side is None

        shape = ShapeModel(
            preset       = preset,
            fill         = fill,
            stroke       = stroke,
            stroke_style = stroke_style,
            stroke_thick = stroke_thick,
            color        = color,
            rounded      = bool(props.get("rounded", False)),
            shadow       = bool(props.get("shadow",  False)),
            inline       = inline,
            float_side   = float_side,
            full_width   = (name == "banner"),
            accent       = accent,
        )

        # Text inside the shape — convert children to ParagraphModels
        if node.children:
            sub = Converter()
            sub._model.default_font    = self._model.default_font
            sub._model.default_size_pt = self._model.default_size_pt

            shape_run = run
            if color:
                shape_run = shape_run.with_color(color)

            for child in node.children:
                if isinstance(child, TextNode):
                    # Bare string inside shape → wrap as paragraph
                    p_model = ParagraphModel(
                        runs=[sub._make_run(child.text, shape_run)],
                    )
                    shape.paragraphs.append(p_model)
                elif isinstance(child, ElementNode):
                    sub._walk(child, para, shape_run)
                    shape.paragraphs.extend(
                        [c for c in sub._model.content
                         if isinstance(c, ParagraphModel)]
                    )
                    sub._model.content.clear()

        return shape

    # ------------------------------------------------------------------
    # Page break
    # ------------------------------------------------------------------

    def _emit_page_break(self) -> None:
        self._model.content.append(PageBreakModel())