"""
dok.converter
~~~~~~~~~~~~~
Walks the node tree and produces a DocxModel.

Print-friendly: box/callout/banner/badge/line use native Word formatting.
Only circle/diamond/chevron use drawing shapes.
"""

from __future__ import annotations
from pathlib import Path
from typing import Any
from .nodes   import (Node, ElementNode, TextNode, ArrowNode, SHAPE_PRESETS)
from . import registry
from .context import ParaCtx, RunCtx
from .colors    import resolve as resolve_color
from .constants import INCH_TO_EMU
from .units     import parse_to_twips, parse_to_emu, parse_to_pt
from .models    import (
    RunModel, ParagraphModel, LineModel, BoxModel, BannerModel, BadgeModel,
    ShapeModel, RowModel, TableModel, TableRowModel, TableCellModel,
    DataTableModel, DataTableRowModel, DataTableCellModel,
    ImageModel, SpacerModel, HeaderModel, FooterModel,
    PageBreakModel, SectionModel, DocxModel,
    TocModel, TocEntry,
    CheckboxModel, TextInputModel, DropdownModel, ToggleModel, FrameModel,
    SPACING_PRESETS,
)

# Re-export models so existing `from .converter import RunModel` keeps working
__all__ = [
    "RunModel", "ParagraphModel", "LineModel", "BoxModel", "BannerModel",
    "BadgeModel", "ShapeModel", "RowModel", "TableModel", "TableRowModel",
    "TableCellModel", "DataTableModel", "DataTableRowModel",
    "DataTableCellModel", "ImageModel", "SpacerModel", "HeaderModel",
    "FooterModel", "PageBreakModel", "SectionModel", "DocxModel",
    "TocModel", "TocEntry", "SPACING_PRESETS", "Converter",
]

# Backward compat alias
_SPACING_PRESETS = SPACING_PRESETS

_INCH_TO_EMU = INCH_TO_EMU  # backward compat alias


# ---------------------------------------------------------------------------
# Converter
# ---------------------------------------------------------------------------

class Converter:

    def __init__(self) -> None:
        self._model       = DocxModel()
        self._float_next  = None
        self._toc_entries: list[TocEntry] = []
        self._bookmark_id = 0
        self._available_width_twips: int = 0  # 0 = use page width

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
            # Typography settings
            if "kerning" in doc.props:
                self._model.kerning = bool(doc.props["kerning"])
            if "ligatures" in doc.props:
                self._model.ligatures = bool(doc.props["ligatures"])
            if "widow-orphan" in doc.props:
                self._model.widow_orphan = int(doc.props["widow-orphan"])
            if "hyphenate" in doc.props:
                self._model.hyphenate = bool(doc.props["hyphenate"])
            nodes = doc.children

        para = ParaCtx()
        run  = RunCtx()
        for node in nodes:
            self._walk(node, para, run)

        # Post-convert: fill TOC entries
        for item in self._model.content:
            if isinstance(item, TocModel):
                item.entries = [e for e in self._toc_entries if e.level <= item.depth]

        return self._model

    # ------------------------------------------------------------------
    # Sub-converter — DRY helper for nested content
    # ------------------------------------------------------------------

    def _sub_convert(self, children: list[Node],
                     para: ParaCtx, run: RunCtx,
                     available_width: int = 0) -> list:
        """Convert children in a fresh scope. Returns the content list."""
        sub = Converter()
        sub._model.default_font    = self._model.default_font
        sub._model.default_size_pt = self._model.default_size_pt
        sub._available_width_twips = available_width or self._available_width_twips
        for child in children:
            sub._walk(child, para, run)
        return sub._model.content

    def _sub_convert_mixed(self, children: list[Node],
                           para: ParaCtx, run: RunCtx,
                           available_width: int = 0) -> list:
        """Like _sub_convert but wraps bare TextNodes as paragraphs."""
        sub = Converter()
        sub._model.default_font    = self._model.default_font
        sub._model.default_size_pt = self._model.default_size_pt
        sub._available_width_twips = available_width or self._available_width_twips
        for child in children:
            if isinstance(child, TextNode):
                sub._model.content.append(ParagraphModel(
                    runs=[self._make_run(child.text, run)], align=para.align,
                ))
            else:
                sub._walk(child, para, run)
        return sub._model.content

    # ------------------------------------------------------------------
    # Walker
    # ------------------------------------------------------------------

    def _walk(self, node: Node, para: ParaCtx, run: RunCtx) -> None:
        if isinstance(node, TextNode):
            self._model.content.append(ParagraphModel(
                runs=[self._make_run(node.text, run, para)],
                style=para.style, align=para.align,
                direction=para.direction,
                indent_twips=para.indent_twips(),
                space_before=para.space_before,
                space_after=para.space_after,
            ))
            return

        if isinstance(node, ArrowNode):
            return

        if not isinstance(node, ElementNode):
            return

        name = node.name
        elem = registry.get(name)

        if not elem:
            # Unknown element: recurse children (never silently drop)
            for child in node.children:
                self._walk(child, para, run)
            return

        cat = elem.category

        # Category-based dispatch
        if cat == "style":
            run = self._apply_style_props(node, run)
            for child in node.children:
                self._walk(child, para, run)
        elif cat == "layout" and not elem.handler:
            self._handle_layout(node, para, run)
        elif cat == "block":
            self._emit_paragraph(node, para, run)
        elif elem.handler:
            getattr(self, elem.handler)(node, para, run)
        else:
            for child in node.children:
                self._walk(child, para, run)

    # ------------------------------------------------------------------
    # Doc / Page / Float / PageBreak / Spacer
    # ------------------------------------------------------------------

    def _handle_doc(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        self._model.default_font    = str(node.props.get("font", "Calibri"))
        self._model.default_size_pt = int(node.props.get("size", 11))
        self._model.spacing         = str(node.props.get("spacing", "normal"))
        for child in node.children:
            self._walk(child, para, run)

    def _handle_page(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        section = SectionModel(
            margin=str(node.props.get("margin", "normal")),
            paper =str(node.props.get("paper",  "a4")),
            cols  =int(node.props.get("cols",   1)),
        )
        # Exact margin overrides — supports units (pt, cm, mm, in)
        for side in ("top", "right", "bottom", "left"):
            val = node.props.get(f"margin-{side}")
            if val is not None:
                setattr(section, f"margin_{side}", parse_to_twips(val))
            pad = node.props.get(f"padding-{side}")
            if pad is not None:
                setattr(section, f"padding_{side}", parse_to_twips(pad))

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
    # Style — single method for applying style props to RunCtx
    # ------------------------------------------------------------------

    def _apply_style_props(self, node: ElementNode, run: RunCtx) -> RunCtx:
        """Apply style from node name and props to the run context."""
        name = node.name
        # Style from node name (e.g. bold{}, color(value:red){})
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

        # Inline style props (e.g. bold(color: red, size: 14){})
        if "color" in node.props:
            h = resolve_color(str(node.props["color"]))
            if h: run = run.with_color(h)
        if "size" in node.props:  run = run.with_size(int(node.props["size"]))
        if "font" in node.props:  run = run.with_font(str(node.props["font"]))
        if node.props.get("bold"):   run = run.with_bold()
        if node.props.get("italic"): run = run.with_italic()
        return run

    def _handle_float(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        self._float_next = node.props.get("side", "right")
        for child in node.children:
            self._walk(child, para, run)
        self._float_next = None

    def _handle_page_break(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        self._model.content.append(PageBreakModel())

    def _emit_spacer(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        raw = node.props.get("size", 10)
        self._model.content.append(SpacerModel(height_twips=parse_to_twips(raw)))

    def _emit_toc(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        depth = int(node.props.get("depth", 4))
        title = str(node.props.get("title", "Table of Contents"))
        # Insert a TocModel placeholder — entries are filled in post-convert
        self._model.content.append(TocModel(depth=depth, title=title))

    def _emit_ref(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        target = str(node.props.get("to", ""))
        # Collect display text from children
        text_parts: list[str] = []
        for child in node.children:
            if isinstance(child, TextNode):
                text_parts.append(child.text)
        display = "".join(text_parts) or target
        # Create a run with internal hyperlink
        ref_run = RunModel(text=display, hyperlink_url=f"#{target}",
                           color="0563C1", underline=True)
        self._model.content.append(ParagraphModel(
            runs=[ref_run], style=para.style, align=para.align,
            direction=para.direction,
        ))

    # ------------------------------------------------------------------
    # Row
    # ------------------------------------------------------------------

    def _handle_row(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        row_model = RowModel()
        for child in node.children:
            if isinstance(child, ArrowNode):
                row_model.arrows.append(child.label)
            elif isinstance(child, ElementNode):
                if child.name in SHAPE_PRESETS:
                    row_model.items.append(self._build_drawing_shape(child, para, run))
                else:
                    content = self._sub_convert_mixed([child], para, run)
                    row_model.items.extend(content)
            elif isinstance(child, TextNode):
                row_model.items.append(ParagraphModel(
                    runs=[self._make_run(child.text, run, para)],
                    align=para.align,
                ))
        self._model.content.append(row_model)

    # ------------------------------------------------------------------
    # Cols
    # ------------------------------------------------------------------

    def _handle_cols(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        col_nodes = [c for c in node.children
                    if isinstance(c, ElementNode) and c.name == "col"]
        n_cols = len(col_nodes) or 2

        # Parse ratio — default to equal distribution matching actual col count
        ratio_str = str(node.props.get("ratio", ":".join(["1"] * n_cols)))
        parts     = [int(p) for p in ratio_str.split(":")]

        # Pad or truncate ratio to match actual col count
        while len(parts) < n_cols:
            parts.append(parts[-1] if parts else 1)
        parts = parts[:n_cols]

        total  = sum(parts)
        widths = [round(p / total * 100) for p in parts]

        # Ensure widths sum to exactly 100
        diff = 100 - sum(widths)
        if diff and widths:
            widths[-1] += diff

        gap_raw    = node.props.get("gap", 0)
        padding_raw = node.props.get("padding", 0)
        fill       = node.props.get("fill")
        border     = bool(node.props.get("border", False))

        table = TableModel(
            border=border,
            gap_twips=parse_to_twips(gap_raw),
            cell_padding_twips=parse_to_twips(padding_raw),
            fill=resolve_color(str(fill)) if fill else None,
        )
        row = TableRowModel()
        for i, col_node in enumerate(col_nodes):
            pct = widths[i] if i < len(widths) else 100 // n_cols
            col_padding_raw = col_node.props.get("padding", 0)
            col_fill    = col_node.props.get("fill")
            col_align   = col_node.props.get("align")

            col_padding_twips = parse_to_twips(col_padding_raw)
            cell = TableCellModel(
                width_pct=pct,
                padding_twips=col_padding_twips,
                fill=resolve_color(str(col_fill)) if col_fill else None,
                align=str(col_align) if col_align else None,
            )

            # Apply column-level alignment to child context
            col_para = para
            if col_align:
                col_para = para.with_align(str(col_align))

            # Calculate actual cell width in twips for child width propagation
            from .constants import content_width_twip, MARGIN_PRESETS
            section = self._model.sections[-1] if self._model.sections else SectionModel()
            cw = content_width_twip(section.paper,
                                    margins=section.resolved_margins(MARGIN_PRESETS))
            gap_twips = parse_to_twips(gap_raw)
            total_gap = gap_twips * max(0, n_cols - 1)
            usable_w = cw - total_gap
            total_pct_sum = sum(widths)
            cell_w_twips = int(usable_w * pct / total_pct_sum) if total_pct_sum else usable_w // n_cols
            cell_inner = max(0, cell_w_twips - col_padding_twips * 2 - parse_to_twips(padding_raw) * 2)

            cell.content = self._sub_convert(col_node.children, col_para, run,
                                             available_width=cell_inner)
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

        # Per-paragraph spacing override
        spacing_name = node.props.get("spacing")
        line_spacing = 0
        if spacing_name:
            from .models import SPACING_PRESETS
            preset = SPACING_PRESETS.get(spacing_name)
            if preset:
                para = para.__class__(
                    align=para.align, direction=para.direction, indent=para.indent,
                    style=para.style, space_before=para.space_before,
                    space_after=preset[0],
                )
                line_spacing = preset[2]
        if "line-height" in node.props:
            # line-height in tenths: 10 = single, 15 = 1.5x, 20 = double
            lh = int(node.props["line-height"])
            line_spacing = lh * 24   # 10 * 24 = 240 twips (single)

        runs  = self._collect_runs(node.children, run, para.direction)
        runs  = self._merge_runs(runs)

        # Bookmark: explicit id prop or auto-generated for headings
        bookmark = None
        explicit_id = node.props.get("id")
        if explicit_id:
            bookmark = str(explicit_id)
        elif name in ("h1", "h2", "h3", "h4"):
            self._bookmark_id += 1
            bookmark = f"_heading_{self._bookmark_id}"

        para_model = ParagraphModel(
            runs=runs, style=para.style, align=para.align,
            direction=para.direction, indent_twips=para.indent_twips(),
            space_before=para.space_before, space_after=para.space_after,
            line_spacing=line_spacing, bookmark=bookmark,
        )
        self._model.content.append(para_model)

        # Collect TOC entry for headings
        if name in ("h1", "h2", "h3", "h4") and bookmark:
            level = int(name[1])
            text = "".join(r.text for r in runs)
            self._toc_entries.append(TocEntry(text=text, level=level, anchor=bookmark))

    # ------------------------------------------------------------------
    # Line
    # ------------------------------------------------------------------

    def _emit_line(self, node: ElementNode, para: ParaCtx = None, run: RunCtx = None) -> None:
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

    # Per-variant default overrides (convention over configuration)
    _BOX_DEFAULTS: dict[str, dict[str, Any]] = {
        "callout": {"fill": "FFF8E1", "stroke": "FFB300", "accent": "FFB300"},
        "banner":  {"fill": "E8EAF6", "stroke": "3F51B5", "accent": "3F51B5"},
        "badge":   {"fill": "1F3864", "color": "FFFFFF", "inline": True},
    }

    def _emit_box(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        defaults = self._BOX_DEFAULTS.get(node.name, {})

        fill    = resolve_color(str(props["fill"]))    if "fill"    in props else defaults.get("fill")
        stroke  = resolve_color(str(props["stroke"]))  if "stroke"  in props else defaults.get("stroke", "BFBFBF")
        color   = resolve_color(str(props["color"]))   if "color"   in props else defaults.get("color")
        accent  = resolve_color(str(props["accent"]))  if "accent"  in props else defaults.get("accent")
        inline  = bool(props.get("inline", defaults.get("inline", False)))
        rounded = bool(props.get("rounded", False))
        shadow  = bool(props.get("shadow", False))
        width_pct = min(int(props.get("width", 0)), 100)  # cap at 100%
        height_pt = int(props.get("height", 0))

        # Badge: inline text box
        if inline:
            text_parts = [c.text for c in node.children if isinstance(c, TextNode)]
            self._model.content.append(BoxModel(
                text=" ".join(text_parts), fill=fill, color=color,
                inline=True, align=para.align,
            ))
            return

        # Compute available width for child elements
        parent_width = self._available_width_twips
        # Box padding: 160 twips left + 160 twips right = 320 twips
        box_padding = 320
        inner_width = max(0, (parent_width - box_padding)) if parent_width else 0

        # All boxes (including callout/banner) use the same rendering path
        box_run = run.with_color(color) if color else run
        content = self._sub_convert_mixed(node.children, para, box_run,
                                          available_width=inner_width)

        def _bool_prop(key: str, default: bool = True) -> bool:
            v = props.get(key, default)
            if isinstance(v, str):
                return v.lower() not in ("false", "no", "0", "off")
            return bool(v)

        self._model.content.append(BoxModel(
            content=content, fill=fill, stroke=stroke, accent=accent,
            rounded=rounded, shadow=shadow, color=color,
            width_pct=width_pct, height_pt=height_pt,
            max_width_twips=parent_width,
            border_top=_bool_prop("border-top"),
            border_right=_bool_prop("border-right"),
            border_bottom=_bool_prop("border-bottom"),
            border_left=_bool_prop("border-left"),
            border_width=int(props.get("border-width", 1)),
        ))

    # ------------------------------------------------------------------
    # Lists
    # ------------------------------------------------------------------

    def _emit_list(self, node: ElementNode, para: ParaCtx, run: RunCtx,
                   ilvl: int = 0) -> None:
        is_ordered = (node.name == "ol")
        marker = node.props.get("marker")
        list_format = None

        if is_ordered:
            num_id = 2
            if marker:
                list_format = str(marker)  # "decimal", "alpha", "roman"
        else:
            num_id = 1

        # Custom bullet marker uses num_id=3+
        custom_marker = None
        if not is_ordered and marker:
            custom_marker = str(marker)
            if custom_marker not in self._model.custom_markers:
                self._model.custom_markers.append(custom_marker)
            num_id = 3 + self._model.custom_markers.index(custom_marker)

        self._model.has_lists = True

        for child in node.children:
            if isinstance(child, ElementNode) and child.name == "li":
                self._emit_list_item(child, para, run, num_id, ilvl,
                                     custom_marker=custom_marker,
                                     list_format=list_format)
            elif isinstance(child, ElementNode) and child.name in ("ul", "ol"):
                self._emit_list(child, para, run, ilvl + 1)

    def _emit_list_item(self, node: ElementNode, para: ParaCtx, run: RunCtx,
                        num_id: int, ilvl: int,
                        custom_marker: str | None = None,
                        list_format: str | None = None) -> None:
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
            if runs:
                self._model.content.append(ParagraphModel(
                    runs=runs, align=para.align, direction=para.direction,
                    num_id=num_id, num_ilvl=ilvl,
                    list_marker=custom_marker,
                    list_format=list_format,
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
            direction=para.direction,
        )

        for child in node.children:
            if isinstance(child, ElementNode) and child.name == "tr":
                row = DataTableRowModel()
                has_th = False
                # Row-level overrides
                row_align = child.props.get("align")
                row_dir   = child.props.get("direction")
                for cell_node in child.children:
                    if isinstance(cell_node, ElementNode) and cell_node.name in ("td", "th"):
                        is_th = (cell_node.name == "th")
                        if is_th: has_th = True
                        colspan = int(cell_node.props.get("colspan", 1))

                        # Cell-level overrides → row-level → parent context
                        cell_dir   = str(cell_node.props.get("direction", row_dir or ""))
                        cell_align = str(cell_node.props.get("align", row_align or ""))
                        cell_fill  = cell_node.props.get("fill")

                        cell_para = para
                        if cell_dir:
                            cell_para = cell_para.with_direction(cell_dir)
                        if cell_align:
                            cell_para = cell_para.with_align(cell_align)

                        cell_run = run.with_bold() if is_th else run
                        cell_content = self._sub_convert_mixed(
                            cell_node.children, cell_para, cell_run)
                        row.cells.append(DataTableCellModel(
                            content=cell_content, is_th=is_th, colspan=colspan,
                            align=cell_align or None,
                            direction=cell_dir or None,
                            fill=resolve_color(str(cell_fill)) if cell_fill else None,
                        ))
                row.is_header = has_th
                table.rows.append(row)

        # Auto-calculate column widths based on content length
        table.col_widths = self._calc_table_col_widths(table)

        self._model.content.append(table)

    @staticmethod
    def _text_length(content: list) -> int:
        """Estimate text length from a list of model items."""
        total = 0
        for item in content:
            if isinstance(item, ParagraphModel):
                for r in item.runs:
                    total += len(r.text) if r.text else 0
        return total

    @staticmethod
    def _calc_table_col_widths(table: DataTableModel) -> list[int]:
        """Calculate proportional column widths (pct) based on content length."""
        if not table.rows:
            return []

        max_cols = max((sum(c.colspan for c in r.cells) for r in table.rows), default=1)
        if max_cols <= 1:
            return [100]

        # Accumulate character counts per logical column
        col_chars: list[float] = [0.0] * max_cols
        for row in table.rows:
            col_idx = 0
            for cell in row.cells:
                length = Converter._text_length(cell.content)
                # Headers weigh slightly more
                weight = 1.2 if cell.is_th else 1.0
                chars = length * weight
                # Spread across spanned columns
                per_col = chars / cell.colspan
                for i in range(cell.colspan):
                    if col_idx + i < max_cols:
                        col_chars[col_idx + i] += per_col
                col_idx += cell.colspan

        total = sum(col_chars)
        if total == 0:
            # No text content — equal widths
            return [100 // max_cols] * max_cols

        # Calculate proportional percentages with minimum 8% per column
        MIN_PCT = 8
        raw = [c / total * 100 for c in col_chars]

        # Clamp columns below minimum, redistribute
        widths: list[int] = []
        clamped_total = 0
        free_total = 0.0
        for r in raw:
            if r < MIN_PCT:
                clamped_total += MIN_PCT
            else:
                free_total += r

        remaining = 100 - clamped_total
        for r in raw:
            if r < MIN_PCT:
                widths.append(MIN_PCT)
            else:
                widths.append(max(MIN_PCT, int(r / free_total * remaining))) if free_total > 0 else widths.append(MIN_PCT)

        # Adjust rounding so widths sum to 100
        diff = 100 - sum(widths)
        if diff != 0 and widths:
            # Add/subtract from the widest column
            widest = widths.index(max(widths))
            widths[widest] += diff

        return widths

    # ------------------------------------------------------------------
    # Image
    # ------------------------------------------------------------------

    def _emit_image(self, node: ElementNode, para: ParaCtx = None, run: RunCtx = None) -> None:
        props = node.props
        src   = str(props.get("src", ""))
        raw_w = props.get("width", 300)    # default 300pt (~4.17 inches)
        raw_h = props.get("height", 0)

        width_emu = parse_to_emu(raw_w)
        height_emu = parse_to_emu(raw_h) if raw_h else 0

        if height_emu == 0:
            height_emu = self._auto_image_height(src, width_emu)

        self._model.content.append(ImageModel(
            src=src,
            width_emu=width_emu,
            height_emu=height_emu,
            align=para.align if para else "left",
        ))

    def _auto_image_height(self, src: str, width_emu: int) -> int:
        """Calculate height EMU from actual image aspect ratio."""
        try:
            from .image import image_dimensions
            base = self._model.base_dir
            path = (base / src) if base else Path(src)
            w, h = image_dimensions(path)
            if w > 0 and h > 0:
                return int(width_emu * h / w)
        except Exception:
            pass
        # Fallback: assume square
        return width_emu

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

    def _emit_page_number(self, node: ElementNode = None, para: ParaCtx = None, run: RunCtx = None) -> None:
        self._model.content.append(ParagraphModel(
            runs=[RunModel(text="", field="PAGE")],
            align=para.align,
        ))

    # ------------------------------------------------------------------
    # Header / Footer
    # ------------------------------------------------------------------

    def _emit_header(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        content = self._sub_convert(node.children, para, run)
        paras = [c for c in content if isinstance(c, ParagraphModel)]
        self._model.header = HeaderModel(paragraphs=paras)

    def _emit_footer(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        content = self._sub_convert(node.children, para, run)
        paras = [c for c in content if isinstance(c, ParagraphModel)]
        self._model.footer = FooterModel(paragraphs=paras)

    # ------------------------------------------------------------------
    # Form fields
    # ------------------------------------------------------------------

    def _emit_frame(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        fill = resolve_color(str(props["fill"])) if "fill" in props else None
        stroke = resolve_color(str(props["stroke"])) if "stroke" in props else None
        content = self._sub_convert_mixed(node.children, para, run)
        self._model.content.append(FrameModel(
            x_twips=parse_to_twips(props.get("x", 0)),
            y_twips=parse_to_twips(props.get("y", 0)),
            width_twips=parse_to_twips(props.get("width", 200)),
            height_twips=parse_to_twips(props.get("height", 100)),
            fill=fill, stroke=stroke,
            rounded=bool(props.get("rounded", False)),
            shadow=bool(props.get("shadow", False)),
            anchor=str(props.get("anchor", "page")),
            content=content,
        ))

    def _emit_toggle(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        title = str(props.get("title", "Details"))
        is_open = props.get("open", False)
        if isinstance(is_open, str):
            is_open = is_open.lower() not in ("false", "no", "0", "off")
        content = self._sub_convert_mixed(node.children, para, run)
        self._model.content.append(ToggleModel(
            title=title, open=bool(is_open), content=content,
        ))

    def _emit_checkbox(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        checked = props.get("checked", False)
        if isinstance(checked, str):
            checked = checked.lower() not in ("false", "no", "0", "off")
        label = str(props.get("label", ""))
        if not label:
            # Get label from child text
            for child in node.children:
                if isinstance(child, TextNode):
                    label = child.text
                    break
        self._model.content.append(CheckboxModel(checked=bool(checked), label=label))

    def _emit_text_input(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        props = node.props
        self._model.content.append(TextInputModel(
            placeholder=str(props.get("placeholder", "")),
            value=str(props.get("value", "")),
            width_pct=int(props.get("width", 100)),
        ))

    def _emit_dropdown(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
        options: list[str] = []
        for child in node.children:
            if isinstance(child, ElementNode) and child.name == "option":
                # Get option text from children or value prop
                opt_text = str(child.props.get("value", ""))
                if not opt_text:
                    for gc in child.children:
                        if isinstance(gc, TextNode):
                            opt_text = gc.text
                            break
                options.append(opt_text)
        self._model.content.append(DropdownModel(
            options=options,
            value=str(node.props.get("value", "")),
        ))

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
            shape_run = run.with_color(color) if color else run
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
                elem = registry.get(node.name)
                if elem and elem.category == "style":
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
            if (prev.style_key() == r.style_key()
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
