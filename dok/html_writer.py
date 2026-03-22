"""
dok.html_writer
~~~~~~~~~~~~~~~
Converts a DocxModel to a standalone .html file.

Mirrors docx_writer.py feature-for-feature:
  paragraphs, runs (bold/italic/underline/strike/color/highlight/super/sub),
  hyperlinks, lists (bullet + ordered), tables, data-tables, boxes, banners,
  badges, images (base64-embedded), spacers, lines, shapes, rows, page breaks,
  header / footer, sections (paper width → max-width, margins).
"""

from __future__ import annotations

import base64
import html
import io
import mimetypes
from pathlib import Path

from .converter import (
    DocxModel, ParagraphModel, RunModel, ShapeModel,
    RowModel, TableModel, TableCellModel,
    PageBreakModel, SectionModel,
    LineModel, BoxModel, BannerModel, BadgeModel,
    DataTableModel, DataTableRowModel, DataTableCellModel,
    ImageModel, SpacerModel, HeaderModel, FooterModel,
    _SPACING_PRESETS,
)


# ---------------------------------------------------------------------------
# Units  (twips / EMUs → CSS)
# ---------------------------------------------------------------------------

def twip_to_px(t: int) -> float:   return round(t / 14.4, 2)     # 1440 twip = 100 px (96 dpi)
def twip_to_pt(t: int) -> float:   return round(t / 20, 2)        # 1 twip = 1/20 pt
def emu_to_px(e: int) -> float:    return round(e / 9144, 2)       # 914400 EMU = 100 px (96 dpi)


MARGIN_PRESETS_PX = {
    "normal": {"top": 96, "right": 96, "bottom": 96, "left": 96},
    "narrow": {"top": 48, "right": 48, "bottom": 48, "left": 48},
    "wide":   {"top": 96, "right": 192, "bottom": 96, "left": 192},
    "none":   {"top":  0, "right":   0, "bottom":  0, "left":   0},
}

# Approximate A4 / Letter content widths in px at 96 dpi
PAPER_MAX_WIDTH_PX = {
    "a4":     794,
    "letter": 816,
    "a3":     1123,
}

HIGHLIGHT_CSS = {
    "yellow":   "#FFFF00", "green":  "#00FF00", "cyan":  "#00FFFF",
    "magenta":  "#FF00FF", "blue":   "#0000FF", "red":   "#FF0000",
    "darkBlue": "#000080", "darkCyan": "#008080", "darkGreen": "#008000",
    "darkMagenta": "#800080", "darkRed": "#800000", "darkYellow": "#808000",
    "darkGray": "#808080", "lightGray": "#C0C0C0", "black": "#000000",
    "white":    "#FFFFFF",
}

# Preset geometry → CSS border-radius approximation
SHAPE_PRESETS = {
    "rect":       "0",
    "roundRect":  "8px",
    "ellipse":    "50%",
    "diamond":    "0",     # handled specially
    "chevron":    "0",     # handled specially
    "hexagon":    "0",
}


# ---------------------------------------------------------------------------
# HtmlWriter
# ---------------------------------------------------------------------------

class HtmlWriter:

    def __init__(self, model: DocxModel) -> None:
        self._model = model
        self._list_counters: dict[tuple[int, int], int] = {}  # (num_id, ilvl) → counter

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def write(self, dest: str | Path | io.StringIO | io.BytesIO) -> None:
        html_text = self._build_html()
        if isinstance(dest, (str, Path)):
            Path(dest).write_text(html_text, encoding="utf-8")
        elif isinstance(dest, io.BytesIO):
            dest.write(html_text.encode("utf-8"))
        else:
            dest.write(html_text)

    # ------------------------------------------------------------------
    # Top-level document
    # ------------------------------------------------------------------

    def _build_html(self) -> str:
        section = self._model.sections[-1] if self._model.sections else SectionModel()
        margins = MARGIN_PRESETS_PX.get(section.margin, MARGIN_PRESETS_PX["normal"])
        max_w   = PAPER_MAX_WIDTH_PX.get(section.paper, PAPER_MAX_WIDTH_PX["a4"])
        content_w = max_w - margins["left"] - margins["right"]

        para_after, h_scale, line_sp = _SPACING_PRESETS.get(
            self._model.spacing, _SPACING_PRESETS["normal"]
        )

        font      = self._model.default_font
        font_size = self._model.default_size_pt

        css = self._build_css(font, font_size, margins, max_w, content_w,
                              para_after, h_scale, line_sp)

        parts: list[str] = []
        parts.append(f"<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n")
        parts.append('<meta charset="UTF-8">\n')
        parts.append('<meta name="viewport" content="width=device-width, initial-scale=1.0">\n')
        parts.append(f"<style>\n{css}\n</style>\n</head>\n<body>\n")

        # Header
        if self._model.header:
            parts.append('<header class="doc-header">\n')
            for para in self._model.header.paragraphs:
                parts.append(self._render_paragraph(para))
            parts.append("</header>\n")

        parts.append('<main class="doc-body">\n')
        for item in self._model.content:
            parts.append(self._render_item(item))
        parts.append("</main>\n")

        # Footer
        if self._model.footer:
            parts.append('<footer class="doc-footer">\n')
            for para in self._model.footer.paragraphs:
                parts.append(self._render_paragraph(para))
            parts.append("</footer>\n")

        parts.append("</body>\n</html>")
        return "".join(parts)

    # ------------------------------------------------------------------
    # CSS
    # ------------------------------------------------------------------

    def _build_css(self, font: str, font_size: float,
                   margins: dict, max_w: int, content_w: int,
                   para_after: int, h_scale: float, line_sp: int) -> str:

        para_after_pt = twip_to_pt(para_after)
        line_height    = round(line_sp / 240, 3)   # 240 = single, 276 ≈ 1.15
        font_size_pt   = font_size

        h_sizes = {1: 27, 2: 24, 3: 21, 4: 19.5}  # approx from half-point values /2
        h_colors = {1: "#1F3864", 2: "#1F3864", 3: "#404040", 4: "#404040"}

        heading_css_parts = []
        for lvl in (1, 2, 3, 4):
            before = round(int({1: 480, 2: 360, 3: 280, 4: 240}[lvl] * h_scale) / 20, 1)
            after  = round(int(120 * h_scale) / 20, 1)
            heading_css_parts.append(f"""
h{lvl} {{
  font-family: '{font}', sans-serif;
  font-size: {h_sizes[lvl]}pt;
  font-weight: bold;
  color: {h_colors[lvl]};
  margin-top: {before}pt;
  margin-bottom: {after}pt;
  line-height: {min(line_height, 1.0)};
}}""")
        heading_css = "\n".join(heading_css_parts)

        return f"""
/* ---- Reset & base ---- */
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

body {{
  font-family: '{font}', Calibri, sans-serif;
  font-size: {font_size_pt}pt;
  line-height: {line_height};
  background: #f0f0f0;
  color: #000;
}}

/* ---- Page shell ---- */
.doc-body {{
  max-width: {max_w}px;
  margin: 0 auto;
  padding: {margins['top']}px {margins['right']}px {margins['bottom']}px {margins['left']}px;
  background: #fff;
  box-shadow: 0 2px 12px rgba(0,0,0,.15);
  min-height: 100vh;
}}

.doc-header, .doc-footer {{
  max-width: {max_w}px;
  margin: 0 auto;
  padding: 8px {margins['right']}px;
  background: #fafafa;
  border-bottom: 1px solid #ddd;
  font-size: {font_size_pt - 1}pt;
  color: #555;
}}
.doc-footer {{
  border-top: 1px solid #ddd;
  border-bottom: none;
}}

/* ---- Paragraphs ---- */
p {{
  margin-bottom: {para_after_pt}pt;
  line-height: {line_height};
}}
p:last-child {{ margin-bottom: 0; }}

/* ---- Headings ---- */
{heading_css}

/* ---- Inline runs ---- */
.run-sup  {{ vertical-align: super;  font-size: 0.75em; }}
.run-sub  {{ vertical-align: sub;    font-size: 0.75em; }}

/* ---- Block quote / block text ---- */
.block-text {{
  margin: 0.5em 3em;
  font-style: italic;
  color: #404040;
}}

/* ---- Source code ---- */
.source-code {{
  font-family: 'Courier New', monospace;
  font-size: 10pt;
  line-height: 1.0;
  white-space: pre-wrap;
  background: #f8f8f8;
  padding: 2px 4px;
}}

/* ---- Horizontal rule ---- */
.doc-line {{
  border: none;
  margin: 0;
}}
.doc-line.thick {{ border-top-width: 3px !important; }}
.doc-line.dashed {{ border-top-style: dashed !important; }}
.doc-line.dotted {{ border-top-style: dotted !important; }}

/* ---- Box ---- */
.doc-box {{
  width: 100%;
  padding: 10px 12px;
  margin-bottom: 8px;
}}

/* ---- Banner ---- */
.doc-banner {{ margin-bottom: 4px; }}

/* ---- Badge ---- */
.doc-badge {{ display: inline-block; padding: 1px 6px; margin: 4px 0; font-size: 0.85em; }}

/* ---- Data table ---- */
.data-table {{
  width: 100%;
  border-collapse: collapse;
  margin-bottom: 8px;
  table-layout: fixed;
  word-wrap: break-word;
}}
.data-table td, .data-table th {{
  padding: 3px 7px;
  vertical-align: top;
}}
.data-table.border td,
.data-table.border th {{
  border: 1px solid #BFBFBF;
}}
.data-table .header-row {{ background: #E8E8E8; font-weight: bold; }}
.data-table .stripe-row  {{ background: #F9F9F9; }}

/* ---- Layout table (cols) ---- */
.layout-table {{
  width: 100%;
  border-collapse: collapse;
  margin-bottom: 6px;
}}
.layout-table td {{ vertical-align: top; }}

/* ---- Image ---- */
.doc-image {{ margin: 4px 0; }}
.doc-image img {{ max-width: 100%; display: block; }}
.doc-image.center {{ text-align: center; }}
.doc-image.right  {{ text-align: right; }}

/* ---- Spacer ---- */
.doc-spacer {{ display: block; }}

/* ---- Shape ---- */
.doc-shape {{
  display: inline-flex;
  align-items: center;
  justify-content: center;
  text-align: center;
  padding: 6px 10px;
  box-sizing: border-box;
  overflow: hidden;
  vertical-align: middle;
  font-size: 0.85em;
}}

/* ---- Row of shapes ---- */
.shape-row {{
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  margin: 8px 0;
  flex-wrap: wrap;
}}
.shape-row .arrow-label {{ font-size: 1.1em; color: #555; }}

/* ---- Page break ---- */
.page-break {{ page-break-after: always; display: block; height: 0; }}

/* ---- Lists ---- */
ul, ol {{
  margin: 0 0 {para_after_pt}pt 0;
  padding-left: 2.5em;
  line-height: {line_height};
}}
li {{ margin-bottom: 2px; }}

/* ---- Hyperlinks ---- */
a {{ color: #0563C1; text-decoration: underline; }}
a:hover {{ color: #033E8F; }}

@media print {{
  body      {{ background: white; }}
  .doc-body {{ box-shadow: none; }}
  .page-break {{ page-break-after: always; }}
}}
""".strip()

    # ------------------------------------------------------------------
    # Item dispatcher
    # ------------------------------------------------------------------

    def _render_item(self, item) -> str:
        if isinstance(item, ParagraphModel):   return self._render_paragraph(item)
        if isinstance(item, LineModel):        return self._render_line(item)
        if isinstance(item, BoxModel):         return self._render_box(item)
        if isinstance(item, BannerModel):      return self._render_banner(item)
        if isinstance(item, BadgeModel):       return self._render_badge(item)
        if isinstance(item, DataTableModel):   return self._render_data_table(item)
        if isinstance(item, ImageModel):       return self._render_image(item)
        if isinstance(item, SpacerModel):      return self._render_spacer(item)
        if isinstance(item, ShapeModel):       return self._render_shape(item)
        if isinstance(item, RowModel):         return self._render_row(item)
        if isinstance(item, TableModel):       return self._render_table(item)
        if isinstance(item, PageBreakModel):   return '<div class="page-break"></div>\n'
        return ""

    # ------------------------------------------------------------------
    # Paragraph
    # ------------------------------------------------------------------

    def _render_paragraph(self, para: ParagraphModel) -> str:
        tag, css_class, extra_style = self._para_style_to_tag(para.style)

        styles: list[str] = list(extra_style)

        align_map = {"center": "center", "right": "right", "justify": "justify"}
        if para.align != "left" and para.align in align_map:
            styles.append(f"text-align:{align_map[para.align]}")
        if para.direction == "rtl":
            styles.append("direction:rtl")
        if para.indent_twips:
            styles.append(f"padding-left:{twip_to_px(para.indent_twips)}px")
        if para.space_before:
            styles.append(f"margin-top:{twip_to_pt(para.space_before)}pt")
        if para.space_after:
            styles.append(f"margin-bottom:{twip_to_pt(para.space_after)}pt")
        if para.shading:
            styles.append(f"background-color:#{para.shading}")
        if para.border_left:
            sz_px = para.border_left_sz * 0.5
            styles.append(
                f"border-left:{sz_px}px solid #{para.border_left};padding-left:8px"
            )

        style_attr  = f' style="{";".join(styles)}"' if styles else ""
        class_attr  = f' class="{css_class}"' if css_class else ""

        inner = self._render_runs(para.runs)

        # Empty paragraph → non-breaking space so it takes height
        if not inner.strip():
            inner = "&nbsp;"

        return f"<{tag}{class_attr}{style_attr}>{inner}</{tag}>\n"

    def _para_style_to_tag(self, style: str) -> tuple[str, str, list[str]]:
        """Return (html_tag, css_class, extra_style_list)."""
        if style.startswith("Heading"):
            lvl = style.replace("Heading", "").strip()
            if lvl.isdigit() and 1 <= int(lvl) <= 6:
                return f"h{lvl}", "", []
        if style == "BlockText":
            return "p", "block-text", []
        if style == "SourceCode":
            return "pre", "source-code", []
        if style in ("ListParagraph",):
            return "p", "list-paragraph", []
        return "p", "", []

    # ------------------------------------------------------------------
    # Runs
    # ------------------------------------------------------------------

    def _render_runs(self, runs: list[RunModel]) -> str:
        parts: list[str] = []
        i = 0
        while i < len(runs):
            run = runs[i]
            if run.hyperlink_url:
                url = run.hyperlink_url
                link_runs: list[RunModel] = []
                while i < len(runs) and runs[i].hyperlink_url == url:
                    link_runs.append(runs[i])
                    i += 1
                inner = "".join(self._render_run(r) for r in link_runs)
                safe_url = html.escape(url, quote=True)
                parts.append(f'<a href="{safe_url}">{inner}</a>')
            else:
                parts.append(self._render_run(run))
                i += 1
        return "".join(parts)

    def _render_run(self, run: RunModel) -> str:
        if run.field:
            return self._render_field_run(run)
        if not run.text:
            return ""

        styles: list[str] = []
        if run.font:
            styles.append(f"font-family:'{run.font}',sans-serif")
        if run.size_pt is not None:
            styles.append(f"font-size:{run.size_pt}pt")
        if run.bold:
            styles.append("font-weight:bold")
        if run.italic:
            styles.append("font-style:italic")
        if run.underline:
            styles.append("text-decoration:underline")
        if run.strike:
            styles.append("text-decoration:line-through")
        if run.underline and run.strike:
            styles.append("text-decoration:underline line-through")
        if run.color:
            styles.append(f"color:#{run.color}")
        if run.highlight:
            bg = HIGHLIGHT_CSS.get(run.highlight, f"#{run.highlight}")
            styles.append(f"background-color:{bg}")
        if run.shading:
            styles.append(f"background-color:#{run.shading}")
        if run.rtl:
            styles.append("direction:rtl")

        text = html.escape(run.text)

        if run.sup:
            text = f'<sup class="run-sup">{text}</sup>'
        elif run.sub:
            text = f'<sub class="run-sub">{text}</sub>'

        if not styles and not run.sup and not run.sub:
            return text

        style_attr = f' style="{";".join(styles)}"' if styles else ""
        if run.sup or run.sub:
            # Already wrapped above; attach styles to the wrapper
            tag_inner = html.escape(run.text)
            wrap = "sup" if run.sup else "sub"
            cls  = "run-sup" if run.sup else "run-sub"
            return f'<{wrap} class="{cls}"{style_attr}>{tag_inner}</{wrap}>'

        return f'<span{style_attr}>{text}</span>'

    def _render_field_run(self, run: RunModel) -> str:
        """Render PAGE / NUMPAGES fields as static placeholders."""
        field = (run.field or "").strip().upper()
        text = {"PAGE": "1", "NUMPAGES": "1", "DATE": _today()}.get(field, f"[{field}]")
        styles: list[str] = []
        if run.font:   styles.append(f"font-family:'{run.font}',sans-serif")
        if run.size_pt: styles.append(f"font-size:{run.size_pt}pt")
        if run.color:  styles.append(f"color:#{run.color}")
        style_attr = f' style="{";".join(styles)}"' if styles else ""
        return f'<span class="field-placeholder"{style_attr}>{html.escape(text)}</span>'

    # ------------------------------------------------------------------
    # Line
    # ------------------------------------------------------------------

    def _render_line(self, line: LineModel) -> str:
        styles = [f"border-top:1px solid #{line.color}"]
        if line.space_before:
            styles.append(f"margin-top:{twip_to_pt(line.space_before)}pt")
        if line.space_after:
            styles.append(f"margin-bottom:{twip_to_pt(line.space_after)}pt")

        classes = ["doc-line"]
        if line.thick:   classes.append("thick")
        if line.style == "dashed": classes.append("dashed")
        if line.style == "dotted": classes.append("dotted")

        cls   = " ".join(classes)
        style = ";".join(styles)
        return f'<hr class="{cls}" style="{style}">\n'

    # ------------------------------------------------------------------
    # Box
    # ------------------------------------------------------------------

    def _render_box(self, box: BoxModel) -> str:
        styles: list[str] = []
        if box.fill:
            styles.append(f"background-color:#{box.fill}")
        if box.stroke:
            styles.append(f"border:1px solid #{box.stroke}")
        style_attr = f' style="{";".join(styles)}"' if styles else ""

        inner_parts = [self._render_item(i) for i in (box.content or [])]
        inner = "".join(inner_parts) or "&nbsp;"
        return f'<div class="doc-box"{style_attr}>\n{inner}</div>\n'

    # ------------------------------------------------------------------
    # Banner
    # ------------------------------------------------------------------

    def _render_banner(self, banner: BannerModel) -> str:
        parts = [self._render_paragraph(p) for p in banner.paragraphs]
        return f'<div class="doc-banner">{"".join(parts)}</div>\n'

    # ------------------------------------------------------------------
    # Badge
    # ------------------------------------------------------------------

    def _render_badge(self, badge: BadgeModel) -> str:
        styles: list[str] = []
        if badge.color: styles.append(f"color:#{badge.color}")
        if badge.fill:  styles.append(f"background-color:#{badge.fill}")

        align_map = {"center": "center", "right": "right"}
        wrap_style = ""
        if badge.align in align_map:
            wrap_style = f' style="text-align:{align_map[badge.align]}"'

        style_attr = f' style="{";".join(styles)}"' if styles else ""
        text = html.escape(badge.text)
        return (
            f'<p{wrap_style}>'
            f'<span class="doc-badge"{style_attr}>{text}</span>'
            f'</p>\n'
        )

    # ------------------------------------------------------------------
    # Data table
    # ------------------------------------------------------------------

    def _render_data_table(self, table: DataTableModel) -> str:
        classes = ["data-table"]
        if table.border:
            classes.append("border")
        cls = " ".join(classes)

        rows_html: list[str] = []
        for row_idx, row in enumerate(table.rows):
            tr_class = ""
            if row.is_header:
                tr_class = ' class="header-row"'
            elif table.striped and row_idx % 2 == 0:
                tr_class = ' class="stripe-row"'

            cells_html: list[str] = []
            for cell in row.cells:
                tag = "th" if row.is_header else "td"
                colspan_attr = f' colspan="{cell.colspan}"' if cell.colspan > 1 else ""
                inner = "".join(self._render_item(i) for i in (cell.content or []))
                cells_html.append(f"<{tag}{colspan_attr}>{inner or '&nbsp;'}</{tag}>")

            rows_html.append(f"<tr{tr_class}>{''.join(cells_html)}</tr>")

        return f'<table class="{cls}">\n{"".join(rows_html)}\n</table>\n'

    # ------------------------------------------------------------------
    # Image
    # ------------------------------------------------------------------

    def _render_image(self, img: ImageModel) -> str:
        base = self._model.base_dir
        img_path = (base / img.src) if base else Path(img.src)

        align_map = {"center": "center", "right": "right"}
        div_class = "doc-image " + align_map.get(img.align, "")

        width_px  = emu_to_px(img.width_emu)
        height_px = emu_to_px(img.height_emu)

        if not img_path.exists():
            return (
                f'<div class="{div_class.strip()}">'
                f'<span style="color:red">[Image not found: {html.escape(str(img.src))}]</span>'
                f'</div>\n'
            )

        raw   = img_path.read_bytes()
        mime  = mimetypes.guess_type(str(img_path))[0] or "image/png"
        b64   = base64.b64encode(raw).decode()
        src   = f"data:{mime};base64,{b64}"

        return (
            f'<div class="{div_class.strip()}">'
            f'<img src="{src}" '
            f'width="{width_px:.0f}" height="{height_px:.0f}" '
            f'alt="{html.escape(img_path.name)}">'
            f'</div>\n'
        )

    # ------------------------------------------------------------------
    # Spacer
    # ------------------------------------------------------------------

    def _render_spacer(self, spacer: SpacerModel) -> str:
        px = twip_to_px(spacer.height_twips)
        return f'<div class="doc-spacer" style="height:{px}px"></div>\n'

    # ------------------------------------------------------------------
    # Shape
    # ------------------------------------------------------------------

    def _render_shape(self, shape: ShapeModel,
                      width_px: float | None = None,
                      height_px: float | None = None) -> str:
        if width_px is None:  width_px = 134.0
        if height_px is None:
            if shape.paragraphs:
                height_px = min(48 + 34 * len(shape.paragraphs), 384)
            else:
                height_px = 76.8

        fill   = f"#{shape.fill}" if shape.fill else "transparent"
        stroke = f"#{shape.stroke}" if shape.stroke else "transparent"
        stroke_w = "2px" if shape.stroke_thick else "1px"
        stroke_style = {"dashed": "dashed", "dotted": "dotted"}.get(shape.stroke_style, "solid")

        radius = SHAPE_PRESETS.get(shape.preset, "0")
        if shape.rounded:
            radius = "8px"

        extra_css = ""
        if shape.preset == "diamond":
            # CSS clip-path diamond
            extra_css = "clip-path:polygon(50% 0%,100% 50%,50% 100%,0% 50%);"
        elif shape.preset == "ellipse":
            radius = "50%"
        elif shape.preset == "chevron":
            extra_css = "clip-path:polygon(0 0,85% 0,100% 50%,85% 100%,0 100%,15% 50%);"

        inner = "".join(self._render_paragraph(p) for p in (shape.paragraphs or []))

        style = (
            f"width:{width_px:.0f}px;"
            f"height:{height_px:.0f}px;"
            f"background-color:{fill};"
            f"border:{stroke_w} {stroke_style} {stroke};"
            f"border-radius:{radius};"
            f"{extra_css}"
        )
        return f'<div class="doc-shape" style="{style}">{inner}</div>'

    # ------------------------------------------------------------------
    # Row of shapes
    # ------------------------------------------------------------------

    def _render_row(self, row: RowModel) -> str:
        parts: list[str] = ['<div class="shape-row">\n']
        shape_w, shape_h = 134.0, 57.6

        for i, shape in enumerate(row.shapes):
            parts.append(self._render_shape(shape, shape_w, shape_h))
            if i < len(row.arrows):
                label = row.arrows[i] or "→"
                parts.append(f'<span class="arrow-label">{html.escape(label)}</span>')

        parts.append("\n</div>\n")
        return "".join(parts)

    # ------------------------------------------------------------------
    # Table (layout columns)
    # ------------------------------------------------------------------

    def _render_table(self, table: TableModel) -> str:
        rows_html: list[str] = []
        for tbl_row in table.rows:
            total_pct = sum(c.width_pct for c in tbl_row.cells) or 100
            cells_html: list[str] = []
            for cell in tbl_row.cells:
                pct = round(cell.width_pct / total_pct * 100, 2)
                inner = "".join(self._render_item(i) for i in (cell.content or []))
                cells_html.append(
                    f'<td style="width:{pct}%;vertical-align:top">'
                    f'{inner or "&nbsp;"}</td>'
                )
            rows_html.append(f"<tr>{''.join(cells_html)}</tr>")

        border_attr = ""
        if table.border:
            border_attr = ' style="border:1px solid #BFBFBF;border-collapse:collapse"'

        return (
            f'<table class="layout-table"{border_attr}>\n'
            f'{"".join(rows_html)}\n'
            f'</table>\n'
        )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _today() -> str:
    from datetime import date
    return date.today().strftime("%B %d, %Y")