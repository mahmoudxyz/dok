"""
dok.docx_writer
~~~~~~~~~~~~~~~
Converts a DocxModel to a .docx file (ZIP + XML).

Supports: paragraphs, shapes, tables, lists (numbering.xml),
images (embedded media), hyperlinks, header/footer, page numbers.
"""

from __future__ import annotations
import io
import zipfile
from pathlib import Path

from .constants import (
    inch_to_emu, inch_to_twip,
    MARGIN_PRESETS, PAPER_SIZES, content_width_twip,
    HYPERLINK_COLOR, DEFAULT_BORDER_COLOR,
)
from .docx_packaging import (
    W_NS as _W_NS, R_NS as _R_NS, A_NS as _A_NS,
    WPS_URI as _WPS_URI, PIC_URI as _PIC_URI,
    DOCUMENT_NS as _NS, PACKAGE_RELS as _RELS, SETTINGS_XML as _SETTINGS,
    build_numbering_xml, build_doc_rels, build_content_types, build_settings_xml,
)
from .docx_styles import build_styles_xml
from .models import (
    DocxModel, ParagraphModel, RunModel, ShapeModel,
    RowModel, TableModel, TableCellModel,
    PageBreakModel, SectionModel,
    LineModel, BoxModel,
    DataTableModel, DataTableRowModel, DataTableCellModel,
    ImageModel, SpacerModel, HeaderModel, FooterModel,
    TocModel, TocEntry,
    CheckboxModel, TextInputModel, DropdownModel,
    ToggleModel, FrameModel,
)
from .writer_utils import group_runs_by_hyperlink
from .xml_writer import XmlWriter
from xml.sax.saxutils import escape as _xml_escape


# ---------------------------------------------------------------------------
# Writer
# ---------------------------------------------------------------------------

class DocxWriter:

    def __init__(self, model: DocxModel) -> None:
        self._model    = model
        self._shape_id = 0
        self._rel_id   = 2   # rId1=styles, rId2=settings
        # Accumulated during writing
        self._image_entries: list[tuple[str, str, bytes]] = []  # (relId, filename, data)
        self._hyperlink_rels: list[tuple[str, str]] = []        # (relId, url)
        self._hyperlink_cache: dict[str, str] = {}              # url → relId
        self._image_counter = 0
        self._bookmark_counter = 0

    def _next_bookmark_id(self) -> int:
        self._bookmark_counter += 1
        return self._bookmark_counter

    def _next_rel_id(self) -> str:
        self._rel_id += 1
        return f"rId{self._rel_id}"

    def _next_shape_id(self) -> int:
        self._shape_id += 1
        return self._shape_id

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def write(self, dest: str | Path | io.BytesIO) -> None:
        # Build main document (accumulates rels for images/hyperlinks)
        doc_xml = self._build_document_xml()

        # Build optional parts
        header_xml = self._build_header_xml() if self._model.header else None
        footer_xml = self._build_footer_xml() if self._model.footer else None
        numbering_xml = build_numbering_xml(
            custom_markers=self._model.custom_markers
        ) if self._model.has_lists else None

        # Assign rel IDs for header/footer/numbering
        header_rel_id = footer_rel_id = numbering_rel_id = None
        if numbering_xml:
            numbering_rel_id = self._next_rel_id()
        if header_xml:
            header_rel_id = self._next_rel_id()
        if footer_xml:
            footer_rel_id = self._next_rel_id()

        # Build dynamic rels and content types
        doc_rels_xml    = build_doc_rels(
            self._image_entries, self._hyperlink_rels,
            header_rel_id, footer_rel_id, numbering_rel_id)
        content_types   = build_content_types(
            self._image_entries,
            has_header=header_xml is not None,
            has_footer=footer_xml is not None,
            has_numbering=numbering_xml is not None)

        # If we have header/footer, rebuild doc xml with correct rel IDs in sectPr
        if header_rel_id or footer_rel_id:
            doc_xml = self._build_document_xml(header_rel_id, footer_rel_id)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",          content_types)
            zf.writestr("_rels/.rels",                  _RELS)
            zf.writestr("word/document.xml",            doc_xml)
            zf.writestr("word/styles.xml",              self._build_styles_xml())
            settings = build_settings_xml(hyphenate=self._model.hyphenate)
            zf.writestr("word/settings.xml",            settings)
            zf.writestr("word/_rels/document.xml.rels", doc_rels_xml)

            if numbering_xml:
                zf.writestr("word/numbering.xml", numbering_xml)
            if header_xml:
                zf.writestr("word/header1.xml", header_xml)
            if footer_xml:
                zf.writestr("word/footer1.xml", footer_xml)

            for _, filename, data in self._image_entries:
                zf.writestr(f"word/media/{filename}", data)

        data = buf.getvalue()
        if isinstance(dest, (str, Path)):
            Path(dest).write_bytes(data)
        else:
            dest.write(data)

    # ------------------------------------------------------------------
    # document.xml
    # ------------------------------------------------------------------

    def _build_document_xml(self, header_rel_id: str | None = None,
                            footer_rel_id: str | None = None) -> str:
        # Reset rels on rebuild
        self._image_entries.clear()
        self._hyperlink_rels.clear()
        self._hyperlink_cache.clear()
        self._image_counter = 0
        self._shape_id = 0
        self._rel_id = 2

        w = XmlWriter()
        w.declaration()
        w.open("w:document", _NS)
        w.open("w:body")

        for item in self._model.content:
            self._write_item(w, item)

        self._write_sect_pr(w, header_rel_id, footer_rel_id)
        w.close("w:body")
        w.close("w:document")
        return w.getvalue()

    _DISPATCH: dict[type, str] = {
        ParagraphModel: "_write_paragraph",
        LineModel:      "_write_line",
        BoxModel:       "_write_box",
        DataTableModel: "_write_data_table",
        ImageModel:     "_write_image",
        SpacerModel:    "_write_spacer",
        ShapeModel:     "_write_shape",
        RowModel:       "_write_row",
        TableModel:     "_write_table",
        PageBreakModel: "_write_page_break",
        TocModel:       "_write_toc",
        FrameModel:     "_write_frame",
        ToggleModel:    "_write_toggle",
        CheckboxModel:  "_write_checkbox",
        TextInputModel: "_write_text_input",
        DropdownModel:  "_write_dropdown",
    }

    def _write_item(self, w: XmlWriter, item) -> None:
        handler = self._DISPATCH.get(type(item))
        if handler:
            getattr(self, handler)(w, item)

    def _content_width(self) -> int:
        """Content width in twips for the current section."""
        section = self._model.sections[-1] if self._model.sections else SectionModel()
        return content_width_twip(section.paper,
                                  margins=section.resolved_margins(MARGIN_PRESETS))

    # ------------------------------------------------------------------
    # Paragraph
    # ------------------------------------------------------------------

    def _write_paragraph(self, w: XmlWriter, para: ParagraphModel) -> None:
        # SourceCode with newlines: split into one <w:p> per line
        if para.style == "SourceCode" and any("\n" in r.text for r in para.runs):
            self._write_code_lines(w, para)
            return

        self._write_single_paragraph(w, para)

    def _write_single_paragraph(self, w: XmlWriter, para: ParagraphModel) -> None:
        w.open("w:p")
        # Bookmark start
        if para.bookmark:
            bm_id = str(self._next_bookmark_id())
            w.tag("w:bookmarkStart", {"w:id": bm_id, "w:name": para.bookmark})
        self._write_para_props(w, para)
        self._write_runs_with_hyperlinks(w, para.runs)
        # Bookmark end
        if para.bookmark:
            w.tag("w:bookmarkEnd", {"w:id": bm_id})
        w.close("w:p")

    def _write_para_props(self, w: XmlWriter, para: ParagraphModel,
                          space_before: int | None = None,
                          space_after: int | None = None) -> None:
        w.open("w:pPr")

        w.tag("w:pStyle", {"w:val": para.style})

        # Widow/orphan control
        if self._model.widow_orphan > 0:
            w.tag("w:widowControl", {})

        if para.align != "left":
            jc = {"center": "center", "right": "right",
                  "justify": "both"}.get(para.align, "left")
            w.tag("w:jc", {"w:val": jc})

        if para.direction == "rtl":
            w.tag("w:bidi", {})

        if para.indent_twips > 0:
            w.tag("w:ind", {"w:left": str(para.indent_twips)})

        # Always write spacing to override style defaults precisely
        sb = space_before if space_before is not None else para.space_before
        sa = space_after if space_after is not None else para.space_after
        spacing: dict = {"w:before": str(sb), "w:after": str(sa)}
        if para.line_spacing:
            spacing["w:line"] = str(para.line_spacing)
            spacing["w:lineRule"] = "auto"
        w.tag("w:spacing", spacing)

        if para.shading:
            w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": para.shading})

        if para.border_left:
            w.open("w:pBdr")
            w.tag("w:left", {
                "w:val": "single", "w:sz": str(para.border_left_sz),
                "w:space": "8", "w:color": para.border_left,
            })
            w.close("w:pBdr")

        # List numbering
        if para.num_id:
            w.open("w:numPr")
            w.tag("w:ilvl", {"w:val": str(para.num_ilvl)})
            w.tag("w:numId", {"w:val": str(para.num_id)})
            w.close("w:numPr")

        if para.direction == "rtl":
            w.open("w:rPr")
            w.tag("w:rtl", {})
            w.close("w:rPr")

        w.close("w:pPr")

    def _write_code_lines(self, w: XmlWriter, para: ParagraphModel) -> None:
        """Split a SourceCode paragraph with newlines into one <w:p> per line."""
        # Concatenate all run text
        full_text = "".join(r.text for r in para.runs)
        lines = full_text.split("\n")
        # Take run formatting from first run
        template_run = para.runs[0] if para.runs else RunModel(text="")

        for i, line_text in enumerate(lines):
            w.open("w:p")
            sb = para.space_before if i == 0 else 0
            sa = para.space_after if i == len(lines) - 1 else 0
            self._write_para_props(w, para, space_before=sb, space_after=sa)
            text = line_text if line_text else " "
            run = RunModel(
                text=text, bold=template_run.bold, italic=template_run.italic,
                underline=template_run.underline, color=template_run.color,
                size_pt=template_run.size_pt, font=template_run.font,
                rtl=template_run.rtl,
            )
            self._write_run(w, run)
            w.close("w:p")

    def _write_runs_with_hyperlinks(self, w: XmlWriter, runs: list[RunModel]) -> None:
        """Write runs, grouping consecutive hyperlinked runs into <w:hyperlink>."""
        for url, group in group_runs_by_hyperlink(runs):
            if url:
                if url.startswith("#"):
                    # Internal bookmark reference
                    w.open("w:hyperlink", {"w:anchor": url[1:]})
                else:
                    rel_id = self._get_hyperlink_rel(url)
                    w.open("w:hyperlink", {"r:id": rel_id})
                for r in group:
                    self._write_run(w, r)
                w.close("w:hyperlink")
            else:
                for r in group:
                    self._write_run(w, r)

    def _get_hyperlink_rel(self, url: str) -> str:
        if url in self._hyperlink_cache:
            return self._hyperlink_cache[url]
        rel_id = self._next_rel_id()
        self._hyperlink_rels.append((rel_id, url))
        self._hyperlink_cache[url] = rel_id
        return rel_id

    def _write_run(self, w: XmlWriter, run: RunModel) -> None:
        # Field run (page number etc.)
        if run.field:
            self._write_field_run(w, run)
            return

        if not run.text:
            return

        w.open("w:r")
        w.open("w:rPr")

        if run.font:
            w.tag("w:rFonts", {"w:ascii": run.font, "w:hAnsi": run.font, "w:cs": run.font})
        if run.size_pt is not None:
            size = run.size_pt * 2
            w.tag("w:sz", {"w:val": str(size)})
            w.tag("w:szCs", {"w:val": str(size)})

        if run.bold:      w.tag("w:b", {})
        if run.italic:    w.tag("w:i", {})
        if run.underline: w.tag("w:u", {"w:val": "single"})
        if run.strike:    w.tag("w:strike", {})

        if run.color:     w.tag("w:color", {"w:val": run.color})
        if run.highlight: w.tag("w:highlight", {"w:val": run.highlight})
        if run.shading:
            w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": run.shading})

        va = "superscript" if run.sup else "subscript" if run.sub else None
        if va: w.tag("w:vertAlign", {"w:val": va})
        if run.rtl: w.tag("w:rtl", {})

        # Kerning: apply to all runs when enabled
        if self._model.kerning:
            kern_pt = (run.size_pt or self._model.default_size_pt) * 2
            w.tag("w:kern", {"w:val": str(kern_pt)})

        w.close("w:rPr")
        w.w_t(run.text)
        w.close("w:r")

    def _write_field_run(self, w: XmlWriter, run: RunModel) -> None:
        """Emit a Word field code (PAGE, NUMPAGES, etc.)."""
        w.open("w:r")
        if run.color or run.size_pt or run.font:
            w.open("w:rPr")
            if run.font:
                w.tag("w:rFonts", {"w:ascii": run.font, "w:hAnsi": run.font, "w:cs": run.font})
            if run.size_pt:
                sz = run.size_pt * 2
                w.tag("w:sz", {"w:val": str(sz)})
                w.tag("w:szCs", {"w:val": str(sz)})
            if run.color:
                w.tag("w:color", {"w:val": run.color})
            w.close("w:rPr")
        w.tag("w:fldChar", {"w:fldCharType": "begin"})
        w.close("w:r")

        w.open("w:r")
        w.open("w:instrText", {"xml:space": "preserve"})
        w.text(f" {run.field} ")
        w.close("w:instrText")
        w.close("w:r")

        w.open("w:r")
        w.tag("w:fldChar", {"w:fldCharType": "separate"})
        w.close("w:r")

        w.open("w:r")
        w.w_t("1")  # placeholder
        w.close("w:r")

        w.open("w:r")
        w.tag("w:fldChar", {"w:fldCharType": "end"})
        w.close("w:r")

    # ------------------------------------------------------------------
    # Line
    # ------------------------------------------------------------------

    def _write_line(self, w: XmlWriter, line: LineModel) -> None:
        border_sz = "12" if line.thick else "4"
        border_val = {"dashed": "dashed", "dotted": "dotted"}.get(line.style, "single")
        w.open("w:p")
        w.open("w:pPr")
        spacing: dict = {}
        if line.space_before: spacing["w:before"] = str(line.space_before)
        if line.space_after:  spacing["w:after"]  = str(line.space_after)
        if spacing: w.tag("w:spacing", spacing)
        w.open("w:pBdr")
        w.tag("w:bottom", {
            "w:val": border_val, "w:sz": border_sz,
            "w:space": "1", "w:color": line.color,
        })
        w.close("w:pBdr")
        w.close("w:pPr")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Box
    # ------------------------------------------------------------------

    def _write_box(self, w: XmlWriter, box: BoxModel) -> None:
        # Badge mode: inline text
        if box.inline:
            self._write_badge_inline(w, box)
            return

        # All boxes use the same table-cell rendering path
        cw = box.max_width_twips if box.max_width_twips else self._content_width()
        box_w = min((cw * box.width_pct // 100) if box.width_pct else cw, cw)

        w.open("w:tbl")
        w.open("w:tblPr")
        w.tag("w:tblW", {"w:w": str(box_w), "w:type": "dxa"})
        w.open("w:tblCellMar")
        w.tag("w:top",    {"w:w": "100", "w:type": "dxa"})
        w.tag("w:left",   {"w:w": "160", "w:type": "dxa"})
        w.tag("w:bottom", {"w:w": "100", "w:type": "dxa"})
        w.tag("w:right",  {"w:w": "160", "w:type": "dxa"})
        w.close("w:tblCellMar")
        w.open("w:tblBorders")
        stroke_color = box.stroke or DEFAULT_BORDER_COLOR
        border_sz = str(box.border_width * 8)  # pt → OOXML eighth-points
        side_enabled = {
            "top": box.border_top, "bottom": box.border_bottom,
            "right": box.border_right, "left": box.border_left,
        }
        if box.stroke:
            for side in ("top", "bottom", "right"):
                if side_enabled[side]:
                    w.tag(f"w:{side}", {"w:val": "single", "w:sz": border_sz,
                                        "w:space": "0", "w:color": stroke_color})
                else:
                    w.tag(f"w:{side}", {"w:val": "none", "w:sz": "0",
                                        "w:space": "0", "w:color": "auto"})
            # Left border: thick accent or normal stroke
            if side_enabled["left"]:
                left_color = box.accent or stroke_color
                left_sz = "24" if box.accent else border_sz
                w.tag("w:left", {"w:val": "single", "w:sz": left_sz,
                                  "w:space": "0", "w:color": left_color})
            else:
                w.tag("w:left", {"w:val": "none", "w:sz": "0",
                                  "w:space": "0", "w:color": "auto"})
        else:
            if box.accent and side_enabled["left"]:
                w.tag("w:left", {"w:val": "single", "w:sz": "24",
                                  "w:space": "0", "w:color": box.accent})
            else:
                w.tag("w:left", {"w:val": "none", "w:sz": "0",
                                  "w:space": "0", "w:color": "auto"})
            for side in ("top", "bottom", "right"):
                w.tag(f"w:{side}", {"w:val": "none", "w:sz": "0",
                                    "w:space": "0", "w:color": "auto"})
        for side in ("insideH", "insideV"):
            w.tag(f"w:{side}", {"w:val": "none", "w:sz": "0",
                                "w:space": "0", "w:color": "auto"})
        w.close("w:tblBorders")
        w.close("w:tblPr")

        w.open("w:tr")
        if box.height_pt:
            w.open("w:trPr")
            w.tag("w:trHeight", {"w:val": str(box.height_pt * 20), "w:hRule": "atLeast"})
            w.close("w:trPr")
        w.open("w:tc")
        w.open("w:tcPr")
        w.tag("w:tcW", {"w:w": str(box_w), "w:type": "dxa"})
        if box.fill:
            w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": box.fill})
        w.close("w:tcPr")

        if not box.content:
            w.raw("<w:p/>")
        else:
            for item in box.content:
                self._write_item(w, item)
        w.close("w:tc")
        w.close("w:tr")
        w.close("w:tbl")

        # Minimal trailing paragraph (required by OOXML after tables)
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:spacing", {"w:before": "0", "w:after": "0"}); w.close("w:pPr")
        w.close("w:p")

    def _write_badge_inline(self, w: XmlWriter, box: BoxModel) -> None:
        w.open("w:p")
        w.open("w:pPr")
        if box.align != "left":
            jc = {"center": "center", "right": "right"}.get(box.align, "left")
            w.tag("w:jc", {"w:val": jc})
        w.tag("w:spacing", {"w:after": "160"})
        w.close("w:pPr")

        padded = f"  {box.text or ''}  "
        w.open("w:r")
        w.open("w:rPr")
        if box.color: w.tag("w:color", {"w:val": box.color})
        if box.fill:  w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": box.fill})
        w.close("w:rPr")
        w.w_t(padded)
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Data Table
    # ------------------------------------------------------------------

    def _write_data_table(self, w: XmlWriter, table: DataTableModel) -> None:
        cw = self._content_width()
        is_rtl = table.direction == "rtl"

        # Determine number of columns from widest row
        max_cols = max((sum(c.colspan for c in r.cells) for r in table.rows), default=1)

        w.open("w:tbl")
        w.open("w:tblPr")
        w.tag("w:tblW", {"w:w": str(cw), "w:type": "dxa"})
        w.tag("w:tblLayout", {"w:type": "fixed"})

        # RTL column order
        if is_rtl:
            w.tag("w:bidiVisual", {})

        if table.border:
            w.open("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                w.tag(f"w:{side}", {"w:val": "single", "w:sz": "4",
                                    "w:space": "0", "w:color": DEFAULT_BORDER_COLOR})
            w.close("w:tblBorders")

        w.open("w:tblCellMar")
        w.tag("w:top",    {"w:w": "40",  "w:type": "dxa"})
        w.tag("w:left",   {"w:w": "100", "w:type": "dxa"})
        w.tag("w:bottom", {"w:w": "40",  "w:type": "dxa"})
        w.tag("w:right",  {"w:w": "100", "w:type": "dxa"})
        w.close("w:tblCellMar")

        w.close("w:tblPr")

        # Grid columns — use pre-calculated proportional widths if available
        if table.col_widths and len(table.col_widths) == max_cols:
            col_widths = [cw * pct // 100 for pct in table.col_widths]
        else:
            col_widths = [cw // max_cols] * max_cols
        w.open("w:tblGrid")
        for cw_i in col_widths:
            w.tag("w:gridCol", {"w:w": str(cw_i)})
        w.close("w:tblGrid")

        for row_idx, row in enumerate(table.rows):
            w.open("w:tr")
            col_idx = 0
            for cell in row.cells:
                cell_w = sum(col_widths[col_idx:col_idx + cell.colspan])
                col_idx += cell.colspan
                w.open("w:tc")
                w.open("w:tcPr")
                w.tag("w:tcW", {"w:w": str(cell_w), "w:type": "dxa"})
                if cell.colspan > 1:
                    w.tag("w:gridSpan", {"w:val": str(cell.colspan)})

                # Cell background: explicit fill > header shading > striped
                if cell.fill:
                    w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": cell.fill})
                elif row.is_header:
                    w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": "E8E8E8"})
                elif table.striped and not row.is_header and row_idx % 2 == 0:
                    w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": "F9F9F9"})

                w.close("w:tcPr")

                if not cell.content:
                    w.raw("<w:p/>")
                else:
                    for item in cell.content:
                        self._write_item(w, item)

                w.close("w:tc")
            w.close("w:tr")

        w.close("w:tbl")

        # Minimal trailing paragraph (required by OOXML after tables)
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:spacing", {"w:before": "0", "w:after": "0"}); w.close("w:pPr")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Image
    # ------------------------------------------------------------------

    def _write_image(self, w: XmlWriter, img: ImageModel) -> None:
        # Read the image file
        base = self._model.base_dir
        img_path = (base / img.src) if base else Path(img.src)

        if not img_path.exists():
            # Fallback: write a placeholder paragraph
            w.open("w:p")
            w.open("w:r")
            w.open("w:rPr"); w.tag("w:color", {"w:val": "FF0000"}); w.close("w:rPr")
            w.w_t(f"[Image not found: {img.src}]")
            w.close("w:r")
            w.close("w:p")
            return

        image_data = img_path.read_bytes()
        self._image_counter += 1
        ext = img_path.suffix or ".png"
        filename = f"image{self._image_counter}{ext}"
        rel_id = self._next_rel_id()
        self._image_entries.append((rel_id, filename, image_data))

        sid = self._next_shape_id()

        w.open("w:p")
        w.open("w:pPr")
        if img.align != "left":
            jc = {"center": "center", "right": "right"}.get(img.align, "left")
            w.tag("w:jc", {"w:val": jc})
        w.close("w:pPr")

        w.open("w:r")
        w.open("w:drawing")
        w.open("wp:inline", {"distT": "0", "distB": "0", "distL": "0", "distR": "0"})
        w.tag("wp:extent", {"cx": str(img.width_emu), "cy": str(img.height_emu)})
        w.tag("wp:effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
        w.tag("wp:docPr", {"id": str(sid), "name": f"Picture{sid}"})

        w.open("a:graphic", {"xmlns:a": _A_NS})
        w.open("a:graphicData", {"uri": _PIC_URI})
        w.open("pic:pic", {"xmlns:pic": _PIC_URI})

        w.open("pic:nvPicPr")
        w.tag("pic:cNvPr", {"id": str(sid), "name": f"Picture{sid}"})
        w.open("pic:cNvPicPr")
        w.tag("a:picLocks", {"noChangeAspect": "1"})
        w.close("pic:cNvPicPr")
        w.close("pic:nvPicPr")

        w.open("pic:blipFill")
        w.tag("a:blip", {"r:embed": rel_id})
        w.open("a:stretch")
        w.tag("a:fillRect")
        w.close("a:stretch")
        w.close("pic:blipFill")

        w.open("pic:spPr")
        w.open("a:xfrm")
        w.tag("a:off", {"x": "0", "y": "0"})
        w.tag("a:ext", {"cx": str(img.width_emu), "cy": str(img.height_emu)})
        w.close("a:xfrm")
        w.open("a:prstGeom", {"prst": "rect"})
        w.tag("a:avLst")
        w.close("a:prstGeom")
        w.close("pic:spPr")

        w.close("pic:pic")
        w.close("a:graphicData")
        w.close("a:graphic")
        w.close("wp:inline")
        w.close("w:drawing")
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Spacer
    # ------------------------------------------------------------------

    def _write_spacer(self, w: XmlWriter, spacer: SpacerModel) -> None:
        w.open("w:p")
        w.open("w:pPr")
        w.tag("w:spacing", {"w:before": str(spacer.height_twips), "w:after": "0"})
        # Small font so the empty paragraph is minimal
        w.close("w:pPr")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Drawing Shape (circle, diamond, chevron)
    # ------------------------------------------------------------------

    def _write_shape(self, w: XmlWriter, shape: ShapeModel,
                     width_emu: int | None = None,
                     height_emu: int | None = None) -> None:
        if width_emu is None:  width_emu = inch_to_emu(1.4)
        if height_emu is None:
            if shape.paragraphs:
                estimated = 0.5 + 0.35 * len(shape.paragraphs)
                height_emu = inch_to_emu(min(estimated, 4.0))
            else:
                height_emu = inch_to_emu(0.8)

        sid = self._next_shape_id()
        w.open("w:p")
        w.open("w:pPr"); w.close("w:pPr")
        w.open("w:r")
        w.open("w:rPr"); w.close("w:rPr")
        w.open("w:drawing")
        w.open("wp:inline", {"distT": "0", "distB": "0", "distL": "0", "distR": "0"})
        w.tag("wp:extent", {"cx": str(width_emu), "cy": str(height_emu)})
        w.tag("wp:effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
        w.tag("wp:docPr", {"id": str(sid), "name": f"Shape{sid}"})

        w.open("a:graphic", {"xmlns:a": _A_NS})
        w.open("a:graphicData", {"uri": _WPS_URI})
        w.open("wps:wsp")
        w.open("wps:spPr")
        preset = "roundRect" if shape.rounded else shape.preset
        w.open("a:prstGeom", {"prst": preset})
        w.tag("a:avLst")
        w.close("a:prstGeom")
        self._write_fill(w, shape.fill)
        self._write_stroke(w, shape)
        w.close("wps:spPr")

        if shape.paragraphs:
            w.open("wps:txbx")
            w.open("w:txbxContent")
            for para in shape.paragraphs:
                self._write_paragraph(w, para)
            w.close("w:txbxContent")
            w.close("wps:txbx")

        w.open("wps:bodyPr", {"wrap": "square",
                               "lIns": "91440", "tIns": "45720",
                               "rIns": "91440", "bIns": "45720"})
        w.close("wps:bodyPr")
        w.close("wps:wsp")
        w.close("a:graphicData")
        w.close("a:graphic")
        w.close("wp:inline")
        w.close("w:drawing")
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Row
    # ------------------------------------------------------------------

    def _write_row(self, w: XmlWriter, row: RowModel) -> None:
        # Check if row contains only drawing shapes (old behavior: inline shapes in a paragraph)
        all_shapes = all(isinstance(item, ShapeModel) for item in row.items)

        if all_shapes and row.items:
            self._write_row_shapes(w, row)
        else:
            self._write_row_mixed(w, row)

    def _write_row_shapes(self, w: XmlWriter, row: RowModel) -> None:
        """Row of drawing shapes rendered as inline shapes in a centered paragraph."""
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:jc", {"w:val": "center"}); w.close("w:pPr")
        shape_w = inch_to_emu(1.4); shape_h = inch_to_emu(0.6)

        for i, shape in enumerate(row.items):
            self._write_shape_inline(w, shape, shape_w, shape_h)

            if i < len(row.arrows):
                label = row.arrows[i]
                arrow_text = f"  {label}  " if label else "  \u2192  "
                w.open("w:r"); w.open("w:rPr"); w.close("w:rPr")
                w.w_t(arrow_text); w.close("w:r")
        w.close("w:p")

    def _write_row_mixed(self, w: XmlWriter, row: RowModel) -> None:
        """Row of mixed content rendered as equal-width table columns."""
        if not row.items:
            return
        cw = self._content_width()
        col_w = cw // len(row.items)

        w.open("w:tbl")
        w.open("w:tblPr")
        w.tag("w:tblW", {"w:w": str(cw), "w:type": "dxa"})
        w.tag("w:tblLayout", {"w:type": "fixed"})
        w.close("w:tblPr")
        w.open("w:tblGrid")
        for _ in row.items:
            w.tag("w:gridCol", {"w:w": str(col_w)})
        w.close("w:tblGrid")

        w.open("w:tr")
        for item in row.items:
            w.open("w:tc")
            w.open("w:tcPr")
            w.tag("w:tcW", {"w:w": str(col_w), "w:type": "dxa"})
            w.close("w:tcPr")
            self._write_item(w, item)
            # Ensure at least one paragraph in cell
            if not isinstance(item, ParagraphModel):
                pass  # _write_item handles its own paragraphs
            w.close("w:tc")
        w.close("w:tr")
        w.close("w:tbl")

    def _write_shape_inline(self, w: XmlWriter, shape: ShapeModel,
                            shape_w: int, shape_h: int) -> None:
        """Write a single drawing shape as an inline element."""
        sid = self._next_shape_id()
        w.open("w:r")
        w.open("w:rPr"); w.close("w:rPr")
        w.open("w:drawing")
        w.open("wp:inline", {"distT": "0", "distB": "0", "distL": "57150", "distR": "57150"})
        w.tag("wp:extent", {"cx": str(shape_w), "cy": str(shape_h)})
        w.tag("wp:effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
        w.tag("wp:docPr", {"id": str(sid), "name": f"Shape{sid}"})
        w.open("a:graphic", {"xmlns:a": _A_NS})
        w.open("a:graphicData", {"uri": _WPS_URI})
        w.open("wps:wsp")
        w.open("wps:spPr")
        preset = "roundRect" if shape.rounded else shape.preset
        w.open("a:prstGeom", {"prst": preset})
        w.tag("a:avLst"); w.close("a:prstGeom")
        self._write_fill(w, shape.fill)
        self._write_stroke(w, shape)
        w.close("wps:spPr")
        if shape.paragraphs:
            w.open("wps:txbx"); w.open("w:txbxContent")
            for para in shape.paragraphs:
                self._write_paragraph(w, para)
            w.close("w:txbxContent"); w.close("wps:txbx")
        w.open("wps:bodyPr", {"wrap": "square",
                               "lIns": "91440", "tIns": "45720",
                               "rIns": "91440", "bIns": "45720"})
        w.close("wps:bodyPr")
        w.close("wps:wsp"); w.close("a:graphicData"); w.close("a:graphic")
        w.close("wp:inline"); w.close("w:drawing"); w.close("w:r")

    # ------------------------------------------------------------------
    # Fill / Stroke helpers
    # ------------------------------------------------------------------

    def _write_fill(self, w: XmlWriter, hex_color: str | None) -> None:
        if hex_color is None:
            w.tag("a:noFill")
        else:
            w.open("a:solidFill"); w.tag("a:srgbClr", {"val": hex_color}); w.close("a:solidFill")

    def _write_stroke(self, w: XmlWriter, shape: ShapeModel) -> None:
        thickness = "25400" if shape.stroke_thick else "9525"
        if shape.stroke is None:
            w.open("a:ln", {"w": thickness}); w.tag("a:noFill"); w.close("a:ln")
            return
        w.open("a:ln", {"w": thickness})
        w.open("a:solidFill"); w.tag("a:srgbClr", {"val": shape.stroke}); w.close("a:solidFill")
        dash = {"dashed": "dash", "dotted": "dot"}.get(shape.stroke_style, "solid")
        w.tag("a:prstDash", {"val": dash})
        w.close("a:ln")

    # ------------------------------------------------------------------
    # Table (cols layout)
    # ------------------------------------------------------------------

    def _write_table(self, w: XmlWriter, table: TableModel) -> None:
        cw = self._content_width()
        n_cells = max(len(r.cells) for r in table.rows) if table.rows else 1
        total_gap = table.gap_twips * max(0, n_cells - 1)
        usable_w = cw - total_gap

        w.open("w:tbl")
        w.open("w:tblPr")
        w.tag("w:tblW", {"w:w": str(cw), "w:type": "dxa"})
        if not table.border:
            w.open("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                w.tag(f"w:{side}", {"w:val": "none", "w:sz": "0",
                                    "w:space": "0", "w:color": "auto"})
            w.close("w:tblBorders")
        # Cell margins (gap is implemented as cell margins on sides)
        if table.gap_twips or table.cell_padding_twips:
            half_gap = table.gap_twips // 2
            pad = table.cell_padding_twips
            w.open("w:tblCellMar")
            w.tag("w:top",    {"w:w": str(pad), "w:type": "dxa"})
            w.tag("w:left",   {"w:w": str(half_gap + pad), "w:type": "dxa"})
            w.tag("w:bottom", {"w:w": str(pad), "w:type": "dxa"})
            w.tag("w:right",  {"w:w": str(half_gap + pad), "w:type": "dxa"})
            w.close("w:tblCellMar")
        if table.fill:
            w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": table.fill})
        w.close("w:tblPr")

        for tbl_row in table.rows:
            w.open("w:tr")
            total_pct = sum(c.width_pct for c in tbl_row.cells) or 100
            for cell in tbl_row.cells:
                cell_w = int(usable_w * cell.width_pct / total_pct)
                w.open("w:tc")
                w.open("w:tcPr")
                w.tag("w:tcW", {"w:w": str(cell_w), "w:type": "dxa"})
                if cell.fill:
                    w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": cell.fill})
                if cell.padding_twips:
                    w.open("w:tcMar")
                    p = str(cell.padding_twips)
                    for s in ("top", "left", "bottom", "right"):
                        w.tag(f"w:{s}", {"w:w": p, "w:type": "dxa"})
                    w.close("w:tcMar")
                w.close("w:tcPr")
                if not cell.content:
                    w.raw("<w:p/>")
                else:
                    for item in cell.content:
                        self._write_item(w, item)
                w.close("w:tc")
            w.close("w:tr")

        w.close("w:tbl")
        # Minimal trailing paragraph (required by OOXML after tables)
        w.open("w:p"); w.open("w:pPr")
        w.tag("w:spacing", {"w:before": "0", "w:after": "0"})
        w.close("w:pPr"); w.close("w:p")

    # ------------------------------------------------------------------
    # Table of Contents
    # ------------------------------------------------------------------

    def _write_toc(self, w: XmlWriter, toc: TocModel) -> None:
        # Title paragraph
        w.open("w:p")
        w.open("w:pPr")
        w.tag("w:pStyle", {"w:val": "Heading1"})
        w.tag("w:spacing", {"w:before": "0", "w:after": "200"})
        w.close("w:pPr")
        w.open("w:r")
        w.open("w:rPr"); w.tag("w:b", {}); w.close("w:rPr")
        w.raw(f"<w:t>{_xml_escape(toc.title)}</w:t>")
        w.close("w:r")
        w.close("w:p")

        # TOC entries as hyperlinked paragraphs
        for entry in toc.entries:
            indent = (entry.level - 1) * 360  # twips indent per level
            w.open("w:p")
            w.open("w:pPr")
            w.tag("w:pStyle", {"w:val": "Normal"})
            if indent:
                w.tag("w:ind", {"w:left": str(indent)})
            w.tag("w:spacing", {"w:before": "20", "w:after": "20"})
            w.close("w:pPr")
            # Internal hyperlink via bookmark ref
            w.open("w:hyperlink", {"w:anchor": entry.anchor})
            w.open("w:r")
            w.open("w:rPr")
            w.tag("w:rStyle", {"w:val": "Hyperlink"})
            w.close("w:rPr")
            w.raw(f"<w:t>{_xml_escape(entry.text)}</w:t>")
            w.close("w:r")
            w.close("w:hyperlink")
            w.close("w:p")

    # ------------------------------------------------------------------
    # Form fields
    # ------------------------------------------------------------------

    def _write_toggle(self, w: XmlWriter, toggle: ToggleModel) -> None:
        """DOCX fallback: render as a box with a title (no native toggle support)."""
        # Title paragraph with indicator
        indicator = "▼" if toggle.open else "▶"
        w.open("w:p")
        w.open("w:pPr")
        w.tag("w:spacing", {"w:before": "80", "w:after": "40"})
        w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": "F5F5F5"})
        w.close("w:pPr")
        w.open("w:r")
        w.open("w:rPr"); w.tag("w:b", {}); w.close("w:rPr")
        w.w_t(f"{indicator} {toggle.title}")
        w.close("w:r")
        w.close("w:p")
        # Content (always shown in DOCX since there's no interactivity)
        for item in toggle.content:
            self._write_item(w, item)

    def _write_checkbox(self, w: XmlWriter, cb: CheckboxModel) -> None:
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:spacing", {"w:after": "80"}); w.close("w:pPr")
        # Checkbox field
        w.open("w:r")
        w.tag("w:fldChar", {"w:fldCharType": "begin"})
        w.close("w:r")
        w.open("w:r")
        checked_val = "1" if cb.checked else "0"
        w.open("w:instrText", {"xml:space": "preserve"})
        w.text(f" FORMCHECKBOX ")
        w.close("w:instrText")
        w.close("w:r")
        w.open("w:r")
        w.tag("w:fldChar", {"w:fldCharType": "separate"})
        w.close("w:r")
        w.open("w:r")
        # Display as checked/unchecked symbol
        w.w_t("\u2612" if cb.checked else "\u2610")
        w.close("w:r")
        w.open("w:r")
        w.tag("w:fldChar", {"w:fldCharType": "end"})
        w.close("w:r")
        # Label text
        if cb.label:
            w.open("w:r")
            w.w_t(f" {cb.label}")
            w.close("w:r")
        w.close("w:p")

    def _write_text_input(self, w: XmlWriter, inp: TextInputModel) -> None:
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:spacing", {"w:after": "80"}); w.close("w:pPr")
        # Show placeholder or value as underlined text (simulating a form field)
        display = inp.value or inp.placeholder or "________________"
        w.open("w:r")
        w.open("w:rPr")
        w.tag("w:u", {"w:val": "single"})
        if not inp.value and inp.placeholder:
            w.tag("w:color", {"w:val": "808080"})
        w.close("w:rPr")
        w.w_t(f" {display} ")
        w.close("w:r")
        w.close("w:p")

    def _write_dropdown(self, w: XmlWriter, dd: DropdownModel) -> None:
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:spacing", {"w:after": "80"}); w.close("w:pPr")
        # Display as a bracketed selection
        display = dd.value or (dd.options[0] if dd.options else "—")
        w.open("w:r")
        w.open("w:rPr")
        w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": "F0F0F0"})
        w.close("w:rPr")
        w.w_t(f" {display} ▾")
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Frame (positioned text box)
    # ------------------------------------------------------------------

    def _write_frame(self, w: XmlWriter, frame: FrameModel) -> None:
        """Render a positioned frame as a wp:anchor floating text box."""
        # Convert twips → EMU (1 twip = 914400/1440 = 635 EMU)
        TWIP_TO_EMU = 635
        cx = frame.width_twips * TWIP_TO_EMU if frame.width_twips else inch_to_emu(3.0)
        cy = frame.height_twips * TWIP_TO_EMU if frame.height_twips else inch_to_emu(2.0)
        pos_x = frame.x_twips * TWIP_TO_EMU
        pos_y = frame.y_twips * TWIP_TO_EMU

        sid = self._next_shape_id()

        # Anchor type mapping
        rel_h = "page" if frame.anchor == "page" else "column"
        rel_v = "page" if frame.anchor == "page" else "paragraph"

        w.open("w:p")
        w.open("w:r")
        w.open("w:drawing")
        w.open("wp:anchor", {
            "distT": "0", "distB": "0", "distL": "114300", "distR": "114300",
            "simplePos": "0", "relativeHeight": str(251658240 + sid),
            "behindDoc": "0", "locked": "0", "layoutInCell": "1",
            "allowOverlap": "1",
        })
        w.tag("wp:simplePos", {"x": "0", "y": "0"})
        w.open("wp:positionH", {"relativeFrom": rel_h})
        w.open("wp:posOffset"); w.text(str(pos_x)); w.close("wp:posOffset")
        w.close("wp:positionH")
        w.open("wp:positionV", {"relativeFrom": rel_v})
        w.open("wp:posOffset"); w.text(str(pos_y)); w.close("wp:posOffset")
        w.close("wp:positionV")
        w.tag("wp:extent", {"cx": str(cx), "cy": str(cy)})
        w.tag("wp:effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
        w.tag("wp:wrapNone")
        w.tag("wp:docPr", {"id": str(sid), "name": f"Frame{sid}"})

        w.open("a:graphic", {"xmlns:a": _A_NS})
        w.open("a:graphicData", {"uri": _WPS_URI})
        w.open("wps:wsp")

        # Shape properties — rectangle
        w.open("wps:spPr")
        preset = "roundRect" if frame.rounded else "rect"
        w.open("a:prstGeom", {"prst": preset})
        w.tag("a:avLst")
        w.close("a:prstGeom")

        # Fill
        if frame.fill:
            w.open("a:solidFill")
            w.tag("a:srgbClr", {"val": frame.fill.lstrip("#")})
            w.close("a:solidFill")
        else:
            w.tag("a:noFill")

        # Stroke
        if frame.stroke:
            w.open("a:ln", {"w": "12700"})
            w.open("a:solidFill")
            w.tag("a:srgbClr", {"val": frame.stroke.lstrip("#")})
            w.close("a:solidFill")
            w.close("a:ln")
        else:
            w.open("a:ln")
            w.tag("a:noFill")
            w.close("a:ln")

        # Shadow
        if frame.shadow:
            w.open("a:effectLst")
            w.open("a:outerShdw", {"blurRad": "50800", "dist": "38100", "dir": "2700000", "algn": "tl"})
            w.open("a:srgbClr", {"val": "000000"})
            w.tag("a:alpha", {"val": "40000"})
            w.close("a:srgbClr")
            w.close("a:outerShdw")
            w.close("a:effectLst")

        w.close("wps:spPr")

        # Text box content
        w.open("wps:txbx")
        w.open("w:txbxContent")
        if frame.content:
            for child in frame.content:
                self._write_item(w, child)
        else:
            w.raw("<w:p/>")
        w.close("w:txbxContent")
        w.close("wps:txbx")

        # Body properties
        w.open("wps:bodyPr", {
            "wrap": "square",
            "lIns": "91440", "tIns": "45720",
            "rIns": "91440", "bIns": "45720",
            "anchor": "t",
        })
        w.close("wps:bodyPr")

        w.close("wps:wsp")
        w.close("a:graphicData")
        w.close("a:graphic")
        w.close("wp:anchor")
        w.close("w:drawing")
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Page break
    # ------------------------------------------------------------------

    def _write_page_break(self, w: XmlWriter, _item=None) -> None:
        w.open("w:p"); w.open("w:r")
        w.tag("w:br", {"w:type": "page"})
        w.close("w:r"); w.close("w:p")

    # ------------------------------------------------------------------
    # Section properties
    # ------------------------------------------------------------------

    def _write_sect_pr(self, w: XmlWriter,
                       header_rel_id: str | None = None,
                       footer_rel_id: str | None = None) -> None:
        section = self._model.sections[-1] if self._model.sections else SectionModel()
        pw, ph  = PAPER_SIZES.get(section.paper, PAPER_SIZES["a4"])
        margins = section.resolved_margins(MARGIN_PRESETS)

        w.open("w:sectPr")
        if header_rel_id:
            w.tag("w:headerReference", {"w:type": "default", "r:id": header_rel_id})
        if footer_rel_id:
            w.tag("w:footerReference", {"w:type": "default", "r:id": footer_rel_id})
        w.tag("w:pgSz", {"w:w": str(pw), "w:h": str(ph)})
        w.tag("w:pgMar", {
            "w:top": str(margins["top"]), "w:right": str(margins["right"]),
            "w:bottom": str(margins["bottom"]), "w:left": str(margins["left"]),
            "w:header": "709", "w:footer": "709", "w:gutter": "0",
        })
        if section.cols > 1:
            w.tag("w:cols", {"w:num": str(section.cols), "w:space": "720"})
        w.close("w:sectPr")

    # ------------------------------------------------------------------
    # Header / Footer XML
    # ------------------------------------------------------------------

    def _build_header_xml(self) -> str:
        return self._build_hf_xml("w:hdr", self._model.header.paragraphs)

    def _build_footer_xml(self) -> str:
        return self._build_hf_xml("w:ftr", self._model.footer.paragraphs)

    def _build_hf_xml(self, tag: str, paragraphs: list[ParagraphModel]) -> str:
        w = XmlWriter()
        w.declaration()
        w.open(tag, {"xmlns:w": _W_NS, "xmlns:r": _R_NS})
        for para in paragraphs:
            self._write_paragraph(w, para)
        if not paragraphs:
            w.raw("<w:p/>")
        w.close(tag)
        return w.getvalue()

    # ------------------------------------------------------------------
    # Numbering XML
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # Styles XML
    # ------------------------------------------------------------------

    def _build_styles_xml(self) -> str:
        return build_styles_xml(
            font=self._model.default_font,
            size_pt=self._model.default_size_pt,
            spacing=self._model.spacing,
        )
