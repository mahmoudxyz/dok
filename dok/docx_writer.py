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

from .converter import (
    DocxModel, ParagraphModel, RunModel, ShapeModel,
    RowModel, TableModel, TableCellModel,
    PageBreakModel, SectionModel,
    LineModel, BoxModel, BannerModel, BadgeModel,
    DataTableModel, DataTableRowModel, DataTableCellModel,
    ImageModel, SpacerModel, HeaderModel, FooterModel,
    _SPACING_PRESETS,
)
from .xml_writer import XmlWriter


# ---------------------------------------------------------------------------
# Units
# ---------------------------------------------------------------------------

def inch_to_emu(i: float) -> int:  return int(i * 914400)
def inch_to_twip(i: float)-> int:  return int(i * 1440)

MARGIN_PRESETS = {
    "normal": {"top": 1440, "right": 1440, "bottom": 1440, "left": 1440},
    "narrow": {"top":  720, "right":  720, "bottom":  720, "left":  720},
    "wide":   {"top": 1440, "right": 2880, "bottom": 1440, "left": 2880},
    "none":   {"top":    0, "right":    0, "bottom":    0, "left":    0},
}

PAPER_SIZES = {
    "a4":     (11906, 16838),
    "letter": (12240, 15840),
    "a3":     (16838, 23811),
}

def content_width_twip(paper: str = "a4", margin: str = "normal") -> int:
    pw, _ = PAPER_SIZES.get(paper, PAPER_SIZES["a4"])
    m = MARGIN_PRESETS.get(margin, MARGIN_PRESETS["normal"])
    return pw - m["left"] - m["right"]


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
        numbering_xml = self._build_numbering_xml() if self._model.has_lists else None

        # Assign rel IDs for header/footer/numbering
        header_rel_id = footer_rel_id = numbering_rel_id = None
        if numbering_xml:
            numbering_rel_id = self._next_rel_id()
        if header_xml:
            header_rel_id = self._next_rel_id()
        if footer_xml:
            footer_rel_id = self._next_rel_id()

        # Build dynamic rels and content types
        doc_rels_xml    = self._build_doc_rels(header_rel_id, footer_rel_id, numbering_rel_id)
        content_types   = self._build_content_types(header_xml, footer_xml, numbering_xml)

        # If we have header/footer, rebuild doc xml with correct rel IDs in sectPr
        if header_rel_id or footer_rel_id:
            doc_xml = self._build_document_xml(header_rel_id, footer_rel_id)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",          content_types)
            zf.writestr("_rels/.rels",                  _RELS)
            zf.writestr("word/document.xml",            doc_xml)
            zf.writestr("word/styles.xml",              self._build_styles_xml())
            zf.writestr("word/settings.xml",            _SETTINGS)
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

    def _write_item(self, w: XmlWriter, item) -> None:
        if isinstance(item, ParagraphModel):   self._write_paragraph(w, item)
        elif isinstance(item, LineModel):      self._write_line(w, item)
        elif isinstance(item, BoxModel):       self._write_box(w, item)
        elif isinstance(item, BannerModel):    self._write_banner(w, item)
        elif isinstance(item, BadgeModel):     self._write_badge(w, item)
        elif isinstance(item, DataTableModel): self._write_data_table(w, item)
        elif isinstance(item, ImageModel):     self._write_image(w, item)
        elif isinstance(item, SpacerModel):    self._write_spacer(w, item)
        elif isinstance(item, ShapeModel):     self._write_shape(w, item)
        elif isinstance(item, RowModel):       self._write_row(w, item)
        elif isinstance(item, TableModel):     self._write_table(w, item)
        elif isinstance(item, PageBreakModel): self._write_page_break(w)

    # ------------------------------------------------------------------
    # Paragraph
    # ------------------------------------------------------------------

    def _write_paragraph(self, w: XmlWriter, para: ParagraphModel) -> None:
        w.open("w:p")
        w.open("w:pPr")

        w.tag("w:pStyle", {"w:val": para.style})

        if para.align != "left":
            jc = {"center": "center", "right": "right",
                  "justify": "both"}.get(para.align, "left")
            w.tag("w:jc", {"w:val": jc})

        if para.direction == "rtl":
            w.tag("w:bidi", {})

        if para.indent_twips > 0:
            w.tag("w:ind", {"w:left": str(para.indent_twips)})

        spacing: dict = {}
        if para.space_before: spacing["w:before"] = str(para.space_before)
        if para.space_after:  spacing["w:after"]  = str(para.space_after)
        if spacing:
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

        # Group runs by hyperlink URL for proper wrapping
        self._write_runs_with_hyperlinks(w, para.runs)

        w.close("w:p")

    def _write_runs_with_hyperlinks(self, w: XmlWriter, runs: list[RunModel]) -> None:
        """Write runs, grouping consecutive hyperlinked runs into <w:hyperlink>."""
        i = 0
        while i < len(runs):
            run = runs[i]
            if run.hyperlink_url:
                url = run.hyperlink_url
                rel_id = self._get_hyperlink_rel(url)
                w.open("w:hyperlink", {"r:id": rel_id})
                while i < len(runs) and runs[i].hyperlink_url == url:
                    self._write_run(w, runs[i])
                    i += 1
                w.close("w:hyperlink")
            else:
                self._write_run(w, run)
                i += 1

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
        section = self._model.sections[-1] if self._model.sections else SectionModel()
        cw = content_width_twip(section.paper, section.margin)

        w.open("w:tbl")
        w.open("w:tblPr")
        w.tag("w:tblW", {"w:w": str(cw), "w:type": "dxa"})
        w.open("w:tblCellMar")
        w.tag("w:top",    {"w:w": "100", "w:type": "dxa"})
        w.tag("w:left",   {"w:w": "160", "w:type": "dxa"})
        w.tag("w:bottom", {"w:w": "100", "w:type": "dxa"})
        w.tag("w:right",  {"w:w": "160", "w:type": "dxa"})
        w.close("w:tblCellMar")
        w.open("w:tblBorders")
        stroke_color = box.stroke or "BFBFBF"
        if box.stroke:
            for side in ("top", "left", "bottom", "right"):
                w.tag(f"w:{side}", {"w:val": "single", "w:sz": "4",
                                    "w:space": "0", "w:color": stroke_color})
        else:
            for side in ("top", "left", "bottom", "right"):
                w.tag(f"w:{side}", {"w:val": "none", "w:sz": "0",
                                    "w:space": "0", "w:color": "auto"})
        for side in ("insideH", "insideV"):
            w.tag(f"w:{side}", {"w:val": "none", "w:sz": "0",
                                "w:space": "0", "w:color": "auto"})
        w.close("w:tblBorders")
        w.close("w:tblPr")

        w.open("w:tr")
        w.open("w:tc")
        w.open("w:tcPr")
        w.tag("w:tcW", {"w:w": str(cw), "w:type": "dxa"})
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

        w.open("w:p")
        w.open("w:pPr"); w.tag("w:spacing", {"w:before": "0", "w:after": "120"}); w.close("w:pPr")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Banner
    # ------------------------------------------------------------------

    def _write_banner(self, w: XmlWriter, banner: BannerModel) -> None:
        for para in banner.paragraphs:
            self._write_paragraph(w, para)

    # ------------------------------------------------------------------
    # Badge
    # ------------------------------------------------------------------

    def _write_badge(self, w: XmlWriter, badge: BadgeModel) -> None:
        w.open("w:p")
        w.open("w:pPr")
        if badge.align != "left":
            jc = {"center": "center", "right": "right"}.get(badge.align, "left")
            w.tag("w:jc", {"w:val": jc})
        w.tag("w:spacing", {"w:after": "160"})
        w.close("w:pPr")

        padded = f"  {badge.text}  "
        w.open("w:r")
        w.open("w:rPr")
        if badge.color: w.tag("w:color", {"w:val": badge.color})
        if badge.fill:  w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": badge.fill})
        w.close("w:rPr")
        w.w_t(padded)
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Data Table
    # ------------------------------------------------------------------

    def _write_data_table(self, w: XmlWriter, table: DataTableModel) -> None:
        section = self._model.sections[-1] if self._model.sections else SectionModel()
        cw = content_width_twip(section.paper, section.margin)

        # Determine number of columns from widest row
        max_cols = max((sum(c.colspan for c in r.cells) for r in table.rows), default=1)

        w.open("w:tbl")
        w.open("w:tblPr")
        w.tag("w:tblW", {"w:w": str(cw), "w:type": "dxa"})
        w.tag("w:tblLayout", {"w:type": "fixed"})

        if table.border:
            w.open("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                w.tag(f"w:{side}", {"w:val": "single", "w:sz": "4",
                                    "w:space": "0", "w:color": "BFBFBF"})
            w.close("w:tblBorders")

        w.open("w:tblCellMar")
        w.tag("w:top",    {"w:w": "40",  "w:type": "dxa"})
        w.tag("w:left",   {"w:w": "100", "w:type": "dxa"})
        w.tag("w:bottom", {"w:w": "40",  "w:type": "dxa"})
        w.tag("w:right",  {"w:w": "100", "w:type": "dxa"})
        w.close("w:tblCellMar")

        w.close("w:tblPr")

        # Grid columns
        col_w = cw // max_cols
        w.open("w:tblGrid")
        for _ in range(max_cols):
            w.tag("w:gridCol", {"w:w": str(col_w)})
        w.close("w:tblGrid")

        for row_idx, row in enumerate(table.rows):
            w.open("w:tr")
            for cell in row.cells:
                cell_w = col_w * cell.colspan
                w.open("w:tc")
                w.open("w:tcPr")
                w.tag("w:tcW", {"w:w": str(cell_w), "w:type": "dxa"})
                if cell.colspan > 1:
                    w.tag("w:gridSpan", {"w:val": str(cell.colspan)})

                # Header row: shaded background
                if row.is_header:
                    w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": "E8E8E8"})

                # Striped rows
                if table.striped and not row.is_header and row_idx % 2 == 0:
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

        # Spacing after table
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:spacing", {"w:before": "0", "w:after": "120"}); w.close("w:pPr")
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
        w.open("w:p")
        w.open("w:pPr"); w.tag("w:jc", {"w:val": "center"}); w.close("w:pPr")
        shape_w = inch_to_emu(1.4); shape_h = inch_to_emu(0.6)

        for i, shape in enumerate(row.shapes):
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

            if i < len(row.arrows):
                label = row.arrows[i]
                arrow_text = f"  {label}  " if label else "  \u2192  "
                w.open("w:r"); w.open("w:rPr"); w.close("w:rPr")
                w.w_t(arrow_text); w.close("w:r")
        w.close("w:p")

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
        section = self._model.sections[-1] if self._model.sections else SectionModel()
        cw = content_width_twip(section.paper, section.margin)
        w.open("w:tbl")
        w.open("w:tblPr")
        w.tag("w:tblW", {"w:w": str(cw), "w:type": "dxa"})
        if not table.border:
            w.open("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                w.tag(f"w:{side}", {"w:val": "none", "w:sz": "0",
                                    "w:space": "0", "w:color": "auto"})
            w.close("w:tblBorders")
        w.close("w:tblPr")

        for tbl_row in table.rows:
            w.open("w:tr")
            total_pct = sum(c.width_pct for c in tbl_row.cells) or 100
            for cell in tbl_row.cells:
                cell_w = int(cw * cell.width_pct / total_pct)
                w.open("w:tc")
                w.open("w:tcPr"); w.tag("w:tcW", {"w:w": str(cell_w), "w:type": "dxa"}); w.close("w:tcPr")
                if not cell.content:
                    w.raw("<w:p/>")
                else:
                    for item in cell.content:
                        self._write_item(w, item)
                w.close("w:tc")
            w.close("w:tr")

        w.close("w:tbl")
        w.open("w:p"); w.open("w:pPr")
        w.tag("w:spacing", {"w:before": "0", "w:after": "80"})
        w.close("w:pPr"); w.close("w:p")

    # ------------------------------------------------------------------
    # Page break
    # ------------------------------------------------------------------

    def _write_page_break(self, w: XmlWriter) -> None:
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
        margins = MARGIN_PRESETS.get(section.margin, MARGIN_PRESETS["normal"])

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

    def _build_numbering_xml(self) -> str:
        w = XmlWriter()
        w.declaration()
        w.open("w:numbering", {"xmlns:w": _W_NS})

        # Abstract 0: bullet list
        w.open("w:abstractNum", {"w:abstractNumId": "0"})
        w.tag("w:multiLevelType", {"w:val": "hybridMultilevel"})
        bullets = ["\u2022", "\u25E6", "\u2013"]  # •, ◦, –
        for lvl in range(3):
            w.open("w:lvl", {"w:ilvl": str(lvl)})
            w.tag("w:start", {"w:val": "1"})
            w.tag("w:numFmt", {"w:val": "bullet"})
            w.tag("w:lvlText", {"w:val": bullets[lvl % len(bullets)]})
            w.tag("w:lvlJc", {"w:val": "left"})
            w.open("w:pPr")
            indent = 720 * (lvl + 1)
            w.tag("w:ind", {"w:left": str(indent), "w:hanging": "360"})
            w.close("w:pPr")
            w.open("w:rPr")
            w.tag("w:rFonts", {"w:ascii": "Calibri", "w:hAnsi": "Calibri", "w:hint": "default"})
            w.close("w:rPr")
            w.close("w:lvl")
        w.close("w:abstractNum")

        # Abstract 1: ordered list
        w.open("w:abstractNum", {"w:abstractNumId": "1"})
        w.tag("w:multiLevelType", {"w:val": "hybridMultilevel"})
        formats = [("decimal", "%1."), ("lowerLetter", "%2."), ("lowerRoman", "%3.")]
        for lvl in range(3):
            fmt, text = formats[lvl % len(formats)]
            w.open("w:lvl", {"w:ilvl": str(lvl)})
            w.tag("w:start", {"w:val": "1"})
            w.tag("w:numFmt", {"w:val": fmt})
            w.tag("w:lvlText", {"w:val": text})
            w.tag("w:lvlJc", {"w:val": "left"})
            w.open("w:pPr")
            indent = 720 * (lvl + 1)
            w.tag("w:ind", {"w:left": str(indent), "w:hanging": "360"})
            w.close("w:pPr")
            w.close("w:lvl")
        w.close("w:abstractNum")

        # Concrete instances
        w.open("w:num", {"w:numId": "1"})
        w.tag("w:abstractNumId", {"w:val": "0"})
        w.close("w:num")
        w.open("w:num", {"w:numId": "2"})
        w.tag("w:abstractNumId", {"w:val": "1"})
        w.close("w:num")

        w.close("w:numbering")
        return w.getvalue()

    # ------------------------------------------------------------------
    # Dynamic relationships
    # ------------------------------------------------------------------

    def _build_doc_rels(self, header_rel_id: str | None = None,
                        footer_rel_id: str | None = None,
                        numbering_rel_id: str | None = None) -> str:
        w = XmlWriter()
        w.declaration()
        w.open("Relationships", {"xmlns": _RELS_NS})

        w.tag("Relationship", {"Id": "rId1",
              "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
              "Target": "styles.xml"})
        w.tag("Relationship", {"Id": "rId2",
              "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
              "Target": "settings.xml"})

        if numbering_rel_id:
            w.tag("Relationship", {"Id": numbering_rel_id,
                  "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                  "Target": "numbering.xml"})

        if header_rel_id:
            w.tag("Relationship", {"Id": header_rel_id,
                  "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                  "Target": "header1.xml"})

        if footer_rel_id:
            w.tag("Relationship", {"Id": footer_rel_id,
                  "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
                  "Target": "footer1.xml"})

        for rel_id, filename, _ in self._image_entries:
            w.tag("Relationship", {"Id": rel_id,
                  "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                  "Target": f"media/{filename}"})

        for rel_id, url in self._hyperlink_rels:
            w.tag("Relationship", {"Id": rel_id,
                  "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                  "Target": url, "TargetMode": "External"})

        w.close("Relationships")
        return w.getvalue()

    def _build_content_types(self, header_xml, footer_xml, numbering_xml) -> str:
        w = XmlWriter()
        w.declaration()
        w.open("Types", {"xmlns": "http://schemas.openxmlformats.org/package/2006/content-types"})

        w.tag("Default", {"Extension": "rels",
              "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
        w.tag("Default", {"Extension": "xml", "ContentType": "application/xml"})

        # Image types
        exts_seen: set[str] = set()
        for _, filename, _ in self._image_entries:
            ext = Path(filename).suffix.lstrip(".")
            if ext not in exts_seen:
                exts_seen.add(ext)
                from .image import image_content_type
                ct = image_content_type(filename)
                w.tag("Default", {"Extension": ext, "ContentType": ct})

        w.tag("Override", {"PartName": "/word/document.xml",
              "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"})
        w.tag("Override", {"PartName": "/word/styles.xml",
              "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"})
        w.tag("Override", {"PartName": "/word/settings.xml",
              "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"})

        if numbering_xml:
            w.tag("Override", {"PartName": "/word/numbering.xml",
                  "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"})
        if header_xml:
            w.tag("Override", {"PartName": "/word/header1.xml",
                  "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"})
        if footer_xml:
            w.tag("Override", {"PartName": "/word/footer1.xml",
                  "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"})

        w.close("Types")
        return w.getvalue()

    # ------------------------------------------------------------------
    # Styles XML
    # ------------------------------------------------------------------

    def _build_styles_xml(self) -> str:
        font = self._model.default_font
        size = self._model.default_size_pt * 2
        para_after, h_scale, line_sp = _SPACING_PRESETS.get(
            self._model.spacing, _SPACING_PRESETS["normal"])

        w = XmlWriter()
        w.declaration()
        w.open("w:styles", {"xmlns:w": _W_NS})

        w.open("w:docDefaults")
        w.open("w:rPrDefault"); w.open("w:rPr")
        w.tag("w:rFonts", {"w:ascii": font, "w:hAnsi": font, "w:cs": font})
        w.tag("w:sz", {"w:val": str(size)}); w.tag("w:szCs", {"w:val": str(size)})
        w.tag("w:lang", {"w:val": "en-US", "w:eastAsia": "en-US", "w:bidi": "ar-SA"})
        w.close("w:rPr"); w.close("w:rPrDefault")
        w.open("w:pPrDefault"); w.open("w:pPr")
        w.tag("w:spacing", {"w:after": str(para_after), "w:line": str(line_sp), "w:lineRule": "auto"})
        w.close("w:pPr"); w.close("w:pPrDefault")
        w.close("w:docDefaults")

        self._write_style(w, "Normal", "Normal", is_default=True, font=font, size=size,
                          spacing_after=para_after, line_spacing=line_sp)

        for lvl, sz in {1: 36, 2: 32, 3: 28, 4: 26}.items():
            clr = "1F3864" if lvl <= 2 else "404040"
            before = int({1: 480, 2: 360, 3: 280, 4: 240}[lvl] * h_scale)
            after  = int(120 * h_scale)
            self._write_style(w, f"Heading{lvl}", f"Heading {lvl}",
                              font=font, size=sz, bold=True, color=clr,
                              outline_lvl=lvl - 1, spacing_before=before,
                              spacing_after=after, line_spacing=min(line_sp, 240))

        self._write_style(w, "BlockText", "Block Text", font=font, size=size,
                          italic=True, color="404040", indent_left=720, indent_right=720,
                          spacing_before=int(120 * h_scale),
                          spacing_after=int(120 * h_scale), line_spacing=line_sp)

        self._write_style(w, "SourceCode", "Source Code", font="Courier New", size=20,
                          spacing_before=0, spacing_after=0, line_spacing=240)

        # List Paragraph style (for bullet/numbered lists)
        list_after = max(0, min(para_after, 40))
        self._write_style(w, "ListParagraph", "List Paragraph", font=font, size=size,
                          indent_left=720, spacing_after=list_after, line_spacing=line_sp)

        # Hyperlink character style
        w.open("w:style", {"w:type": "character", "w:styleId": "Hyperlink"})
        w.tag("w:name", {"w:val": "Hyperlink"})
        w.open("w:rPr")
        w.tag("w:color", {"w:val": "0563C1"})
        w.tag("w:u", {"w:val": "single"})
        w.close("w:rPr")
        w.close("w:style")

        w.close("w:styles")
        return w.getvalue()

    def _write_style(self, w: XmlWriter, style_id: str, name: str,
                     is_default: bool = False,
                     font: str = "Calibri", size: int = 22,
                     bold: bool = False, italic: bool = False,
                     color: str | None = None,
                     outline_lvl: int | None = None,
                     spacing_before: int = 0, spacing_after: int = 160,
                     indent_left: int = 0, indent_right: int = 0,
                     line_spacing: int = 276) -> None:
        attrs: dict = {"w:type": "paragraph", "w:styleId": style_id}
        if is_default: attrs["w:default"] = "1"
        w.open("w:style", attrs)
        w.tag("w:name", {"w:val": name})
        if not is_default: w.tag("w:basedOn", {"w:val": "Normal"})

        w.open("w:pPr")
        if outline_lvl is not None:
            w.tag("w:outlineLvl", {"w:val": str(outline_lvl)})
        sp: dict = {"w:line": str(line_spacing), "w:lineRule": "auto"}
        if spacing_before: sp["w:before"] = str(spacing_before)
        if spacing_after:  sp["w:after"]  = str(spacing_after)
        w.tag("w:spacing", sp)
        if indent_left or indent_right:
            ind: dict = {}
            if indent_left:  ind["w:left"]  = str(indent_left)
            if indent_right: ind["w:right"] = str(indent_right)
            w.tag("w:ind", ind)
        w.close("w:pPr")

        w.open("w:rPr")
        w.tag("w:rFonts", {"w:ascii": font, "w:hAnsi": font, "w:cs": font})
        w.tag("w:sz", {"w:val": str(size)}); w.tag("w:szCs", {"w:val": str(size)})
        if bold:  w.tag("w:b", {})
        if italic: w.tag("w:i", {})
        if color: w.tag("w:color", {"w:val": color})
        w.close("w:rPr")
        w.close("w:style")


# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------

_W_NS    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_A_NS    = "http://schemas.openxmlformats.org/drawingml/2006/main"
_WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_NS = {
    "xmlns:wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "xmlns:mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:r":   _R_NS,
    "xmlns:wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xmlns:w":   _W_NS,
    "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "xmlns:a":   _A_NS,
    "xmlns:pic": _PIC_URI,
}

_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

_SETTINGS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word"
      w:val="15"/>
  </w:compat>
</w:settings>"""
