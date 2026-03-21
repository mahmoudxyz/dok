"""
dok.docx_writer
~~~~~~~~~~~~~~~
Converts a DocxModel to a .docx file (ZIP + XML).
"""

from __future__ import annotations
import io
import zipfile
from pathlib import Path

from .converter import (
    DocxModel, ParagraphModel, RunModel, ShapeModel,
    RowModel, TableModel, TableCellModel,
    PageBreakModel, SectionModel
)
from .xml_writer import XmlWriter


# ---------------------------------------------------------------------------
# Unit conversions
# ---------------------------------------------------------------------------

def pt_to_twip(pt: float) -> int:   return int(pt * 20)
def pt_to_emu(pt: float)  -> int:   return int(pt * 12700)
def inch_to_emu(i: float) -> int:   return int(i * 914400)
def inch_to_twip(i: float)-> int:   return int(i * 1440)


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

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def write(self, dest: str | Path | io.BytesIO) -> None:
        doc_xml = self._build_document_xml()

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",          _CONTENT_TYPES)
            zf.writestr("_rels/.rels",                  _RELS)
            zf.writestr("word/document.xml",            doc_xml)
            zf.writestr("word/styles.xml",              self._build_styles_xml())
            zf.writestr("word/settings.xml",            _SETTINGS)
            zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)

        data = buf.getvalue()
        if isinstance(dest, (str, Path)):
            Path(dest).write_bytes(data)
        else:
            dest.write(data)

    # ------------------------------------------------------------------
    # document.xml
    # ------------------------------------------------------------------

    def _build_document_xml(self) -> str:
        w = XmlWriter()
        w.declaration()
        w.open("w:document", _NS)
        w.open("w:body")

        for item in self._model.content:
            if isinstance(item, ParagraphModel):
                self._write_paragraph(w, item)
            elif isinstance(item, ShapeModel):
                self._write_shape(w, item)
            elif isinstance(item, RowModel):
                self._write_row(w, item)
            elif isinstance(item, TableModel):
                self._write_table(w, item)
            elif isinstance(item, PageBreakModel):
                self._write_page_break(w)

        self._write_sect_pr(w)
        w.close("w:body")
        w.close("w:document")
        return w.getvalue()

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

        if para.direction == "rtl":
            w.open("w:rPr")
            w.tag("w:rtl", {})
            w.close("w:rPr")

        w.close("w:pPr")

        for run in para.runs:
            self._write_run(w, run)

        w.close("w:p")

    def _write_run(self, w: XmlWriter, run: RunModel) -> None:
        if not run.text:
            return

        w.open("w:r")
        w.open("w:rPr")

        # Only emit font/size when EXPLICITLY set on this run.
        # If we always emit sz=default, it overrides the paragraph style —
        # e.g. Heading1 defines sz=36, but run sz=22 would win (wrong).
        if run.font:
            w.tag("w:rFonts", {
                "w:ascii": run.font, "w:hAnsi": run.font, "w:cs": run.font,
            })

        if run.size_pt is not None:
            size = run.size_pt * 2
            w.tag("w:sz",   {"w:val": str(size)})
            w.tag("w:szCs", {"w:val": str(size)})

        if run.bold:      w.tag("w:b",      {})
        if run.italic:    w.tag("w:i",      {})
        if run.underline: w.tag("w:u",      {"w:val": "single"})
        if run.strike:    w.tag("w:strike", {})

        if run.color:
            w.tag("w:color", {"w:val": run.color})
        if run.highlight:
            w.tag("w:highlight", {"w:val": run.highlight})

        va = "superscript" if run.sup else "subscript" if run.sub else None
        if va:
            w.tag("w:vertAlign", {"w:val": va})
        if run.rtl:
            w.tag("w:rtl", {})

        w.close("w:rPr")
        w.w_t(run.text)
        w.close("w:r")

    # ------------------------------------------------------------------
    # Shape
    # ------------------------------------------------------------------

    def _write_shape(
        self,
        w: XmlWriter,
        shape: ShapeModel,
        width_emu:  int | None = None,
        height_emu: int | None = None,
    ) -> None:

        if width_emu is None:
            width_emu = inch_to_emu(2)
        if height_emu is None:
            if shape.paragraphs:
                estimated_in = 0.5 + 0.35 * len(shape.paragraphs)
                height_emu = inch_to_emu(min(estimated_in, 4.0))
            else:
                height_emu = inch_to_emu(0.8)

        if shape.full_width:
            section   = self._model.sections[-1] if self._model.sections else SectionModel()
            cw        = content_width_twip(section.paper, section.margin)
            width_emu = int(cw / 1440 * 914400)
            if not shape.paragraphs:
                height_emu = inch_to_emu(0.8)

        sid = self._next_shape_id()

        w.open("w:p")
        w.open("w:pPr")

        # Banner accent: thick left paragraph border — no separate shape needed
        if shape.full_width and shape.accent:
            w.open("w:pBdr")
            w.tag("w:left", {
                "w:val":   "single",
                "w:sz":    "48",
                "w:space": "4",
                "w:color": shape.accent,
            })
            w.close("w:pBdr")

        w.close("w:pPr")

        w.open("w:r")
        w.open("w:rPr")
        w.close("w:rPr")
        w.open("w:drawing")

        if shape.inline:
            w.open("wp:inline", {
                "distT": "0", "distB": "0",
                "distL": "0", "distR": "0",
            })
        else:
            side = shape.float_side or "right"
            w.open("wp:anchor", {
                "distT": "114300", "distB": "114300",
                "distL": "114300", "distR": "114300",
                "simplePos": "0", "relativeHeight": "251658240",
                "behindDoc": "0", "locked": "0",
                "layoutInCell": "1", "allowOverlap": "1",
            })
            w.tag("wp:simplePos", {"x": "0", "y": "0"})
            pos_h_align = "right" if side == "right" else "left"
            w.open("wp:positionH", {"relativeFrom": "margin"})
            w.raw(f"<wp:alignment>{pos_h_align}</wp:alignment>")
            w.close("wp:positionH")
            w.open("wp:positionV", {"relativeFrom": "paragraph"})
            w.raw("<wp:alignment>top</wp:alignment>")
            w.close("wp:positionV")

        w.tag("wp:extent", {"cx": str(width_emu), "cy": str(height_emu)})
        w.tag("wp:effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})

        if not shape.inline:
            w.tag("wp:wrapSquare", {"wrapText": "bothSides"})

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
        if shape.shadow:
            w.open("a:effectLst")
            w.open("a:outerShdw", {
                "blurRad": "40000", "dist": "23000",
                "dir": "5400000", "rotWithShape": "0",
            })
            w.open("a:srgbClr", {"val": "000000"})
            w.tag("a:alpha", {"val": "35000"})
            w.close("a:srgbClr")
            w.close("a:outerShdw")
            w.close("a:effectLst")
        w.close("wps:spPr")

        if shape.paragraphs:
            w.open("wps:txbx")
            w.open("w:txbxContent")
            for para in shape.paragraphs:
                self._write_paragraph(w, para)
            w.close("w:txbxContent")
            w.close("wps:txbx")

        w.open("wps:bodyPr", {
            "wrap":  "square",
            "lIns": "91440", "tIns": "45720",
            "rIns": "91440", "bIns": "45720",
        })
        w.close("wps:bodyPr")

        w.close("wps:wsp")
        w.close("a:graphicData")
        w.close("a:graphic")

        if shape.inline:
            w.close("wp:inline")
        else:
            w.close("wp:anchor")

        w.close("w:drawing")
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Row
    # ------------------------------------------------------------------

    def _write_row(self, w: XmlWriter, row: RowModel) -> None:
        w.open("w:p")
        w.open("w:pPr")
        w.tag("w:jc", {"w:val": "center"})
        w.close("w:pPr")

        shape_width  = inch_to_emu(1.4)
        shape_height = inch_to_emu(0.6)

        for i, shape in enumerate(row.shapes):
            sid = self._next_shape_id()

            w.open("w:r")
            w.open("w:rPr"); w.close("w:rPr")
            w.open("w:drawing")

            w.open("wp:inline", {
                "distT": "0", "distB": "0",
                "distL": "57150", "distR": "57150",
            })
            w.tag("wp:extent", {"cx": str(shape_width), "cy": str(shape_height)})
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

            w.open("wps:bodyPr", {
                "wrap":  "square",
                "lIns": "91440", "tIns": "45720",
                "rIns": "91440", "bIns": "45720",
            })
            w.close("wps:bodyPr")
            w.close("wps:wsp")
            w.close("a:graphicData")
            w.close("a:graphic")
            w.close("wp:inline")
            w.close("w:drawing")
            w.close("w:r")

            # Arrow between shapes
            if i < len(row.arrows):
                label      = row.arrows[i]
                arrow_text = f"  {label}  " if label else "  →  "
                w.open("w:r")
                w.open("w:rPr"); w.close("w:rPr")
                w.w_t(arrow_text)
                w.close("w:r")

        w.close("w:p")

    # ------------------------------------------------------------------
    # Fill and stroke
    # ------------------------------------------------------------------

    def _write_fill(self, w: XmlWriter, hex_color: str | None) -> None:
        if hex_color is None:
            w.tag("a:noFill")
        else:
            w.open("a:solidFill")
            w.tag("a:srgbClr", {"val": hex_color})
            w.close("a:solidFill")

    def _write_stroke(self, w: XmlWriter, shape: ShapeModel) -> None:
        thickness = "25400" if shape.stroke_thick else "9525"

        if shape.stroke is None:
            w.open("a:ln", {"w": thickness})
            w.tag("a:noFill")
            w.close("a:ln")
            return

        w.open("a:ln", {"w": thickness})
        w.open("a:solidFill")
        w.tag("a:srgbClr", {"val": shape.stroke})
        w.close("a:solidFill")

        if shape.stroke_style == "dashed":
            w.tag("a:prstDash", {"val": "dash"})
        elif shape.stroke_style == "dotted":
            w.tag("a:prstDash", {"val": "dot"})
        else:
            w.tag("a:prstDash", {"val": "solid"})

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
                w.tag(f"w:{side}", {
                    "w:val": "none", "w:sz": "0",
                    "w:space": "0", "w:color": "auto",
                })
            w.close("w:tblBorders")

        w.close("w:tblPr")

        for tbl_row in table.rows:
            w.open("w:tr")
            total_pct = sum(c.width_pct for c in tbl_row.cells) or 100

            for cell in tbl_row.cells:
                cell_w = int(cw * cell.width_pct / total_pct)
                w.open("w:tc")
                w.open("w:tcPr")
                w.tag("w:tcW", {"w:w": str(cell_w), "w:type": "dxa"})
                w.close("w:tcPr")

                if not cell.content:
                    w.raw("<w:p/>")
                else:
                    for item in cell.content:
                        if isinstance(item, ParagraphModel):
                            self._write_paragraph(w, item)
                        elif isinstance(item, ShapeModel):
                            self._write_shape(w, item)
                        elif isinstance(item, RowModel):
                            self._write_row(w, item)

                w.close("w:tc")
            w.close("w:tr")

        w.close("w:tbl")
        w.raw("<w:p/>")  # table must be followed by an empty paragraph

    # ------------------------------------------------------------------
    # Page break
    # ------------------------------------------------------------------

    def _write_page_break(self, w: XmlWriter) -> None:
        w.open("w:p")
        w.open("w:r")
        w.tag("w:br", {"w:type": "page"})
        w.close("w:r")
        w.close("w:p")

    # ------------------------------------------------------------------
    # Section properties
    # ------------------------------------------------------------------

    def _write_sect_pr(self, w: XmlWriter) -> None:
        section = self._model.sections[-1] if self._model.sections else SectionModel()
        pw, ph  = PAPER_SIZES.get(section.paper, PAPER_SIZES["a4"])
        margins = MARGIN_PRESETS.get(section.margin, MARGIN_PRESETS["normal"])

        w.open("w:sectPr")
        w.tag("w:pgSz", {"w:w": str(pw), "w:h": str(ph)})
        w.tag("w:pgMar", {
            "w:top":    str(margins["top"]),
            "w:right":  str(margins["right"]),
            "w:bottom": str(margins["bottom"]),
            "w:left":   str(margins["left"]),
            "w:header": "709",
            "w:footer": "709",
            "w:gutter": "0",
        })
        if section.cols > 1:
            w.tag("w:cols", {"w:num": str(section.cols), "w:space": "720"})
        w.close("w:sectPr")

    # ------------------------------------------------------------------
    # Styles XML
    # ------------------------------------------------------------------

    def _build_styles_xml(self) -> str:
        font = self._model.default_font
        size = self._model.default_size_pt * 2

        w = XmlWriter()
        w.declaration()
        w.open("w:styles", {
            "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        })

        w.open("w:docDefaults")
        w.open("w:rPrDefault")
        w.open("w:rPr")
        w.tag("w:rFonts", {"w:ascii": font, "w:hAnsi": font, "w:cs": font})
        w.tag("w:sz",   {"w:val": str(size)})
        w.tag("w:szCs", {"w:val": str(size)})
        w.tag("w:lang", {"w:val": "en-US", "w:eastAsia": "en-US", "w:bidi": "ar-SA"})
        w.close("w:rPr")
        w.close("w:rPrDefault")
        w.close("w:docDefaults")

        self._write_style(w, "Normal", "Normal", is_default=True,
                          font=font, size=size, spacing_after=160)

        heading_sizes  = {1: 36, 2: 32, 3: 28, 4: 26}
        heading_before = {1: 480, 2: 360, 3: 280, 4: 240}
        for lvl, sz in heading_sizes.items():
            self._write_style(w, f"Heading{lvl}", f"Heading {lvl}",
                              font=font, size=sz, bold=True,
                              outline_lvl=lvl - 1,
                              spacing_before=heading_before[lvl],
                              spacing_after=120)

        self._write_style(w, "BlockText", "Block Text",
                          font=font, size=size, italic=True,
                          indent_left=720, indent_right=720,
                          spacing_before=120, spacing_after=120)

        self._write_style(w, "SourceCode", "Source Code",
                          font="Courier New", size=20,
                          spacing_before=0, spacing_after=0)

        w.close("w:styles")
        return w.getvalue()

    def _write_style(self, w: XmlWriter, style_id: str, name: str,
                     is_default: bool = False,
                     font: str = "Calibri", size: int = 22,
                     bold: bool = False, italic: bool = False,
                     outline_lvl: int | None = None,
                     spacing_before: int = 0, spacing_after: int = 160,
                     indent_left: int = 0, indent_right: int = 0) -> None:
        attrs: dict = {"w:type": "paragraph", "w:styleId": style_id}
        if is_default:
            attrs["w:default"] = "1"
        w.open("w:style", attrs)
        w.tag("w:name", {"w:val": name})
        if not is_default:
            w.tag("w:basedOn", {"w:val": "Normal"})

        w.open("w:pPr")
        if outline_lvl is not None:
            w.tag("w:outlineLvl", {"w:val": str(outline_lvl)})
        sp: dict = {}
        if spacing_before: sp["w:before"] = str(spacing_before)
        if spacing_after:  sp["w:after"]  = str(spacing_after)
        if sp:
            w.tag("w:spacing", sp)
        if indent_left or indent_right:
            ind: dict = {}
            if indent_left:  ind["w:left"]  = str(indent_left)
            if indent_right: ind["w:right"] = str(indent_right)
            w.tag("w:ind", ind)
        w.close("w:pPr")

        w.open("w:rPr")
        w.tag("w:rFonts", {"w:ascii": font, "w:hAnsi": font, "w:cs": font})
        w.tag("w:sz",   {"w:val": str(size)})
        w.tag("w:szCs", {"w:val": str(size)})
        if bold:   w.tag("w:b", {})
        if italic: w.tag("w:i", {})
        w.close("w:rPr")

        w.close("w:style")

    # ------------------------------------------------------------------
    # Helper
    # ------------------------------------------------------------------

    def _next_shape_id(self) -> int:
        self._shape_id += 1
        return self._shape_id


# ---------------------------------------------------------------------------
# Namespaces and static files
# ---------------------------------------------------------------------------

_A_NS    = "http://schemas.openxmlformats.org/drawingml/2006/main"
_WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

_NS = {
    "xmlns:wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "xmlns:mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xmlns:w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "xmlns:a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
}

_CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

_DOC_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
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