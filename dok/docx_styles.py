"""
dok.docx_styles
~~~~~~~~~~~~~~~
Builds the styles.xml part of the .docx package.

Pure functions — no dependency on DocxWriter state.
"""

from __future__ import annotations

from .constants import HYPERLINK_COLOR
from .models   import SPACING_PRESETS
from .xml_writer import XmlWriter

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def build_styles_xml(
    font: str = "Calibri",
    size_pt: int = 11,
    spacing: str = "normal",
) -> str:
    """Build the complete styles.xml content for a .docx file."""
    size = size_pt * 2  # half-points
    para_after, h_scale, line_sp = SPACING_PRESETS.get(
        spacing, SPACING_PRESETS["normal"])

    w = XmlWriter()
    w.declaration()
    w.open("w:styles", {"xmlns:w": _W_NS})

    # --- docDefaults ---
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

    # --- Normal ---
    _write_style(w, "Normal", "Normal", is_default=True, font=font, size=size,
                 spacing_after=para_after, line_spacing=line_sp)

    # --- Headings ---
    for lvl, sz in {1: 36, 2: 32, 3: 28, 4: 26}.items():
        clr = "1F3864" if lvl <= 2 else "404040"
        before = int({1: 480, 2: 360, 3: 280, 4: 240}[lvl] * h_scale)
        after  = int(120 * h_scale)
        _write_style(w, f"Heading{lvl}", f"Heading {lvl}",
                     font=font, size=sz, bold=True, color=clr,
                     outline_lvl=lvl - 1, spacing_before=before,
                     spacing_after=after, line_spacing=min(line_sp, 240))

    # --- BlockText (quote) ---
    _write_style(w, "BlockText", "Block Text", font=font, size=size,
                 italic=True, color="404040", indent_left=720, indent_right=720,
                 spacing_before=int(120 * h_scale),
                 spacing_after=int(120 * h_scale), line_spacing=line_sp)

    # --- SourceCode ---
    _write_style(w, "SourceCode", "Source Code", font="Courier New", size=20,
                 spacing_before=40, spacing_after=40, line_spacing=240,
                 shading="F5F5F5", border_color="E0E0E0")

    # --- List Paragraph ---
    list_after = max(0, min(para_after, 40))
    _write_style(w, "ListParagraph", "List Paragraph", font=font, size=size,
                 indent_left=720, spacing_after=list_after, line_spacing=line_sp)

    # --- Hyperlink character style ---
    w.open("w:style", {"w:type": "character", "w:styleId": "Hyperlink"})
    w.tag("w:name", {"w:val": "Hyperlink"})
    w.open("w:rPr")
    w.tag("w:color", {"w:val": HYPERLINK_COLOR})
    w.tag("w:u", {"w:val": "single"})
    w.close("w:rPr")
    w.close("w:style")

    w.close("w:styles")
    return w.getvalue()


def _write_style(
    w: XmlWriter, style_id: str, name: str,
    is_default: bool = False,
    font: str = "Calibri", size: int = 22,
    bold: bool = False, italic: bool = False,
    color: str | None = None,
    outline_lvl: int | None = None,
    spacing_before: int = 0, spacing_after: int = 160,
    indent_left: int = 0, indent_right: int = 0,
    line_spacing: int = 276,
    shading: str | None = None,
    border_color: str | None = None,
) -> None:
    """Write a single paragraph style element."""
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
    sp: dict = {"w:line": str(line_spacing), "w:lineRule": "auto"}
    if spacing_before:
        sp["w:before"] = str(spacing_before)
    if spacing_after:
        sp["w:after"] = str(spacing_after)
    w.tag("w:spacing", sp)
    if indent_left or indent_right:
        ind: dict = {}
        if indent_left:
            ind["w:left"] = str(indent_left)
        if indent_right:
            ind["w:right"] = str(indent_right)
        w.tag("w:ind", ind)
    if shading:
        w.tag("w:shd", {"w:val": "clear", "w:color": "auto", "w:fill": shading})
    if border_color:
        w.open("w:pBdr")
        for side in ("top", "left", "bottom", "right"):
            w.tag(f"w:{side}", {"w:val": "single", "w:sz": "4",
                                "w:space": "4", "w:color": border_color})
        w.close("w:pBdr")
    w.close("w:pPr")

    w.open("w:rPr")
    w.tag("w:rFonts", {"w:ascii": font, "w:hAnsi": font, "w:cs": font})
    w.tag("w:sz", {"w:val": str(size)}); w.tag("w:szCs", {"w:val": str(size)})
    if bold:
        w.tag("w:b", {})
    if italic:
        w.tag("w:i", {})
    if color:
        w.tag("w:color", {"w:val": color})
    w.close("w:rPr")
    w.close("w:style")
