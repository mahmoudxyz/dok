"""
dok.docx_packaging
~~~~~~~~~~~~~~~~~~
Static .docx packaging: XML namespace constants, relationships,
content types, numbering, and static XML parts.
"""

from __future__ import annotations

from pathlib import Path
from .xml_writer import XmlWriter


# ---------------------------------------------------------------------------
# XML Namespaces
# ---------------------------------------------------------------------------

W_NS     = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS     = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_URI  = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
PIC_URI  = "http://schemas.openxmlformats.org/drawingml/2006/picture"
RELS_NS  = "http://schemas.openxmlformats.org/package/2006/relationships"

DOCUMENT_NS = {
    "xmlns:wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "xmlns:mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:r":   R_NS,
    "xmlns:wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xmlns:w":   W_NS,
    "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "xmlns:a":   A_NS,
    "xmlns:pic": PIC_URI,
}


# ---------------------------------------------------------------------------
# Static XML parts
# ---------------------------------------------------------------------------

PACKAGE_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

SETTINGS_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word"
      w:val="15"/>
  </w:compat>
</w:settings>"""


def build_settings_xml(*, hyphenate: bool = False) -> str:
    """Build settings.xml with optional typography features."""
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '  <w:defaultTabStop w:val="720"/>',
    ]
    if hyphenate:
        parts.append('  <w:autoHyphenation/>')
        parts.append('  <w:consecutiveHyphenLimit w:val="3"/>')
        parts.append('  <w:hyphenationZone w:val="425"/>')
    parts.append('  <w:compat>')
    parts.append('    <w:compatSetting w:name="compatibilityMode"')
    parts.append('      w:uri="http://schemas.microsoft.com/office/word"')
    parts.append('      w:val="15"/>')
    parts.append('  </w:compat>')
    parts.append('</w:settings>')
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Numbering XML (bullet + ordered lists)
# ---------------------------------------------------------------------------

def build_numbering_xml(custom_markers: list[str] | None = None) -> str:
    """Build numbering.xml with bullet (numId=1), ordered (numId=2),
    and custom marker lists (numId=3+)."""
    w = XmlWriter()
    w.declaration()
    w.open("w:numbering", {"xmlns:w": W_NS})

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

    # Abstract 2+: custom marker lists
    for idx, marker_char in enumerate(custom_markers or []):
        abstract_id = str(idx + 2)
        w.open("w:abstractNum", {"w:abstractNumId": abstract_id})
        w.tag("w:multiLevelType", {"w:val": "hybridMultilevel"})
        for lvl in range(3):
            w.open("w:lvl", {"w:ilvl": str(lvl)})
            w.tag("w:start", {"w:val": "1"})
            w.tag("w:numFmt", {"w:val": "bullet"})
            w.tag("w:lvlText", {"w:val": marker_char})
            w.tag("w:lvlJc", {"w:val": "left"})
            w.open("w:pPr")
            indent = 720 * (lvl + 1)
            w.tag("w:ind", {"w:left": str(indent), "w:hanging": "360"})
            w.close("w:pPr")
            w.open("w:rPr")
            w.tag("w:rFonts", {"w:ascii": "Calibri", "w:hAnsi": "Calibri",
                                "w:cs": "Calibri", "w:hint": "default"})
            w.close("w:rPr")
            w.close("w:lvl")
        w.close("w:abstractNum")

    # Concrete instances
    w.open("w:num", {"w:numId": "1"})
    w.tag("w:abstractNumId", {"w:val": "0"})
    w.close("w:num")
    w.open("w:num", {"w:numId": "2"})
    w.tag("w:abstractNumId", {"w:val": "1"})
    w.close("w:num")
    for idx in range(len(custom_markers or [])):
        num_id = str(idx + 3)
        abstract_id = str(idx + 2)
        w.open("w:num", {"w:numId": num_id})
        w.tag("w:abstractNumId", {"w:val": abstract_id})
        w.close("w:num")

    w.close("w:numbering")
    return w.getvalue()


# ---------------------------------------------------------------------------
# Document relationships (word/_rels/document.xml.rels)
# ---------------------------------------------------------------------------

def build_doc_rels(
    image_entries: list[tuple[str, str, bytes]],
    hyperlink_rels: list[tuple[str, str]],
    header_rel_id: str | None = None,
    footer_rel_id: str | None = None,
    numbering_rel_id: str | None = None,
) -> str:
    """Build word/_rels/document.xml.rels."""
    w = XmlWriter()
    w.declaration()
    w.open("Relationships", {"xmlns": RELS_NS})

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

    for rel_id, filename, _ in image_entries:
        w.tag("Relationship", {"Id": rel_id,
              "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              "Target": f"media/{filename}"})

    for rel_id, url in hyperlink_rels:
        w.tag("Relationship", {"Id": rel_id,
              "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
              "Target": url, "TargetMode": "External"})

    w.close("Relationships")
    return w.getvalue()


# ---------------------------------------------------------------------------
# Content types ([Content_Types].xml)
# ---------------------------------------------------------------------------

def build_content_types(
    image_entries: list[tuple[str, str, bytes]],
    has_header: bool = False,
    has_footer: bool = False,
    has_numbering: bool = False,
) -> str:
    """Build [Content_Types].xml."""
    w = XmlWriter()
    w.declaration()
    w.open("Types", {"xmlns": "http://schemas.openxmlformats.org/package/2006/content-types"})

    w.tag("Default", {"Extension": "rels",
          "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
    w.tag("Default", {"Extension": "xml", "ContentType": "application/xml"})

    # Image types
    exts_seen: set[str] = set()
    for _, filename, _ in image_entries:
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

    if has_numbering:
        w.tag("Override", {"PartName": "/word/numbering.xml",
              "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"})
    if has_header:
        w.tag("Override", {"PartName": "/word/header1.xml",
              "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"})
    if has_footer:
        w.tag("Override", {"PartName": "/word/footer1.xml",
              "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"})

    w.close("Types")
    return w.getvalue()
