"""
Integration tests for dok — end-to-end from source to DOCX.

Tests individual elements, mixed element combinations, measurements,
and validates the generated DOCX structure.
"""
import pytest
import zipfile
import io
import xml.etree.ElementTree as ET

import dok
from dok.nodes import ElementNode, TextNode
from dok.converter import Converter
from dok.models import (
    ParagraphModel, RunModel, BoxModel, LineModel, SpacerModel,
    DataTableModel, TableModel, ImageModel, HeaderModel, FooterModel,
    PageBreakModel, ShapeModel, RowModel, TocModel,
    CheckboxModel, TextInputModel, DropdownModel,
    FrameModel, ToggleModel, DocxModel, SectionModel,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def parse_and_bytes(source: str) -> bytes:
    """Parse source and produce DOCX bytes."""
    node = dok.parse(source)
    return dok.to_bytes(node)


def parse_and_model(source: str) -> DocxModel:
    """Parse source and produce the intermediate DocxModel."""
    node = dok.parse(source)
    return Converter().convert([node])


def docx_document_xml(data: bytes) -> ET.Element:
    """Extract and parse document.xml from DOCX bytes."""
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        xml_bytes = zf.read("word/document.xml")
    return ET.fromstring(xml_bytes)


def docx_files(data: bytes) -> list[str]:
    """List all files in the DOCX ZIP."""
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return zf.namelist()


def find_elements(root: ET.Element, tag: str) -> list[ET.Element]:
    """Find all elements with the given local tag name (ignoring namespace)."""
    results = []
    for elem in root.iter():
        local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if local == tag:
            results.append(elem)
    return results


def get_text_content(root: ET.Element) -> str:
    """Extract all text content from XML."""
    texts = []
    for elem in root.iter():
        local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if local == "t" and elem.text:
            texts.append(elem.text)
    return "".join(texts)


# ===========================================================================
# INDIVIDUAL ELEMENT TESTS
# ===========================================================================


class TestParagraph:
    def test_simple_paragraph(self):
        data = parse_and_bytes('p { "Hello world" }')
        root = docx_document_xml(data)
        assert "Hello world" in get_text_content(root)

    def test_paragraph_alignment_center(self):
        model = parse_and_model('center { p { "Centered" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].align == "center"

    def test_paragraph_alignment_right(self):
        model = parse_and_model('right { p { "Right" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].align == "right"

    def test_paragraph_alignment_justify(self):
        model = parse_and_model('justify { p { "Justified" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].align == "justify"

    def test_paragraph_rtl(self):
        model = parse_and_model('rtl { p { "مرحبا" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].direction == "rtl"

    def test_paragraph_ltr(self):
        model = parse_and_model('ltr { p { "Hello" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].direction == "ltr"


class TestHeadings:
    def test_h1(self):
        model = parse_and_model('h1 { "Title" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].style == "Heading1"

    def test_h2(self):
        model = parse_and_model('h2 { "Section" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].style == "Heading2"

    def test_h3(self):
        model = parse_and_model('h3 { "Subsection" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].style == "Heading3"

    def test_h4(self):
        model = parse_and_model('h4 { "Minor" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].style == "Heading4"

    def test_heading_generates_bookmark(self):
        model = parse_and_model('h1 { "Bookmarked" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].bookmark is not None

    def test_heading_with_explicit_id(self):
        model = parse_and_model('h1(id: "my-id") { "Custom ID" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].bookmark == "my-id"

    def test_heading_to_docx(self):
        data = parse_and_bytes('h1 { "Big Title" }')
        root = docx_document_xml(data)
        assert "Big Title" in get_text_content(root)
        # Should have Heading1 style
        styles = find_elements(root, "pStyle")
        assert any(s.get(f"{{{W_NS}}}val") == "Heading1" for s in styles)


class TestTextStyles:
    def test_bold(self):
        model = parse_and_model('p { bold { "Strong" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].bold is True
        assert paras[0].runs[0].text == "Strong"

    def test_italic(self):
        model = parse_and_model('p { italic { "Emphasis" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].italic is True

    def test_underline(self):
        model = parse_and_model('p { underline { "Underlined" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].underline is True

    def test_strike(self):
        model = parse_and_model('p { strike { "Struck" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].strike is True

    def test_superscript(self):
        model = parse_and_model('p { sup { "2" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].sup is True

    def test_subscript(self):
        model = parse_and_model('p { sub { "n" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].sub is True

    def test_color(self):
        model = parse_and_model('p { color(value: red) { "Red" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].color is not None

    def test_font_size(self):
        model = parse_and_model('p { size(value: 24) { "Big" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].size_pt == 24

    def test_font_family(self):
        model = parse_and_model('p { font(value: "Arial") { "Arial" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].font == "Arial"

    def test_highlight(self):
        model = parse_and_model('p { highlight(value: yellow) { "Highlighted" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].highlight == "yellow"

    def test_nested_bold_italic(self):
        model = parse_and_model('p { bold { italic { "Both" } } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        run = paras[0].runs[0]
        assert run.bold is True
        assert run.italic is True
        assert run.text == "Both"

    def test_bold_with_color_prop(self):
        model = parse_and_model('p { bold(color: navy) { "Navy bold" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        run = paras[0].runs[0]
        assert run.bold is True
        assert run.color is not None

    def test_bold_with_size_prop(self):
        model = parse_and_model('p { bold(size: 18) { "Big bold" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        run = paras[0].runs[0]
        assert run.bold is True
        assert run.size_pt == 18

    def test_all_styles_to_docx(self):
        data = parse_and_bytes('''
            p {
                bold { "B " }
                italic { "I " }
                underline { "U " }
                strike { "S " }
                sup { "sup" }
                sub { "sub" }
            }
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "B " in text
        assert "I " in text


class TestLists:
    def test_unordered_list_model(self):
        model = parse_and_model('ul { li { "A" } li { "B" } }')
        assert model.has_lists is True
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert len(paras) == 2
        assert all(p.num_id == 1 for p in paras)

    def test_ordered_list_model(self):
        model = parse_and_model('ol { li { "1" } li { "2" } }')
        assert model.has_lists is True
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert all(p.num_id == 2 for p in paras)

    def test_custom_bullet_marker(self):
        model = parse_and_model('ul(marker: "→") { li { "A" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].list_marker == "→"
        assert "→" in model.custom_markers

    def test_ordered_list_alpha(self):
        model = parse_and_model('ol(marker: alpha) { li { "A" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].list_format == "alpha"

    def test_ordered_list_roman(self):
        model = parse_and_model('ol(marker: roman) { li { "I" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].list_format == "roman"

    def test_nested_list(self):
        model = parse_and_model('''
            ul {
                li { "Parent" }
                ul {
                    li { "Child" }
                }
            }
        ''')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert len(paras) == 2
        # Child should be at a deeper level
        assert paras[1].num_ilvl > paras[0].num_ilvl

    def test_list_to_docx(self):
        data = parse_and_bytes('ul { li { "A" } li { "B" } li { "C" } }')
        files = docx_files(data)
        assert "word/numbering.xml" in files

    def test_ordered_list_to_docx(self):
        data = parse_and_bytes('ol { li { "First" } li { "Second" } }')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "First" in text
        assert "Second" in text


class TestDataTables:
    def test_simple_table(self):
        model = parse_and_model('''
            table {
                tr { td { "A" } td { "B" } }
                tr { td { "C" } td { "D" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, DataTableModel)]
        assert len(tables) == 1
        assert len(tables[0].rows) == 2
        assert len(tables[0].rows[0].cells) == 2

    def test_table_with_headers(self):
        model = parse_and_model('''
            table {
                tr { th { "Name" } th { "Value" } }
                tr { td { "A" } td { "1" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, DataTableModel)]
        assert tables[0].rows[0].is_header is True
        assert tables[0].rows[0].cells[0].is_th is True

    def test_table_border_prop(self):
        model = parse_and_model('table(border: true) { tr { td { "X" } } }')
        tables = [c for c in model.content if isinstance(c, DataTableModel)]
        assert tables[0].border is True

    def test_table_striped_prop(self):
        model = parse_and_model('table(striped: true) { tr { td { "X" } } }')
        tables = [c for c in model.content if isinstance(c, DataTableModel)]
        assert tables[0].striped is True

    def test_table_col_widths_auto(self):
        model = parse_and_model('''
            table {
                tr { td { "Short" } td { "This is a much longer text" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, DataTableModel)]
        widths = tables[0].col_widths
        assert len(widths) == 2
        assert sum(widths) == 100
        # Longer column should get more width
        assert widths[1] > widths[0]

    def test_table_to_docx(self):
        data = parse_and_bytes('''
            table(border: true) {
                tr { th { "H1" } th { "H2" } }
                tr { td { "A" } td { "B" } }
            }
        ''')
        root = docx_document_xml(data)
        tbls = find_elements(root, "tbl")
        assert len(tbls) >= 1
        text = get_text_content(root)
        assert "H1" in text
        assert "A" in text


class TestColumns:
    def test_two_columns(self):
        model = parse_and_model('''
            cols(ratio: 1:1) {
                col { p { "Left" } }
                col { p { "Right" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, TableModel)]
        assert len(tables) == 1
        assert len(tables[0].rows[0].cells) == 2

    def test_three_columns(self):
        model = parse_and_model('''
            cols(ratio: 1:2:1) {
                col { p { "A" } }
                col { p { "B" } }
                col { p { "C" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, TableModel)]
        cells = tables[0].rows[0].cells
        assert len(cells) == 3
        # Middle column should be wider
        assert cells[1].width_pct > cells[0].width_pct

    def test_columns_with_gap(self):
        model = parse_and_model('''
            cols(ratio: 1:1, gap: 12) {
                col { p { "A" } }
                col { p { "B" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, TableModel)]
        assert tables[0].gap_twips > 0

    def test_columns_with_padding(self):
        model = parse_and_model('''
            cols(ratio: 1:1, padding: 8) {
                col { p { "A" } }
                col { p { "B" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, TableModel)]
        assert tables[0].cell_padding_twips > 0

    def test_column_fill(self):
        model = parse_and_model('''
            cols(ratio: 1:1) {
                col(fill: lightgray) { p { "Shaded" } }
                col { p { "Normal" } }
            }
        ''')
        tables = [c for c in model.content if isinstance(c, TableModel)]
        assert tables[0].rows[0].cells[0].fill is not None

    def test_columns_to_docx(self):
        data = parse_and_bytes('''
            cols(ratio: 2:1) {
                col { p { "Main content" } }
                col { p { "Sidebar" } }
            }
        ''')
        root = docx_document_xml(data)
        assert "Main content" in get_text_content(root)


class TestBoxes:
    def test_simple_box(self):
        model = parse_and_model('box { p { "Content" } }')
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert len(boxes) == 1

    def test_box_with_fill(self):
        model = parse_and_model('box(fill: lightblue) { p { "Blue" } }')
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert boxes[0].fill is not None

    def test_box_with_stroke(self):
        model = parse_and_model('box(stroke: red) { p { "Red border" } }')
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert boxes[0].stroke is not None

    def test_banner(self):
        model = parse_and_model('banner { p { "Notice" } }')
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert len(boxes) == 1
        assert boxes[0].accent is not None

    def test_callout(self):
        model = parse_and_model('callout { p { "Warning" } }')
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert len(boxes) == 1

    def test_badge(self):
        model = parse_and_model('badge { "OK" }')
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert len(boxes) == 1
        assert boxes[0].inline is True
        assert boxes[0].text == "OK"

    def test_box_selective_borders(self):
        model = parse_and_model('''
            box(stroke: black, border-top: false, border-right: false) {
                p { "Bottom-left only" }
            }
        ''')
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert boxes[0].border_top is False
        assert boxes[0].border_right is False
        assert boxes[0].border_bottom is True
        assert boxes[0].border_left is True

    def test_box_to_docx(self):
        data = parse_and_bytes('box(fill: #E0E0E0, stroke: black) { p { "Boxed" } }')
        root = docx_document_xml(data)
        assert "Boxed" in get_text_content(root)


class TestLine:
    def test_simple_line(self):
        model = parse_and_model('line { }')
        lines = [c for c in model.content if isinstance(c, LineModel)]
        assert len(lines) == 1

    def test_line_with_color(self):
        model = parse_and_model('line(stroke: red) { }')
        lines = [c for c in model.content if isinstance(c, LineModel)]
        assert lines[0].color != "BFBFBF"

    def test_dashed_line(self):
        model = parse_and_model('line(dashed) { }')
        lines = [c for c in model.content if isinstance(c, LineModel)]
        assert lines[0].style == "dashed"

    def test_thick_line(self):
        model = parse_and_model('line(thick) { }')
        lines = [c for c in model.content if isinstance(c, LineModel)]
        assert lines[0].thick is True


class TestSpacer:
    def test_default_spacer(self):
        model = parse_and_model('space { }')
        spacers = [c for c in model.content if isinstance(c, SpacerModel)]
        assert len(spacers) == 1
        assert spacers[0].height_twips > 0

    def test_spacer_with_size(self):
        model = parse_and_model('space(size: 24) { }')
        spacers = [c for c in model.content if isinstance(c, SpacerModel)]
        assert spacers[0].height_twips == 480  # 24pt = 480 twips

    def test_spacer_with_unit(self):
        model = parse_and_model('space(size: 72) { }')
        spacers = [c for c in model.content if isinstance(c, SpacerModel)]
        assert spacers[0].height_twips == 1440  # 72pt = 1 inch = 1440 twips


class TestPageBreak:
    def test_page_break(self):
        model = parse_and_model('p { "Before" } --- p { "After" }')
        breaks = [c for c in model.content if isinstance(c, PageBreakModel)]
        assert len(breaks) == 1


class TestHyperlinks:
    def test_link_model(self):
        model = parse_and_model('p { link(href: "https://example.com") { "Click" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert any(r.hyperlink_url == "https://example.com" for r in paras[0].runs)

    def test_link_styling(self):
        model = parse_and_model('p { link(href: "https://example.com") { "Click" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        link_run = [r for r in paras[0].runs if r.hyperlink_url][0]
        assert link_run.underline is True
        assert link_run.color is not None

    def test_link_to_docx(self):
        data = parse_and_bytes('link(href: "https://example.com") { "Click" }')
        root = docx_document_xml(data)
        hyperlinks = find_elements(root, "hyperlink")
        assert len(hyperlinks) >= 1


class TestHeaderFooter:
    def test_header(self):
        model = parse_and_model('doc { header { p { "My Header" } } p { "Body" } }')
        assert model.header is not None
        assert len(model.header.paragraphs) >= 1

    def test_footer(self):
        model = parse_and_model('doc { footer { p { "My Footer" } } p { "Body" } }')
        assert model.footer is not None

    def test_header_in_docx(self):
        data = parse_and_bytes('doc { header { p { "Header Text" } } p { "Body" } }')
        files = docx_files(data)
        assert "word/header1.xml" in files

    def test_footer_in_docx(self):
        data = parse_and_bytes('doc { footer { p { "Footer Text" } } p { "Body" } }')
        files = docx_files(data)
        assert "word/footer1.xml" in files


class TestPageNumber:
    def test_page_number_model(self):
        model = parse_and_model('page-number { }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert any(r.field == "PAGE" for r in paras[0].runs)


class TestQuote:
    def test_quote_style(self):
        model = parse_and_model('quote { "Wise words" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].style == "BlockText"


class TestCode:
    def test_code_style(self):
        model = parse_and_model('code { "print(hello)" }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].style == "SourceCode"

    def test_multiline_code(self):
        data = parse_and_bytes('code { "line1\\nline2\\nline3" }')
        root = docx_document_xml(data)
        # Should split into multiple paragraphs
        paras = find_elements(root, "p")
        assert len(paras) >= 3


class TestIndent:
    def test_indent_level_1(self):
        model = parse_and_model('indent { p { "Indented" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].indent_twips > 0

    def test_indent_level_2(self):
        model = parse_and_model('indent(level: 2) { p { "Deep" } }')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].indent_twips > 720  # > level 1


class TestDrawingShapes:
    def test_circle(self):
        model = parse_and_model('circle(fill: navy) { }')
        shapes = [c for c in model.content if isinstance(c, ShapeModel)]
        assert len(shapes) == 1
        assert shapes[0].preset == "ellipse"

    def test_diamond(self):
        model = parse_and_model('diamond(fill: gold) { }')
        shapes = [c for c in model.content if isinstance(c, ShapeModel)]
        assert shapes[0].preset == "diamond"

    def test_chevron(self):
        model = parse_and_model('chevron(fill: red) { }')
        shapes = [c for c in model.content if isinstance(c, ShapeModel)]
        assert shapes[0].preset == "chevron"


class TestFormFields:
    def test_checkbox(self):
        model = parse_and_model('checkbox(checked: true, label: "Agree") { }')
        checks = [c for c in model.content if isinstance(c, CheckboxModel)]
        assert len(checks) == 1
        assert checks[0].checked is True
        assert checks[0].label == "Agree"

    def test_text_input(self):
        model = parse_and_model('text-input(placeholder: "Name") { }')
        inputs = [c for c in model.content if isinstance(c, TextInputModel)]
        assert len(inputs) == 1
        assert inputs[0].placeholder == "Name"

    def test_dropdown(self):
        model = parse_and_model('''
            dropdown(value: "a") {
                option(value: "a") { }
                option(value: "b") { }
            }
        ''')
        drops = [c for c in model.content if isinstance(c, DropdownModel)]
        assert len(drops) == 1
        assert len(drops[0].options) == 2


class TestFrame:
    def test_frame_model(self):
        model = parse_and_model('frame(x: 50, y: 50, width: 200, height: 100) { p { "Float" } }')
        frames = [c for c in model.content if isinstance(c, FrameModel)]
        assert len(frames) == 1
        assert frames[0].width_twips == 4000  # 200pt = 4000 twips
        assert frames[0].height_twips == 2000


class TestToggle:
    def test_toggle_model(self):
        model = parse_and_model('toggle(title: "Details", open: true) { p { "Content" } }')
        toggles = [c for c in model.content if isinstance(c, ToggleModel)]
        assert len(toggles) == 1
        assert toggles[0].title == "Details"
        assert toggles[0].open is True


class TestToc:
    def test_toc_collects_headings(self):
        model = parse_and_model('''
            toc(depth: 3) { }
            h1 { "Chapter 1" }
            h2 { "Section 1.1" }
            h3 { "Sub 1.1.1" }
            h4 { "Deep" }
        ''')
        tocs = [c for c in model.content if isinstance(c, TocModel)]
        assert len(tocs) == 1
        assert len(tocs[0].entries) == 3  # depth:3 excludes h4

    def test_toc_entry_levels(self):
        model = parse_and_model('''
            toc { }
            h1 { "One" }
            h2 { "Two" }
        ''')
        tocs = [c for c in model.content if isinstance(c, TocModel)]
        entries = tocs[0].entries
        assert entries[0].level == 1
        assert entries[1].level == 2


class TestRef:
    def test_internal_ref(self):
        model = parse_and_model('''
            h1(id: "section1") { "Section 1" }
            ref(to: "section1") { "See section 1" }
        ''')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        ref_para = paras[-1]
        assert any(r.hyperlink_url == "#section1" for r in ref_para.runs)


# ===========================================================================
# MIXED ELEMENT INTEGRATION TESTS
# ===========================================================================


class TestMixedElements:
    def test_heading_with_styled_text(self):
        data = parse_and_bytes('h1 { bold { "Important " } "Title" }')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Important " in text
        assert "Title" in text

    def test_paragraph_with_link(self):
        data = parse_and_bytes('''
            p { "Visit " link(href: "https://example.com") { "here" } " for info." }
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Visit " in text
        assert "here" in text

    def test_list_with_bold_items(self):
        model = parse_and_model('''
            ul {
                li { bold { "Important: " } "First item" }
                li { "Normal item" }
            }
        ''')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].bold is True
        assert paras[0].runs[1].bold is False

    def test_table_with_styled_cells(self):
        data = parse_and_bytes('''
            table(border: true) {
                tr { th { "Name" } th { "Status" } }
                tr { td { bold { "Widget A" } } td { color(value: green) { "Active" } } }
            }
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Widget A" in text
        assert "Active" in text

    def test_columns_with_lists(self):
        data = parse_and_bytes('''
            cols(ratio: 1:1) {
                col {
                    h2 { "Features" }
                    ul { li { "Fast" } li { "Simple" } }
                }
                col {
                    h2 { "Benefits" }
                    ol { li { "Saves time" } li { "Reduces cost" } }
                }
            }
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Fast" in text
        assert "Saves time" in text

    def test_box_with_table(self):
        data = parse_and_bytes('''
            box(fill: #F0F0F0) {
                h3 { "Summary" }
                table(border: true) {
                    tr { th { "Metric" } th { "Value" } }
                    tr { td { "Score" } td { "95%" } }
                }
            }
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Summary" in text
        assert "Score" in text

    def test_full_document_structure(self):
        """Test a realistic document with multiple element types."""
        data = parse_and_bytes('''
            doc(font: "Calibri", size: 11) {
                header { center { p { "Company Report" } } }
                footer { center { page-number { } } }

                center { h1 { "Annual Report 2024" } }
                line { }
                p { "This report covers our annual performance." }

                h2 { "Key Metrics" }
                table(border: true) {
                    tr { th { "Metric" } th { "Value" } th { "Change" } }
                    tr { td { "Revenue" } td { "$1.2M" } td { color(value: green) { "+15%" } } }
                    tr { td { "Users" } td { "50K" } td { color(value: green) { "+25%" } } }
                }

                space(size: 12) { }

                h2 { "Highlights" }
                ul {
                    li { bold { "Growth: " } "Revenue up 15%" }
                    li { bold { "Expansion: " } "3 new markets" }
                    li { bold { "Team: " } "Hired 20 engineers" }
                }

                ---

                h2 { "Looking Forward" }
                cols(ratio: 1:1) {
                    col {
                        h3 { "Goals" }
                        ol { li { "Double revenue" } li { "Enter 5 markets" } }
                    }
                    col {
                        h3 { "Risks" }
                        ul { li { "Competition" } li { "Regulation" } }
                    }
                }
            }
        ''')
        assert zipfile.is_zipfile(io.BytesIO(data))
        files = docx_files(data)
        assert "word/document.xml" in files
        assert "word/header1.xml" in files
        assert "word/footer1.xml" in files
        assert "word/numbering.xml" in files

        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Annual Report 2024" in text
        assert "Revenue" in text
        assert "Double revenue" in text

    def test_rtl_mixed_document(self):
        """Test RTL document with mixed content."""
        data = parse_and_bytes('''
            doc {
                rtl {
                    h1 { "تقرير" }
                    p { "هذا تقرير المالي" }
                }
                ltr {
                    p { bold { "Q4 2024" } }
                    p { "English section" }
                }
            }
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "تقرير" in text
        assert "Q4 2024" in text
        assert "English section" in text


# ===========================================================================
# MEASUREMENT INTEGRATION TESTS
# ===========================================================================


class TestMeasurements:
    def test_spacer_pt(self):
        model = parse_and_model('space(size: 36) { }')
        spacers = [c for c in model.content if isinstance(c, SpacerModel)]
        assert spacers[0].height_twips == 720  # 36pt = 720 twips

    def test_margin_values(self):
        model = parse_and_model('page(margin-top: 72, margin-bottom: 72) { p { "X" } }')
        section = model.sections[-1]
        assert section.margin_top == 1440  # 72pt = 1 inch = 1440 twips

    def test_image_default_size(self):
        model = parse_and_model('img(src: "test.png") { }')
        images = [c for c in model.content if isinstance(c, ImageModel)]
        assert images[0].width_emu > 0
        assert images[0].height_emu > 0

    def test_image_custom_size(self):
        model = parse_and_model('img(src: "test.png", width: 144, height: 72) { }')
        images = [c for c in model.content if isinstance(c, ImageModel)]
        # 144pt = 2 inches = 1828800 EMU
        assert images[0].width_emu == 1828800
        assert images[0].height_emu == 914400

    def test_frame_dimensions(self):
        model = parse_and_model('frame(x: 72, y: 36, width: 144, height: 72) { p { "F" } }')
        frames = [c for c in model.content if isinstance(c, FrameModel)]
        assert frames[0].x_twips == 1440    # 72pt
        assert frames[0].y_twips == 720     # 36pt
        assert frames[0].width_twips == 2880   # 144pt
        assert frames[0].height_twips == 1440  # 72pt


# ===========================================================================
# DOCX STRUCTURE VALIDATION TESTS
# ===========================================================================


class TestDocxStructure:
    def test_valid_zip(self):
        data = parse_and_bytes('p { "Test" }')
        assert zipfile.is_zipfile(io.BytesIO(data))

    def test_required_files(self):
        data = parse_and_bytes('p { "Test" }')
        files = docx_files(data)
        assert "[Content_Types].xml" in files
        assert "_rels/.rels" in files
        assert "word/document.xml" in files
        assert "word/styles.xml" in files
        assert "word/settings.xml" in files
        assert "word/_rels/document.xml.rels" in files

    def test_content_types_xml(self):
        data = parse_and_bytes('p { "Test" }')
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            ct = zf.read("[Content_Types].xml").decode()
        assert "application/vnd.openxmlformats" in ct

    def test_document_xml_structure(self):
        data = parse_and_bytes('p { "Hello" }')
        root = docx_document_xml(data)
        body = find_elements(root, "body")
        assert len(body) == 1

    def test_styles_xml(self):
        data = parse_and_bytes('p { "Test" }')
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            styles = zf.read("word/styles.xml").decode()
        assert "Normal" in styles
        assert "Heading1" in styles

    def test_paragraph_xml_has_text(self):
        data = parse_and_bytes('p { "Specific text" }')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Specific text" in text

    def test_bold_xml(self):
        data = parse_and_bytes('p { bold { "Bold" } }')
        root = docx_document_xml(data)
        bolds = find_elements(root, "b")
        assert len(bolds) >= 1

    def test_italic_xml(self):
        data = parse_and_bytes('p { italic { "Italic" } }')
        root = docx_document_xml(data)
        italics = find_elements(root, "i")
        assert len(italics) >= 1

    def test_underline_xml(self):
        data = parse_and_bytes('p { underline { "UL" } }')
        root = docx_document_xml(data)
        uls = find_elements(root, "u")
        assert len(uls) >= 1

    def test_hyperlink_xml(self):
        data = parse_and_bytes('link(href: "https://example.com") { "Link" }')
        root = docx_document_xml(data)
        hyperlinks = find_elements(root, "hyperlink")
        assert len(hyperlinks) >= 1

    def test_bookmark_xml(self):
        data = parse_and_bytes('h1(id: "test") { "Bookmarked" }')
        root = docx_document_xml(data)
        starts = find_elements(root, "bookmarkStart")
        assert len(starts) >= 1

    def test_page_break_xml(self):
        data = parse_and_bytes('p { "Before" } --- p { "After" }')
        root = docx_document_xml(data)
        breaks = find_elements(root, "br")
        page_breaks = [b for b in breaks if b.get(f"{{{W_NS}}}type") == "page"]
        assert len(page_breaks) >= 1


# ===========================================================================
# DOC DEFAULTS AND TYPOGRAPHY TESTS
# ===========================================================================


class TestDocDefaults:
    def test_custom_font(self):
        model = parse_and_model('doc(font: "Arial") { p { "X" } }')
        assert model.default_font == "Arial"

    def test_custom_size(self):
        model = parse_and_model('doc(size: 14) { p { "X" } }')
        assert model.default_size_pt == 14

    def test_spacing_compact(self):
        model = parse_and_model('doc(spacing: compact) { p { "X" } }')
        assert model.spacing == "compact"

    def test_spacing_relaxed(self):
        model = parse_and_model('doc(spacing: relaxed) { p { "X" } }')
        assert model.spacing == "relaxed"

    def test_kerning(self):
        model = parse_and_model('doc(kerning: true) { p { "X" } }')
        assert model.kerning is True

    def test_ligatures(self):
        model = parse_and_model('doc(ligatures: true) { p { "X" } }')
        assert model.ligatures is True

    def test_hyphenate(self):
        model = parse_and_model('doc(hyphenate: true) { p { "X" } }')
        assert model.hyphenate is True


class TestPageSettings:
    def test_paper_a4(self):
        model = parse_and_model('page(paper: a4) { p { "X" } }')
        section = model.sections[-1]
        assert section.paper == "a4"

    def test_paper_letter(self):
        model = parse_and_model('page(paper: letter) { p { "X" } }')
        section = model.sections[-1]
        assert section.paper == "letter"

    def test_margin_narrow(self):
        model = parse_and_model('page(margin: narrow) { p { "X" } }')
        section = model.sections[-1]
        assert section.margin == "narrow"

    def test_columns(self):
        model = parse_and_model('page(cols: 2) { p { "X" } }')
        section = model.sections[-1]
        assert section.cols == 2


# ===========================================================================
# SUGAR SYNTAX INTEGRATION TESTS
# ===========================================================================


class TestSugarIntegration:
    def test_heading_sugar_to_docx(self):
        data = parse_and_bytes("# My Document Title")
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "My Document Title" in text
        styles = find_elements(root, "pStyle")
        assert any(s.get(f"{{{W_NS}}}val") == "Heading1" for s in styles)

    def test_bullet_sugar_to_docx(self):
        data = parse_and_bytes("- Apple\n- Banana\n- Cherry")
        files = docx_files(data)
        assert "word/numbering.xml" in files
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Apple" in text
        assert "Cherry" in text

    def test_numbered_sugar_to_docx(self):
        data = parse_and_bytes("1. First step\n2. Second step")
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "First step" in text

    def test_quote_sugar_to_docx(self):
        data = parse_and_bytes("> Wise words from someone")
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Wise words from someone" in text

    def test_mixed_sugar_and_dok(self):
        """Sugar and regular dok syntax mixed together."""
        data = parse_and_bytes('''
# Introduction

p { bold { "Welcome" } " to the doc." }

- Feature one
- Feature two

> Important note

table(border: true) {
    tr { th { "Name" } th { "Value" } }
    tr { td { "Alpha" } td { "100" } }
}
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Introduction" in text
        assert "Welcome" in text
        assert "Feature one" in text
        assert "Important note" in text
        assert "Alpha" in text

    def test_inline_sugar_in_paragraph(self):
        data = parse_and_bytes('p { "Hello **bold** and *italic* text" }')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "bold" in text
        assert "italic" in text
        bolds = find_elements(root, "b")
        assert len(bolds) >= 1


# ===========================================================================
# EDGE CASES AND ERROR HANDLING
# ===========================================================================


class TestEdgeCases:
    def test_empty_document(self):
        data = parse_and_bytes('doc { }')
        assert zipfile.is_zipfile(io.BytesIO(data))

    def test_deeply_nested(self):
        data = parse_and_bytes('''
            center { bold { italic { underline { p { "Deep" } } } } }
        ''')
        root = docx_document_xml(data)
        assert "Deep" in get_text_content(root)

    def test_empty_paragraph(self):
        data = parse_and_bytes('p { "" }')
        assert zipfile.is_zipfile(io.BytesIO(data))

    def test_special_characters_in_text(self):
        data = parse_and_bytes('p { "Hello <world> & \\"friends\\"" }')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "<world>" in text
        assert "&" in text

    def test_unicode_text(self):
        data = parse_and_bytes('p { "Hello 世界 مرحبا 🌍" }')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "世界" in text

    def test_very_long_text(self):
        long_text = "A" * 10000
        data = parse_and_bytes(f'p {{ "{long_text}" }}')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert len(text) >= 10000

    def test_many_paragraphs(self):
        paras = "\n".join(f'p {{ "Paragraph {i}" }}' for i in range(100))
        data = parse_and_bytes(paras)
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Paragraph 0" in text
        assert "Paragraph 99" in text

    def test_multiple_tables(self):
        data = parse_and_bytes('''
            table { tr { td { "T1" } } }
            table { tr { td { "T2" } } }
        ''')
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "T1" in text
        assert "T2" in text

    def test_mixed_runs_in_paragraph(self):
        """Multiple styled runs in a single paragraph."""
        model = parse_and_model('''
            p {
                "Normal "
                bold { "bold " }
                italic { "italic " }
                underline { "underlined " }
                color(value: red) { "red " }
                size(value: 20) { "big" }
            }
        ''')
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert len(paras) == 1
        runs = paras[0].runs
        assert len(runs) >= 5
        assert runs[0].text == "Normal "
        assert runs[0].bold is False
        assert runs[1].bold is True

    def test_template_variables(self):
        node = dok.parse('''
            def show(t) { h1 { t } }
            show(t: "My Report")
        ''')
        data = dok.to_bytes(node)
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "My Report" in text

    def test_template_each_loop(self):
        node = dok.parse('''
            let items = ["Alpha", "Beta", "Gamma"]
            each item in items {
                p { item }
            }
        ''')
        data = dok.to_bytes(node)
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Alpha" in text
        assert "Gamma" in text

    def test_function_definitions(self):
        node = dok.parse('''
            def section(title) {
                h2 { title }
                line { }
            }
            section(title: "Overview")
            section(title: "Details")
        ''')
        data = dok.to_bytes(node)
        root = docx_document_xml(data)
        text = get_text_content(root)
        assert "Overview" in text
        assert "Details" in text
