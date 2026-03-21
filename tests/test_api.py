"""Tests for the dok public API."""
import pytest
import zipfile
import io
from pathlib import Path

import dok
from dok.nodes import ElementNode
from dok.errors import LexError, ParseError, ValidationErrors


class TestParse:
    def test_simple_parse(self):
        node = dok.parse('doc { p { "hello" } }')
        assert isinstance(node, ElementNode)
        assert node.name == "doc"

    def test_implicit_doc_wrapper(self):
        node = dok.parse('p { "hello" }')
        assert node.name == "doc"
        assert node.children[0].name == "p"

    def test_function_expansion(self):
        node = dok.parse('''
            def greet(name) { bold { name } }
            greet(name: "World")
        ''')
        assert node.name == "doc"
        assert node.children[0].name == "bold"

    def test_parse_with_imports(self, tmp_path):
        lib = tmp_path / "lib.dok"
        lib.write_text('def hi(name) { p { name } }')

        node = dok.parse(
            'import "lib.dok"\nhi(name: "test")',
            base_dir=tmp_path,
        )
        assert node.children[0].name == "p"


class TestParseErrors:
    def test_lex_error(self):
        with pytest.raises(LexError):
            dok.parse('"unterminated')

    def test_parse_error(self):
        with pytest.raises(ParseError):
            dok.parse("bold {")

    def test_validation_error(self):
        with pytest.raises(ValidationErrors):
            dok.parse('doc { li { "orphan" } }')


class TestToDocx:
    def test_write_file(self, tmp_path):
        node = dok.parse('doc { p { "hello" } }')
        out = tmp_path / "test.docx"
        dok.to_docx(node, out)
        assert out.exists()
        # Verify it's a valid zip
        assert zipfile.is_zipfile(out)

    def test_contains_document_xml(self, tmp_path):
        node = dok.parse('doc { p { "hello" } }')
        out = tmp_path / "test.docx"
        dok.to_docx(node, out)
        with zipfile.ZipFile(out) as z:
            assert "word/document.xml" in z.namelist()


class TestToBytes:
    def test_returns_bytes(self):
        node = dok.parse('doc { p { "hello" } }')
        data = dok.to_bytes(node)
        assert isinstance(data, bytes)
        assert zipfile.is_zipfile(io.BytesIO(data))


class TestBuilderAPI:
    def test_builder_to_docx(self, tmp_path):
        root = dok.doc(
            dok.page(
                dok.h1("Test"),
                dok.p("Hello ", dok.bold("world")),
                margin="normal",
            ),
        )
        out = tmp_path / "builder.docx"
        dok.to_docx(root, out)
        assert out.exists()

    def test_all_builder_functions_exist(self):
        # Verify all builder functions are exported
        for name in [
            "doc", "page", "center", "right", "justify", "rtl", "ltr",
            "indent", "row", "cols", "col", "bold", "italic", "underline",
            "strike", "sup", "sub", "color", "size", "font", "highlight",
            "span", "h1", "h2", "h3", "h4", "p", "quote", "code",
            "box", "circle", "diamond", "chevron", "callout", "badge",
            "banner", "line", "page_break", "arrow",
            "ul", "ol", "li", "table", "tr", "td", "th",
            "img", "link", "page_number", "header", "footer", "space",
        ]:
            assert hasattr(dok, name), f"dok.{name} not exported"

    def test_lists_builder(self, tmp_path):
        root = dok.doc(
            dok.ul(dok.li("a"), dok.li("b")),
            dok.ol(dok.li("1"), dok.li("2")),
        )
        out = tmp_path / "lists.docx"
        dok.to_docx(root, out)
        assert out.exists()

    def test_table_builder(self, tmp_path):
        root = dok.doc(
            dok.table(
                dok.tr(dok.th("H1"), dok.th("H2")),
                dok.tr(dok.td("A"), dok.td("B")),
                border=True,
            ),
        )
        out = tmp_path / "table.docx"
        dok.to_docx(root, out)
        assert out.exists()

    def test_link_builder(self, tmp_path):
        root = dok.doc(
            dok.p(dok.link("https://example.com", "click")),
        )
        out = tmp_path / "link.docx"
        dok.to_docx(root, out)
        assert out.exists()
