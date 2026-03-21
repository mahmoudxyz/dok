"""Tests for the dok parser."""
import pytest
from dok.lexer import Lexer
from dok.parser import Parser
from dok.nodes import ElementNode, TextNode, FunctionDefNode, ImportNode
from dok.errors import ParseError


def parse(source: str):
    tokens = Lexer(source).tokenize()
    return Parser(tokens).parse()


class TestBasicParsing:
    def test_bare_string(self):
        nodes = parse('"hello"')
        assert len(nodes) == 1
        assert isinstance(nodes[0], TextNode)
        assert nodes[0].text == "hello"

    def test_empty_element(self):
        nodes = parse("bold { }")
        assert len(nodes) == 1
        assert isinstance(nodes[0], ElementNode)
        assert nodes[0].name == "bold"
        assert nodes[0].children == []

    def test_element_with_text(self):
        nodes = parse('bold { "hello" }')
        assert nodes[0].name == "bold"
        assert len(nodes[0].children) == 1
        assert isinstance(nodes[0].children[0], TextNode)

    def test_nested_elements(self):
        nodes = parse('bold { italic { "text" } }')
        assert nodes[0].name == "bold"
        inner = nodes[0].children[0]
        assert isinstance(inner, ElementNode)
        assert inner.name == "italic"


class TestProps:
    def test_string_prop(self):
        nodes = parse('img(src: "photo.png")')
        assert nodes[0].props["src"] == "photo.png"

    def test_int_prop(self):
        nodes = parse("size(value: 14)")
        assert nodes[0].props["value"] == 14

    def test_name_prop(self):
        nodes = parse("color(value: red)")
        assert nodes[0].props["value"] == "red"

    def test_bool_prop(self):
        nodes = parse("table(border: true)")
        assert nodes[0].props["border"] in (True, "true")

    def test_multiple_props(self):
        nodes = parse("page(margin: normal, paper: a4)")
        assert nodes[0].props["margin"] == "normal"
        assert nodes[0].props["paper"] == "a4"

    def test_ratio_prop(self):
        nodes = parse("cols(ratio: 1:1:1)")
        assert nodes[0].props["ratio"] == "1:1:1"


class TestFunctionDef:
    def test_simple_function(self):
        nodes = parse('def greeting(name) { bold { name } }')
        assert len(nodes) == 1
        assert isinstance(nodes[0], FunctionDefNode)
        assert nodes[0].name == "greeting"
        assert nodes[0].params == ["name"]

    def test_multi_param_function(self):
        nodes = parse('def card(title, body) { h1 { title } p { body } }')
        fn = nodes[0]
        assert fn.params == ["title", "body"]
        assert len(fn.body) == 2


class TestImport:
    def test_import_statement(self):
        nodes = parse('import "lib.dok"')
        assert len(nodes) == 1
        assert isinstance(nodes[0], ImportNode)
        assert nodes[0].path == "lib.dok"


class TestPageBreak:
    def test_page_break(self):
        nodes = parse("---")
        assert len(nodes) == 1
        assert isinstance(nodes[0], ElementNode)
        assert nodes[0].name == "---"


class TestErrors:
    def test_missing_closing_brace(self):
        with pytest.raises(ParseError):
            parse("bold {")

    def test_error_has_location(self):
        with pytest.raises(ParseError) as exc_info:
            parse("bold {")
        assert exc_info.value.loc is not None
