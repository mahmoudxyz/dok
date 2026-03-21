"""Tests for the dok resolver (function expansion and imports)."""
import pytest
from dok.lexer import Lexer
from dok.parser import Parser
from dok.resolver import resolve, resolve_imports
from dok.nodes import ElementNode, TextNode
from dok.errors import ResolveError


def parse_raw(source: str):
    return Parser(Lexer(source).tokenize()).parse()


def resolve_source(source: str):
    return resolve(parse_raw(source))


class TestFunctionExpansion:
    def test_simple_function(self):
        nodes = resolve_source('''
            def greeting(name) { bold { name } }
            greeting(name: "Alice")
        ''')
        assert len(nodes) == 1
        assert isinstance(nodes[0], ElementNode)
        assert nodes[0].name == "bold"
        assert isinstance(nodes[0].children[0], TextNode)
        assert nodes[0].children[0].text == "Alice"

    def test_multi_param_function(self):
        nodes = resolve_source('''
            def card(title, subtitle) {
                h1 { title }
                p { subtitle }
            }
            card(title: "Hello", subtitle: "World")
        ''')
        assert len(nodes) == 2
        assert nodes[0].name == "h1"
        assert nodes[1].name == "p"

    def test_children_keyword(self):
        nodes = resolve_source('''
            def wrapper(color_val) {
                box(fill: color_val) { children }
            }
            wrapper(color_val: blue) { "inner content" }
        ''')
        assert nodes[0].name == "box"
        assert isinstance(nodes[0].children[0], TextNode)
        assert nodes[0].children[0].text == "inner content"

    def test_function_strips_def(self):
        nodes = resolve_source('''
            def unused(x) { p { x } }
            h1 { "title" }
        ''')
        assert len(nodes) == 1
        assert nodes[0].name == "h1"


class TestFunctionErrors:
    def test_duplicate_function(self):
        with pytest.raises(ResolveError, match="Duplicate"):
            resolve_source('''
                def foo(x) { p { x } }
                def foo(y) { p { y } }
            ''')

    def test_builtin_name_conflict(self):
        with pytest.raises(ResolveError, match="built-in"):
            resolve_source('def bold(x) { p { x } }')

    def test_missing_parameter(self):
        with pytest.raises(ResolveError, match="Missing parameter"):
            resolve_source('''
                def greet(name) { p { name } }
                greet()
            ''')

    def test_unknown_parameter(self):
        with pytest.raises(ResolveError, match="Unknown parameter"):
            resolve_source('''
                def greet(name) { p { name } }
                greet(name: "A", extra: "B")
            ''')


class TestImports:
    def test_import_resolution(self, tmp_path):
        lib = tmp_path / "lib.dok"
        lib.write_text('def greet(name) { bold { name } }')

        main_source = 'import "lib.dok"\ngreet(name: "World")'
        nodes = parse_raw(main_source)
        nodes = resolve_imports(nodes, tmp_path)
        nodes = resolve(nodes)
        assert len(nodes) == 1
        assert nodes[0].name == "bold"

    def test_circular_import(self, tmp_path):
        a = tmp_path / "a.dok"
        b = tmp_path / "b.dok"
        a.write_text('import "b.dok"')
        b.write_text('import "a.dok"')

        nodes = parse_raw('import "a.dok"')
        with pytest.raises(ResolveError, match="Circular"):
            resolve_imports(nodes, tmp_path)

    def test_missing_import(self, tmp_path):
        nodes = parse_raw('import "nonexistent.dok"')
        with pytest.raises(ResolveError, match="not found"):
            resolve_imports(nodes, tmp_path)
