"""Tests for the dok lexer."""
import pytest
from dok.lexer import Lexer, Token
from dok.errors import LexError


def lex(source: str) -> list[Token]:
    return Lexer(source).tokenize()


def types(source: str) -> list[str]:
    return [t.type for t in lex(source)]


def types_no_eof(source: str) -> list[str]:
    return [t.type for t in lex(source) if t.type != "EOF"]


class TestBasicTokens:
    def test_empty(self):
        # Lexer may return just EOF
        non_eof = types_no_eof("")
        assert non_eof == []

    def test_braces(self):
        assert types_no_eof("{ }") == ["LBRACE", "RBRACE"]

    def test_parens(self):
        assert types_no_eof("( )") == ["LPAREN", "RPAREN"]

    def test_colon(self):
        assert types_no_eof(":") == ["COLON"]

    def test_comma(self):
        assert types_no_eof(",") == ["COMMA"]

    def test_arrow(self):
        assert types_no_eof("->") == ["ARROW"]

    def test_page_break(self):
        toks = types_no_eof("---")
        assert len(toks) == 1
        assert "BREAK" in toks[0].upper() or "PAGE" in toks[0].upper()


class TestStrings:
    def test_double_quoted(self):
        tokens = [t for t in lex('"hello"') if t.type != "EOF"]
        assert len(tokens) == 1
        assert tokens[0].type == "STRING"
        assert tokens[0].value == "hello"

    def test_escape_sequences(self):
        tokens = lex(r'"line\nbreak"')
        assert tokens[0].value == "line\nbreak"


class TestNames:
    def test_simple_name(self):
        tokens = lex("bold")
        assert tokens[0].type == "NAME"
        assert tokens[0].value == "bold"

    def test_hyphenated_name(self):
        tokens = lex("page-number")
        assert tokens[0].type == "NAME"
        assert tokens[0].value == "page-number"

    def test_hash_color(self):
        tokens = lex("#FF0000")
        # May be COLOR or NAME type
        assert tokens[0].value == "#FF0000" or "FF0000" in tokens[0].value


class TestNumbers:
    def test_integer(self):
        tokens = lex("42")
        assert tokens[0].type == "NUMBER"
        assert tokens[0].value == "42"


class TestComments:
    def test_line_comment(self):
        tokens = [t for t in lex("// this is a comment\nbold") if t.type != "EOF"]
        assert len(tokens) == 1
        assert tokens[0].value == "bold"


class TestKeywords:
    def test_def(self):
        non_eof = [t for t in lex("def") if t.type != "EOF"]
        assert len(non_eof) == 1
        assert non_eof[0].value == "def"

    def test_import(self):
        tokens = lex("import")
        assert tokens[0].value == "import"


class TestSourceLocation:
    def test_token_has_location(self):
        tokens = lex("bold")
        assert tokens[0].line >= 1

    def test_multiline_location(self):
        tokens = [t for t in lex('bold\n"text"') if t.type != "EOF"]
        assert tokens[0].line == 1
        assert tokens[1].line == 2


class TestErrors:
    def test_unterminated_string(self):
        with pytest.raises(LexError):
            lex('"unterminated')

    def test_error_has_location(self):
        with pytest.raises(LexError) as exc_info:
            lex('"bad')
        assert exc_info.value.loc is not None
