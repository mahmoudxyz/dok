"""
dok.parser
~~~~~~~~~~
Converts a flat token list (from Lexer) into a node tree.

Supports function definitions:
    def header(title, subtitle) {
        center { bold { title } }
    }
"""

from __future__ import annotations
from typing import Any
from .lexer import Token
from .nodes import Node, ElementNode, TextNode, ArrowNode, FunctionDefNode, ImportNode
from .errors import ParseError, SourceLoc


class Parser:
    """
    Recursive descent parser. Consumes a token list, produces a Node list.
    Function definitions are stored separately from content nodes.
    """

    def __init__(self, tokens: list[Token]) -> None:
        self._tokens = tokens
        self._pos    = 0

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def parse(self) -> list[Node]:
        nodes: list[Node] = []
        while not self._at_end():
            nodes.append(self._parse_node())
        return nodes

    # ------------------------------------------------------------------
    # Node parsing
    # ------------------------------------------------------------------

    def _parse_node(self) -> Node:
        if self._peek_type("NAME"):
            tok_val = self._peek().value
            if tok_val == "def":
                return self._parse_function_def()
            if tok_val == "import":
                return self._parse_import()
            return self._parse_element()

        if self._peek_type("STRING"):
            return self._parse_text()

        if self._peek_type("ARROW"):
            return self._parse_arrow()

        if self._peek_type("PAGEBREAK"):
            tok = self._consume("PAGEBREAK")
            return ElementNode(name="---", props={}, children=[], loc=tok.loc)

        tok = self._peek()
        raise self._error(
            f"Unexpected token {tok.type} ({tok.value!r})",
            tok,
            hint="Expected an element name, quoted string, or ->.",
        )

    def _parse_import(self) -> ImportNode:
        imp_tok = self._consume("NAME")  # consume "import"
        if not self._peek_type("STRING"):
            raise self._error(
                "Expected file path after 'import'",
                self._peek(),
                hint='import "components.dok"',
            )
        path_tok = self._consume("STRING")
        return ImportNode(path=path_tok.value, loc=imp_tok.loc)

    def _parse_function_def(self) -> FunctionDefNode:
        def_tok = self._consume("NAME")  # consume "def"

        # Function name
        if not self._peek_type("NAME"):
            raise self._error(
                "Expected function name after 'def'",
                self._peek(),
                hint="def my_function(param1, param2) { ... }",
            )
        name_tok = self._consume("NAME")
        name = name_tok.value

        # Parameters
        params: list[str] = []
        if self._peek_type("LPAREN"):
            self._consume("LPAREN")
            while not self._peek_type("RPAREN"):
                if not self._peek_type("NAME"):
                    raise self._error(
                        f"Expected parameter name, got {self._peek().type}",
                        self._peek(),
                    )
                params.append(self._consume("NAME").value)
                if self._peek_type("COMMA"):
                    self._consume("COMMA")
            self._consume("RPAREN")

        # Body
        if not self._peek_type("LBRACE"):
            raise self._error(
                "Expected '{' for function body",
                self._peek(),
                hint=f"def {name}({', '.join(params)}) {{ ... }}",
            )
        body = self._parse_block()

        return FunctionDefNode(
            name=name, params=params, body=body, loc=def_tok.loc,
        )

    def _parse_element(self) -> ElementNode:
        name_tok = self._consume("NAME")
        name     = name_tok.value

        # Optional props
        props: dict[str, Any] = {}
        if self._peek_type("LPAREN"):
            props = self._parse_props()

        # Body: block, shorthand string, or nothing
        children: list[Node] = []
        if self._peek_type("LBRACE"):
            children = self._parse_block()
        elif self._peek_type("STRING"):
            children = [self._parse_text()]

        return ElementNode(name=name, props=props, children=children, loc=name_tok.loc)

    def _parse_props(self) -> dict[str, Any]:
        self._consume("LPAREN")
        props: dict[str, Any] = {}

        while not self._peek_type("RPAREN"):
            if self._at_end():
                raise self._error(
                    "Unclosed '(' — missing ')'",
                    self._peek(),
                )
            # Each prop: NAME (COLON value)?
            key_tok = self._consume("NAME")
            key     = key_tok.value

            if self._peek_type("COLON"):
                self._consume("COLON")
                props[key] = self._parse_value()
            else:
                # Bare flag
                props[key] = True

            # Optional comma between props
            if self._peek_type("COMMA"):
                self._consume("COMMA")

        self._consume("RPAREN")
        return props

    def _parse_block(self) -> list[Node]:
        self._consume("LBRACE")
        children: list[Node] = []

        while not self._peek_type("RBRACE"):
            if self._at_end():
                raise self._error(
                    "Unclosed block — missing '}'",
                    self._peek(),
                    hint="Check that every '{' has a matching '}'.",
                )
            children.append(self._parse_node())

        self._consume("RBRACE")
        return children

    def _parse_arrow(self) -> ArrowNode:
        tok = self._consume("ARROW")

        # Labeled arrow: -> "label" ->
        if self._peek_type("STRING"):
            label = self._consume("STRING").value
            self._consume("ARROW")
            return ArrowNode(label=label, loc=tok.loc)

        # Plain arrow: ->
        return ArrowNode(label=None, loc=tok.loc)

    def _parse_value(self) -> Any:
        tok = self._peek()

        if tok.type == "NAME":
            self._consume("NAME")
            return tok.value

        if tok.type == "STRING":
            self._consume("STRING")
            return tok.value

        if tok.type == "NUMBER":
            self._consume("NUMBER")
            n1 = int(tok.value)
            # Handle ratio values: 2:1  or  1:1:1  (any number of parts)
            if self._peek_type("COLON") and self._peek(1).type == "NUMBER":
                parts = [n1]
                while self._peek_type("COLON") and self._peek(1).type == "NUMBER":
                    self._consume("COLON")
                    parts.append(int(self._consume("NUMBER").value))
                return ":".join(str(p) for p in parts)
            return n1

        if tok.type == "COLOR":
            self._consume("COLOR")
            return tok.value

        raise self._error(
            f"Expected a value (name, string, number, color), "
            f"got {tok.type} ({tok.value!r})",
            tok,
        )

    def _parse_text(self) -> TextNode:
        tok = self._consume("STRING")
        return TextNode(text=tok.value, loc=tok.loc)

    # ------------------------------------------------------------------
    # Token helpers
    # ------------------------------------------------------------------

    def _peek(self, offset: int = 0) -> Token:
        pos = self._pos + offset
        if pos >= len(self._tokens):
            return self._tokens[-1]   # EOF
        return self._tokens[pos]

    def _peek_type(self, *types: str, offset: int = 0) -> bool:
        return self._peek(offset).type in types

    def _consume(self, expected_type: str | None = None) -> Token:
        tok = self._tokens[self._pos]
        if expected_type and tok.type != expected_type:
            raise self._error(
                f"Expected {expected_type}, got {tok.type} ({tok.value!r})",
                tok,
            )
        self._pos += 1
        return tok

    def _at_end(self) -> bool:
        return self._peek().type == "EOF"

    def _error(self, message: str, tok: Token, hint: str | None = None) -> ParseError:
        return ParseError(message, loc=tok.loc, hint=hint)
