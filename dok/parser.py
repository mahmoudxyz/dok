"""
dok.parser
~~~~~~~~~~
Converts a flat token list (from Lexer) into a node tree.
"""

from __future__ import annotations
from typing import Any
from .lexer import Token
from .nodes import Node, ElementNode, TextNode, ArrowNode


class ParseError(Exception):
    def __init__(self, message: str, token: Token) -> None:
        super().__init__(f"Line {token.line}, col {token.col}: {message}")
        self.token = token


class Parser:
    """
    Recursive descent parser. Consumes a token list, produces a Node list.
    """

    def __init__(self, tokens: list[Token]) -> None:
        self._tokens = tokens
        self._pos    = 0

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def parse(self) -> list[Node]:
        nodes = []
        while not self._at_end():
            nodes.append(self._parse_node())
        return nodes

    # ------------------------------------------------------------------
    # Node parsing
    # ------------------------------------------------------------------

    def _parse_node(self) -> Node:
        if self._peek_type("NAME"):
            return self._parse_element()

        if self._peek_type("STRING"):
            return self._parse_text()

        if self._peek_type("ARROW"):
            return self._parse_arrow()

        if self._peek_type("PAGEBREAK"):
            self._consume("PAGEBREAK")
            return ElementNode(name="---", props={}, children=[])

        tok = self._peek()
        raise ParseError(
            f"Unexpected token {tok.type} ({tok.value!r})", tok
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

        return ElementNode(name=name, props=props, children=children)

    def _parse_props(self) -> dict[str, Any]:
        self._consume("LPAREN")
        props: dict[str, Any] = {}

        while not self._peek_type("RPAREN"):
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
                tok = self._peek()
                raise ParseError("Unclosed block — missing '}'", tok)
            children.append(self._parse_node())

        self._consume("RBRACE")
        return children

    def _parse_arrow(self) -> ArrowNode:
        self._consume("ARROW")

        # Labeled arrow: -> "label" ->
        if self._peek_type("STRING"):
            label = self._consume("STRING").value
            self._consume("ARROW")
            return ArrowNode(label=label)

        # Plain arrow: ->
        return ArrowNode(label=None)

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

        raise ParseError(
            f"Expected a value (name, string, number, color), "
            f"got {tok.type} ({tok.value!r})",
            tok,
        )
    def _parse_text(self) -> TextNode:
        tok = self._consume("STRING")
        return TextNode(text=tok.value)

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
            raise ParseError(
                f"Expected {expected_type}, got {tok.type} ({tok.value!r})",
                tok,
            )
        self._pos += 1
        return tok

    def _at_end(self) -> bool:
        return self._peek().type == "EOF"