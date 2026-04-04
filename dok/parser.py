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
from .nodes import (
    Node, ElementNode, TextNode, ArrowNode, FunctionDefNode, ImportNode,
    LetNode, EachNode, IfNode,
)
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
            if tok_val == "let":
                return self._parse_let()
            if tok_val == "each":
                return self._parse_each()
            if tok_val == "if":
                return self._parse_if()
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

    # ------------------------------------------------------------------
    # Template constructs
    # ------------------------------------------------------------------

    def _parse_let(self) -> LetNode:
        """Parse: let name = value  or  let name = [item1, item2, ...]"""
        let_tok = self._consume("NAME")  # consume "let"
        if not self._peek_type("NAME"):
            raise self._error("Expected variable name after 'let'", self._peek(),
                              hint='let myvar = "value"')
        name_tok = self._consume("NAME")
        var_name = name_tok.value

        if not self._peek_type("EQUALS"):
            raise self._error("Expected '=' after variable name", self._peek(),
                              hint=f'let {var_name} = "value"')
        self._consume("EQUALS")

        # List literal: [item1, item2, ...]
        if self._peek_type("LBRACKET"):
            self._consume("LBRACKET")
            items: list = []
            while not self._peek_type("RBRACKET"):
                if self._at_end():
                    raise self._error("Unclosed '[' — missing ']'", self._peek())
                items.append(self._parse_let_value())
                if self._peek_type("COMMA"):
                    self._consume("COMMA")
            self._consume("RBRACKET")
            return LetNode(name=var_name, value=items, loc=let_tok.loc)

        value = self._parse_let_value()
        return LetNode(name=var_name, value=value, loc=let_tok.loc)

    def _parse_let_value(self):
        """Parse a single value in a let expression."""
        tok = self._peek()
        if tok.type == "STRING":
            self._consume("STRING")
            return tok.value
        if tok.type == "NUMBER":
            self._consume("NUMBER")
            raw = tok.value
            has_unit = any(raw.endswith(u) for u in ("pt", "cm", "mm", "in", "px", "emu", "twip"))
            if has_unit or "." in raw:
                return raw
            return int(raw)
        if tok.type == "NAME":
            self._consume("NAME")
            # Could be a boolean
            if tok.value in ("true", "True"):
                return True
            if tok.value in ("false", "False"):
                return False
            return tok.value  # variable reference
        if tok.type == "COLOR":
            self._consume("COLOR")
            return tok.value
        raise self._error(f"Expected a value, got {tok.type}", tok)

    def _parse_each(self) -> EachNode:
        """Parse: each item in items { body }  or  each item, idx in items { body }"""
        each_tok = self._consume("NAME")  # consume "each"
        if not self._peek_type("NAME"):
            raise self._error("Expected variable name after 'each'", self._peek(),
                              hint='each item in mylist { ... }')
        var_tok = self._consume("NAME")
        var_name = var_tok.value

        # Optional index variable: each item, idx in items
        index_var = None
        if self._peek_type("COMMA"):
            self._consume("COMMA")
            if not self._peek_type("NAME"):
                raise self._error("Expected index variable name after ','", self._peek())
            index_var = self._consume("NAME").value

        # Expect "in"
        if not self._peek_type("NAME") or self._peek().value != "in":
            raise self._error("Expected 'in' after variable name", self._peek(),
                              hint=f'each {var_name} in mylist {{ ... }}')
        self._consume("NAME")  # consume "in"

        # Iterable (variable name)
        if not self._peek_type("NAME"):
            raise self._error("Expected iterable name after 'in'", self._peek())
        iterable = self._consume("NAME").value

        # Body
        if not self._peek_type("LBRACE"):
            raise self._error("Expected '{' for each body", self._peek())
        body = self._parse_block()

        return EachNode(var_name=var_name, iterable=iterable, index_var=index_var,
                        body=body, loc=each_tok.loc)

    def _parse_if(self) -> IfNode:
        """Parse: if expr { then } elif expr { ... } else { else }"""
        if_tok = self._consume("NAME")  # consume "if"
        condition = self._parse_condition()

        if not self._peek_type("LBRACE"):
            raise self._error("Expected '{' after if condition", self._peek())
        then_body = self._parse_block()

        elif_clauses: list[tuple[list, list[Node]]] = []
        else_body: list[Node] = []

        while self._peek_type("NAME") and self._peek().value == "elif":
            self._consume("NAME")  # consume "elif"
            elif_cond = self._parse_condition()
            if not self._peek_type("LBRACE"):
                raise self._error("Expected '{' after elif condition", self._peek())
            elif_body = self._parse_block()
            elif_clauses.append((elif_cond, elif_body))

        if self._peek_type("NAME") and self._peek().value == "else":
            self._consume("NAME")  # consume "else"
            if not self._peek_type("LBRACE"):
                raise self._error("Expected '{' after else", self._peek())
            else_body = self._parse_block()

        return IfNode(condition=condition, then_body=then_body,
                      elif_clauses=elif_clauses, else_body=else_body,
                      loc=if_tok.loc)

    def _parse_condition(self) -> list:
        """Parse a condition expression (tokens until '{')."""
        tokens: list = []
        while not self._peek_type("LBRACE") and not self._at_end():
            tok = self._peek()
            if tok.type == "NAME":
                tokens.append(("name", self._consume("NAME").value))
            elif tok.type == "NUMBER":
                tokens.append(("num", int(self._consume("NUMBER").value)))
            elif tok.type == "STRING":
                tokens.append(("str", self._consume("STRING").value))
            elif tok.type == "OP":
                tokens.append(("op", self._consume("OP").value))
            elif tok.type == "DOLLAR":
                self._consume("DOLLAR")
                if self._peek_type("NAME"):
                    tokens.append(("var", self._consume("NAME").value))
                else:
                    raise self._error("Expected variable name after '$'", self._peek())
            else:
                break
        return tokens

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
            raw = tok.value
            # Check if it has a unit suffix or decimal point — keep as string
            has_unit = any(raw.endswith(u) for u in ("pt", "cm", "mm", "in", "px", "emu", "twip"))
            has_decimal = "." in raw
            if has_unit or has_decimal:
                # Return as string so units.py can parse it
                # But first check for ratio syntax
                if self._peek_type("COLON") and self._peek(1).type == "NUMBER":
                    # Ratio not supported with unit-suffixed numbers
                    pass
                return raw
            n1 = int(raw)
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

    def _parse_text(self) -> TextNode | ElementNode:
        tok = self._consume("STRING")
        # Check for inline sugar (**bold**, *italic*, etc.)
        from .sugar import desugar_inline
        expanded = desugar_inline(tok.value)
        if expanded:
            # Re-parse the expanded inline sugar as a sequence of nodes
            # Wrap in a temporary container to parse multiple nodes
            from .lexer import Lexer
            try:
                sub_tokens = Lexer(expanded).tokenize()
                sub_parser = Parser(sub_tokens)
                sub_nodes = sub_parser.parse()
                if len(sub_nodes) == 1:
                    return sub_nodes[0]
                # Multiple nodes: wrap in a span element
                return ElementNode(
                    name="span", props={}, children=sub_nodes, loc=tok.loc,
                )
            except Exception:
                pass  # Fall through to plain text
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
