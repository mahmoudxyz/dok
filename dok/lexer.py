"""
dok.lexer
~~~~~~~~~
Tokenises Dok string syntax into a flat list of tokens.
Fully implemented — the parser consumes these.

Token types:
  NAME      identifier:  doc  box  h1  fill  rounded
  STRING    quoted text: "hello world"
  NUMBER    integer:     11  14  2
  COLOR     hex color:   #4472C4
  LPAREN    (
  RPAREN    )
  LBRACE    {
  RBRACE    }
  COLON     :
  COMMA     ,
  ARROW     ->
  PAGEBREAK ---
  EOF
"""

from __future__ import annotations
import re
from dataclasses import dataclass
from typing import Iterator

from .errors import LexError, SourceLoc


# ---------------------------------------------------------------------------
# Token
# ---------------------------------------------------------------------------

TOKEN_TYPES = (
    "NAME", "STRING", "NUMBER", "COLOR",
    "LPAREN", "RPAREN", "LBRACE", "RBRACE",
    "COLON", "COMMA", "ARROW", "PAGEBREAK",
    "EOF",
)


@dataclass
class Token:
    type:  str
    value: str
    line:  int    # 1-based, for error messages
    col:   int    # 1-based

    @property
    def loc(self) -> SourceLoc:
        return SourceLoc(self.line, self.col)

    def __repr__(self) -> str:
        return f"Token({self.type}, {self.value!r}, line={self.line})"


# ---------------------------------------------------------------------------
# Lexer
# ---------------------------------------------------------------------------

class Lexer:
    """
    Converts a Dok source string into a list of tokens.

    Usage:
        tokens = Lexer(source).tokenize()

    All tokens including EOF are returned.
    Comments (// to end of line) are stripped before tokenising.
    """

    # Ordered: longest match first where ambiguous
    _PATTERNS = [
        ("PAGEBREAK", r"---"),
        ("ARROW",     r"->"),
        ("COLOR",     r"#[0-9a-fA-F]{3,6}"),
        ("NUMBER",    r"\d+"),
        ("STRING",    r'"(?:[^"\\]|\\.)*"'),
        ("LPAREN",    r"\("),
        ("RPAREN",    r"\)"),
        ("LBRACE",    r"\{"),
        ("RBRACE",    r"\}"),
        ("COLON",     r":"),
        ("COMMA",     r","),
        ("NAME",      r"[a-zA-Z\u0600-\u06FF\u0750-\u077F\u0590-\u05FF\u00C0-\u024F\u1E00-\u1EFF][a-zA-Z0-9_\u0600-\u06FF\u0750-\u077F\u0590-\u05FF\u00C0-\u024F\u1E00-\u1EFF-]*"),
    ]

    _MASTER = re.compile(
        "|".join(f"(?P<{name}>{pat})" for name, pat in _PATTERNS),
        re.UNICODE,
    )

    def __init__(self, source: str) -> None:
        self._source = self._strip_comments(source)

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def tokenize(self) -> list[Token]:
        """Return all tokens including a final EOF token."""
        return list(self._scan())

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _strip_comments(self, source: str) -> str:
        """Remove // line comments, preserving line numbers."""
        lines = []
        for line in source.splitlines():
            # Only strip comment if // is not inside a string
            in_string = False
            i = 0
            while i < len(line):
                ch = line[i]
                if ch == '"' and not in_string:
                    in_string = True
                elif ch == '"' and in_string:
                    in_string = False
                elif ch == '/' and not in_string and i + 1 < len(line) and line[i+1] == '/':
                    line = line[:i]
                    break
                i += 1
            lines.append(line)
        return "\n".join(lines)

    def _scan(self) -> Iterator[Token]:
        source = self._source
        pos    = 0
        line   = 1
        line_start = 0

        while pos < len(source):
            # Skip whitespace
            if source[pos] in " \t\r\n":
                if source[pos] == "\n":
                    line += 1
                    line_start = pos + 1
                pos += 1
                continue

            col = pos - line_start + 1
            m   = self._MASTER.match(source, pos)

            if not m:
                raise LexError(
                    f"Unexpected character {source[pos]!r}",
                    loc=SourceLoc(line, col),
                    hint="Check for stray characters or unsupported symbols.",
                )

            tok_type  = m.lastgroup
            tok_value = m.group()

            if tok_type == "STRING":
                # Unescape the value, strip surrounding quotes
                tok_value = tok_value[1:-1].replace('\\"', '"').replace("\\n", "\n")

            yield Token(tok_type, tok_value, line, col)

            # Advance line counter for multi-line strings
            newlines = m.group().count("\n")
            if newlines:
                line += newlines
                line_start = pos + m.group().rfind("\n") + 1

            pos = m.end()

        yield Token("EOF", "", line, pos - line_start + 1)