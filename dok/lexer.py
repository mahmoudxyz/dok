"""
dok.lexer
~~~~~~~~~
Tokenises Dok string syntax into a flat list of tokens.
Fully implemented — the parser consumes these.

Token types:
  NAME      identifier:  doc  box  h1  fill  rounded
  STRING    quoted text: "hello world" or triple-quoted \"\"\"multiline\"\"\"
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
    "LBRACKET", "RBRACKET",
    "COLON", "COMMA", "ARROW", "PAGEBREAK",
    "EQUALS", "DOLLAR", "OP",
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
        ("OP",        r"[!=<>]=|[<>+*/]"),   # ==, !=, <=, >=, <, >, +, *, /
        ("COLOR",     r"#[0-9a-fA-F]{3,6}"),
        ("NUMBER",    r"\d+"),
        ("STRING",    r'"(?:[^"\\]|\\.)*"'),
        ("LPAREN",    r"\("),
        ("RPAREN",    r"\)"),
        ("LBRACE",    r"\{"),
        ("RBRACE",    r"\}"),
        ("LBRACKET",  r"\["),
        ("RBRACKET",  r"\]"),
        ("EQUALS",    r"="),
        ("DOLLAR",    r"\$"),
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
        """Remove // line comments, preserving line numbers.
        Respects both single-quoted and triple-quoted strings."""
        result: list[str] = []
        i = 0
        while i < len(source):
            # Triple-quoted string — pass through verbatim
            if source[i:i+3] == '"""':
                end = source.find('"""', i + 3)
                if end == -1:
                    result.append(source[i:])
                    break
                result.append(source[i:end+3])
                i = end + 3
            # Single-quoted string — pass through verbatim
            elif source[i] == '"':
                j = i + 1
                while j < len(source) and source[j] != '"':
                    if source[j] == '\\':
                        j += 1
                    j += 1
                result.append(source[i:j+1])
                i = j + 1
            # Line comment
            elif source[i:i+2] == '//':
                end = source.find('\n', i)
                if end == -1:
                    break
                i = end  # keep the newline
            else:
                result.append(source[i])
                i += 1
        return "".join(result)

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

            # Triple-quoted string: """..."""
            if source[pos:pos+3] == '"""':
                end = source.find('"""', pos + 3)
                if end == -1:
                    raise LexError(
                        "Unterminated triple-quoted string",
                        loc=SourceLoc(line, col),
                        hint='Triple-quoted strings must end with """.',
                    )
                raw = source[pos+3:end]
                # Strip common leading indentation
                raw = _dedent(raw)
                yield Token("STRING", raw, line, col)
                # Advance line counter
                full = source[pos:end+3]
                newlines = full.count("\n")
                if newlines:
                    line += newlines
                    line_start = pos + full.rfind("\n") + 1
                pos = end + 3
                continue

            m = self._MASTER.match(source, pos)

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


def _dedent(text: str) -> str:
    """Strip common leading whitespace from a triple-quoted string.
    Also strips the first line if it's blank (right after opening quotes)
    and the last line if it's blank (right before closing quotes)."""
    lines = text.split("\n")
    # Strip leading blank line
    if lines and not lines[0].strip():
        lines = lines[1:]
    # Strip trailing blank line
    if lines and not lines[-1].strip():
        lines = lines[:-1]
    if not lines:
        return ""
    # Find minimum indentation of non-blank lines
    indents = [len(l) - len(l.lstrip()) for l in lines if l.strip()]
    if not indents:
        return "\n".join(lines)
    min_indent = min(indents)
    return "\n".join(l[min_indent:] for l in lines)