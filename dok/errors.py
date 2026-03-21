"""
dok.errors
~~~~~~~~~~
Unified error hierarchy for every stage of the pipeline.

Every error carries a source location (line, col) and an optional hint
that tells the user how to fix the problem.

Stages:
  LexError        — tokenisation (bad characters, unterminated strings)
  ParseError      — syntax (unexpected tokens, unclosed blocks)
  ResolveError    — function expansion (undefined function, missing param)
  ValidationError — semantic checks (bad nesting, invalid props, out-of-range)
"""

from __future__ import annotations
from dataclasses import dataclass


@dataclass(frozen=True)
class SourceLoc:
    """Points to a position in the .dok source."""
    line: int   # 1-based
    col:  int   # 1-based

    def __str__(self) -> str:
        return f"line {self.line}, col {self.col}"


class DokError(Exception):
    """Base for all dok errors.  Always has a location and optional hint."""

    def __init__(
        self,
        message: str,
        loc: SourceLoc | None = None,
        hint: str | None = None,
    ) -> None:
        self.msg  = message
        self.loc  = loc
        self.hint = hint
        super().__init__(self._formatted())

    def _formatted(self) -> str:
        parts: list[str] = []
        if self.loc:
            parts.append(f"{self.loc}: {self.msg}")
        else:
            parts.append(self.msg)
        if self.hint:
            parts.append(f"  hint: {self.hint}")
        return "\n".join(parts)


class LexError(DokError):
    """Raised during tokenisation."""


class ParseError(DokError):
    """Raised during parsing."""


class ResolveError(DokError):
    """Raised during function/component expansion."""


class ValidationError(DokError):
    """A single validation problem."""


class ValidationErrors(DokError):
    """Collects multiple ValidationError instances so the user sees them all."""

    def __init__(self, errors: list[ValidationError]) -> None:
        self.errors = errors
        combined = "\n".join(str(e) for e in errors)
        super().__init__(f"{len(errors)} validation error(s):\n{combined}")
