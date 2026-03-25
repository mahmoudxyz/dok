"""
dok.nodes
~~~~~~~~~
The node tree. This is the single data model that:
  - the parser produces
  - the builder API produces
  - the converter consumes
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any

from .errors import SourceLoc


# ---------------------------------------------------------------------------
# Base
# ---------------------------------------------------------------------------

class Node:
    """Base class for all nodes in a Dok document tree."""
    loc: SourceLoc | None = None

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}()"


# ---------------------------------------------------------------------------
# Concrete node types
# ---------------------------------------------------------------------------

@dataclass
class ElementNode(Node):
    name:     str
    props:    dict[str, Any]  = field(default_factory=dict)
    children: list[Node]      = field(default_factory=list)
    loc:      SourceLoc | None = None

    def __repr__(self) -> str:
        return f"ElementNode({self.name!r}, props={self.props}, children={len(self.children)})"

    def prop(self, key: str, default: Any = None) -> Any:
        return self.props.get(key, default)

    def flag(self, key: str) -> bool:
        v = self.props.get(key, False)
        return v is True or v == "true"


@dataclass
class TextNode(Node):
    text: str
    loc:  SourceLoc | None = None

    def __repr__(self) -> str:
        return f"TextNode({self.text!r})"


@dataclass
class ArrowNode(Node):
    label: str | None = None
    loc:   SourceLoc | None = None

    def __repr__(self) -> str:
        return f"ArrowNode(label={self.label!r})"


@dataclass
class FunctionDefNode(Node):
    name:   str
    params: list[str]
    body:   list[Node]
    loc:    SourceLoc | None = None

    def __repr__(self) -> str:
        return f"FunctionDefNode({self.name!r}, params={self.params})"


@dataclass
class ImportNode(Node):
    """import "path.dok" directive — resolved before function expansion."""
    path: str
    loc:  SourceLoc | None = None

    def __repr__(self) -> str:
        return f"ImportNode({self.path!r})"


# ---------------------------------------------------------------------------
# Template nodes — resolved before function expansion
# ---------------------------------------------------------------------------

@dataclass
class LetNode(Node):
    """Variable assignment: let name = value"""
    name:  str
    value: Any   # string, int, list
    loc:   SourceLoc | None = None

    def __repr__(self) -> str:
        return f"LetNode({self.name!r}, {self.value!r})"


@dataclass
class EachNode(Node):
    """Loop: each item in items { body }"""
    var_name:  str
    iterable:  str        # variable name to iterate
    index_var: str | None = None  # optional index variable name
    body:      list[Node] = field(default_factory=list)
    loc:       SourceLoc | None = None

    def __repr__(self) -> str:
        return f"EachNode({self.var_name!r} in {self.iterable!r})"


@dataclass
class IfNode(Node):
    """Conditional: if expr { then } else { else }"""
    condition:  list  # list of tokens/values representing the expression
    then_body:  list[Node] = field(default_factory=list)
    elif_clauses: list[tuple[list, list[Node]]] = field(default_factory=list)
    else_body:  list[Node] = field(default_factory=list)
    loc:        SourceLoc | None = None

    def __repr__(self) -> str:
        return f"IfNode(condition={self.condition!r})"


# ---------------------------------------------------------------------------
# DOCX preset geometry names — used by the converter for drawing shapes
# ---------------------------------------------------------------------------

SHAPE_PRESETS: dict[str, str] = {
    "circle":   "ellipse",
    "diamond":  "diamond",
    "chevron":  "chevron",
}
