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
# Layer classification
# ---------------------------------------------------------------------------

# Layer 1 — document defaults
DOC_NODES = {"doc"}

# Layer 2 — physical space
PAGE_NODES = {"page"}

# Layer 3 — layout / arrangement
LAYOUT_NODES = {"center", "right", "left", "justify", "rtl", "ltr",
                "indent", "row", "cols", "col", "float"}

# Layer 4 — run style
STYLE_NODES  = {"bold", "italic", "underline", "strike", "sup", "sub",
                "color", "size", "font", "highlight"}

# Layer 5 — content atoms
BLOCK_NODES  = {"h1", "h2", "h3", "h4", "p", "quote", "code"}
SHAPE_NODES  = {"box", "circle", "diamond", "chevron", "callout",
                "badge", "banner", "line"}
LIST_NODES   = {"ul", "ol", "li"}
TABLE_NODES  = {"table", "tr", "td", "th"}
INLINE_NODES = {"link", "img", "page-number"}
META_NODES   = {"header", "footer", "space"}
SPECIAL_NODES = {"---"}

CONTENT_NODES = BLOCK_NODES | SHAPE_NODES | SPECIAL_NODES | LIST_NODES | TABLE_NODES | INLINE_NODES | META_NODES

ALL_KNOWN_NODES = DOC_NODES | PAGE_NODES | LAYOUT_NODES | STYLE_NODES | CONTENT_NODES


def node_layer(name: str) -> str:
    if name in DOC_NODES:    return "doc"
    if name in PAGE_NODES:   return "page"
    if name in LAYOUT_NODES: return "layout"
    if name in STYLE_NODES:  return "style"
    if name in CONTENT_NODES:return "content"
    return "unknown"


# ---------------------------------------------------------------------------
# DOCX preset geometry names — used by the converter for drawing shapes
# ---------------------------------------------------------------------------

SHAPE_PRESETS: dict[str, str] = {
    "circle":   "ellipse",
    "diamond":  "diamond",
    "chevron":  "chevron",
}
