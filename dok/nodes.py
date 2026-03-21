"""
dok.nodes
~~~~~~~~~
The node tree. This is the single data model that:
  - the parser produces
  - the builder API produces
  - the converter consumes

Nothing else. No logic, no DOCX knowledge, no parsing.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any


# ---------------------------------------------------------------------------
# Base
# ---------------------------------------------------------------------------

class Node:
    """Base class for all nodes in a Dok document tree."""

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}()"


# ---------------------------------------------------------------------------
# Concrete node types
# ---------------------------------------------------------------------------

@dataclass
class ElementNode(Node):
    """
    A named element with optional props and children.

    Covers every Dok construct: doc, page, center, bold, box, h1, p, ...

    Examples in string syntax:
        bold { "hello" }
        box(fill: navy, rounded) { p { "text" } }
        h1 { "Title" }
    """
    name:     str
    props:    dict[str, Any]  = field(default_factory=dict)
    children: list[Node]      = field(default_factory=list)

    def __repr__(self) -> str:
        return f"ElementNode({self.name!r}, props={self.props}, children={len(self.children)})"

    # Convenience: get a prop with a default
    def prop(self, key: str, default: Any = None) -> Any:
        return self.props.get(key, default)

    # Convenience: check a boolean flag
    def flag(self, key: str) -> bool:
        v = self.props.get(key, False)
        return v is True or v == "true"


@dataclass
class TextNode(Node):
    """
    A bare text node. Always a leaf. Always inside an ElementNode.

    Example in string syntax:
        "hello world"
    """
    text: str

    def __repr__(self) -> str:
        return f"TextNode({self.text!r})"


@dataclass
class ArrowNode(Node):
    """
    A connector arrow between two siblings inside a row { }.

    Example in string syntax:
        box { "A" }  ->  box { "B" }
        box { "A" }  -> "label" ->  box { "B" }
    """
    label: str | None = None

    def __repr__(self) -> str:
        return f"ArrowNode(label={self.label!r})"


# ---------------------------------------------------------------------------
# Layer classification
# These are the five layers from the architecture.
# The converter uses these to decide what context to update vs. what to emit.
# ---------------------------------------------------------------------------

# Layer 1 — document defaults
DOC_NODES = {"doc"}

# Layer 2 — physical space
PAGE_NODES = {"page"}

# Layer 3 — layout / arrangement (update ParaCtx, do not emit directly)
LAYOUT_NODES = {"center", "right", "left", "justify", "rtl", "ltr",
                "indent", "row", "cols", "col", "float"}

# Layer 4 — run style (update RunCtx, do not emit directly)
STYLE_NODES  = {"bold", "italic", "underline", "strike", "sup", "sub",
                "color", "size", "font", "highlight"}

# Layer 5 — content atoms (consume context and emit DOCX)
BLOCK_NODES  = {"h1", "h2", "h3", "h4", "p", "quote", "code"}
SHAPE_NODES  = {"box", "circle", "diamond", "chevron", "callout",
                "badge", "banner", "line"}
SPECIAL_NODES = {"---"}   # page break

CONTENT_NODES = BLOCK_NODES | SHAPE_NODES | SPECIAL_NODES


def node_layer(name: str) -> str:
    """Return which layer a node name belongs to."""
    if name in DOC_NODES:    return "doc"
    if name in PAGE_NODES:   return "page"
    if name in LAYOUT_NODES: return "layout"
    if name in STYLE_NODES:  return "style"
    if name in CONTENT_NODES:return "content"
    return "unknown"


# ---------------------------------------------------------------------------
# DOCX preset geometry names — used by the converter
# ---------------------------------------------------------------------------

SHAPE_PRESETS: dict[str, str] = {
    "box":      "rect",
    "circle":   "ellipse",
    "diamond":  "diamond",
    "chevron":  "chevron",
    "callout":  "wedgeRectCallout",
    "badge":    "rect",    # small inline rect
    "banner":   "rect",    # full-width rect
    "line":     "line",
}