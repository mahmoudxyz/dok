"""
dok.builder
~~~~~~~~~~~
The Python builder API.

Every function returns a Node. Nodes are plain data — no side effects.
None / False children are silently dropped.

This is how you use Dok from Python code:

    import dok

    doc = dok.doc(
        dok.page(
            dok.banner("Acme Corp", fill="navy", accent="gold", color="white"),
            dok.h1("Q4 Report"),
            dok.p("Revenue grew by ", dok.bold("42%", color="green"), " this year."),
            margin="normal",
        )
    )

    dok.to_docx(doc, "report.docx")
"""

from __future__ import annotations
from typing import Any
from .nodes import ElementNode, TextNode, ArrowNode, Node


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _node(name: str,
          *children: Node | str | None | bool,
          **props: Any) -> ElementNode:
    """
    Build an ElementNode, wrapping strings as TextNodes
    and silently dropping None/False children.
    """
    clean: list[Node] = []
    for child in children:
        if child is None or child is False:
            continue
        if isinstance(child, str):
            clean.append(TextNode(child))
        else:
            clean.append(child)
    return ElementNode(name=name, props=props, children=clean)


def _inline(name: str,
            content: str | Node | None = None,
            *extra: Node | str | None,
            **props: Any) -> ElementNode:
    """
    Shorthand for nodes that commonly hold a single string.

    dok.h1("Title")           → ElementNode("h1", {}, [TextNode("Title")])
    dok.bold("text", color="red")  → ElementNode("bold", {color: "red"}, [TextNode("text")])
    """
    children: list[Node | str | None] = []
    if content is not None:
        children.append(content)
    children.extend(extra)
    return _node(name, *children, **props)


# ---------------------------------------------------------------------------
# Layer 1 — document root
# ---------------------------------------------------------------------------

def doc(*children: Node | str | None, **props: Any) -> ElementNode:
    """
    Document root. Sets font and size defaults.

    dok.doc(
        dok.page(...),
        font="Calibri",
        size=11,
    )
    """
    return _node("doc", *children, **props)


# ---------------------------------------------------------------------------
# Layer 2 — page
# ---------------------------------------------------------------------------

def page(*children: Node | str | None, **props: Any) -> ElementNode:
    """
    Page section. Sets margin, paper, and column count.

    dok.page(
        ...,
        margin="normal",   # normal | narrow | wide | none
        paper="a4",        # a4 | letter | a3
        cols=1,            # column count
    )
    """
    return _node("page", *children, **props)


# ---------------------------------------------------------------------------
# Layer 3 — layout
# ---------------------------------------------------------------------------

def center(*children: Node | str | None) -> ElementNode:
    """Center-align all children."""
    return _node("center", *children)

def right(*children: Node | str | None) -> ElementNode:
    """Right-align all children."""
    return _node("right", *children)

def justify(*children: Node | str | None) -> ElementNode:
    """Justify all children."""
    return _node("justify", *children)

def rtl(*children: Node | str | None) -> ElementNode:
    """Right-to-left block. All children get RTL paragraph direction."""
    return _node("rtl", *children)

def ltr(*children: Node | str | None) -> ElementNode:
    """Left-to-right override (inside rtl blocks)."""
    return _node("ltr", *children)

def indent(*children: Node | str | None, level: int = 1) -> ElementNode:
    """Indent children by *level* levels (each = 0.5 inch)."""
    return _node("indent", *children, level=level)

def row(*children: Node | str | None) -> ElementNode:
    """
    Horizontal flow. Children are placed side by side.
    Use arrow() between children to add connectors.

    dok.row(
        dok.box("Input",   fill="blue"),
        dok.arrow(),
        dok.box("Output",  fill="green"),
    )
    """
    return _node("row", *children)

def cols(*children: Node | str | None, ratio: str = "1:1") -> ElementNode:
    """
    Side-by-side columns. Children should be col() nodes.

    dok.cols(
        dok.col(dok.p("Left.")),
        dok.col(dok.p("Right.")),
        ratio="2:1",
    )
    """
    return _node("cols", *children, ratio=ratio)

def col(*children: Node | str | None) -> ElementNode:
    """One column inside cols()."""
    return _node("col", *children)

def float_right(*children: Node | str | None) -> ElementNode:
    """Float children to the right (text wraps around them)."""
    return _node("float", *children, side="right")

def float_left(*children: Node | str | None) -> ElementNode:
    """Float children to the left."""
    return _node("float", *children, side="left")


# ---------------------------------------------------------------------------
# Layer 4 — run style
# ---------------------------------------------------------------------------

def bold(*children: Node | str | None, **props: Any) -> ElementNode:
    """Bold text. Can carry other run props: dok.bold("text", color="red")"""
    return _node("bold", *children, **props)

def italic(*children: Node | str | None, **props: Any) -> ElementNode:
    """Italic text."""
    return _node("italic", *children, **props)

def underline(*children: Node | str | None, **props: Any) -> ElementNode:
    """Underlined text."""
    return _node("underline", *children, **props)

def strike(*children: Node | str | None, **props: Any) -> ElementNode:
    """Strikethrough text."""
    return _node("strike", *children, **props)

def sup(*children: Node | str | None) -> ElementNode:
    """Superscript."""
    return _node("sup", *children)

def sub(*children: Node | str | None) -> ElementNode:
    """Subscript."""
    return _node("sub", *children)

def color(hex_or_name: str, *children: Node | str | None) -> ElementNode:
    """Colored text. dok.color("red", "text") or dok.color("#FF0000", "text")"""
    return _node("color", *children, value=hex_or_name)

def size(pt: int, *children: Node | str | None) -> ElementNode:
    """Font size in points."""
    return _node("size", *children, value=pt)

def font(name: str, *children: Node | str | None) -> ElementNode:
    """Font family."""
    return _node("font", *children, value=name)

def highlight(color_name: str, *children: Node | str | None) -> ElementNode:
    """Highlighted text (DOCX highlight, limited color set)."""
    return _node("highlight", *children, value=color_name)

def span(*children: Node | str | None, **props: Any) -> ElementNode:
    """
    Generic inline style wrapper for when you need multiple props at once.

    dok.span("text", bold=True, color="red", size=14)
    """
    return _node("span", *children, **props)


# ---------------------------------------------------------------------------
# Layer 5 — content: text blocks
# ---------------------------------------------------------------------------

def h1(content: str | Node, *rest: Node | str | None, **props: Any) -> ElementNode:
    return _node("h1", content, *rest, **props)

def h2(content: str | Node, *rest: Node | str | None, **props: Any) -> ElementNode:
    return _node("h2", content, *rest, **props)

def h3(content: str | Node, *rest: Node | str | None, **props: Any) -> ElementNode:
    return _node("h3", content, *rest, **props)

def h4(content: str | Node, *rest: Node | str | None, **props: Any) -> ElementNode:
    return _node("h4", content, *rest, **props)

def p(*children: Node | str | None, **props: Any) -> ElementNode:
    """
    Paragraph. Can hold text and inline style nodes.

    dok.p("Revenue grew by ", dok.bold("42%", color="green"), " this year.")
    """
    return _node("p", *children, **props)

def quote(*children: Node | str | None) -> ElementNode:
    """Block quote paragraph."""
    return _node("quote", *children)

def code(content: str) -> ElementNode:
    """Code block (monospace, no spacing)."""
    return _node("code", content)


# ---------------------------------------------------------------------------
# Layer 5 — content: shapes
# ---------------------------------------------------------------------------

def box(*children: Node | str | None, **props: Any) -> ElementNode:
    """
    Rectangle shape. Can contain text and paragraphs.

    dok.box("Important notice", fill="lightblue", rounded=True)
    dok.box(dok.bold("Key insight"), dok.p("Details here."), fill="navy", color="white")
    """
    return _node("box", *children, **props)

def circle(*children: Node | str | None, **props: Any) -> ElementNode:
    """Circle / ellipse shape."""
    return _node("circle", *children, **props)

def diamond(*children: Node | str | None, **props: Any) -> ElementNode:
    """Diamond shape (decisions in flowcharts)."""
    return _node("diamond", *children, **props)

def chevron(*children: Node | str | None, **props: Any) -> ElementNode:
    """Chevron / process step shape."""
    return _node("chevron", *children, **props)

def callout(*children: Node | str | None, **props: Any) -> ElementNode:
    """
    Speech callout shape.

    dok.callout("Note: this is preliminary.", fill="lightyellow", stroke="orange")
    """
    return _node("callout", *children, **props)

def badge(content: str, **props: Any) -> ElementNode:
    """
    Small inline label shape.

    dok.badge("APPROVED", fill="green", color="white")
    """
    return _node("badge", content, **props)

def banner(*children: Node | str | None, **props: Any) -> ElementNode:
    """
    Full-width decorative block. Good for page headers.

    dok.banner(
        dok.bold("Acme Corporation", color="white", size=16),
        fill="navy",
        accent="gold",  ← colored left-edge bar
    )
    """
    return _node("banner", *children, **props)

def line(**props: Any) -> ElementNode:
    """
    Horizontal rule / divider.

    dok.line()                      ← default thin gray line
    dok.line(stroke="blue", thick=True)
    dok.line(stroke="gray", dashed=True)
    """
    return _node("line", **props)

def page_break() -> ElementNode:
    """Insert a page break."""
    return ElementNode(name="---", props={}, children=[])


# ---------------------------------------------------------------------------
# Connectors (used inside row())
# ---------------------------------------------------------------------------

def arrow(label: str | None = None) -> ArrowNode:
    """
    Arrow connector between siblings in a row().

    dok.row(
        dok.box("A"),
        dok.arrow(),
        dok.box("B"),
        dok.arrow("validates"),
        dok.box("C"),
    )
    """
    return ArrowNode(label=label)