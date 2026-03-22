"""
dok.registry
~~~~~~~~~~~~
Central element registry — single source of truth for every element's
name, category, allowed props, parent constraints, and converter handler.

Adding a new element = one register() call here.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any


# ---------------------------------------------------------------------------
# Schema types
# ---------------------------------------------------------------------------

@dataclass
class PropDef:
    type: str                                # "color"|"int"|"string"|"bool"|"ratio"|"enum"
    default: Any = None
    required: bool = False
    choices: tuple[str, ...] | None = None   # valid values for "enum" type


@dataclass
class ElementDef:
    name: str
    category: str                            # doc|page|layout|style|block|container|list|table|inline|meta|drawing
    props: dict[str, PropDef] = field(default_factory=dict)
    parent_must_be: set[str] | None = None   # e.g. {"ul","ol"} for li
    handler: str | None = None               # converter method name


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

ELEMENTS: dict[str, ElementDef] = {}


def register(e: ElementDef) -> None:
    ELEMENTS[e.name] = e


def get(name: str) -> ElementDef | None:
    return ELEMENTS.get(name)


def categories(cat: str) -> set[str]:
    """Return all element names in a category."""
    return {e.name for e in ELEMENTS.values() if e.category == cat}


# ---------------------------------------------------------------------------
# Prop shorthand helpers
# ---------------------------------------------------------------------------

_COLOR  = PropDef("color")
_INT    = PropDef("int")
_STRING = PropDef("string")
_BOOL   = PropDef("bool")
_RATIO  = PropDef("ratio")
_MARGIN = PropDef("enum", choices=("normal", "narrow", "wide", "none"))
_PAPER  = PropDef("enum", choices=("a4", "letter", "a3"))
_ALIGN  = PropDef("enum", choices=("left", "right"))
_TAIL   = PropDef("enum", choices=("top-left", "top-right", "bottom-left", "bottom-right"))

# Common prop sets reused across multiple elements
_STYLE_PROPS = {"color": _COLOR, "size": _INT, "font": _STRING}


# ---------------------------------------------------------------------------
# Register all elements
# ---------------------------------------------------------------------------

# Document structure
register(ElementDef("doc", "doc",
    props={"font": _STRING, "size": _INT, "spacing": PropDef("enum", choices=("compact", "tight", "normal", "relaxed"))},
    handler="_handle_doc"))
register(ElementDef("page", "page",
    props={"margin": _MARGIN, "paper": _PAPER, "cols": _INT},
    handler="_handle_page"))

# Layout
register(ElementDef("center",  "layout"))
register(ElementDef("right",   "layout"))
register(ElementDef("left",    "layout"))
register(ElementDef("justify", "layout"))
register(ElementDef("rtl",     "layout"))
register(ElementDef("ltr",     "layout"))
register(ElementDef("indent",  "layout", props={"level": _INT}))
register(ElementDef("row",     "layout", handler="_handle_row"))
register(ElementDef("cols",    "layout", props={"ratio": _RATIO}, handler="_handle_cols"))
register(ElementDef("col",     "layout", parent_must_be={"cols"}))
register(ElementDef("float",   "layout", props={"side": _ALIGN}, handler="_handle_float"))

# Style
for _name in ("bold", "italic", "underline", "strike"):
    register(ElementDef(_name, "style", props=dict(_STYLE_PROPS)))
register(ElementDef("sup",       "style"))
register(ElementDef("sub",       "style"))
register(ElementDef("color",     "style", props={"value": _COLOR, "size": _INT, "font": _STRING}))
register(ElementDef("size",      "style", props={"value": _INT, "color": _COLOR, "font": _STRING}))
register(ElementDef("font",      "style", props={"value": _STRING, "size": _INT, "color": _COLOR}))
register(ElementDef("highlight", "style", props={"value": _STRING, "size": _INT, "color": _COLOR}))
register(ElementDef("span",      "style", props={"bold": _BOOL, "italic": _BOOL, "underline": _BOOL,
                                                   "color": _COLOR, "size": _INT, "font": _STRING}))

# Block content
_BLOCK_PROPS = {"color": _COLOR, "size": _INT,
                "spacing": PropDef("enum", choices=("compact", "tight", "normal", "relaxed")),
                "line-height": _INT}
for _name in ("h1", "h2", "h3", "h4", "p"):
    register(ElementDef(_name, "block", props=dict(_BLOCK_PROPS)))
register(ElementDef("quote", "block"))
register(ElementDef("code",  "block"))

# Containers — box is the universal container (banner/callout/badge are variants)
_BOX_PROPS = {
    "fill": _COLOR, "stroke": _COLOR, "color": _COLOR,
    "rounded": _BOOL, "shadow": _BOOL,
    "accent": _COLOR, "inline": _BOOL,
    "width": _INT, "height": _INT,
}
register(ElementDef("box",     "container", props=dict(_BOX_PROPS), handler="_emit_box"))
register(ElementDef("callout", "container", props={**_BOX_PROPS, "tail": _TAIL}, handler="_emit_box"))
register(ElementDef("banner",  "container", props=dict(_BOX_PROPS), handler="_emit_box"))
register(ElementDef("badge",   "container", props=dict(_BOX_PROPS), handler="_emit_box"))
register(ElementDef("line",    "container",
    props={"stroke": _COLOR, "dashed": _BOOL, "thick": _BOOL},
    handler="_emit_line"))

# Drawing shapes
for _name in ("circle", "diamond", "chevron"):
    register(ElementDef(_name, "drawing",
        props={"fill": _COLOR, "stroke": _COLOR, "color": _COLOR, "shadow": _BOOL},
        handler="_emit_drawing_shape"))

# Lists
register(ElementDef("ul", "list", handler="_emit_list"))
register(ElementDef("ol", "list", props={"start": _INT}, handler="_emit_list"))
register(ElementDef("li", "list", parent_must_be={"ul", "ol"}))

# Tables
register(ElementDef("table", "table",
    props={"border": _BOOL, "striped": _BOOL},
    handler="_emit_data_table"))
register(ElementDef("tr", "table", parent_must_be={"table"}))
register(ElementDef("td", "table", props={"colspan": _INT}, parent_must_be={"tr"}))
register(ElementDef("th", "table", props={"colspan": _INT}, parent_must_be={"tr"}))

# Inline / meta
register(ElementDef("img", "inline",
    props={"src": PropDef("string", required=True), "width": _INT, "height": _INT, "alt": _STRING},
    handler="_emit_image"))
register(ElementDef("link", "inline",
    props={"href": PropDef("string", required=True)},
    handler="_emit_link_paragraph"))
register(ElementDef("page-number", "inline", handler="_emit_page_number"))
register(ElementDef("space", "meta", props={"size": _INT}, handler="_emit_spacer"))
register(ElementDef("header", "meta",
    parent_must_be={"doc", "page"},
    handler="_emit_header"))
register(ElementDef("footer", "meta",
    parent_must_be={"doc", "page"},
    handler="_emit_footer"))

# Special
register(ElementDef("---", "special", handler="_handle_page_break"))
