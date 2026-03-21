"""
dok.validator
~~~~~~~~~~~~~
Validates the resolved AST before conversion.

Checks three categories:
  1. **Structure** — nesting rules
  2. **Props**     — known prop names, correct types, valid values
  3. **Printable** — constraints that keep output looking good on paper
"""

from __future__ import annotations
from typing import Any

from .nodes import (
    Node, ElementNode, TextNode, ArrowNode,
    DOC_NODES, PAGE_NODES, LAYOUT_NODES, STYLE_NODES,
    BLOCK_NODES, SHAPE_NODES, SHAPE_PRESETS, ALL_KNOWN_NODES,
    LIST_NODES, TABLE_NODES, INLINE_NODES, META_NODES,
)
from .colors import resolve as resolve_color
from .errors import ValidationError, ValidationErrors, SourceLoc


# ---------------------------------------------------------------------------
# Prop type constants
# ---------------------------------------------------------------------------

_COLOR  = "color"
_INT    = "int"
_STRING = "string"
_BOOL   = "bool"
_RATIO  = "ratio"
_MARGIN = "margin"
_PAPER  = "paper"
_ALIGN  = "align"
_TAIL   = "tail"

# ---------------------------------------------------------------------------
# Prop schemas
# ---------------------------------------------------------------------------

_PROP_SCHEMAS: dict[str, dict[str, str]] = {
    "doc":       {"font": _STRING, "size": _INT},
    "page":      {"margin": _MARGIN, "paper": _PAPER, "cols": _INT},
    "indent":    {"level": _INT},
    "cols":      {"ratio": _RATIO},
    "float":     {"side": _ALIGN},

    # Style nodes
    "color":     {"value": _COLOR, "size": _INT, "font": _STRING},
    "size":      {"value": _INT, "color": _COLOR, "font": _STRING},
    "font":      {"value": _STRING, "size": _INT, "color": _COLOR},
    "highlight": {"value": _STRING, "size": _INT, "color": _COLOR},
    "bold":      {"color": _COLOR, "size": _INT, "font": _STRING},
    "italic":    {"color": _COLOR, "size": _INT, "font": _STRING},
    "underline": {"color": _COLOR, "size": _INT, "font": _STRING},
    "strike":    {"color": _COLOR, "size": _INT, "font": _STRING},
    "span":      {"bold": _BOOL, "italic": _BOOL, "underline": _BOOL,
                  "color": _COLOR, "size": _INT, "font": _STRING},

    # Block nodes
    "h1": {"color": _COLOR, "size": _INT},
    "h2": {"color": _COLOR, "size": _INT},
    "h3": {"color": _COLOR, "size": _INT},
    "h4": {"color": _COLOR, "size": _INT},
    "p":  {"color": _COLOR, "size": _INT},

    # Shape-like nodes
    "box":     {"fill": _COLOR, "stroke": _COLOR, "color": _COLOR,
                "rounded": _BOOL, "shadow": _BOOL},
    "circle":  {"fill": _COLOR, "stroke": _COLOR, "color": _COLOR, "shadow": _BOOL},
    "diamond": {"fill": _COLOR, "stroke": _COLOR, "color": _COLOR, "shadow": _BOOL},
    "chevron": {"fill": _COLOR, "stroke": _COLOR, "color": _COLOR, "shadow": _BOOL},
    "callout": {"fill": _COLOR, "stroke": _COLOR, "color": _COLOR,
                "tail": _TAIL, "shadow": _BOOL},
    "badge":   {"fill": _COLOR, "stroke": _COLOR, "color": _COLOR},
    "banner":  {"fill": _COLOR, "stroke": _COLOR, "color": _COLOR, "accent": _COLOR},
    "line":    {"stroke": _COLOR, "dashed": _BOOL, "thick": _BOOL},

    # Lists
    "ul": {},
    "ol": {"start": _INT},
    "li": {},

    # Tables
    "table": {"border": _BOOL, "striped": _BOOL},
    "tr":    {},
    "td":    {"colspan": _INT},
    "th":    {"colspan": _INT},

    # Inline
    "img":         {"src": _STRING, "width": _INT, "height": _INT, "alt": _STRING},
    "link":        {"href": _STRING},
    "page-number": {},

    # Meta
    "header": {},
    "footer": {},
    "space":  {"size": _INT},
}

_VALID_MARGINS = {"normal", "narrow", "wide", "none"}
_VALID_PAPERS  = {"a4", "letter", "a3"}
_VALID_ALIGNS  = {"left", "right"}
_VALID_TAILS   = {"top-left", "top-right", "bottom-left", "bottom-right"}

_MIN_FONT_SIZE = 6
_MAX_FONT_SIZE = 96
_MAX_NESTING   = 12


# ---------------------------------------------------------------------------
# Public
# ---------------------------------------------------------------------------

def validate(nodes: list[Node]) -> None:
    ctx = _Ctx()
    for node in nodes:
        ctx.walk(node, parent_stack=[])
    if ctx.errors:
        raise ValidationErrors(ctx.errors)


# ---------------------------------------------------------------------------
# Internal
# ---------------------------------------------------------------------------

class _Ctx:
    def __init__(self) -> None:
        self.errors: list[ValidationError] = []

    def _err(self, msg: str, loc: SourceLoc | None, hint: str | None = None) -> None:
        self.errors.append(ValidationError(msg, loc=loc, hint=hint))

    def walk(self, node: Node, parent_stack: list[str]) -> None:
        if isinstance(node, TextNode):
            return

        if isinstance(node, ArrowNode):
            if not parent_stack or parent_stack[-1] != "row":
                self._err("Arrow '->' used outside of a row { }",
                          loc=node.loc,
                          hint="Arrows connect shapes inside row { shape -> shape }.")
            return

        if not isinstance(node, ElementNode):
            return

        name = node.name
        loc  = node.loc

        if len(parent_stack) > _MAX_NESTING:
            self._err(f"Nesting too deep ({len(parent_stack)} levels)", loc=loc,
                      hint=f"Maximum nesting depth is {_MAX_NESTING}.")
            return

        self._check_structure(name, parent_stack, loc)
        self._check_required_props(name, node.props, loc)

        if name in _PROP_SCHEMAS:
            self._check_props(node, _PROP_SCHEMAS[name], loc)
        elif name in ALL_KNOWN_NODES:
            if node.props:
                self._err(f"'{name}' does not accept any properties", loc=loc)

        self._check_font_size(node, loc)

        child_stack = parent_stack + [name]
        for child in node.children:
            self.walk(child, child_stack)

    def _check_structure(self, name: str, parents: list[str], loc: SourceLoc | None) -> None:
        parent = parents[-1] if parents else None

        if name == "col" and parent != "cols":
            self._err("'col' must be a direct child of 'cols'", loc=loc,
                      hint="cols(ratio: 1:1) { col { ... } col { ... } }")

        elif name == "li" and parent not in ("ul", "ol"):
            self._err("'li' must be inside 'ul' or 'ol'", loc=loc,
                      hint="ul { li { \"item\" } }")

        elif name == "tr" and parent != "table":
            self._err("'tr' must be inside 'table'", loc=loc,
                      hint="table { tr { td { \"cell\" } } }")

        elif name in ("td", "th") and parent != "tr":
            self._err(f"'{name}' must be inside 'tr'", loc=loc)

        elif name == "page" and parent and parent != "doc":
            self._err(f"'page' should be inside 'doc', not '{parent}'", loc=loc)

        elif name == "doc" and parents:
            self._err("'doc' must be the root element", loc=loc)

        elif name in ("header", "footer"):
            if parent and parent not in ("doc", "page"):
                self._err(f"'{name}' should be inside 'doc' or 'page'", loc=loc)

        elif name in SHAPE_NODES and parent and parent in STYLE_NODES:
            self._err(f"Shape '{name}' inside style '{parent}' — "
                      f"shapes should be at block level", loc=loc,
                      hint=f"Move '{name}' outside of '{parent}'.")

    def _check_required_props(self, name: str, props: dict, loc: SourceLoc | None) -> None:
        if name == "img" and "src" not in props:
            self._err("'img' requires a 'src' property", loc=loc,
                      hint='img(src: "photo.png", width: 4)')

        elif name == "link" and "href" not in props:
            self._err("'link' requires an 'href' property", loc=loc,
                      hint='link(href: "https://example.com") { "click here" }')

    def _check_props(self, node: ElementNode, schema: dict[str, str],
                     loc: SourceLoc | None) -> None:
        for key, value in node.props.items():
            if key not in schema:
                known = ", ".join(sorted(schema.keys())) if schema else "none"
                self._err(f"Unknown property '{key}' on '{node.name}'", loc=loc,
                          hint=f"Known properties: {known}.")
                continue
            self._check_prop_value(node.name, key, value, schema[key], loc)

    def _check_prop_value(self, elem: str, key: str, value: Any,
                          expected: str, loc: SourceLoc | None) -> None:
        if expected == _COLOR:
            if isinstance(value, str) and value not in ("none", "dashed", "dotted", "thick", "thin"):
                if resolve_color(value) is None:
                    self._err(f"Invalid color '{value}' for '{elem}.{key}'", loc=loc,
                              hint="Use a named color (red, navy, gold, ...) or hex (#FF0000, #ABC).")

        elif expected == _INT:
            if not isinstance(value, int):
                try:
                    int(value)
                except (ValueError, TypeError):
                    self._err(f"'{elem}.{key}' must be an integer, got '{value}'", loc=loc)

        elif expected == _BOOL:
            if value not in (True, False, "true", "false"):
                self._err(f"'{elem}.{key}' must be a boolean flag", loc=loc,
                          hint=f"Use: {elem}({key}) for true, or omit for false.")

        elif expected == _MARGIN:
            if isinstance(value, str) and value not in _VALID_MARGINS:
                self._err(f"Invalid margin '{value}'", loc=loc,
                          hint=f"Valid margins: {', '.join(sorted(_VALID_MARGINS))}.")

        elif expected == _PAPER:
            if isinstance(value, str) and value not in _VALID_PAPERS:
                self._err(f"Invalid paper size '{value}'", loc=loc,
                          hint=f"Valid paper sizes: {', '.join(sorted(_VALID_PAPERS))}.")

        elif expected == _RATIO:
            if isinstance(value, str):
                parts = value.split(":")
                if not all(p.isdigit() and int(p) > 0 for p in parts):
                    self._err(f"Invalid ratio '{value}'", loc=loc,
                              hint="Ratios: 1:1, 2:1, 1:1:1")

        elif expected == _ALIGN:
            if isinstance(value, str) and value not in _VALID_ALIGNS:
                self._err(f"Invalid alignment '{value}'", loc=loc)

        elif expected == _TAIL:
            if isinstance(value, str) and value not in _VALID_TAILS:
                self._err(f"Invalid tail position '{value}'", loc=loc,
                          hint=f"Valid: {', '.join(sorted(_VALID_TAILS))}.")

    def _check_font_size(self, node: ElementNode, loc: SourceLoc | None) -> None:
        size_val = None
        if node.name in ("size", "doc"):
            size_val = node.props.get("value" if node.name == "size" else "size")
        elif node.name != "space" and "size" in node.props:
            size_val = node.props["size"]

        if size_val is not None:
            try:
                pt = int(size_val)
                if pt < _MIN_FONT_SIZE or pt > _MAX_FONT_SIZE:
                    self._err(f"Font size {pt}pt is outside printable range", loc=loc,
                              hint=f"Use a size between {_MIN_FONT_SIZE} and {_MAX_FONT_SIZE} points.")
            except (ValueError, TypeError):
                pass
