"""
dok.validator
~~~~~~~~~~~~~
Validates the resolved AST before conversion.

Checks three categories:
  1. **Structure** — nesting rules (from registry parent_must_be)
  2. **Props**     — known prop names, correct types, valid values
  3. **Printable** — constraints that keep output looking good on paper
"""

from __future__ import annotations
from typing import Any

from .nodes import Node, ElementNode, TextNode, ArrowNode, SHAPE_PRESETS
from . import registry
from .colors import resolve as resolve_color
from .errors import ValidationError, ValidationErrors, SourceLoc


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

        elem = registry.get(name)

        self._check_structure(name, elem, parent_stack, loc)
        self._check_required_props(name, elem, node.props, loc)

        if elem:
            self._check_props(node, elem, loc)
        else:
            if node.props:
                self._err(f"Unknown element '{name}'", loc=loc)

        self._check_font_size(node, loc)

        child_stack = parent_stack + [name]
        for child in node.children:
            self.walk(child, child_stack)

    def _check_structure(self, name: str, elem: registry.ElementDef | None,
                         parents: list[str], loc: SourceLoc | None) -> None:
        parent = parents[-1] if parents else None

        # Registry-driven parent constraints
        if elem and elem.parent_must_be:
            if parent not in elem.parent_must_be:
                allowed = ", ".join(sorted(elem.parent_must_be))
                self._err(f"'{name}' must be inside {{{allowed}}}", loc=loc)

        # Additional structural rules not captured by parent_must_be
        if name == "page" and parent and parent != "doc":
            self._err(f"'page' should be inside 'doc', not '{parent}'", loc=loc)

        elif name == "doc" and parents:
            self._err("'doc' must be the root element", loc=loc)

        elif name in SHAPE_PRESETS and parent:
            style_names = registry.categories("style")
            if parent in style_names:
                self._err(f"Shape '{name}' inside style '{parent}' — "
                          f"shapes should be at block level", loc=loc,
                          hint=f"Move '{name}' outside of '{parent}'.")

    def _check_required_props(self, name: str, elem: registry.ElementDef | None,
                              props: dict, loc: SourceLoc | None) -> None:
        if not elem:
            return
        for key, prop_def in elem.props.items():
            if prop_def.required and key not in props:
                self._err(f"'{name}' requires a '{key}' property", loc=loc,
                          hint=f'{name}({key}: "value")')

    def _check_props(self, node: ElementNode, elem: registry.ElementDef,
                     loc: SourceLoc | None) -> None:
        for key, value in node.props.items():
            prop_def = elem.props.get(key)
            if not prop_def:
                known = ", ".join(sorted(elem.props.keys())) if elem.props else "none"
                self._err(f"Unknown property '{key}' on '{node.name}'", loc=loc,
                          hint=f"Known properties: {known}.")
                continue
            self._check_prop_value(node.name, key, value, prop_def, loc)

    def _check_prop_value(self, elem_name: str, key: str, value: Any,
                          prop_def: registry.PropDef, loc: SourceLoc | None) -> None:
        ptype = prop_def.type

        if ptype == "color":
            if isinstance(value, str) and value not in ("none", "dashed", "dotted", "thick", "thin"):
                if resolve_color(value) is None:
                    self._err(f"Invalid color '{value}' for '{elem_name}.{key}'", loc=loc,
                              hint="Use a named color (red, navy, gold, ...) or hex (#FF0000, #ABC).")

        elif ptype == "int":
            if not isinstance(value, int):
                try:
                    int(value)
                except (ValueError, TypeError):
                    self._err(f"'{elem_name}.{key}' must be an integer, got '{value}'", loc=loc)

        elif ptype == "bool":
            if value not in (True, False, "true", "false"):
                self._err(f"'{elem_name}.{key}' must be a boolean flag", loc=loc,
                          hint=f"Use: {elem_name}({key}) for true, or omit for false.")

        elif ptype == "ratio":
            if isinstance(value, str):
                parts = value.split(":")
                if not all(p.isdigit() and int(p) > 0 for p in parts):
                    self._err(f"Invalid ratio '{value}'", loc=loc,
                              hint="Ratios: 1:1, 2:1, 1:1:1")

        elif ptype == "enum":
            if prop_def.choices and isinstance(value, str) and value not in prop_def.choices:
                valid = ", ".join(sorted(prop_def.choices))
                self._err(f"Invalid value '{value}' for '{elem_name}.{key}'", loc=loc,
                          hint=f"Valid values: {valid}.")

        elif ptype == "string":
            pass  # strings are always valid

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
