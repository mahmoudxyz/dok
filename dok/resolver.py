"""
dok.resolver
~~~~~~~~~~~~
Two-phase resolution:
  1. Import resolution — reads imported files, injects their nodes.
  2. Function expansion — expands function calls by substituting parameters.

Pipeline position:  parse → **resolve_imports** → **resolve** → validate → convert
"""

from __future__ import annotations
import copy
from pathlib import Path
from typing import Any

from .nodes import (
    Node, ElementNode, TextNode, ArrowNode, FunctionDefNode,
    ImportNode, ALL_KNOWN_NODES,
)
from .errors import ResolveError, SourceLoc

_MAX_DEPTH = 16


# ---------------------------------------------------------------------------
# Import resolution
# ---------------------------------------------------------------------------

def resolve_imports(
    nodes: list[Node],
    base_dir: Path | None = None,
    _seen: set[str] | None = None,
) -> list[Node]:
    """Replace ImportNode instances with their parsed content.
    Must be called before resolve() so imported functions are available."""
    if base_dir is None:
        return nodes  # no base dir → skip imports

    if _seen is None:
        _seen = set()

    from .lexer import Lexer
    from .parser import Parser

    result: list[Node] = []
    for node in nodes:
        if isinstance(node, ImportNode):
            import_path = (base_dir / node.path).resolve()
            canon = str(import_path)

            if canon in _seen:
                raise ResolveError(
                    f"Circular import: {node.path}",
                    loc=node.loc,
                    hint="This file was already imported earlier in the chain.",
                )

            if not import_path.exists():
                raise ResolveError(
                    f"Import not found: {node.path}",
                    loc=node.loc,
                    hint=f"Looked for: {import_path}",
                )

            _seen.add(canon)
            source = import_path.read_text(encoding="utf-8")
            tokens = Lexer(source).tokenize()
            imported = Parser(tokens).parse()
            # Recursively resolve imports in the imported file
            imported = resolve_imports(imported, import_path.parent, _seen)
            result.extend(imported)
        else:
            result.append(node)

    return result


# ---------------------------------------------------------------------------
# Function resolution
# ---------------------------------------------------------------------------

def resolve(nodes: list[Node]) -> list[Node]:
    """Expand all function definitions and calls. Returns a new node list."""
    funcs: dict[str, FunctionDefNode] = {}
    rest:  list[Node] = []

    for node in nodes:
        if isinstance(node, FunctionDefNode):
            if node.name in ALL_KNOWN_NODES:
                raise ResolveError(
                    f"Cannot define function '{node.name}' — "
                    f"it conflicts with a built-in element",
                    loc=node.loc,
                    hint="Choose a different name for your function.",
                )
            if node.name in funcs:
                raise ResolveError(
                    f"Duplicate function definition '{node.name}'",
                    loc=node.loc,
                    hint=f"A function named '{node.name}' was already defined.",
                )
            funcs[node.name] = node
        else:
            rest.append(node)

    if not funcs:
        return rest

    return _expand_list(rest, funcs, depth=0)


def _expand_list(
    nodes: list[Node],
    funcs: dict[str, FunctionDefNode],
    depth: int,
) -> list[Node]:
    result: list[Node] = []
    for node in nodes:
        result.extend(_expand_node(node, funcs, depth))
    return result


def _expand_node(
    node: Node,
    funcs: dict[str, FunctionDefNode],
    depth: int,
) -> list[Node]:
    if depth > _MAX_DEPTH:
        loc = node.loc if hasattr(node, "loc") else None
        raise ResolveError(
            "Maximum function expansion depth exceeded",
            loc=loc,
            hint="Check for circular function calls.",
        )

    if isinstance(node, (TextNode, ArrowNode)):
        return [node]

    if not isinstance(node, ElementNode):
        return [node]

    if node.name in funcs:
        func = funcs[node.name]
        return _instantiate(func, node, funcs, depth + 1)

    new_children = _expand_list(node.children, funcs, depth)
    return [ElementNode(
        name=node.name, props=node.props,
        children=new_children, loc=node.loc,
    )]


def _instantiate(
    func: FunctionDefNode,
    call: ElementNode,
    funcs: dict[str, FunctionDefNode],
    depth: int,
) -> list[Node]:
    param_values: dict[str, Any] = {}
    for param in func.params:
        if param not in call.props:
            raise ResolveError(
                f"Missing parameter '{param}' in call to '{func.name}'",
                loc=call.loc,
                hint=f"Usage: {func.name}({', '.join(p + ': ...' for p in func.params)})",
            )
        param_values[param] = call.props[param]

    for key in call.props:
        if key not in func.params:
            raise ResolveError(
                f"Unknown parameter '{key}' in call to '{func.name}'",
                loc=call.loc,
                hint=f"'{func.name}' accepts: {', '.join(func.params) or 'no parameters'}",
            )

    body = copy.deepcopy(func.body)
    body = _substitute(body, param_values, call.children)
    return _expand_list(body, funcs, depth)


def _substitute(
    nodes: list[Node],
    params: dict[str, Any],
    call_children: list[Node],
) -> list[Node]:
    result: list[Node] = []

    for node in nodes:
        if isinstance(node, ElementNode):
            if (node.name in params
                    and not node.props
                    and not node.children):
                value = params[node.name]
                result.append(TextNode(text=str(value), loc=node.loc))
                continue

            if (node.name == "children"
                    and not node.props
                    and not node.children):
                result.extend(copy.deepcopy(call_children))
                continue

            node.children = _substitute(node.children, params, call_children)

            new_props: dict[str, Any] = {}
            for k, v in node.props.items():
                if isinstance(v, str) and v in params:
                    new_props[k] = params[v]
                else:
                    new_props[k] = v
            node.props = new_props

            result.append(node)
        else:
            result.append(node)

    return result
