"""
dok.template
~~~~~~~~~~~~
Template resolution: variables, loops, conditionals.

Pipeline position: parse → resolve_imports → **resolve_templates** → resolve → validate → convert

Expands LetNode, EachNode, IfNode into regular ElementNode/TextNode.
Substitutes $var references in text and prop values.
"""

from __future__ import annotations
import copy
import re
from typing import Any

from .nodes import (
    Node, ElementNode, TextNode, ArrowNode,
    FunctionDefNode, ImportNode,
    LetNode, EachNode, IfNode,
)
from .errors import ResolveError


_VAR_RE = re.compile(r'\$([a-zA-Z_][a-zA-Z0-9_]*)')


def resolve_templates(nodes: list[Node]) -> list[Node]:
    """Expand all template constructs (let/each/if) in the node tree."""
    scope: dict[str, Any] = {}
    return _resolve_list(nodes, scope)


def _resolve_list(nodes: list[Node], scope: dict[str, Any]) -> list[Node]:
    result: list[Node] = []
    for node in nodes:
        result.extend(_resolve_node(node, scope))
    return result


def _resolve_node(node: Node, scope: dict[str, Any]) -> list[Node]:
    if isinstance(node, LetNode):
        # Evaluate the value (could reference other variables)
        value = _eval_value(node.value, scope)
        scope[node.name] = value
        return []  # let nodes produce no output

    if isinstance(node, EachNode):
        iterable = scope.get(node.iterable)
        if iterable is None:
            raise ResolveError(
                f"Undefined variable '{node.iterable}' in each loop",
                loc=node.loc,
            )
        if not isinstance(iterable, (list, tuple)):
            raise ResolveError(
                f"Variable '{node.iterable}' is not iterable (got {type(iterable).__name__})",
                loc=node.loc,
            )
        result: list[Node] = []
        for idx, item in enumerate(iterable):
            child_scope = dict(scope)
            child_scope[node.var_name] = item
            if node.index_var:
                child_scope[node.index_var] = idx
            body_copy = copy.deepcopy(node.body)
            result.extend(_resolve_list(body_copy, child_scope))
        return result

    if isinstance(node, IfNode):
        if _eval_condition(node.condition, scope):
            return _resolve_list(copy.deepcopy(node.then_body), scope)
        for elif_cond, elif_body in node.elif_clauses:
            if _eval_condition(elif_cond, scope):
                return _resolve_list(copy.deepcopy(elif_body), scope)
        if node.else_body:
            return _resolve_list(copy.deepcopy(node.else_body), scope)
        return []

    if isinstance(node, TextNode):
        new_text = _substitute_vars(node.text, scope)
        return [TextNode(text=new_text, loc=node.loc)]

    if isinstance(node, ElementNode):
        # If element name matches a scope variable and has no props/children,
        # replace with a TextNode (e.g. `item` inside `each item in items`)
        if (node.name in scope
                and not node.props
                and not node.children):
            val = scope[node.name]
            return [TextNode(text=str(val), loc=node.loc)]

        # Substitute in props
        new_props = {}
        for k, v in node.props.items():
            if isinstance(v, str):
                new_props[k] = _substitute_vars(v, scope)
            else:
                new_props[k] = v
        # Recurse into children
        new_children = _resolve_list(node.children, scope)
        return [ElementNode(name=node.name, props=new_props,
                            children=new_children, loc=node.loc)]

    # Pass through FunctionDefNode, ImportNode, ArrowNode, etc.
    return [node]


def _substitute_vars(text: str, scope: dict[str, Any]) -> str:
    """Replace $var references in text with their values."""
    def replacer(m: re.Match) -> str:
        var_name = m.group(1)
        if var_name in scope:
            return str(scope[var_name])
        return m.group(0)  # leave unresolved vars as-is
    return _VAR_RE.sub(replacer, text)


def _eval_value(value: Any, scope: dict[str, Any]) -> Any:
    """Evaluate a value, resolving variable references."""
    if isinstance(value, str):
        # Check if it's a bare variable name
        if value in scope:
            return scope[value]
        return _substitute_vars(value, scope)
    if isinstance(value, list):
        return [_eval_value(item, scope) for item in value]
    return value


def _eval_condition(condition: list, scope: dict[str, Any]) -> bool:
    """Evaluate a condition expression.
    Condition is a list of (type, value) tuples from the parser."""
    values = _condition_to_values(condition, scope)

    if len(values) == 1:
        return bool(values[0])

    if len(values) == 3:
        left, op, right = values
        if op == "==":  return left == right
        if op == "!=":  return left != right
        if op == ">":   return _num(left) > _num(right)
        if op == "<":   return _num(left) < _num(right)
        if op == ">=":  return _num(left) >= _num(right)
        if op == "<=":  return _num(left) <= _num(right)

    # Default: truthy check on first value
    return bool(values[0]) if values else False


def _condition_to_values(condition: list, scope: dict[str, Any]) -> list:
    """Convert condition tokens to actual values."""
    values = []
    for tok_type, tok_val in condition:
        if tok_type == "var":
            values.append(scope.get(tok_val, None))
        elif tok_type == "name":
            if tok_val == "true":
                values.append(True)
            elif tok_val == "false":
                values.append(False)
            elif tok_val in scope:
                values.append(scope[tok_val])
            else:
                values.append(tok_val)
        elif tok_type == "num":
            values.append(tok_val)
        elif tok_type == "str":
            values.append(tok_val)
        elif tok_type == "op":
            values.append(tok_val)
    return values


def _num(v: Any) -> float:
    """Coerce a value to a number for comparison."""
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0
