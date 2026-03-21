"""
dok.api
~~~~~~~
The three public functions that glue everything together.

  parse(source)          string → Node tree
  to_docx(node, path)    Node tree → .docx file
  to_bytes(node)         Node tree → bytes (for web responses, etc.)
"""

from __future__ import annotations
import io
from pathlib import Path

from .nodes      import Node, ElementNode
from .lexer      import Lexer
from .parser     import Parser
from .converter  import Converter
from .docx_writer import DocxWriter


def parse(source: str) -> ElementNode:
    """
    Parse a Dok string and return the root node.

    If the source contains a top-level doc { } node, that is returned.
    Otherwise the nodes are wrapped in an implicit doc node.

    Raises:
        LexError   on tokenisation errors
        ParseError on syntax errors
    """
    tokens = Lexer(source).tokenize()
    nodes  = Parser(tokens).parse()

    # If the user wrote a top-level doc { }, return it directly
    if len(nodes) == 1 and isinstance(nodes[0], ElementNode) and nodes[0].name == "doc":
        return nodes[0]

    # Otherwise wrap in an implicit doc
    return ElementNode(name="doc", props={}, children=nodes)


def to_docx(node: Node, dest: str | Path) -> None:
    """
    Convert a node tree to a .docx file at *dest*.

    Args:
        node:  Root node (from parse() or builder API)
        dest:  Output file path (created or overwritten)

    Raises:
        NotImplementedError  until Converter and DocxWriter are implemented
    """
    model  = Converter().convert(_ensure_list(node))
    writer = DocxWriter(model)
    writer.write(dest)


def to_bytes(node: Node) -> bytes:
    """
    Convert a node tree to .docx bytes (useful for HTTP responses, S3, etc.).

    Returns:
        bytes of a valid .docx file
    """
    model  = Converter().convert(_ensure_list(node))
    writer = DocxWriter(model)
    buf    = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Internal
# ---------------------------------------------------------------------------

def _ensure_list(node: Node) -> list[Node]:
    """Converter expects a list; wrap a single root node."""
    return [node]