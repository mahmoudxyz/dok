"""
dok.api
~~~~~~~
The public functions that glue the pipeline together.

Pipeline:  source → lex → parse → resolve_imports → resolve → validate → convert → write

  parse(source)          string → Node tree (resolved + validated)
  to_docx(node, path)    Node tree → .docx file
  to_bytes(node)         Node tree → bytes (for web responses, etc.)
"""

from __future__ import annotations
import io
from pathlib import Path

from .nodes       import Node, ElementNode
from .lexer       import Lexer
from .parser      import Parser
from .resolver    import resolve_imports, resolve
from .validator   import validate
from .converter   import Converter
from .docx_writer import DocxWriter
from .html_writer import HtmlWriter



def parse(source: str, *, base_dir: Path | str | None = None) -> ElementNode:
    """
    Parse a Dok string and return the root node.

    Full pipeline: lex → parse → resolve imports → resolve functions → validate.

    Args:
        source:    Dok source code
        base_dir:  Directory for resolving imports (usually the input file's parent)

    If the source contains a top-level doc { } node, that is returned.
    Otherwise the nodes are wrapped in an implicit doc node.

    Raises:
        LexError          on tokenisation errors
        ParseError        on syntax errors
        ResolveError      on function/import expansion errors
        ValidationErrors  on semantic validation errors
    """
    tokens = Lexer(source).tokenize()
    nodes  = Parser(tokens).parse()

    # Resolve imports (must come before function resolution)
    if base_dir is not None:
        nodes = resolve_imports(nodes, Path(base_dir))

    # Expand function definitions and calls
    nodes = resolve(nodes)

    # Wrap in doc if needed (before validation, so validator sees the root)
    if len(nodes) == 1 and isinstance(nodes[0], ElementNode) and nodes[0].name == "doc":
        root = nodes[0]
    else:
        root = ElementNode(name="doc", props={}, children=nodes)

    # Validate the resolved tree
    validate([root])

    return root


def to_docx(node: Node, dest: str | Path, *, base_dir: Path | str | None = None) -> None:
    """
    Convert a node tree to a .docx file at *dest*.

    Args:
        node:      Root node (from parse() or builder API)
        dest:      Output file path (created or overwritten)
        base_dir:  Directory for resolving image paths
    """
    model  = Converter().convert(_ensure_list(node))
    if base_dir is not None:
        model.base_dir = Path(base_dir)
    writer = DocxWriter(model)
    writer.write(dest)


def to_html(node: Node, dest: str | Path, *, base_dir: Path | str | None = None) -> None:
    """
    Convert a node tree to a .html file at *dest*.

    Args:
        node:      Root node (from parse() or builder API)
        dest:      Output file path (created or overwritten)
        base_dir:  Directory for resolving image paths
    """
    model  = Converter().convert(_ensure_list(node))
    if base_dir is not None:
        model.base_dir = Path(base_dir)
    writer = HtmlWriter(model)
    writer.write(dest)



def to_bytes(node: Node, *, base_dir: Path | str | None = None) -> bytes:
    """
    Convert a node tree to .docx bytes (useful for HTTP responses, S3, etc.).

    Returns:
        bytes of a valid .docx file
    """
    model  = Converter().convert(_ensure_list(node))
    if base_dir is not None:
        model.base_dir = Path(base_dir)
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
