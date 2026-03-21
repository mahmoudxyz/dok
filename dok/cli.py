#!/usr/bin/env python3
"""
dok — command-line interface

Usage:
    python -m dok input.dok output.docx
    python -m dok input.dok              # writes input.docx
    python -m dok --check input.dok      # parse only, report errors
    python -m dok --tree  input.dok      # print node tree, no output

Examples:
    python -m dok report.dok report.docx
    python -m dok --check template.dok
"""

import argparse
import sys
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="dok",
        description="Convert a .dok file to .docx",
    )
    parser.add_argument(
        "input",
        help="Input .dok file",
    )
    parser.add_argument(
        "output",
        nargs="?",
        help="Output .docx file (default: same name as input with .docx extension)",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Parse only — check for syntax errors, don't write output",
    )
    parser.add_argument(
        "--tree",
        action="store_true",
        help="Print the node tree to stdout, don't write output",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    source = input_path.read_text(encoding="utf-8")

    # --- Parse ---
    try:
        import dok
        node = dok.parse(source)
    except Exception as e:
        print(f"Parse error: {e}", file=sys.stderr)
        sys.exit(1)

    if args.check:
        print(f"OK: {input_path}")
        return

    if args.tree:
        _print_tree(node)
        return

    # --- Write ---
    output_path = Path(args.output) if args.output else input_path.with_suffix(".docx")

    try:
        import dok
        dok.to_docx(node, output_path)
        print(f"Written: {output_path}")
    except NotImplementedError:
        print("Error: converter not yet implemented", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


def _print_tree(node, indent: int = 0) -> None:
    """Print a node tree to stdout for debugging."""
    from dok.nodes import ElementNode, TextNode, ArrowNode

    prefix = "  " * indent

    if isinstance(node, TextNode):
        preview = node.text[:40].replace("\n", "\\n")
        print(f'{prefix}"{preview}"')

    elif isinstance(node, ArrowNode):
        label = f' "{node.label}"' if node.label else ""
        print(f"{prefix}->{label}")

    elif isinstance(node, ElementNode):
        props = ""
        if node.props:
            parts = [f"{k}={v!r}" for k, v in node.props.items()]
            props = f"({', '.join(parts)})"
        print(f"{prefix}{node.name}{props}")
        for child in node.children:
            _print_tree(child, indent + 1)

    else:
        print(f"{prefix}{node!r}")


if __name__ == "__main__":
    main()