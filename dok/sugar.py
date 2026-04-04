"""
dok.sugar
~~~~~~~~~
Markdown-inspired syntax sugar — a preprocessor that converts
shorthand notations to standard Dok syntax before lexing.

Supported shortcuts:

  Line-level (must start at beginning of line):
    # Title            →  h1 "Title"
    ## Title           →  h2 "Title"
    ### Title          →  h3 "Title"
    #### Title         →  h4 "Title"
    > quote text       →  quote { "quote text" }
    - bullet item      →  (collected into ul { li { "..." } ... })
    * bullet item      →  (collected into ul { li { "..." } ... })
    1. numbered item   →  (collected into ol { li { "..." } ... })
    2. numbered item   →  (same, continues the list)

  Inline (within quoted strings):
    **bold text**      →  bold { "bold text" }
    *italic text*      →  italic { "italic text" }
    ~~struck text~~    →  strike { "struck text" }
    __underlined__     →  underline { "underlined" }
    `code text`        →  code { "code text" }
    [text](url)        →  link(href: "url") { "text" }

The preprocessor only transforms lines that are NOT inside
braces/parentheses or triple-quoted strings, so existing Dok
syntax passes through untouched.
"""

from __future__ import annotations
import re


def desugar(source: str) -> str:
    """Transform markdown-style sugar into standard Dok syntax."""
    lines = source.split("\n")
    result: list[str] = []
    in_triple_quote = False
    brace_depth = 0
    paren_depth = 0

    i = 0
    while i < len(lines):
        line = lines[i]

        # Track triple-quoted strings
        count = line.count('"""')
        if count % 2 == 1:
            in_triple_quote = not in_triple_quote

        if in_triple_quote:
            result.append(line)
            i += 1
            continue

        # Track brace/paren depth (rough — ignores strings but good enough
        # since sugar lines don't appear inside braces)
        for ch in line:
            if ch == '{':
                brace_depth += 1
            elif ch == '}':
                brace_depth = max(0, brace_depth - 1)
            elif ch == '(':
                paren_depth += 1
            elif ch == ')':
                paren_depth = max(0, paren_depth - 1)

        # Only transform when at top level (not inside braces/parens)
        if brace_depth > 0 or paren_depth > 0:
            result.append(line)
            i += 1
            continue

        stripped = line.strip()

        # Skip empty lines and comments
        if not stripped or stripped.startswith("//"):
            result.append(line)
            i += 1
            continue

        # --- Heading shortcuts: # ## ### #### ---
        heading_match = re.match(r'^(\s*)(#{1,4})\s+(.+)$', line)
        if heading_match:
            indent, hashes, text = heading_match.groups()
            level = len(hashes)
            text = _escape_dok_string(text.strip())
            result.append(f'{indent}h{level} "{text}"')
            i += 1
            continue

        # --- Quote shortcut: > text ---
        quote_match = re.match(r'^(\s*)>\s+(.+)$', line)
        if quote_match:
            indent, text = quote_match.groups()
            text = _escape_dok_string(text.strip())
            result.append(f'{indent}quote {{ "{text}" }}')
            i += 1
            continue

        # --- Bullet list: - item  or  * item (at line start) ---
        bullet_match = re.match(r'^(\s*)[-*]\s+(.+)$', line)
        if bullet_match and not stripped.startswith("---"):
            indent = bullet_match.group(1)
            items, i = _collect_list_items(lines, i, r'^(\s*)[-*]\s+(.+)$')
            result.append(f'{indent}ul {{')
            for item_text in items:
                item_text = _escape_dok_string(item_text)
                result.append(f'{indent}  li {{ "{item_text}" }}')
            result.append(f'{indent}}}')
            continue

        # --- Numbered list: 1. item, 2. item ---
        num_match = re.match(r'^(\s*)\d+\.\s+(.+)$', line)
        if num_match:
            indent = num_match.group(1)
            items, i = _collect_list_items(lines, i, r'^(\s*)\d+\.\s+(.+)$')
            result.append(f'{indent}ol {{')
            for item_text in items:
                item_text = _escape_dok_string(item_text)
                result.append(f'{indent}  li {{ "{item_text}" }}')
            result.append(f'{indent}}}')
            continue

        # --- No line-level sugar matched: pass through ---
        result.append(line)
        i += 1

    return "\n".join(result)


def desugar_inline(text: str) -> str:
    """Transform inline markdown patterns within a string into Dok elements.

    This is called on string literals during parsing to expand inline markup.
    Returns Dok source fragment with the inline elements expanded.

    Example:
        "Hello **world** and *everyone*"
        → '"Hello " bold { "world" } " and " italic { "everyone" }'
    """
    if not _has_inline_sugar(text):
        return ""  # empty = no transformation needed

    parts: list[str] = []
    pos = 0

    while pos < len(text):
        # Try each inline pattern
        match = _INLINE_PATTERN.search(text, pos)
        if not match:
            # Rest is plain text
            remainder = text[pos:]
            if remainder:
                parts.append(f'"{_escape_dok_string(remainder)}"')
            break

        # Add text before match
        if match.start() > pos:
            before = text[pos:match.start()]
            parts.append(f'"{_escape_dok_string(before)}"')

        # Determine which pattern matched
        if match.group("bold"):
            parts.append(f'bold {{ "{_escape_dok_string(match.group("bold"))}" }}')
        elif match.group("strike"):
            parts.append(f'strike {{ "{_escape_dok_string(match.group("strike"))}" }}')
        elif match.group("underline"):
            parts.append(f'underline {{ "{_escape_dok_string(match.group("underline"))}" }}')
        elif match.group("italic"):
            parts.append(f'italic {{ "{_escape_dok_string(match.group("italic"))}" }}')
        elif match.group("code"):
            parts.append(f'code {{ "{_escape_dok_string(match.group("code"))}" }}')
        elif match.group("link_text") is not None:
            link_text = _escape_dok_string(match.group("link_text"))
            link_url = _escape_dok_string(match.group("link_url"))
            parts.append(f'link(href: "{link_url}") {{ "{link_text}" }}')

        pos = match.end()

    return " ".join(parts)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

# Inline patterns — order matters (** before *, __ before _)
_INLINE_PATTERN = re.compile(
    r'\*\*(?P<bold>[^*]+)\*\*'           # **bold**
    r'|~~(?P<strike>[^~]+)~~'            # ~~strike~~
    r'|__(?P<underline>[^_]+)__'         # __underline__
    r'|\*(?P<italic>[^*]+)\*'            # *italic*
    r'|`(?P<code>[^`]+)`'               # `code`
    r'|\[(?P<link_text>[^\]]+)\]\((?P<link_url>[^)]+)\)'  # [text](url)
)


def _has_inline_sugar(text: str) -> bool:
    """Quick check if text contains any inline sugar markers."""
    return bool(_INLINE_PATTERN.search(text))


def _escape_dok_string(text: str) -> str:
    """Escape text for use inside a Dok quoted string."""
    return text.replace('\\', '\\\\').replace('"', '\\"')


def _collect_list_items(lines: list[str], start: int,
                        pattern: str) -> tuple[list[str], int]:
    """Collect consecutive list items matching the pattern."""
    items: list[str] = []
    i = start
    regex = re.compile(pattern)
    while i < len(lines):
        m = regex.match(lines[i])
        if m:
            items.append(m.group(2).strip())
            i += 1
        else:
            break
    return items, i
