# Contributing to dok

This guide explains the codebase architecture so you can understand, debug, and extend dok.

## Project structure

```
dok/
├── __init__.py        # Public API re-exports
├── __main__.py        # CLI entry point
├── api.py             # Pipeline glue: parse(), to_docx(), to_html(), to_bytes()
├── lexer.py           # Source → tokens
├── parser.py          # Tokens → AST (node tree)
├── resolver.py        # Import resolution + function expansion
├── validator.py       # AST validation (structure, props, types)
├── registry.py        # Element registry — single source of truth
├── nodes.py           # Node dataclasses (ElementNode, TextNode, etc.)
├── context.py         # ParaCtx and RunCtx — immutable context structs
├── converter.py       # AST → DocxModel (intermediate representation)
├── models.py          # DocxModel dataclasses (ParagraphModel, BoxModel, etc.)
├── docx_writer.py     # DocxModel → .docx (OOXML ZIP)
├── html_writer.py     # DocxModel → .html
├── docx_styles.py     # OOXML style XML generation
├── docx_packaging.py  # OOXML packaging (namespaces, content types, rels)
├── xml_writer.py      # Low-level XML writer (shared by docx_writer)
├── constants.py       # Unit conversions, margins, paper sizes
├── colors.py          # Named color resolution
├── writer_utils.py    # Shared writer helpers
├── image.py           # Image dimension reading
├── errors.py          # Error types with source locations
├── builder.py         # Python builder API (dok.h1(), dok.p(), etc.)
├── template.py        # Template resolution (let/each/if → node expansion)
└── cli.py             # CLI argument parsing and dispatch
```

## The pipeline

Every document flows through this pipeline:

```
source string
    ↓
  Lexer          (lexer.py)        string → Token[]
    ↓
  Parser         (parser.py)       Token[] → Node[]
    ↓
  resolve_imports (resolver.py)    inline imported files
    ↓
  resolve_templates (template.py)  expand let/each/if, substitute $vars
    ↓
  resolve        (resolver.py)     expand function defs/calls
    ↓
  validate       (validator.py)    check structure + props
    ↓
  Converter      (converter.py)    Node[] → DocxModel
    ↓
  Writer         (docx_writer.py   DocxModel → .docx
                  html_writer.py)  DocxModel → .html
```

The builder API (`dok.h1()`, `dok.p()`, etc.) produces the same Node tree that the parser does, then feeds it through the same converter and writer.

## Key concepts

### The registry (`registry.py`)

The registry is the single source of truth for all elements. Every element has:

```python
@dataclass
class ElementDef:
    name: str                          # "box", "h1", "ul", etc.
    category: str                      # "doc"|"page"|"layout"|"style"|"block"|"container"|"list"|"table"|"inline"|"meta"|"drawing"
    props: dict[str, PropDef]          # allowed properties with types
    parent_must_be: set[str] | None    # structural constraint (e.g. "li" → {"ul", "ol"})
    handler: str | None                # converter method name (e.g. "_emit_box")
```

The validator, converter, and resolver all read from the registry. When you add a new element, this is where you register it.

### Nodes (`nodes.py`)

The AST consists of four node types:

- **`ElementNode`** — has a name, props dict, and children list
- **`TextNode`** — holds a string
- **`ArrowNode`** — arrow connector (`->`) inside rows
- **`FunctionDefNode`** / **`ImportNode`** — resolved before conversion
- **`LetNode`** / **`EachNode`** / **`IfNode`** — template nodes, resolved before function expansion

Everything the parser produces and the builder creates is one of these types.

### Context passing (`context.py`)

The converter walks the tree carrying two frozen dataclasses:

- **`ParaCtx`** — paragraph-level context: alignment, direction, indent, style, spacing
- **`RunCtx`** — run-level context: bold, italic, color, size, font

Wrapper nodes (layout, style) call `ctx.with_*()` to produce a new context. Leaf nodes consume the context. No globals, no mutation.

```
doc
  └── center              ← ParaCtx(align="center")
        └── bold          ← RunCtx(bold=True)
              └── h1      ← produces ParagraphModel with both contexts
```

### Models (`models.py`)

The converter produces a `DocxModel` containing a flat list of model objects:

- `ParagraphModel` — text paragraph with runs, style, spacing
- `BoxModel` — bordered/shaded container (selective borders, nesting-aware)
- `LineModel` — horizontal rule
- `DataTableModel` — data table with rows, cells, and auto-calculated column widths
- `TableModel` — layout table (cols/col)
- `ShapeModel` — drawing shape (circle, diamond, chevron)
- `RowModel` — horizontal flow of items
- `ImageModel` — embedded image
- `SpacerModel` — vertical gap
- `PageBreakModel` — page break
- `FrameModel` — positioned floating text box
- `ToggleModel` — collapsible section (HTML details/summary)
- `CheckboxModel` — form checkbox
- `TextInputModel` — form text input
- `DropdownModel` — form dropdown selector
- `TocModel` — table of contents

These models are the intermediate representation between the AST and the final output. Both writers consume the same models.

### Converter dispatch (`converter.py`)

The converter uses registry-driven dispatch instead of if/elif chains:

```python
def _walk(self, node, para, run):
    elem = registry.get(name)
    cat = elem.category

    if cat == "style":
        run = self._apply_style_props(node, run)
        for child in node.children: self._walk(child, para, run)
    elif cat == "layout" and not elem.handler:
        self._handle_layout(node, para, run)
    elif cat == "block":
        self._emit_paragraph(node, para, run)
    elif elem.handler:
        getattr(self, elem.handler)(node, para, run)
    else:
        for child in node.children: self._walk(child, para, run)
```

Categories group elements by behavior. Elements with special handling have a `handler` method name in the registry.

### Writer dispatch (`docx_writer.py`, `html_writer.py`)

Both writers use a type→method dict for dispatching model objects:

```python
_DISPATCH: dict[type, str] = {
    ParagraphModel: "_write_paragraph",
    BoxModel:       "_write_box",
    LineModel:      "_write_line",
    # ... etc
}

def _write_item(self, w, item):
    handler = self._DISPATCH.get(type(item))
    if handler:
        getattr(self, handler)(w, item)
```

## How to add a new element

Adding a new element requires changes in 4 places:

### 1. Register it (`registry.py`)

```python
register(ElementDef("my-element", "container",
    props={"fill": _COLOR, "size": _INT, "label": _STRING},
    handler="_emit_my_element"))
```

This immediately gives you:
- Prop validation (the validator knows what props are allowed and their types)
- Structure validation (if you set `parent_must_be`)
- The name is recognized as a known element

### 2. Add a model if needed (`models.py`)

If your element needs its own model (not reusing ParagraphModel, BoxModel, etc.):

```python
@dataclass
class MyElementModel:
    label: str = ""
    fill: str | None = None
    size: int = 0
```

### 3. Add a converter handler (`converter.py`)

```python
def _emit_my_element(self, node: ElementNode, para: ParaCtx, run: RunCtx) -> None:
    label = node.props.get("label", "")
    fill = resolve_color(str(node.props["fill"])) if "fill" in node.props else None
    self._model.content.append(MyElementModel(label=label, fill=fill))
```

The handler receives the AST node plus the current paragraph and run contexts.

### 4. Add render methods to both writers

**DOCX writer** (`docx_writer.py`):

```python
# Add to _DISPATCH dict:
MyElementModel: "_write_my_element",

# Add the method:
def _write_my_element(self, w: XmlWriter, item: MyElementModel) -> None:
    # Emit OOXML using w.open(), w.close(), w.tag(), w.raw()
    ...
```

**HTML writer** (`html_writer.py`):

```python
# Add to _DISPATCH dict:
MyElementModel: "_render_my_element",

# Add the method:
def _render_my_element(self, item: MyElementModel) -> str:
    return f'<div class="my-element">{html.escape(item.label)}</div>\n'
```

### 5. (Optional) Add a builder function (`builder.py`)

```python
def my_element(*children, **props):
    return _node("my-element", *children, **props)
```

And export it from `__init__.py`.

That's it. The validator, converter dispatch, and writer dispatch are all data-driven — no if/elif chains to update.

## How to add a new prop to an existing element

1. Add the prop to the element's `props` dict in `registry.py`
2. Add the field to the model dataclass in `models.py` (if it needs to reach the writers)
3. Read the prop in the converter handler and pass it to the model
4. Use the field in both writers' render methods

## Architecture decisions

### Why models exist between converter and writers

The converter walks the AST with context inheritance (ParaCtx, RunCtx). The writers know nothing about the AST — they just serialize model objects. This means:

- You can add a new output format by writing a new writer (PDF, Markdown, etc.)
- The converter logic is tested independently from output formatting
- Writers are simple serializers with no tree-walking logic

### Why box/callout/banner/badge are the same element

They are all rendered by the same `_emit_box` converter handler and produce `BoxModel` objects. The element name sets defaults:

```python
_BOX_DEFAULTS = {
    "callout": {"fill": "FFF2CC", "stroke": "FFC000"},
    "badge":   {"fill": "1F3864", "color": "FFFFFF", "inline": True},
}
```

The writer checks `box.accent` for banner/callout mode and `box.inline` for badge mode. This is composition over specialization — one model, multiple conventions.

### Why the lexer handles triple-quotes manually

Triple-quoted strings (`"""..."""`) are handled before the regex engine in `_scan()`. This is because:
- The regex `STRING` pattern matches single-line `"..."` strings
- Triple-quotes can contain unescaped `"` characters and newlines
- Dedenting is applied after extraction (stripping common leading whitespace)

The comment stripper also handles triple-quotes to avoid stripping `//` inside multiline strings.

## Testing

```bash
# Run all tests
python -m pytest tests/ -v

# Validate all examples
for f in examples/*.dok; do python -m dok --check "$f"; done

# End-to-end: compile an example
python -m dok examples/01-simple-letter.dok /tmp/test.docx
python -m dok examples/01-simple-letter.dok /tmp/test.html
```

### Test organization

```
tests/
├── test_lexer.py       # Token-level tests
├── test_parser.py      # AST structure tests
├── test_resolver.py    # Function/import expansion tests
├── test_validator.py   # Validation rule tests
├── test_converter.py   # Model generation tests
├── test_colors.py      # Color resolution tests
└── test_api.py         # End-to-end API tests
```

## OOXML primer

The `.docx` format is a ZIP archive containing XML files:

```
[Content_Types].xml     — file type declarations
_rels/.rels             — package relationships
word/document.xml       — the main document body
word/styles.xml         — paragraph and character styles
word/settings.xml       — document settings
word/numbering.xml      — list numbering definitions (if lists exist)
word/_rels/document.xml.rels — relationships (images, etc.)
word/media/             — embedded images
```

Key OOXML concepts:
- **Paragraph** (`<w:p>`) — block-level element with `<w:pPr>` (properties)
- **Run** (`<w:r>`) — inline text with `<w:rPr>` (formatting)
- **Table** (`<w:tbl>`) — used for layout tables, data tables, and box containers
- **Drawing** (`<w:drawing>`) — shapes, images
- Units: **twips** (1/20 of a point, 1440 per inch) and **EMU** (914400 per inch)

The `xml_writer.py` module provides a thin XML builder (`XmlWriter`) used throughout the DOCX writer.

## Code style

- No external dependencies (pure Python 3.10+)
- Dataclasses for all data structures
- Frozen dataclasses for immutable context objects
- Type hints everywhere
- Short methods — most are under 30 lines
- Dict dispatch over if/elif chains
- Convention over configuration — good defaults, everything customizable
