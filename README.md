# Dok

> A document language built on two ideas:
> great design comes from great limitations,
> and everything is composable by default.

---

## The core idea

A document is a stack of contexts.

```
doc(font: Calibri, size: 11) {          // sets document defaults
  page(margin: normal) {                // sets physical space
    center {                            // sets alignment
      bold(color: navy) {               // sets run style
        h1 { "Quarterly Report" }       // content atom
      }
    }
  }
}
```

Each layer adds one thing to the context. Content at the bottom inherits everything above it.
No cascade. No specificity. No global state. Just a tree where context flows inward.

This is the whole language. Every feature is a node at one of five layers:

```
doc    →  document defaults   (font, size)
page   →  physical space      (margin, paper, columns)
layout →  arrangement         (center, right, rtl, row, cols, indent)
style  →  appearance          (bold, italic, color, size, font)
content → atoms               (h1-h4, p, box, line, table, ul/ol, img, link, "text")
```

---

## What Dok deliberately does not do

Dok has no variables. No loops. No conditionals. No expressions.

This is not a limitation to work around — it is the design.

Loops and conditionals live in your programming language, which already does them better
than any template engine could. Dok's job is to describe *what the document looks like*.
Your language's job is to decide *what goes into it*.

**Dok handles:** structure, appearance, shapes, layout, direction, composition, functions.

**Your language handles:** which sections to include, how many items in a list,
what color a status badge should be, whether to show a warning block.

The boundary is clean. The converter is simple. The output is predictable.

---

## Installation

```bash
pip install dok
```

## Quick start

**From the command line:**
```bash
python -m dok report.dok report.docx
python -m dok --check report.dok      # validate only
python -m dok --tree report.dok       # print node tree
```

**From Python:**
```python
import dok

node = dok.parse('doc { h1 { "Hello World" } }')
dok.to_docx(node, "hello.docx")
```

---

## Syntax

One rule for every node in the language:

```
name(props) { children }    // node with props and children
name { children }           // node with no props
name(props)                 // self-closing node
"text"                      // bare text node
->                          // arrow connector (inside row only)
---                         // page break
// comment                  // line comment, stripped before parsing
```

Props are always `key: value` pairs or bare flags (boolean true):

```
fill: navy          // key: named color
fill: #4472C4       // key: hex color
size: 14            // key: number
src: "image.png"    // key: string
rounded             // bare flag (equivalent to rounded: true)
```

That is the complete grammar. There are no operators, no expressions, no special cases.

---

## The five layers

### Layer 1 — doc

Sets document-wide defaults. Everything inside inherits these unless overridden.

```
doc(font: Calibri, size: 11) {
  // all content here
}
```

| Prop | Values | Default |
|------|--------|---------|
| `font` | any system font name | `Calibri` |
| `size` | number in points | `11` |

`doc` is optional. If omitted, all defaults apply.

---

### Layer 2 — page

Sets the physical page. Multiple `page` nodes = multiple sections with different layouts.

```
doc {
  page(margin: normal) {
    // standard content
  }

  page(margin: wide, cols: 2) {
    // two-column section
  }
}
```

| Prop | Values | Default |
|------|--------|---------|
| `margin` | `normal` `narrow` `wide` `none` | `normal` |
| `paper` | `a4` `letter` `a3` | `a4` |
| `cols` | `1` `2` `3` | `1` |

`page` is optional. If omitted, `margin: normal` and `paper: a4` apply.

---

### Layer 3 — layout

Arrangement nodes. They affect how their children are placed on the page.

```
center { h1 { "Title" } }
right  { p  { "Page 1" } }
rtl    { p  { "نص عربي" } }

indent(level: 2) {
  p { "Indented two levels." }
}

row {
  box(fill: blue, color: white)   { "Step 1" }  ->
  box(fill: orange, color: white) { "Step 2" }  ->
  box(fill: green, color: white)  { "Step 3" }
}

cols(ratio: 2:1) {
  col { p { "Wider column." }  }
  col { p { "Narrower column." } }
}
```

| Node | What it does |
|------|-------------|
| `center` | Centers children horizontally |
| `right` | Right-aligns children |
| `left` | Left-aligns children (default, rarely needed) |
| `justify` | Justifies text in children |
| `rtl` | Sets right-to-left direction for all children |
| `ltr` | Sets left-to-right direction (override inside rtl) |
| `indent(level: N)` | Indents children N levels (each level = 0.5 inch) |
| `row { A -> B }` | Places children side by side with optional arrow connectors |
| `cols { col col }` | Splits page into columns |
| `float(side: right)` | Floats content to the right, text wraps around it |

---

### Layer 4 — style

Style wrappers modify how text renders inside them. They nest freely.

```
p {
  "Normal text, then "
  bold { "bold text" }
  ", then "
  italic(color: navy) { "italic navy text" }
  "."
}
```

| Node | Effect |
|------|--------|
| `bold` | Bold text |
| `italic` | Italic text |
| `underline` | Underlined text |
| `strike` | Strikethrough text |
| `sup` | Superscript |
| `sub` | Subscript |
| `color(value: red)` | Text color (named or hex) |
| `size(value: 14)` | Font size in points |
| `font(value: Georgia)` | Font family |
| `highlight(value: yellow)` | Highlight color |
| `span(bold, color: red, size: 14)` | Multiple styles at once |

Style nodes can carry inline props:

```
// These are equivalent:
bold { color(value: red) { "hello" } }
bold(color: red) { "hello" }
```

---

### Layer 5 — content

Content nodes are the atoms. They produce visible output.

**Text blocks:**

| Node | Description |
|------|-------------|
| `h1` to `h4` | Headings |
| `p` | Paragraph |
| `quote` | Block quote |
| `code` | Code block (monospace) |
| `"text"` | Bare text node |
| `---` | Page break |

**Lists:**

```
ul {
  li { "First bullet" }
  li { "Second with " bold { "formatting" } }
  li { "Third item" }
}

ol {
  li { "Step one" }
  li { "Step two" }
  li { "Step three" }
}
```

Lists produce native Word bullet/numbered lists with proper indentation.

**Tables:**

```
table(border: true, striped: true) {
  tr {
    th { "Name" }
    th { "Score" }
  }
  tr {
    td { "Alice" }
    td { bold(color: green) { "95" } }
  }
  tr {
    td { "Bob" }
    td { bold(color: blue) { "88" } }
  }
}
```

| Prop | Values | Default |
|------|--------|---------|
| `border` | boolean | `false` |
| `striped` | boolean | `false` |

`th` cells are bold with a shaded background. `td(colspan: 2)` spans columns.

**Images:**

```
img(src: "photo.png", width: 4)
img(src: "logo.jpg", width: 2, height: 1)
```

| Prop | Description |
|------|-------------|
| `src` | Image file path (required) |
| `width` | Width in inches |
| `height` | Height in inches (auto-calculated from aspect ratio if omitted) |
| `alt` | Alt text |

Supports PNG and JPEG. Images are embedded in the docx file.

**Hyperlinks:**

```
p {
  "Visit "
  link(href: "https://example.com") { "our website" }
  " for more info."
}
```

Links are styled with blue underlined text automatically.

**Page numbers:**

```
footer {
  center {
    p {
      "Page "
      page-number
    }
  }
}
```

`page-number` inserts a live Word field that updates automatically.

**Visual elements:**

| Node | Description |
|------|-------------|
| `box` | Bordered/shaded content block |
| `callout` | Side-bordered note block |
| `banner` | Full-width colored block (great for page headers) |
| `badge` | Small inline label |
| `line` | Horizontal divider |

These use native Word formatting (tables, borders, shading) for clean printing.

```
box(fill: #E8F4FD, stroke: #2196F3, rounded: true) {
  bold { "Info Box" }
  p { "Content inside the box." }
}

callout(fill: #FFF3CD, stroke: #FFC107, tail: top-left) {
  bold { "Warning:" }
  " Check your input data."
}

banner(fill: #1F3864, accent: gold, color: white) {
  center { bold(size: 16) { "Company Name" } }
}

p { "Status: " badge(fill: green, color: white) { "ACTIVE" } }

line(stroke: gray, dashed: true)
```

**Shape props:**

| Prop | Values | Description |
|------|--------|-------------|
| `fill` | color | Background color |
| `stroke` | color | Border color |
| `color` | color | Text color inside |
| `rounded` | flag | Rounded corners (box only) |
| `shadow` | flag | Drop shadow |
| `accent` | color | Left-edge accent bar (banner only) |
| `tail` | `top-left` `top-right` `bottom-left` `bottom-right` | Callout pointer |
| `dashed` | flag | Dashed line style |
| `thick` | flag | Thick line |

**Spacer:**

```
space(size: 24)    // 24pt vertical space
```

**Header and footer:**

```
doc {
  header {
    right { italic(color: gray, size: 9) { "Document Title" } }
  }

  footer {
    center {
      p {
        color(value: gray, size: 9) { "Page " }
        page-number
      }
    }
  }

  page { /* content */ }
}
```

Headers and footers appear on every page. Place them inside `doc` or `page`.

**Color values:**

Named: `red orange yellow green blue navy purple gray black white gold silver`
Light variants: `lightblue lightgreen lightyellow lightgray lightpink`
Hex: `#4472C4` `#1F3864` `#FF0000` `#ABC`
Transparent: `none`

---

## Functions and reuse

Define reusable components with `def`:

```
def alert-box(message) {
  box(fill: #FFF3CD, stroke: #FFC107, rounded: true) {
    bold { "Warning: " }
    message
  }
}

def info-row(label, value) {
  p {
    bold(color: gray) { label }
    value
  }
}

// Use them
alert-box(message: "Check your data before submitting.")
info-row(label: "Name: ", value: "Alice Johnson")
```

Functions support:
- **Parameters** — referenced by name in the function body
- **Children** — use the `children` keyword for composable wrappers:

```
def card(title) {
  box(fill: #F5F5F5, rounded: true) {
    bold(size: 14) { title }
    children
  }
}

card(title: "Summary") {
  p { "This content goes where 'children' is." }
  p { "Multiple children work too." }
}
```

---

## Imports

Split large documents into reusable modules:

```
// components.dok
def company-header(title) {
  banner(fill: #1F3864, accent: gold, color: white) {
    center { bold(size: 16) { title } }
  }
}

def signature-block(name) {
  center {
    line
    color(value: gray) { name }
  }
}
```

```
// report.dok
import "components.dok"

doc {
  page {
    company-header(title: "Annual Report")
    h1 { "Introduction" }
    p { "..." }
    signature-block(name: "CEO")
  }
}
```

Imports are resolved relative to the importing file's directory. Circular imports are detected and rejected.

---

## Builder API (Python)

The builder API is for when content is dynamic — lists of items, conditional sections,
data-driven content. Your language handles the logic. Dok handles the shape.

Every builder function returns a `Node`. Nodes are plain data — name, props, children.
They are immutable, composable, and have no side effects.

```python
import dok

doc = dok.doc(
    dok.page(
        dok.banner("Acme Corp", fill="#1F3864", accent="gold", color="white"),
        dok.h1("Report"),
        dok.p("Revenue grew by ", dok.bold("42%", color="green"), "."),
        margin="normal",
    )
)

dok.to_docx(doc, "report.docx")
```

**All builder functions:**

```python
# Document structure
dok.doc(*children, font=..., size=...)
dok.page(*children, margin=..., paper=..., cols=...)

# Layout
dok.center(*children)
dok.right(*children)
dok.justify(*children)
dok.rtl(*children)
dok.ltr(*children)
dok.indent(*children, level=1)
dok.row(*children)
dok.cols(*children, ratio="1:1")
dok.col(*children)
dok.float_right(*children)
dok.float_left(*children)

# Style
dok.bold(*children, color=..., size=..., font=...)
dok.italic(*children, ...)
dok.underline(*children, ...)
dok.strike(*children, ...)
dok.sup(*children)
dok.sub(*children)
dok.color("red", *children)
dok.size(14, *children)
dok.font("Georgia", *children)
dok.highlight("yellow", *children)
dok.span(*children, bold=True, color=..., size=...)

# Text blocks
dok.h1("Title", ...)    # also h2, h3, h4
dok.p(*children, ...)
dok.quote(*children)
dok.code("source code")

# Lists
dok.ul(dok.li("item"), dok.li("item"))
dok.ol(dok.li("first"), dok.li("second"), start=1)
dok.li(*children)

# Tables
dok.table(
    dok.tr(dok.th("Header"), dok.th("Header")),
    dok.tr(dok.td("Cell"), dok.td("Cell")),
    border=True, striped=True,
)

# Visual elements
dok.box(*children, fill=..., stroke=..., rounded=True)
dok.callout(*children, fill=..., stroke=..., tail="top-left")
dok.banner(*children, fill=..., accent=..., color=...)
dok.badge("label", fill=..., color=...)
dok.line(stroke=..., dashed=True)

# Inline elements
dok.img("photo.png", width=4)
dok.link("https://example.com", "click here")
dok.page_number()

# Meta
dok.header(*children)
dok.footer(*children)
dok.space(size=12)
dok.page_break()
dok.arrow(label=None)
```

**Dynamic content — just Python:**

```python
# Loops
items = [dok.li(item.name) for item in data]
doc = dok.doc(dok.ul(*items))

# Conditionals — None children are silently dropped
doc = dok.doc(
    dok.page(
        dok.callout("CONFIDENTIAL", fill="red", color="white") if classified else None,
        dok.h1(report.title),
        *[dok.p(section.text) for section in report.sections],
    )
)
```

**Output formats:**

```python
# Write to file
dok.to_docx(node, "report.docx")

# Get bytes (for HTTP responses, S3, etc.)
data = dok.to_bytes(node)

# Parse .dok string
node = dok.parse(source_string)
node = dok.parse(source_string, base_dir=Path("./templates"))  # for imports
```

---

## Pipeline

The compilation pipeline:

```
source → lex → parse → resolve_imports → resolve_functions → validate → convert → write
```

1. **Lexer** — tokenizes the source string
2. **Parser** — builds an AST of Node trees
3. **Import resolver** — reads imported files, injects their nodes
4. **Function resolver** — expands function calls by substituting parameters
5. **Validator** — checks structure rules, prop types, printable constraints
6. **Converter** — walks the AST with context inheritance, produces a DocxModel
7. **Writer** — serializes the DocxModel to OOXML inside a ZIP archive

All errors carry source locations (line, column) and human-readable hints.

---

## Error handling

All errors include the source location and a hint for fixing the problem:

```
ParseError at line 5, col 12: Expected '}' to close block
  Hint: Every '{' needs a matching '}'.

ValidationError at line 8, col 3: 'li' must be inside 'ul' or 'ol'
  Hint: ul { li { "item" } }

ValidationError at line 12, col 5: Invalid color 'notacolor' for 'bold.color'
  Hint: Use a named color (red, navy, gold, ...) or hex (#FF0000, #ABC).

ResolveError at line 3: Missing parameter 'name' in call to 'greeting'
  Hint: Usage: greeting(name: ...)
```

The validator collects all errors in a single pass — you see everything that's wrong at once, not one error at a time.

---

## Complete example

```
import "components.dok"

def metric(label, value, trend_color) {
  box(fill: #F8F9FA, rounded: true) {
    p { color(value: gray) { label } }
    bold(size: 20, color: trend_color) { value }
  }
}

doc(font: Calibri, size: 11) {

  header {
    right { italic(color: gray, size: 9) { "Q4 2024 Report" } }
  }

  footer {
    center { p { color(value: gray, size: 9) { "Page " } page-number } }
  }

  page(margin: normal) {

    banner(fill: #1F3864, accent: gold, color: white) {
      center { bold(size: 16) { "Acme Corporation" } }
    }

    center {
      h1 { "Q4 2024 Financial Report" }
      italic(color: gray) { "For internal distribution only" }
    }

    // Key metrics
    row {
      metric(label: "Revenue", value: "$4.2M", trend_color: green)
      metric(label: "Customers", value: "1,840", trend_color: green)
      metric(label: "Churn", value: "2.1%", trend_color: red)
    }

    h2 { "Regional Breakdown" }

    table(border: true, striped: true) {
      tr {
        th { "Region" }
        th { "Revenue" }
        th { "Growth" }
      }
      tr {
        td { "EMEA" }
        td { "$1.8M" }
        td { bold(color: green) { "+38%" } }
      }
      tr {
        td { "Americas" }
        td { "$1.5M" }
        td { bold(color: green) { "+22%" } }
      }
      tr {
        td { "APAC" }
        td { "$0.9M" }
        td { bold(color: orange) { "+8%" } }
      }
    }

    space(size: 12)

    callout(fill: #FFF2CC, stroke: #FFC000, tail: bottom-left) {
      bold { "Note:" }
      p { "These figures are preliminary and subject to audit." }
    }

    h2 { "Key Highlights" }

    ul {
      li { "Enterprise contracts in Germany and UK drove EMEA growth" }
      li { "Customer acquisition cost decreased by " bold { "15%" } }
      li { "New product line launched in Q3 reaching " bold { "$400K" } " revenue" }
    }

    space(size: 24)

    cols(ratio: 1:1:1) {
      col { center { line  color(value: gray) { "CEO" } } }
      col { center { line  color(value: gray) { "CFO" } } }
      col { center { line  color(value: gray) { "Board Secretary" } } }
    }

  }
}
```

---

## File extension

`.dok` for string syntax files. UTF-8 text. Version-control friendly.
Diffs cleanly. Readable by non-developers.

For purely dynamic documents generated in code, no `.dok` file is needed —
build the node tree directly with the builder API and pass it to the converter.
