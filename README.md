# Dok

> A document language built on two ideas:
> great design comes from great limitations,
> and everything is composable by default.

---

## The core idea

A document is a stack of contexts.

```
doc(font: Calibri, size: 11) {          ← sets document defaults
  page(margin: normal) {                ← sets physical space
    center {                            ← sets alignment
      bold(color: navy) {               ← sets run style
        h1 { "Quarterly Report" }       ← content atom
      }
    }
  }
}
```

Each layer adds one thing to the context. Content at the bottom inherits everything above it.
No cascade. No specificity. No global state. Just a tree where context flows inward.

This is the whole language. Every feature is a node at one of five layers:

```
doc    →  document defaults   (font, size, language)
page   →  physical space      (margin, paper, columns)
layout →  arrangement         (center, right, rtl, row, cols, indent)
style  →  appearance          (bold, italic, color, size, font)
content → atoms               (h1 h2 h3 h4, p, box, line, "text")
```

---

## What Dok deliberately does not do

Dok has no variables. No loops. No conditionals. No expressions.

This is not a limitation to work around — it is the design.

Loops and conditionals live in your programming language, which already does them better
than any template engine could. Dok's job is to describe *what the document looks like*.
Your language's job is to decide *what goes into it*.

**Dok handles:** structure, appearance, shapes, layout, direction, composition.

**Your language handles:** which sections to include, how many items in a list,
what color a status badge should be, whether to show a warning block.

The boundary is clean. The converter is simple. The output is predictable.

---

## Syntax

One rule for every node in the language:

```
name(props) { children }    ← node with props and children
name { children }            ← node with no props
name(props) { "text" }      ← node with text content
"text"                       ← bare text node
->                           ← arrow connector (inside row only)
// comment                   ← line comment, stripped before parsing
```

Props are always `key: value` pairs or bare flags (boolean true):

```
fill: navy          ← key: named color
fill: #4472C4       ← key: hex color
fill: none          ← key: keyword
size: 14            ← key: number
rounded             ← bare flag (equivalent to rounded: true)
shadow              ← bare flag
```

That is the complete grammar. A parser for this is under 200 lines in any language.
There are no operators, no expressions, no special cases.

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
| `lang` | `en` `ar` `he` `fr` ... | `en` |
| `direction` | `ltr` `rtl` | `ltr` |

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

| Prop | Values | Meaning |
|------|--------|---------|
| `margin` | `normal` `narrow` `wide` `none` | Page margins preset |
| `margin` | `normal-top` `none-sides` | Directional override |
| `paper` | `a4` `letter` `a3` | Paper size |
| `cols` | `2` `3` | Column count (newspaper-style) |

`page` is optional. If omitted, `margin: normal` and `paper: a4` apply.

---

### Layer 3 — layout

Arrangement nodes. They affect how their children are placed on the page.
They do not produce any visible element themselves — only their children do.

```
center { h1 { "Title" } }
right  { p  { "Page 1" } }
rtl    { p  { "نص عربي" } }

indent(level: 2) {
  p { "Indented two levels." }
}

row {
  box(fill: blue)   { "Step 1" }  ->
  box(fill: orange) { "Step 2" }  ->
  box(fill: green)  { "Step 3" }
}

cols {
  col { p { "Left column content." }  }
  col { p { "Right column content." } }
}

cols(ratio: 2:1) {
  col { p { "Wider column." }  }
  col { p { "Narrower column." } }
}
```

**Layout nodes:**

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
| `float(right)` | Floats a shape to the right, text wraps around it |

**`row` and the `->` connector:**

Inside a `row`, `->` between two children draws an arrow connecting them.
`-> "label" ->` draws a labeled arrow.

```
row {
  box(fill: blue)   { "Input"   }
  -> "data" ->
  box(fill: orange) { "Process" }
  ->
  box(fill: green)  { "Output"  }
}
```

---

### Layer 4 — style

Style wrappers modify how text renders inside them. They can nest freely.
`bold` inside `color(red)` inside `center` all compose correctly.

```
p {
  "Normal text, then "
  bold { "bold text" }
  ", then "
  italic(color: navy) { "italic navy text" }
  "."
}
```

**Style nodes:**

| Node | DOCX | Notes |
|------|------|-------|
| `bold` | `<w:b/>` | |
| `italic` | `<w:i/>` | |
| `underline` | `<w:u val="single"/>` | |
| `strike` | `<w:strike/>` | |
| `sup` | `<w:vertAlign val="superscript"/>` | |
| `sub` | `<w:vertAlign val="subscript"/>` | |
| `color(red)` | `<w:color/>` | Named or #hex |
| `size(14)` | `<w:sz val="28"/>` | Points |
| `font(Georgia)` | `<w:rFonts/>` | |
| `highlight(yellow)` | `<w:highlight/>` | Named colors only |

Style nodes can carry props directly instead of as children:

```
// These are equivalent:
bold { color(red) { "hello" } }
bold(color: red) { "hello" }
```

When a style node has a prop that belongs to a different layer — like `color` on a paragraph —
it applies to the layer it belongs to:

```
// color on a paragraph-level node sets the default text color for that paragraph
p(color: gray) { "All text in this paragraph is gray by default." }

// bold wrapping a p sets bold on all runs inside
bold { p { "This whole paragraph is bold." } }
```

---

### Layer 5 — content

Content nodes are the atoms. They produce visible output.

**Text blocks:**

| Node | DOCX | Notes |
|------|------|-------|
| `h1` to `h4` | `<w:pStyle val="Heading1"/>` | Headings |
| `p` | `<w:p>` | Paragraph |
| `quote` | `<w:pStyle val="BlockText"/>` | Block quote |
| `code` | `<w:pStyle val="SourceCode"/>` | Code block (monospace) |
| `"text"` | `<w:t>` | Bare text node, always inside a block |
| `---` | `<w:pageBreak/>` | Page break |

**Shape atoms:**

| Node | Shape | DOCX preset |
|------|-------|-------------|
| `box` | Rectangle | `prstGeom prst="rect"` |
| `box(rounded)` | Rounded rectangle | `prstGeom prst="roundRect"` |
| `circle` | Circle / ellipse | `prstGeom prst="ellipse"` |
| `diamond` | Diamond | `prstGeom prst="diamond"` |
| `chevron` | Chevron | `prstGeom prst="chevron"` |
| `callout` | Speech callout | `prstGeom prst="wedgeRectCallout"` |
| `badge` | Small inline rect | inline `<w:drawing>` |
| `banner` | Full-width block | full-width `<w:drawing>` |
| `line` | Horizontal rule | `prstGeom prst="line"` |

**Shape props:**

| Prop | Values | Meaning |
|------|--------|---------|
| `fill` | color or `none` | Background fill |
| `stroke` | color or `none` | Border color |
| `stroke` | `dashed` `dotted` `thick` `thin` | Border style / width |
| `color` | color | Text color inside the shape |
| `rounded` | flag | Rounded corners |
| `shadow` | flag | Drop shadow |
| `accent` | color | Colored left-edge bar (banner only) |
| `tail` | `left` `right` `bottom-left` | Callout tail direction |

Text inside a shape is just children:

```
box(fill: navy, color: white, rounded) {
  bold { "Important" }
  p { "This content sits inside the box." }
}
```

**Color values:**

Named: `red orange yellow green blue navy purple gray black white gold silver`
Light variants: `lightblue lightgreen lightyellow lightgray lightpink`
Hex: `#4472C4` `#1F3864` `#FF0000`
Transparent: `none`

---

## Complete example

```
doc(font: Calibri, size: 11) {

  page(margin: normal) {

    // Decorative header banner
    banner(fill: #1F3864, accent: gold) {
      bold(color: white, size: 16) { "Acme Corporation" }
    }

    // Centered title block
    center {
      h1 { "Q4 2024 Financial Report" }
      italic(color: gray) { p { "For internal distribution only" } }
    }

    // Body
    h2 { "Executive Summary" }

    p {
      "Total revenue reached "
      bold(color: #0070C0) { "$4.2 million" }
      ", representing a "
      bold(color: green) { "42% increase" }
      " year-over-year."
    }

    // Warning callout
    callout(fill: #FFF2CC, stroke: #FFC000, tail: bottom-left) {
      bold { "Note:" }
      p { "These figures are preliminary and subject to audit." }
    }

    h2 { "Process Overview" }

    // Flow diagram
    center {
      row {
        box(fill: #4472C4, color: white, rounded) { "Collect" }
        ->
        box(fill: #ED7D31, color: white, rounded) { "Analyse" }
        ->
        box(fill: #70AD47, color: white, rounded) { "Report"  }
      }
    }

    h2 { "Regional Breakdown" }

    // Two-column layout
    cols(ratio: 2:1) {
      col {
        p { "EMEA showed strongest growth, driven by new enterprise contracts
             in Germany and the UK. Headcount increased by 12 across the region." }
      }
      col {
        box(fill: lightblue, rounded) {
          center { bold(size: 18) { "+38%" } }
          center { p { "EMEA growth" } }
        }
      }
    }

    // Arabic section
    rtl {
      h2 { "الملخص التنفيذي" }
      p { "حقق إجمالي الإيرادات " bold { "4.2 مليون دولار" } "." }
    }

  }

}
```

---

## Builder API

The builder API is for when content is dynamic — lists of items, conditional sections,
data-driven shapes. Your language handles the logic. Dok handles the shape.

Every builder function returns a `Node`. Nodes are plain data — name, props, children.
They are immutable, composable, and have no side effects.

### Python

```python
import dok

# Static document
doc = dok.doc(
    dok.page(
        dok.banner("Acme Corp", fill="#1F3864", accent="gold", color="white"),
        dok.h1("Report"),
        dok.p("Summary here."),
        margin="normal",
    )
)

dok.to_docx(doc, "report.docx")
```

**Loops — just Python:**
```python
region_nodes = [
    dok.box(
        dok.bold(region.name),
        dok.p(region.summary),
        fill="lightblue", rounded=True,
    )
    for region in regions
]

doc = dok.doc(
    dok.page(
        dok.h1("Regional Performance"),
        *region_nodes,
    )
)
```

**Conditionals — just Python:**
```python
def status_badge(status):
    color = {"active": "green", "pending": "orange", "closed": "gray"}.get(status, "gray")
    return dok.badge(status.upper(), fill=color, color="white")

doc = dok.doc(
    dok.page(
        # None children are silently ignored
        dok.callout("CONFIDENTIAL", fill="red", color="white") if is_confidential else None,
        dok.h1("Report"),
        *[dok.p(dok.bold(item.name), " — ", status_badge(item.status))
          for item in items],
    )
)
```

**Composable helpers — just functions:**
```python
def metric_card(label, value, up=None):
    trend = dok.span("↑", color="green") if up is True  else \
            dok.span("↓", color="red")   if up is False else None
    return dok.box(
        dok.p(label, color="gray"),
        dok.bold(value, size=20),
        trend,
        fill="white", stroke="lightgray", rounded=True,
    )

dashboard = dok.row(
    metric_card("Revenue",   "$4.2M", up=True),
    metric_card("Customers", "1,840", up=True),
    metric_card("Churn",     "2.1%",  up=False),
)
```

### JavaScript / TypeScript

```typescript
import * as dok from 'dok'

// Loops — just JavaScript
const regionNodes = regions.map(region =>
    dok.box(
        dok.bold(region.name),
        dok.p(region.summary),
        { fill: "lightblue", rounded: true }
    )
)

// Conditionals — just JavaScript
const doc = dok.doc(
    dok.page(
        isConfidential && dok.callout("CONFIDENTIAL", { fill: "red", color: "white" }),
        dok.h1("Report"),
        ...regionNodes,
        { margin: "normal" }
    )
)

// false / null / undefined children are silently ignored
await dok.toDocx(doc, "report.docx")
```

**Composable helpers:**
```typescript
const statusBadge = (status: string) => {
    const colors: Record<string, string> = {
        active: "green", pending: "orange", closed: "gray"
    }
    return dok.badge(status.toUpperCase(), {
        fill: colors[status] ?? "gray",
        color: "white",
    })
}

const section = (title: string, items: Item[]) => [
    dok.h2(title),
    ...items.map(item =>
        dok.p(dok.bold(item.name), " — ", statusBadge(item.status))
    ),
]
```

### Kotlin

Kotlin's trailing lambda syntax makes Dok read like Compose:

```kotlin
val doc = doc(font = "Calibri", size = 11) {
    page(margin = "normal") {

        banner(fill = "#1F3864", accent = "gold") {
            bold(color = "white", size = 16) { +"Acme Corp" }
        }

        center {
            h1 { +"Q4 Report" }
        }

        // Loops — just Kotlin
        for (region in regions) {
            box(fill = "lightblue", rounded = true) {
                bold { +region.name }
                p   { +region.summary }
            }
        }

        // Conditionals — just Kotlin
        if (isConfidential) {
            callout(fill = "red", color = "white") { +"CONFIDENTIAL" }
        }
    }
}

doc.toDocx("report.docx")
```

### Go

```go
doc := dok.Doc(
    dok.Page(
        dok.Banner("Acme Corp",
            dok.Fill("#1F3864"),
            dok.Accent("gold"),
            dok.Color("white"),
        ),
        dok.H1("Q4 Report"),
        buildRegions(regions)...,  // your function returning []dok.Node
        dok.Margin("normal"),
    ),
    dok.Font("Calibri"),
    dok.Size(11),
)

dok.ToDocx(doc, "report.docx")
```

---

## Mixing string syntax and builder API

String syntax is for static content. Builder API is for dynamic content. Mix freely.
Both produce the same internal tree — the converter sees no difference.

```python
import dok

# Load static templates once
header = dok.parse("""
  banner(fill: #1F3864, accent: gold) {
    bold(color: white, size: 16) { "Acme Corporation" }
  }
""")

footer = dok.parse("""
  center {
    italic(color: gray) { p { "Confidential — Internal Use Only" } }
  }
""")

# Build dynamic body in Python
body = [
    dok.h1(report.title),
    dok.p(report.summary),
    *[dok.box(dok.bold(s.title), dok.p(s.body), fill="lightblue")
      for s in report.sections
      if s.visible],
]

# Compose
doc = dok.doc(dok.page(header, *body, footer))
dok.to_docx(doc, "report.docx")
```

---

## Architecture — how to implement Dok

### The node tree

Both string parsing and the builder API produce the same structure:

```
Node:
  name:     str            // "box", "h1", "bold", "center", "doc", ...
  props:    dict[str, any] // { "fill": "navy", "rounded": True }
  children: list[Node]     // other nodes

TextNode:
  text: str

ArrowNode:
  label: str | None        // None for plain ->
```

`None` / `false` / `undefined` in children lists are silently dropped during construction.
This enables conditional children without if-else guards at every callsite.

### Parsing the string syntax

The grammar is LL(1). No ambiguity, no backtracking:

```
document  = node*
node      = NAME props? block
          | STRING                    // bare text node
          | '->' STRING? '->'         // arrow with optional label
          | '->'                      // plain arrow

props     = '(' prop (',' prop)* ')'
prop      = NAME ':' value
          | NAME                      // bare flag → true

block     = '{' node* '}'
value     = NAME | STRING | NUMBER | '#' HEX+
```

Tokeniser produces: `NAME STRING NUMBER COLOR LPAREN RPAREN LBRACE RBRACE COLON COMMA ARROW`

### The converter — context inheritance

The converter carries two small context structs as it walks the tree.
Wrapper nodes update context for their subtree. Leaf nodes consume context.

```
ParaCtx {
  align:     left | center | right | justify
  direction: ltr | rtl
  indent:    int
  spacing:   { before, after }
}

RunCtx {
  bold:      bool
  italic:    bool
  underline: bool
  strike:    bool
  color:     str | None
  highlight: str | None
  size:      int | None
  font:      str | None
}

convert(node, para_ctx, run_ctx):
  match node.name:
    "doc":      set_doc_defaults(node.props);  recurse children
    "page":     begin_section(node.props);     recurse children;  end_section()
    "center":   recurse with para_ctx.align = center
    "right":    recurse with para_ctx.align = right
    "justify":  recurse with para_ctx.align = justify
    "rtl":      recurse with para_ctx.direction = rtl
    "indent":   recurse with para_ctx.indent += level
    "bold":     recurse with run_ctx.bold = true
    "italic":   recurse with run_ctx.italic = true
    "color(x)": recurse with run_ctx.color = x
    "size(n)":  recurse with run_ctx.size = n
    "p":        emit_paragraph(node.children, para_ctx, run_ctx)
    "h1"–"h4":  emit_heading(level, node.children, para_ctx, run_ctx)
    "quote":    emit_paragraph(node.children, para_ctx.as_quote(), run_ctx)
    "code":     emit_paragraph(node.children, para_ctx.as_code(), run_ctx)
    "box":      emit_shape("rect",     node.props, node.children)
    "circle":   emit_shape("ellipse",  node.props, node.children)
    "diamond":  emit_shape("diamond",  node.props, node.children)
    "chevron":  emit_shape("chevron",  node.props, node.children)
    "banner":   emit_banner(node.props, node.children)
    "badge":    emit_inline_shape(node.props, node.children)
    "callout":  emit_shape("wedgeRectCallout", node.props, node.children)
    "line":     emit_line(node.props)
    "row":      emit_row(node.children)
    "cols":     emit_cols(node.children, node.props)
    TextNode:   emit_run(node.text, run_ctx)
    "---":      emit_page_break()
```

Key properties of this design:
- Context flows downward only. Children cannot affect parents.
- Each node type has exactly one converter function.
- No cascade resolution. No specificity calculation.
- The converter is stateless except for the two context structs.
- Time complexity: O(n) where n is number of nodes.

### DOCX output

The converter writes directly to a ZIP archive containing:

```
[Content_Types].xml   ← fixed boilerplate
_rels/.rels           ← fixed boilerplate
word/document.xml     ← THE OUTPUT — generated from node tree
word/styles.xml       ← fixed set of named styles (Heading1–4, SourceCode, BlockText)
word/settings.xml     ← fixed boilerplate
```

Adjacent text runs with identical `RunCtx` are merged into one `<w:r>` before emission.
The output never has one `<w:r>` per word.

---

## DOCX mapping — complete

| Dok | DOCX construct |
|-----|----------------|
| `doc(font: X)` | `<w:docDefaults>` in `styles.xml` |
| `page(margin: normal)` | `<w:sectPr>` with standard margins |
| `page(cols: 2)` | `<w:sectPr><w:cols w:num="2"/>` |
| `center { }` | `<w:pPr><w:jc w:val="center"/>` |
| `right { }` | `<w:pPr><w:jc w:val="right"/>` |
| `justify { }` | `<w:pPr><w:jc w:val="both"/>` |
| `rtl { }` | `<w:pPr><w:bidi/>` + `<w:pPr><w:rPr><w:rtl/>` |
| `indent(level: 2)` | `<w:pPr><w:ind w:left="1440"/>` |
| `bold` | `<w:rPr><w:b/>` |
| `italic` | `<w:rPr><w:i/>` |
| `underline` | `<w:rPr><w:u w:val="single"/>` |
| `strike` | `<w:rPr><w:strike/>` |
| `color(red)` | `<w:rPr><w:color w:val="FF0000"/>` |
| `size(14)` | `<w:rPr><w:sz w:val="28"/>` |
| `font(Georgia)` | `<w:rPr><w:rFonts w:ascii="Georgia"/>` |
| `sup` | `<w:rPr><w:vertAlign w:val="superscript"/>` |
| `sub` | `<w:rPr><w:vertAlign w:val="subscript"/>` |
| `h1` | `<w:pPr><w:pStyle w:val="Heading1"/>` |
| `p` | `<w:p>` |
| `quote` | `<w:pPr><w:pStyle w:val="BlockText"/>` |
| `code` | `<w:pPr><w:pStyle w:val="SourceCode"/>` |
| `---` | `<w:p><w:r><w:br w:type="page"/>` |
| `box` | `<w:drawing><wps:wsp><a:prstGeom prst="rect"/>` |
| `box(rounded)` | `<a:prstGeom prst="roundRect"/>` |
| `circle` | `<a:prstGeom prst="ellipse"/>` |
| `diamond` | `<a:prstGeom prst="diamond"/>` |
| `chevron` | `<a:prstGeom prst="chevron"/>` |
| `callout` | `<a:prstGeom prst="wedgeRectCallout"/>` |
| `fill: navy` | `<a:solidFill><a:srgbClr val="000080"/>` |
| `fill: none` | `<a:noFill/>` |
| `stroke: dashed` | `<a:ln><a:prstDash val="dash"/>` |
| `shadow` | `<a:effectLst><a:outerShdw/>` |
| `banner` | Full-width inline `<w:drawing>` rect + optional accent rect |
| `badge` | Inline `<w:drawing>` small rect |
| `line` | `<w:drawing>` line shape, full width |
| `row { A -> B }` | `<w:drawing>` group (`wpg:wgp`) + connector shapes |
| `cols { col col }` | `<w:tbl>` single-row borderless table |
| `float(right)` | `<wp:anchor wrapSquare wrapText="bothSides"/>` |
| Text inside shape | `<wps:txbx><w:txbxContent><w:p><w:r><w:t>` |

---

## File extension

`.dok` for string syntax files. UTF-8 text. Version-control friendly.
Diffs cleanly. Readable by non-developers.

For purely dynamic documents generated in code, no `.dok` file is needed —
build the node tree directly with the builder API and pass it to the converter.
EOF