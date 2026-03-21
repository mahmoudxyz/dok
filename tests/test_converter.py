"""Tests for the dok converter (AST → DocxModel)."""
import pytest
from dok.nodes import ElementNode, TextNode
from dok.converter import (
    Converter, ParagraphModel, RunModel, DocxModel,
    LineModel, BoxModel, BannerModel, BadgeModel,
    DataTableModel, SpacerModel, HeaderModel, FooterModel,
)


def el(name, props=None, children=None):
    return ElementNode(name=name, props=props or {}, children=children or [])

def text(t):
    return TextNode(text=t)


def convert(source_nodes):
    return Converter().convert(source_nodes)


class TestBasicConversion:
    def test_empty_doc(self):
        model = convert([el("doc")])
        assert isinstance(model, DocxModel)

    def test_paragraph(self):
        model = convert([el("doc", children=[
            el("p", children=[text("hello")])
        ])])
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert len(paras) == 1
        assert paras[0].runs[0].text == "hello"

    def test_heading(self):
        model = convert([el("doc", children=[
            el("h1", children=[text("Title")])
        ])])
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].style == "Heading1"

    def test_bold_run(self):
        model = convert([el("doc", children=[
            el("p", children=[el("bold", children=[text("strong")])])
        ])])
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].bold is True

    def test_italic_run(self):
        model = convert([el("doc", children=[
            el("p", children=[el("italic", children=[text("em")])])
        ])])
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].runs[0].italic is True


class TestAlignment:
    def test_center(self):
        model = convert([el("doc", children=[
            el("center", children=[el("p", children=[text("centered")])])
        ])])
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].align == "center"

    def test_right(self):
        model = convert([el("doc", children=[
            el("right", children=[el("p", children=[text("right")])])
        ])])
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert paras[0].align == "right"


class TestShapes:
    def test_line(self):
        model = convert([el("doc", children=[el("line")])])
        lines = [c for c in model.content if isinstance(c, LineModel)]
        assert len(lines) == 1

    def test_box(self):
        model = convert([el("doc", children=[
            el("box", props={"fill": "lightblue"}, children=[text("content")])
        ])])
        boxes = [c for c in model.content if isinstance(c, BoxModel)]
        assert len(boxes) == 1

    def test_banner(self):
        model = convert([el("doc", children=[
            el("banner", props={"fill": "navy"}, children=[text("title")])
        ])])
        banners = [c for c in model.content if isinstance(c, BannerModel)]
        assert len(banners) == 1

    def test_badge(self):
        model = convert([el("doc", children=[
            el("badge", props={"fill": "green"}, children=[text("OK")])
        ])])
        badges = [c for c in model.content if isinstance(c, BadgeModel)]
        assert len(badges) == 1
        assert badges[0].text == "OK"


class TestLists:
    def test_unordered_list(self):
        model = convert([el("doc", children=[
            el("ul", children=[
                el("li", children=[text("item 1")]),
                el("li", children=[text("item 2")]),
            ])
        ])])
        assert model.has_lists is True
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert len(paras) >= 2

    def test_ordered_list(self):
        model = convert([el("doc", children=[
            el("ol", children=[
                el("li", children=[text("first")]),
                el("li", children=[text("second")]),
            ])
        ])])
        assert model.has_lists is True


class TestTables:
    def test_basic_table(self):
        model = convert([el("doc", children=[
            el("table", children=[
                el("tr", children=[
                    el("td", children=[text("A")]),
                    el("td", children=[text("B")]),
                ]),
            ])
        ])])
        tables = [c for c in model.content if isinstance(c, DataTableModel)]
        assert len(tables) == 1
        assert len(tables[0].rows) == 1


class TestHyperlinks:
    def test_link_inline(self):
        model = convert([el("doc", children=[
            el("p", children=[
                el("link", props={"href": "https://example.com"},
                   children=[text("click")])
            ])
        ])])
        paras = [c for c in model.content if isinstance(c, ParagraphModel)]
        assert any(r.hyperlink_url for r in paras[0].runs)


class TestSpacer:
    def test_spacer(self):
        model = convert([el("doc", children=[
            el("space", props={"size": 24})
        ])])
        spacers = [c for c in model.content if isinstance(c, SpacerModel)]
        assert len(spacers) == 1
        assert spacers[0].height_twips > 0


class TestHeaderFooter:
    def test_header(self):
        model = convert([el("doc", children=[
            el("header", children=[el("p", children=[text("Header")])])
        ])])
        assert model.header is not None

    def test_footer(self):
        model = convert([el("doc", children=[
            el("footer", children=[el("p", children=[text("Footer")])])
        ])])
        assert model.footer is not None


class TestDocDefaults:
    def test_default_font(self):
        model = convert([el("doc", props={"font": "Arial", "size": 14})])
        assert model.default_font == "Arial"
        assert model.default_size_pt == 14
