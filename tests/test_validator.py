"""Tests for the dok validator."""
import pytest
from dok.nodes import ElementNode, TextNode
from dok.validator import validate
from dok.errors import ValidationErrors


def el(name, props=None, children=None):
    return ElementNode(name=name, props=props or {}, children=children or [])


def text(t):
    return TextNode(text=t)


class TestStructure:
    def test_valid_simple_doc(self):
        # Should not raise
        validate([el("doc", children=[
            el("page", children=[el("p", children=[text("hello")])])
        ])])

    def test_li_outside_list(self):
        with pytest.raises(ValidationErrors):
            validate([el("li", children=[text("bad")])])

    def test_li_inside_ul(self):
        validate([el("ul", children=[el("li", children=[text("ok")])])])

    def test_tr_outside_table(self):
        with pytest.raises(ValidationErrors):
            validate([el("tr", children=[el("td", children=[text("bad")])])])

    def test_td_outside_tr(self):
        with pytest.raises(ValidationErrors):
            validate([el("td", children=[text("bad")])])

    def test_col_outside_cols(self):
        with pytest.raises(ValidationErrors):
            validate([el("col", children=[text("bad")])])


class TestProps:
    def test_unknown_prop(self):
        with pytest.raises(ValidationErrors):
            validate([el("bold", props={"nonexistent": "x"})])

    def test_valid_color_prop(self):
        validate([el("bold", props={"color": "red"}, children=[text("ok")])])

    def test_invalid_color(self):
        with pytest.raises(ValidationErrors):
            validate([el("bold", props={"color": "notacolor"}, children=[text("x")])])

    def test_valid_margin(self):
        validate([el("page", props={"margin": "narrow"})])

    def test_invalid_margin(self):
        with pytest.raises(ValidationErrors):
            validate([el("page", props={"margin": "huge"})])

    def test_valid_paper(self):
        validate([el("page", props={"paper": "a4"})])

    def test_invalid_paper(self):
        with pytest.raises(ValidationErrors):
            validate([el("page", props={"paper": "a5"})])


class TestRequiredProps:
    def test_img_requires_src(self):
        with pytest.raises(ValidationErrors):
            validate([el("img")])

    def test_link_requires_href(self):
        with pytest.raises(ValidationErrors):
            validate([el("link")])


class TestFontSize:
    def test_size_too_small(self):
        with pytest.raises(ValidationErrors):
            validate([el("size", props={"value": 2})])

    def test_size_too_large(self):
        with pytest.raises(ValidationErrors):
            validate([el("size", props={"value": 200})])

    def test_valid_size(self):
        validate([el("size", props={"value": 12}, children=[text("ok")])])


class TestMultipleErrors:
    def test_collects_all_errors(self):
        with pytest.raises(ValidationErrors) as exc_info:
            validate([
                el("li", children=[text("orphan li")]),
                el("td", children=[text("orphan td")]),
            ])
        assert len(exc_info.value.errors) >= 2
