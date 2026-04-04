"""Tests for the dok.sugar module — markdown-style syntax sugar."""
import pytest
from dok.sugar import desugar, desugar_inline


class TestHeadingSugar:
    def test_h1(self):
        assert desugar("# Hello") == 'h1 "Hello"'

    def test_h2(self):
        assert desugar("## Section") == 'h2 "Section"'

    def test_h3(self):
        assert desugar("### Subsection") == 'h3 "Subsection"'

    def test_h4(self):
        assert desugar("#### Minor") == 'h4 "Minor"'

    def test_heading_with_leading_spaces(self):
        assert desugar("  # Indented") == '  h1 "Indented"'

    def test_heading_preserves_trailing_text(self):
        assert desugar("# Hello World 123") == 'h1 "Hello World 123"'

    def test_heading_escapes_quotes(self):
        assert desugar('# Say "hello"') == 'h1 "Say \\"hello\\""'

    def test_no_heading_without_space(self):
        # #tag should not be treated as a heading
        assert desugar("#tag") == "#tag"


class TestQuoteSugar:
    def test_simple_quote(self):
        assert desugar("> This is a quote") == 'quote { "This is a quote" }'

    def test_quote_with_leading_spaces(self):
        assert desugar("  > Indented quote") == '  quote { "Indented quote" }'


class TestBulletListSugar:
    def test_single_bullet(self):
        result = desugar("- Item one")
        assert 'ul {' in result
        assert 'li { "Item one" }' in result

    def test_multiple_bullets(self):
        result = desugar("- One\n- Two\n- Three")
        assert result.count("li {") == 3

    def test_star_bullets(self):
        result = desugar("* Alpha\n* Beta")
        assert 'ul {' in result
        assert result.count("li {") == 2

    def test_does_not_match_page_break(self):
        result = desugar("---")
        assert "ul" not in result
        assert "---" in result


class TestNumberedListSugar:
    def test_simple_numbered(self):
        result = desugar("1. First\n2. Second")
        assert 'ol {' in result
        assert result.count("li {") == 2

    def test_non_sequential_numbers(self):
        result = desugar("1. A\n5. B\n10. C")
        assert result.count("li {") == 3


class TestMixedSugar:
    def test_heading_and_list(self):
        result = desugar("# Title\n\n- One\n- Two")
        assert 'h1 "Title"' in result
        assert 'ul {' in result

    def test_preserves_regular_dok(self):
        source = 'doc { p { "hello" } }'
        assert desugar(source) == source

    def test_preserves_triple_quotes(self):
        source = '"""\n# Not a heading\n"""'
        result = desugar(source)
        assert "# Not a heading" in result
        assert "h1" not in result

    def test_no_sugar_inside_braces(self):
        source = 'box {\n# Not a heading\n}'
        result = desugar(source)
        # Inside braces, the # should be left as-is
        assert "h1" not in result

    def test_comments_preserved(self):
        source = "// This is a comment\n# Title"
        result = desugar(source)
        assert "// This is a comment" in result
        assert 'h1 "Title"' in result


class TestInlineSugar:
    def test_bold(self):
        result = desugar_inline("Hello **world**")
        assert 'bold { "world" }' in result
        assert '"Hello "' in result

    def test_italic(self):
        result = desugar_inline("Hello *world*")
        assert 'italic { "world" }' in result

    def test_strike(self):
        result = desugar_inline("Hello ~~old~~")
        assert 'strike { "old" }' in result

    def test_underline(self):
        result = desugar_inline("Hello __important__")
        assert 'underline { "important" }' in result

    def test_link(self):
        result = desugar_inline("Visit [Google](https://google.com)")
        assert 'link(href: "https://google.com") { "Google" }' in result

    def test_code_inline(self):
        result = desugar_inline("Use `print()` here")
        assert 'code { "print()" }' in result

    def test_no_inline_sugar(self):
        result = desugar_inline("Plain text here")
        assert result == ""  # empty means no transformation

    def test_multiple_inline_styles(self):
        result = desugar_inline("Hello **bold** and *italic* text")
        assert 'bold { "bold" }' in result
        assert 'italic { "italic" }' in result
        assert '" and "' in result

    def test_bold_before_italic(self):
        # ** should match before *
        result = desugar_inline("**bold** then *italic*")
        assert 'bold { "bold" }' in result
        assert 'italic { "italic" }' in result


class TestInlineSugarIntegration:
    """Test inline sugar through the full pipeline."""

    def test_bold_in_paragraph(self):
        import dok
        node = dok.parse('p { "Hello **world**" }')
        p = node.children[0]
        # Should have expanded inline sugar
        assert len(p.children) > 0

    def test_link_in_paragraph(self):
        import dok
        node = dok.parse('p { "Visit [here](https://example.com) now" }')
        data = dok.to_bytes(node)
        assert len(data) > 0

    def test_heading_sugar_to_docx(self):
        import dok
        node = dok.parse("# My Document\n\n## Section One")
        data = dok.to_bytes(node)
        assert len(data) > 0

    def test_bullet_sugar_to_docx(self):
        import dok
        node = dok.parse("- Alpha\n- Beta\n- Gamma")
        data = dok.to_bytes(node)
        assert len(data) > 0
