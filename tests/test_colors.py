"""Tests for the dok color resolver."""
from dok.colors import resolve


class TestColors:
    def test_named_color(self):
        assert resolve("red") == "FF0000"

    def test_hex_6(self):
        assert resolve("#FF0000") == "FF0000"

    def test_hex_3(self):
        assert resolve("#F00") == "FF0000"

    def test_unknown(self):
        assert resolve("notacolor") is None

    def test_case_insensitive(self):
        assert resolve("Red") == "FF0000"
        assert resolve("RED") == "FF0000"
