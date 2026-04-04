"""Tests for the dok.units module — measurement parsing and conversion."""
import pytest
from dok.units import (
    parse_to_twips, parse_to_emu, parse_to_half_points, parse_to_pt,
    pt_to_twips, pt_to_emu, twips_to_pt, twips_to_emu, emu_to_twips,
)


class TestParseToTwips:
    def test_plain_int_defaults_to_pt(self):
        assert parse_to_twips(12) == 240  # 12pt = 240 twips

    def test_plain_float(self):
        assert parse_to_twips(1.5) == 30  # 1.5pt = 30 twips

    def test_pt_suffix(self):
        assert parse_to_twips("12pt") == 240

    def test_inch(self):
        assert parse_to_twips("1in") == 1440

    def test_cm(self):
        # 2.54cm = 1 inch = 1440 twips
        assert parse_to_twips("2.54cm") == 1440

    def test_mm(self):
        # 25.4mm = 1 inch = 1440 twips
        assert parse_to_twips("25.4mm") == 1440

    def test_px(self):
        # 96px = 1 inch at 96 DPI = 1440 twips
        assert parse_to_twips("96px") == 1440

    def test_twip_identity(self):
        assert parse_to_twips("1440twip") == 1440

    def test_zero(self):
        assert parse_to_twips(0) == 0

    def test_string_int(self):
        assert parse_to_twips("72") == 1440  # 72pt

    def test_case_insensitive(self):
        assert parse_to_twips("12PT") == 240
        assert parse_to_twips("1IN") == 1440

    def test_negative(self):
        assert parse_to_twips(-10) == -200

    def test_invalid_raises(self):
        with pytest.raises(ValueError):
            parse_to_twips("abc")


class TestParseToEmu:
    def test_pt(self):
        assert parse_to_emu(72) == 914400  # 72pt = 1 inch

    def test_inch(self):
        assert parse_to_emu("1in") == 914400

    def test_cm(self):
        assert parse_to_emu("2.54cm") == 914400

    def test_px(self):
        # 1px = 9525 EMU
        assert parse_to_emu("1px") == 9525

    def test_zero(self):
        assert parse_to_emu(0) == 0


class TestParseToHalfPoints:
    def test_11pt(self):
        assert parse_to_half_points(11) == 22

    def test_14pt(self):
        assert parse_to_half_points("14pt") == 28

    def test_zero(self):
        assert parse_to_half_points(0) == 0


class TestParseToPt:
    def test_identity(self):
        assert parse_to_pt(12) == 12.0

    def test_inch(self):
        assert parse_to_pt("1in") == 72.0

    def test_cm(self):
        assert parse_to_pt("2.54cm") == 72.0


class TestDirectConversions:
    def test_pt_to_twips(self):
        assert pt_to_twips(1) == 20
        assert pt_to_twips(72) == 1440

    def test_pt_to_emu(self):
        assert pt_to_emu(72) == 914400

    def test_twips_to_pt(self):
        assert twips_to_pt(240) == 12.0

    def test_twips_to_emu(self):
        assert twips_to_emu(1440) == 914400

    def test_emu_to_twips(self):
        assert emu_to_twips(914400) == 1440

    def test_roundtrip_pt_twips(self):
        for pt in [6, 11, 12, 14, 24, 36, 72]:
            assert twips_to_pt(pt_to_twips(pt)) == pt

    def test_roundtrip_twips_emu(self):
        for tw in [20, 240, 720, 1440]:
            assert emu_to_twips(twips_to_emu(tw)) == tw
