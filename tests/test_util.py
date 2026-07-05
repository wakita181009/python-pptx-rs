"""Unit tests for pptx_rs.util length types (API-compatible with pptx.util)."""

from __future__ import annotations

import pytest

from pptx_rs.util import Centipoints, Cm, Emu, Inches, Length, Mm, Pt


class TestLengthConstructors:
    @pytest.mark.parametrize(
        ("value", "expected_emu"),
        [
            (Inches(1), 914400),
            (Inches(0.5), 457200),
            (Centipoints(100), 12700),
            (Cm(1), 360000),
            (Cm(2.54), 914400),
            (Emu(914400), 914400),
            (Mm(10), 360000),
            (Pt(1), 12700),
            (Pt(72), 914400),
        ],
    )
    def test_constructs_expected_emu_value(self, value: Length, expected_emu: int) -> None:
        assert value == expected_emu

    def test_length_is_int(self) -> None:
        assert isinstance(Inches(1), int)
        assert isinstance(Emu(1), Length)


class TestLengthProperties:
    def test_inches(self) -> None:
        assert Inches(2).inches == 2.0

    def test_centipoints(self) -> None:
        assert Pt(1).centipoints == 100

    def test_cm(self) -> None:
        assert Cm(3).cm == 3.0

    def test_emu(self) -> None:
        assert Inches(1).emu == 914400

    def test_mm(self) -> None:
        assert Mm(25).mm == 25.0

    def test_pt(self) -> None:
        assert Inches(1).pt == 72.0
