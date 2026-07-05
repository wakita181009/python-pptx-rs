"""Utility value types, API-compatible with pptx.util."""

from __future__ import annotations


class Length(int):
    """Base class for length classes such as Emu, Inches, Pt.

    Behaves as an int count of English Metric Units (914,400 per inch,
    360,000 per centimeter).
    """

    _EMUS_PER_INCH = 914400
    _EMUS_PER_CENTIPOINT = 127
    _EMUS_PER_CM = 360000
    _EMUS_PER_MM = 36000
    _EMUS_PER_PT = 12700

    @property
    def inches(self) -> float:
        return self / self._EMUS_PER_INCH

    @property
    def centipoints(self) -> int:
        return self // self._EMUS_PER_CENTIPOINT

    @property
    def cm(self) -> float:
        return self / self._EMUS_PER_CM

    @property
    def emu(self) -> int:
        return self

    @property
    def mm(self) -> float:
        return self / self._EMUS_PER_MM

    @property
    def pt(self) -> float:
        return self / self._EMUS_PER_PT


class Inches(Length):
    """Convenience constructor for length in inches."""

    def __new__(cls, inches: float) -> Inches:
        return Length.__new__(cls, round(inches * cls._EMUS_PER_INCH))


class Centipoints(Length):
    """Convenience constructor for length in hundredths of a point."""

    def __new__(cls, centipoints: int) -> Centipoints:
        return Length.__new__(cls, round(centipoints * cls._EMUS_PER_CENTIPOINT))


class Cm(Length):
    """Convenience constructor for length in centimeters."""

    def __new__(cls, cm: float) -> Cm:
        return Length.__new__(cls, round(cm * cls._EMUS_PER_CM))


class Emu(Length):
    """Convenience constructor for length in English Metric Units."""

    def __new__(cls, emu: int) -> Emu:
        return Length.__new__(cls, int(emu))


class Mm(Length):
    """Convenience constructor for length in millimeters."""

    def __new__(cls, mm: float) -> Mm:
        return Length.__new__(cls, round(mm * cls._EMUS_PER_MM))


class Pt(Length):
    """Convenience constructor for length in points."""

    def __new__(cls, points: float) -> Pt:
        return Length.__new__(cls, round(points * cls._EMUS_PER_PT))
