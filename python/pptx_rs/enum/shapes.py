"""Enumerations used by shapes (``pptx.enum.shapes`` counterpart)."""

from __future__ import annotations

from pptx_rs._core import MSO_SHAPE_TYPE

# python-pptx exposes this alias
MSO = MSO_SHAPE_TYPE

__all__ = ["MSO", "MSO_SHAPE_TYPE"]
