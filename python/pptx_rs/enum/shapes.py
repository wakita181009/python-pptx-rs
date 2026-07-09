"""Enumerations used by shapes (``pptx.enum.shapes`` counterpart)."""

from __future__ import annotations

from pptx_rs._core import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE

# python-pptx exposes these aliases
MSO = MSO_SHAPE_TYPE
MSO_SHAPE = MSO_AUTO_SHAPE_TYPE

__all__ = ["MSO", "MSO_AUTO_SHAPE_TYPE", "MSO_SHAPE", "MSO_SHAPE_TYPE"]
