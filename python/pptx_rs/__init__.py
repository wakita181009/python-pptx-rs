"""python-pptx-rs: Rust-powered, python-pptx-compatible .pptx library."""

from __future__ import annotations

import os
from importlib.resources import files
from typing import IO

from pptx_rs import _core

__version__ = _core.__version__

__all__ = ["Presentation", "__version__"]


def Presentation(pptx: str | IO[bytes] | os.PathLike[str] | None = None) -> _core.Presentation:
    """Return a |Presentation| object loaded from `pptx`.

    `pptx` is a path, a path-like object, or a binary file-like object. When missing
    or None, the built-in default presentation template is loaded (matching
    python-pptx behavior).
    """
    if pptx is None:
        data = (files("pptx_rs") / "templates" / "default.pptx").read_bytes()
        return _core.Presentation.from_bytes(data)
    if isinstance(pptx, (str, os.PathLike)):
        return _core.Presentation.from_path(os.fspath(pptx))
    return _core.Presentation.from_bytes(pptx.read())
