"""behave environment for running python-pptx's own feature suite against pptx_rs.

Loaded by behave before any step module. python-pptx stays importable (its
enums, chart data, etc. are used by step modules), but its `Presentation`
entry point is replaced with pptx_rs's, so every scenario exercises the Rust
implementation.
"""

import os

import pptx

import pptx_rs

pptx.Presentation = pptx_rs.Presentation  # ty: ignore[invalid-assignment]

scratch_dir = os.path.abspath(os.path.join(os.path.split(__file__)[0], "_scratch"))


def before_all(context):
    if not os.path.isdir(scratch_dir):
        os.mkdir(scratch_dir)
