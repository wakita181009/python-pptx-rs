# python-pptx compatibility suite

Runs python-pptx's own behave feature suite against `pptx_rs`.

`features/` vendors a subset of python-pptx's `features/*.feature` plus its
`steps/` directory (including `steps/test_files/` fixtures). Our
`environment.py` (loaded by behave before step modules) replaces
`pptx.Presentation` with `pptx_rs.Presentation`, so every scenario exercises
the Rust implementation while python-pptx's enums/helpers stay importable for
the step definitions (python-pptx itself is a dev dependency from PyPI).

```bash
uv run behave tests/compat/features -f plain
```

## Status (2026-07-06, after MarkItDown read APIs)

| feature | pass | fail/error | total | notes |
|---|---|---|---|---|
| txt-text | 6 | 0 | 6 | run/paragraph text incl. `_xHHHH_` ctrl-char escapes |
| prs-default-template | 1 | 0 | 1 | default template + slide_masters |
| shp-shared | 39 | 33 | 72 | id/name/geometry/shape_type pass; rotation, shadow, click_action, part not implemented |
| sld-slide | 11 | 16 | 27 | shapes/slide_id/slide_layout/notes_slide pass; background, adding notes not implemented |
| sld-slides | 7 | 7 | 14 | len/iter/index/add_slide pass; get(slide_id), clone-across-files not implemented |
| prs-open-save | 4 | 2 | 6 | path/stream round-trip + image part access pass; dir-package not implemented |
| prs-presentation-props | 3 | 3 | 6 | slide_width/height/masters pass; notes_master, core_properties n/a |
| txt-textframe | 4 | 15 | 19 | text/paragraphs pass; auto_size, word_wrap, margins, font not implemented |
| txt-paragraph | 6 | 21 | 27 | text/runs/add_run pass; alignment, level, spacing, font not implemented |
| shp-shapes | 3 | 88 | 91 | add_textbox, shapes.title pass; add_chart/picture/table/shape, freeform n/a |
| **TOTAL** | **84** | **185** | **269** | |

Every remaining failure is an unimplemented API (shape/placeholder proxy
class hierarchy, write-side chart/picture/table adds, masters/layouts,
formatting properties), not a behavioral divergence in the implemented
surface.

## MarkItDown integration

`tests/integrations/test_markitdown.py` runs [MarkItDown](https://github.com/microsoft/markitdown)
with `pptx.Presentation` monkeypatched to `pptx_rs.Presentation` and asserts
the produced Markdown is byte-identical to the real-python-pptx baseline, on
both MarkItDown's own test deck (vendored under `tests/fixtures/markitdown/`)
and a synthetic deck with speaker notes. Covered read surface: stream open,
`shapes.title`, `shape_type` (IntEnum, cross-library `==`), placeholder
geometry inheritance (slide → layout → master), tables, pictures
(`image.blob/content_type/filename`, `_element._nvXxPr.cNvPr.attrib` alt
text), group shapes, category charts, and `notes_slide`.
