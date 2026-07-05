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

## Status (2026-07-06, after add_slide + layouts)

| feature | pass | fail/error | total | notes |
|---|---|---|---|---|
| txt-text | 6 | 0 | 6 | run/paragraph text incl. `_xHHHH_` ctrl-char escapes |
| prs-default-template | 1 | 0 | 1 | default template + slide_masters |
| shp-shared | 35 | 37 | 72 | id/name/geometry pass; rotation, shadow, click_action, part not implemented |
| sld-slide | 8 | 19 | 27 | shapes/slide_id/slide_layout pass; notes_slide, background not implemented |
| sld-slides | 7 | 7 | 14 | len/iter/index/add_slide pass; get(slide_id), clone-across-files not implemented |
| prs-open-save | 3 | 3 | 6 | path/stream round-trip pass; dir-package, image part access not implemented |
| prs-presentation-props | 3 | 3 | 6 | slide_width/height/masters pass; notes_master, core_properties n/a |
| txt-textframe | 4 | 15 | 19 | text/paragraphs pass; auto_size, word_wrap, margins, font not implemented |
| txt-paragraph | 6 | 21 | 27 | text/runs/add_run pass; alignment, level, spacing, font not implemented |
| shp-shapes | 2 | 89 | 91 | add_textbox passes; add_chart/picture/table/shape, placeholders, freeform n/a |
| **TOTAL** | **75** | **194** | **269** | |

Every remaining failure is an unimplemented API (charts, pictures, tables,
placeholders, masters/layouts, formatting properties, shape-type proxy
dispatch), not a behavioral divergence in the implemented surface.
