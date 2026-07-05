# python-pptx compatibility suite

Runs python-pptx's own behave feature suite against `pptx_rs`.

`features/` symlinks a subset of `../python-pptx/features/*.feature` plus its
`steps/` directory. Our `environment.py` (loaded by behave before step modules)
replaces `pptx.Presentation` with `pptx_rs.Presentation`, so every scenario
exercises the Rust implementation while python-pptx's enums/helpers stay
importable for the step definitions.

Requires the python-pptx repo checked out as a sibling:
`../python-pptx` (feature files and fixture .pptx files are read from there).

```bash
uv run behave tests/compat/features -f plain
```

## Status (2026-07-06)

| feature | pass | fail/error | total | notes |
|---|---|---|---|---|
| txt-text | 6 | 0 | 6 | run/paragraph text incl. `_xHHHH_` ctrl-char escapes |
| shp-shared | 35 | 37 | 72 | id/name/geometry pass; rotation, shadow, click_action, part not implemented |
| sld-slide | 4 | 23 | 27 | shapes/slide_id pass; layouts, notes_slide, background not implemented |
| prs-open-save | 3 | 3 | 6 | path/stream round-trip pass; dir-package, image part access not implemented |
| txt-textframe | 4 | 15 | 19 | text/paragraphs pass; auto_size, word_wrap, margins, font not implemented |
| txt-paragraph | 6 | 21 | 27 | text/runs/add_run pass; alignment, level, spacing, font not implemented |
| sld-slides | 3 | 11 | 14 | len/iter/index pass; add_slide, get(slide_id) not implemented |
| prs-presentation-props | 2 | 4 | 6 | slide_width/height pass; masters, notes_master, core_properties n/a |
| shp-shapes | 2 | 89 | 91 | add_textbox passes; add_chart/picture/table/shape, placeholders, freeform n/a |
| prs-default-template | 0 | 1 | 1 | needs `Presentation.slide_masters` |
| **TOTAL** | **65** | **204** | **269** | |

Every remaining failure is an unimplemented API (charts, pictures, tables,
placeholders, masters/layouts, formatting properties, shape-type proxy
dispatch), not a behavioral divergence in the implemented surface.
