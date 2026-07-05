# python-pptx-rs

Rust + PyO3 reimplementation of [python-pptx](https://github.com/scanny/python-pptx)
that keeps its **public API** while replacing the engine. The Python object layer is
python-pptx's real bottleneck (its XML parsing is already C via lxml), so proxies are
implemented as Rust pyclasses — one attribute access = one Rust call.

## Commands

```bash
uv sync                              # venv + dev deps
uv run maturin develop --uv          # build & install extension (debug)
uv run maturin develop --uv --release  # REQUIRED before running benchmarks
uv run pytest                        # own test suite (coverage gate 80%)
uv run behave tests/compat/features -f plain  # python-pptx compat suite
uv run ruff check python tests benches && uv run ty check python tests benches
cargo test -p pptx-core && cargo clippy --workspace
uv run python benches/bench_vs_python_pptx.py
```

- `cargo build -p pptx-python` fails at link on macOS — that is expected for
  `extension-module`; use `cargo check` for iteration and maturin to build.
- Always run tools from the repo root: maturin needs the `.venv` here, and a stale
  shell cwd (e.g. `crates/pptx-python`) makes uv/maturin resolve the wrong project.

## Architecture

- `crates/pptx-core` — no Python deps. `dom.rs`: arena XML DOM (quick-xml) with
  namespace resolution. `opc.rs`: zip package, relationships, content types.
- `crates/pptx-python` — all pyclasses (`_core` module). Every proxy =
  `Py<Presentation>` + part name + `NodeId`; all state lives in `Presentation`,
  accessed via its pyclass borrow (never hold a borrow across a Python call).
- `python/pptx_rs` — thin shim only: `Presentation()` entry point, `util` Length
  types, `_core.pyi` stubs (update stubs whenever the Rust API changes).

## Invariants (breaking these breaks round-trip or compat)

- **NodeIds are forever**: the DOM arena is append-only; "removing" a node only
  detaches it from its parent. Never compact or reuse arena slots — live Python
  proxies hold NodeIds.
- **Byte-exact round-trip**: parts never touched are written back verbatim. Use
  `pkg.doc()` for reads and `pkg.doc_mut()` only for writes — `doc_mut()` marks the
  part dirty and switches save to re-serialization for that part. A read path that
  calls `doc_mut()` silently destroys byte fidelity (there is a pytest guarding this).
- **O(1) id allocation**: `next_shape_ids` caches per-part max shape id; any new code
  path that inserts shapes must go through `take_next_shape_id` (or update the cache).
  Slide ids start at 256 (python-pptx convention).
- **python-pptx text semantics** (verified by the compat suite): control chars
  invalid in XML are stored as `_xHHHH_` escapes; `_Paragraph.text` setter converts
  both `\n` and `\v` to `a:br`; getters render `a:br` as `\v` and join paragraphs
  with `\n`; `a:endParaRPr` must stay the last child of `a:p`.
- **add_slide**: clones layout placeholders *except* date/footer/slide-number
  ("latent" types), builds fresh minimal `p:sp` elements (no deep copy), and must
  register all three: slide→layout rel, presentation→slide rel, `[Content_Types].xml`
  Override.

## Compatibility workflow

- `../python-pptx` checkout is the reference. Before implementing any API, read its
  source for exact semantics (e.g. `oxml/` templates, `shapes/shapetree.py`) — do not
  guess from the docs.
- `tests/compat/` runs python-pptx's own behave features against us by monkeypatching
  `pptx.Presentation = pptx_rs.Presentation` in `environment.py` (loaded before step
  modules). Step files under `tests/compat/features/steps/` are **vendored
  third-party code** — excluded from ruff/ty, never "fix" them.
- After adding an API: re-run behave, update the matrix in `tests/compat/README.md`.
  A behave *failure* in implemented surface is a bug; an AttributeError in
  unimplemented surface is roadmap.
- Every write-path pytest re-opens the output with real python-pptx to validate.

## Conventions

- Latest-version policy for all deps: query registries (`cargo add`, PyPI) — never
  trust remembered version numbers. Exception: avoid pre-releases (e.g. zip 9.0.0-preX;
  stay on latest stable).
- PyO3 uses `abi3-py310`: one wheel covers Python 3.10–3.14. Don't add
  version-specific FFI.
- Rust edition 2024 (let-chains OK). No `unwrap()` outside tests/invariants;
  map core errors via `to_py`.
- GitHub Actions: `astral-sh/setup-uv` has no `v8` major alias tag — pin exact
  versions (`v8.3.0`).
