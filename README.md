# python-pptx-rs

Rust-powered, [python-pptx](https://github.com/scanny/python-pptx)-compatible library for
creating, reading, and updating PowerPoint (.pptx) files.

- Drop-in compatible public API: `pip install python-pptx-rs`, then `import pptx_rs as pptx`
- Rust core (quick-xml + zip) via PyO3, single abi3 wheel for Python 3.10-3.14
- O(1) shape-id assignment and relationship lookup (quadratic in python-pptx)
- Untouched parts round-trip byte-exactly

## Benchmarks

1000 slides x 10 textboxes, macOS arm64, median of 3 (`benches/bench_vs_python_pptx.py`):

| workload | python-pptx | python-pptx-rs | speedup |
|---|---|---|---|
| open | 85.0ms | 8.9ms | 9.5x |
| open + traverse all text | 316.2ms | 73.5ms | 4.3x |
| traverse (pre-opened) | 227.8ms | 10.4ms | 21.9x |
| open + 1 edit + save | 199.3ms | 48.8ms | 4.1x |
| add 500 textboxes to one slide | 191.0ms | 10.8ms | 17.7x |

## Compatibility

python-pptx's own behave feature suite runs against python-pptx-rs — see
[tests/compat/README.md](tests/compat/README.md) for the current matrix
(75 scenarios passing; all remaining gaps are unimplemented APIs, not
behavioral divergences).

## Development

Requires Rust (stable), [uv](https://docs.astral.sh/uv/), and for the
compatibility suite a python-pptx checkout at `../python-pptx`.

```bash
uv sync
uv run maturin develop --uv
uv run pytest  # includes Python coverage (fails under 80%)
```

Rust coverage (merges `cargo test` with pytest run against an instrumented
build, since the PyO3 binding crate is only exercised through Python):

```bash
cargo install cargo-llvm-cov
source <(cargo llvm-cov show-env --export-prefix)
cargo llvm-cov clean --workspace
cargo test --workspace
uv run --no-sync maturin develop --uv
uv run --no-sync pytest --no-cov
cargo llvm-cov report --fail-under-lines 80
```
