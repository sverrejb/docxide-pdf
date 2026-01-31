# docxside-pdf

> ⚠️ **Work in progress.** The API, output quality, and supported features are all actively changing.

A Rust library for converting DOCX files to PDF, with the goal of matching Microsoft Word's PDF export as closely as possible.

## Goal

Given a `.docx` file, produce a `.pdf` that is visually indistinguishable from what Word would export. This is harder than it sounds — Word's layout engine handles fonts, spacing, line breaking, and page geometry in ways that are not fully documented.

## Current State

Basic text rendering works for simple documents. Font embedding is functional. Multi-paragraph layout with heading styles and spacing is supported. Complex content (tables, images, charts) is not yet handled.

Similarity scores against Word-generated reference PDFs:

| Case | Jaccard | SSIM (ink blocks) |
|------|---------|-------------------|
| case1 (simple body text) | ~44% | ~42% |
| case2 (headings + body) | ~27% | ~24% |
| case3 (multi-paragraph) | ~34% | ~29% |

Scores reflect font shape differences and layout imprecision. There is meaningful room for improvement.

## Showcase

> Run `python tools/showcase.py` to regenerate.

<!-- showcase-start -->
<table>
  <tr><th>MS Word</th><th>Docxside-PDF</th></tr>
  <tr>
    <td align="center"><img src="showcase/case2_ref.png"/><br/><sub>case2 — reference</sub></td>
    <td align="center"><img src="showcase/case2_gen.png"/><br/><sub>case2 — 93.3% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="showcase/case3_ref.png"/><br/><sub>case3 — reference</sub></td>
    <td align="center"><img src="showcase/case3_gen.png"/><br/><sub>case3 — 88.1% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="showcase/case1_ref.png"/><br/><sub>case1 — reference</sub></td>
    <td align="center"><img src="showcase/case1_gen.png"/><br/><sub>case1 — 61.2% SSIM</sub></td>
  </tr>
</table>
<!-- showcase-end -->

## Usage

```rust
use docxside_pdf::convert_docx_to_pdf;
use std::path::Path;

convert_docx_to_pdf(
    Path::new("input.docx"),
    Path::new("output.pdf"),
)?;
```

## Architecture

```
src/
  lib.rs      — public API
  error.rs    — Error enum
  model.rs    — Document/Paragraph/Run intermediate representation
  docx.rs     — DOCX ZIP + XML → Document parser
  pdf.rs      — Document → PDF renderer
tests/
  visual_comparison.rs  — Jaccard + SSIM comparison against Word reference PDFs
  fixtures/<case>/      — input.docx + reference.pdf pairs
  output/<case>/        — generated.pdf, screenshots, diff images
tools/
  docx-inspect          — inspect ZIP entries and XML inside a DOCX
  docx-fonts            — print font/style info from a DOCX
  jaccard               — compute Jaccard similarity between two PNGs or directories
  case-diff             — render and compare a fixture, print per-page scores
  graph.py              — live-updating similarity score graph over time
```

## Testing

Tests require [muto](https://mupdf.com/docs/mutool.html) (`mutool`) on `PATH` for PDF-to-PNG rendering.

```bash
# Run all tests
cargo test -- --nocapture

# Run only Jaccard visual comparison
cargo test visual_comparison -- --nocapture

# Run only SSIM comparison
cargo test ssim_comparison -- --nocapture
```

Each test prints a summary table at the end:

```
+-------+---------+------+
| Case  | Jaccard | Pass |
+-------+---------+------+
| case1 |   44.2% | ✓    |
| case2 |   27.0% | ✓    |
| case3 |   33.5% | ✓    |
+-------+---------+------+
  threshold: 25%
```

Results are appended to `tests/output/results.csv` and `tests/output/ssim_results.csv`. Run `python tools/graph.py` to see a live-updating graph of scores over time.

## Debugging Tools

Build the tools once:

```bash
cd tools && cargo build
```

Then run from the project root:

```bash
# Inspect XML inside a DOCX
cargo run --manifest-path tools/Cargo.toml --bin docx-inspect -- input.docx

# Print font information
cargo run --manifest-path tools/Cargo.toml --bin docx-fonts -- input.docx

# Compare two rendered pages
cargo run --manifest-path tools/Cargo.toml --bin jaccard -- a.png b.png

# Full fixture diff
cargo run --manifest-path tools/Cargo.toml --bin case-diff -- case1
```

## License

Not yet specified.
