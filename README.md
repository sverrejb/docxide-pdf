# docxside-pdf

## ⚠️ Work in progress. This is in no way ready for use in production. The API, output quality, and supported features are all actively changing.

A Rust library and CLI tool for converting DOCX files to PDF, with the goal of matching Microsoft Word's PDF export as closely as possible.<sup>*</sup>

*<sub>Reference PDFs are generated using Microsoft Word for Mac (16.106.1) with the "Best for electronic distribution and accessibility (uses Microsoft online service)" export option.</sub>

## Goal

Given a `.docx` file, produce a `.pdf` that is visually indistinguishable from what Word would export. This is harder than it sounds — Word's layout engine handles fonts, spacing, line breaking, and page geometry in ways that are not fully documented.

## Agent disclaimer 🤖

While the idea, architecture, testing strategy and validation of output are all human, the vast majority of the code as of now is written by Claude Opus 4.6 with access to the PDF specification (ISO-32000) and the Office Open XML File Formats specification (ECMA-376).

## Sort-of supported features ✅

These *kind of* work:

- **Text**: font embedding (TTF/OTF), bold, italic, underline, strikethrough, font size, text color, theme fonts
- **Paragraphs**: left/center/right/justify alignment, space before/after, line spacing, indentation, contextual spacing, keep-next, bottom borders
- **Styles**: paragraph style inheritance (`basedOn` chains), document defaults from `docDefaults`
- **Lists**: bullet and numbered lists with nesting levels
- **Tables**: column widths with auto-fit, cell borders, cell text with alignment
- **Images**: inline JPEG embedding with sizing
- **Page layout**: page size, margins, document grid, automatic page breaking with widow/orphan control
- **Fonts**: cross-platform font search (macOS/Linux/Windows), embedded DOCX font extraction, `DOCXSIDE_FONTS` env var for custom font directories

### Not yet supported

Explicit page/section breaks, headers/footers, footnotes, tab stops, clickable hyperlinks, non-JPEG images, table merged cells, table cell shading, text boxes, charts, SmartArt, superscript/subscript, multi-column layouts, and many other features.

## Examples

See more examples in the [showcase](https://github.com/sverrejb/docxside-pdf/tree/main/showcase#readme)

<!-- showcase-start -->
<table>
  <tr><th>MS Word</th><th>Docxside-PDF</th></tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case4_ref.png"/><br/><sub>case4 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case4_gen.png"/><br/><sub>case4 — 89.5% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case6_ref.png"/><br/><sub>case6 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case6_gen.png"/><br/><sub>case6 — 54.8% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case7_ref.png"/><br/><sub>case7 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case7_gen.png"/><br/><sub>case7 — 91.7% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case8_ref.png"/><br/><sub>case8 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase/case8_gen.png"/><br/><sub>case8 — 94.1% SSIM</sub></td>
  </tr>
</table>
<!-- showcase-end -->

## Installation

```bash
# Install the CLI
cargo install docxside-pdf
```

## Usage

### CLI

```bash
# Convert a DOCX file to PDF
docxside-pdf input.docx

# Specify output path (defaults to input.pdf)
docxside-pdf input.docx output.pdf
```

### Library

```bash
cargo add docxside-pdf --no-default-features
```

This avoids pulling in the CLI dependency (`clap`).

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

Tests require [mutool](https://mupdf.com/docs/mutool.html) (`mutool`) on `PATH` for PDF-to-PNG rendering.

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

Apache-2.0
