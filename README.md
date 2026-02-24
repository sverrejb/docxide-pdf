# docxside-pdf

## ⚠️ Work in progress. This is in no way ready for use in production. The API, output quality, and supported features are all actively changing.

A Rust library and CLI tool for converting DOCX files to PDF, with the goal of matching Microsoft Word's PDF export as closely as possible.<sup>*</sup>

*<sub>Reference PDFs are generated using Microsoft Word for Mac (16.106.1) with the "Best for electronic distribution and accessibility (uses Microsoft online service)" export option.</sub>

## Goals

**Accurate:** Given a `.docx` file, produce a `.pdf` that is visually identical to what Word would export. This is harder than it sounds — Word's layout engine handles fonts, spacing, line breaking, and page geometry in ways that are not fully documented.

**Fast:** Typical conversions complete in under 100ms.

**Small files:** Output PDFs should be the same size or smaller than Word's exports. Font subsetting and content stream compression keep file sizes down.

## AI useage disclaimer 🤖

While the idea, architecture, testing strategy and validation of output are all human, the vast majority of the code as of now is written by Claude Opus 4.6 with access to the PDF specification (ISO-32000) and the Office Open XML File Formats specification (ECMA-376).

## Sort-of supported features ✅

These *kind of* work:

- **Text**: font embedding (TTF/OTF), bold, italic, underline, strikethrough, font size, text color, superscript/subscript, theme fonts
- **Paragraphs**: left/center/right/justify alignment, space before/after, line spacing, indentation, contextual spacing, keep-next, bottom borders
- **Styles**: paragraph style inheritance (`basedOn` chains), document defaults from `docDefaults`
- **Lists**: bullet and numbered lists with nesting levels
- **Tables**: column widths with auto-fit, merged cells (horizontal and vertical), row heights, cell borders with color/width, cell shading, vertical alignment, cell text with alignment
- **Images**: inline JPEG and PNG embedding with sizing and alpha transparency
- **Page layout**: page size, margins, document grid, explicit page breaks, automatic page breaking with widow/orphan control
- **Headers/footers**: default and first-page variants, page number and page count fields
- **Tab stops**: left, center, right, decimal with leader dots
- **Fonts**: cross-platform font search (macOS/Linux/Windows), embedded DOCX font extraction, font subsetting, disk-cached font index
- **Output optimization**: font subsetting, content stream compression

### Not yet supported

Footnotes, clickable hyperlinks, text boxes, charts, SmartArt, multi-column layouts, section breaks with different page sizes/orientations, and many other features.

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

## Configuration

### Environment Variables

| Variable | Description |
|---|---|
| `DOCXSIDE_FONTS` | Additional font directories to search, colon-separated (`;` on Windows). Searched before system font directories. |
| `DOCXSIDE_NO_FONT_CACHE` | Set to any value to disable the font index disk cache. Forces a full font scan on every conversion. Useful for debugging font resolution issues. |

Font scanning results are cached to disk (per-directory, invalidated by mtime). The cache is stored at:
- **macOS**: `~/Library/Caches/docxside-pdf/font-index.tsv`
- **Linux**: `$XDG_CACHE_HOME/docxside-pdf/font-index.tsv` (default `~/.cache/`)
- **Windows**: `%LOCALAPPDATA%\docxside-pdf\cache\font-index.tsv`

## Architecture

```
src/
  lib.rs          — public API
  error.rs        — Error enum
  model.rs        — Document/Paragraph/Run intermediate representation
  fonts.rs        — font discovery, metrics, subsetting
  docx/
    mod.rs        — DOCX ZIP + XML → Document parser
    styles.rs     — theme, style parsing, style inheritance
  pdf/
    mod.rs        — main render loop, header/footer rendering
    layout.rs     — text layout, line building, paragraph rendering
    table.rs      — table layout, auto-fit, table rendering
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

Tests require `mutool` on `PATH` for PDF-to-PNG rendering:

```bash
brew install mupdf        # macOS
apt install mupdf-tools   # Debian/Ubuntu
```

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

## Contributing

Pull requests are welcome!

### Got a weird DOCX?

If you have a `.docx` file that produces ugly, broken, or just plain wrong output, send it to me! Real-world documents with surprising formatting are the best way to improve the converter. Open an issue or PR with the file included and I will try to make it work.

## License

Apache-2.0
