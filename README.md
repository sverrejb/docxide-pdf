# docxide-pdf

## ⚠️ Work in progress. This is in no way ready for use in production. The API, output quality, and supported features are all actively changing.

A Rust library and CLI tool for converting DOCX files to PDF, with the goal of matching Microsoft Word's PDF export as closely as possible.<sup>*</sup>

*<sub>Reference PDFs are generated using Microsoft Word for Mac (16.106.1) with the "Best for electronic distribution and accessibility (uses Microsoft online service)" export option.</sub>

## Goals

**Accurate:** Given a `.docx` file, produce a `.pdf` that is visually identical to what Word would export.

**Fast:** Typical conversions complete in under 100ms.

**Small files:** Output PDFs should be the same size or smaller than Word's export.

## AI usage disclaimer 🤖

While the idea, architecture, testing strategy and validation of output are all human, the vast majority of the code as of now is written by Claude Opus 4.6 with access to the PDF specification (ISO-32000) and the Office Open XML File Formats specification (ECMA-376).

## Supported features

- **Text**: font embedding (TTF/OTF/TTC), bold, italic, underline, strikethrough, double strikethrough, font size, text color, superscript/subscript, small caps, all caps, character spacing, text expansion/compression (`w:w`), hidden text (`w:vanish`), kerning (legacy kern table + GPOS PairAdjustment)
- **Paragraphs**: left/center/right/justify alignment, space before/after, line spacing (auto, exact, at-least), first-line and hanging indentation, left/right indentation, contextual spacing, keep-next, keep-lines, paragraph borders (top/bottom/left/right/between) with color, paragraph shading, run highlighting
- **Styles**: paragraph and run style inheritance (`basedOn` chains), document defaults from `docDefaults`, theme fonts and colors
- **Lists**: bullet and numbered lists with multi-level nesting, custom number formats, list style inheritance
- **Tables**: column widths with auto-fit, merged cells (horizontal `gridSpan` and vertical `vMerge`), row heights (exact and minimum), per-cell borders with color/width, cell shading, vertical alignment, cell margins, floating/positioned tables (`tblpPr`)
- **Images**: inline JPEG/PNG embedding with sizing and alpha transparency, anchored/floating images (all wrap modes), floating image positioning relative to page/margin/column
- **Text boxes**: DrawingML textboxes (`wps:txbx`) and VML fallback (`v:textbox`), shape fills (solid color with theme color support including lumMod/lumOff), textbox body margins
- **Charts**: bar (clustered/stacked, vertical/horizontal), line, pie, area, doughnut, radar, scatter, bubble — with axis labels, gridlines, legends
- **Page layout**: page size, margins, document grid (`linePitch`), explicit page breaks, `pageBreakBefore`, automatic page breaking with widow/orphan control
- **Sections**: multiple sections with `nextPage`/`continuous`/`oddPage`/`evenPage` breaks, per-section page size and margins
- **Multi-column layout**: 2+ columns with custom widths and spacing, column breaks, column separators
- **Headers/footers**: default, first-page, and even/odd variants, per-section headers/footers, STYLEREF field resolution, page number and page count fields, images in headers/footers
- **Footnotes**: footnote references, footnote rendering at page bottom with separator line
- **Fields**: PAGE, NUMPAGES, STYLEREF (with spec-compliant search order), field code cached results for non-dynamic fields
- **Hyperlinks**: clickable links in PDF output (URI link annotations)
- **Tab stops**: left, center, right, decimal with leader dots
- **Track changes**: final mode (insertions included, deletions removed — matches Word's PDF export)
- **Compatibility**: `mc:AlternateContent` fallback, structured document tag (`w:sdt`) content extraction, `altChunk` HTML content parsing
- **Fonts**: cross-platform font search (macOS/Linux/Windows), embedded DOCX font extraction and deobfuscation, font subsetting (CIDFont/Type0), disk-cached font index, font substitution via `fontTable.xml` altName and family-class fallback
- **Output optimization**: font subsetting, content stream compression

### Not yet supported

- **Text**: ligatures, complex script shaping (Arabic, Devanagari, etc.), CJK fallback fonts
- **Tables**: conditional formatting (`tblLook`/`tblStylePr` — banded rows, first/last column styles), nested tables, text direction in cells (`textDirection`)
- **Images**: text wrapping around floating images/textboxes, EMF/WMF vector images
- **Layout**: distribute alignment (`w:jc val="distribute"`), vertical page alignment (`w:vAlign` on section), right-to-left (bidi) text
- **PDF features**: bookmarks/outline, document metadata (title, author)
- **Features**: table of contents generation, endnotes, SmartArt, OLE objects
- **Fonts**: bundled fallback fonts, text shaping (ligatures, complex scripts)

## Examples

See more examples in the [showcase](https://github.com/sverrejb/docxide-pdf/tree/main/showcase#readme)

<!-- showcase-start -->
<table>
  <tr><th>MS Word</th><th>Docxside-PDF</th></tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case4_ref.png"/><br/><sub>case4 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case4_gen.png"/><br/><sub>case4 — 89.5% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case6_ref.png"/><br/><sub>case6 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case6_gen.png"/><br/><sub>case6 — 54.8% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case7_ref.png"/><br/><sub>case7 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case7_gen.png"/><br/><sub>case7 — 91.7% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case8_ref.png"/><br/><sub>case8 — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case8_gen.png"/><br/><sub>case8 — 94.1% SSIM</sub></td>
  </tr>
</table>
<!-- showcase-end -->

## Installation

```bash
# Install the CLI
cargo install docxide-pdf
```

## Usage

### CLI

```bash
# Convert a DOCX file to PDF
docxide-pdf input.docx

# Specify output path (defaults to input.pdf)
docxide-pdf input.docx output.pdf
```

### Library

```bash
cargo add docxide-pdf --no-default-features
```

This avoids pulling in the CLI dependency (`clap`).

```rust
use docxide_pdf::convert_docx_to_pdf;
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
- **macOS**: `~/Library/Caches/docxide-pdf/font-index.tsv`
- **Linux**: `$XDG_CACHE_HOME/docxide-pdf/font-index.tsv` (default `~/.cache/`)
- **Windows**: `%LOCALAPPDATA%\docxide-pdf\cache\font-index.tsv`

## Architecture

```
src/
  lib.rs              — public API
  main.rs             — CLI binary (behind `cli` feature)
  error.rs            — Error enum
  model.rs            — Document/Section/Paragraph/Run intermediate representation
  fonts/
    mod.rs            — font registration, metrics, fallback
    discovery.rs      — cross-platform font search, disk cache
    embed.rs          — font embedding, kern/GPOS extraction, subsetting
    encoding.rs       — WinAnsi encoding, glyph mapping
    cache.rs          — font index disk cache (TSV)
  docx/
    mod.rs            — DOCX ZIP + XML → Document parser, shared utilities
    styles.rs         — theme, style parsing, style inheritance
    runs.rs           — run-level XML → Vec<Run>
    numbering.rs      — list/numbering parsing, counter management
    images.rs         — image extraction (inline + floating), chart detection
    charts.rs         — chart XML parsing (8 chart types)
    textbox.rs        — textbox parsing (DrawingML + VML)
    tables.rs         — table + cell parsing
    embedded_fonts.rs — DOCX font extraction and deobfuscation
    sections.rs       — section properties, columns, headers/footers refs
    headers_footers.rs— header/footer/footnote XML parsing
    settings.rs       — document settings (tab stops, mirror margins)
    alt_chunk.rs      — altChunk HTML content parsing
  pdf/
    mod.rs            — main render loop, header/footer rendering
    layout.rs         — text layout, line building, paragraph rendering
    table.rs          — table layout, auto-fit, table rendering
    charts.rs         — cartesian chart rendering (bar/line/area/scatter/bubble/radar)
    charts_radial.rs  — pie/doughnut chart rendering
    chart_legend.rs   — shared legend rendering
    header_footer.rs  — header/footer rendering
    footnotes.rs      — footnote rendering
tests/
  visual_comparison.rs  — Jaccard + SSIM comparison against Word reference PDFs
  text_boundary.rs      — page/line-level text boundary tests
  fixtures/<case>/      — input.docx + reference.pdf pairs
  output/<case>/        — generated.pdf, screenshots, diff images
tools/
  analyze-fixtures      — fixture score table, feature audit, XML grep
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
