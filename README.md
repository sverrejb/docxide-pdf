# docxide-pdf

Library and CLI for converting DOCX files to PDF, matching Microsoft Word's output as closely as possible.

## ⚠️ Work in progress. This is in no way ready for use in production. The API, output quality, and supported features are all actively changing.

### Got a weird DOCX?

If you have a `.docx` file that produces ugly, broken, or just plain wrong output, send it to me! Real-world documents with surprising formatting are the best way to improve the converter. Open an issue or PR with the file included and I will try to make it work.


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
- **Styles**: paragraph and run style inheritance (`basedOn` chains), document defaults from `docDefaults` (all run properties: bold, italic, caps, smallCaps, vanish, strikethrough, dstrike, underline, color, char_spacing), theme fonts and colors
- **Lists**: bullet and numbered lists with multi-level nesting, custom number formats, list style inheritance
- **Tables**: column widths with auto-fit, merged cells (horizontal `gridSpan` and vertical `vMerge`), row heights (exact and minimum), per-cell borders with color/width, inline `w:tblBorders`, cell shading, vertical alignment, cell margins, floating/positioned tables (`tblpPr`)
- **Images**: inline JPEG/PNG embedding with sizing and alpha transparency, grayscale and CMYK JPEG support, anchored/floating images (all wrap modes), floating image positioning relative to page/margin/column, behind-document z-ordering
- **Text boxes**: DrawingML textboxes (`wps:txbx`) and VML fallback (`v:textbox`), shape fills (solid color with theme color support including lumMod/lumOff, linear gradients with multiple color stops), textbox body margins
- **Shapes & geometry**: all 187 OOXML preset shapes via formula-based geometry engine (guide formulas, adjustment values), custom geometry paths (`a:custGeom` with moveTo, lineTo, cubicBezTo, arcTo), shape fills and strokes
- **Charts**: bar (clustered/stacked, vertical/horizontal), line, pie, area, doughnut, radar, scatter, bubble — with axis labels, tick marks, gridlines, legends, bubble fill opacity
- **Page layout**: page size, margins, document grid (`linePitch`), explicit page breaks, `pageBreakBefore`, automatic page breaking with widow/orphan control
- **Sections**: multiple sections with `nextPage`/`continuous`/`oddPage`/`evenPage` breaks, per-section page size and margins, blank page insertion for odd/even page alignment
- **Multi-column layout**: 2+ columns with custom widths and spacing, column breaks, column separators
- **Headers/footers**: default, first-page, and even/odd variants, per-section headers/footers, STYLEREF field resolution (spec-compliant backward search), page number and page count fields, images in headers/footers, correct z-ordering (behind body content)
- **Footnotes**: footnote references, footnote rendering at page bottom with separator line
- **Fields**: PAGE, NUMPAGES, STYLEREF (with spec-compliant search order), field code cached results for non-dynamic fields
- **Hyperlinks**: clickable links in PDF output (URI link annotations)
- **Tab stops**: left, center, right, decimal with leader dots
- **Track changes**: final mode (insertions included, deletions removed — matches Word's PDF export)
- **SmartArt**: rendering via pre-flattened drawing shapes (`dsp:drawing`) with full geometry engine support — all 187 preset shapes, custom geometry, fills (solid, gradient), strokes, and text
- **Document settings**: `word/settings.xml` parsing — even/odd headers, default tab stop interval, mirror margins
- **Compatibility**: `mc:AlternateContent` fallback, structured document tag (`w:sdt`) content extraction, `altChunk` HTML content parsing, smart tag handling
- **Fonts**: cross-platform font search (macOS/Linux/Windows), embedded DOCX font extraction and deobfuscation, font subsetting (CIDFont/Type0), disk-cached font index, font substitution via `fontTable.xml` altName and family-class fallback
- **Output optimization**: font subsetting, content stream compression

### Not yet supported

- **Text**: text shaping/ligatures (fi, fl), complex script shaping (Arabic, Devanagari, etc.), Unicode line breaking for CJK/Thai
- **Tables**: conditional formatting (`tblLook`/`tblStylePr` — banded rows, first/last column styles), nested tables, text direction in cells (`textDirection`)
- **Images**: text wrapping around floating images/textboxes/shapes, EMF/WMF vector images, shape clipping to bounding box
- **Layout**: distribute alignment (`w:jc val="distribute"`), vertical page alignment (`w:vAlign` on section), right-to-left (bidi) text
- **Charts**: 3D charts, stock charts, combo charts, stacked bar rendering, data labels, chart titles, secondary axes
- **SmartArt**: only pre-flattened `dsp:drawing` fallback; no layout engine for documents missing the fallback (see roadmap)
- **PDF features**: bookmarks/outline, document metadata (title, author)
- **Features**: table of contents generation, endnotes, OLE objects, radial/pattern gradient fills
- **Fonts**: bundled fallback fonts, CJK fallback font chain, text shaping via rustybuzz (ligatures, complex scripts)

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
./tools/target/debug/docx-inspect input.docx

# Print font information
./tools/target/debug/docx-fonts input.docx

# Compare two rendered pages
./tools/target/debug/jaccard a.png b.png

# Full fixture diff
./tools/target/debug/case-diff case1
```

## Contributing

Pull requests are welcome!

## License

Apache-2.0
