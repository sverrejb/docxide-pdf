# docxide-pdf

Library and CLI for converting DOCX files to PDF, matching Microsoft Word's output as closely as possible.

[Try the demo!](https://docxide-demo.fly.dev/)

## ⚠️ Work in progress.

This crate **might** work for your production case, do give it a try! The API, output quality, and supported features are all actively changing.

## Comparison with other converters

**[Live comparison → sverrejb.github.io/docxide-pdf](https://sverrejb.github.io/docxide-pdf/)**

Every fixture in the test corpus rendered side by side by Word (the reference),
docxide-pdf, LibreOffice, [MiniPdf](https://github.com/mini-software/MiniPdf)'s Rust crate,
and [rdocx](https://crates.io/crates/rdocx).

Each engine is scored against the Word reference with the same three metrics the test
suite uses:

| Metric | What it measures |
|---|---|
| **J** (Jaccard) | Overlap of ink pixels at 150 DPI. Strict: a one-line vertical shift sends it toward zero. |
| **SSIM** | Structural similarity on 8×8 windows with ±8 px vertical tolerance, so small drift is forgiven. |
| **TB** (text boundary) | Share of lines whose first and last word match the reference. Measures line breaking and pagination, independent of fonts. |


## Got a weird DOCX?

If you have a `.docx` file that produces ugly, broken, or just plain wrong output, send it to me! Real-world documents with surprising formatting are the best way to improve. Open an issue or PR with the file included and I will try to make it work.


## Goals

A Rust library and CLI tool for converting DOCX files to PDF, with the goal of matching Microsoft Word's PDF export as closely as possible.<sup>*</sup>

**Accurate:** Given a `.docx` file, produce a `.pdf` that is visually identical to what Word would export.

**Fast:** Typical conversions complete in under 100ms.

**Small files:** Output PDFs should be the same size or smaller than Word's export.

*<sub>Reference PDFs are generated using Microsoft Word for Mac (16.106.1) with the "Best for electronic distribution and accessibility (uses Microsoft online service)" export option.</sub>

## AI usage disclaimer 🤖

While the idea, architecture, testing strategy and validation of output are all human, the vast majority of the code as of now is written by Claude Opus 4.6 with access to the PDF specification (ISO-32000) and the Office Open XML File Formats specification (ECMA-376). This project was done as an exercise to get experience with the usage of coding agents.

## Supported features

- **Text**: font embedding (TTF/OTF/TTC), bold, italic, underline, strikethrough, double strikethrough, font size, text color, superscript/subscript, small caps, all caps, character spacing, text expansion/compression (`w:w`), hidden text (`w:vanish`), kerning (legacy kern table + GPOS PairAdjustment), vertical text (CJK), run borders with color/width/spacing, legacy text-effect toggles (`w:outline`, `w:shadow`, `w:emboss`, `w:imprint`), UAX #14 line breaking
- **Paragraphs**: left/center/right/justify/distributed alignment (`distribute`), space before/after, line spacing (auto, exact, at-least), first-line and hanging indentation, left/right indentation, contextual spacing, keep-next, keep-lines, paragraph borders (top/bottom/left/right/between) with color, paragraph shading, run highlighting
- **Styles**: paragraph and run style inheritance (`basedOn` chains), document defaults from `docDefaults` (all run properties: bold, italic, caps, smallCaps, vanish, strikethrough, dstrike, underline, color, char_spacing), theme fonts and colors
- **Lists**: bullet and numbered lists with multi-level nesting, custom number formats (incl. CJK: `decimalEnclosedCircle`, `decimalFullWidth`, `aiueoFullWidth`), list style inheritance, `w:lvlRestart`, `w:pStyle` level association
- **Tables**: column widths with auto-fit, merged cells (horizontal `gridSpan` and vertical `vMerge`), row heights (exact and minimum), per-cell borders with color/width, inline `w:tblBorders`, cell shading, pattern/hatch shading, vertical alignment, cell text direction (rotated cells), cell margins, floating/positioned tables (`tblpPr`), nested tables, conditional formatting (`tblLook`/`tblStylePr` — banded rows/columns, first/last row and column), Word-compatible row splitting across pages
- **CJK text**: CIDFont/Identity-H/ToUnicode encoding, platform-specific font fallback chains (Hiragino/Noto/Yu Gothic), per-character font fallback at render time, script-based run splitting via `w:rFonts @eastAsia`
- **Images**: inline JPEG/PNG embedding with sizing and alpha transparency, grayscale and CMYK JPEG support, EMF/WMF vector translation to PDF form XObjects, anchored/floating images with wrap modes (square, tight, through, topAndBottom), floating image positioning relative to page/margin/column, rotation, clipping to shape geometry, behind-document z-ordering
- **Picture effects**: outer shadow (`a:outerShdw`), inner shadow, glow, soft edges, reflection — rasterized blur masks via SMask
- **Text boxes**: DrawingML textboxes (`wps:txbx`) and VML fallback (`v:textbox`), shape fills (solid color with theme color support including lumMod/lumOff, linear gradients with multiple color stops), textbox body margins
- **WordArt**: modern DrawingML WordArt with all 40 `prstTxWarp` presets — two-path envelope warping (wave, slant, inflate, etc.) and single-path text-on-a-path (arch, circle), text outlines, shadows, glow effects, bold/italic font variant selection, VML WordArt fallback
- **Shapes & geometry**: all 187 OOXML preset shapes via formula-based geometry engine (guide formulas, adjustment values), custom geometry paths (`a:custGeom` with moveTo, lineTo, cubicBezTo, arcTo), shape fills and strokes, drawing canvases (`wpc:wpc`) and shape groups (`wpg:wgp`/`grpSp`) flattened with nested transforms, connectors
- **Charts**: bar (clustered/stacked/percent-stacked, vertical/horizontal), line, pie, area, doughnut, radar, scatter, bubble — with axis labels, tick marks, gridlines, legends, series markers, bubble fill opacity
- **Math**: Office Math (OMML) equations, inline and display (`m:oMathPara`) with justification
- **Page layout**: page size, margins, gutter margins, document grid (`linePitch`), page borders (`w:pgBorders`), vertical page alignment (`w:vAlign`), line numbering (`w:lnNumType`), explicit page breaks, `pageBreakBefore`, automatic page breaking with widow/orphan control
- **Sections**: multiple sections with `nextPage`/`continuous`/`oddPage`/`evenPage` breaks, per-section page size and margins, blank page insertion for odd/even page alignment
- **Multi-column layout**: 2+ columns with custom widths and spacing, column breaks, column separators
- **Headers/footers**: default, first-page, and even/odd variants, per-section headers/footers, STYLEREF field resolution (spec-compliant backward search), page number and page count fields, images in headers/footers, correct z-ordering (behind body content)
- **Footnotes & endnotes**: footnote references and page-bottom rendering with separator line, endnotes flowed at document end, per-section mark numbering formats, shading on reference marks
- **Comments**: `word/comments.xml` rendered in Word's right-hand review pane with callouts and body scaling
- **Fields**: PAGE, NUMPAGES, PAGEREF, STYLEREF (with spec-compliant search order), field code cached results for non-dynamic fields
- **Hyperlinks**: clickable links in PDF output (URI link annotations)
- **Tab stops**: left, center, right, decimal with leader dots
- **Track changes**: final mode (insertions included, deletions removed — matches Word's PDF export)
- **SmartArt**: rendering via pre-flattened drawing shapes (`dsp:drawing`) with full geometry engine support — all 187 preset shapes, custom geometry, fills (solid, gradient, image), strokes, and text
- **Document settings**: `word/settings.xml` parsing — even/odd headers, default tab stop interval, mirror margins
- **Compatibility**: `mc:AlternateContent` fallback, structured document tag (`w:sdt`) content extraction, `w:customXml` transparent wrappers, `altChunk` HTML content parsing, smart tag handling, VML fallbacks for shapes, textboxes, WordArt and `w:object` embeds
- **Fonts**: cross-platform font search (macOS/Linux/Windows), embedded DOCX font extraction and deobfuscation, font subsetting (CIDFont/Type0), disk-cached font index, font substitution via `fontTable.xml` altName and family-class fallback
- **Output optimization**: font subsetting, content stream compression

### Not yet supported

- **Text**: text shaping/ligatures (fi, fl), complex script shaping (Arabic, Devanagari, etc.), automatic hyphenation (parked — Word's online converter doesn't hyphenate either)
- **Images**: look-back text wrapping (text before a float anchor wrapping beside the image), tight vs through wrapping distinction
- **Layout**: mirror margins (parsed but not applied to even pages), right-to-left (bidi) text, kashida justification (`mediumKashida`/`highKashida`/`lowKashida` render as plain justify — glyph elongation needs Arabic shaping)
- **Charts**: 3D charts, stock charts, combo charts, data labels, chart titles, secondary axes
- **Shape effects**: 3D bevel/rotation (`a:scene3d`, `a:sp3d`), preset shadows (`a:prstShdw`), radial/path gradient fills (axial only)
- **SmartArt**: no layout engine for documents missing the `dsp:drawing` fallback (see roadmap)
- **Features**: table of contents generation, OLE objects
- **Fonts**: bundled fallback fonts, text shaping via rustybuzz (ligatures, complex scripts)

## Examples

See more examples in the [showcase](https://github.com/sverrejb/docxide-pdf/tree/main/showcase#readme)

<!-- showcase-start -->
<table>
  <tr><th>MS Word</th><th>docxide-pdf</th></tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case11_ref.png"/><br/><sub>Report with headers, footers & page numbers — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case11_gen.png"/><br/><sub>92.0% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case34_ref.png"/><br/><sub>20 preset shapes via geometry engine — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case34_gen.png"/><br/><sub>96.9% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case8_ref.png"/><br/><sub>Embedded fonts & mixed typography — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case8_gen.png"/><br/><sub>94.1% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case22_ref.png"/><br/><sub>Three-column newsletter layout — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case22_gen.png"/><br/><sub>93.3% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case30_ref.png"/><br/><sub>Line, pie & area charts — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/case30_gen.png"/><br/><sub>83.3% SSIM</sub></td>
  </tr>
  <tr>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/centrifugal_ref.png"/><br/><sub>Real-world document (scraped) — reference</sub></td>
    <td align="center"><img src="https://raw.githubusercontent.com/sverrejb/docxide-pdf/main/showcase/centrifugal_gen.png"/><br/><sub>94.0% SSIM</sub></td>
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

## Works well with `docxide-template`

[`docxide-template`](https://github.com/sverrejb/docxide-template) is a sibling crate for type-safe MS Word templates. It scans a folder of `.docx` files at compile time and generates a Rust struct per template, with `{Placeholder}` patterns turned into snake_case fields. Pair it with `docxide-pdf` to go from template → filled DOCX → PDF in a single, fully in-memory pipeline:

```rust
use docxide_pdf::convert_docx_bytes_to_pdf;
use docxide_template::generate_templates;
use std::path::Path;

generate_templates!("templates");

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let doc = HelloWorld {
        first_name: "Alice".into(),
        company: "Acme Corp".into(),
    };

    let docx_bytes = doc.to_bytes()?;
    convert_docx_bytes_to_pdf(&docx_bytes, Path::new("output/greeting.pdf"))?;
    Ok(())
}
```

100% Rust, end to end — no temporary files, and no Word or LibreOffice install required on the host. Fill the template in memory, hand the bytes to `convert_docx_bytes_to_pdf`, and write the PDF. Combined with `docxide-template`'s `embed` feature, you get a single self-contained binary that turns structured data into a polished PDF.

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
