# docxide-pdf

Library and CLI for converting DOCX files to PDF, matching Microsoft Word's output as closely as possible.

[Try the demo!](https://docxide-demo.fly.dev/)

**🫶 Accessible:** Output PDFs should be just as accessible as Word's export, or better: tagged structure, reading order, language and metadata that pass the same PDF/UA checks.

**🎯 Accurate:** Given a `.docx` file, produce a `.pdf` that is visually identical to what Word would export.

**⚡️ Fast:** Typical conversions complete in under 100ms.

**🤏 Small files:** Output PDFs should be the same size or smaller than Word's export.

Reference PDFs are generated using Microsoft Word for Mac (16.106.1 and later). Most (173 of 234) use the "Best for electronic distribution and accessibility (uses Microsoft online service)" export option; the other 61 were made with "Best for printing", macOS's untagged print path.

## ⚠️ Work in progress.

This crate **might** work for your production case, do give it a try! The API, output quality, and supported features are all actively changing.

## Comparison with other converters

[Side-by-side comparison](https://sverrejb.github.io/docxide-pdf/)

[Score table](https://sverrejb.github.io/docxide-pdf/#scores)

Every fixture in the test corpus rendered side by side by Word (the reference) and 

* docxide-pdf (duh)
* LibreOffice
* [MiniPdf](https://github.com/mini-software/MiniPdf)'s Rust crate
* [rdocx](https://crates.io/crates/rdocx)
* [office2pdf](https://github.com/developer0hye/office2pdf)
* [jubarte-redlines](https://crates.io/crates/jubarte-redlines)
* [dxpdf](https://github.com/nerdy-pro/dxpdf).

Each engine is scored against the Word reference with the same metrics the test
suite uses: three for how the pages look, and one for how accessible the PDF is.

| Metric | What it measures |
|---|---|
| **J** (Jaccard) | Overlap of ink pixels at 150 DPI. Strict: a one-line vertical shift sends it toward zero. |
| **SSIM** | Structural similarity on 8×8 windows with ±8 px vertical tolerance, so small drift is forgiven. |
| **TB** (text boundary) | Share of lines whose first and last word match the reference. Measures line breaking and pagination, independent of fonts. |
| **a11y** | PDF/UA-1 rules failed (veraPDF), how many of those Word passes, and how closely the tag structure and screen-reader text order match Word's. |


## Got a weird DOCX?

If you have a `.docx` file that produces ugly, broken, or just plain wrong output, send it to me! Real-world documents with surprising formatting are the best way to improve. Open an issue or PR with the file included and I will try to make it work.


## AI usage disclaimer 🤖

While the idea, architecture, testing strategy and validation of output are all human, the vast majority of the code as of now is written by various Claude models with access to the PDF specification (ISO-32000) and the Office Open XML File Formats specification (ECMA-376). This project was done as an exercise to get experience with the usage of coding agents.

## Supported features

- **Text**: font embedding (TTF/OTF/TTC), bold, italic (sheared when the face has no italic), single and double underline, strikethrough, double strikethrough, font size, text color, superscript/subscript, raised/lowered text (`w:position`), small caps, all caps, character spacing, text expansion/compression (`w:w`), hidden text (`w:vanish`), kerning (legacy kern table + GPOS PairAdjustment), run borders with color/width/spacing, run shading (`w:shd`) and highlighting, legacy text-effect toggles (`w:outline`, `w:shadow`, `w:emboss`, `w:imprint`), Word 2010 text effects (`w14:glow`, `w14:shadow`, `w14:textOutline`, gradient `w14:textFill`), symbols (`w:sym`), positional tabs (`w:ptab`), non-breaking hyphens, UAX #14 line breaking, Word's per-character glyph fallback for symbols missing from the run's font (Segoe UI Symbol, Arial, Cambria Math)
- **Paragraphs**: left/center/right/justify/distributed alignment (`distribute`), space before/after, line spacing (auto, exact, at-least), first-line and hanging indentation, left/right indentation, character-unit indents (`*Chars`), contextual spacing, keep-next, keep-lines, paragraph borders (top/bottom/left/right/between) with color, paragraph shading, vertical text alignment within the line (`w:textAlignment`), text frames (`w:framePr`)
- **Styles**: paragraph and run style inheritance (`basedOn` chains), document defaults from `docDefaults` (all run properties: bold, italic, caps, smallCaps, vanish, strikethrough, dstrike, underline, color, char_spacing), theme fonts and colors
- **Lists**: bullet and numbered lists with multi-level nesting, custom number formats (incl. CJK: `decimalEnclosedCircle`, `decimalFullWidth`, `aiueoFullWidth`), list style inheritance, `w:lvlRestart`, `w:lvlOverride`/`w:startOverride`, `w:isLgl`, `w:suff`, `w:pStyle` level association
- **Tables**: column widths with auto-fit, merged cells (horizontal `gridSpan` and vertical `vMerge`), row heights (exact and minimum), per-cell borders with color/width, inline `w:tblBorders`, cell shading, pattern/hatch shading, vertical alignment, cell text direction (rotated cells, vertical CJK), cell margins, floating/positioned tables (`tblpPr`), nested tables, conditional formatting (`tblLook`/`tblStylePr` — banded rows/columns, first/last row and column), repeated header rows (`tblHeader`), `cantSplit`, Word-compatible row splitting across pages
- **CJK text**: CIDFont/Identity-H/ToUnicode encoding, Word-compatible substitution of missing CJK fonts by fontTable charset and family (Batang/Malgun Gothic/MS Mincho/SimSun/…, then Apple and Noto faces), per-character font fallback at render time, script-based run splitting via `w:rFonts @eastAsia`, Word's East Asian line height, `compressPunctuation`, `autoSpaceDE`/`autoSpaceDN`, mixed-script justification
- **Right-to-left text**: `w:bidi` paragraphs, Unicode bidi reordering of mixed-direction lines, Arabic contextual joining forms and the lam-alef ligature, right-to-left list labels
- **Images**: inline JPEG/PNG/BMP/GIF/TIFF embedding with sizing and alpha transparency, grayscale and CMYK JPEG support, cropping (`a:srcRect`), EMF/WMF vector translation to PDF form XObjects, SVG pictures (`asvg:svgBlip`) drawn as vectors (shapes, paths, solid paints, strokes, transforms, class stylesheets; everything else falls back to the embedded PNG), anchored/floating images with wrap modes (square, tight, through, topAndBottom), floating image positioning relative to page/margin/column, rotation, clipping to shape geometry, behind-document z-ordering, OLE objects (`w:object`) drawn from their preview picture
- **Picture effects**: outer shadow (`a:outerShdw`), inner shadow, glow, soft edges, reflection — rasterized blur masks via SMask
- **Text boxes**: DrawingML textboxes (`wps:txbx`) and VML fallback (`v:textbox`), shape fills (solid color with theme color support including lumMod/lumOff, linear gradients with multiple color stops), textbox body margins
- **WordArt**: modern DrawingML WordArt with all 40 `prstTxWarp` presets — two-path envelope warping (wave, slant, inflate, etc.) and single-path text-on-a-path (arch, circle), text outlines, shadows, glow effects, bold/italic font variant selection, VML WordArt fallback
- **Shapes & geometry**: all 187 OOXML preset shapes via formula-based geometry engine (guide formulas, adjustment values), custom geometry paths (`a:custGeom` with moveTo, lineTo, cubicBezTo, arcTo), shape fills and strokes, drawing canvases (`wpc:wpc`) and shape groups (`wpg:wgp`/`grpSp`) flattened with nested transforms, connectors
- **Charts**: bar (clustered/stacked/percent-stacked, vertical/horizontal), line, pie, area, doughnut, radar, scatter, bubble — with axis labels, tick marks, gridlines, legends, series markers, bubble fill opacity
- **Math**: Office Math (OMML) equations, inline and display (`m:oMathPara`) with justification — fractions, radicals, n-ary operators with limits, delimiters, matrices, equation arrays, accents, bars, boxes and border boxes, functions, sub/superscripts and pre-scripts, group characters, phantoms, upper/lower limits
- **Page layout**: page size, margins, gutter margins, mirror margins (`w:mirrorMargins`), document grid (`linePitch`), page borders (`w:pgBorders`), vertical page alignment (`w:vAlign`), line numbering (`w:lnNumType`), explicit page breaks, `pageBreakBefore`, automatic page breaking with widow/orphan control
- **Sections**: multiple sections with `nextPage`/`continuous`/`oddPage`/`evenPage` breaks, per-section page size and margins, blank page insertion for odd/even page alignment
- **Multi-column layout**: 2+ columns with custom widths and spacing, column breaks, column separators, column balancing before a continuous section break
- **Headers/footers**: default, first-page, and even/odd variants, per-section headers/footers, STYLEREF field resolution (spec-compliant backward search), page number and page count fields, images in headers/footers, correct z-ordering (behind body content)
- **Footnotes & endnotes**: footnote references and page-bottom rendering with separator line, endnotes flowed at document end, per-section mark numbering formats, shading on reference marks
- **Comments**: `word/comments.xml` rendered in Word's right-hand review pane with callouts and body scaling
- **Fields**: PAGE, NUMPAGES, PAGEREF, STYLEREF (with spec-compliant search order) and IF evaluated in complex `w:fldChar` fields, MACROBUTTON shows its display text or picture, cached results for every other field (and for `w:fldSimple`)
- **Hyperlinks**: clickable external links (URI link annotations) and internal links (bookmarks, footnote marks)
- **Tab stops**: left, center, right, decimal; dot, hyphen and underscore leaders
- **Track changes**: final mode (insertions and moved-to text included, deletions and moved-from text removed — matches Word's PDF export)
- **SmartArt**: rendering via pre-flattened drawing shapes (`dsp:drawing`) with full geometry engine support — all 187 preset shapes, custom geometry, fills (solid, gradient, image), strokes, and text
- **Document settings**: `word/settings.xml` parsing — even/odd headers, mirror margins, default tab stop interval, `gutterAtTop`, `themeFontLang`, `characterSpacingControl`, `compatibilityMode`, `linkStyles`, `doNotExpandShiftReturn`
- **Compatibility**: `mc:AlternateContent` fallback, structured document tag (`w:sdt`) content extraction, `w:customXml` transparent wrappers, `altChunk` HTML content parsing, smart tag handling, VML fallbacks for shapes, textboxes, WordArt and `w:object` embeds
- **Fonts**: cross-platform font search (macOS/Linux/Windows), embedded DOCX font extraction and deobfuscation, font subsetting (CIDFont/Type0), disk-cached font index, Word's missing-font substitution (`fontTable.xml` altName, then Cambria or Calibri by family class)
- **Accessibility**: tagged PDF structure tree (headings, lists, tables, figures with alt text, links, notes, TOC), document and per-run language, XMP metadata, bookmarks from headings, PDF/UA-1 claimed when the document allows it
- **Output optimization**: font subsetting, content stream compression, compressed object streams

### Not yet supported

- **Text**: Latin ligatures (fi, fl), complex script shaping beyond Arabic joining (Devanagari, Thai, etc.), automatic hyphenation (parked — Word's online converter doesn't hyphenate either), underline styles other than single/double, underline color, emphasis marks (`w:em`), `w:fitText`, ruby, drop caps, vertical text outside table cells
- **Right-to-left text**: kashida justification (`mediumKashida`/`highKashida`/`lowKashida` render as plain justify), GPOS mark positioning (harakat), joining across runs with different formatting, mirrored tab stops and `bidiVisual` tables
- **Images**: wrapping around floats anchored more than one paragraph below the text, more than one wrapping float at a time, tight vs through wrapping distinction, SVG `text`/`use`/`image`, gradients, opacity, clip paths and filters (PNG fallback)
- **Layout**: page background color, line numbering `w:restart` (the counter never restarts per page or section)
- **Tab stops**: bar tabs, middle-dot and heavy leaders
- **Track changes**: revision markup view (insertion/deletion colors, change bars, balloons) — only the final view is rendered
- **Charts**: 3D charts (3D pie is drawn flat), stock charts, combo charts (first chart type only), data labels, chart titles, secondary axes
- **Shape effects**: shadow/glow/soft edges on shapes and text boxes (pictures only), dashed outlines (`a:prstDash`), 3D bevel/rotation (`a:scene3d`, `a:sp3d`), preset shadows (`a:prstShdw`), radial/path gradient fills (drawn as linear)
- **SmartArt**: no layout engine for documents missing the `dsp:drawing` fallback (see roadmap)
- **Features**: table of contents generation (cached TOC results are drawn), embedded OLE object data (only the preview picture is drawn), lists inside `altChunk` HTML (dropped)
- **Fonts**: bundled fallback fonts (without the document's fonts installed, Arial/Liberation Sans/DejaVu Sans or Helvetica stand in)

## Examples

Every test case rendered as a Word reference and with docxide-pdf on the [comparison page](https://sverrejb.github.io/docxide-pdf/).

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

The CLI never overwrites: if the output exists it writes `output(2).pdf`, `output(3).pdf` and so on. `RUST_LOG=info` prints timings.

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

To keep everything in memory, `convert_docx_bytes_to_pdf_bytes` takes the DOCX as bytes and returns the PDF as bytes:

```rust
let pdf: Vec<u8> = docxide_pdf::convert_docx_bytes_to_pdf_bytes(&docx_bytes)?;
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
| `DOCXIDE_FONTS` | Additional font directories to search, colon-separated (`;` on Windows). Highest priority: a face found there wins over the system's copy of the same family and style, and the first directory listed wins over the ones after it. System fonts still fill in the families and styles the directories lack. |
| `DOCXIDE_NO_FONT_CACHE` | Set to any value to disable the font index disk cache. Forces a full font scan on every conversion. Useful for debugging font resolution issues. |

Font scanning results are cached to disk (per-directory, invalidated by mtime). The cache is stored at:
- **macOS**: `~/Library/Caches/docxide-pdf/font-index.tsv`
- **Linux**: `$XDG_CACHE_HOME/docxide-pdf/font-index.tsv` (default `~/.cache/`)
- **Windows**: `%LOCALAPPDATA%\docxide-pdf\cache\font-index.tsv`

### Fonts

docxide-pdf lays text out with the fonts it finds. The PDF can only match Word's if the font files match too.

- **Where fonts come from**: first fonts embedded in the DOCX, then the directories in `DOCXIDE_FONTS` (the first one listed wins), then the system font directories. Lookup is by family plus bold and italic, so a Regular you supply does not block the system's Bold.
- **Same name, different file, different pages.** Line height comes from numbers inside the font file, and two files of the same family can carry different numbers. For example, the Arial and Times New Roman that ship with Office for Mac give 13.4pt lines at 12pt, while the files Word itself exports with give 13.8pt. That is enough to move every page break in a long document. The version number does not tell you which one you have.
- **A font missing some characters is still used.** If the Bold has no Greek letters, say, the Greek comes from another font, one letter at a time, with no warning.
- **To get the same PDF everywhere**, embed the fonts in the DOCX (Word: "Embed fonts in the file"), or put a checked set of font files in one directory and point `DOCXIDE_FONTS` at it. Run with `RUST_LOG=info` to see which fonts were substituted.

## Testing

Tests require `mutool` on `PATH` for PDF-to-PNG rendering:

```bash
brew install mupdf        # macOS
apt install mupdf-tools   # Debian/Ubuntu
```

```bash
# Run all tests with a compact report of what changed
./tools/run-tests.sh

# Run only the visual comparison (Jaccard and SSIM)
./tools/run-tests.sh --test visual_comparison
```

`.cargo/config.toml` points `DOCXIDE_FONTS` at a `fonts/` directory that isn't in the repository (the fixtures' fonts can't be redistributed). Without those fonts many fixtures fall back to substitutes and score lower. The accessibility test also needs veraPDF and Poppler (`brew install verapdf poppler`) and skips without them.

[SCORING.md](SCORING.md) explains every score: the visual metrics, the accessibility metrics and what fails the suite.

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

# Browse reference vs generated pages per fixture, with notes
./tools/target/debug/case-browser

# Fixture features and scores, e.g. only the failing ones
./tools/target/debug/analyze-fixtures --failing

# Are our font files the builds Word embedded in these references? (needs mutool)
python3 tools/font_builds.py tests/fixtures/fonts/multi_font/reference.pdf -- fonts
```

## Contributing

Pull requests are welcome!

## License

Apache-2.0
