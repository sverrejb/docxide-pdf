[![AI:W WRITE](https://img.shields.io/badge/AI-WRITE-grey?style=flat&logoSize=auto&labelColor=13171f)](https://0rwx.com)

# docxide-pdf

Library and CLI for converting DOCX files to PDF, matching Microsoft Word's output as closely as possible.

[Try the demo!](https://docxide.fly.dev/)

**🫶 Accessible:** Output PDFs should be just as accessible as Word's export, or better: tagged structure, reading order, language and metadata that pass the same PDF/UA checks.

**🎯 Accurate:** Given a `.docx` file, produce a `.pdf` that is visually identical to what Word would export.

**⚡️ Fast:** Typical conversions complete in under 100ms.

**🤏 Small files:** Output PDFs should be the same size or smaller than Word's export.

Reference PDFs are generated using Microsoft Word for Mac (16.106.1 and later). Most use the "Best for electronic distribution and accessibility (uses Microsoft online service)" export option; some have been converted with the "Best for printing" option where some features were not supported by the online conversion.

## ⚠️ Work in progress.

This crate **might** work for your production case, do give it a try! The API, output quality, and supported features are all actively changing.

## Comparison with other converters

[![Mean Jaccard and SSIM per engine](https://sverrejb.github.io/docxide-pdf/scores.svg)](https://docxide.fly.dev/compare#scores)

[Side-by-side comparison](https://docxide.fly.dev/compare)

[Score table](https://docxide.fly.dev/compare#scores)

Every fixture in the test corpus rendered side by side by Word (the reference) and 

* docxide-pdf
* LibreOffice
* [MiniPdf](https://github.com/mini-software/MiniPdf)'s Rust crate
* [rdocx](https://crates.io/crates/rdocx)
* [office2pdf](https://github.com/developer0hye/office2pdf)
* [jubarte-redlines](https://crates.io/crates/jubarte-redlines)
* [dxpdf](https://github.com/nerdy-pro/dxpdf)
* [WordCraft](https://github.com/storytold/wordcraft)
* [Stirling Office Convert](https://github.com/Stirling-Tools/Stirling-Office-Convert).

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

[![AI:W WRITE](https://img.shields.io/badge/AI-WRITE-grey?style=flat&logoSize=auto&labelColor=13171f)](https://0rwx.com)

While the idea, architecture, testing strategy and validation of output are all human, the vast majority of the code as of now is written by various Claude models with access to the PDF specification (ISO-32000) and the Office Open XML File Formats specification (ECMA-376). This project was done as an exercise to get experience with the usage of coding agents.

## Not yet supported

- **Text**: Latin ligatures (fi, fl), complex script shaping beyond Arabic joining (Devanagari, Thai, etc.), automatic hyphenation (parked — Word's online converter doesn't hyphenate either), theme tint and shade on underline colour (`w:u` `w:themeTint`/`w:themeShade`), emphasis marks (`w:em`), vertical text outside table cells
- **Right-to-left text**: kashida justification (lines break as Word does, but are justified by widening spaces, not by stretching letters), GPOS mark positioning (harakat), joining across runs with different formatting, mirrored tab stops
- **Images**: wrapping around floats anchored more than three paragraphs below the text, wrapping around floats from more than one anchor paragraph at once (a later anchor's floats replace the earlier ones), tight vs through wrapping distinction, SVG `text`/`use`/`image`, nested `svg`, gradients, opacity, clip paths, masks, markers and filters (PNG fallback)
- **Layout**: page background color
- **Tab stops**: bar tabs (treated as left stops, no line drawn)
- **Track changes**: revision balloons for text changes (the markup view shows them inline; only deleted and inserted table cells get balloons)
- **Charts**: true-perspective 3-D for horizontal bars and pies (drawn oblique or orthographic), surface, stock, pie-of-pie and bar-of-pie charts, combo charts (only one chart type is drawn), secondary axes
- **Shape effects**: shadow and soft edges on shapes and text boxes (pictures only), dashed outlines on connectors, picture borders and VML shapes, 3D bevel/rotation (`a:scene3d`, `a:sp3d`), preset shadows (`a:prstShdw`), radial/path gradient fills (drawn as linear)
- **SmartArt**: no layout engine for documents missing the `dsp:drawing` fallback (not planned)
- **Features**: regeneration of a TOC Word has already filled in (its cached entries are drawn, with re-measured page numbers; only a one-paragraph placeholder is generated), embedded OLE object data (only the preview picture is drawn), lists inside `altChunk` HTML (dropped)
- **Fonts**: no bundled fallback fonts. Without the document's fonts (installed, embedded in the DOCX, or passed with `.font()` or `.font_dirs()`), Arial, Liberation Sans, Arimo, Helvetica or DejaVu Sans stand in, then the standard-14 Helvetica

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

# Keep every picture's source pixels (Word's online export resamples pictures above 192 ppi)
docxide-pdf --no-resample input.docx

# Fonts from a directory of your own first (repeatable; the first wins)
docxide-pdf --fonts ./fonts input.docx
```

The CLI never overwrites: if the output exists it writes `output(2).pdf`, `output(3).pdf` and so on. `RUST_LOG=info` prints timings.

### Library

```bash
cargo add docxide-pdf --no-default-features
```

or in `Cargo.toml`: `docxide-pdf = { version = "0.22", default-features = false }`. This skips the CLI's dependencies (`clap`, `env_logger`).

One function, `convert`, takes the DOCX as bytes and returns a `Conversion` with the PDF as bytes. File I/O is yours:

```rust
use docxide_pdf::{ConvertOptions, convert};

let out = convert(&std::fs::read("input.docx")?, &ConvertOptions::new())?;
std::fs::write("output.pdf", &out.pdf)?;
```

`ConvertOptions::new()` reproduces Word's online export; each builder method moves one behaviour towards desktop Word or adapts the conversion to its host (see [Configuration](#configuration)). The library reads no environment variables. Build the options once and reuse them: they are cheap to clone and can be shared across threads.

Besides `pdf`, a `Conversion` carries:

- `bookmarks`: where each bookmarked paragraph was laid out (below);
- `missing_fonts`: the fonts the document asks for that no embedded, supplied or installed face provides, each with the requested bold and italic. A missing style is listed even when the regular face stands in with synthetic bold or italic. When Word draws a different font in place of the document's (Arial for Helvetica, Times New Roman for Times, Courier New for Courier), `name` is that font and `replaces` holds the document's name, since supplying the document's font would change nothing. Use it to tell the user which font to supply: `ConvertOptions::font` takes the font file's bytes, and converting again uses it;
- `display`: with `.display_list(true)`, what each page draws, with the fonts and pictures it draws them with ([`display`](https://docs.rs/docxide-pdf/latest/docxide_pdf/display/)), recorded from the same render as the PDF, for a caller that writes PDF or renders pages itself.

`convert` fails only with `Error::InvalidDocx`, when the input is not a DOCX (not a ZIP archive, or no readable `word/document.xml`). A new output becomes a new `Conversion` field and a new failure a new `Error` variant, so adding either doesn't break callers.

### Bookmark positions

`Conversion::bookmarks` lists where each bookmarked paragraph was laid out, so a
host can map a place in the PDF back to the document (to scroll to it or
highlight it):

```rust
let out = docxide_pdf::convert(&docx_bytes, &ConvertOptions::new())?;
for p in &out.bookmarks {
    println!("{}: page {} top {} → page {} bottom {}", p.name, p.page, p.top, p.end_page, p.bottom);
}
```

A position describes the paragraph containing `bookmarkStart`: its first
fragment's page/top and last fragment's page/bottom, including table cells,
nested tables and fragments spanning pages or columns. Positions come from the
final render pass and include section vertical alignment. Pages are zero-based
and coordinates are PDF points measured upward from the page bottom. This is
not a multi-paragraph bookmark-range API; headers, footers and text boxes are
excluded.

### Range links

`ConvertOptions::range_links` writes invisible link annotations over the text
laid out between two bookmarks, so a viewer can link a place in the PDF back
to whatever produced it:

```rust
use docxide_pdf::{ConvertOptions, RangeLink, convert};

let options = ConvertOptions::new().range_links([
    RangeLink::new("start", "end", "app:item/1"),
    // A second link over the range's heading: its leading bold lines.
    RangeLink::new("start2", "end2", "app:item/2").heading_uri("app:heading/2"),
]);
let out = convert(&docx_bytes, &options)?;
```

A range runs from the `bookmarkStart` named `start` to the one named `end`,
wherever they sit: between paragraphs and tables, or between the runs of a line
(the link then starts or stops inside the line, inside a word if the runs
split one). Every line holding text or a picture of the range gets the link,
as the box of the range's glyphs on it (advance widths, the face's `hhea`
ascender and descender) or of the picture. Per page, the range's first and
last lines keep their own rectangle and the lines in between share one: at
most three rectangles a page, plus one per floating picture anchored in the
range, drawn over the picture's box (rotated: that box's bounding box) and
never merged with the lines. `heading_uri` adds one more link over the
range's leading lines on its first page while every run on them is bold (at
most four lines, no gap taller than a line between them). Only the body is
covered (table cells, text boxes and pictures included; headers, footers,
notes and list labels are not), and links follow section vertical alignment.
A range whose bookmarks are missing or out of order links nothing; without range
links the PDF is unchanged.

### Hosts without a filesystem (WebAssembly)

From a source checkout, compile the library for the bare wasm target:

```sh
rustup target add wasm32-unknown-unknown
cargo build --locked --lib --no-default-features --target wasm32-unknown-unknown
```

This produces `target/wasm32-unknown-unknown/debug/libdocxide_pdf.rlib`, for linking
into a Rust host that exports its own wasm interface. It does not produce a
JavaScript package or browser bindings.

On this target font directories are not scanned and there are no diagnostic
timings. Text uses DOCX-embedded fonts and the fonts passed with
`ConvertOptions::font`, and standard-14 Helvetica (not embedded) for everything
else; `Conversion::missing_fonts` names the fonts worth supplying.

Validation from a source checkout (tested with Rust 1.96, 1.99 and 1.100, and Node.js):

```sh
./tools/check-wasm-host.sh
```

The smoke script builds the `wasm_host_smoke` example
(`target/wasm32-unknown-unknown/debug/examples/wasm_host_smoke.wasm`) and runs it in Node. That small host converts an embedded fixture without JS
clock or filesystem implementations; see `tools/wasm_host_smoke.rs` for its
Rust entry point and `tools/check-wasm-host.mjs` for the Node runner.

## Works well with `docxide-template`

[`docxide-template`](https://github.com/sverrejb/docxide-template) is a sibling crate for type-safe MS Word templates. It scans a folder of `.docx` files at compile time and generates a Rust struct per template, with `{Placeholder}` patterns turned into snake_case fields. Pair it with `docxide-pdf` to go from template → filled DOCX → PDF in a single, fully in-memory pipeline:

```rust
use docxide_pdf::{ConvertOptions, convert};
use docxide_template::generate_templates;

generate_templates!("templates");

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let doc = HelloWorld {
        first_name: "Alice".into(),
        company: "Acme Corp".into(),
    };

    let docx_bytes = doc.to_bytes()?;
    let out = convert(&docx_bytes, &ConvertOptions::new())?;
    std::fs::write("output/greeting.pdf", &out.pdf)?;
    Ok(())
}
```

100% Rust, end to end — no temporary files, and no Word or LibreOffice install required on the host. Fill the template in memory, hand the bytes to `convert`, and write the PDF. Combined with `docxide-template`'s `embed` feature, you get a single self-contained binary that turns structured data into a polished PDF.

## Configuration

Each setting is a `ConvertOptions` builder method in the library, and a CLI flag or a `DOCXIDE_*` environment variable in the CLI. Without any of them you get Word's online export. The library itself never reads the environment.

| Setting | Default | Library (`ConvertOptions`) | CLI flag | CLI environment variable |
|---|---|---|---|---|
| Print tracked changes as Word's All Markup view: insertions underlined, deletions struck through in the author's colour, change bars in the margin | Final text, no markup | `.revisions(RevisionView::Markup)` | `--markup` | `DOCXIDE_REVISIONS=markup` |
| Keep every picture's source pixels instead of resampling pictures above 192 ppi down to 192 ppi | Resample, like Word's online export | `.resample_pictures(false)` | `--no-resample` | `DOCXIDE_NO_RESAMPLE=1` |
| Convert faster into a larger PDF, for a local preview: light stream compression, JPEGs embedded as stored, pictures resampled only above 384 ppi | Word-sized output | `.compression(Compression::Fast)` | `--fast-compression` | `DOCXIDE_FAST_COMPRESSION=1` |
| Extra font directories, searched before the system fonts. The first one listed wins, and system fonts fill in what they lack (see [Fonts](#fonts)) | None | `.font_dirs(["DIR", ...])` | `--fonts DIR` (repeat for more) | `DOCXIDE_FONTS=DIR1:DIR2` (`;` on Windows) |
| A font file's bytes, used before installed faces of the same family and style; repeat for more | None | `.font(bytes)` | | |
| Search the system font directories | On (off on bare wasm) | `.system_fonts(false)` | | |
| Scan the font directories on every conversion instead of reading the font index disk cache. Useful when debugging font resolution | Cache on | `.font_cache(false)` | `--no-font-cache` | `DOCXIDE_NO_FONT_CACHE=1` |
| Invisible links over bookmarked text (see [Range links](#range-links)) | None | `.range_links([RangeLink::new(start, end, uri)])` | `--range-link "START END URI [HEADING_URI]"` | `DOCXIDE_RANGE_LINKS`, one per line |
| Also return the pages as a display list (`Conversion::display`) | Off | `.display_list(true)` | | |
| Draw the float wrap zones on the page (debugging) | Off | | `--debug-wrap` | `DOCXIDE_DEBUG_WRAP=1` |

```rust
use docxide_pdf::{ConvertOptions, RevisionView, convert};

let options = ConvertOptions::new()
    .revisions(RevisionView::Markup)
    .resample_pictures(false)
    .font_dirs(["./fonts", "/opt/company-fonts"])
    .font(std::fs::read("Aptos-Bold.ttf")?)
    .font_cache(false);

let out = convert(&std::fs::read("input.docx")?, &options)?;
```

In the CLI an on/off variable counts as off when it is empty, `0`, `false`, `no` or `off`; its flag can only switch the setting on, so to switch a variable off, unset it or set it to `0`. `--fonts` replaces `DOCXIDE_FONTS` instead of adding to it. There is no configuration file: an application that wants one fills in `ConvertOptions` from its own.

Font scanning results are cached to disk (per-directory, invalidated by mtime). The cache is stored at:
- **macOS**: `~/Library/Caches/docxide-pdf/font-index.tsv`
- **Linux**: `$XDG_CACHE_HOME/docxide-pdf/font-index.tsv` (default `~/.cache/`)
- **Windows**: `%LOCALAPPDATA%\docxide-pdf\cache\font-index.tsv`

### Fonts

docxide-pdf lays text out with the fonts it finds. The PDF can only match Word's if the font files match too.

- **Where fonts come from**: first the fonts passed with `.font()`, then the font directories (`.font_dirs()` or `DOCXIDE_FONTS`; the first one listed wins), then the system font directories. A font embedded in the DOCX stands in for a family none of them has, as in Word. Lookup is by family plus bold and italic, so a Regular you supply does not block the system's Bold.
- **Same name, different file, different pages.** Line height comes from numbers inside the font file, and two files of the same family can carry different numbers. For example, the Arial and Times New Roman that ship with Office for Mac give 13.4pt lines at 12pt, while the files Word itself exports with give 13.8pt. That is enough to move every page break in a long document. The version number does not tell you which one you have.
- **A font missing some characters is still used.** If the Bold has no Greek letters, say, the Greek comes from another font, one letter at a time, with no warning.
- **To get the same PDF everywhere**, embed the fonts in the DOCX (Word: "Embed fonts in the file"), or put a checked set of font files in one directory and pass it with `.font_dirs()` (CLI: `--fonts` or `DOCXIDE_FONTS`). Run with `RUST_LOG=info` to see which fonts were substituted.

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
