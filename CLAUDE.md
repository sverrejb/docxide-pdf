# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

docxide-pdf is a Rust library for converting DOCX files to PDF with the goal of matching Microsoft Word's PDF export as closely as possible. The project is in early development.

## Common Commands

```bash
# Build the library
cargo build

# Run tests (compact output — ALWAYS use this instead of raw cargo test)
./tools/run-tests.sh                              # all tests, compact report
./tools/run-tests.sh --test visual_comparison      # one test suite
./tools/run-tests.sh --case case5                  # one fixture
./tools/run-tests.sh --verbose                     # full cargo output (debugging only)
./tools/run-tests.sh --test accessibility          # PDF/UA-1 + structure-tree parity vs Word (needs: brew install verapdf poppler)

# Run a specific test (verbose — avoid unless debugging)
cargo test test_name

# Build in release mode
cargo build --release

# Check code without building
cargo check

# Format code
cargo fmt

# Run linter
cargo clippy

# Build analysis tools
cd tools && cargo build && cd ..

# Analyze scraped fixture features and scores
./tools/target/debug/analyze-fixtures              # full table
./tools/target/debug/analyze-fixtures --failing    # only failing fixtures
./tools/target/debug/analyze-fixtures --fonts      # include font lists
./tools/target/debug/analyze-fixtures --audit      # feature audit: which XML features appear in failing vs passing
./tools/target/debug/analyze-fixtures --grep "w:caps"  # search for specific XML pattern across all fixtures

# Side-by-side engine viewer: Word reference | ours | LibreOffice | MiniPdf | rdocx | office2pdf → comparison/index.html (static site, gitignored)
python3 tools/engine_compare.py --open                 # all fixtures (reuses tests/output PNGs, caches the rest in comparison/work/)
python3 tools/engine_compare.py --case case41 --case 'case2*'   # exact name or glob
python3 tools/engine_compare.py --html-only            # rebuild index.html from comparison/work/manifest.json (no re-scoring)
python3 tools/engine_compare.py --shard 0/8            # one of 8 shards, by name hash (CI runs 8 such jobs, then --merge joins their manifests)
tools/deploy_comparison.sh [remote] [branch]           # publish comparison/ (work/ excluded) as an orphan gh-pages commit (DRY_RUN=1 to preview)
# CI does both on every push to main: .github/workflows/comparison.yml (fonts come from the private sverrejb/docxide-pdf-assets repo; only tracked cases/ are compared)
# MiniPdf = the Rust crate's CLI (`cargo install minipdf-cli`), never the .NET engine; rdocx via `cargo install rdocx`; office2pdf via `cargo install office2pdf-cli`; jubarte-redlines via `cargo install jubarte-redlines` (binary `jubarte`); LibreOffice via soffice

# Word for Mac, unattended: export .docx → .pdf with Word itself (probe documents, what-if experiments)
python3 tools/word_export.py <file-or-dir>... [--out DIR] [--force] [--timeout S] [--preset online|print]   # default online = tagged refs; stages inside Word's container, answers dialogs; needs the sandbox off
python3 tools/word_export.py --check                   # preflight: Word present, Automation granted

# Accuracy work: verify each rule change on all fixtures and, where available, an external corpus
tools/score_snapshot.sh <label> [<prev label>]         # full visual suite → tests/output/snapshots/<label>.json, compared with <prev>
python3 tools/compare_scores.py <a.json> <b.json>      # group means + every case that moved
python3 tools/ab_view.py <label> <cli> <group/case>... [-v]   # render with a given CLI build, print vdiff page drift (VDIFF=path/to/vdiff)
python3 tools/pdf_lines.py <pdf> <page> [ymin ymax]    # rules (borders) and baselines from a PDF, y from page top
python3 tools/line_diff.py <ref.pdf> <gen.pdf> [n]     # first places where text lines differ (where reflow starts)
python3 tools/word_x_diff.py <ref.pdf> <gen.pdf> <page>  # per-word start-x drift (width/justification errors)
python3 tools/ink_diff.py <ref.pdf> <gen.pdf> <page> <out.png>  # blue = ref-only ink, red = ours-only, grey = both
python3 tools/docx_edit.py <in> <out> <part> <old> <new>  # literal edit of one docx part for what-if experiments
python3 tools/convert_scan.py [<cli>]                  # convert every fixture with a timeout; reports hangs/crashes
python3 tools/corpus_score.py <dir> <cli> <label>      # score docx/+pdf/ pairs → tests/output/corpus/<label>.json (needs release page-metrics)
python3 tools/corpus_compare.py <before> <after>       # corpus mean + biggest movers
python3 tools/corpus_view.py <dir> <label> <cli> <prefix>... [-v] [--score]  # vdiff (and metrics) for chosen corpus docs
```

## Architecture

The project aims to:
1. Parse DOCX (Office Open XML) format
2. Render to PDF following Word's layout and formatting rules
3. Test against real-world DOCX files with Word-generated reference PDFs
4. Use visual comparison testing for pixel-level accuracy verification

Testing strategy emphasizes:
- Large corpus of diverse DOCX files
- Reference PDFs generated by MS Word
- Visual diff testing via image conversion

## Roadmap

See `roadmap.md` for planned work prioritized by impact. **Always consult the roadmap** before starting work to understand priorities and avoid duplicating completed work. **Update the roadmap** when completing items, discovering new issues, or when priorities shift.

## Reference Material

- PDF spec is ingested into local RAG system at `/Users/sverrejb/specs/pdf-spec.pdf` for semantic search queries about PDF internals
- Office Open XML spec is ingested into local RAG system at `/Users/sverrejb/specs/office.pdf` for semantic search queries about DOCX format and WordprocessingML

**IMPORTANT**: Always use the `mcp__local-rag__query_documents` tool to search the specs when you need information about PDF internals, DOCX format, or WordprocessingML. The specs contain authoritative technical details that are essential for correct implementation.

## Module Structure

```
src/
  lib.rs              — public API: convert_docx_to_pdf(input, output), convert_docx_bytes_to_pdf(bytes, output) and convert_docx_bytes_to_pdf_bytes(bytes) -> Vec<u8>; paths are impl AsRef<Path>
  error.rs            — Error enum (InvalidDocx, Zip, Xml, Pdf, Io variants)
  main.rs             — CLI binary (behind `cli` feature)
  model/
    mod.rs            — Document, Section, Paragraph, Run IR, font table, settings
    chart.rs          — Chart IR (ChartType: bar/line/pie/area/scatter/bubble/doughnut/radar)
    drawing.rs        — shape geometry (presets + custom geometry), fills, textbox types
    table.rs          — table IR
  fonts/
    mod.rs            — font registration, metrics, fallback
    discovery.rs      — cross-platform font search, disk cache
    embed.rs          — font embedding, kern/GPOS extraction, subsetting
    encoding.rs       — WinAnsi encoding, glyph mapping
    cache.rs          — font index disk cache (TSV)
  geometry/
    mod.rs            — preset/custom shape evaluation entry point
    definitions.rs    — all 187 OOXML preset shape definitions (generated by tools/src/bin/generate_shapes.rs)
    text_warp_definitions.rs — the 40 WordArt text-warp presets (same generator)
    formulas.rs       — guide formula interpreter (adjustment values, trig, arithmetic)
    path.rs           — path command evaluation → ResolvedCommand list (MoveTo/LineTo/CubicTo/Close, PDF y-up)
  docx/
    mod.rs            — XML utils, relationships, parse orchestrator
    paragraph.rs      — paragraph XML → Paragraph (runs, numbering, drawings)
    styles.rs         — theme parsing, style structs, style inheritance
    runs.rs           — run-level XML → Vec<Run>
    numbering.rs      — list/numbering parsing, counter management
    images.rs         — image/chart drawing detection, floating/inline extraction
    charts.rs         — parse c:chartSpace XML → Chart model
    textbox.rs        — DrawingML + VML textbox parsing, shape fills (solid + gradient)
    tables.rs         — table + cell parsing
    smartart.rs       — SmartArt diagram parsing (dsp:drawing shape trees)
    group.rs          — drawing canvas (wpc) and shape group (wpg) flattening
    wordart.rs        — WordArt properties (wps:bodyPr warps, w14:textOutline)
    color.rs          — DrawingML colour resolution (srgbClr/schemeClr/prstClr)
    comments.rs       — comments part parsing
    emf.rs, wmf.rs    — minimal EMF/WMF metafile parsing
    embedded_fonts.rs — embedded font extraction and deobfuscation
    sections.rs       — section properties (page size, margins, columns)
    headers_footers.rs — header/footer/footnote XML parsing
    settings.rs       — document settings (tab stops, mirror margins, even/odd headers)
    alt_chunk.rs      — altChunk HTML content parsing
  pdf/
    mod.rs            — main render loop, behind-doc/body/foreground z-ordering
    assembly.rs       — page assembly: page borders, content streams, document-level objects
    layout.rs         — text layout, line building, paragraph rendering
    positioning.rs    — floating object placement and text wrap zones
    table_layout.rs   — table layout, auto-fit columns
    table.rs          — table rendering
    fonts.rs          — per-document font collection and registration
    images.rs         — image XObjects and picture effects
    textbox_render.rs — textbox rendering
    smartart.rs       — SmartArt shape rendering via geometry engine
    wordart.rs        — WordArt outlines and text warps
    charts.rs         — cartesian chart rendering (bar/line/area/scatter/bubble/radar)
    charts_radial.rs  — pie/doughnut chart rendering
    chart_legend.rs   — shared legend rendering
    header_footer.rs  — header/footer rendering
    footnotes.rs      — footnote rendering
    list_label.rs     — list label (bullet/number) rendering
    comments.rs       — comments pane rendering
    emf.rs            — EMF → PDF Form XObject
    color.rs, helpers.rs — fill/stroke colour and small drawing helpers
    tagging.rs        — tagged PDF (PDF/UA) structure tree
    objstm.rs         — compressed object streams post-pass
tests/
  suite/main.rs        — the one integration-test binary (`cargo test --test suite <module>`), modules:
    visual_comparison.rs — Jaccard/SSIM similarity test against Word-generated reference PDFs
    accessibility.rs     — PDF/UA-1 + structure-tree parity vs Word (needs verapdf, poppler)
    text_boundary.rs     — line start/end word metric
    font_validation.rs, image_count.rs, page_geometry.rs, file_size.rs — font/image/page-size/size parity
    libreoffice_comparison.rs — opt-in (DOCXIDE_LO_COMPARE=1) comparison with LibreOffice
    speed.rs             — conversion timing (#[ignore])
  common/              — shared harness (suite's `common` module via #[path]): mod.rs (fixtures, conversion, baselines, per-case score caches), metrics.rs (rasterization, Jaccard/SSIM; also tools/page-metrics), a11y.rs, text_boundary.rs
  baselines.json, visual_hashes.json — accepted scores and page hashes
  fixtures/<group>/<case>/ — input.docx + reference.pdf; groups: cases/ (79 handcrafted), fonts/ (7), hyphenation/ (8), scraped/ (164); SKIPLIST
  output/<group>/<case>/ — generated.pdf, reference/, generated/, diff/ screenshots (gitignored)
```

## Environment Variables

- `DOCXIDE_FONTS` — colon-separated (`;` on Windows) list of additional font directories to search (`.cargo/config.toml` points it at the repo's `fonts/` dir for cargo runs)
- `DOCXIDE_NO_FONT_CACHE` — set to any value to disable the font index disk cache (forces full scan every time)
- `DOCXIDE_DEBUG_WRAP` — set to draw the float wrap zones on the page
- Test harness: `DOCXIDE_CASE` (one fixture by name), `DOCXIDE_GROUP` (one fixture group), `DOCXIDE_IMAGES` (also save side-by-side comparison PNGs), `DOCXIDE_LO_COMPARE` (opt into the LibreOffice comparison), `DOCXIDE_A11Y_GEN` (which PDF in tests/output to score for accessibility), `LIBREOFFICE_PATH` (soffice binary)
- `DOCX_BIN` — CLI binary for `tools/engine_compare.py`

## Dependencies

- `zip = "8.1"` — DOCX is a ZIP container
- `roxmltree = "0.21"` — read-only XML tree parser
- `pdf-writer = "0.14"` — low-level PDF generation
- `ttf-parser = "0.25"`, `subsetter = "0.2"`, `memmap2` — font parsing, subsetting, memory-mapped font files
- `image = "0.25"` (png/bmp/jpeg/gif/tiff only), `png = "0.18"`, `miniz_oxide` — image decoding and recompression
- `unicode-linebreak` — line break opportunities; `log`; `clap` + `env_logger` behind the `cli` feature
- dev: `image` (png only), `rayon`, `serde`/`serde_json`, `sha2`, `env_logger`

## Development Notes

- Rust edition: 2024
- Test output is compared using **Jaccard similarity on ink pixels** (luma < 200 = ink) and **SSIM** with spatial tolerance (±8px). Run tests with `cargo test -- --nocapture` to see scores.
- Jaccard threshold: **20.5%**, SSIM threshold: **75%** (defined in `tests/suite/visual_comparison.rs`)
- Accessibility (`tests/suite/accessibility.rs`, `tests/common/a11y.rs`): `ua_deficit` (veraPDF PDF/UA-1 rules we fail worse than Word; 0 = as good as Word), `a11y_struct` (tag sequence vs Word's), `a11y_text` (block text in structure order vs Word's), `a11y_missing` (DOCX letters and digits that never reach our structure tree; needs no reference, any increase fails). Untagged references score N/A on the Word-relative three. See the Accessibility section of `roadmap.md`
- 79 handcrafted test cases (`tests/fixtures/cases/`) covering text, tables, images, charts, shapes, SmartArt, and more, plus the fonts/, hyphenation/ and scraped/ groups

## Working Alongside Other Sessions

Several agent sessions often run on this repo at once. **Before your first edit**, check for another active session: `git status` (uncommitted changes that are not yours) and `pgrep -fl "claude|cargo"`. If one is active, do your work in a worktree, never in the main checkout:

```bash
git worktree add .worktrees/<topic> -b <topic>
ln -s ../../fonts .worktrees/<topic>/fonts     # fonts/ is gitignored; without it every score collapses to the Arial fallback
cd .worktrees/<topic> && ./tools/run-tests.sh
git push origin <topic>:main                   # when done; leaves the shared checkout alone
```

Two suites in one checkout share `tests/output/`, each deleting and rewriting `latest_scores.json`: both reports are garbage. Never `git checkout`/`switch` branches in the main checkout either — the other session's edits silently move with you.

## Word Layout Learnings

### Default Font (Modern Word)
Modern Word (Office 365+) uses **Aptos 12pt** as the default Normal style font, not Calibri 11pt. Always parse `word/styles.xml` `docDefaults` for the actual defaults:
- `w:rPrDefault/w:rPr/w:sz @val` — font size in half-points (24 = 12pt)
- `w:pPrDefault/w:pPr/w:spacing @after` — default space after in twips (160 = 8pt)
- `w:pPrDefault/w:pPr/w:spacing @line @lineRule` — line spacing (278 auto = ~1.16× single)

### Baseline Y Position
The first line's baseline in PDF coordinates (from bottom):

```
baseline_y = slot_top - font_size
slot_top   = page_height - margin_top - space_before
```

This formula aligns cap-tops of different fonts (Helvetica, Aptos, etc.) at the same position, because cap height ≈ 0.72–0.94 × font_size and the formula compensates for this.

Measured for case1: reference baseline at y=708.72pt from bottom (83.28pt from top). Formula gives 708pt — 0.72pt difference, visually negligible.

### Document Grid
`w:sectPr/w:docGrid @w:linePitch` (in twips) defines the baseline-to-baseline distance for grid-snapped text. Divide by 20 to get points (360 twips = 18pt for case1).

### East Asian Line Height
Word lays out East Asian fonts (any face with CJK/Hangul/kana glyphs) at **1.3 × (hhea ascent + descent)** per line (the win metrics for most faces; Yu Gothic's hhea box is smaller, so 10.5pt steps 15.0, not 17.6), with no hhea lineGap and the extra leading split evenly above and below the glyphs (online export; local Mac Word put it all above). This is why 10.5pt SimSun gives the classic 15.6pt line, why 16pt Microsoft YaHei takes two cells of an 18pt grid, and why Yu Mincho lines double past 10.5pt. Exact-height boxes put the baseline at 0.8 × the height like Latin ones (online probes: KaiTi, SimSun, YaHei, Meiryo, MS Gothic/Mincho, Yu Gothic, Malgun); at-least boxes whose minimum wins bottom-align at winDescent. A run of only spaces uses the plain metrics (it must not raise a Latin line); empty paragraph marks, tabs and blank lines after a break keep the font's real metrics. Implemented in `src/fonts/embed.rs` (`compute_line_metrics`) and `src/pdf/layout.rs` (`run_line_metrics`).

### Missing CJK Fonts
Word substitutes by fontTable `w:charset` (hex: 80 Shift-JIS, 81 Hangul, 86 GB2312, 88 Big5) and `w:family` (roman → Batang/MS Mincho/SimSun/PMingLiU, otherwise Malgun Gothic/MS Gothic/Microsoft YaHei), rescuing single missing glyphs per character (kanji missing from Batang → MS Mincho). One platform-independent list in `src/fonts/mod.rs` (`cjk_fallback_fonts`) leads with the vendored Word fonts so CI and macOS agree.

### Case Browser Annotations

The case browser (`tools/src/bin/case_browser/`) supports click-to-annotate. All annotations are saved in a single file: `tests/output/annotations.json`. Each annotation includes the case name, page, source (Reference/Generated), PDF coordinates (x from left, y from bottom, in points), note text, and a `fixed` flag. The annotations capture rendering issues spotted during visual inspection — use them as precise bug reports when fixing rendering problems.

**Workflow for fixing issues from annotations:**
1. Read `tests/output/annotations.json` and filter for the case you're working on
2. The coordinates tell you exactly where on the page the issue is — cross-reference with the DOCX source using `docx-inspect` to find the relevant XML elements at that position
3. Use the overlay view (key `O`) in the case browser to see exactly what differs between reference and generated output at the annotated location
4. Annotations with `"fixed": true` have been resolved — focus on unfixed ones

**Keyboard shortcuts:** `A` toggle markers, `N` toggle notes panel, click image to add note, `Cmd+Enter` to save, `Escape` to cancel. The **Fix** button next to a note hands it to a headless `claude -p` run in its own worktree (`.worktrees/fix-<id>`, branch `fix/ann-<id>`, log in `tests/output/fix-<id>.log`), so several can run at once; the top bar shows each run, **Stop** kills one (with any cargo run it started, the worktree is kept), and when one exits the GUI mirrors the `fixed` flag Claude set in the worktree. Merge the branch and press Re-generate to see the result.

### pdf-writer Patterns
Object IDs must be written exactly once. Build the content stream first, then write the page object:

```rust
pdf.stream(content_id, &content.finish());
pdf.page(page_id)
    .media_box(...)
    .parent(pages_id)
    .contents(content_id)   // must come before .resources()
    .resources()
    .fonts()
    .pair(Name(b"F1"), font_id);
```
