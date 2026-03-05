# Roadmap

## Kerning (DONE — conditional via `w:kern`)

Kern pairs extracted from both legacy `kern` table and GPOS PairAdjustment (Format1 + Format2) during font embedding. Applied conditionally in `word_width()` when `w:kern` threshold is met. Parsed from `docDefaults`, paragraph/character styles (with `basedOn` inheritance), and inline run properties.

Results: case1 +14.4pp, case18 +13.6pp, zero regressions.

Remaining:
- **PDF rendering kerning**: currently kerning only affects line breaking (text measurement). Rendering still uses `Tj` without kern adjustments. Adding `TJ` arrays with positioning would improve visual quality for justified text.
- **`enableOpenTypeFeatures`**: this compat setting controls ligatures/contextual alternates, NOT kerning. Investigated and confirmed — all test documents have it enabled but Word does not use it to trigger kerning.

## The Mongolian Case

The `mongolian_human_rights_law` scraped fixture scores 17.9% Jaccard (needs 20%). Improved from 13.6% via:
1. **"Standard" style recognition** (DONE) — LibreOffice exports its default paragraph style as a custom `w:customStyle="1"` style named "Standard". When the document's `docDefaults` lacks `w:kern`, we now merge `kern_threshold` from "Standard" if present. Found in 4/39 scraped fixtures; 2 of those carry `w:kern val="3"`.
2. **Multi-space preservation** (DONE) — `build_paragraph_lines` previously used `split_whitespace()` which collapsed consecutive spaces to single gaps. Now uses `split_preserving_spaces()` which counts actual space characters between words and accumulates space width across runs. Fixed the date line (66 consecutive spaces used for positioning) wrapping correctly.

Remaining gap to 20%: page 1 scores 47.6% but pages 2-8 score 7-13% due to cascading vertical shifts from multiple small differences: title formatting (missing space in "МОНГОЛ УЛСЫНХУУЛЬ"), header image positioning, paragraph spacing precision, and font metrics differences between Word and our rendering.

## Font Substitution (DONE — fontTable.xml altName + family fallback)

Parses `word/fontTable.xml` metadata (`w:altName`, `w:family`) into `FontTable` on the `Document` model. When a font isn't found via the semicolon-split candidate list or system font index, the substitution chain tries:
1. `altName` from fontTable.xml (e.g. "Liberation Serif" → "Times New Roman")
2. Family-class fallback: roman→"Times New Roman", swiss→"Arial", modern→"Courier New"
3. Only then falls back to Helvetica

Also added `w:hAnsi`/`w:hAnsiTheme` fallback in `resolve_font_from_node()` for documents that only specify hAnsi font variants.

Remaining:
- **Panose matching** — fontTable.xml also contains panose classification bytes; could use these for more precise substitution
- **Bundle fallback fonts** (Liberation, Noto) for consistent output without system fonts
- **CJK fallback** — CJK characters render as blanks; need fallback to system CJK fonts (Hiragino on macOS, Noto CJK on Linux)

## Scraped Fixture Improvements

8 passing, ~22 failing out of ~30 non-skipped scraped fixtures.
Run `./tools/target/debug/analyze-fixtures --failing --fonts` for current breakdown.

### Floating Tables (DONE — positioning; inline `w:tblBorders` DONE)

Floating table positioning (`w:tblpPr`) was already implemented. Inline `w:tblBorders` (borders specified directly on `w:tblPr` rather than via a named `w:tblStyle`) are now parsed and merged with style borders (inline overrides style). Test case32 covers floating table + inline borders. Affected scraped fixtures (`italian_project_proposal`, `polish_municipal_letter`, etc.) still score below 20% Jaccard due to other gaps (font metrics, complex layout).

### Textbox / Shape Rendering (DONE — fills, margins, header z-order)

DrawingML textboxes (`wps:txbx` → `w:txbxContent`) and VML fallback (`v:textbox`) render text content at the correct anchor position. Shape fills (`a:solidFill` with `a:srgbClr` and `a:schemeClr` theme colors including lumMod/lumOff modifiers) render as filled rectangles. Textbox body margins (`wps:bodyPr` lIns/tIns/rIns/bIns) are respected. Header/footer content renders behind body content via content stream prepending (correct z-order). Floating images render after textbox shapes for correct layering (images on top of fills). c9211737 scores 91.5% Jaccard; 5811dabc/d0252e2f/f271d69a remain skipped (page count mismatches / font issues). Remaining gaps: text wrapping around textboxes, clipping to bounding box, shape borders/outlines, proper z-index interleaving of shapes and images.

### Anchored Image Positioning (DONE — all wrap modes)

`wp:anchor` images are now positioned absolutely regardless of wrap mode (wrapNone, wrapTight, wrapSquare, etc.). Previously only `wrapNone` anchors were treated as floating; all others fell through to inline rendering. `compute_drawing_info()` now skips all anchors (parse_runs handles them), preventing image duplication. Remaining gap: text wrapping around anchored images is not implemented (content flows through/behind images).

### Tab Stop Line Wrapping (MEDIUM — causes 42.5% SSIM on czech_tree_cutting_permit)

When many consecutive tabs cause content to overflow a line, Word wraps the content to the next line. Our renderer fails to handle this — tab-wrapped content is lost or pushed off-page. This is the primary cause of the worst SSIM score among non-skipped fixtures. Affects tab-heavy form-layout documents.

### Tab Stop Precision (LOW)

Tab stop alignment and leader rendering has small positioning errors that accumulate in tab-heavy documents (e.g. table of contents). Header tab stops (center/right) also need proper handling.

## Charts (DONE — Bar, Line, Pie, Area, Doughnut, Radar, Scatter, Bubble)

Inline charts parsed from DrawingML chart parts (`word/charts/chartN.xml`). Detected via `a:graphicData` URI in `images.rs`, parsed in `docx/charts.rs`, rendered in `pdf/charts.rs` (radial charts in `pdf/charts_radial.rs`).

Supported chart types: `c:barChart` (vertical/horizontal, clustered/stacked), `c:lineChart`, `c:pieChart`/`c:pie3DChart`, `c:areaChart`, `c:doughnutChart`, `c:radarChart`, `c:scatterChart`, `c:bubbleChart`. Series data, category labels, axis config, legend, plot borders all extracted.

Rendering: bar rects, line smooth Catmull-Rom curves with per-series markers, pie/doughnut polygon wedges with theme accent colors, area filled polygons, radar polygons with concentric gridlines, scatter/bubble point markers. Content-aware margins, nice tick steps, gridlines, axis labels, legend (right/bottom).

Test fixtures: case29 (4 bar chart variations), case30 (line + pie + area).

Remaining:
- **Edge cases**: charts with very many bars/sectors (50+), verify rendering doesn't break
- **3D charts**: `c:bar3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` — not parsed
- **Stock charts**: `c:stockChart` — not parsed
- **Combo charts**: two chart types overlaid on the same plot area — not handled
- **Stacked bar rendering**: parsed but rendering treats as clustered
- **Data labels on chart**: not parsed or rendered
- **Chart title**: not parsed or rendered
- **Secondary axes**: not handled
- **Radar axis auto-scale** (DONE): headroom threshold changed from 0.9 to 0.98 to match scatter charts, fixing axis max 12→10 for case31 data.
- **Bubble chart legend markers** (DONE): bubble chart legend now renders circles instead of diamond/square.
- **Radar chart "0" label** (DONE): renders "0" at center of radar chart axis.
- **Radar legend line+marker style** (DONE): radar chart legend now draws line segments through markers, matching Word's style.
- **Bubble chart fill alpha** (DONE): parses `a:alpha` from series `c:spPr`, renders via PDF ExtGState `/ca` fill opacity. Bubble chart stroke outlines also rendered.
- **Axis tick marks** (DONE): outward tick marks rendered on both axes for all cartesian chart types.
- case31 SSIM improved from 64.1% → 76.1%. Radar chart: pentagon size, label placement (angle-based continuous formula), value label gap, legend stroke colors, legend line length all tuned. Fixture colors updated to match python-docx theme accents.
- **Chart label positioning**: axis labels on all case31 charts (scatter, doughnut, radar, bubble) still have small offsets vs Word. `text_width_approx` (len × fs × 0.5) is crude — real font metrics would help. Revisit per-chart-type label placement.
- **Legend placement fine-tuning**: pie and line/bar chart legends have small positional offsets vs Word (few pt). Centering formula and spacing need per-chart-type calibration.
- **Font selection in chart labels**: picks arbitrary font from seen_fonts, not theme font

## Unimplemented Spec Features

- **`w:tblLook` / `w:tblStylePr`** — table conditional formatting (firstRow, lastRow, firstCol, bands, etc.)
- **`w:jc val="distribute"`** — distribute alignment (equal spacing including edges), different from justified
- **`w:textDirection`** — text direction in table cells (btLr, tbRl)
- **`w:vAlign` on sectPr** — vertical alignment of text on the page (top/center/bottom/both)

### Line Height: OS/2 Win Metrics (MEDIUM — correct but causes regressions)

Word computes line height using OS/2 `usWinAscent + usWinDescent` when the font's `USE_TYPO_METRICS` flag is not set (most fonts). We currently use `hhea ascender - descender + line_gap`, which produces tighter line spacing. The fix is straightforward in `src/fonts/embed.rs` (`face.tables().os2` → `windows_ascender()`/`windows_descender()`; note `windows_descender()` returns a negated value so use `win_asc - win_desc`). However, changing this globally causes 23 regressions (some -50pp) because other layout code has been calibrated against the wrong `line_h_ratio`. Should be landed alongside a pass to fix compensating layout issues.

### Partially Implemented

- **Line spacing** — Auto and Exact work. AtLeast parsed but may not enforce minimum correctly.
- **Tab stops** — basic left/center/right tabs work but leader rendering and decimal alignment have precision issues.

## Code Structure

### Refactor `pdf/mod.rs` `render()` (LOW)

The `render()` function in `pdf/mod.rs` is ~1450 lines with many closures and shared mutable state (`y`, `current_page`, `effective_margin_bottom`, etc.). It could benefit from extraction into submodules:

- `pdf/headers_footers.rs` — `render_header_footer` (~220 lines, already a free fn)
- `pdf/footnotes.rs` — footnote height computation + rendering (~120 lines)
- `pdf/images.rs` — `embed_image` closure → free fn (~140 lines)
- `pdf/list_labels.rs` — `label_for_run`, `label_for_paragraph` (~30 lines)

The core page loop is tightly coupled through shared state, so breaking it into phases would require introducing a render context struct — a bigger undertaking that should be done incrementally.

## Performance

### Known Bottlenecks

- **Double font reads** — scan reads each font file for indexing, then `register_font` reads again for embedding. Keep the data from the first read.
- **Repeated WinAnsi conversion** — same text is converted in line-building, rendering, and table auto-fit. Pre-compute once and store in `WordChunk`.
- **String allocations** — `font_key()` allocates on every call; `WordChunk` clones font name strings per word. Use indices or interning.

### Profiling Setup

- Add phase timing to parse/render split
- Add Criterion benchmarks (full pipeline, parse-only, render-only, font scan)
- Use `samply` for flamegraph profiling

### Parallelism (rayon)

- Font directory scanning — embarrassingly parallel, biggest win
- Font metric computation — parse face, compute widths per font independently
- Paragraph line wrapping — independent per paragraph once font metrics are ready
- ZIP decompression + XML parsing — read all entries into memory, parse in parallel

### Other

- Compress font file streams with FlateDecode (currently uncompressed)
- Memory usage for large DOCX files with many images

## Test Corpus Expansion

Additional fixture ideas not yet covered:
- Deep style inheritance (3+ level chains with run vs style vs paragraph conflicts)
- Hyperlinks and bookmarks
- Nested/multi-level lists (outline numbering: `1. → a. → i. → •`)
- Nested tables (tables inside table cells)
- Table of Contents (right-aligned tabs + dot leaders + page field codes)
- Stacked bar chart rendering
- Charts with extreme data (50+ categories, very small/large values)
