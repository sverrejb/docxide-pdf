# Roadmap

## Kerning (DONE — conditional via `w:kern`)

Kern pairs extracted from both legacy `kern` table and GPOS PairAdjustment (Format1 + Format2) during font embedding. Applied conditionally in `word_width()` when `w:kern` threshold is met. Parsed from `docDefaults`, paragraph/character styles (with `basedOn` inheritance), and inline run properties.

Results: case1 +14.4pp, case18 +13.6pp, zero regressions.

Remaining:
- **PDF rendering kerning**: currently kerning only affects line breaking (text measurement). Rendering still uses `Tj` without kern adjustments. Adding `TJ` arrays with positioning would improve visual quality for justified text.
- **`enableOpenTypeFeatures`**: this compat setting controls ligatures/contextual alternates, NOT kerning. Investigated and confirmed — all test documents have it enabled but Word does not use it to trigger kerning.

## The Mongolian Case

The `mongolian_human_rights_law` scraped fixture scores 13.5% Jaccard (needs 20%). The `w:dstrike` in rPrDefault is now correctly inherited and overridden by Normal style's `w:dstrike val="0"` via the full style chain. Previously improved from 13.6% via:
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
- **Bundle fallback fonts** — see "Bundled Fallback Fonts" section below
- **CJK fallback** — see "CJK Font Support" section below

## Text Shaping with rustybuzz (TODO — HIGH IMPACT)

We do manual char→glyph mapping with no OpenType shaping. This means ligatures (fi, fl), contextual alternates, and complex scripts (Arabic, Indic) all render incorrectly. Integrating `rustybuzz` (pure-Rust HarfBuzz port) would:
1. **Fix ligatures** — OpenType GSUB table support for all Latin ligatures
2. **Fix complex scripts** — Arabic reordering/joining, Indic conjuncts, Thai marks
3. **Improve kerning** — GPOS kerning from shaping output replaces our manual kern table + GPOS extraction
4. **Fix CJK contextual forms** — proper glyph selection for CJK fonts

This is the single highest-impact improvement for international document fidelity. Would also simplify the font pipeline — shaping returns glyph IDs and advances directly, eliminating manual width computation.

## Unicode Line Breaking (TODO — HIGH IMPACT)

We split text on whitespace only via `split_preserving_spaces()`. This fails for CJK (no spaces between words), Thai, and other scripts where break opportunities are Unicode-defined. Integrating `unicode-linebreak` crate would:
1. **Fix CJK line breaking** — break at correct positions without spaces
2. **Fix other scripts** — Thai, Khmer, Lao, Myanmar word boundaries
3. **Improve Latin handling** — proper break opportunities around hyphens, punctuation

Small integration effort, high impact for non-Latin documents.

## CJK Font Support (TODO — HIGH IMPACT)

Japanese/Chinese/Korean text renders with completely wrong metrics or as blanks. The `japanese_interlibrary_loan` scraped fixture (1-page form) scores only 3.2% Jaccard because CJK glyphs fall back to Latin fonts with wrong character widths, causing massive text displacement in tables. Needs:
1. **CJK font fallback chain** — detect CJK codepoints, fall back to system CJK fonts (Hiragino Sans on macOS, Noto Sans CJK on Linux)
2. **CJK-aware text measurement** — full-width characters need correct advance widths
3. **CJK encoding in PDF** — CJK fonts require CIDFont/ToUnicode mapping, not WinAnsiEncoding

Would unblock all CJK-language documents (Japanese, Chinese, Korean).

## Bundled Fallback Fonts (TODO — MEDIUM IMPACT)

We rely entirely on system fonts and fall back to Helvetica Type1 as a last resort. This produces inconsistent output across environments (servers, Docker, CI). Should bundle metric-compatible open fonts behind a feature flag:
- **Carlito** — metric-compatible with Calibri (the most common Word font)
- **Caladea** — metric-compatible with Cambria
- **Liberation Sans/Serif/Mono** — metric-compatible with Arial/Times New Roman/Courier New

Metric compatibility means identical advance widths, so layout stays correct even with substitution. Ensures consistent output without requiring specific system fonts.

## docDefaults Run Properties (DONE)

`StyleDefaults` now carries all run-level properties from `rPrDefault/rPr` (bold, italic, caps, smallCaps, vanish, strikethrough, dstrike, underline, color, char_spacing). Previously only font_size, font_name, and kern_threshold were parsed — all other properties silently defaulted to false/none. `parse_runs()` now falls back to these defaults instead of hardcoded `false`/`0.0`. `ParagraphStyle` also carries `underline`, `strikethrough`, `dstrike`, and `char_spacing` with full `basedOn` inheritance, completing the style chain: direct rPr → character style → paragraph style → docDefaults.

Impact: `italian_project_proposal` improved from 7.2% → 10.9% Jaccard (entire document defaults to smallCaps). `mongolian_human_rights_law` correctly inherits dstrike from docDefaults and overrides it via Normal style's `w:dstrike val="0"`.

## Scraped Fixture Improvements

14 passing, ~16 failing out of ~30 non-skipped scraped fixtures.
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

## Track Changes (DONE — final mode)

`w:ins` (insertions) and `w:del` (deletions) are now handled in `collect_run_nodes()` in `src/docx/runs.rs`. Final mode: inserted content is included as normal text, deleted content is skipped entirely. This matches Word's "Accept All Changes" / PDF export behavior.

Remaining:
- **Markup mode** — rendering deletions with red strikethrough, insertions with red underline (for documents exported with markup visible)
- **Paragraph-level changes** — `w:ins`/`w:del` wrapping entire `w:p` elements at `w:body` level (not seen in test fixtures yet)
- **Property changes** — `w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tblPrChange` (formatting revisions)

## Field Code Display Results (DONE)

Text between `fldChar separate` and `fldChar end` is now rendered for non-dynamic fields (DocProperty, etc.). Previously all text inside field regions was suppressed. Dynamic fields (PAGE, NUMPAGES, STYLEREF) suppress cached result text and use computed values at render time. The cached text is still stored in the field_code run via `field_result_text` accumulator for body text display. Affects 4 scraped fixtures.

## Even/Odd Headers & STYLEREF Resolution (DONE)

Even-page headers/footers (`w:evenAndOddHeaders` + `type="even"` references) are now parsed and rendered. Section-relative page numbers are computed when `pgNumType w:start` is explicitly set. STYLEREF fields in headers/footers resolve per spec §17.16.5.59: search current page top-to-bottom first, then backward through previous pages. Tracked via per-page `styleref_running` and `styleref_page_first` maps. bush_fires_act_comparison improved from 8.4% → 8.7% Jaccard, 20.5% → 21.0% SSIM.

## Document Settings (DONE — `word/settings.xml`)

New module `src/docx/settings.rs` parses `word/settings.xml` into `DocumentSettings`:
- `even_and_odd_headers` — different headers for even/odd pages
- `default_tab_stop` — default tab stop interval in points
- `mirror_margins` — mirror margins for book-style layout

`even_and_odd_headers` combined with `pgNumType w:start` triggers blank page insertion for odd/even page alignment in section breaks.

## Odd/Even Page Section Breaks (DONE)

`SectionBreakType::OddPage`/`EvenPage` now insert blank pages when the next physical page has wrong parity. Also handles implicit odd-page alignment when `evenAndOddHeaders` is enabled and the section has an explicit `pgNumType w:start`.

## Paginator Extraction (TODO — MEDIUM IMPACT, HIGH ARCHITECTURAL VALUE)

The `render()` function in `pdf/mod.rs` mixes pagination with rendering. Extracting a dedicated pagination pass would:
1. **Enable widow/orphan control** — `w:widowControl` (default on) requires knowing whether ≥2 lines fit before committing a paragraph to the page. Currently we can't split paragraphs across pages line-by-line.
2. **Enable table header row repeat** — `w:tblHeader` marks rows that should repeat when a table breaks across pages. Requires pagination to know where the break falls.
3. **Enable keep-with-next / keep-lines** — `w:keepNext` and `w:keepLines` paragraph properties need look-ahead during pagination.
4. **Enable post-pagination field resolution** — PAGE/NUMPAGES fields could be resolved after layout instead of during rendering, which is cleaner and more correct.

Architecture: a `Paginator` takes the document model and produces `Vec<Page>` where each `Page` contains positioned elements. The PDF renderer then simply draws them. This is a significant refactor but unlocks multiple features that are impossible without it.

## PDF Bookmarks (TODO — SMALL EFFORT, MEDIUM IMPACT)

We don't generate PDF outline/bookmarks from heading styles. This is a commonly expected feature — most PDF viewers show a sidebar navigation panel from the outline. Implementation:
1. During render, track heading paragraphs (style with `w:outlineLvl` or Heading1-9 style names) with their page index and y-position
2. Build a hierarchical outline tree from heading levels
3. Write PDF Outline objects with `pdf-writer`'s outline API

Small effort, high perceived quality improvement.

## PDF Metadata (TODO — TRIVIAL EFFORT, MEDIUM IMPACT)

We don't write document metadata (title, author, subject, keywords) to the PDF. DOCX stores this in `docProps/core.xml` (Dublin Core). Implementation:
1. Parse `docProps/core.xml` during DOCX loading — extract `dc:title`, `dc:creator`, `dc:subject`, `cp:keywords`
2. Write PDF document info dictionary via `pdf-writer`

Trivial effort, improves PDF viewer display (title in tab/title bar instead of filename).

## DrawingML Preset Geometry (TODO — MEDIUM EFFORT, MEDIUM IMPACT)

We currently handle only a handful of preset shapes (`ellipse`, `rect`, `notchedRightArrow`, `line`, `arc`). OOXML defines **187 preset shapes** (`ST_ShapeType`), each with guide formulas that compute the actual path from adjustment values. Any shape we don't recognize falls back to a rectangle.

### Current state

Basic connector/arc rendering added for the vaccine fixture (lines and arcs forming letters on circles). Arc angle conversion from OOXML to PDF coordinates works but doesn't handle ellipse rotation transforms properly (only angle-shifting, not true rotation).

### No existing Rust crate helps with the hard part

- **[msoffice_shared](https://docs.rs/msoffice_shared/0.1.1/msoffice_shared/drawingml/index.html)** — Has DrawingML type definitions (enums, structs) but doesn't compute actual paths. Parsing only, v0.1.1 (2020).
- **[ooxmlsdk](https://github.com/KaiserY/ooxmlsdk)** — .NET Open XML SDK port. Document read/write, not rendering.
- **[ooxml](https://lib.rs/crates/ooxml)** — xlsx only.

None provide a function that takes `prst` name + adjustment values → path commands. The geometry formulas are defined in the OOXML spec Part 4 and would need to be implemented from scratch or ported from LibreOffice's C++ implementation.

### Incremental path

1. **Add common shapes** (LOW EFFORT) — roundRect, diamond, chevron, pentagon, hexagon, triangle, etc. Hand-code the most frequent ~20 shapes.
2. **Implement guide formula interpreter** (MEDIUM EFFORT) — parse the `a:gd` formula language (`val`, `*/`, `+-`, `sin`, `cos`, `at2`, etc.) to compute paths from adjustment values. This unlocks all 187 shapes at once.
3. **Custom geometry** (`a:custGeom`) (MEDIUM EFFORT) — shapes defined with explicit path commands rather than presets. Already appears in some documents.

## Unimplemented Spec Features

- **`w:tblLook` / `w:tblStylePr`** — table conditional formatting (firstRow, lastRow, firstCol, bands, etc.)
- **`w:jc val="distribute"`** — distribute alignment (equal spacing including edges), different from justified
- **`w:textDirection`** — text direction in table cells (btLr, tbRl)
- **`w:vAlign` on sectPr** — vertical alignment of text on the page (top/center/bottom/both)

### Header Height Trailing space_after (DONE)

`compute_header_height` now includes the last header paragraph's `space_after` in the returned height. Previously this trailing spacing was excluded, causing `effective_slot_top` to be too high when the header constrains body text positioning. Improved croatian_grant_guidelines page 1 alignment (+0.1pp Jaccard, +0.2pp SSIM). Pages 2+ offset remains (~2-3pt) — caused by the OS/2 Win Metrics issue below.

### Line Height: OS/2 Win Metrics (MEDIUM — correct but causes regressions)

Word computes line height using OS/2 `usWinAscent + usWinDescent` when the font's `USE_TYPO_METRICS` flag is not set (most fonts). We currently use `hhea ascender - descender + line_gap`, which produces tighter line spacing. The fix is straightforward in `src/fonts/embed.rs` (`face.tables().os2` → `windows_ascender()`/`windows_descender()`; note `windows_descender()` returns a negated value so use `win_asc - win_desc`). However, changing this globally causes 23 regressions (some -50pp) because other layout code has been calibrated against the wrong `line_h_ratio`. Should be landed alongside a pass to fix compensating layout issues.

### Partially Implemented

- **Line spacing** — Auto and Exact work. AtLeast parsed but may not enforce minimum correctly.
- **Tab stops** — basic left/center/right tabs work but leader rendering and decimal alignment have precision issues.

## Code Structure

### Refactor `pdf/mod.rs` `render()` (LOW → see "Paginator Extraction")

The `render()` function in `pdf/mod.rs` is ~1450 lines with many closures and shared mutable state (`y`, `current_page`, `effective_margin_bottom`, etc.). The right fix is the paginator extraction described above — separating pagination from rendering. In the meantime, smaller extractions are possible:

- `pdf/headers_footers.rs` — `render_header_footer` (~220 lines, already a free fn)
- `pdf/footnotes.rs` — footnote height computation + rendering (~120 lines)
- `pdf/images.rs` — `embed_image` closure → free fn (~140 lines)
- `pdf/list_labels.rs` — `label_for_run`, `label_for_paragraph` (~30 lines)

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

## SmartArt (IN PROGRESS — basic fallback rendering; full support is VERY HIGH EFFORT)

### Current state

Basic rendering works by reading the pre-flattened `dsp:drawing` shape tree from `word/diagrams/drawingN.xml`. This contains positioned shapes (rect, ellipse, notchedRightArrow) with fills, strokes, and text — essentially a snapshot that Word pre-computes from the layout engine. Parsing in `src/docx/smartart.rs`, rendering in `src/pdf/smartart.rs`.

### Scale of the full feature

SmartArt is one of the largest features in OOXML. Microsoft ships **~190 built-in layouts** across 9 categories:

| Category | Layouts |
|---|---|
| Process | 42 |
| Relationship | 35 |
| Picture | 34 |
| List | 31 |
| Cycle | 16 |
| Hierarchy | 15 |
| Office.com | 9 |
| Matrix | 4 |
| Pyramid | 4 |

Each SmartArt diagram in a DOCX consists of **4 XML parts**: data model (`dgm:dataModel`), layout definition (`dgm:layoutDef`), colors (`dgm:colorsDef`), and style (`dgm:styleDef`).

### Why it's hard

The layout definitions are a **mini programming language**. Each layout is a tree of `layoutNode` elements with an algorithm type that controls positioning:

- `composite` — absolute positioning of children
- `lin` — linear flow
- `snake` — wrapping linear flow (multi-row/column)
- `cycle` — circular arrangement
- `hierRoot` / `hierChild` — org chart trees
- `pyra` — pyramid/trapezoid stacking
- `conn` — connector routing
- `sp` — spacing
- `tx` — text fitting

Each algorithm has dozens of parameters. The OOXML spec calls layout "the single largest aspect of DrawingML."

**No fallback is guaranteed** — Word does NOT always include pre-rendered shapes in the DOCX. The `dsp:drawing` part exists in many files but is not mandatory. Files without it require a full layout engine to display anything.

LibreOffice spent years implementing SmartArt import and describes their support as working "nearly perfectly" for ~200 layouts, with editing still "experimental only."

### Incremental path

1. **Current: `dsp:drawing` fallback** (DONE) — renders pre-flattened shapes when present. Covers many real-world files since Word typically includes this part.
2. **More shape types** (DONE) — geometry engine supports all 187 OOXML preset shapes via formula interpreter.
3. **Gradient fills** (LOW EFFORT) — SmartArt shapes frequently use linear gradients; extend fill parsing.
4. **Group shapes** (MEDIUM EFFORT) — `dsp:grpSp` groups with nested transforms. Need recursive parsing.
5. **Connector shapes** (MEDIUM EFFORT) — `dsp:cxnSp` connectors between shapes (arrows, lines).
6. **Image shapes** (MEDIUM EFFORT) — shapes that contain embedded images (`a:blipFill`).
7. **Full layout engine** (VERY HIGH EFFORT) — implement the constraint-based layout algorithm that interprets ~200 XML layout recipes. This is essentially building a small layout engine from scratch. Only needed for files that lack the `dsp:drawing` fallback.

### Not planned

- **SmartArt editing/creation** — out of scope for a PDF converter
- **Office.com layout download** — dynamic layouts fetched from Microsoft's servers

## Test Corpus Expansion

Additional fixture ideas not yet covered:
- Deep style inheritance (3+ level chains with run vs style vs paragraph conflicts)
- Hyperlinks and bookmarks
- Nested/multi-level lists (outline numbering: `1. → a. → i. → •`)
- Nested tables (tables inside table cells)
- Table of Contents (right-aligned tabs + dot leaders + page field codes)
- Stacked bar chart rendering
- Charts with extreme data (50+ categories, very small/large values)
