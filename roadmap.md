# Roadmap

## Unicode Line Breaking (TODO — HIGH IMPACT)

We split text on whitespace only via `split_preserving_spaces()`. This fails for CJK (no spaces between words), Thai, and other scripts where break opportunities are Unicode-defined. Integrating `unicode-linebreak` crate would:
1. **Fix CJK line breaking** — break at correct positions without spaces
2. **Fix other scripts** — Thai, Khmer, Lao, Myanmar word boundaries
3. **Improve Latin handling** — proper break opportunities around hyphens, punctuation

Small integration effort, high impact for non-Latin documents.

## CJK Rendering Polish (TODO — MEDIUM IMPACT)

Core CJK support is implemented: CIDFont/Identity-H/ToUnicode encoding, platform-specific font fallback chains (Hiragino/Noto/Yu Gothic), per-character font fallback at render time, script-based run splitting via `w:rFonts @eastAsia`, and vertical text rendering (`w:textDirection="tbRlV"`). CJK fixtures render readable output but score low (4–9% Jaccard) due to spacing/positioning precision issues:

1. **`w:snapToGrid`** (DONE) — implemented. Paragraphs with `snap_to_grid: true` (spec default) snap line heights to grid pitch multiples when the section's `docGrid @type` is "lines", "linesAndChars", or "snapToChars". Paragraphs with `snapToGrid="0"` skip grid snapping. `DocGridType` enum and `grid_type` field added to `SectionProperties`.
2. **Vanished paragraph mark** (DONE) — `w:pPr/w:rPr/w:vanish` on the paragraph mark now produces zero height and zero spacing for empty paragraphs. Prevents phantom page breaks from trailing vanished paragraphs.
3. **`w:firstLineChars`** (MEDIUM) — character-based indent (e.g. `firstLineChars="100"` = 1 character width). Not parsed; we only handle `w:firstLine` (twip-based). In practice, twip fallback is always present alongside firstLineChars.
4. **Vertical text centering** — `render_vertical_cjk_cell` uses a simplistic height calculation (chars × font_size) that doesn't account for paragraph spacing, causing vertical misalignment in merged cells.

## Bundled Fallback Fonts (TODO — MEDIUM IMPACT)

We rely entirely on system fonts and fall back to Helvetica Type1 as a last resort. This produces inconsistent output across environments (servers, Docker, CI). Should bundle metric-compatible open fonts behind a feature flag:
- **Carlito** — metric-compatible with Calibri (the most common Word font)
- **Caladea** — metric-compatible with Cambria
- **Liberation Sans/Serif/Mono** — metric-compatible with Arial/Times New Roman/Courier New

Metric compatibility means identical advance widths, so layout stays correct even with substitution. Ensures consistent output without requiring specific system fonts.

## Paginator Extraction (TODO — MEDIUM IMPACT, HIGH ARCHITECTURAL VALUE)

The `render()` function in `pdf/mod.rs` mixes pagination with rendering. Extracting a dedicated pagination pass would:
1. **Enable widow/orphan control** — `w:widowControl` (default on) requires knowing whether ≥2 lines fit before committing a paragraph to the page. Currently we can't split paragraphs across pages line-by-line.
2. **Enable table header row repeat** — `w:tblHeader` marks rows that should repeat when a table breaks across pages. Requires pagination to know where the break falls.
3. **Enable keep-with-next / keep-lines** — `w:keepNext` and `w:keepLines` paragraph properties need look-ahead during pagination.
4. **Enable post-pagination field resolution** — PAGE/NUMPAGES fields could be resolved after layout instead of during rendering, which is cleaner and more correct.

Architecture: a `Paginator` takes the document model and produces `Vec<Page>` where each `Page` contains positioned elements. The PDF renderer then simply draws them. This is a significant refactor but unlocks multiple features that are impossible without it.

## TOC Internal Link Navigation (DONE)

`w:hyperlink w:anchor="name"` links in TOC entries now produce PDF GoTo annotations that jump to the target heading. `w:bookmarkStart w:name="..."` elements are collected during DOCX parsing and their page/y positions are tracked during rendering. The annotation writer emits `/S /GoTo` with `/XYZ` destinations for `#name` URLs and falls back to `/S /URI` for external links. Tested in case39.

## PDF Bookmarks (DONE)

PDF outline/bookmarks (sidebar navigation panel) are now generated from heading styles. Implementation tracks heading paragraphs via `w:outlineLvl` attributes in styles or paragraphs, builds a hierarchical outline tree using parent-child relationships from heading levels, and writes PDF Outline objects. The catalog sets `PageMode::UseOutlines` so the outline sidebar opens automatically. Tested in case39 (Introduction and Methods headings with proper nesting).

## PDF Metadata (DONE)

Document metadata (title, author, subject, keywords) is now parsed from `docProps/core.xml` and written to the PDF info dictionary. Producer is set to "docxside-pdf".

## Line Height: OS/2 Win Metrics (DONE)

OS/2 Win Metrics are already implemented in `compute_line_metrics()` in `src/fonts/embed.rs`. The function correctly uses `usWinAscent + usWinDescent` when `USE_TYPO_METRICS` is not set. The remaining vertical drift in failing text-only fixtures has other root causes — not line metrics.

## Vertical Drift Investigation (TODO — HIGH IMPACT)

The 8 failing text-only fixtures still show accumulated vertical shift. Investigation of three hypothesized root causes found:
1. **Image paragraph height rounding** — fixed: removed unconditional `line_pitch` floor for image paragraphs, using exact `content_height` instead.
2. **Table trailing spacing** — investigated and disproven: `render_table` already accounts for cell margins and borders, so `prev_space_after = 0.0` is approximately correct. Adding the last paragraph's `space_after` double-counts spacing.
3. **List paragraph spacing** — investigated and disproven for `polish_archery_range_plan`: the fixture uses manual numbering in tables, not `w:numPr` list formatting, and no styles define inter-paragraph spacing.

Remaining avenues to investigate:
- **Text wrapping around floating tables** — Word wraps body text around `tblpPr` tables, pushing subsequent paragraphs below. We render them as overlapping, causing large visual diffs (case32).
- **Per-font line height calibration** — different fonts may have subtle differences in how their Win Metrics translate to actual line spacing in Word.

## SmartArt Remaining Work

Basic fallback rendering via pre-flattened `dsp:drawing` shape trees is done, with full geometry engine support (all 187 preset shapes). Remaining:

1. **Group shapes** (MEDIUM EFFORT) — `dsp:grpSp` groups with nested transforms. Need recursive parsing.
2. **Connector shapes** (MEDIUM EFFORT) — `dsp:cxnSp` connectors between shapes (arrows, lines).
3. **Image shapes** (MEDIUM EFFORT) — shapes that contain embedded images (`a:blipFill`).
4. **Full layout engine** (VERY HIGH EFFORT) — implement the constraint-based layout algorithm that interprets ~200 XML layout recipes. Only needed for files that lack the `dsp:drawing` fallback. Not planned for the near term.

## Charts Remaining Work

All 8 chart types are supported (bar, line, pie, area, doughnut, radar, scatter, bubble). Remaining:

- **3D charts**: `c:bar3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` — not parsed
- **Stock charts**: `c:stockChart` — not parsed
- **Combo charts**: two chart types overlaid on the same plot area — not handled
- **Stacked bar rendering**: parsed but rendering treats as clustered
- **Data labels**: not parsed or rendered
- **Chart title**: not parsed or rendered
- **Secondary axes**: not handled
- **Chart label positioning**: axis labels still have small offsets vs Word. `text_width_approx` (len × fs × 0.5) is crude — real font metrics would help.
- **Legend placement fine-tuning**: small positional offsets vs Word. Centering formula and spacing need per-chart-type calibration.
- **Font selection in chart labels**: picks arbitrary font from seen_fonts, not theme font

## Track Changes Remaining Work

Final mode (insertions included, deletions removed) is done. Remaining:

- **Markup mode** — rendering deletions with red strikethrough, insertions with red underline (for documents exported with markup visible)
- **Paragraph-level changes** — `w:ins`/`w:del` wrapping entire `w:p` elements at `w:body` level
- **Property changes** — `w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tblPrChange` (formatting revisions)

## WordArt (MOSTLY DONE — LOW IMPACT)

WordArt appears in two forms: modern DrawingML (current Word) and legacy VML (older docs).

**Level 1 — Flat rendering (DONE):** Dedicated `src/docx/wordart.rs` and `src/pdf/wordart.rs` modules. Parses `fromWordArt`, `prstTxWarp`, `spAutoFit` from `bodyPr`. Parses `w14:textOutline` and `w14:textFill` from WML run properties. VML `v:textpath @string` fallback renders as flat text. Text outlines render via `TextRenderingMode::FillStroke`. Auto-fit (`spAutoFit`) computes textbox height from content.

**Level 2 — Text effects (DONE):** Solid `w14:textFill` color override applied during rendering. Text shadows (`w14:shadow`) parsed and rendered as offset pre-pass with blended shadow color. Text glow (`w14:glow`) parsed and rendered as thick stroke pre-pass with round joins. Gradient text fills parsed but rendering deferred (requires PDF clip-then-pattern).

**Level 3 — Two-path envelope warping (DONE):** All 40 `prstTxWarp` presets auto-generated from ECMA-376 spec. Glyph outlines extracted via `ttf_parser::OutlineBuilder` with correct bold/italic font variant selection. Warp algorithm: evaluate preset → sample top/bottom boundary curves → transform each glyph point through envelope interpolation → emit as filled PDF paths. Boundary sizing: natural text width horizontally (centered in textbox), shape height vertically. Text height normalized using `font_size × (ascender / glyph_extent)` to match Word's cap-height fill ratio. Shared helpers (`collect_text_info`, `emit_glyph_commands`, `fill_and_stroke_glyphs`) eliminate code duplication between renderers.

**Level 4 — Single-path text-on-a-path (DONE):** Arch/circle presets (`textArchUp`, `textArchDown`, `textCircle`) use a text-on-a-path algorithm: arc-length parameterize the single curve path, place each character along the arc with tangent-angle rotation, anchor text at `font_size / 2` from the path. Boundary sized to natural text advance (produces correct curvature). Case43 Jaccard improved from 22% → 50%.

**Level 5 — Legacy VML enhancement (TODO):** VML fill types (gradient/pattern), VML shadow, VML shapetype-to-prstTxWarp mapping. Basic flat rendering already done in Level 1.

## Unimplemented Spec Features

- **`w:tblLook` / `w:tblStylePr`** — table conditional formatting (firstRow, lastRow, firstCol, bands, etc.)
- **`w:jc val="distribute"`** — distribute alignment (equal spacing including edges), different from justified
- **`w:textDirection`** — text direction in table cells (btLr, tbRl)
- **`w:vAlign` on sectPr** — vertical alignment of text on the page (top/center/bottom/both)
- **Panose font matching** — fontTable.xml contains panose classification bytes; could use for more precise substitution

### Partially Implemented

- **Line spacing** — Auto and Exact work. AtLeast parsed but may not enforce minimum correctly.
- **Tab stops** — basic left/center/right tabs work but leader rendering and decimal alignment have precision issues.

## Code Structure

### Refactor `pdf/mod.rs` `render()` (see "Paginator Extraction")

The `render()` function in `pdf/mod.rs` is ~2400 lines with many closures and shared mutable state. The right fix is the paginator extraction described above. In the meantime, smaller extractions are possible:

- `pdf/headers_footers.rs` — `render_header_footer` (~220 lines, already a free fn)
- `pdf/footnotes.rs` — footnote height computation + rendering (~120 lines)
- `pdf/images.rs` — `embed_image` closure → free fn (~140 lines)
- `pdf/list_labels.rs` — `label_for_run`, `label_for_paragraph` (~30 lines)

## Performance

### Known Bottlenecks

- **Double font reads** — scan reads each font file for indexing, then `register_font` reads again for embedding. Keep the data from the first read.
- **Repeated WinAnsi conversion** — same text is converted in line-building, rendering, and table auto-fit. Pre-compute once and store in `WordChunk`.
- **String allocations** — `font_key()` allocates on every call; `WordChunk` clones font name strings per word. Use indices or interning.

### Parallelism (rayon)

- Font directory scanning — embarrassingly parallel, biggest win
- Font metric computation — parse face, compute widths per font independently
- Paragraph line wrapping — independent per paragraph once font metrics are ready
- ZIP decompression + XML parsing — read all entries into memory, parse in parallel

### Other

- Compress font file streams with FlateDecode (currently uncompressed)
- Memory usage for large DOCX files with many images

## Floating Image Wrapping

### Done

- **`lines_beside` bias**: `.round()` over-counted narrow lines when the last line barely overlapped the float zone bottom. Added -0.1 bias. Case41 Jaccard +5.7pp.
- **Self-wrapping**: Paragraphs wrap around their own floating images (float zone set before line building). Polygon data from `wp:wrapPolygon` parsed and used for per-line contour-aware exclusion. Self-wrapping now always triggers for paragraphs with floating images (previous float zone replaced, not blocked by pre-existing zone).
- **`wrapText` variants**: All four side-selection locations and the per-paragraph minimum-width gate now respect `wrapText`. Previously `wrapText` was parsed but completely ignored — all modes behaved as `Largest`.
  - `Left`: force text to left region only
  - `Right`: force text to right region only
  - `Largest`: pick the wider side (pre-existing default behavior)
  - `BothSides`: dual-region line builder fills left region first, overflows to right region on the same line. Combined-width threshold (both sides together must meet 72pt) instead of single-side threshold.
- **Both-sides wrapping**: `wrapText="bothSides"` flows text on both left and right of the image simultaneously. Implementation: `RightRegion` struct on `TextLine`, `per_line_dual` parameter on `build_paragraph_lines`, per-region alignment/justify in `render_paragraph_lines`. Works with all wrap types (Square, Tight, Through). Case42 (Mario): text wraps both sides. Case41 page 3 (centered image): text wraps both sides.

### Remaining

- **Remaining y-shift (page 2 only)**: Word places page 2's image (180×144pt) 14.8pt higher than all other images, despite identical `posOffset=0`. Pages 1,3,4,5,7 match perfectly (delta <0.02pt). Pages 2 and 6 (both cy=1828800/144pt) are the outliers. Likely Word snapping to grid/text boundaries based on image dimensions.
- **Look-back wrapping (TODO — MEDIUM IMPACT)**: Paragraphs BEFORE the image anchor cannot wrap beside the image because the float zone isn't set until the anchor paragraph renders. In Word, text from preceding paragraphs also wraps (e.g. case41 page 3 — the first paragraph's lower lines should wrap beside the centered image). Requires either a paginator or a two-pass layout with look-back: after rendering the image paragraph, go back and re-render preceding paragraphs that overlap the float zone.
- **Image in text paragraph**: Case41 page 6 — last line of text paragraph overlaps the image. Look-ahead only fires for the NEXT block, not same-paragraph floats.
- **Tight vs Through distinction**: Both currently use convex-hull polygon scanline (`poly_scanline` returns min/max x). For Through wrapping, text should fill polygon concavities. Requires returning per-line interval segments instead of hull bounds. Rare in practice — most polygons are convex.
- **Word-break precision**: BothSides wrapping produces correct structure but slightly different word breaks from Word, causing ~2pp Jaccard differences on case41. Likely due to font metric differences and rounding in the dual-region fill algorithm.

## Scraped Fixture Status

33 passing, 16 failing, 0 skipped out of 49 scraped fixtures. Breakdown of 16 failures by dominant issue:
- **text/layout only**: 8 fixtures (all show accumulated vertical shift from wrong line metrics)
- **anchored images**: 4 fixtures
- **floating tables**: 3 fixtures
- **structured doc tags**: 1 fixture

Run `./tools/target/debug/analyze-fixtures --failing` for current breakdown.

## Test Corpus Expansion

Additional fixture ideas not yet covered:
- Deep style inheritance (3+ level chains with run vs style vs paragraph conflicts)
- Nested tables (tables inside table cells)
- Table of Contents (right-aligned tabs + dot leaders + page field codes)
- Stacked bar chart rendering
- Charts with extreme data (50+ categories, very small/large values)
