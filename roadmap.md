# Roadmap

## Text Shaping with rustybuzz (TODO — HIGH IMPACT)

We do manual char→glyph mapping with no OpenType shaping. This means ligatures (fi, fl), contextual alternates, and complex scripts (Arabic, Indic) all render incorrectly. Integrating `rustybuzz` (pure-Rust HarfBuzz port) would:
1. **Fix ligatures** — OpenType GSUB table support for all Latin ligatures
2. **Fix complex scripts** — Arabic reordering/joining, Indic conjuncts, Thai marks
3. **Improve kerning** — GPOS kerning from shaping output replaces our manual kern table + GPOS extraction
4. **Fix CJK contextual forms** — proper glyph selection for CJK fonts

This is the single highest-impact improvement for international document fidelity. Would also simplify the font pipeline — shaping returns glyph IDs and advances directly, eliminating manual width computation.

Would also subsume the current "PDF rendering kerning" gap: we currently use `Tj` without kern adjustments in the PDF content stream. Shaping output would enable `TJ` arrays with per-glyph positioning for visually correct justified text.

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

## Line Height: OS/2 Win Metrics (TODO — MEDIUM, correct but causes regressions)

Word computes line height using OS/2 `usWinAscent + usWinDescent` when the font's `USE_TYPO_METRICS` flag is not set (most fonts). We currently use `hhea ascender - descender + line_gap`, which produces tighter line spacing. The fix is straightforward in `src/fonts/embed.rs` (`face.tables().os2` → `windows_ascender()`/`windows_descender()`; note `windows_descender()` returns a negated value so use `win_asc - win_desc`). However, changing this globally causes 23 regressions (some -50pp) because other layout code has been calibrated against the wrong `line_h_ratio`. Should be landed alongside a pass to fix compensating layout issues.

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

## Scraped Fixture Status

19 passing, 0 failing, 30 skipped (font issues) out of 39 scraped fixtures.
Run `./tools/target/debug/analyze-fixtures --failing --fonts` for current breakdown.

## Test Corpus Expansion

Additional fixture ideas not yet covered:
- Deep style inheritance (3+ level chains with run vs style vs paragraph conflicts)
- Nested tables (tables inside table cells)
- Table of Contents (right-aligned tabs + dot leaders + page field codes)
- Stacked bar chart rendering
- Charts with extreme data (50+ categories, very small/large values)
