# Roadmap

## Kerning / GPOS (PARTIALLY DONE)

Render-only GPOS kerning is implemented via `rustybuzz`: per-word shaping extracts GPOS kern adjustments and emits TJ operators with per-glyph positioning. Line breaking still uses unkerned widths.

**Why render-only**: Full kerning (layout + rendering) was tested but caused widespread line-break regressions. Investigation revealed Word's kerning values differ from standard GPOS — Word kerns pairs (e.g. e→l, l→o in Aptos) that have no GPOS table entry at all. The source of Word's extra kerning is unknown (possibly proprietary heuristics, or the malformed legacy kern table).

Remaining work:
1. **Match Word's kerning source** — investigate where Word gets kerning for pairs absent from the GPOS table; once understood, apply to both layout and rendering
2. **Cross-word kerning** — currently shape per-word; shaping full runs would capture comma→space and space→W adjustments that Word includes

## Font Substitution (HIGH)

When a font isn't found, we fall back unconditionally to Helvetica (built-in PDF Type1, no TrueType metrics). This causes two problems:
1. **Wrong font family** — serif fonts get Helvetica instead of a serif fallback
2. **No TrueType metrics** — forces `font_size * 1.2` line height estimate, causing layout drift across pages

Implementation ideas (increasing quality):
- Map font family → system default (serif→Times New Roman, sans→Arial, mono→Courier New)
- Parse panose classification from `fontTable.xml` and match against installed fonts
- Bundle fallback fonts (Liberation, Noto) for consistent output without system fonts

Related: CJK characters render as blanks — need fallback to system CJK fonts (Hiragino on macOS, Noto CJK on Linux).

## Scraped Fixture Improvements

9 passing, ~25 failing, 10 skipped (font issues) out of ~44 scraped fixtures.
Run `./tools/target/debug/analyze-fixtures --failing --fonts` for current breakdown.

### Floating Tables (MEDIUM — 5 failing)

Tables with `w:tblpPr` positioning attributes render as normal flow tables instead of being positioned absolutely. Existing table rendering can be reused; only the positioning pass needs to be added.

### Textbox Rendering (MEDIUM — 3 failing)

VML textboxes (`v:textbox`, `w:txbxContent`) and DrawingML `wps:txbx` content are unhandled. Content inside `w:txbxContent` is regular WordprocessingML so existing rendering can be reused in a clipped bounding box.

### Anchored Image Positioning (MEDIUM — 10 failing w/ anchors)

`wp:anchor` images lack proper positioning (horizontal/vertical offsets relative to page/column/margin) and text wrapping. Currently rendered inline.

### Tab Stop Precision (LOW)

Tab stop alignment and leader rendering has small positioning errors that accumulate in tab-heavy documents (e.g. table of contents). Header tab stops (center/right) also need proper handling.

## Unimplemented Spec Features

- **`w:tblLook` / `w:tblStylePr`** — table conditional formatting (firstRow, lastRow, firstCol, bands, etc.)
- **`w:jc val="distribute"`** — distribute alignment (equal spacing including edges), different from justified
- **`w:textDirection`** — text direction in table cells (btLr, tbRl)
- **`w:vAlign` on sectPr** — vertical alignment of text on the page (top/center/bottom/both)

### Partially Implemented

- **Line spacing** — Auto and Exact work. AtLeast parsed but may not enforce minimum correctly.
- **Tab stops** — basic left/center/right tabs work but leader rendering and decimal alignment have precision issues.

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
