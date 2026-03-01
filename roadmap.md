# Roadmap

## Kerning / GPOS (HIGH)

The single biggest factor limiting Jaccard scores — text progressively drifts right across dense-text cases, losing ~5-10pp each. A prototype using the legacy `kern` table was implemented and reverted (improved Aptos but regressed Calibri) because Word uses GPOS kerning, not the legacy table.

To do it properly:
1. Use GPOS table for kerning lookups (requires OpenType layout engine, e.g. `rustybuzz`)
2. Apply kerning to both word width calculation (line breaking) and PDF rendering (TJ operator)

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

8 passing, ~22 failing out of ~30 non-skipped scraped fixtures.
Run `./tools/target/debug/analyze-fixtures --failing --fonts` for current breakdown.

### Floating Tables (MEDIUM — 5 failing)

Tables with `w:tblpPr` positioning attributes render as normal flow tables instead of being positioned absolutely. Existing table rendering can be reused; only the positioning pass needs to be added.

### Textbox / Shape Rendering (DONE — fills, margins, header z-order)

DrawingML textboxes (`wps:txbx` → `w:txbxContent`) and VML fallback (`v:textbox`) render text content at the correct anchor position. Shape fills (`a:solidFill` with `a:srgbClr` and `a:schemeClr` theme colors including lumMod/lumOff modifiers) render as filled rectangles. Textbox body margins (`wps:bodyPr` lIns/tIns/rIns/bIns) are respected. Header/footer content renders behind body content via content stream prepending (correct z-order). Floating images render after textbox shapes for correct layering (images on top of fills). c9211737 scores 91.5% Jaccard; 5811dabc/d0252e2f/f271d69a remain skipped (page count mismatches / font issues). Remaining gaps: text wrapping around textboxes, clipping to bounding box, shape borders/outlines, proper z-index interleaving of shapes and images.

### Anchored Image Positioning (DONE — all wrap modes)

`wp:anchor` images are now positioned absolutely regardless of wrap mode (wrapNone, wrapTight, wrapSquare, etc.). Previously only `wrapNone` anchors were treated as floating; all others fell through to inline rendering. `compute_drawing_info()` now skips all anchors (parse_runs handles them), preventing image duplication. Remaining gap: text wrapping around anchored images is not implemented (content flows through/behind images).

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
