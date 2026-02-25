# Roadmap

## Kerning

Prototype implemented and reverted — render-only kern table kerning improved Aptos cases (SSIM +3-8pp) but caused small regressions (1-3pp) for Calibri/other fonts. Root cause: Word uses GPOS kerning, not the legacy `kern` table, and the values differ.

To do kerning properly:
1. Use GPOS table for kerning lookups (requires OpenType layout engine, e.g. `rustybuzz`)
2. With correct GPOS values, apply kerning to both word width calculation (for accurate line breaking) and PDF rendering (TJ operator)
3. The kern table approach is in git history if needed as a reference

## Font resolver

Already implemented with layered strategy:
1. Embedded fonts from DOCX — ✅
2. `DOCXIDE_FONTS` env var — ✅
3. Cross-platform system font search (macOS, Linux, Windows) — ✅
4. Helvetica Type1 fallback — ✅

Remaining:
- **Font substitution** — when a font isn't found (embedded or system), we unconditionally fall back to Helvetica (a built-in PDF Type1 font with no TrueType metrics). Word instead uses a sophisticated substitution system that picks the closest available font based on family, metrics, and panose classification. This causes two problems:
  1. **Wrong font family** — a missing sans-serif font (e.g. DejaVu Sans) gets Helvetica, which is OK, but a missing serif font (e.g. Liberation Serif) also gets Helvetica instead of Times New Roman. Should at least respect serif/sans-serif/monospace family.
  2. **No TrueType metrics** — Helvetica fallback returns `line_h_ratio: None` and `ascender_ratio: None`, forcing the renderer to use a `font_size * 1.2` estimate for line height. This compounds across pages and causes layout drift (wrong page breaks, wrong vertical positions). Real font metrics are critical for accurate layout.

  Implementation ideas:
  - Minimal: map font family → default system font (serif→Times New Roman, sans→Arial, mono→Courier New)
  - Better: parse the font's panose classification from the DOCX `fontTable.xml` and match against installed fonts
  - Best: bundle fallback fonts (Liberation, Noto) so output is consistent without system fonts installed
  - The semicolon-separated fallback lists in font names (e.g. `"Liberation Serif;Times New Roman"`) are now tried in order, but many DOCX fonts have no fallback list (e.g. `"DejaVu Sans"` with no alternative)

## Output file size — ✅ Done

Font subsetting (CIDFont/Type0/Identity-H via `subsetter` crate), FlateDecode content stream compression, and BT/ET consolidation (one per line instead of per word). Case13 went from 7.5MB (7.6× reference) to 1.0MB (1.0× reference). Most cases are at or below Word's output size.

Remaining:
- Compress font file streams with FlateDecode (currently uncompressed, ~38kB for case13)

## Performance

### Profiling setup
- Add phase timing (`log::info!`) to parse/render split for quick feedback
- Add Criterion benchmarks (full pipeline, parse-only, render-only, font scan) for regression tracking
- Use `samply` for flamegraph profiling to identify actual bottlenecks

### Known bottlenecks
- **Font scanning** — ✅ Done. Directory-level disk cache (`font-index.tsv`) with mtime invalidation + mmap for font parsing. ~500ms → ~33ms warm cache (release). Disable with `DOCXIDE_NO_FONT_CACHE=1`.
- **Double font reads** — scan reads each font file for indexing, then `register_font` reads the same file again for embedding. Keep the data from the first read
- **Kerning extraction** — O(n²) brute-force over all WinAnsi glyph pairs. Iterate actual kern table entries instead
- **Per-word text objects** — ✅ Done. One BT/ET per line with relative Td positioning and deduplicated Tf calls
- **Repeated WinAnsi conversion** — same text is converted in line-building, rendering, and table auto-fit. Pre-compute once and store in `WordChunk`
- **String allocations** — `font_key()` allocates on every call; `WordChunk` clones font name strings per word. Use indices or interning

### Parallelism (rayon)
- Font directory scanning — embarrassingly parallel, biggest win
- Font metric computation — parse face, compute widths, extract kern pairs per font independently, then write to PDF sequentially
- Paragraph line wrapping — independent per paragraph once font metrics are ready
- ZIP decompression + XML parsing — read all entries into memory, parse in parallel

### Other
- Font subsetting (related to output file size)
- Memory usage for large DOCX files with many images

## Test corpus

Build a larger, more diverse test corpus by scraping public DOCX files from the internet. Current fixtures (case1-9) cover limited scenarios. A broad corpus would surface edge cases in layout, font handling, and feature coverage that manual test cases miss.

Additional fixture ideas:
- ~~Explicit page breaks (`w:br w:type="page"`)~~ — covered by case10
- ~~Headers and footers~~ — covered by case11
- ~~Mixed inline formatting within a single line~~ — covered by case9
- ~~Inline images (PNG, JPEG, varying sizes)~~ — covered by case16
- ~~Paragraph borders and shading (all sides, combined with background color)~~ — covered by case17
- Multi-section documents (different page sizes/orientations per section)
- Deep style inheritance (3+ level chains with run vs style vs paragraph conflicts)
- Hyperlinks and bookmarks
- Multi-column layouts
- Footnotes and endnotes (parsing `footnotes.xml`, separator line, superscript references, page-bottom rendering)
- Nested/multi-level lists (outline numbering: `1. → a. → i. → •`)
- Line spacing modes (`w:lineRule="exact"` and `"atLeast"`, not just `"auto"`)
- First-line indent (`w:ind @firstLine`) and right indent (`w:ind @right`)
- Soft line breaks (`w:br` without type attribute)
- Nested tables (tables inside table cells)
- Table of Contents (right-aligned tabs + dot leaders + page field codes)
