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
2. `DOCXSIDE_FONTS` env var — ✅
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

## Output file size

Generated PDFs are larger than Word's PDF export. Likely causes:
- Full TTF font embedding — we embed the entire font file; Word subsets to only used glyphs
- Investigate font subsetting (e.g. `subsetter` crate or manual subsetting)
- Compare file sizes across test cases to quantify the overhead

## Performance

### Profiling setup
- Add phase timing (`log::info!`) to parse/render split for quick feedback
- Add Criterion benchmarks (full pipeline, parse-only, render-only, font scan) for regression tracking
- Use `samply` for flamegraph profiling to identify actual bottlenecks

### Known bottlenecks
- **Font scanning** — ✅ Done. Directory-level disk cache (`font-index.tsv`) with mtime invalidation + mmap for font parsing. ~500ms → ~33ms warm cache (release). Disable with `DOCXSIDE_NO_FONT_CACHE=1`.
- **Double font reads** — scan reads each font file for indexing, then `register_font` reads the same file again for embedding. Keep the data from the first read
- **Kerning extraction** — O(n²) brute-force over all WinAnsi glyph pairs. Iterate actual kern table entries instead
- **Per-word text objects** — each word emits its own BT/Tf/Td/Tj/ET sequence. Batch consecutive words sharing font+color into single text objects to reduce output size and CPU
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
- Explicit page breaks (`w:br w:type="page"`)
- Headers and footers
- Mixed inline formatting within a single line (multiple font sizes, styles, colors mid-sentence)
- Hyperlinks and bookmarks
- Multi-section documents (different page sizes/orientations)
- Multi-column layouts
