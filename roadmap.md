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

Remaining: bundle open-source fallback fonts (Liberation, Noto) so output is consistent even without system fonts installed.

## Performance

Profile and optimize conversion speed. Areas to investigate:
- Font scanning (`scan_font_dirs`) — currently walks all system font directories on every invocation via `OnceLock`, but could be slow on first call
- Font embedding — full TTF data is embedded per font; consider subsetting
- PDF generation — measure `pdf-writer` overhead vs our layout logic
- Memory usage — large DOCX files with many images

## Test corpus

Build a larger, more diverse test corpus by scraping public DOCX files from the internet. Current fixtures (case1-9) cover limited scenarios. A broad corpus would surface edge cases in layout, font handling, and feature coverage that manual test cases miss.

Additional fixture ideas:
- Explicit page breaks (`w:br w:type="page"`)
- Headers and footers
- Mixed inline formatting within a single line (multiple font sizes, styles, colors mid-sentence)
- Hyperlinks and bookmarks
- Multi-section documents (different page sizes/orientations)
- Multi-column layouts
