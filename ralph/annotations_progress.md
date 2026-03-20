# Annotations Progress

## 2026-03-20: Fixed text justification from docDefaults (annotations #27, #43)

**Problem**: Text in `scraped/brazilian_logistics_study` was rendered left-aligned instead of justified. The document's `docDefaults/pPrDefault` specified `w:jc w:val="both"` (justify), but this default alignment was not being parsed or applied.

**Root cause**: `StyleDefaults` struct lacked an `alignment` field. The `parse_styles()` function parsed spacing, indents, and widow control from `pPrDefault` but not `w:jc`. Paragraph alignment resolution fell back to hardcoded `Alignment::Left` instead of the docDefaults value.

**Fix**:
- Added `alignment: Alignment` field to `StyleDefaults` in `src/docx/styles.rs`
- Parse `w:jc` from `pPrDefault` during style parsing
- Changed fallback in paragraph alignment resolution from `Alignment::Left` to `styles.defaults.alignment`

**Impact**:
- `scraped/brazilian_logistics_study`: Jaccard +0.8pp (12.9→13.6%), SSIM +1.5pp (24.4→25.9%)
- `scraped/lithuanian_ethics`: Jaccard +2.7pp (40.2→42.9%), SSIM +3.6pp (58.7→62.3%)
- No regressions across all 91 test cases

---

## 2026-03-20: Investigated annotation #5 (TOC position in croatian_grant_guidelines)

**Problem**: "TOC starts a bit too high up on the page compared to reference" — annotation at (156.05, 767.32) on page 1 (0-indexed).

**Analysis**: The first paragraph on page 2 is directly the Sadraj1 (TOC 1) entry with `space_before=6pt` correctly suppressed at page top. The reference fits 2-3 more lines on the page. Investigation revealed:
- Font line metrics are computed correctly (formula verified against Win metrics)
- The ~2-3pt position difference at the top accumulates across 50+ TOC entries due to subtle text width/wrapping differences
- An attempt to preserve `space_before` after natural page overflow caused regressions in 6 cases (up to -39.5pp)
- Word suppresses `space_before` at ALL page tops, not just after explicit breaks

**Conclusion**: This is a font metrics / text width accumulation issue, not a spacing bug. Requires broader improvements to text width calculation to fix.
