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

---

## 2026-03-20: Fixed floating table page splitting (annotation #6)

**Problem**: Green-shaded "Važno!" box in `scraped/croatian_grant_guidelines` (page 3-4, 0-indexed) was not breaking across pages like the reference. The reference shows the green box starting on page 4 and continuing onto page 5, but our generated output rendered all the content on a single page.

**Root cause**: The green box is a single-row, single-cell table with `w:tblpPr` (floating table positioning). The table rendering code in `src/pdf/table.rs` had `!is_floating` guards on both row-splitting conditions (lines 1312, 1353), which prevented floating tables from ever splitting across page boundaries. When the row content exceeded the available space, the table was simply rendered at its floating position without any page-break handling.

**Fix** (`src/pdf/table.rs`):
- Removed `!is_floating` guard from both page-break conditions in `render_table()`
- For the row-splitting condition, changed `row_h > available_h && row_h > page_content_h` to also trigger when `is_floating` (since floating tables start mid-page, the row may not fit even if it would fit on a fresh page)
- Added `did_flush_while_floating` flag to track when a floating table causes a page break
- When a floating table spans pages, skip the cursor restoration and float zone registration (these are invalid across page boundaries)

**Impact**:
- `scraped/croatian_grant_guidelines`: Jaccard -1.4pp (8.5→7.0%), SSIM -1.7pp (19.9→18.2%) — small score drop because the split point differs from reference due to text width differences, but the structural behavior (table splitting across pages) now matches the reference
- No regressions across all other 90 test cases (including 5 other fixtures with floating tables)

---

## 2026-03-20: Fixed text wrapping around floating images in headers (annotation #7)

**Problem**: In `scraped/czech_municipal_grant_form`, the heading text "OBEC TUHAŇ" and address were rendered on top of/overlapping the coat of arms logo in the header, instead of wrapping to the right of it. The reference shows the text correctly positioned to the right of the floating image.

**Root cause**: The header/footer rendering code (`src/pdf/header_footer.rs`) had no support for text wrapping around floating images. While the body text rendering had a sophisticated `FloatZone` system, the header code simply emitted floating images at their positions and laid out text at full page width, ignoring `wrapSquare` wrap mode entirely.

Additionally, the floating image was in a separate paragraph (paragraph 0) from the text (paragraphs 1 and 2), so the float zone needed to persist across paragraph boundaries within the header.

**Fix** (`src/pdf/header_footer.rs`):
- Added a cross-paragraph `hdr_fz` variable that tracks the float zone (position, dimensions, distances) across header paragraphs
- When rendering a floating image with `WrapType::Square | Tight | Through`, register the float zone in `hdr_fz`
- For each subsequent text paragraph, check if it vertically overlaps the float zone
- When overlapping, narrow `para_text_x` and `para_text_width` to avoid the image area
- Build per-line geometry for multi-line paragraphs that span into/through the float zone
- Pass per-line widths to `build_paragraph_lines` and `render_paragraph_lines` via a new `build_lines_with_float` helper

**Impact**:
- `scraped/czech_municipal_grant_form`: Jaccard +0.8pp (10.5→11.3%), SSIM +2.2pp (27.8→30.1%)
- No regressions across all 91 test cases

---

## 2026-03-20: Fixed standard HR thickness (annotation #45)

**Problem**: In `scraped/croatian_grant_guidelines`, the horizontal rule separating the header from the body text was rendered as a thick 1.5pt filled rectangle, while the reference shows a thin ~0.5pt line.

**Root cause**: The VML horizontal rule element `<v:rect o:hr="t" o:hrstd="t" style="height:1.5pt" .../>` has the `o:hrstd="t"` attribute indicating a "standard" HR. In Word, standard HRs render as a thin 0.5pt line, with the `height` style attribute controlling the total spacing consumed (not the line thickness). Our parser captured the height correctly but ignored `o:hrstd`, causing the full 1.5pt to be used as the drawn rectangle height.

**Fix**:
- Added `is_standard: bool` field to `HorizontalRule` struct in `src/model/mod.rs`
- Parse `o:hrstd` attribute from VML shape in `src/docx/runs.rs`
- In rendering (`src/pdf/mod.rs`), when `is_standard` is true, draw a 0.5pt line centered within the specified height space instead of filling the full height

**Impact**:
- `scraped/croatian_grant_guidelines`: HR now visually matches reference (thin line vs thick bar)
- `scraped/mandated_reporter_child_abuse`: Jaccard +0.3pp (26.0→26.3%)
- No regressions across all 91 test cases

---

## 2026-03-20: Fixed page-relative textbox height in header (annotation #9)

**Problem**: In `scraped/education_consultant_posting`, body content ("Section A" table) started too far down the page, creating excessive space below "TERMS OF REFERENCE" in the header. The annotation noted "Too much space below 'Terms of Reference', and too little above it."

**Root cause**: The header contains a VML textbox (address text "United Nations Children's Fund | Pakistan Country Office") with `wp:wrapTopAndBottom` and `VRelativeFrom::Page` positioning at 69pt from page top (height 13.5pt). In `compute_header_height()`, the textbox's absolute page position (`v_offset_pt + height_pt + dist_bottom = 82.5pt`) was being added directly to `content_h` via the `_ =>` catch-all branch. This treated the absolute distance from page top as a relative content height, inflating the header height by ~48pt and pushing all body content down.

**Fix** (`src/pdf/header_footer.rs`):
- Added `sp: &SectionProperties` and `is_header: bool` parameters to `compute_header_height()`
- For `VRelativeFrom::Page` textboxes in headers, convert the absolute page position to header-relative: `contribution = max(0, tb_bottom_from_page_top - header_margin - accumulated_height)`
- Updated all 3 call sites to pass the new parameters

**Impact**:
- `scraped/education_consultant_posting`: Jaccard +1.8pp (10.4→12.2%), SSIM +3.4pp (20.7→24.1%)
- `cases/case43`: -0.9pp (pre-existing stale baseline, verified unrelated to this change)
- No regressions across remaining test cases
