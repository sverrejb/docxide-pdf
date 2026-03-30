# Annotations Progress

## 2026-03-30: Render image drop shadow from a:effectLst/a:outerShdw (annotation #31)

**Problem**: In `scraped/parish_housing_data_profile`, the map image on page 1 had no drop shadow. The reference shows a blurred gray shadow offset to the bottom-right of the image.

**Root cause**: Image drop shadow effects (`a:outerShdw` inside `pic:spPr/a:effectLst`) were entirely unimplemented. The DOCX parsing code extracted image blip data and dimensions but ignored the effect list element on the shape properties.

The XML structure: `pic:spPr > a:effectLst > a:outerShdw blurRad="292100" dist="139700" dir="2700000" > a:srgbClr val="333333" > a:alpha val="65000"` — a shadow with ~11pt distance at 45° (southeast), color #333333 at 65% opacity.

**Fix** (two commits):
- `src/model/drawing.rs`: Added `ImageShadow` struct (`offset_x`, `offset_y`, `blur_radius`, `color`) and `shadow: Option<ImageShadow>` to `EmbeddedImage`
- `src/docx/images.rs`: Added `parse_pic_shadow()` to extract `a:outerShdw` from `pic:spPr/a:effectLst`. Parses blurRad, distance/direction → x,y offsets, color with alpha pre-blended against white. Applied in all three image creation paths: `parse_run_drawing` (inline + floating) and `compute_drawing_info` (paragraph-level).
- `src/pdf/color.rs`: Added `draw_image_shadow()` helper — draws shadow rect expanded by `blur_radius * 0.5` on each side to approximate gaussian blur spread
- `src/pdf/mod.rs`: Draw shadow before body-level paragraph images (the main code path for this fixture)
- `src/pdf/layout.rs`, `positioning.rs`, `table.rs`, `header_footer.rs`: Draw shadow before images in all other rendering paths
- Shadow is pre-blended with white rather than using ExtGState alpha to avoid threading alpha_states through all rendering paths

**Impact**:
- `scraped/parish_housing_data_profile`: Drop shadow clearly visible on map image. Jaccard 33.5% → 33.9% (+0.4pp), SSIM 63.5% → 63.7% (+0.2pp)
- No regressions across all 114 test cases

---

## 2026-03-30: Deep investigation of remaining unfixed annotations

**Annotations investigated**: #5, #8, #25, #30, #55, #77, #78

**Findings**: All remaining unfixed annotations fall into two categories:

### Category 1: Page break positioning (annotations #5, #8, #25)
These annotations report content being on the wrong page — either too much or too little content on page 1, causing subsequent pages to be misaligned. Root cause analysis of annotation #5 (`scraped/croatian_grant_guidelines`) using `mutool trace` revealed:

- Reference page 2: empty paragraph baseline at y=760.72, first TOC entry "1 Opće informacije" at y=734.22 (gap of 26.5pt)
- Generated page 2: first TOC entry directly at y=759.624 (no empty paragraph)
- The empty SDT paragraph that should be on page 2 stays on page 1 because our page 1 consumes ~17pt less vertical space than Word's
- The 17pt difference accumulates from multiple empty paragraphs on the cover page with large paragraph mark font sizes (18pt from `pPr/rPr/w:sz=36`)
- Attempted fix: use `paragraph_mark_font_size` for empty paragraph heights → no effect because the "empty" paragraphs already have text-empty runs with the correct font sizes from `tallest_run_metrics`
- True root cause: cumulative font metric and line height precision differences

### Category 2: Font metric / text width differences (#30, #77, #78)
- Annotation #30 (`mongolian_human_rights_law`): extra space after "4.1.5." is a tab/text width difference
- Annotation #77 (`polish_municipal_letter`): text alignment IS correct (justified), confirmed via debug that `is_justified=true` with `extra_per_gap` applied. The visual difference is from font width differences causing different line widths (content_w=322pt vs eff_w=454pt on a line ending with `w:br`)
- Annotation #78 (`stem_partnerships_guide`): text overflow likely from column width calculation differences

### Already working: annotation #55
Annotation #55 (`japanese_interlibrary_loan`) — vertical centering in table cells IS implemented and working correctly. Debug confirmed `valign_offset` producing offsets of 5-12pt. The low overall score (4.1% Jaccard) is due to missing Japanese fonts, not vAlign.

**Conclusion**: The remaining unfixed annotations require broader improvements to font metrics precision, text width calculation, and line breaking algorithms — not targeted bug fixes. Future priorities:
- Improve OS/2 font metric handling for less common fonts
- Better paragraph mark height calculation for empty paragraphs
- Investigate Word's exact line pitch / document grid behavior

---

## 2026-03-30: Fix table cell inline image spacing (annotation #29)

**Problem**: In `scraped/mandated_reporter_child_abuse`, the "JCS-Inc. Form" text below the logo image had too much spacing compared to the reference. The text was 17.5pt too far down the page.

**Root cause**: The table cell rendering code advanced the cursor by the full `content_height` (which includes `display_height + layout_extra_height`) after rendering an inline image. The `layout_extra_height` consists of `distT + distB + effectExtent` from the `wp:inline` element (22pt total for this image: 9pt distT + 9pt distB + 2pt effectExtent top + 2pt effectExtent bottom). Word does not apply `distT`/`distB` as inter-content spacing inside table cells — they only contribute to the overall row height calculation, not to cursor advancement between paragraphs.

**Fix** (`src/pdf/table.rs`):
- Changed the cursor advancement for image paragraphs in `render_cell_content` from `cursor_y -= para.content_height` to `cursor_y -= para.image_height`
- The `para_block_height()` function still returns the full `content_height` (including `layout_extra_height`), preserving correct row height computation

**Impact**:
- `scraped/mandated_reporter_child_abuse`: "JCS-Inc. Form" positioning error reduced from 17.5pt to 4.5pt. Jaccard +0.1pp (46.9→47.0%)
- No regressions across all 114 test cases

---

## 2026-03-30: Render image outline/border from pic:spPr/a:ln (annotation #28)

**Problem**: In `scraped/mandated_reporter_child_abuse`, the JCS Inc. logo in the first-page header had no blue rectangular border. The reference shows a dark blue (#1E4D78) 2pt solid outline around the logo image.

**Root cause**: Image outline/border properties (`a:ln` inside `pic:spPr`) were entirely unimplemented. The DOCX parsing code extracted the image's blip data and dimensions but ignored the shape properties element that defines the outline. The rendering code placed images with no border support at all.

The XML structure: `wp:inline > a:graphic > a:graphicData > pic:pic > pic:spPr > a:ln w="25400" > a:solidFill > a:srgbClr val="1E4D78"` — a 2pt solid dark blue border.

**Fix**:
- `src/model/drawing.rs`: Added `stroke_color: Option<[u8; 3]>` and `stroke_width: f32` to `EmbeddedImage`
- `src/docx/images.rs`: Added `parse_pic_outline()` to extract `a:ln` from `pic:spPr` (color + width in EMU → points). Applied to both inline and floating image paths. Key detail: `pic:spPr` uses the `pic` namespace, not DML_NS.
- `src/pdf/layout.rs`: Draw stroke rectangle after inline image XObject placement
- `src/pdf/positioning.rs`: Draw stroke rectangle after floating image placement
- `src/pdf/table.rs`: Draw stroke rectangle after table cell image placement
- `src/pdf/header_footer.rs`: Draw stroke rectangle after header/footer image placement
- `src/pdf/table_layout.rs`: Propagate stroke fields through `CellParagraphLayout`

**Impact**:
- `scraped/mandated_reporter_child_abuse`: SSIM +0.1pp (67.2→67.3%), Jaccard stable at ~47%
- No regressions across all test cases

---

## 2026-03-30: Fixed decimal tab stop reuse preventing dot leader rendering (annotation #26)

**Problem**: In `scraped/air_pollution_permit_form`, the "V:" line at the bottom of the form was missing dotted leader lines between "dňa" and "2020". The reference shows `V: .................., dňa.............................................2020` but the generated output showed `V: ..................., dňa2020` with no dots.

**Root cause**: The paragraph uses two tab stops with dot leaders: a decimal tab at 120.5pt and a left tab at 276.5pt. When the decimal tab aligned the text segment ", dňa " before its tab position (seg_start=93.5pt), the text ended at current_x=117.5pt — slightly below the tab stop at 120.5pt. The next `find_next_tab_stop(117.5)` call found `120.5 > 118.0` and returned the **same** decimal tab stop again instead of advancing to the left tab at 276.5pt. With the same tab stop reused, the "2020" segment was positioned at ~117.5pt (right after "dňa") with zero gap — no room for leader dots.

**Fix** (`src/pdf/layout.rs`):
- After each tab segment's text layout, advance `current_x` to at least the tab stop position (`current_x = current_x.max(effective_tab_target)`)
- This ensures `find_next_tab_stop` skips already-consumed tab stops on subsequent tab characters
- Also properly tracks the effective tab target through line overflow (when a tab causes a new line, the target updates to the new stop's position)

**Impact**:
- `scraped/air_pollution_permit_form`: Dot leaders now render correctly between "dňa" and "2020" on both V: lines. Jaccard score unchanged at 15.0% (dots gained offset by position shift of "2020" to correct location)
- No regressions across all test cases

---

## 2026-03-30: Include list label font metrics in line height calculation (annotation #66)

**Problem**: In `case33`, each bullet point caused a slightly larger vertical drift from the reference. The accumulated drift was ~2pt by the end of the page. The annotation explicitly asked for a thorough examination.

**Root cause**: Word includes the numbering label character's font metrics (e.g., Symbol font for bullet `•`) in the tallest-font-on-the-line calculation for line height. Our code only considered paragraph text runs via `tallest_run_metrics()`, ignoring the label font entirely. When the label font (Symbol) has larger vertical metrics than the text font (Calibri), Word produces taller bullet lines (~16pt) while we produced shorter ones (~15.44pt).

**Analysis**: Extracted precise text Y positions from both reference and generated PDFs using `mutool trace`. Reference bullet-to-bullet gap was 16.00pt (using SymbolMT metrics) while ours was 15.44pt (using only Calibri metrics). The 0.56pt per-bullet error accumulated across 3 bullets to ~1.7pt, then continued growing through subsequent paragraphs.

**Fix** (`src/pdf/mod.rs`):
- Added `label_boosted_line_h()` helper that looks up the label font's metrics from `ctx.fonts` and uses `max(text_line_h, label_line_h)` when the label font produces a taller line
- Only applies when label and text font sizes match (within 1pt) — avoids aggressive boosting for special numbering styles with dramatically different label sizes (e.g., 20pt numbered label on 10pt text)
- Applied to `first_line_h` in the content height calculation, matching Word's behavior where only the first line (containing the bullet) is affected

**Impact**:
- `cases/case33`: Jaccard +1.5pp (55.9→57.4%), SSIM stable at 95.4%
- `cases/case3`: +2.3pp
- `scraped/polish_archery_r..`: +11.8pp Jaccard (18.3→30.1%), +26.7pp SSIM (50.8→77.4%)
- `scraped/indonesian_bench..`: +4.9pp (35.6→40.5%)
- `samples/sample500kB`: +1.6pp
- 3 more fixtures improved by 0.1–0.6pp
- No regressions caused by this change (german_mezzo -4.2pp is pre-existing)

---

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

---

## 2026-03-21: Investigated annotation #8 (east_asia_conference_form page overflow)

**Problem**: In `scraped/east_asia_conference_form`, all content renders on page 1 but the reference has "신청자(申請者):" on page 2. The page overflow is caused by natural content height, not explicit page breaks.

**Analysis**: The document contains a floating table (`tblpPr`) with 6 rows of Korean/Japanese form content. Our table renders as ~471pt total height, but the reference table is ~38pt taller. This difference is caused by font metric disparities:
- Reference (Windows): Uses MalgunGothicBold (맑은 고딕) at 12.96pt with Windows-native metrics
- Generated (macOS): Falls back to HiraginoSansW3 with line_h_ratio=1.33
- The fallback font's larger line_h_ratio inflates cell content heights, causing rows 3-4 to exceed their trHeight minimums (42.75pt) by 15.88pt each
- Row heights that exceed minimums make the table taller in our rendering, but the total table height is STILL shorter than Word's because the correct font would produce different wrapping/metrics overall

**Conclusion**: This is a font availability issue on macOS. Malgun Gothic (맑은 고딕) is not installed, and the fallback font has different line height metrics. Fixing requires either installing the correct fonts or implementing font metric estimation for unavailable fonts.

---

## 2026-03-21: Fixed numbering level tab stop for list paragraph indentation (annotation #17)

**Problem**: In `scraped/polish_building_procurement_spec`, the first numbered list item "1. Nazwa..." had its text indented too far to the right compared to the reference. The "1." label was at x=56.7pt (correct) but "Nazwa" started at x=80.2pt instead of x=70.8pt.

**Root cause**: The first list item had an explicit paragraph-level `w:ind w:left="470" w:hanging="470"` (23.5pt each) that overrode the numbering level's default indent of 283 twips (14.15pt). The numbering level also defined a tab stop at position 283 for positioning text after the label. Two issues:

1. The numbering level's tab stop was not being parsed or propagated to the paragraph's tab list
2. The `text_hanging` for list paragraphs was always set to 0.0, meaning the first line text started at `indent_left` (23.5pt from margin) instead of at the numbering level's tab stop position (14.15pt)

**Fix**:
- Added `tab_stop: Option<f32>` to `LevelDef` and `ListLabelInfo` in `src/docx/numbering.rs`
- Parse `<w:tabs><w:tab>` from the numbering level's `<w:pPr>` element
- Added `num_level_tab_stop: Option<f32>` to the `Paragraph` model
- Propagated the numbering tab stop to the paragraph's tab list in `src/docx/mod.rs`
- In the rendering code (`src/pdf/mod.rs`, `src/pdf/header_footer.rs`), when a list paragraph has `indent_left == indent_hanging` (first line at margin) and a numbering tab stop exists, compute `text_hanging = indent_left - num_tab_stop` to shift the first line text left to the tab stop position

**Impact**:
- `scraped/polish_building_procurement_spec`: Item 1 text now at x=70.8pt (matching items 2-6), reference shows x=74.0pt — much closer
- `cases/case43`: Jaccard +0.9pp (21.1→22.1%)
- No Jaccard regressions across all test cases

---

## 2026-03-21: Fixed hanging indent in table cell rendering (annotation #34)

**Problem**: In `scraped/turkish_ancient_religions_plan`, the heading "ŞAMANİZM" was indented to the right compared to other headings like "BUDİZM" and "ZERDÜŞTLİK". All headings should start at the same x-position within their table cell.

**Root cause**: The paragraph uses the `OkumaParas` style (which has `numId=1` for list formatting) but overrides with `numId=0` (no list) and explicit `w:ind w:left="432" w:hanging="360"`. With hanging indent, the first line should start at `indent_left - indent_hanging = 3.6pt` — same as the `Konu-basligi` style's `ind left=72 twips = 3.6pt`.

The bug was in `src/pdf/table.rs`: the table cell rendering code always passed `0.0` as `first_line_hanging` to `render_paragraph_lines()`, even when the paragraph had a non-zero `indent_hanging`. The lines were correctly built with the hanging indent width in `compute_row_layouts()`, but the rendering didn't shift the first line left to account for it.

**Fix** (`src/pdf/table.rs`):
- At both cell rendering locations, compute `first_line_hanging` based on whether the paragraph is a list or not
- For non-list paragraphs: pass `para.indent_hanging` as `first_line_hanging`
- For list paragraphs: keep `0.0` (label handles positioning separately)

**Impact**:
- `scraped/turkish_ancient_religions_plan`: Jaccard +0.3pp (23.3→23.6%), SSIM +0.4pp (54.0→54.4%)
- `scraped/polish_archery_range_plan`: Jaccard +0.3pp (14.9→15.2%), SSIM +1.3pp (48.4→49.7%)
- No regressions above 2% threshold across all test cases

---

## 2026-03-21: Fixed page break trailing leading tolerance (annotation #10)

**Problem**: In `scraped/feminist_voice_dissertation`, the "Keywords: Feminist, Feminism, Patriarchy..." line was pushed to page 8 instead of fitting at the bottom of page 7 (the ABSTRACT page) like the reference. The abstract text content was identical between reference and generated (same line breaks, same Y positions within ±1pt), but the Keywords line overflowed.

**Root cause**: The page break decision in `src/pdf/mod.rs` checked whether the FULL `content_h` (including the last line's inter-line spacing) fit within the remaining content area. For the Keywords paragraph with Times New Roman 12pt at 1.5× line spacing, `content_h = line_h = 20.70pt`. But only ~13.80pt of actual text height was needed — the remaining 6.90pt was trailing inter-line spacing that serves no purpose for the last line on a page.

The calculation: after the body text, `slot_top ≈ 99.70` (PDF coords from bottom). Keywords `needed = 10pt (inter-paragraph gap) + 20.70pt (full line_h) = 30.70pt`. Check: `99.70 - 30.70 = 69.00 < 72.00 (margin_bottom)` → overflow! But the Keywords baseline at y≈79 from bottom with descent ~2.6pt ends at ~76.4, well within the 72pt margin. Word allows the trailing leading to extend past the bottom margin.

**Fix** (`src/pdf/mod.rs`):
- Compute `last_line_lead`: the excess of `line_h` over the font's single-line height (`font_size * line_h_ratio`), representing trailing inter-line spacing that can overflow into the bottom margin
- Adjust the page break condition from `slot_top - needed < margin_bottom` to `slot_top - needed + last_line_lead < margin_bottom`
- Only applies to text paragraphs with non-Exact line spacing (Auto/AtLeast); images, charts, and SmartArt are excluded

**Impact**:
- `scraped/feminist_voice_dissertation`: Jaccard +6.4pp (33.6→40.0%), SSIM +12.9pp (69.5→82.3%)
- `scraped/brazilian_logistics_study`: Jaccard +0.9pp (13.6→14.6%), SSIM +1.6pp (25.9→27.6%)
- `cases/case43`: Jaccard +0.9pp (21.1→22.1%), SSIM +0.8pp (27.6→28.4%)
- No regressions across all 92 test cases

---

## 2026-03-21: Fixed floating image page overflow for paragraph-relative wrapSquare images (annotation #13)

**Problem**: In `scraped/indonesian_benchmarking_guide`, a large floating anchor image (questionnaire table, 251×374pt with `wrapSquare`) was rendered on top of the "Benchmarking Process" flowchart and other content on page 6, instead of being pushed to page 7 like the reference. The image overlapped existing content because no page break was triggered.

**Root cause**: The page-break decision in `src/pdf/mod.rs` computed `content_h` for each paragraph and checked if `slot_top - needed < margin_bottom`. For floating images with `wrapSquare`, the image height was only included in `content_h` when the image was wider than 90% of text width (meaning text couldn't wrap beside it). The questionnaire image was ~56% of text width, so its 374pt height was excluded from `content_h`. The paragraph's text content was ~14pt (one empty line), so the page-break check saw only 14pt needed — far below the remaining page space — and didn't trigger a page break. The 374pt image was then rendered at the paragraph's position, extending far below the bottom margin and overlapping subsequent content.

**Fix** (`src/pdf/mod.rs`):
- Added a separate `float_overflow_h` variable to track the height of paragraph-relative floating images with `wrapSquare`/`Tight`/`Through` that are narrower than 90% of text width
- These images have text wrapping beside them, so their height should NOT inflate `content_h` (which controls cursor advancement after the paragraph)
- Computed `needed_with_floats = max(needed, inter_gap + float_overflow_h)` for use only in the page-break condition
- This ensures the page-break check accounts for the full image height while cursor advancement still uses the text-only height

**Impact**:
- `scraped/indonesian_benchmarking_guide`: Jaccard +10.0pp (22.9→32.9%), SSIM +9.1pp (43.0→52.1%)
- `cases/case43`: Jaccard -0.9pp (22.1→21.1%) — within noise, this case fluctuates at this level
- No regressions above 2% threshold across all 92 test cases

---

## 2026-03-21: Fixed floating images in table cells not rendering (annotation #18)

**Problem**: In `scraped/polish_municipal_letter`, the municipal coat of arms (red emblem) in the header table was missing from the generated PDF. The reference shows the emblem in the top-left cell of the floating table.

**Root cause**: Three issues prevented floating images in table cells from rendering:

1. **Parsing** (`src/docx/tables.rs`): `parse_runs()` correctly extracted floating images from the cell paragraph's `wp:anchor` elements, but the `Paragraph` struct was constructed with `..Paragraph::default()` which set `floating_images` to an empty vec, discarding the parsed data.

2. **Image embedding** (`src/pdf/images.rs`): The table cell image embedding loop only iterated `para.image` (inline images), not `para.floating_images`. The floating image's XObject was never written to the PDF.

3. **Rendering** (`src/pdf/table.rs`): Two sub-issues:
   - `para_has_visible_content()` only checked for text lines and labels, not floating images. A paragraph containing only a floating image was considered "invisible" and the entire cell was skipped.
   - `render_cell_content()` and `render_partial_cell_content()` had no code to draw floating images at their positions.

The image was a `wp:anchor` with `layoutInCell="1"`, `wrapNone`, positioned `relativeFrom="column"` (h_offset=19.35pt) and `relativeFrom="paragraph"` (v_offset=8.2pt), dimensions 67.5×82.2pt.

**Fix**:
- `src/docx/tables.rs`: Pass `floating_images: parsed.floating_images` to the Paragraph struct
- `src/pdf/images.rs`: Extended table cell image embedding to also iterate `para.floating_images` and embed each one via `embed_single_image`
- `src/pdf/table.rs`:
  - Added `CellFloatingImageLayout` struct to store pre-resolved PDF name, dimensions, and cell-relative offsets
  - Updated `para_has_visible_content()` to return true when floating images are present
  - In `compute_row_layouts()`: resolve FloatingImage positions to cell-relative offsets and look up PDF names from `table_cell_image_names`
  - In `render_cell_content()` and `render_partial_cell_content()`: draw floating images at their cell-relative positions

**Impact**:
- `scraped/polish_municipal_letter`: Jaccard +8.3pp (27.2→35.5%), SSIM +2.5pp (69.4→71.9%)
- `cases/case43`: Jaccard +0.9pp (21.1→22.1%), SSIM +0.8pp (27.6→28.4%)
- No regressions across all 92 test cases

---

## 2026-03-21: Fixed vertical text centering in table cells (annotation #14)

**Problem**: In `scraped/japanese_interlibrary_loan`, vertical text (e.g., "申込者", "申込図書") in the first column of the form table was rendered at the top of its cell instead of being vertically centered like the reference. The cells used `w:textDirection w:val="tbRlV"` for vertical text but had no explicit `w:vAlign` attribute.

**Root cause**: In vertical text cells, `w:jc` (paragraph justification) controls alignment along the text flow direction, which is vertical. So `w:jc w:val="center"` means the text should be centered vertically within the cell. The `render_vertical_cjk_cell()` function in `src/pdf/table.rs` only checked `cell.v_align` (which defaulted to `Top` since no `w:vAlign` was specified) and ignored the paragraph's `jc` alignment entirely.

**Fix** (`src/pdf/table.rs`):
- In `render_vertical_cjk_cell()`, when `cell.v_align` is `Top` (default), check the first paragraph's alignment
- If the paragraph has `Alignment::Center`, use `CellVAlign::Center` for the vertical offset
- If the paragraph has `Alignment::Right`, use `CellVAlign::Bottom` for the vertical offset
- This correctly maps horizontal paragraph alignment to vertical positioning in vertical text cells

**Impact**:
- `scraped/japanese_interlibrary_loan`: Jaccard +0.1pp (4.1→4.3%), SSIM +0.2pp (26.7→26.8%)
- Small improvement because the overall score is dominated by font differences (MS Gothic not available on macOS)
- No regressions across all 92 test cases

---

## 2026-03-21: Fixed table cell baseline positioning using full font size (annotation #15)

**Problem**: In `scraped/japanese_interlibrary_loan`, text in the bottom row of the form table had too little padding above it compared to the reference. The annotation at (41.09, 70.68) on page 0 noted visible gap difference between reference and generated output.

**Root cause**: The table cell baseline formula used `cursor_y - font_size * ascender_ratio` to position the first line's baseline below the cell top. Word instead uses `cursor_y - font_size` (full em-square height, not just the ascender). Precise PDF coordinate analysis confirmed: for MS Gothic 11pt in the reference, the gap from cell top to baseline was exactly 11.0pt (= font_size), while our code produced 9.45pt (= font_size × 0.86 ascender_ratio). The difference of ~1.55pt was visible as insufficient top padding.

**Fix** (`src/pdf/table.rs`):
- Changed both `render_cell_content()` and `render_partial_cell_content()` baseline formula from `cursor_y - para.font_size * para.ascender_ratio` to `cursor_y - para.font_size`
- This shifts text down by `font_size * (1 - ascender_ratio)` in all table cells, matching Word's positioning

**Impact**:
- `scraped/japanese_interlibrary_loan`: Jaccard +0.6pp (4.3→4.8%), SSIM +0.1pp (26.8→26.9%)
- `cases/case6`: Jaccard +3.9pp (39.5→43.4%) — handcrafted test confirms formula correctness
- `scraped/parish_housing_data_profile`: Jaccard +2.7pp (30.0→32.7%)
- `scraped/russian_university_proceedings`: Jaccard +2.1pp (20.2→22.3%)
- `scraped/traditional_skills_heritage`: Jaccard +2.0pp (13.3→15.3%)
- `scraped/polish_archery_range_plan`: Jaccard +1.8pp (15.2→16.9%)
- Two minor Jaccard regressions (SSIM unchanged, confirming shift-only effect): `scraped/italian_project_..` -2.3pp, `scraped/polish_municipal..` -2.8pp
- No structural layout changes in any case (text boundary scores unchanged)

---

## 2026-03-21: Fixed page numbering for continuous sections (annotation #32)

**Problem**: In `scraped/stem_partnerships_guide`, the footer page numbers were off by one. Physical page 3 showed "1" instead of "2" as in the reference. The document has three sections: section 0 (cover, NextPage), section 1 (empty, Continuous), and section 2 (body content, Continuous with `pgNumType start=1`).

**Root cause**: The `page_section_indices` tuple stored only `(hf_section, is_first)`. For continuous section breaks, `hf_section` stays as the previous section (correct for header/footer selection), but page numbering also used `hf_section` to count pages within a section. Pages where section 2's content was rendering but `hf_section = 0` (inherited from section 0) were not counted as part of section 2, giving wrong page numbers.

**Fix** (`src/pdf/mod.rs`, `src/pdf/assembly.rs`):
- Expanded `page_section_indices` from `(usize, bool)` to `(usize, bool, usize)` — `(hf_section, is_first, content_section)`
- `content_section` always tracks the actual section being rendered (set to `sect_idx` in `flush_page`)
- Page numbering now uses `content_section` to count pages within a section
- Geometry lookups (columns, footnotes, page size) also use `content_section`
- Header/footer resolution continues using `hf_section`

**Also verified**: Annotations #16 ("Page 2 should not exist" in `japanese_interlibrary_loan`) and #19 ("Best Practice Guide should not be on page 2" in `stem_partnerships_guide`) were already fixed — marked as fixed.

**Impact**:
- `scraped/stem_partnerships_guide`: Page numbers now match reference (page 3 shows "2" instead of "1")
- `scraped/transition_to_work`: Jaccard +0.2pp (26.1→26.2%)
- No regressions above 2% threshold across all test cases

---

## 2026-03-21: Fixed table cell vAlign content height (annotation #20)

**Problem**: In `scraped/uk_commercial_lease_template`, all text on page 1 was shifted ~9.5pt too low compared to the reference. The annotation at (299.57, 203.63) on page 0 noted "Text needs to be higher up." Precise measurement showed "Dated" text was at 103.59pt from page top vs 94.12pt in reference (9.47pt error). The centered "[LANDLORD]" block was 5.03pt too low.

**Root cause**: The `valign_offset()` calculation for bottom/center-aligned table cells used a `content_h` that excluded the trailing `space_after` of the last paragraph. However, `compute_row_layouts()` correctly included `prev_space_after` in the row's `total_h` (line 896). This mismatch meant the rendering code underestimated the content block size, making `v_offset` too large and pushing text too far down.

For the "Dated" cell (vAlign=bottom, Normal style space_after=9pt):
- Before: content_h=12.65pt → baseline 2.65pt from cell bottom (reference: 12.0pt)
- After: content_h=21.65pt → baseline 11.65pt from cell bottom — within 0.35pt of reference

**Fix** (`src/pdf/table.rs`):
- Added `space_after: f32` field to `CellParagraphLayout` struct
- Created `cell_content_h_for_valign()` helper that computes content height including the last paragraph's `space_after`
- Replaced inline content_h calculations at all three rendering sites (normal rows, split rows, header rows) with the helper

**Impact**:
- `scraped/uk_commercial_lease_template`: "Dated" position error reduced from 9.47pt to 0.47pt; "[LANDLORD]" from 5.03pt to 0.53pt
- `scraped/turkish_prostate_cancer_course`: Jaccard +2.1pp (35.8→37.9%), SSIM +0.3pp
- `scraped/turkish_ancient_religions_plan`: Jaccard -2.8pp (23.3→20.5%), SSIM +0.1pp — the dense center-aligned table shifted text direction in Jaccard overlap, but SSIM confirms spatial improvement
- No other regressions across all 92 test cases

---

## 2026-03-21: Fixed double border rendering thickness (annotation #33)

**Problem**: In `scraped/turkish_ancient_religions_plan`, the horizontal borders separating major table sections (e.g., between "ŞAMANİZM" and "ZERDÜŞTLİK") appeared the same thin width as inner row borders. The reference shows these section separators as visibly thicker/bolder double-line borders.

**Root cause**: The table uses `w:val="double" w:sz="4"` for section separator borders (via `tcBorders`) and `w:val="single" w:sz="4"` for inner borders (via `tblBorders insideH`). The double border rendering in `draw_border()` used `thin = max(w/3, 0.25)` for each line and `gap = max(w, 0.75)`. For `w:sz="4"` (0.5pt), this produced two 0.25pt lines with 0.75pt gap — barely distinguishable from a 0.5pt single line. Word renders each line of a double border at approximately the full specified width, making them clearly thicker than single borders.

**Fix** (`src/pdf/table.rs`):
- Changed double border line thickness from `(w / 3.0).max(0.25)` to `w.max(0.25)` — each line now uses the full border width
- Changed gap from `w.max(0.75)` to `thin` (= the line width) — proportional gap matching Word's rendering

**Impact**:
- `scraped/turkish_ancient_religions_plan`: Jaccard +2.0pp (20.5→22.5%)
- `cases/case43`: Jaccard +0.9pp (21.1→22.1%)
- `scraped/polish_municipal_letter`: Jaccard +0.6pp (32.7→33.3%)
- No Jaccard regressions across all 92 test cases

---

## 2026-03-23: Fixed table auto-fit ignoring cell margins (annotation #35)

**Problem**: In `cases/case6`, the text "/api/auth/refresh" in the Performance Metrics table overflowed its cell boundary into the adjacent column. The reference shows the first column wide enough to contain the text.

**Root cause**: The `auto_fit_columns()` function in `src/pdf/table.rs` computed minimum column widths based on word widths alone, without accounting for cell margins (padding). The default cell margins are 5.4pt left + 5.4pt right = 10.8pt total horizontal padding. For a column of 86.4pt with 10.8pt padding, the available text width is only 75.6pt. The word "/api/auth/refresh" at ~77pt exceeded this, but the auto-fit check compared against the full 86.4pt column width (77 < 86.4) and didn't expand the column.

**Fix** (`src/pdf/table.rs`):
- In `auto_fit_columns()`, resolve each cell's effective margins (`cell.cell_margins` or table-level `cell_margins`)
- Add horizontal padding (`ecm.left + ecm.right`) to each word width before comparing against column widths
- This ensures columns expand to fit text INCLUDING the cell padding

**Impact**:
- `cases/case6`: Jaccard +2.1pp (43.4→45.5%), SSIM +2.4pp (85.2→87.5%)
- No regressions across all test cases

---

## 2026-03-23: Fixed paragraph bottom border extent double-counted in space_after (annotation #2)

**Problem**: In `samples/samtale`, the answer text (e.g., "I stor grad.") was floating above the grey bottom border lines instead of sitting directly above them. This caused cumulative vertical drift — each bordered paragraph added extra space, and by items 4-5 on page 2, the text was visibly shifted down compared to the reference.

**Root cause**: In `src/docx/mod.rs`, the paragraph parsing code added `bdr_bottom_extra = space_pt + width_pt` from the bottom border to `space_after`. This meant the border extent was treated as inter-paragraph spacing rather than paragraph height. Two problems:
1. **Double-counting**: The rendering code separately handled border positioning via `bdr_bottom_pad`, so the border space was accounted for twice — once in `space_after` and once in the border drawing
2. **Incorrect contextualSpacing interaction**: When `contextualSpacing` suppressed `space_after` to 0, the border extent was also suppressed, leaving no room below the border

For the samtale answer paragraphs (Normal style with `w:pBdr bottom sz="18" space="1"`), this added 3.25pt extra per bordered paragraph (space_pt=1 + width_pt=2.25). Over 5 question-answer pairs, this accumulated ~16pt of vertical drift.

**Fix**:
- `src/docx/mod.rs`: Removed `bdr_bottom_extra` from `space_after` calculation — border extent is not spacing
- `src/pdf/mod.rs`: Added `bdr_bottom_extent` (= `space_pt + width_pt`) to the slot_top advancement and `needed` page-break calculation, treating the border extent as part of the paragraph's consumed height rather than inter-paragraph spacing

This ensures the border extent is always consumed (not suppressible by `contextualSpacing`) and doesn't inflate the inter-paragraph gap via `max(prev_space_after, next_space_before)` collapsing.

**Impact**:
- `samples/samtale`: SSIM +3.4pp (54.3→57.7%), Jaccard -0.3pp (12.8→12.6%, within noise)
- No regressions across all test cases

---

## 2026-03-23: Fixed chart category axis label positioning (annotation #37)

**Problem**: In `cases/case30`, the x-axis labels ("Jan", "Feb", etc.) on the line chart were positioned directly on the tick marks instead of being centered between them. The reference shows labels centered in each category segment between adjacent tick marks.

**Root cause**: The label positioning code in `src/pdf/charts.rs` had a special case for line/area charts (`is_point_chart`) that used `plot_w / (num_categories - 1)` spacing — the same edge-to-edge spacing as data points. This placed labels at data point positions (ON tick marks). However, tick marks were already correctly drawn at segment boundaries using `plot_w / num_categories` spacing. The label and tick mark coordinate systems were inconsistent.

**Fix** (`src/pdf/charts.rs`):
- Removed the `is_point_chart` special case for label positioning
- All category axis labels now use the same formula: `plot_x + (ci + 0.5) * (plot_w / num_categories) - tw/2` — centered in each category segment between tick marks
- Data point positions remain unchanged (edge-to-edge spacing for line/area charts)

**Impact**:
- `cases/case30`: Jaccard -0.1pp (81.9→81.8%, within noise), SSIM -1.1pp (83.3→82.2%) — labels shifted to correct positions but font width differences cause imperfect centering
- `cases/case29` (bar charts): unchanged — already used correct centering formula
- `cases/case31`: unchanged
- No regressions above 2% across all test cases

---

## 2026-03-23: Fixed SmartArt text labels overflowing shape bounds (annotation #21)

**Problem**: In `scraped/vaccines_history_chapter`, the SmartArt timeline labels (date labels and descriptions like "(1796) Edward Jenner invented the Small pox vaccine") were rendered too wide, extending far beyond their shape boundaries. The reference shows labels wrapped within their shape bounds.

**Root cause**: The SmartArt text rendering code in `src/pdf/smartart.rs` split text only on explicit `\n` characters (paragraph breaks from the XML) but never word-wrapped within shape bounds. Each paragraph was rendered as a single continuous line. For labels with long descriptions (e.g., "Influenza reccommend as an additonal vaccine for children" in a 36pt-wide shape at 5pt font), the text width far exceeded the shape width (142pt vs 36pt), causing it to overflow the shape boundaries and overlap adjacent labels.

**Fix** (`src/pdf/smartart.rs`):
- Added `wrap_text_into()` helper that word-wraps text into lines fitting within `max_width` using greedy line-breaking
- Modified `render_smartart()` to wrap each paragraph line before rendering
- Vertical centering recalculated based on wrapped line count (not raw paragraph count)

**Impact**:
- `scraped/vaccines_history_chapter`: SSIM +0.1pp (51.8→51.9%) — small score improvement because labels are tiny (5pt) relative to full page, but visually the labels now correctly wrap within shape bounds matching the reference
- No regressions caused by this change across all test cases (pre-existing baseline drift in 3 non-SmartArt fixtures was accepted)

---

## 2026-03-23: Fixed paragraph bottom border positioned too close to text (annotation #22)

**Problem**: In `samples/samtale`, the green underline below "To- og femmånederssamtalen" was rendered almost touching the text descenders, while the reference shows a clear ~13pt gap between the text baseline and the border. The grey separator borders on page 2 were also ~1pt too close.

**Root cause**: The border positioning code in `src/pdf/mod.rs` subtracted `trailing_lead = (line_h - font_size).max(0.0)` from the content height before computing the border position. This was intended to make bottom-only borders "sit close to the text" by measuring the `space` attribute from the text descent rather than the line-height bottom. However, Word actually measures `w:space` from the full line-height content bottom (including trailing leading), not from the text descent.

For the heading paragraph (26pt Calibri Bold, Auto 1.15 line spacing):
- `line_h = 36.50pt`, `font_size = 26pt`, `trailing_lead = 10.50pt`
- With trailing_lead: gap from baseline to border top = **2.24pt**
- Without trailing_lead: gap = **12.75pt** — matching the reference's 12.78pt within 0.03pt
- The grey borders (11pt, Auto 1.15) also improved: gap 3.96pt vs reference 3.76pt (previously 2.65pt)

**Fix** (`src/pdf/mod.rs`):
- Removed the `trailing_lead` computation entirely — it was based on incorrect assumptions about how Word positions paragraph bottom borders
- Simplified both the shading background and border positioning formulas by removing the `+ trailing_lead` term

**Impact**:
- `samples/samtale`: Jaccard -1.6pp (12.6→11.0%, from overall vertical shift), SSIM -0.3pp (57.7→57.4%)
- `cases/case4`: Jaccard +1.6pp (71.3→72.9%), SSIM +0.8pp (89.5→90.2%)
- Green line gap now matches reference within 0.03pt; grey borders within 0.2pt
- No SSIM regressions above 0.3pp across all test cases

---

## 2026-03-23: Fixed line break run inflating paragraph line height (annotation #23)

**Problem**: In `samples/samtale`, the space above "Medarbeiderens egenevaluering" was too large compared to the reference. The subheading was at 122.38pt from page top while the reference had it at 114.78pt — a 7.60pt vertical error.

**Root cause**: The paragraph containing "Medarbeiderens egenevaluering" (14pt font) also had a trailing `<w:br/>` line break run with `w:sz="52"` (26pt). The `tallest_run_metrics()` function in `src/pdf/layout.rs` iterated over ALL runs to find the tallest font, including line break runs. The 26pt break run was selected as the tallest, causing:
- `font_size = 26pt` instead of `14pt` for the paragraph
- `line_h = 26 × 1.22 × 1.15 = 36.5pt` instead of `~19.7pt`
- Baseline positioned 19.5pt below slot_top (26×0.75) instead of 10.5pt (14×0.75)
- This pushed the visible 14pt text 9pt lower than it should be

The gap between heading and subheading measured 36.35pt (generated) vs 25.13pt (reference).

**Fix** (`src/pdf/layout.rs`):
- Skip runs with `is_line_break: true` in `tallest_run_metrics()` — line break runs only affect the empty line they create, not the paragraph's overall line height

**Impact**:
- `samples/samtale`: SSIM +0.3pp (57.4→57.6%), subheading gap now 24.92pt vs reference 25.13pt (0.21pt difference — essentially perfect)
- No regressions across all 100 test cases

---

## 2026-03-29: Fixed leading-space paragraph indent (annotation #71)

**Problem**: In `scraped/vaccines_history_chapter`, the paragraph starting with "Vaccines are an important part..." had no visible first-line indent. The reference shows the paragraph indented ~7 spaces from the left margin. The annotation at (91.9, 266.24) on page 0 noted "Vaccines should be indented here at the start of the paragraph."

**Root cause**: The paragraph's first run contained 7 space characters `"       "` followed by a second run with `"Vaccines are an important..."`. In `build_paragraph_lines()`, the trailing-space accumulation correctly added the 7 spaces to `pending_space_w`. However, when placing the first word on a line (`current_chunks.is_empty()`), the code set `proposed_x = current_x` (= 0.0), ignoring `pending_space_w`. The `need_space` check required `current_chunks` to be non-empty, so leading spaces before the first word were always lost.

**Fix** (`src/pdf/layout.rs`):
- Added a third branch to the `proposed_x` calculation: when `current_chunks.is_empty()` and `pending_space_w > 0.0`, set `proposed_x = pending_space_w` to preserve leading spaces as a visual indent

**Also fixed** (`src/pdf/textbox_render.rs`):
- Added PDF-level clipping (save_state + clip rect + restore_state) for fixed-size textboxes (AutoFit::None)
- Added text-level overflow break: `render_textbox_paragraphs` stops rendering when cursor_y drops below the textbox bottom boundary

**Impact**:
- `scraped/vaccines_history_chapter`: Jaccard +0.5pp (42.2→42.7%), SSIM +3.4pp (51.7→55.2%)
- `scraped/russian_university_proceedings`: Jaccard +1.1pp (22.7→23.8%), SSIM +1.6pp (50.7→52.4%)
- `scraped/feminist_voice_dissertation`: Jaccard +0.5pp (65.9→66.4%), SSIM +0.5pp (85.2→85.7%)
- `scraped/federal_procurement_terms`: Jaccard +0.4pp (53.7→54.2%)
- 16 fixtures improved overall, no regressions above 0.1pp

---

## 2026-03-29: Fixed inline textbox rendering (annotation #53)

**Problem**: In `scraped/federal_procurement_terms`, a bordered text box containing "Applicable to Grants, Subgrants, Cooperative Agreements, and Contracts exceeding $100,000 in federal funds" was completely missing from the generated PDF (page 9). The reference shows it as a bordered box below the "ATTACHMENT B: LOBBYING CERTIFICATION" heading.

**Root cause**: The text box was an **inline drawing** (`wp:inline`) containing a `wps:wsp` with text content, not a floating anchor (`wp:anchor`). In `parse_run_drawing()` (`src/docx/images.rs`), the inline drawing path only handled images (`find_blip_embed`), charts, and SmartArt. Inline textboxes (`wps:wsp` with `wps:txbx`) were silently ignored.

**Fix** (`src/docx/images.rs`):
- Added inline textbox detection in the `is_inline` branch of `parse_run_drawing()`
- Calls `parse_textbox_from_wsp()` (same as the anchor path) to parse the textbox content
- Returns `RunDrawingResult::TextBox` with paragraph-relative positioning and `WrapType::TopAndBottom` so it renders as a block element at the paragraph position
- Preserves fill, stroke, margins, and text anchor from the parsed wsp properties

**Also verified**:
- Annotation #68 (`scraped/slovak_misdemeanor_amendment`, "Dôvodová správa" centering) was already fixed by the leading-space indent fix in annotation #71 — marked as fixed
- Annotation #73 (`scraped/feminist_voice_dissertation`, word splitting "procreation") no longer reproduces — marked as fixed

**Impact**:
- `scraped/federal_procurement_terms`: Jaccard +0.9pp (54.2→55.0%), SSIM +4.5pp (73.9→78.4%)
- No regressions across all test cases

---

## 2026-03-30: Fixed table bottom border missing for vMerge continuation cells (annotation #70)

**Problem**: In `scraped/polish_municipal_letter`, the gray double-line HR separating the header table from the body text did not span the full table width. The border started at column 2 (x=141.9pt) instead of column 1 (x=39.8pt), leaving a gap where the first column (containing the coat of arms via a floating image) has a vertically merged empty cell.

**Root cause**: The `render_table_row()` function in `src/pdf/table.rs` completely skipped vMerge=Continue cells in its border drawing loop (line 623-625: `if cell.v_merge == VMerge::Continue { continue; }`). The design intent was that the vMerge=Restart cell handles all borders for the merged area using `merge_extra` to extend the bottom position. However, the Restart cell uses ROW 0's border definitions — and for a non-last row, the bottom border is `tblBorders/insideH` (which was `none` in this document), not `tblBorders/bottom` (the table's outer bottom border). So the merged cell's bottom got `insideH=none` instead of the table's `bottom=double`.

**Fix** (`src/pdf/table.rs`):
- In `render_table_row()`, when a vMerge=Continue cell is encountered, still draw its bottom border if present
- Creates a temporary `CellBorders` with only the bottom border set (top/left/right remain default/non-present) to avoid drawing duplicate borders for the merged interior

**Impact**:
- `scraped/polish_municipal_letter`: Jaccard +0.3pp (38.4→38.7%), SSIM +0.7pp (73.9→74.5%)
- No regressions across all test cases

---

## 2026-03-30: Fixed trailing line break creating empty line with proper height (annotation #75)

**Problem**: In `samples/samtale`, the text "Medarbeiderens navn: Ola Normann" was positioned too high on the page compared to the reference. The right column content had cumulative vertical drift of ~47pt by the bottom. The main source was a ~30pt gap difference between "Medarbeiderens egenevaluering" and "Slik gjør du" — our generated output had 52pt gap vs the reference's 82pt.

**Root cause**: Two issues with `<w:br/>` (line break) handling in paragraph height calculation:

1. **Trailing breaks didn't create empty lines**: In `build_paragraph_lines()`, a `<w:br/>` at the end of a paragraph finalized the current line but didn't create the subsequent empty line that Word renders. In the samtale document, paragraph 18 ("Medarbeiderens egenevaluering") ends with `<w:br w:sz="52"/>` (26pt), which should produce a text line PLUS an empty break line. Our code only produced the text line.

2. **Break-created lines used wrong font metrics**: When a `<w:br/>` run has a different font size than the paragraph text (e.g., 26pt break vs 14pt text), Word uses the break run's font metrics for the empty line's height. Our code used the paragraph's text-based `line_h` for all lines, ignoring the break run's font size entirely. The annotation #23 fix correctly prevented break runs from inflating text line heights, but went too far — break-created empty lines still need the break run's metrics.

For the samtale paragraph with the 26pt break:
- Our code: content_h = 1 × 19.64pt (text line only) = 19.64pt
- With fix: content_h = 19.64pt (text) + 36.45pt (break line at 26pt) = 56.09pt

Similarly, paragraph 19 (2 breaks at 10pt) went from 2 lines to 3 lines, gaining ~12pt.

**Fix**:
- `src/pdf/layout.rs`: Added trailing empty line generation in both `build_paragraph_lines()` and `build_tabbed_line()` — when the last line ends with a break, push an additional empty `TextLine` with the break run's `font_size` stored in a new `break_font_size` field
- `src/pdf/mod.rs`: Added `break_run_lhr()` helper to look up a break run's font metrics. Updated the content_h calculation to use the break run's font size and line_h_ratio for break-created lines instead of the paragraph's text-based line_h

**Impact**:
- `scraped/mandated_reporter_child_abuse`: Jaccard **+20.6pp** (26.4→47.0%), SSIM **+16.7pp** (50.5→67.2%)
- `scraped/german_mezzo_soprano_bio`: Jaccard +3.3pp (51.0→54.3%), SSIM +11.7pp (53.2→64.9%)
- `samples/samtale`: Jaccard +1.0pp (10.8→11.9%), SSIM +0.9pp (57.6→58.5%)
- No regressions across all test cases

---

## 2026-03-30: Investigated annotations #25, #30, #66, #77 — systemic issues, no code changes

### Annotation #25 (air_pollution_permit_form — extra "Mellékletek / Prílohy:" text at bottom)

**Problem**: Text "Mellékletek / Prílohy:" appears at the bottom of page 1 in the generated PDF but not in the reference.

**Analysis**: The document has three large page-covering textboxes (`wps:wsp txBox="1"` with `wrapNone`). Textbox 4 (the main form body, 481×714pt) is anchored to a paragraph that's 10 empty body paragraphs below the first one. These empty paragraphs advance the cursor by ~152pt, causing the textbox's anchor position to be much lower than expected:
- Textbox 4 starts at y=645.6 (PDF coords) instead of ~797
- Clip boundary: y=-68.1 (below page bottom)
- "Mellékletek" text at y=52.7 is within the clip bounds, so it renders

In Word, the textbox starts at ~797 with clip_bottom=83.6, so "Mellékletek" (which overflows the textbox height) is clipped. The fundamental issue is that our textbox position depends on the cursor advancement of preceding empty body paragraphs, and the textbox content layout is slightly more compact than Word's (different line heights for the form content inside the textbox).

**Conclusion**: Systemic textbox positioning issue when textbox is paragraph-relative and the anchor paragraph is far from the page top. Would require rethinking how paragraph-relative textbox positions interact with body paragraph cursor advancement.

### Annotation #30 (mongolian_human_rights_law — extra space after "4.1.5")

**Problem**: Annotator reported a space between "4.1.5." and the opening quote character that's not in the reference.

**Analysis**: The text `4.1.5."хүний` is a single word in the layout engine — `split_preserving_spaces` correctly keeps `4.1.5."хүний` as one chunk with no internal spacing. Text extraction from both generated and reference PDFs confirms identical text content: `4.1.5."хүний`. The visual difference is from font glyph width differences (the `"` character renders with slightly more sidebearing in our font compared to the reference's Windows font), not an actual space insertion.

**Conclusion**: Font metric difference, not a rendering bug.

### Annotation #66 (case33 — bullet point vertical drift)

**Problem**: Each bullet point causes slightly larger vertical drift from the reference, accumulating over the page.

**Analysis**: Thorough investigation of:
- **Font metrics**: Calibri on macOS has usWinAscent=1950, usWinDescent=550 (total=2500, UPM=2048). USE_TYPO_METRICS is false. line_h_ratio = 1.2207. For 11pt at 1.15x spacing: line_h = 15.44pt.
- **contextualSpacing**: Working correctly — inter_gap=0 between consecutive ListBullet paragraphs.
- **Line height calculation**: `resolve_line_h(Auto(1.15), 11, Some(1.2207))` = 15.4418pt. Word would round to nearest twip: 15.45pt. Per-line difference: 0.008pt.
- **Paragraph spacing**: space_before/space_after values match DOCX specification exactly.

Over ~15 paragraphs, cumulative floating-point error is ~0.12pt — too small to explain the visible ~2-3pt drift. The remaining drift is likely from Word's internal twip-rounded cursor advancement vs our floating-point arithmetic, compounded by different font_size paragraphs (headings) creating slightly different rounding patterns.

**Conclusion**: Very small floating-point precision differences in line height calculation that accumulate. Would require implementing Word-compatible twip rounding throughout the spacing pipeline.

### Annotation #77 (polish_municipal_letter — text justification)

**Problem**: Body text should be "justified or something like that."

**Analysis**: The paragraphs have `<w:jc w:val="both"/>` (justify), and our `parse_alignment("both")` correctly returns `Alignment::Justify`. The justification rendering code in `build_paragraph_lines` and `render_paragraph_lines` applies extra_per_gap spacing between words on non-last lines. The text IS justified in our output. The visual difference is from slightly different text widths (font metric differences for Times New Roman/Calibri on macOS vs Windows), causing different line breaks and inter-word spacing.

**Conclusion**: Font metric difference causing different line breaks, not a justification logic bug.
