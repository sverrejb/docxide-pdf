# Annotations Progress — Remaining Unfixed Findings

This file tracks annotation findings that were investigated but **not fixed** — they represent systemic issues requiring broader improvements rather than targeted bug fixes.

---

## Annotation #5 — Page break positioning (croatian_grant_guidelines)

**Problem**: TOC starts too high on page 2 — content that should overflow from page 1 stays on page 1 because our page 1 consumes ~17pt less vertical space than Word's.

**Analysis**: The ~2-3pt position difference at the top accumulates across 50+ TOC entries due to subtle text width/wrapping differences. The empty SDT paragraph that should be on page 2 stays on page 1 because of cumulative font metric and line height precision differences. An attempt to preserve `space_before` after natural page overflow caused regressions in 6 cases (up to -39.5pp).

**Root cause**: Cumulative font metric / text width accumulation issue, not a spacing bug. Requires broader improvements to text width calculation.

---

## Annotation #8 — Font availability (east_asia_conference_form)

**Problem**: All content renders on page 1 but the reference has "신청자(申請者):" on page 2. The page overflow should come from natural content height.

**Analysis**: The document uses MalgunGothicBold (맑은 고딕) at 12.96pt with Windows-native metrics. On macOS, falls back to HiraginoSansW3 with different line_h_ratio (1.33), causing different row heights and wrapping. The fallback font's metrics make the table shorter overall than Word's rendering, preventing the expected page overflow.

**Root cause**: Font availability issue on macOS. Malgun Gothic is not installed. Fixing requires either installing the correct fonts or implementing font metric estimation for unavailable fonts.

---

## Annotation #25 — Textbox positioning overflow (air_pollution_permit_form)

**Problem**: Text "Mellékletek / Prílohy:" appears at page bottom in generated but not in reference.

**Analysis**: The document has three large page-covering textboxes. Textbox 4 (481×714pt) is anchored to a paragraph 10 empty paragraphs below the first one. These empty paragraphs advance the cursor by ~152pt, causing the textbox's anchor position to be much lower than expected. In Word, the textbox starts higher so "Mellékletek" (which overflows the textbox height) is clipped. In our output, the textbox starts lower so the overflow text falls within clip bounds and renders.

**Root cause**: Systemic textbox positioning issue when textbox is paragraph-relative and the anchor paragraph is far from the page top. Would require rethinking how paragraph-relative textbox positions interact with body paragraph cursor advancement.

---

## Annotation #30 — Font glyph width differences (mongolian_human_rights_law)

**Problem**: Annotator reported extra space after "4.1.5." before the opening quote.

**Analysis**: The text `4.1.5."хүний` is a single word in the layout engine — no actual space insertion. Text extraction from both generated and reference PDFs confirms identical content. The visual difference is from the `"` character rendering with slightly more sidebearing in our font compared to the reference's Windows font.

**Root cause**: Font metric difference, not a rendering bug.

---

## Annotation #66 — Bullet point vertical drift (case33)

**Problem**: Each bullet point causes slightly larger vertical drift from the reference, accumulating ~2-3pt over the page.

**Analysis**: Thorough investigation of Calibri metrics (usWinAscent=1950, usWinDescent=550, UPM=2048, line_h_ratio=1.2207). contextualSpacing works correctly. Line height: `resolve_line_h(Auto(1.15), 11, Some(1.2207))` = 15.4418pt vs Word's twip-rounded 15.45pt. Per-line difference: 0.008pt. Over ~15 paragraphs, cumulative error is ~0.12pt — too small alone, but compounded by different font_size paragraphs (headings) creating different rounding patterns.

**Root cause**: Floating-point precision differences in line height vs Word's internal twip-rounded cursor advancement. Would require implementing Word-compatible twip rounding throughout the spacing pipeline.

---

## Annotation #77 — Apparent justification issue (polish_municipal_letter)

**Problem**: Body text should be "justified or something like that."

**Analysis**: Paragraphs have `w:jc w:val="both"` and our code correctly applies `Alignment::Justify` with `extra_per_gap` spacing. The text IS justified. The visual difference is from slightly different text widths (font metric differences for Times New Roman/Calibri on macOS vs Windows), causing different line breaks and inter-word spacing.

**Root cause**: Font metric difference causing different line breaks, not a justification logic bug.

---

## Annotation #78 — Text overflow (stem_partnerships_guide)

**Problem**: Text overflow likely from column width calculation differences.

**Root cause**: Font metric / text width differences. Not investigated in detail.

---

## Annotation #31 — Image drop shadow quality (parish_housing_data_profile) — 2026-03-30

**Problem**: Annotator reported "We are not rendering drop shadow for this image" for the map image on page 1.

**Analysis**: The drop shadow WAS being rendered via `draw_image_shadow()` in `pdf/color.rs`, but was barely visible due to two issues in the layered rectangle approximation:
1. Blur radius was halved (`blur_radius * 0.5`), making the shadow extension too compact
2. Linear layer distribution spread opacity evenly, leaving the visible edge too faint

The image has `a:outerShdw blurRad="292100" dist="139700" dir="2700000"` with `srgbClr val="333333"` at 65% alpha — a 45-degree bottom-right shadow with 23pt blur radius.

**Fix**: Improved shadow rendering in two ways:
1. Increased blur multiplier from 0.5 to 0.75, extending the visible shadow to better match the `effectExtent` values in the DOCX (l=13.5pt, t=13.5pt, r=27.8pt, b=26.3pt)
2. Changed from linear to quadratic layer distribution (`t = blur * frac²`), which concentrates more layers near the image edge for a denser, more visible shadow body with gentle outer fade

**Result**: Shadow is now visually present on the bottom-right of the image, matching the reference pattern. Jaccard essentially unchanged (-0.04pp), SSIM slightly improved (+0.1pp). No regressions on any other fixture.

---

## Annotation #38 — Chart rendered too wide (case30) — 2026-03-30

**Problem**: Line chart in case30 extends too far to the right compared to reference. The "Jun" x-axis label is 20pt further right in generated (329.8pt) vs reference (310.3pt).

**Analysis**: Measured text positions in generated vs reference PDFs. The legend area is too small: our code allocated ~95pt for the right margin while Word uses ~115pt. Root cause: the legend swatch width for point charts (line/area/scatter/bubble) was 5.5pt, matching bar chart squares, but Word uses wider line+marker swatches (~20pt). Also, inter-element padding was insufficient.

**Fix**: In `pdf/charts.rs`, changed legend margin calculation:
1. Increased legend swatch width from 5.5pt to 20pt for point charts (line/area/scatter/bubble)
2. Increased inter-element padding (6pt gap + 12pt right padding, was 4pt + 8pt)
3. Raised proportional minimum from 12% to 15% of chart width

**Result**: Improvements across all chart test cases: case29 +5.1pp, case30 +0.5pp, case31 +1.6pp, case53 +3.0pp, sample500kB +0.8pp. One small regression: case52 -1.9pp (within threshold). No regression flags.

---

## Annotation #41 — Table overlaps header on continuation pages (education_consultant_posting) — 2026-03-30

**Problem**: On page 7, a table continuation from the previous page rendered on top of the page header content ("United Nations Children's Fund" and "TERMS OF REFERENCE"), overlapping it. The table should start below the header.

**Analysis**: When a table spans multiple pages, `flush_and_render_headers()` in `pdf/table.rs` resets the cursor to `sp.page_height - sp.margin_top` — the raw top margin. This ignores the page header height entirely. The same issue affected `at_page_top` detection, `available_h`, and `page_content_h` calculations, which all used raw `sp.margin_top` / `sp.margin_bottom` instead of the effective values that account for header/footer content height.

Meanwhile, the main render loop in `pdf/mod.rs` consistently uses `effective_slot_top()` and `compute_effective_margin_bottom()` for all page break decisions.

**Fix**: In `pdf/table.rs`:
1. Changed `flush_and_render_headers()` to use `effective_slot_top(sp, false, ctx)` instead of `sp.page_height - sp.margin_top`
2. Changed `at_page_top` check to compare against `effective_slot_top()` instead of raw margin
3. Changed `available_h` to use `compute_effective_margin_bottom()` instead of `sp.margin_bottom`
4. Changed `page_content_h` to use effective top - effective bottom
5. Changed the split-row loop's `avail` calculation to use `compute_effective_margin_bottom()`

**Result**: `education_consultant_posting` Jaccard +2.3pp (12.1% → 14.4%), SSIM +8.0pp (24.2% → 32.1%). `go_math_grade4_guide` Jaccard +2.4pp (27.7% → 30.0%), SSIM +1.2pp. No regressions. Table content now properly appears below headers on continuation pages.

**Note**: The header rendering on this page still has garbled text ("nited ations Childrens und" instead of "United Nations Children's Fund") — this is a separate font encoding issue unrelated to the table positioning fix.

---

## Annotation #39 — Scatter/Bubble chart plot area too narrow (case31) — 2026-03-30

**Problem**: Scatter chart plot area was narrower than Word's reference. The annotation reported the graph as "too wide" but measurements showed the opposite: our plot_w was ~252pt vs reference ~259pt, and our legend was 23pt closer to the plot area.

**Analysis**: The margin_right calculation for right-side legends used `swatch_w = 20.0` for ALL point charts (Line/Area/Scatter/Bubble), matching the width of line+marker legend swatches. However, Scatter and Bubble charts render legend swatches as small marker dots (`SwatchStyle::Marker`, swatch_size=5.5) without any line extension, so 20pt was ~14.5pt more than needed. This over-allocation compressed the plot area.

Measurements confirmed: generated margin_right=76.2pt vs reference ~69pt; generated plot_w=252.1pt vs reference ~259pt.

**Fix**: In `pdf/charts.rs`, differentiated the legend swatch width reservation by chart type:
- Line/Area charts: keep `swatch_w = 20.0` (accounts for line+marker swatches)
- Scatter/Bubble/Bar charts: use `swatch_w = 5.5` (marker dot or rect only)

This brings margin_right from 76.2pt down to ~62pt for scatter charts, producing a wider plot area that better matches Word's layout.

**Result**: case31 Jaccard +1.4pp (58.2% → 59.5%), SSIM +2.5pp (75.0% → 77.6%). No regressions on case29, case30, sample500kB, or any other fixture.

---

## Annotation #44/#74 — Line chart data points at wrong x-positions (case30) — 2026-03-30

**Problem**: Line chart data points started and ended at the plot area edges instead of aligning with category labels. The "Jan" data point was at the left axis and "Jun" at the right edge, but category labels (placed at section midpoints) were inset from the edges.

**Analysis**: The Line chart code used `cat_w = plot_w / (num_categories - 1)` to space data points edge-to-edge. But category labels were correctly placed at section midpoints using `(ci + 0.5) * group_w` where `group_w = plot_w / num_categories`. This caused data markers to misalign with their category labels.

The Area chart uses a different convention — data points at edges to fill the entire plot width — so the fix was only applied to Line charts.

**Fix**: In `pdf/charts.rs`, changed Line chart data point placement from edge-to-edge (`plot_x + ci * plot_w/(n-1)`) to category midpoints (`plot_x + (ci + 0.5) * plot_w/n`), matching the existing category label positions.

**Result**: case30 Jaccard +1.9pp (82.3% → 84.2%), SSIM +1.1pp (82.8% → 83.9%). No regressions on any fixture.

---

## Annotation #46 — Header floating image not counted in header height (czech_municipal_grant_form) — 2026-03-30

**Problem**: On page 2, numbered text overlapped the header logo. The body content started too high, rendering on top of the coat-of-arms logo in the page header.

**Analysis**: The header contains a floating image (45.4×48.7pt) with `WrapType::Square` wrap. The `compute_header_height()` function only counted floating images with `WrapType::TopAndBottom`, ignoring `Square`/`Tight`/`Through` wrap types. Since the logo uses Square wrapping, its height (~48.7pt) was not included in the header height calculation, so `effective_slot_top()` placed body content too high.

**Fix**: In `pdf/header_footer.rs`, changed the floating image filter in `compute_header_height()` from `matches!(fi.wrap_type, WrapType::TopAndBottom)` to `!matches!(fi.wrap_type, WrapType::None)`. In headers/footers, any text-displacing wrap type (TopAndBottom, Square, Tight, Through) should contribute to the header height. Also clamped negative offsets to 0 to prevent negative height contributions.

**Result**: czech_municipal_grant_form SSIM +4.1pp (30.0% → 34.1%), Jaccard -0.2pp (noise). education_consultant_posting also improved: SSIM +0.6pp. No regressions.

---

## Annotation #52 — Header textbox TopAndBottom cursor not advancing (education_consultant_posting) — 2026-03-30

**Problem**: "TERMS OF REFERENCE" heading in the header was rendered 20.5pt too high (y=80.50 vs reference y=101.00). The annotator noted the table below was correctly positioned, confirming this was a header-internal spacing issue.

**Analysis**: The page header contains a TopAndBottom textbox (mc:AlternateContent with DrawingML/VML textbox) at page-relative y=69pt, 13.5pt tall. This textbox holds the subtitle text ("United Nations Children's Fund | Pakistan Country Office"). The textbox acts as a vertical spacer — content below it should start after its bottom edge (82.5pt from top). Two issues found:

1. **mc:AlternateContent not collected in headers**: The `parse_header_footer_xml` function skipped non-WML namespace elements, missing `mc:AlternateContent` blocks that wrap paragraphs in mc:Choice/mc:Fallback. Fixed by using `collect_block_nodes()` (which now also handles mc:AlternateContent via mc:Fallback).

2. **Header rendering cursor didn't advance past TopAndBottom textboxes**: When a paragraph with a TopAndBottom textbox had empty text, the cursor only advanced by `line_h` (14.4pt), ignoring the textbox's vertical extent. The textbox bottom at 82.5pt from page top was never reached by the cursor (which was at ~65pt from top after para 1). Fixed by computing the needed advance to clear the textbox bottom in PDF coordinates.

**Fix**:
1. In `src/docx/mod.rs`: Extended `collect_block_nodes()` to handle `mc:AlternateContent` by descending into `mc:Fallback` children.
2. In `src/docx/headers_footers.rs`: Replaced manual block iteration with `collect_block_nodes(root)`.
3. In `src/pdf/header_footer.rs`: For empty paragraphs with TopAndBottom textboxes, compute the cursor advance needed to clear the textbox bottom (page-relative or paragraph-relative) and use the maximum of that and `line_h`.

**Result**: `education_consultant_posting` Jaccard +1.2pp (14.5% → 15.7%), SSIM +1.9pp (32.7% → 34.5%). "TERMS OF REFERENCE" now at y=100.86 (reference: 101.00, 0.14pt difference). No regressions on any fixture.

---

## Annotation #48 — Floating table breakpoint (croatian_grant_guidelines) — 2026-03-30 (investigated, not fixed)

**Problem**: The green "Važno!" floating table on page 4 splits one paragraph too early. Reference fits 8 items (paragraphs 0-7) on the page, but we fit only 7 (0-6).

**Analysis**: Cumulative item heights: items 0-7 need 386.00pt, but only 383.80pt is available. The 2.2pt shortfall traces to the floating table starting ~16pt lower in our rendering (at 387pt from top vs 371pt in reference). This is the same systemic vertical shift as annotation #47 — accumulated text height differences from pages 1-3 push the floating table down on page 4.

**Root cause**: Consequence of the systemic Y-shift (annotation #47). Not independently fixable.

---

## Annotation #54 — Table cells wrong (go_math_grade4_guide) — 2026-03-30

**Problem**: The "Chapter 10 Rule of Thumb" table's content row rendered as one cell spanning full width instead of two cells (rule text + rationale).

**Analysis**: The table has a 3-column grid (10, 8460, 5940 twips) but Row 1 has only 2 cells with no `w:gridSpan` attribute. Without explicit gridSpan, the parser defaulted to span=1, mapping Cell 0 to the tiny 10-twip column (0.5pt) instead of spanning columns 0+1 (8470 twips = 423pt) as intended by its tcW width.

**Fix**: In `src/docx/tables.rs`, when `w:gridSpan` is not explicit, the parser now infers the span by finding the number of consecutive grid columns whose cumulative width best matches the cell's declared tcW width. For the problematic cell: tcW=8460 matches columns 0+1 (10+8460=8470, diff=0.5pt) much better than column 0 alone (10, diff=422.5pt), so gridSpan=2 is inferred.

**Result**: go_math_grade4_guide Jaccard +0.2pp (30.0% → 30.3%). No regressions on any fixture. The Rule of Thumb table now correctly shows two cells matching the reference layout.

---

## Annotations #49, #50, #51 — Already resolved (croatian_grant_guidelines) — 2026-03-30

**#49** (Links overflow page width): Verified all URL text right edges ≤ 524.46pt (within right margin). URLs wrap correctly via `unicode_linebreak` segments.

**#50** (Footer not rendered): Footers ("Stranica N") render on ALL 20 generated pages with position matching reference within 0.2pt.

**#51** (Footer rendered wrong): Footer positioning matches reference. No horizontal line exists in the DOCX footer XML.

---

## Summary of systemic themes

All remaining unfixed annotations fall into these categories:

1. **Font metric precision** (#5, #66, #77, #78): Cumulative differences from macOS vs Windows font metrics, floating-point vs twip rounding
2. **Font availability** (#8, #30): Missing Windows fonts on macOS causing fallback with different metrics
3. **Textbox positioning** (#25): Paragraph-relative textbox anchor position depends on cursor advancement of preceding empty paragraphs

## Annotation #56 — Text wrapping around wide image (indonesian_benchmarking_guide) — 2026-03-30

**Problem**: On page 7 (0-indexed page 6), text wrapped to the right of a `wrapSquare wrapText="bothSides"` floating image (251pt wide, left-aligned to margin). The reference shows text flowing below the image, not beside it.

**Analysis**: The image (Picture 4, "Capture modul 5.4.PNG") is 251pt wide in a 451pt text area (55.6% of width). With `wrapSquare wrapText="bothSides"` and left alignment, there's ~191pt of space to the right. Our code was treating it as a wrapping image and flowing text beside it. Word treats wide images as effectively TopAndBottom, not wrapping text beside them.

The existing threshold for treating Square/Tight/Through images as TopAndBottom was 90% of text width — far too permissive. Analysis of all wrapSquare images across the test corpus showed a clear gap: all images at ≤44.2% correctly wrap text, while images at ≥55.6% should not wrap.

**Fix**: In `src/pdf/mod.rs`, lowered the width threshold from 90% to 50% of text_width in three locations:
1. The `reserve` check for content_h calculation (line ~1349)
2. The float zone setup in the self-wrapping block (line ~967) — added `fi.image.display_width < text_width * 0.5` condition
3. The float zone setup during rendering (line ~1755) — same condition as a match guard
4. The lookahead for next-paragraph wrapping (line ~1204) — same condition
5. The textbox reserve check — same 50% threshold

**Result**: Text now flows below the image on page 7, matching the reference. Jaccard -4.1pp (40.5% → 36.4%) due to cascading page breaks from the layout change. No regressions on any other fixture (brazilian_logistics_study, german_mezzo_soprano_bio, vaccines_history_chapter, czech_expert_witness_law, go_math_grade4_guide all unchanged).

---

## Annotation #58 — Too much empty space between image and caption (brazilian_logistics_study) — 2026-03-31

**Problem**: On page 8, too much empty space (~120pt) between a floating image and its caption "Fonte: ALARCOM, (2019)."

**Analysis**: The paragraph contains a `wp:anchor` floating image (364.5pt wide, 146.9pt tall) with `wrapSquare wrapText="bothSides"`. The image occupies 80.4% of the 453.5pt text area. Between the image paragraph and the caption, there are 8 empty paragraphs with 1.5x line spacing at 10pt (~17pt each).

Previously, wide wrapSquare images (≥50% of text width) were treated as TopAndBottom — their height was added to content_h, advancing the cursor past the image. The 8 empty paragraphs then added ~120pt MORE space on top of the image height, creating excessive vertical space.

In Word, these empty paragraphs exist within the image's vertical extent (flowing beside it in the narrow wrap zones), contributing no extra space below the image.

**Fix**: Three changes in `src/pdf/mod.rs`:
1. **content_h reservation**: Square/Tight/Through wrapping images no longer add their height to content_h. Their height is tracked via `float_overflow_h` for page break decisions only.
2. **Float zone creation**: All Square/Tight/Through images now create float zones (removed the `< 0.5 * text_width` guard). This means subsequent paragraphs are aware of the image's vertical extent.
3. **Dynamic MIN_WRAP_WIDTH**: Changed from fixed 72pt to `(col_w * 0.5).max(72.0)`. This prevents text from wrapping beside wide images (where combined wrap space < 50% of column width) while still allowing narrow images to wrap correctly.
4. **Empty paragraph absorption**: When a float zone has insufficient wrap space, empty paragraphs are no longer pushed below the zone. They advance the cursor naturally within the zone, effectively being "absorbed" by the image's vertical extent. Non-empty paragraphs are still pushed below.

**Result**: brazilian_logistics_study Jaccard +1.3pp (16.3% → 17.6%), SSIM +3.1pp (27.4% → 30.5%). indonesian_benchmarking_guide also improved: Jaccard +4.5pp (36.4% → 40.8%), SSIM +4.6pp (50.3% → 54.9%). Minor noise on case41 (-0.3pp) and sample500kB (-0.2pp/-0.7pp). No regression flags.

---

## Annotation #59 — Too much white space above image (brazilian_logistics_study) — 2026-03-31 (investigated, not fixed)

**Problem**: On page 9 (0-indexed page 8), excessive white space above "Figura 2" image and caption.

**Analysis**: Generated page 8 has 26 text lines vs reference's 22 — 4 extra lines from different line breaks in justified Arial text. These 4 extra lines consume space at the bottom of page 8 that should hold 7 empty paragraphs (spacers before "Figura 2"). About 4 empty paragraphs overflow to page 9, creating ~83pt of visible white space above the content.

**Root cause**: Systemic text wrapping / line length differences. Not independently fixable.

---

## Annotation #60 — Wrong font on chart labels (sample500kB) — 2026-03-31

**Problem**: Annotator reported chart labels use wrong (serif) font.

**Analysis**: Both reference and generated PDFs use Aptos (sans-serif) for chart labels. Verified via `mutool info`: reference has `AAAAAN+Aptos`, generated has `Aptos`. The chart font correctly defaults to the theme minor font (Aptos). The annotation was likely made against an older build.

**Result**: Already fixed. Marked as fixed in annotations.json.

---

## Annotation #67 — DRAFTING NOTE overlaps MCL logo (uk_commercial_lease_template) — 2026-03-31

**Problem**: "[DRAFTING NOTE: THIS LEASE IS INTENDED..." text was 25pt too low on the cover page, overlapping the MCL logo (positioned via footer with -115pt negative offset).

**Analysis**: The cover page is a single 8-row table. Row 7 (last row) contains a nested 5-row table inside a single cell. After the nested table, the cell has a mandatory empty end-of-cell paragraph (SHNormal style). This trailing paragraph added 12.3pt (line_h for Arial 10pt at 1.1× spacing) + 9pt (space_after from Normal style) = 21.3pt to the row height. In Word, when a cell's content is solely a nested table plus the mandatory end-of-cell mark, the trailing paragraph mark contributes only the ~0.5pt glyph height (already accounted for in the row-height addition), not a full line with spacing.

**Fix**: In `src/pdf/table_layout.rs`, added two conditions:
1. When computing cell content height, if the previous block was a nested table and this is the first empty paragraph after it (`prev_was_nested_table && para_idx == 1`), skip adding `line_h`.
2. When adding trailing `space_after` to cell height, check if the cell is exactly `[nested_table, empty_paragraph]` (`sole_table_plus_mark`). If so, suppress the space_after.

**Result**: DRAFTING NOTE moved from y=640.3 to y=618.7 (reference: y=614.9, difference now 3.8pt — within cumulative font metric tolerance). uk_commercial_lease_template SSIM +0.3pp (33.9% → 34.2%). One expected side-effect: cases/case51 SSIM -2.7pp (nested table test case with a [TABLE, P("")] cell that now renders shorter; baseline updated).

---

## Annotation #72 — Lines between cells wrong (turkish_ancient_religions_plan) — 2026-03-31

**Problem**: Horizontal borders were drawn through vertically merged cells in the left "Haftalar" column, creating lines that should not exist. In Word, merged cells appear as one tall cell without internal horizontal borders.

**Analysis**: The table has `insideH: single` and cell-level `tcBorders` with `dotted` borders. The left column uses `vMerge` to merge groups of 3-4 rows per "Hafta". Four border rendering paths (`render_table_row`, `render_nested_table`, `render_partial_row`, `render_header_footer_table`) handled `VMerge::Continue` cells incorrectly:
- `render_table_row`: Drew the bottom border of every continuation cell, creating lines through the merged area
- The other three paths: Drew all borders for continuation cells without any vMerge handling

Additionally, when a merge group extended to the table edge, the restart cell's `borders.bottom` used `insideH` (set during parsing based on the restart row's position) instead of the table-edge border.

**Fix**: Two-part fix:
1. **Parsing** (`src/docx/tables.rs`): After border resolution, a new pass propagates the last continuation cell's bottom border to the restart cell. This ensures the merged region uses the correct bottom edge style (e.g., table-edge double border rather than interior insideH).
2. **Rendering** (`src/pdf/table.rs`): All four border rendering paths now skip `VMerge::Continue` cells entirely. For `VMerge::Restart` cells, borders extend to `effective_bottom` (row_bottom minus merge_extra), covering the full merged region.

**Result**: turkish_ancient_religions_plan Jaccard +0.2pp (22.3% → 22.5%). Also improved: italian_project +0.6pp, case15 +0.1pp. Minor: japanese_interlibrary_loan -0.4pp (corrected border rendering in tables with vMerge). No SSIM regressions.

---

## Annotation #76 — List label font size boosting line height (samples/samtale) — 2026-03-31

**Problem**: In paragraphs with large list labels (e.g., 20pt numbered "6." on 10pt text), "I stor grad." answer text floated ~17pt above the grey paragraph bottom border instead of the correct ~5pt gap.

**Analysis**: The document uses a 2-column layout with numbered Q&A sections. Each question uses `numId=6` at `ilvl=0` with `w:sz w:val="40"` (20pt) for the list number label, while paragraph text is 10pt. The `content_h` calculation in `pdf/mod.rs` boosted `first_line_h` to `resolve_line_h(Auto(1.0), 20.0, tallest_lhr)` ≈ 24.4pt when `list_label_font_size > font_size`, adding ~12.2pt extra per paragraph. Word does not do this — list labels sit in the margin and don't affect line height.

For Q6 (right column): content_h was 24.4 + 12.2 + 12.2 = 48.8pt instead of correct 36.6pt. The border at `box_bottom = slot_top - content_h - bdr_bottom_pad` was placed 12.2pt too low, creating a 17pt gap between "I stor grad." baseline and the border (reference: ~5.8pt).

**Fix**: In `src/pdf/mod.rs`, removed the `first_line_h` boost for `list_label_font_size > font_size`. Now `first_line_h` always starts at `line_h` (based on text font), and only `label_boosted_line_h()` adjusts it for labels with font size close to text (within 1pt difference, e.g. bullet symbols).

**Result**: Q6 text-to-border gap: 17.0pt → 4.8pt (reference: 5.8pt). Text boundary score improved +7.7pp (20.8% → 28.6%). Jaccard -2.2pp (11.9% → 9.7%) and SSIM -15.3pp due to secondary column-break effect: more compact Q1-Q5 spacing caused "Vurdering" heading to fit in column 1 instead of flowing to column 2. No regressions on any other fixture.

---

## Annotation #79 — Garbled header text (education_consultant_posting) — 2026-03-31

**Problem**: On page 7 (0-indexed page 6), the header text "United Nations Children's Fund | Pakistan Country Office" rendered with missing characters: "□nited □ations Children□s □und". Letters like U, N, F, and the curly apostrophe were replaced with `.notdef` glyphs.

**Analysis**: The header contains a VML textbox (`mc:AlternateContent` → VML shape) with the title text in Calibri Bold at color #00B0F0. The text is rendered via `render_header_footer_textboxes()` in `pdf/header_footer.rs`, which correctly iterates textbox paragraphs and their runs.

However, the font subsetting pipeline in `src/pdf/fonts.rs` (`collect_all_runs()`) collects characters from all runs to determine which glyphs to include in each subsetted font. For **body** paragraphs, it used `para_runs_with_textboxes(para)` which recursively includes runs from nested textbox paragraphs. But for **header/footer** paragraphs, it only used `p.runs.iter()` — skipping textbox runs entirely.

Characters that appeared *only* in header textboxes (and nowhere else in the document) were never added to the font's `char_to_gid` mapping. During PDF rendering, `encode_as_gids()` mapped these unknown characters to glyph ID 0 (`.notdef`), producing garbled output.

The same issue affected **footnote** paragraphs, which also used `p.runs.iter()` instead of `para_runs_with_textboxes(p)`.

**Fix**: In `src/pdf/fonts.rs`, changed both header/footer and footnote run collection to use `para_runs_with_textboxes(p)` instead of `p.runs.iter()`, ensuring all textbox-nested runs are included in font character collection.

**Result**: education_consultant_posting Jaccard +0.3pp (15.7% → 16.0%), SSIM +0.9pp (34.5% → 35.4%), text boundary +23.4pp (43.9% → 67.3%). Header text now renders correctly. No regressions on any fixture.

---

## Annotation #80 — Empty space below image too large (indonesian_benchmarking_guide) — 2026-03-31

**Problem**: On page 7 (0-indexed page 6), excessive empty space (~28pt) between the floating questionnaire image and "3. Metode Benchmarking" heading below it.

**Analysis**: The floating image (Capture modul 5.4.PNG, 251×374pt, wrapSquare bothSides) creates a float zone covering most of the page. Between the image's anchor paragraph and "3. Metode Benchmarking", there are 5 completely empty paragraphs plus 1 paragraph containing only a `w:br` (soft line break) element with no text.

The float zone absorption logic in `pdf/mod.rs` checked `is_text_empty(&p.runs)` to decide whether a paragraph should be absorbed within the float zone. `is_text_empty()` returns false for runs with `is_line_break=true`, so the br-only paragraph was NOT absorbed. Instead, it triggered the float-zone push (setting slot_top to fz.bottom_y) and then added its own content_h of 28.2pt (2 lines × 14.1pt) below the image — creating visible empty space.

**Fix**: In `src/pdf/mod.rs`, expanded the `is_empty_para` check in the float zone absorption code to treat line-break-only runs as empty content. Changed from `is_text_empty(&p.runs)` to an inline check that allows `r.is_line_break` (in addition to vanished and truly empty runs), so paragraphs with only soft line breaks are absorbed within the float zone just like truly empty paragraphs.

**Result**: indonesian_benchmarking_guide Jaccard +1.8pp (40.8% → 42.6%), SSIM +1.7pp (54.9% → 56.6%). "3. Metode Benchmarking" heading moved from 476pt to 448pt from top (reference: 462pt). More content now fits on page 7 (including section 3.1.3), matching the reference layout. No regressions on any fixture.

---

Future priorities to address these:
- Implement Word-compatible twip rounding in the spacing pipeline
- Improve OS/2 font metric handling for less common fonts
- Better paragraph mark height calculation for empty paragraphs
- Investigate Word's exact line pitch / document grid behavior
- Rethink paragraph-relative textbox positioning
