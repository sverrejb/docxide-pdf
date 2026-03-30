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

## Summary of systemic themes

All remaining unfixed annotations fall into these categories:

1. **Font metric precision** (#5, #66, #77, #78): Cumulative differences from macOS vs Windows font metrics, floating-point vs twip rounding
2. **Font availability** (#8, #30): Missing Windows fonts on macOS causing fallback with different metrics
3. **Textbox positioning** (#25): Paragraph-relative textbox anchor position depends on cursor advancement of preceding empty paragraphs

Future priorities to address these:
- Implement Word-compatible twip rounding in the spacing pipeline
- Improve OS/2 font metric handling for less common fonts
- Better paragraph mark height calculation for empty paragraphs
- Investigate Word's exact line pitch / document grid behavior
- Rethink paragraph-relative textbox positioning
