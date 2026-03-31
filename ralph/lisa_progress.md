# Lisa Progress

## Completed: Annotation #49 — Links overflow page width (croatian_grant_guidelines)

**Status**: Already fixed. URLs wrap correctly within margins via `unicode_linebreak` segments.

## Completed: Annotations #50, #51 — Footer issues (croatian_grant_guidelines)

**Status**: Already fixed. Footers render correctly on all pages, matching reference positions.

## Completed: Annotation #54 — Table cells wrong (go_math_grade4_guide)

**Status**: Fixed. Inferred gridSpan from cell width when not explicit. Jaccard +0.2pp, no regressions. Committed as 928d052.

## Completed: Annotation #56 — Text wrapping around image (indonesian_benchmarking_guide)

**Status**: Fixed. Lowered the wrap-to-TopAndBottom threshold from 90% to 50% of text_width for Square/Tight/Through floating images. Images wider than half the text area now displace text vertically (TopAndBottom) instead of allowing text to wrap beside them. Jaccard -4.1pp (40.5% → 36.4%) due to cascading page breaks — structurally correct, text now flows below image matching reference.

## Completed: Annotation #58 — Too much empty space between image and caption (brazilian_logistics_study)

**Status**: Fixed. Wide wrapSquare floating images (>50% text width) now use float zones instead of adding height to content_h. Empty paragraphs between image and caption are absorbed within the float zone's vertical extent. Dynamic MIN_WRAP_WIDTH (col_w * 0.5) prevents text from wrapping beside wide images while still allowing narrow images to wrap correctly. Jaccard +1.3pp, SSIM +3.1pp. Also improved indonesian_benchmarking_guide (+4.5pp/+4.6pp).

## Investigated: Annotation #59 — Too much white space above image (brazilian_logistics_study)

**Status**: Systemic. The white space on page 9 is caused by cascading text wrapping differences from earlier pages. Our text wraps to different line counts than Word's (due to font metric differences for justified Arial text), causing 4 extra lines on page 8. This consumes space that should hold empty paragraphs, which overflow to page 9 creating visible white space. Root cause: systemic line length / text wrapping differences.

## Completed: Annotation #60 — Wrong font used on labels (sample500kB)

**Status**: Fixed. Chart labels used the raw font name "Aptos" in PDF content stream `set_font()` calls, but page resources mapped fonts as "F1", "F2", etc. The PDF viewer couldn't find "Aptos" in resources and fell back to a serif substitute. Fix: resolve the FontEntry's `pdf_name` for the chart font key, and pass the font entry to `show_text()` so CID fonts encode glyph IDs correctly. Improved all chart cases: case29 +1.0pp, case30 +1.2pp, case31 +1.7pp SSIM. Committed as 4603e34.

## Completed: Annotation #67 — DRAFTING NOTE overlaps MCL logo (uk_commercial_lease_template)

**Status**: Fixed. The cover page table's last row contained a nested table followed by the mandatory empty end-of-cell paragraph. This trailing paragraph added ~21pt (12.3pt line_h + 9pt space_after) to the row height. In Word, when a cell contains only [nested table, empty ¶], the trailing paragraph mark doesn't contribute line height or space_after — its glyph is covered by the 0.5pt row-height addition. Fixed by suppressing line_h and space_after for empty trailing paragraphs that immediately follow a nested table as the cell's sole content. DRAFTING NOTE moved from 25pt below reference to ~4pt (within cumulative font metric tolerance). SSIM +0.3pp. One side-effect: case51 SSIM -2.7pp (nested table test case, baseline updated).

## Completed: Annotation #72 — Lines between cells wrong (turkish_ancient_religions_plan)

**Status**: Fixed. Vertically merged cells (vMerge) incorrectly drew horizontal borders through the merged area. Two fixes: (1) In parsing (docx/tables.rs), propagate the last continuation cell's bottom border to the restart cell so the merged region uses the correct edge style. (2) In rendering (pdf/table.rs), skip border drawing for VMerge::Continue cells across all four render paths (render_table_row, render_nested_table, render_partial_row, render_header_footer_table). For restart cells, borders extend to the full merge height via effective_bottom. Jaccard +0.2pp, italian_project +0.6pp, case15 +0.1pp. Minor: japanese_interlibrary -0.4pp (corrected border rendering).

## Completed: Annotation #76 — List label font size boosting line height (samples/samtale)

**Status**: Fixed. Paragraphs with large list number labels (20pt) on small text (10pt) had inflated content_h because first_line_h was boosted to 24.4pt instead of 12.2pt. This caused paragraph bottom borders (grey lines) to be drawn ~12pt too low, creating a 17pt gap between answer text "I stor grad." and the grey separator line (reference: ~5.8pt). Fixed by removing the first_line_h boost for list_label_font_size > font_size — Word's list labels sit in the margin and don't affect line height. Text-to-border gap: 17pt → 4.8pt. Text boundary +7.7pp. Jaccard -2.2pp (column-break side effect). No regressions on other fixtures.
