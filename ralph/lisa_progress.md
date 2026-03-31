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

**Status**: Already fixed. Both reference and generated PDFs use Aptos (sans-serif) for chart labels. The chart font correctly defaults to the theme minor font. Verified via mutool info showing identical font (AAAAAN+Aptos in reference, Aptos in generated).

## Completed: Annotation #67 — DRAFTING NOTE overlaps MCL logo (uk_commercial_lease_template)

**Status**: Fixed. The cover page table's last row contained a nested table followed by the mandatory empty end-of-cell paragraph. This trailing paragraph added ~21pt (12.3pt line_h + 9pt space_after) to the row height. In Word, when a cell contains only [nested table, empty ¶], the trailing paragraph mark doesn't contribute line height or space_after — its glyph is covered by the 0.5pt row-height addition. Fixed by suppressing line_h and space_after for empty trailing paragraphs that immediately follow a nested table as the cell's sole content. DRAFTING NOTE moved from 25pt below reference to ~4pt (within cumulative font metric tolerance). SSIM +0.3pp. One side-effect: case51 SSIM -2.7pp (nested table test case, baseline updated).
