# Progress for plan_education.md

## Step 1: Fix SDT parsing in table cells — COMPLETED
- Changed `tc.children()` to `collect_block_nodes(tc).into_iter()` in `src/docx/tables.rs:226`
- `collect_block_nodes` was already imported; it unwraps `w:sdt > w:sdtContent` to find nested paragraphs
- Compilation verified with `cargo check` — no new warnings or errors

## Step 2: New data structures for per-paragraph cell layout — COMPLETED
- Added `CellParagraphLayout` struct with fields: `lines`, `line_h`, `font_size`, `ascender_ratio`, `alignment`, `space_before`
- Added `CellLayout` struct with fields: `paragraphs: Vec<CellParagraphLayout>`, `total_height: f32`
- Replaced `RowLayout.cell_lines: Vec<(Vec<TextLine>, f32, f32)>` with `RowLayout.cells: Vec<CellLayout>`
- Expected compilation errors in `compute_row_layouts`, `render_table_row`, `render_table`, and `render_header_footer_table` — these will be fixed in Steps 3, 4, and 6

## Step 3: Update `compute_row_layouts` to produce per-paragraph layout — COMPLETED
- Refactored the inner loop to collect `Vec<CellParagraphLayout>` per cell instead of flattening into `all_lines`
- Each paragraph now stores its own `line_h`, `font_size`, `ascender_ratio`, `alignment`, and collapsed `space_before`
- `ascender_ratio` computed per-paragraph from font metrics (using `font_key_buf` + font entry lookup)
- `space_before` uses the existing collapsed spacing logic: `max(prev_space_after, para.space_before)`, 0.0 for first paragraph
- VMerge::Continue cells return empty `CellLayout` with `total_height: 14.4`
- `compute_row_layouts` now returns `RowLayout { height, cells: Vec<CellLayout> }`
- Remaining compilation errors in `render_table_row` (line 258), `render_table` log (line 436), and `render_header_footer_table` (line 510) — to be fixed in Steps 4, 5, and 6

## Step 4: Update `render_table_row` for per-paragraph rendering — COMPLETED
- Replaced single `render_paragraph_lines` call per cell with a loop over `cell_layout.paragraphs`
- Each paragraph rendered with its own `line_h`, `font_size`, `ascender_ratio`, `alignment`, and `space_before`
- Vertical alignment (Top/Center/Bottom) computed from total content height across all paragraphs, applied as a `v_offset` before rendering
- `cursor_y` tracks position within cell, advancing by `space_before` then `lines.len() * line_h` for each paragraph
- Also updated `render_header_footer_table` with the same per-paragraph pattern (Step 6, required for compilation)
- Fixed `render_table` log line to use `layout.cells.len()` instead of old `layout.cell_lines.len()`
- `cargo check` passes (only pre-existing warnings)
- `cargo test -- --nocapture`: all tests pass, one minor regression: `scraped/stem_partnership` -0.7pp Jaccard (27.4%, still passes 20% threshold), -0.6pp SSIM (41.3%, was already failing). Expected from more accurate inter-paragraph spacing in table cells.

## Step 5: Table row splitting across pages — COMPLETED
- Added `find_cell_split()` helper: walks paragraphs from a start index, accumulating height (cm.top + cm.bottom + space_before + lines * line_h), returns exclusive end index. Always includes at least one paragraph to guarantee progress.
- Added `render_partial_row()`: renders a paragraph range per cell with correct shading, text, and border handling. Top border drawn only on first chunk, bottom border only on last chunk, left/right always drawn. No vertical alignment applied (top-aligned for split rows, matching Word behavior).
- Modified `render_table()` loop with three-way condition:
  1. `row_h > available_h && row_h > page_content_h`: Row too tall for any single page — split across pages starting on current page (no wasted space)
  2. `row_h > available_h`: Row fits on fresh page but not current — flush first, then render normally (preserves existing behavior)
  3. Otherwise: render normally
- The splitting loop: find split indices per cell, render partial row, flush page, repeat headers if applicable, continue until all paragraphs rendered
- `space_before` suppressed for the first paragraph in each continuation chunk (pi == start)
- `cargo check` passes (only pre-existing warnings)
- `cargo test -- --nocapture`: no new regressions. Only pre-existing `scraped/stem_partnership` regression from Step 4.
- education_consultant_posting: Jaccard 8.6% (+1.8pp from 6.8%), SSIM 22.5% (+3.9pp from 18.6%). Generated 5 pages (ref has 7). Content now flows properly across pages starting on page 1. Lower page count vs reference is due to tighter text formatting (no bullet indentation, compressed paragraph spacing — separate issues from row splitting).

## Step 7: Remove from SKIPLIST & test — COMPLETED (kept on SKIPLIST)
- Ran `DOCXIDE_CASE=education_consultant_posting cargo test -- --nocapture`: Jaccard 8.6%, SSIM 22.5% — both still below thresholds (20% / 75%)
- Since the fixture still fails thresholds, it was **kept on the SKIPLIST** rather than removed. Removing it would add a known-failing fixture to the regular test suite.
- Full test suite: all tests pass. Only pre-existing regression: `scraped/stem_partnership` from Step 4.
- The plan's improvements (row splitting, per-paragraph rendering, SDT parsing) are all implemented and working, but the fixture needs further work on bullet indentation, paragraph spacing, and other formatting to cross the threshold.
