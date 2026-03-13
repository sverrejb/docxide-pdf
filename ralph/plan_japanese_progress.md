# Progress for ralph/plan_japanese.md

## Step 1: Model — Add `TextDirection` enum and cell field ✅
- Added `TextDirection` enum (`LrTb`, `TbRl`, `BtLr`) with `Default` derive to `src/model.rs` (after `CellVAlign`)
- Added `text_direction: TextDirection` field to `TableCell` struct
- Updated `TableCell` construction in `src/docx/tables.rs` and `src/docx/alt_chunk.rs` with `TextDirection::default()`
- Added `TextDirection` to import in `tables.rs`
- Verified: `cargo check` passes

## Step 2: Parse `w:textDirection` in table cell properties ✅
- Added `textDirection` parsing in `src/docx/tables.rs` after `v_align` parsing (line ~182)
- Maps both legacy values (`tbRlV`, `tbRl`, `rlV`, `rl`) and new values (`tbV`, `tb`) to `TbRl`
- Maps `btLr`, `lr`, `lrV`, `lrTbV` to `BtLr`
- All other values default to `LrTb`
- Replaced `TextDirection::default()` in `TableCell` construction with the parsed value
- Verified: `cargo check` passes

## Step 3: Render rotated text in table cells ✅
- Modified `src/pdf/table.rs`:
  - Added `TextDirection` import and `text_direction` field to `CellLayout`
  - In `compute_row_layouts`: rotated cells (TbRl/BtLr) use unlimited width (10000pt) for single-line layout; `total_h` overridden to `cm.top + cm.bottom + max_line_width` (text width becomes row height)
  - In `render_table_row`: TbRl cells rendered with `save_state/cm(0,-1,1,0,e,f)/restore_state` — 90° CW rotation maps pre-transform x→screen-down, y→screen-right
  - Text block centered horizontally within column width via `h_offset`
  - Vertical alignment (vAlign) applied along the flow direction via `v_offset` in pre-transform x
- Scores: Jaccard 3.4% → 5.5% (+2.1pp), SSIM 18.3% → 19.3% (+1.0pp)
- Zero regressions across full test suite

## Step 4: Fix vMerge border spanning ✅
- Added `compute_merge_spans()` function in `src/pdf/table.rs` that pre-computes extra height from Continue rows for each Restart cell, keyed by `(row_idx, grid_col)`
- Modified `render_table_row` to accept `row_idx` and `merge_spans` parameters
- Border drawing now skips Continue cells entirely (Restart cell's extended borders cover them)
- Restart cells with merge spans draw left/right/bottom borders spanning the full merged height
- Updated all 4 call sites in `render_table` to pass `row_idx` and `&merge_spans`
- Scores: Jaccard 5.5% → 5.9% (+0.4pp), SSIM 19.3% (unchanged)
- Zero regressions across full test suite
