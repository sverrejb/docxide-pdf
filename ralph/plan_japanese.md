# Plan: Table Cell Vertical Text Direction (`w:textDirection`)

## Context

The `japanese_interlibrary_loan` fixture (3.4% Jaccard, 18.3% SSIM) is a Japanese interlibrary loan form with a complex table. The primary rendering gap is **vertical text direction in table cells** (`w:textDirection val="tbRlV"`), which appears 3 times in the leftmost merged column cells ("申込機関", "申込者", "申込図書"). These labels should be rendered as rotated 90° text flowing top-to-bottom, but currently render as horizontal text crammed into a narrow column.

Secondary issue: **vMerge border rendering** — borders for vertically merged cells don't span all merged rows, causing visual breaks where there should be continuous left/right borders.

Both features are general and affect many CJK documents (forms, contracts, etc.) as well as some Western documents that use rotated table headers.

## Discrepancies Identified (reference vs generated)

1. **Vertical text (HIGH IMPACT)** — Left column labels rendered horizontal instead of rotated 90°
2. **vMerge borders (MEDIUM IMPACT)** — Left/right borders of merged cells broken at row boundaries instead of continuous
3. **Row spacing in "申込図書" section** — Rows for 書名/著者名/出版者/出版年/ISBN more compressed than reference (likely consequence of vertical text cell width being wrong)

## Implementation

### Step 1: Model — Add `TextDirection` enum and cell field ✅ COMPLETED

**File: `src/model.rs`**

```rust
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum TextDirection {
    #[default]
    LrTb,    // Normal horizontal (default)
    TbRl,    // Top-to-bottom, right-to-left (tbRlV / tbRl / rlV / rl)
    BtLr,    // Bottom-to-top, left-to-right (btLr / lr / lrV)
}
```

Add to `TableCell`:
```rust
pub text_direction: TextDirection,
```

### Step 2: Parse `w:textDirection` in table cell properties ✅ COMPLETED

**File: `src/docx/tables.rs`**

After parsing `v_align`, parse `textDirection`:
```rust
let text_direction = tc_pr
    .and_then(|pr| wml(pr, "textDirection"))
    .and_then(|n| n.attribute((WML_NS, "val")))
    .map(|v| match v {
        "tbRlV" | "tbRl" | "rlV" | "rl" => TextDirection::TbRl,
        "btLr" | "lr" | "lrV" => TextDirection::BtLr,
        _ => TextDirection::LrTb,
    })
    .unwrap_or(TextDirection::LrTb);
```

Note: OOXML has both legacy values (`tbRlV`, `btLr`, `lrTbV`) and new values (`tb`, `rl`, `lr`, `tbV`, `rlV`, `lrV`). Map both sets.

### Step 3: Render rotated text in table cells ✅ COMPLETED

**File: `src/pdf/table.rs`**

For cells with `TextDirection::TbRl` (the case in this fixture):
- Text layout: swap effective dimensions — text wraps within the cell's *height* as its line width, and lines stack across the cell's *width*
- PDF rendering: use a rotated text matrix. For `TbRl`, rotate 90° clockwise:
  - `Tm` matrix: `[0 -1 1 0 tx ty]` — this rotates text so characters flow top-to-bottom
  - The text origin shifts: x becomes the vertical position, y becomes horizontal

Concrete changes in `compute_row_layouts()`:
- When building lines for a rotated cell, use the cell's *height* (from `trHeight` or content-based) as `cell_text_w` instead of column width
- Since we don't know height before layout, use a two-pass approach:
  1. First pass: lay out rotated cells assuming unlimited height (single line), compute their natural width need
  2. Use max(content_width_need, row_height) as the effective "line width" for the rotated cell

Actually, simpler: for rotated cells in Word, the text typically fits on one or very few lines (it's usually short labels). Compute the text width of the content, then the cell "height contribution" is the text width, and the cell "width consumption" is the line height. This naturally makes the row tall enough for the rotated text.

In `render_table_row()`:
- For TbRl cells: use PDF content stream `cm` (concat matrix) to rotate the coordinate system 90° clockwise before rendering text, then restore
- Position: the text baseline starts at the top of the cell and goes downward

### Step 4: Fix vMerge border spanning ✅ COMPLETED

**File: `src/pdf/table.rs`**

Current issue: each row draws its own borders independently. For vMerge cells:
- The `Restart` row draws top border + left/right borders for only its own row height
- `Continue` rows skip content but still draw left/right borders for their row height — but because the cell is skipped entirely in `render_table_row`, no borders are drawn for Continue rows

Fix: In `render_table_row()`, for a cell with `VMerge::Restart`:
1. Compute the total merged height by summing heights of all subsequent `Continue` rows in the same grid column
2. Draw left/right borders spanning the full merged height (from restart row top to last continue row bottom)
3. Draw top border at restart row top, bottom border at last continue row bottom
4. Skip left/right/bottom border drawing for `Continue` rows (they're covered by the restart cell's extended borders)

This requires passing the full `row_layouts` slice to `render_table_row` so it can look ahead, or pre-computing merge spans.

**Approach**: Pre-compute a `merge_heights: HashMap<(usize, usize), f32>` mapping `(row_idx, grid_col)` of Restart cells to their total merged height (sum of all merged row heights). Pass this to the render function.

## Critical Files to Modify

1. `src/model.rs` — Add `TextDirection` enum, add field to `TableCell`
2. `src/docx/tables.rs` — Parse `w:textDirection` from `tcPr`
3. `src/pdf/table.rs` — Rotated text layout + rendering, vMerge border fix

## Verification

1. `cargo build` — compiles
2. `cargo test` — no regressions (check for "REGRESSION in:" lines)
3. `DOCXIDE_CASE=japanese_interlibrary_loan cargo test -- --nocapture` — check improved Jaccard/SSIM scores
4. Visual comparison of generated vs reference PNG for this fixture
5. Run `./tools/target/debug/analyze-fixtures` to verify no other fixtures regressed
