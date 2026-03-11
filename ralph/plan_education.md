# Plan: Improve education_consultant_posting Fixture Score

## Context

The `education_consultant_posting` scraped fixture scores **Jaccard 6.68%, SSIM 17.47%** — critically failing (thresholds: 20% / 75%). The reference PDF has 7 pages, but generated output has only 5, with most body text crammed into a wall of text on page 2.

**Root cause**: The document is structured as two massive tables. Table 1 ("Section A") has a single row (row 7) with a full-width cell (gridSpan=7) containing **hundreds of paragraphs** — the entire "Work Assignment" body: Background, bullet lists, Scope of Work, Sub-tasks. This cell should span ~5 pages. Our renderer treats table rows as atomic units and never splits them across pages, so this row overflows off the bottom of one page. Additionally, all cell text renders at a uniform line height, losing paragraph spacing.

## Issues & Fixes (priority order)

### 1. Table row splitting across pages (CRITICAL — highest impact)

**Problem**: `render_table()` in `src/pdf/table.rs:395-428` treats each row as atomic. If `row_h > available_space`, the row overflows off the page bottom.

**Fix**: When a row is taller than available page space, split it at paragraph boundaries and render across multiple pages.

### 2. Per-paragraph rendering in table cells (HIGH)

**Problem**: `render_table_row()` at `src/pdf/table.rs:257-302` renders ALL lines from a cell with a single `render_paragraph_lines()` call using only the FIRST paragraph's `line_h`, `font_size`, and `alignment`. Inter-paragraph spacing is lost.

**Fix**: Replace flat `(Vec<TextLine>, f32, f32)` per cell with a `Vec<CellParagraphLayout>` that preserves paragraph boundaries, and render each paragraph individually.

### 3. SDT-wrapped paragraphs inside table cells (MEDIUM)

**Problem**: `src/docx/tables.rs:226-229` uses `tc.children()` to find paragraphs, missing any wrapped in `w:sdt > w:sdtContent`.

**Fix**: Use `collect_block_nodes(tc)` (already imported at line 13).

## Implementation

### Step 1: Fix SDT parsing in table cells ✅ COMPLETED
**File**: `src/docx/tables.rs`

Line 226 — change `tc.children()` to `collect_block_nodes(tc).into_iter()`. `collect_block_nodes` is already imported from `super::`.

### Step 2: New data structures for per-paragraph cell layout ✅ COMPLETED
**File**: `src/pdf/table.rs`

Replace:
```rust
struct RowLayout {
    height: f32,
    cell_lines: Vec<(Vec<TextLine>, f32, f32)>,
}
```

With:
```rust
struct CellParagraphLayout {
    lines: Vec<TextLine>,
    line_h: f32,
    font_size: f32,
    ascender_ratio: f32,
    alignment: Alignment,
    space_before: f32, // collapsed inter-paragraph gap (0.0 for first)
}

struct CellLayout {
    paragraphs: Vec<CellParagraphLayout>,
    total_height: f32,
}

struct RowLayout {
    height: f32,
    cells: Vec<CellLayout>,
}
```

### Step 3: Update `compute_row_layouts` to produce per-paragraph layout ✅ COMPLETED
**File**: `src/pdf/table.rs:112-207`

Refactor the existing loop. Currently it already iterates paragraphs and computes per-paragraph `line_h`, `font_size`, spacing — but then flattens into `all_lines`. Instead, collect into `Vec<CellParagraphLayout>`. For each paragraph, also compute and store `ascender_ratio` and `alignment` (from `cell.paragraphs[pi]`).

Key: the `Paragraph` model (in `src/model.rs`) has `.alignment`, `.space_before`, `.space_after`, `.line_spacing`. These are already accessed during height calculation but discarded.

### Step 4: Update `render_table_row` for per-paragraph rendering ✅ COMPLETED
**File**: `src/pdf/table.rs:209-359`

Replace the single `render_paragraph_lines` call (lines 289-302) with a loop over `cell_layout.paragraphs`. For each paragraph:
- Advance cursor_y by `para.space_before`
- Compute `baseline_y = cursor_y - para.font_size * para.ascender_ratio`
- Call `render_paragraph_lines` with that paragraph's lines, alignment, line_h
- Advance cursor_y by `para.lines.len() * para.line_h`

Handle vertical alignment (Top/Center/Bottom) by computing total content height from all paragraphs first, then applying the offset.

**Test here** — run `cargo test` to verify zero regressions. This alone should improve spacing.

### Step 5: Table row splitting across pages ✅ COMPLETED
**File**: `src/pdf/table.rs`, in `render_table()` (line 361)

When a row is taller than available space:

1. **Find split point**: Walk `CellLayout.paragraphs` for the tallest cell, accumulating height. Find the last paragraph index that fits entirely within `available_h - cm.top - cm.bottom`.

2. **Render partial row**: New helper `render_partial_row()` — takes paragraph index range (`start_para..end_para`) per cell. Draws:
   - Cell shading for partial height only
   - Text for the paragraph subset
   - Borders: top+left+right on first page; left+right only on middle pages; bottom+left+right on last page

3. **Loop**: After rendering first part, flush page, repeat header rows if applicable, then continue with remaining paragraphs. Repeat if content still overflows.

**Key decisions**:
- Split at **paragraph boundaries only** (not mid-paragraph). Covers the main case; line-level splitting can be added later.
- Use **index ranges** (not Clone) to avoid cloning TextLine/WordChunk. `render_partial_row` takes `&CellLayout` plus `Range<usize>` per cell.
- For single-column cells (like our gridSpan=7 case), the split point equals the tallest cell's split since there's only one cell.

### Step 6: Update `render_header_footer_table` for new RowLayout ✅ COMPLETED
**File**: `src/pdf/table.rs:444-590`

Must update to use new `CellLayout` struct (required for compilation). Mirror Step 4 changes. Header/footer tables don't need row splitting.

### Step 7: Remove from SKIPLIST & test ✅ COMPLETED (kept on SKIPLIST — scores still below thresholds)
**File**: `tests/fixtures/SKIPLIST` — tested but kept `education_consultant_posting` (Jaccard 8.6% < 20%, SSIM 22.5% < 75%)

## Critical files
- `src/pdf/table.rs` — all table rendering (main target)
- `src/pdf/layout.rs` — `TextLine`, `render_paragraph_lines` (consumed, not modified)
- `src/docx/tables.rs` — SDT fix (line 226)
- `src/model.rs` — `Paragraph.alignment`, `Table`, `TableRow`, `TableCell` structs (read only)
- `tests/fixtures/SKIPLIST` — remove fixture from skip

## Verification

1. After Step 4: `cargo test -- --nocapture` — check for zero "REGRESSION in:" lines
2. After Step 5: `DOCXIDE_CASE=education_consultant_posting cargo test -- --nocapture` — check scores
3. After Step 5: Full `cargo test -- --nocapture` — verify no regressions across all fixtures
4. Visual check: compare `tests/output/scraped/education_consultant_posting/generated/` pages against reference

## Expected impact
- Steps 2-4 alone (per-paragraph rendering): ~10-12% Jaccard (spacing fixed, still overflows)
- Step 5 (row splitting): **25-40% Jaccard** (content properly paginated across pages)
- Step 1 (SDT): small additional improvement if any cell content is SDT-wrapped
