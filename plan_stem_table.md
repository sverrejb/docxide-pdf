# Plan: Fix Table List Labels, Numbering, and Spacing in stem_partnerships_guide

## Context

The `stem_partnerships_guide` fixture has tables with bullet points and numbering that are not rendered. Comparing reference vs generated output reveals three root causes:

1. **List labels (bullets/numbers) never rendered in table cells** — `render_table_row()` in `table.rs` has NO code to draw `para.list_label`. The label is parsed and stored in `Paragraph` but silently ignored during table rendering.

2. **Style-based numbering not resolved in table cells** — `tables.rs:252` calls `parse_list_info(num_pr, None, None, ...)`. The `ListBullet` style carries `w:numPr` with `numId=16`, but the style's numbering info is never passed. Body text (mod.rs:496-502) correctly passes `para_style.num_id` and `para_style.num_ilvl`.

3. **All cell paragraphs rendered as a single flat block** — `render_table_row()` merges all lines from all paragraphs into one `Vec<TextLine>` and renders with a single alignment, line_h, and baseline. This loses per-paragraph spacing, indentation, alignment, and label positions.

The grey header rows in the second table (numId=28, decimal "1.", "2.", etc.) have **direct** `w:numPr` — so parsing works, but rendering doesn't. The `ListBullet` rows have **style-based** numbering — so both parsing AND rendering fail.

## Implementation

### Step 1: Fix style-based numbering in table cell parsing
**File**: `src/docx/tables.rs:250-252`

Change:
```rust
let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
let (mut indent_left, mut indent_hanging, list_label, list_label_font) =
    parse_list_info(num_pr, None, None, numbering, counters, last_seen_level);
```

To:
```rust
let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
let style_num = para_style.and_then(|s| s.num_id.as_deref());
let style_ilvl = para_style.and_then(|s| s.num_ilvl);
let (mut indent_left, mut indent_hanging, list_label, list_label_font) =
    parse_list_info(num_pr, style_num, style_ilvl, numbering, counters, last_seen_level);
```

This enables `ListBullet` style paragraphs in table cells to resolve their bullet from the style's `numId=16`.

### Step 2: Per-paragraph data structures in table layout
**File**: `src/pdf/table.rs`

Add new structs (before `compute_row_layouts`):
```rust
struct CellParagraphLayout {
    lines: Vec<TextLine>,
    line_h: f32,
    font_size: f32,
    ascender_ratio: f32,
    alignment: Alignment,
    space_gap: f32,           // collapsed spacing gap above this paragraph (0.0 for first)
    indent_left: f32,
    indent_hanging: f32,
    list_label: String,
    list_label_font: Option<String>,
    label_color: Option<[u8; 3]>,
}

struct CellLayout {
    paragraphs: Vec<CellParagraphLayout>,
}
```

Replace `RowLayout.cell_lines: Vec<(Vec<TextLine>, f32, f32)>` with `cells: Vec<CellLayout>`.

### Step 3: Refactor `compute_row_layouts` for per-paragraph layout
**File**: `src/pdf/table.rs:112-207`

The existing loop already iterates paragraphs and computes per-paragraph `line_h`, `font_size`, spacing. Instead of flattening into `all_lines`, collect into `Vec<CellParagraphLayout>`.

For each paragraph, store:
- `lines` from `build_paragraph_lines(runs, fonts, cell_text_w - indent_left - indent_right, indent_hanging, ...)`
  - Note: pass **actual** `indent_hanging` instead of `0.0`, and adjust `cell_text_w` for indents
- `alignment` from `cell.paragraphs[pi].alignment`
- `ascender_ratio` from font metrics
- `space_gap`: For `pi == 0`, include `para.space_before` (Word includes first paragraph's space_before as additional top offset). For `pi > 0`, use `max(prev_space_after, para.space_before)`.
- `indent_left`, `indent_hanging`, `list_label`, `list_label_font`, `label_color` from the paragraph model

Height calculation: `cm.top + cm.bottom + sum(para.space_gap + para.lines.len() * para.line_h) + final_space_after`

### Step 4: Refactor `render_table_row` for per-paragraph rendering
**File**: `src/pdf/table.rs:209-359`

Replace the single `render_paragraph_lines` call (lines 289-302) with a per-paragraph loop:

```rust
// After shading, before borders:
let total_content_h = cell_layout.paragraphs.iter().map(|p| {
    p.space_gap + p.lines.len() as f32 * p.line_h
}).sum::<f32>() + final_space_after;

let start_y = match cell.v_align {
    CellVAlign::Top => row_top - cm.top,
    CellVAlign::Center => {
        let avail = row_h - cm.top - cm.bottom;
        row_top - cm.top - ((avail - total_content_h) / 2.0).max(0.0)
    }
    CellVAlign::Bottom => {
        let avail = row_h - cm.top - cm.bottom;
        row_top - cm.top - (avail - total_content_h).max(0.0)
    }
};

let mut cursor_y = start_y;
for para_layout in &cell_layout.paragraphs {
    cursor_y -= para_layout.space_gap;
    if para_layout.lines.is_empty() || para_layout.lines.iter().all(|l| l.chunks.is_empty()) {
        continue;
    }
    let text_x = cell_x + cm.left + para_layout.indent_left;
    let text_w = (col_w - cm.left - cm.right - para_layout.indent_left).max(0.0);
    let baseline_y = cursor_y - para_layout.font_size * para_layout.ascender_ratio;

    // Draw list label
    if !para_layout.list_label.is_empty() {
        let label_x = cell_x + cm.left + para_layout.indent_left - para_layout.indent_hanging;
        // Use label_for_paragraph (make pub(super)) to get font name + bytes
        // Draw label at (label_x, baseline_y)
    }

    render_paragraph_lines(
        content, &para_layout.lines, &para_layout.alignment,
        text_x, text_w, baseline_y, para_layout.line_h,
        para_layout.lines.len(), 0, &mut Vec::new(),
        0.0, // first_line_hanging already baked into line widths
        ctx.fonts,
    );

    cursor_y -= para_layout.lines.len() as f32 * para_layout.line_h;
}
```

### Step 5: Make `label_for_paragraph` accessible from table.rs
**File**: `src/pdf/mod.rs`

Change `fn label_for_paragraph` and `fn label_for_run` from private to `pub(super)` so `table.rs` can call them.

### Step 6: Update header/footer table rendering
**File**: `src/pdf/table.rs` (the `render_header_footer_table` function)

Must also adapt to the new `CellLayout` struct (same changes as Step 4 but simpler since h/f tables are typically single-paragraph cells).

## Critical Files
- `src/pdf/table.rs` — main target: data structures, layout computation, rendering
- `src/docx/tables.rs:250-252` — style-based numbering fix
- `src/pdf/mod.rs:2374-2405` — make `label_for_paragraph`/`label_for_run` pub(super)
- `src/pdf/layout.rs` — `build_paragraph_lines`, `render_paragraph_lines` (consumed, not modified)
- `src/model.rs` — `Paragraph` fields read: `list_label`, `list_label_font`, `indent_left`, `indent_hanging`, `alignment`

## Verification

1. After all steps: `cargo test -- --nocapture` — check for zero "REGRESSION in:" lines
2. `DOCXIDE_CASE=stem_partnerships_guide cargo test -- --nocapture` — check scores
3. Visual comparison: `tests/output/scraped/stem_partnerships_guide/generated/` pages vs reference
4. Verify: page 4-5 table has bullet points in "What does it look like" column
5. Verify: page 6-7 table has "1.", "2.", etc. in grey header rows and bullets in "How" rows
6. `./tools/target/debug/analyze-fixtures` — full score overview, check no regressions
