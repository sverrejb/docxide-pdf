# Fix: Textboxes with wrapTopAndBottom don't push down content

## Context

In the `vaccines_history_chapter` fixture, the first paragraph contains a large gradient rectangle (a `wps:wsp` textbox inside `wp:anchor`) with `wp:wrapTopAndBottom` wrapping. This rectangle contains "Chapter 1" and subtitle text. Three circles with letters (T, Y, B) float over it with `wrapNone`.

The blue underlined heading "The History of the Vaccine" is in a **subsequent paragraph** and should render **below** the gradient rect + circles. Instead, it renders at the top overlapping the gradient because the textbox doesn't reserve any vertical space.

**Root cause**: The `Textbox` struct has no `wrap_type` field. When textboxes are collected from `wp:anchor` elements, the wrap type is discarded. The content height calculation in the renderer only checks `FloatingImage` elements for `TopAndBottom` wrapping — textboxes are completely ignored.

The `parse_wrap_type()` function already exists in `images.rs` and works correctly — it's just never called for textboxes.

## Step 1: Analyze (read-only)

Already completed. Key findings:

- **DOCX structure**: The gradient rect is `wp:anchor` → `wps:wsp` with `wp:wrapTopAndBottom`, `distT="457200" distB="457200"` (36pt each), `positionV relativeFrom="margin"` offset 0, extent 540×180pt
- **Missing field**: `Textbox` struct (`model.rs:198`) has no `wrap_type` or distance fields
- **Missing parse**: `collect_textboxes_from_paragraph` (`textbox.rs:542-556`) doesn't call `parse_wrap_type()`
- **Missing height reservation**: `content_h` calculation (`pdf/mod.rs:964-977`) only loops over `para.floating_images`
- **Existing utility**: `parse_wrap_type()` at `images.rs:164` already parses all 5 wrap types from `wp:anchor`

## Step 2: Implement the fix ✅ COMPLETED

### 2a. Add fields to `Textbox` struct (`src/model.rs:198-213`)

Add three fields:
```rust
pub wrap_type: WrapType,
pub dist_top: f32,    // spacing above shape (points)
pub dist_bottom: f32, // spacing below shape (points)
```

### 2b. Parse wrap type and distances in `collect_textboxes_from_paragraph` (`src/docx/textbox.rs:534-556`)

After `parse_anchor_position(container)`, also call:
- `parse_wrap_type(container)` (import from `images.rs`, already `pub(super)`)
- Parse `distT` / `distB` attributes from the `wp:anchor` element (EMU → points: divide by 12700)

Pass these values into the `Textbox` construction at line 542.

### 2c. Update VML textbox creation (`src/docx/textbox.rs:489`)

Set `wrap_type: WrapType::None, dist_top: 0.0, dist_bottom: 0.0` for VML fallback textboxes (they don't have anchor wrapping info).

### 2d. Reserve height for TopAndBottom textboxes in renderer (`src/pdf/mod.rs`, after line 977)

Add a loop after the existing floating_images loop:
```rust
for tb in &para.textboxes {
    if tb.wrap_type == WrapType::TopAndBottom {
        let tb_y_top = match tb.v_relative_from {
            "page" => sp.page_height - tb.v_offset_pt,
            "margin" | "topMargin" => sp.page_height - sp.margin_top - tb.v_offset_pt,
            _ => slot_top - tb.v_offset_pt,
        };
        let needed = slot_top - (tb_y_top - tb.height_pt - tb.dist_bottom);
        if needed > content_h {
            content_h = needed;
        }
    }
}
```

This computes the exact space from current `slot_top` to the bottom of the textbox + its bottom distance, handling all position modes (page/margin/paragraph-relative).

### Files to modify

1. `src/model.rs` — add `wrap_type`, `dist_top`, `dist_bottom` to `Textbox`
2. `src/docx/textbox.rs` — parse wrap type + distances from anchor, set defaults for VML path
3. `src/pdf/mod.rs` — add textbox height reservation in content_h calculation

### Verification

1. `cargo build` — confirm compilation
2. `DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture` — check that the heading now appears below the gradient rect
3. View generated page 1 image: `tests/output/scraped/vaccines_history_chapter/generated/page_001.png`
4. Compare with reference: `tests/output/scraped/vaccines_history_chapter/reference/page_001.png`
5. `cargo test -- --nocapture` — full test suite, check for "REGRESSION in:" lines to ensure no regressions
