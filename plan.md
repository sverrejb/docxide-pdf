# ✅ COMPLETED: Fix stem_partnerships_guide: behind-doc floating images + vertical alignment

## Context

The `stem_partnerships_guide` scraped fixture renders incorrectly — a large decorative cover image (colored wave design) appears at the **top** of page 1 covering all text, when it should be at the **bottom** and rendered **behind** the text. The DOCX specifies:

```xml
<wp:anchor behindDoc="1" ...>
  <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
  <wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>
  <wp:extent cx="7653020" cy="5392420"/>  <!-- 602.6pt × 424.6pt -->
  <wp:wrapNone/>
</wp:anchor>
```

Two root causes:
1. **`FloatingImage` lacks `behind_doc` field** — image draws on top of text instead of behind it
2. **Vertical `<wp:align>` not parsed** — only `<wp:posOffset>` is handled, so `bottom` defaults to offset 0 (top of page)

## Changes

### Step 1: Model (`src/model.rs`)

**A. Add `VerticalPosition` enum** (next to `HorizontalPosition` at line 145):
```rust
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum VerticalPosition {
    Offset(f32),
    AlignTop,
    AlignCenter,
    AlignBottom,
}
```

**B. Add `behind_doc` and change `v_offset_pt` → `v_position`** on `FloatingImage` (line 157):
```rust
pub struct FloatingImage {
    pub image: EmbeddedImage,
    pub h_position: HorizontalPosition,
    pub h_relative_from: &'static str,
    pub v_position: VerticalPosition,        // was v_offset_pt: f32
    pub v_relative_from: &'static str,
    pub wrap_type: WrapType,
    pub behind_doc: bool,                     // NEW
}
```

### Step 2: Parsing (`src/docx/images.rs`)

**A. Update `parse_anchor_position()`** (line 107) to return `VerticalPosition` instead of `f32`:

Return type: `(HorizontalPosition, &'static str, VerticalPosition, &'static str)`

Parse vertical alignment from `<wp:align>` child (same pattern as horizontal):
```rust
let v_position = if let Some(align_node) = pos_v.and_then(|n| n.children().find(|c| c.tag_name().name() == "align")) {
    match align_node.text().unwrap_or("") {
        "bottom" => VerticalPosition::AlignBottom,
        "center" => VerticalPosition::AlignCenter,
        _ => VerticalPosition::AlignTop,
    }
} else if let Some(offset_node) = pos_v.and_then(|n| n.children().find(|c| c.tag_name().name() == "posOffset")) {
    VerticalPosition::Offset(offset_node.text().unwrap_or("0").parse::<f32>().unwrap_or(0.0) / 12700.0)
} else {
    VerticalPosition::Offset(0.0)
};
```

**B. Update `FloatingImage` construction** (line 257): add `behind_doc` and use `v_position`:
```rust
let behind_doc = container.attribute("behindDoc") == Some("1");
return Some(RunDrawingResult::Floating(FloatingImage {
    image: img, h_position, h_relative_from: h_relative,
    v_position, v_relative_from: v_relative, wrap_type, behind_doc,
}));
```

### Step 3: Update textbox callers (`src/docx/textbox.rs`)

Two call sites at lines 708 and ~line 340 (connectors) destructure `parse_anchor_position()`. They use `v_offset` as `f32`. Extract the offset value:
```rust
let (h_position, h_relative, v_pos, v_relative) = parse_anchor_position(container);
let v_offset = match v_pos { VerticalPosition::Offset(o) => o, _ => 0.0 };
```
(Textboxes don't yet support vertical alignment — this preserves existing behavior.)

### Step 4: Rendering (`src/pdf/mod.rs`)

**A. Extract `resolve_fi_y_top()` helper** to compute vertical position from `VerticalPosition`:
```rust
fn resolve_fi_y_top(fi: &FloatingImage, sp: &SectionProperties, slot_top: f32) -> f32 {
    let img = &fi.image;
    match fi.v_position {
        VerticalPosition::Offset(v_offset) => match fi.v_relative_from {
            "page" => sp.page_height - v_offset,
            "margin" | "topMargin" => sp.page_height - sp.margin_top - v_offset,
            _ => slot_top - v_offset,
        },
        VerticalPosition::AlignTop => match fi.v_relative_from {
            "page" => sp.page_height,
            _ => sp.page_height - sp.margin_top,
        },
        VerticalPosition::AlignCenter => match fi.v_relative_from {
            "page" => (sp.page_height + img.display_height) / 2.0,
            _ => {
                let area = sp.page_height - sp.margin_top - sp.margin_bottom;
                sp.page_height - sp.margin_top - (area - img.display_height) / 2.0
            }
        },
        VerticalPosition::AlignBottom => match fi.v_relative_from {
            "page" => img.display_height,
            _ => sp.margin_bottom + img.display_height,
        },
    }
}
```

**B. Split floating image rendering into two passes** (around lines 1427-1536):

Before behind-doc textboxes (~line 1427):
```rust
// Behind-doc floating images
for (fi_idx, fi) in para.floating_images.iter().enumerate().filter(|(_, f)| f.behind_doc) {
    // ... render using resolve_fi_y_top() ...
}
// Behind-doc textboxes (existing code)
```

At current floating image location (~line 1478):
```rust
// Normal floating images (not behind doc)
for (fi_idx, fi) in para.floating_images.iter().enumerate().filter(|(_, f)| !f.behind_doc) {
    // ... same rendering code ...
}
```

**C. Update space reservation** (line 1179-1191):

Behind-doc images with `wrapNone` already skip reservation (`reserve = false`). Add guard for aligned images:
```rust
if reserve {
    let fi_h = match fi.v_position {
        VerticalPosition::Offset(o) => o + fi.image.display_height,
        _ => fi.image.display_height,
    };
    content_h = content_h.max(fi_h);
}
```

### Step 5: Header/footer rendering (`src/pdf/header_footer.rs`)

Update `fi_y_top` computation at line 353 to use the same `VerticalPosition` match (or call the shared helper).

## Files to modify

| File | Changes |
|------|---------|
| `src/model.rs` | Add `VerticalPosition` enum, update `FloatingImage` struct |
| `src/docx/images.rs` | Parse `<wp:align>` in `parse_anchor_position()`, add `behind_doc` to `FloatingImage` |
| `src/docx/textbox.rs` | Extract offset from `VerticalPosition` at 2 call sites |
| `src/pdf/mod.rs` | Add `resolve_fi_y_top()`, split rendering into behind/in-front passes, update space reservation |
| `src/pdf/header_footer.rs` | Update `fi_y_top` computation for `VerticalPosition` |

## Verification

1. `cargo check` — compilation
2. `cargo test -- --nocapture` — check for `REGRESSION in:` lines (expect zero regressions)
3. Compare `stem_partnerships_guide` page 1 output: decorative image should be at bottom, text visible on top
4. `./tools/target/debug/analyze-fixtures` — verify overall scores
