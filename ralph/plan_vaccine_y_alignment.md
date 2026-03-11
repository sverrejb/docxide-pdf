# Fix: "The History of the Vaccine" heading rendered too high

## Context

In the `vaccines_history_chapter` scraped fixture, the heading "The History of the Vaccine" renders **15.6pt too high** compared to the Word reference PDF.

- Generated: 274.8pt from page top
- Reference: 290.4pt from page top

The document structure is: Para 0 (empty text, contains a large `wrapTopAndBottom` title textbox positioned relative to margin) → Para 1 (Heading1 "The History of the Vaccine").

## Root Cause

In `src/pdf/mod.rs` lines 1580-1592, the textbox space reservation uses `max()`:

```rust
content_h = content_h.max(tb_bottom);  // line 1591
```

For Para 0: `content_h = max(line_h, tb_bottom) = max(~17, 175) = 175`

The `max()` **swallows the paragraph's own line height**. When a `wrapTopAndBottom` textbox is positioned relative to the margin (not the paragraph), its displacement zone is independent of the paragraph's own line slot. Word accounts for BOTH — the empty paragraph's line slot AND the textbox pushdown — making them additive.

**Expected**: `content_h = tb_bottom + base_content_h = 175 + ~17 ≈ 192`
**Result**: heading moves from 274.8pt → ~288pt from top (within ~2pt of reference 290.4pt)

## Plan

### Step 1: Fix textbox content_h reservation to be additive for margin/page-relative textboxes

**File**: `src/pdf/mod.rs` ~line 1580-1592

Change the textbox reservation loop: when the textbox is positioned relative to the margin or page (not the paragraph), ADD its pushdown to the paragraph's base content_h instead of taking the max:

```rust
for tb in &para.textboxes {
    let reserve = match tb.wrap_type {
        crate::model::WrapType::TopAndBottom => true,
        crate::model::WrapType::Square => {
            tb.width_pt >= text_width * 0.9
        }
        _ => false,
    };
    if reserve {
        let tb_bottom = tb.v_offset_pt + tb.height_pt + tb.dist_bottom;
        match tb.v_relative_from {
            VRelativeFrom::Paragraph => {
                // Paragraph-relative: textbox and text share the same
                // coordinate origin, so max is correct
                content_h = content_h.max(tb_bottom);
            }
            _ => {
                // Margin/page-relative: textbox is positioned independently
                // of the paragraph's line slot. Both contribute to height.
                content_h += tb_bottom;
            }
        }
    }
}
```

Need to add `use crate::model::VRelativeFrom;` if not already imported at the usage site.

### Step 2: Verify and run tests

1. `DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture` — verify heading moves closer to reference
2. `cargo test -- --nocapture` — check for regressions (look for "REGRESSION in:" lines)
3. `./tools/target/debug/analyze-fixtures` — full score overview

### Step 3: Update baselines

Update `tests/baselines.json` with new scores for `vaccines_history_chapter`.

## Files to modify

- `src/pdf/mod.rs` — textbox content_h reservation loop (~line 1580-1592)
- `tests/baselines.json` — baseline scores

## Regression risk

Low. Only the `vaccines_history_chapter` fixture has a `wrapTopAndBottom` textbox. The change only affects margin/page-relative textboxes, and the current `max()` behavior is preserved for paragraph-relative ones.
