# Fix vaccines_history_chapter layout: body text too low, heading spacing, timeline alignment

## Context

Three visual defects on page 1 of the `vaccines_history_chapter` scraped fixture:
1. **"The Beginning" heading at ~75% of page** — should be at ~38% (matching reference)
2. **"The History of the Vaccine" heading spacing** — too close to content above
3. **Timeline shapes misaligned vertically** vs reference

All three are consequences of a single bug: narrow `wrapSquare` textboxes unconditionally inflate paragraph `content_h`, pushing all subsequent body text far too low.

## Root Cause

**`src/pdf/mod.rs:1552-1562`** — The textbox content_h reservation loop:

```rust
for tb in &para.textboxes {
    let reserve = match tb.wrap_type {
        WrapType::TopAndBottom => true,
        WrapType::Square => true,   // <-- BUG: no width check
        _ => false,
    };
    if reserve {
        let tb_bottom = tb.v_offset_pt + tb.height_pt + tb.dist_bottom;
        content_h = content_h.max(tb_bottom);
    }
}
```

Text Box 4 (wrapSquare, **144pt wide** on a **~468pt column** = 31%) gets full space reservation: `127.7 + 144 + 0 = 271.7pt`. This inflates Paragraph 3's content_h from **117.8pt** (inline SmartArt) to **271.7pt** — adding **154pt of spurious vertical space**.

The equivalent floating image code (lines 1534-1550) already correctly checks width:
```rust
WrapType::Square | WrapType::Tight | WrapType::Through => {
    fi.image.display_width >= text_width * 0.9
}
```

## Fix

### Step 1: Add width check for wrapSquare textboxes ✅

**File**: `src/pdf/mod.rs`, lines ~1553-1556

Change:
```rust
crate::model::WrapType::Square => true,
```

To:
```rust
crate::model::WrapType::Square => {
    tb.width_pt >= text_width * 0.9
},
```

`text_width` (line 1334) is in scope. This mirrors the floating image pattern.

### Step 2: Visual verification and baseline update ✅

1. `cargo check` — verify compilation
2. `DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture` — check score improvement
3. `cargo test -- --nocapture` — check for "REGRESSION in:" lines
4. `./tools/target/debug/analyze-fixtures` — full score overview
5. Inspect `tests/output/scraped/vaccines_history_chapter/diff/page_001.png` — confirm "The Beginning" moved up
6. Update `tests/baselines.json` with new scores

## Regression Risk

- **No handcrafted cases** have `wrapSquare` textboxes
- Only `vaccines_history_chapter` among active scraped fixtures is affected
- The change makes textbox handling consistent with the existing floating image pattern
