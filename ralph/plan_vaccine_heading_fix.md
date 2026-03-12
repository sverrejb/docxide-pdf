# Fix "The Beginning:" heading position in vaccine case

## Context

In the scraped vaccine fixture (`tests/fixtures/scraped/vaccines_history_chapter`), the heading "The Beginning:" renders ~24.5pt too low (further from the top of the page) compared to the Word reference PDF.

**Measured offset**: Reference "The Beginning:" at y=499pt from top (stext coords), generated at y=523.5pt — a 24.5pt gap.

## Revised Root Cause Analysis (2026-03-12)

The original plan blamed the SmartArt paragraph's caption textbox. That was incorrect. Debug tracing of the actual paragraph layout reveals the problem is in **PARA[0]** (the first paragraph), not PARA[2] (the SmartArt paragraph).

### Page 1 paragraph layout (generated)

| Para | Text | content_h | inter_gap | slot_top_after | Key info |
|------|------|-----------|-----------|----------------|----------|
| 0 | *(empty)* | 189.6 | 0.0 | 530.4 | 4 textboxes (3×wrapNone + 1×TopAndBottom), no smartart |
| 1 | "The History of the Vaccine" | 24.8 | 24.0 | 481.5 | Heading1 |
| 2 | *(empty, SmartArt)* | 117.8 | 6.0 | 357.7 | SmartArt diagram + 1 caption textbox (wrapSquare, skipped by width heuristic) |
| 3-4 | *(empty)* | 14.6 | 10.0 | 333.1 / 308.4 | Spacer paragraphs |
| 5 | *(empty)* | 17.6 | 10.0 | 280.8 | Spacer (12pt font) |
| 6 | "The Beginning:" | 16.6 | 10.0 | 254.3 | **Target heading** |

### The bug: PARA[0] TB[3] inflates paragraph height via `+=` path

PARA[0] contains 4 textboxes:
- TB[0-2]: `wrapNone` — correctly skipped (`reserve=false`)
- **TB[3]**: `wrap=TopAndBottom, w=540, h=180, v_off=-41, v_rel=Margin, dist_b=36`

TB[3] is a large decorative textbox (the timeline banner) positioned **relative to the page margin** (`v_relative_from=Margin`), not relative to the paragraph. In the current code (from commit 74f19a6), the `_ =>` branch handles non-Paragraph textboxes:

```rust
match tb.v_relative_from {
    VRelativeFrom::Paragraph => {
        content_h = content_h.max(tb_bottom);
    }
    _ => {
        content_h += tb_bottom;  // BUG: adds 175pt to paragraph height
    }
}
```

The calculation: `tb_bottom = -41.0 + 180.0 + 36.0 = 175.0`, then `content_h = 14.6 + 175.0 = 189.6`.

**This is wrong.** A textbox positioned relative to the page margin is absolutely positioned on the page — it does not affect the paragraph's vertical footprint. Word renders margin-relative textboxes at their absolute position without pushing subsequent paragraphs down.

Before commit 74f19a6, the code used `.max()` for all textboxes, which gave `content_h = max(14.6, 175.0) = 175.0` — still wrong but a different value. The `+=` change in 74f19a6 made it `14.6 + 175.0 = 189.6` (+14.6pt worse).

### Expected behavior

Textboxes with `v_relative_from` != `Paragraph` should NOT contribute to the paragraph's content height at all. They are absolutely positioned relative to the page/margin/column and don't push content down.

Only `VRelativeFrom::Paragraph` textboxes extend the paragraph's vertical space (using `.max()`).

## Fix

### Step 1: Skip non-paragraph-relative textboxes in height calculation

In `src/pdf/mod.rs`, the textbox reservation loop (lines 1582-1601), change the handling so that only paragraph-relative textboxes contribute to content_h:

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
        match tb.v_relative_from {
            VRelativeFrom::Paragraph => {
                let tb_bottom = tb.v_offset_pt + tb.height_pt + tb.dist_bottom;
                content_h = content_h.max(tb_bottom);
            }
            _ => {
                // Margin/Page/Column-relative textboxes are absolutely positioned
                // and do not affect the paragraph's vertical footprint
            }
        }
    }
}
```

This reverts the `_ => { content_h += tb_bottom; }` branch from commit 74f19a6 to a no-op. The pre-74f19a6 code used `.max()` for everything, which was also incorrect (but less wrong). The correct fix is to skip them entirely.

**Expected effect on PARA[0]**: `content_h` stays at `14.6` (the line_h for the empty paragraph) instead of being inflated to `189.6`. This saves `189.6 - 14.6 = 175.0pt`.

### Step 2: Evaluate remaining gap

After Step 1, "The Beginning:" should move up by ~175pt... which is far more than the 24.5pt gap. This means something else is constraining the layout too. The heading position depends on the cumulative effect of all paragraphs above it. Let me trace the expected new positions:

With PARA[0] content_h = 14.6 (instead of 189.6):
- PARA[0]: slot_top = 720.0 - 14.6 = 705.4 (was 530.4)
- PARA[1]: slot_top = 705.4 - 24.0 - 24.8 = 656.6 (was 481.5)
- PARA[2]: slot_top = 656.6 - 6.0 - 117.8 = 532.8 (was 357.7)
- ...continuing would place "The Beginning:" much higher

This would be a **massive** change and likely cause the heading to be too HIGH, since the decorative shapes in PARA[0] DO occupy visual space that Word accounts for differently. The issue may be that PARA[0]'s content_h should reflect the SmartArt/shape area that IS paragraph-relative, just not the margin-relative textbox.

### Step 2 (revised): Investigate what PARA[0] should actually contain

The 3 wrapNone textboxes (TB[0-2]) are at v_offset ~92-95pt relative to the paragraph. These are decorative shapes that ARE positioned relative to the paragraph. In Word, wrapNone shapes don't reserve vertical space — they float over text. But in this case, PARA[0] is empty and its entire purpose is to anchor these shapes + the TopAndBottom banner.

**Alternative hypothesis**: PARA[0] might actually be the SmartArt "anchor" paragraph that Word uses to position the diagram. Its vertical extent should come from the shapes it anchors, but only the paragraph-relative ones. The correct content_h for PARA[0] might be something like `max(TB[0-2] bottoms) ≈ 95 + 106 = 201pt` — but those are wrapNone so they shouldn't count either.

**This needs more investigation.** The relationship between PARA[0] (decorative shapes), PARA[1] (heading), and PARA[2] (SmartArt diagram) in Word's layout model needs to be understood before making changes. A simple fix to just one branch might cascade into other problems.

### Recommended next steps

1. ~~**Measure reference positions more carefully**~~: COMPLETED — see progress file. Key finding: content above "The Beginning:" matches reference within ~1pt; the 24.5pt gap is introduced between the figure caption (y≈449) and "The Beginning:", in the PARA[3-5] spacer region. PARA[0]'s 175pt inflation does NOT produce a 175pt shift — only ~24.5pt leaks through.
2. ~~**Understand PARA[0]'s role**~~: COMPLETED — see progress file. Key finding: PARA[0]'s 189.6pt content_h is approximately correct; it properly positions the heading and SmartArt. The 24.5pt error is NOT from PARA[0] — it's from the spacer paragraphs (PARA[3-5]) which consume ~28pt more space than in the reference
3. **Check if the caption textbox on PARA[2] should contribute**: The original plan's hypothesis about the caption textbox (TB[0] on PARA[2]: `wrapSquare, v_off=127.7, h=144`) might still be partially correct — just not the primary issue
4. **Test the fix incrementally**: First try making non-Paragraph textboxes skip content_h entirely, measure the result, then adjust

## Files to modify

| File | Change |
|------|--------|
| `src/pdf/mod.rs` | Modify textbox reservation loop: skip content_h for non-paragraph-relative textboxes |

## Verification

```bash
# Run vaccine case specifically
DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture

# Check for regressions across all cases
cargo test -- --nocapture 2>&1 | grep "REGRESSION"

# Full score overview
./tools/target/debug/analyze-fixtures

# Measure exact position (should be close to reference y=499pt from top)
mutool draw -F stext tests/output/scraped/vaccines_history_chapter/generated.pdf 1 | grep "Beginning"
```

Expected result: "The Beginning:" moves closer to y=499pt from top (reference position), from current y=523.5pt.
