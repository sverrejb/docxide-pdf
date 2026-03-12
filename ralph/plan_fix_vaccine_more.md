# Fix vaccine heading position: zero-height section-break paragraphs

## Context

In the vaccine fixture (`vaccines_history_chapter`), "The Beginning:" renders 24.5pt too low. Prior investigation (see `ralph/plan_vaccine_heading_fix_progress.md`) established:

- **PARA[0]'s 189.6pt content_h is correct** — properly positions the heading and SmartArt above
- **The 24.5pt error is in PARA[3-5]** (spacer paragraphs between SmartArt and "The Beginning:")
- Spacer region consumes 76.9pt vs ~49pt in Word's reference — excess ~27.7pt
- **PARA[5] is a section-break-only paragraph** (empty text + `w:sectPr` in its `w:pPr`)

## Root cause

In Word, a paragraph whose only role is to carry `w:sectPr` has **zero rendered height**. It marks a section boundary but doesn't consume vertical space. Our code treats it as a regular empty paragraph, giving it `content_h = line_h` (~17.6pt) plus inter_gap (~10pt) = 27.6pt of unnecessary space.

This 27.6pt closely matches the observed 24.5pt error (the ~3pt difference is font metrics).

## Plan

### Step 1: Add `is_section_break` flag to Paragraph model

**File: `src/model.rs`** (~line 329, Paragraph struct)

Add field: `pub is_section_break: bool`
Add to Default impl: `is_section_break: false`

### Step 2: Set the flag during parsing

**File: `src/docx/mod.rs`** (~line 626)

The section break is already detected at line 629: `if let Some(sect_node) = ppr.and_then(|ppr| wml(ppr, "sectPr"))`. Before this check, the paragraph was already pushed to `blocks` at ~line 626. After the sectPr is detected:

```rust
// After line 626 (blocks.push), before line 628 (sectPr check):
// Mark the last paragraph as a section break carrier
if ppr.and_then(|ppr| wml(ppr, "sectPr")).is_some() {
    if let Some(Block::Paragraph(last_para)) = blocks.last_mut() {
        last_para.is_section_break = true;
    }
}
```

Note: The existing sectPr detection at line 629 can remain as-is since it does the actual section splitting. We just need to mark the paragraph before that happens.

### Step 3: Skip empty section-break paragraphs in the render loop

**File: `src/pdf/mod.rs`** (~line 1393, right after `Block::Paragraph(para) =>`)

Add early skip, similar to the existing `page_break_before` empty-paragraph skip at lines 1411-1414:

```rust
// Skip empty section-break paragraphs — Word gives these zero height
if para.is_section_break
    && is_text_empty(&para.runs)
    && para.image.is_none()
    && para.inline_chart.is_none()
    && para.smartart.is_none()
    && para.floating_images.is_empty()
    && para.textboxes.is_empty()
{
    // Don't update prev_space_after — the previous paragraph's
    // space_after propagates through to the next section
    global_block_idx += 1;
    continue;
}
```

Key details:
- Only skip if truly empty (no text, no images, no shapes, no charts)
- **Don't update `prev_space_after`** — preserves the spacing from the last visible paragraph
- Section-break paragraphs with visible content are NOT skipped (they render normally, just the mark is invisible)

### Expected impact

- PARA[5]'s total contribution (inter_gap 10.0 + content_h 17.6 = 27.6pt) is removed
- "The Beginning:" moves from 523.5 → ~496pt from top (reference: 499pt)
- Residual error: ~3pt — within visual comparison tolerance

### Regression risk

- **case25, case26, case28**: Have multiple empty section-break paragraphs (confirmed in case25). These currently consume unnecessary space that would be removed. Scores may change (likely improve or be neutral).
- **Other multi-section fixtures** (learning_cultures, transition_to_work, stem_partnerships, etc.): May be affected if they have empty section-break paragraphs.
- **Single-section fixtures**: Unaffected (only body-level sectPr, which is not on a paragraph).

## Files to modify

| File | Change |
|------|--------|
| `src/model.rs` | Add `is_section_break: bool` to `Paragraph` |
| `src/docx/mod.rs` | Set flag when `w:sectPr` found in paragraph's `w:pPr` |
| `src/pdf/mod.rs` | Skip empty section-break paragraphs in render loop |

## Verification

```bash
# 1. Run vaccine case — "The Beginning:" should move to ~496-499pt from top
DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture

# 2. Check for regressions across ALL cases
cargo test -- --nocapture 2>&1 | grep "REGRESSION"

# 3. Full score overview
./tools/target/debug/analyze-fixtures

# 4. Verify exact position improvement
mutool draw -F stext tests/output/scraped/vaccines_history_chapter/generated.pdf 1 | grep "Beginning"
```

Expected: zero regressions, vaccine Jaccard/SSIM scores improve.

## After implementation

- Update `ralph/plan_vaccine_heading_fix_progress.md` with results
- Copy this plan to `ralph/plan_fix_vaccine_more.md` as requested by user
