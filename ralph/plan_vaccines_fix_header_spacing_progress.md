# Progress for ralph/plan_vaccines_fix_header_spacing.md

## Step 1: Add width check for wrapSquare textboxes — DONE
- Changed `src/pdf/mod.rs` line ~1555: `WrapType::Square => true` → `WrapType::Square => { tb.width_pt >= text_width * 0.9 }`
- This mirrors the existing floating image width check pattern (lines 1534-1550)
- `cargo check` passes

## Step 2: Visual verification and baseline update — DONE
- `cargo check` passes (8 warnings, no errors)
- `DOCXIDE_CASE=vaccines_history_chapter cargo test`: Jaccard 41.1% (+1.7pp), SSIM 48.6% (+4.6pp) — both improved
- `cargo test -- --nocapture`: zero "REGRESSION in:" lines across all cases
- `analyze-fixtures` confirms vaccines_history_chapter at 41.1% Jaccard, 48.6% SSIM
- Diff image page_001.png confirms "The Beginning" heading moved up significantly (no longer at ~75%)
- `tests/baselines.json` auto-updated by test run with new scores (Jaccard 0.4114, SSIM 0.4861)
