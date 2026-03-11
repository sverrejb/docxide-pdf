# Progress for plan_archery.md

## Step 1: Fix `line_h_ratio` and `ascender_ratio` in embed.rs — COMPLETED

**What was done:**
- Modified `src/fonts/embed.rs` lines 264-283 to use OS/2 table metrics instead of hhea metrics
- Logic: checks `face.tables().os2` for the OS/2 table
  - If `use_typographic_metrics()` is set (bit 7 of fsSelection, version >= 4): uses `typographic_ascender - typographic_descender + typographic_line_gap`
  - Otherwise (most fonts): uses `windows_ascender - windows_descender` (note: `windows_descender()` already negated)
  - Fallback to hhea metrics if OS/2 table is absent
- `ascender_ratio` uses `windows_ascender` or `typographic_ascender` respectively
- Compilation verified with `cargo check` — no new warnings or errors

## Step 2: Analyze and fix regressions — COMPLETED

**Regression analysis:**

Ran full test suite after Step 1's OS/2 change. Found 8 Jaccard regressions:
- `cases/case7`: 69.3% → 26.8% (-42.5pp) — severe
- `scraped/centrifugal_water_chillers`: 79.2% → 20.2% (-59.0pp)
- `scraped/seminary_hill_board_meeting`: 49.1% → 9.6% (-39.4pp)
- `scraped/czech_expert_witness_law`: 64.8% → 14.0% (-50.8pp)
- `scraped/czech_tree_cutting_permit`: 57.5% → 15.8% (-41.7pp)
- `scraped/bush_fires_act_comparison`: 41.0% → 13.1% (-27.9pp)
- `samples/sample500kB`: 32.6% → 14.6% (-18.0pp)
- `cases/case27`: 96.3% → 93.9% (-2.4pp)

The target fixture (polish_archery_range_plan) also **regressed** from 15.0% → 9.2%.

**Root cause investigation:**

Added debug metrics printing to `embed.rs` to compare old (hhea) vs new (OS/2 win) values for all fonts. Key findings:

1. **Only fonts with non-zero hhea `lineGap` were affected**: Times New Roman (gap=87) and Arial (gap=67)
2. For these fonts, `win_asc == hhea_asc` and `win_desc == hhea_desc` — the only difference is the missing line gap
3. The OS/2 Windows path (`win_asc - win_desc`) drops the hhea `lineGap`, making `line_h_ratio` 3-4% smaller
4. For TNR: OLD lhr=1.1499, NEW lhr=1.1074 (delta -0.0425, ~4.25% shorter lines)
5. For Arial: OLD lhr=1.1499, NEW lhr=1.1172 (delta -0.0327, ~3.27% shorter lines)
6. Fonts with hhea_gap=0 (Calibri, Georgia, Verdana, Courier New) were completely unaffected
7. Aptos (the only USE_TYPO_METRICS=true font): typo metrics = hhea metrics, delta=0

The `0.75` ascender_ratio fallback, `1.2` line_h fallback, and `+0.5` table row padding were NOT involved — they only fire for missing/unresolved fonts. All regressions were caused solely by the smaller `line_h_ratio`.

**Fix applied:**

Modified the non-USE_TYPO OS/2 Windows path in `embed.rs:271-277` to include `face.line_gap()`:
```rust
let win_asc = os2.windows_ascender() as f32;
let win_desc = os2.windows_descender() as f32;
let gap = face.line_gap() as f32;  // external leading
((win_asc - win_desc + gap) / units, win_asc / units)
```

Rationale: `usWinAscent/Descent` define glyph clipping bounds, not line spacing. The hhea `lineGap` provides external leading between lines. Word's text rendering includes this leading. For all fonts tested on this system, `win_asc == hhea_asc` and `win_desc == hhea_desc`, so this produces identical results to the original hhea formula. For fonts where win ≠ hhea (e.g., fonts with extended clipping regions), it correctly uses the OS/2 win extent + hhea gap.

**Verification:** Full test suite passes with zero regressions. All scores restored to baseline.

## Step 3: Update handcrafted fixture expectations — COMPLETED

**What was done:**

Ran full test suite and compared current scores against stored baselines in `tests/baselines.json`.

**Findings:**
- Zero REGRESSION flags across all cases
- All handcrafted cases (case1-case37) pass both Jaccard (≥20%) and SSIM (≥75%) thresholds
- Only noise-level deltas observed: case3 -1.0pp Jaccard, case33 -0.1pp, sample500kB -0.6pp, classroom_weekly -0.2pp
- One tiny auto-improvement recorded: case7 SSIM 91.71% → 91.72% (+0.01pp)
- Baselines.json auto-updates via `merge_max` during test runs (only tracks max-ever scores)

**Conclusion:** No fixture changes needed. The OS/2 fix with hhea lineGap restoration produces identical line heights for all fonts used in handcrafted cases (since `win_asc == hhea_asc` and `win_desc == hhea_desc` for all system fonts). The small deltas are rendering noise, not metric changes.

## Step 4: Un-skip the polish_archery fixture — COMPLETED (fixture remains skipped)

**What was done:**

Ran the polish_archery_range_plan fixture to check current scores after the OS/2 metrics fix:
- Jaccard: **15.0%** (threshold: 20%) — NOT passing
- SSIM: **48.4%** (threshold: 75%) — NOT passing

The OS/2 fix was net-neutral for this fixture. Times New Roman's OS/2 `usWinAscent/Descent` values match the hhea `ascender/descender`, so adding the hhea `lineGap` back produces identical `line_h_ratio` as before (1.1499). The cumulative vertical drift observed in this fixture is NOT caused by font metrics.

**Decision:** Fixture remains on the SKIPLIST since it does not meet the 20% Jaccard threshold. Further investigation into the actual cause of the vertical drift (paragraph spacing, table layout, or other layout differences) would be needed to improve this fixture's score.
