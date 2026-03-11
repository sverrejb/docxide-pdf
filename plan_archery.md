# Plan: Improve polish_archery_range_plan via OS/2 Font Metrics Fix

## Context

The `polish_archery_range_plan` scraped fixture scores **15.0% Jaccard / 48.9% SSIM** (needs 20%/75%). It's a 4-page Polish archery range construction plan using Times New Roman 12pt with a mix of body text, bullet lists, and two tables.

Visual comparison shows **cumulative vertical drift** — our lines are slightly more compact than Word's, so more text fits per page, causing content to land on different pages. The drift starts from paragraph 1 and compounds across all 4 pages.

**Root cause:** `line_h_ratio` in `src/fonts/embed.rs:264-265` uses hhea table metrics (`face.ascender()`, `face.descender()`, `face.line_gap()`) instead of OS/2 table metrics that Word uses. This is a **known issue** documented in `roadmap.md` line 223, which notes that fixing it globally causes 23 regressions because other layout code has been calibrated against the wrong metrics.

**Why fix now:** The OS/2 metrics issue is the single largest source of layout inaccuracy across the entire fixture corpus. Every document with Auto line spacing is affected. Landing this fix (with regression remediation) is the highest-leverage improvement possible.

## The Fix

### Step 1: Fix `line_h_ratio` and `ascender_ratio` in embed.rs

**File:** `src/fonts/embed.rs` (lines 264-266)

Replace:
```rust
let line_gap = face.line_gap() as f32;
let line_h_ratio = (face.ascender() as f32 - face.descender() as f32 + line_gap) / units;
let ascender_ratio = face.ascender() as f32 / units;
```

With logic that:
1. Parse the OS/2 table from `face.tables().os2` via `ttf_parser::os2::Table::parse()`
2. Check `os2.use_typographic_metrics()` (bit 7 of fsSelection; returns false for OS/2 version < 4)
3. If `USE_TYPO_METRICS` is set: use `typographic_ascender() - typographic_descender() + typographic_line_gap()`
4. If NOT set (most fonts): use `windows_ascender() - windows_descender()` (note: `windows_descender()` already returns negated value)
5. For `ascender_ratio`: use `windows_ascender()` or `typographic_ascender()` respectively
6. Fallback to current hhea metrics if OS/2 table is absent

The ttf-parser 0.25 API provides everything needed:
- `ttf_parser::os2::Table::parse(data: &[u8]) -> Option<Table>`
- `os2.use_typographic_metrics() -> bool`
- `os2.windows_ascender() -> i16`, `os2.windows_descender() -> i16` (already negated)
- `os2.typographic_ascender() -> i16`, `os2.typographic_descender() -> i16`, `os2.typographic_line_gap() -> i16`

### Step 2: Analyze and fix regressions

After the metrics fix, run `cargo test -- --nocapture` and `./tools/target/debug/analyze-fixtures` to identify all regressions. The roadmap warns of ~23 regressions. Expected categories:

1. **Genuine improvements** — cases where the wrong metrics accidentally matched Word (unlikely many)
2. **Compensating-code regressions** — layout code that assumed the old smaller line heights:
   - The `0.75` ascender_ratio fallback (used when no font metrics available) — may need adjustment
   - The `1.2` line height fallback in `resolve_line_h` — may need adjustment
   - The `+ 0.5` end-of-cell-mark height in table row calculation (`table.rs:299`)
   - Baseline positioning formula `slot_top - font_size * ascender_ratio` everywhere
   - Any hardcoded spacing/padding values that were tuned to compensate
3. **Page break shifts** — increased line height means earlier page breaks, shifting content between pages. Many handcrafted cases were built against the old metrics.

For each regression, determine the root cause and fix. Some cases may genuinely get worse before other features (widow/orphan control, better page break logic) catch up.

### Step 3: Update handcrafted fixture expectations

Some handcrafted cases (case1-case37) may have been designed with the wrong line height. After fixing regressions, update any stored score baselines.

### Step 4: Un-skip the polish_archery fixture

If the score improves above 20% Jaccard (expected — fixing vertical drift should dramatically improve page alignment), remove it from the SKIP list. The fixture uses these features which are all already implemented:
- Exact line spacing ✅
- Small spacing values ✅
- Multi-section margins ✅
- Paragraph indent overrides ✅
- Table vAlign, trHeight, tblInd ✅
- smartTag handling ✅
- Tab stop clearing ✅

## Files to Modify

| File | Change |
|------|--------|
| `src/fonts/embed.rs` | OS/2 metrics for `line_h_ratio` and `ascender_ratio` (lines 264-266) |
| `src/pdf/mod.rs` | Possibly adjust fallback constants in `resolve_line_h` (line 460) |
| `src/pdf/layout.rs` | Possibly adjust `0.75` ascender_ratio fallback (line 916) |
| `src/pdf/table.rs` | Possibly adjust `0.75` fallback and `+ 0.5` row height padding |
| `roadmap.md` | Update OS/2 metrics section to reflect completion |
| `tests/common/mod.rs` | Un-skip polish_archery if score improves |

## Key API Reference (ttf-parser 0.25)

```rust
// Access raw OS/2 data
let os2_data: Option<&[u8]> = face.tables().os2;

// Parse into structured table
let os2 = ttf_parser::os2::Table::parse(os2_data)?;

// Check flag
os2.use_typographic_metrics() -> bool  // false for version < 4

// Win metrics (most fonts)
os2.windows_ascender() -> i16   // positive
os2.windows_descender() -> i16  // ALREADY NEGATED (negative)
// line_h_ratio = (win_asc - win_desc) / upm

// Typo metrics (when USE_TYPO_METRICS set)
os2.typographic_ascender() -> i16   // positive
os2.typographic_descender() -> i16  // negative
os2.typographic_line_gap() -> i16   // positive
// line_h_ratio = (typo_asc - typo_desc + typo_gap) / upm
```

## Verification

1. `cargo test -- --nocapture` — check for REGRESSION lines
2. `./tools/target/debug/analyze-fixtures` — compare scores before/after
3. `DOCXIDE_CASE=polish_archery_range_plan cargo test` — verify the target fixture improves
4. Visually inspect `tests/output/scraped/polish_archery_range_plan/diff/` pages
5. Check that no non-skipped fixture drops below the 20% Jaccard threshold
6. Verify handcrafted cases (case1-case37) don't regress significantly

## Execution Strategy

Since this is known to cause ~23 regressions, the approach should be:
1. Make the OS/2 fix
2. Run full test suite, capture before/after scores for ALL fixtures
3. Triage each regression — is it a real problem or was the old score artificially good?
4. Fix compensating constants/code for genuine regressions
5. Accept that some cases may temporarily regress until other layout features catch up
