# Vertical Drift Investigation Progress

## Summary

What appears as "vertical drift" — text gradually shifting down relative to Word's output across multi-page documents — is actually caused by **horizontal text width precision**: our glyph advance widths are systematically ~0.003pt/char wider than Word's, causing slightly earlier line breaks, more total lines, and eventually extra pages.

## Affected Fixtures

| Fixture | Pages (gen/ref) | Extra lines | Primary font |
|---------|----------------|-------------|-------------|
| russian_volunteerism_essay | 49/48 | +296 total | TNR 14pt |
| english_town_council_report | 9/8 | ~+40 | Calibri 12pt |
| zimbabwe_gold_mining_judgment | 8/8 | visible drift | TNR 12pt |
| czech_works_contract_proposal | 6/6 | visible drift | Calibri/TNR |
| czech_health_statement_form | 2/2 | visible drift | Calibri 12pt |

## The Evidence Chain

### 1. Line heights are correct

Measured via debug instrumentation in the render loop (`mod.rs`):

- TNR 14pt, Auto(1.0): `line_h_ratio=1.1499`, `line_h=16.10pt`
- Verified font file: MS Times New Roman at `/System/Library/Fonts/Supplemental/Times New Roman.ttf` has `winAsc=1825, winDesc=443, hheaGap=87, UPM=2048`
- Formula: `(1825 - (-443) + 87) / 2048 = 1.1499` → `14.0 * 1.1499 = 16.10pt` ✓
- Note: ttf-parser's `windows_descender()` returns `-443` (negated from the unsigned OS/2 value)

### 2. Per-page bottom positions match

Last text y-position on page 1:
- Generated: 747.33pt
- Reference: 747.49pt
- Delta: -0.16pt (negligible)

Same pattern on pages 2 and 10 — content fills to the same depth per page.

### 3. But we produce MORE lines total

Russian essay page-by-page line count comparison (mutool stext extraction):
- Generated: 2926 total lines across 50 pages
- Reference: 2630 total lines across 50 pages
- **296 extra lines** — every single page has more lines in our output

### 4. Character positions diverge progressively

Measured on case4 (Calibri 12pt, left-aligned body text):

```
Char  Gen x      Ref x      Delta
  1   72.000     72.025     -0.025
  5   92.836     92.761     +0.075
 10  118.720    118.501     +0.219
 20  187.013    186.778     +0.235
 50  ...        ...         ~+0.35
 85  517.302    516.830     +0.472
 89  506.374    506.101     +0.273
```

Average drift: **+0.003pt per character** at Calibri 12pt.

### 5. Our widths match the font file exactly

Verified with Python fontTools for Calibri at `/Applications/Microsoft Word.app/Contents/Resources/DFonts/Calibri.ttf`:

```
Char  Font advance  Computed x    Gen x      Ref x
'T'   998 units     72.000        72.000     72.025
'h'   1076          77.848        77.844     77.773
'e'   1019          84.152        84.144     84.025
' '   463           90.123        90.120     90.001
```

Computed (fontTools) ≈ Generated (our code) — differences < 0.01pt.
Both are systematically WIDER than the reference (Word).

### 6. Word's advance widths are narrower than the font file

Implied advance widths from consecutive character positions in the reference PDF:

```
Char  Our advance  Word's advance  Delta
'T'   5.844pt      5.748pt         +0.096pt
'h'   6.300pt      6.252pt         +0.048pt
'e'   5.976pt      5.976pt          0.000pt
' '   2.716pt      2.760pt         -0.044pt
'q'   6.300pt      6.252pt         +0.048pt
```

Word's advances are NOT quantized to twips (not multiples of 0.05pt).

### 7. The practical effect

At Calibri 12pt with 468pt text width: we fit 89 chars per line, Word fits 90. One character fewer per line, because the accumulated +0.27pt over 89 chars pushes us past the margin for the 90th character.

Over 48 pages × ~40 lines/page = ~1920 lines, ~40 lines are affected → net +1 page.

## Hypotheses Tested and Rejected

### A. hhea lineGap in line_h_ratio (REJECTED)

**Theory:** Word excludes hhea lineGap from line spacing, only using `(winAsc + winDescent) / UPM`.

**Test:** Removed `+ gap` from `compute_line_metrics` in `embed.rs:289`.

**Result:** 80+ regressions, 15 improvements. The lineGap is needed — many fonts (TNR hheaGap=87, Arial hheaGap=67) depend on it. Without it, line heights are too short, causing content to overflow pages differently.

### B. ceil() rounding of Auto line heights (REJECTED)

**Theory:** Word rounds auto line heights up to the nearest whole point: `ceil(font_size * ratio * mult)`.

**Test:** Added `.ceil()` to the Auto branch of `resolve_line_h` in `helpers.rs`.

**Result:** Massive regressions across the board (90%+ cases). The rounding is too aggressive for fonts with fractional line heights. Many fonts produce line heights like 14.35pt that should NOT round to 15pt.

### C. Margin calculation error (REJECTED)

**Theory:** Our margin calculation is wrong, producing a narrower text area than Word's.

**Test:** Measured left/right text edges in multiple cases:
- case4: Gen left=90.00, Ref left=90.05 (match); Gen right=522.00, Ref right=524.75 (+2.75pt)
- Russian: Gen right=552.81, Ref right=559.39 (+6.58pt)
- Zimbabwe: Gen right=540.01, Ref right=545.80 (+5.79pt)

**Result:** Our margins match the DOCX spec exactly. The "wider" text in the reference is from Word fitting more characters per line (not a margin issue). The right-edge overshoot is the last character's advance width extending past the margin because Word fits one more word.

### D. Twip-resolution rounding of widths (INVESTIGATED, INCONCLUSIVE)

**Theory:** Word computes widths in twips internally (1/20 pt).

**Analysis:** `round(advance / UPM * fontSize * 20) / 20` — tested mathematically. For Calibri 'T' at 12pt: raw=5.8477pt, rounded to twips=5.85pt. The rounding actually makes widths WIDER (by 0.002pt), opposite to what we need.

### E. EMU-resolution rounding (NOT YET TESTED)

**Theory:** Office uses EMUs (1/914400 inch = 1/12700 pt) internally.

**Status:** Not tested. Would require: `round(advance / UPM * fontSize * 12700) / 12700`. The very fine resolution makes it unlikely to produce the ~0.003pt/char effect we see.

### F. Device-pixel rounding at reference DPI (NOT YET TESTED)

**Theory:** Word computes advance widths at a reference ppem (e.g., 96 DPI) with pixel-level rounding: `round(advance * ppem / UPM) / ppem * fontSize` where ppem = fontSize * DPI / 72.

**Status:** Not tested. This is the most promising unexplored theory because:
- At 96 DPI, Calibri 12pt has ppem = 16
- Rounding to 1/16th of an em would produce systematic width changes
- The effect would be font-size-dependent, matching our observations

### G. Integer font-unit accumulation (NOT YET TESTED)

**Theory:** Word accumulates glyph positions in integer font units (not floating-point points), converting to points only for final output.

**Status:** Not tested. This would produce a different error pattern — positions would be quantized to `1/UPM * fontSize` steps rather than smooth accumulation.

### H. Small tolerance / "slop" in line breaking (TESTED — FRAGILE)

**Theory:** Instead of matching Word's exact width formula, add a small tolerance to the line-break decision. If the last word overshoots by less than ~0.5pt, still fit it on the line.

**Tested (April 2026):** Exhaustive sweep of tolerance values from 0.07pt to 0.75pt.

Results:
- **0.07pt**: Zero effect (same as baseline 0.05pt for all test cases)
- **0.10pt**: 2 regressions (case11 -5.7pp TxtBnd, education_consul -9.4pp), 1 improvement
- **0.15pt**: Same regressions as 0.10pt — the flip threshold is between 0.07pt and 0.10pt
- **0.35pt**: 20 regressions, 3 improvements (Lithuanian +45pp, but too many losses)
- **Adaptive per-char tolerance** (0.002–0.005pt × char_count): Same regression pattern as fixed tolerance because the per-char scaling doesn't help — the issue is font-specific, not line-length-specific

**Conclusion:** Tolerance approach is fundamentally fragile. It can't distinguish between "barely overflows due to width bias" (should fit) and "genuinely doesn't fit but our metrics accidentally agree" (should not fit). Case11 (Cambria body text, near-zero width bias) regresses at ANY tolerance above 0.07pt because its correct line break happens at ~0.05pt overflow, and any additional tolerance overrides it.

### I. Global layout width correction factor (TESTED — PROMISING BUT FONT-DEPENDENT)

**Theory:** Multiply all layout-time advance widths by a small factor (e.g., 0.9998 = 0.02% narrower) to compensate for Word's systematic narrowing. PDF-embedded widths stay exact.

**Tested (April 2026):** Sweep from 0.9990 to 0.9999, both global and per-font.

**Measured per-font signed width bias** (from reference PDF character positions via mutool stext):

| Font | Size | Mean signed diff | Implied factor | Cum error/80 chars | Sample size |
|------|------|-----------------|---------------|-------------------|-------------|
| Calibri | 12pt (body text, case4) | +0.007pt/char | 0.9987 | 0.55pt | 5052 chars |
| Calibri | 12pt (repeated, case54) | +0.016pt/char | 0.9974 | 1.31pt | 1200 chars |
| Calibri | 11pt | +0.008pt/char | 0.9984 | 0.64pt | 2160 chars |
| Arial | 12pt | +0.003pt/char | 0.9995 | 0.20pt | 1360 chars |

Note: Calibri-Bold 13/14pt showed NEGATIVE bias (our widths narrower) but this may be due to font-path mismatch in the analysis script (loading regular Calibri.ttf for bold text).

**Global factor results:**

| Factor | Regressions | Improvements | Notes |
|--------|------------|-------------|-------|
| 0.999 | 46 | 8 | Way too aggressive — Lithuanian +45pp but massive regressions |
| 0.9997 | 8 | 7 | Near-balanced but too many regressions |
| 0.9998 | 2 | 4 | Better ratio but case11 (Cambria) and seminary_hill regress |
| 0.9999 | 1 (case11) | 2 | Even minimal correction breaks case11 |

**Per-font factor results** (Calibri/TNR get correction, Cambria/Aptos/others get 1.0):

| Calibri/TNR factor | Regressions | Improvements | Notes |
|---------------------|------------|-------------|-------|
| 0.9997 | 4 | 4 | feminist_voice, mandated_reporter, seminary_hill regress |
| 0.9998 | 1 (seminary_hill) | 4 | case49 +3.4pp, case54 +2.6pp, polish +2.8pp |
| 0.99985 | 0 | 3 | **CLEAN — zero regressions**, case49 +2.7pp, case54 +2.2pp, polish +2.8pp |
| 0.99982 | 0 | 3 | **CLEAN**, case49 +3.0pp, case54 +2.4pp, polish +2.8pp |

**Conclusion:** Per-font factors work but are limited:
1. The optimal factor varies by font (Cambria needs 1.0, Calibri needs ~0.9998)
2. It also varies by font size (TJ adjustments change sign between sizes — see kerning_and_shaping.md)
3. A single per-font factor is too simplistic: it helps for the average case but can't track size-dependent hinting
4. The clean-improvement range (0.99982–0.99985 for Calibri/TNR) captures only ~30% of the theoretical correction (need ~0.9987 but can only safely apply ~0.99985)
5. A per-font × per-size lookup table is needed for full correction — confirms the data-driven approach in the roadmap

## Diagnostic Fixtures Created

### case54

Test fixture with repeated short words at various fonts/sizes:
- Block 1: "mm " × 200 — Calibri 12pt
- Block 2: "test " × 200 — Calibri 12pt
- Block 3: "test " × 200 — Times New Roman 12pt
- Block 4: "test " × 200 — Times New Roman 14pt
- Block 5: "test " × 200 — Arial 12pt
- Block 6: "The quick brown fox..." × 30 — Calibri 11pt

Result: 69.5% Jaccard, 87.1% SSIM. The "mm" and "test" blocks match Word's line breaks exactly (21 words/line for Calibri 12pt "mm"). The natural-text block shows the subtle width difference.

## What We Know For Sure

1. Our glyph widths match the raw font file metrics (verified via fontTools)
2. Word uses slightly narrower widths (~0.003–0.016pt/char at 12pt depending on font and text)
3. Word's widths are NOT quantized to twips
4. The effect is per-character, not per-line or per-paragraph
5. Line heights, paragraph spacing, and margins are all correct
6. The issue affects ALL fonts tested (Calibri, TNR, Arial) but magnitude varies greatly:
   - Calibri 12pt: +0.007pt/char signed bias (0.55pt over 80 chars)
   - Arial 12pt: +0.003pt/char signed bias (0.20pt over 80 chars)
   - Cambria 12pt: ~0pt signed bias (near-zero, from shaping_compare: 0.07 per-mille absolute)
7. The effect compounds over long documents, producing extra pages
8. A per-font correction factor DOES help (Calibri/TNR 0.99985 gives 3 improvements, 0 regressions) but captures only ~30% of the needed correction — size-dependent effects prevent going further
9. Tolerance-based approaches are fundamentally fragile — can't distinguish genuine overflows from bias-caused overflows
10. The correction is size-dependent: TJ adjustments change sign between font sizes (e.g., Aptos H→e is +7 at 10pt but -6 at 14pt) — a per-font × per-size table is necessary for full correction

## What We Don't Know

1. Word's exact advance width computation formula
2. The precise per-glyph, per-ppem correction values for each common font
3. Whether a formula (ppem rounding variant) can approximate the corrections, or if a lookup table from Word's actual output is required
4. Whether the `hdmx` table (device-specific metrics) plays a role (Calibri has no hdmx; TNR has one — untested)

## Recommended Next Steps

The data-driven approach outlined in the roadmap (Phases 1-4) is confirmed as the right path forward. Specific priorities:

1. **Phase 1: Data collection** — Generate single-glyph and bigram width sheets for Calibri, TNR, Arial, Aptos, Cambria at sizes 8-24pt. Convert via Word online. Extract TJ adjustments as ground truth.
2. **Phase 2: Formula testing** — Test ppem rounding variants against the ground truth. Also check TNR's hdmx table. If a formula fits → implement directly.
3. **Phase 3: Lookup table** — If no formula, build `{font, ppem, glyph_id} → correction` table from the collected data. ~20K entries, ~80KB.
4. **Interim: consider shipping the per-font factor** — Calibri/TNR at 0.99985, Arial at 0.9999, others at 1.0 gives 3 improvements with zero regressions. Small but safe win while the data pipeline is built.
5. **Analysis script preserved** — `tools/experiments/width_analysis.py` extracts per-character signed width errors from reference PDFs via mutool stext. Useful for validating future corrections.
