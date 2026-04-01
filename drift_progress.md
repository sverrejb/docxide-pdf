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

### H. Small tolerance / "slop" in line breaking (NOT YET TESTED)

**Theory:** Instead of matching Word's exact width formula, add a small tolerance to the line-break decision. If the last word overshoots by less than ~0.5pt, still fit it on the line.

**Status:** Not tested. This is a pragmatic fix rather than a root-cause fix. Would need careful tuning to avoid regressions on cases that currently break correctly. Could be implemented as: `let overflows = proposed_x + ww > cur_max + TOLERANCE;`

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
2. Word uses slightly narrower widths (~0.003pt/char at 12pt, ~0.05% systematic bias)
3. Word's widths are NOT quantized to twips
4. The effect is per-character, not per-line or per-paragraph
5. Line heights, paragraph spacing, and margins are all correct
6. The issue affects ALL fonts tested (Calibri, TNR, Arial)
7. The effect compounds over long documents, producing extra pages

## What We Don't Know

1. Word's exact advance width computation formula
2. Whether the narrowing is DPI-dependent, font-size-dependent, or both
3. Whether it's a rounding operation on individual glyphs or on accumulated positions
4. Whether the `hdmx` table (device-specific metrics) plays a role (Calibri has no hdmx; TNR has one — untested)
5. Whether Word uses a different text shaping engine that produces different widths

## Recommended Next Steps

1. **Test ppem-based rounding** at 96 DPI (the most promising unexplored theory)
2. **Check TNR's hdmx table** — TNR has one, which could explain device-specific width differences
3. **Test the tolerance approach** — add ~0.5pt line-break slop as a pragmatic fix
4. **Create more diagnostic fixtures** with varying font sizes (8pt, 10pt, 14pt, 18pt) to see if the error scales with font size or ppem
5. **Compare with LibreOffice** — does LibreOffice match Word's widths or ours? This would indicate whether the issue is Word-specific or a known font rendering convention
