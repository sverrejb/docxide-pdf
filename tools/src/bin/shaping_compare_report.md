# shaping-compare: rustybuzz vs per-char width accuracy

## Question

Does rustybuzz (OpenType shaping) produce text widths closer to Word's than our
current per-character `glyph_hor_advance` approach?

## Method

Parse reference PDF content streams to extract Word's actual glyph widths from
TJ arrays and /W width tables. For each text run, compute the same text's width
via (a) rustybuzz shaping and (b) per-char hmtx advances. Compare both against
the PDF ground truth.

## Results (2026-03-16)

### Handcrafted cases (37 fixtures, 10,681 runs)

All use Aptos via local Word installation — font files match exactly.

- **rustybuzz closer: 1.1%** (120/10,681 runs)
- Per-char is closer or equal in 98.9% of runs
- Per-character deltas: rb=0.84pt/ch, pc=0.87pt/ch (nearly identical)

### All fixtures (62,196 reliable runs after filtering bad ToUnicode decodes)

| Font | Runs | rb wins | |d_rb|/ch | |d_pc|/ch |
|------|------|---------|-----------| ----------|
| Arial | 16,833 | 1.5% | 0.33 | 0.32 |
| Times New Roman | 13,548 | 1.2% | 0.80 | 0.80 |
| Calibri | 10,324 | 17.9% | 1.65 | 1.65 |
| Cambria | 9,783 | 0.9% | 0.08 | 0.07 |
| Aptos | 4,833 | 47.5% | 2.37 | 2.38 |

**Aggregate: rustybuzz closer in 8.4% of runs.**

### Key observations

1. **Aptos is the only font where rustybuzz consistently helps** (~48% win rate).
   For Arial, TNR, and Cambria — the bulk of real-world documents — per-char wins.

2. **The difference between methods is negligible.** Mean per-character delta:
   rb=0.825pt/ch vs pc=0.823pt/ch. The two approaches produce essentially
   identical widths for the vast majority of text.

3. **Font version mismatch is NOT the issue** (verified: system Aptos glyph
   advances match the PDF /W array values exactly for case1). The small deltas
   are genuine shaping differences (GPOS kerning adjustments in TJ arrays).

4. **32% of extracted runs had unreliable text decoding** — CJK fonts, custom
   encodings, and broken ToUnicode CMaps cause the Unicode roundtrip
   (CID → ToUnicode → Unicode → cmap → GID) to produce wrong glyphs. These
   were filtered out (per-char width >20% off from PDF width).

## Conclusion

**rustybuzz would NOT improve our layout accuracy.** The per-character approach
matches Word's output as well as (or better than) rustybuzz for the fonts that
dominate real-world documents. The previous sessions' attempts to integrate
rustybuzz caused regressions due to error cancellation — this data confirms
there was no accuracy benefit to justify the risk.

rustybuzz remains relevant for:
- Complex scripts (Arabic, Indic) where shaping is mandatory
- Ligature rendering (fi, fl) — cosmetic, not layout-critical

But for text width computation (which drives line breaking and layout), the
current per-char approach is sufficient.

## Usage

```
shaping-compare <fixture-path>       # single fixture
shaping-compare --all [--summary]    # all fixtures
shaping-compare --verbose            # show font file matching
```
