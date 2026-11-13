# File size vs Word

Goal: our PDFs are no larger than Word's online export, for any length of
document. Measured 2026-11-10; images are out of scope for now (text first).

## Where we stand

All 1207 fixtures, `generated.pdf` vs `reference.pdf` (per-file table:
`tests/output/pdf_size_vs_reference.tsv`, gitignored; regenerate with the
script in "Measuring" below):

| | Word | Ours | Ours ÷ Word |
|---|---|---|---|
| All 1207 | 102.7 MB | 49.2 MB | 0.48 |
| Median per file | | | 0.53 |
| Larger than Word | | | 15 files |

The gap narrows with length: median ratio 0.54 for 1-page references, 0.77
for the six 100+ page ones. `cases/case13` (205 pages of dense body text, two
fonts) is 1.33× (+328 KB), the only long fixture above 1.

**Why:** our savings are one-off (fonts embedded at a fraction of Word's
size: Cambria 4.8 KB vs 44.5 KB compressed), our losses are per page. case13,
compressed bytes per page:

| | Word | Ours |
|---|---|---|
| Content stream | 3.53 KB | 6.03 KB |
| Other (structure tree, page dicts) | 0.89 KB | 0.34 KB |
| Total | 4.42 KB | 6.37 KB |

About 1.95 KB lost per dense page against roughly 22–40 KB saved per font:
a dense single-font document breaks even at about 15–25 pages, after which
Word's file is smaller. The model predicts case13's +338 KB; measured +328 KB.

## Causes in the content stream

From case13 page 50 (Word 18.7 KB raw / 3.6 KB compressed, ours 48.8 KB /
5.8 KB):

1. **One TJ per word.** The chunk loop in `src/pdf/layout.rs` (`move_to`
   then the `GlyphDraw` match, ~line 6060) writes `x 0 Td`, the word's TJ and
   the trailing space as a separate `Tj` for every chunk: 388 `Td` on the
   page. Word writes one TJ per line with the word gaps as adjustments.
2. **Escaped glyph codes.** Identity-H codes are subset GIDs, so the high
   byte is 0 and low ones are control bytes. pdf-writer's `Str` writes those
   as octal (`\000\021`, 8 chars per glyph), and any byte ≥ 0x80 turns the
   whole string into hex. Worst on case13; little effect on fonts whose used
   GIDs land in the printable range.
3. **Number precision.** `Td` operands carry f32 noise (`36.949005`); Word
   writes 3 decimals.
4. **TJ spacing.** pdf-writer writes `(..) 9 (..)`; Word `(..)9(..)`.
5. **Deflate level.** Content streams use level 6 (`src/pdf/assembly.rs:423`,
   `src/pdf/objstm.rs`).

## Simulated fixes

Real page streams rewritten offline and recompressed (zlib), cumulative,
compressed KB:

| Page | Word | Now | +1 line TJ | +2/3 printable codes, 3 dec | +4 tight TJ | +5 level 9 |
|---|---|---|---|---|---|---|
| learning_cultures_dissertation p30 | 3.32 | 4.16 | 3.53 | 3.42 | 3.31 | 3.20 |
| learning_cultures_dissertation p60 | 3.37 | 3.93 | 3.41 | 3.33 | 3.23 | 3.12 |
| feminist_voice_dissertation p30 | 3.40 | 4.42 | 3.74 | 3.63 | 3.50 | 3.37 |
| feminist_voice_dissertation p60 | 3.94 | 4.88 | 3.98 | 3.88 | 3.74 | 3.56 |
| case13 p30 | 3.57 | 6.28 | 5.24 | 4.51 | 4.25 | 4.14 |
| case13 p60 | 3.57 | 6.37 | 5.32 | 4.55 | 4.27 | 4.14 |

Fix 1 alone gets the dissertations level with Word but leaves case13 ~1 KB
per page over; case13 needs fix 2 as well. With all of them plus our smaller
per-page overhead (0.34 vs 0.89 KB), every tested page ends below Word's. The
fix-1 simulation turns each `Td` into one adjustment number, so byte counts
are estimates (the real adjustment is computed differently, see below).

## Done 2026-11-10 (the five content-stream fixes)

All 1209 fixtures: 48.99 MB → 46.30 MB (0.477 → 0.451 of Word's), 15 → 14
files above Word. case13 1.33× → 0.82×; learning_cultures_dissertation 0.78 →
0.66; feminist_voice_dissertation 0.81 → 0.60. Conversion time, median of
10: case13 0.437 → 0.451 s (+3.1%), learning_cultures +3.1%, feminist_voice
+1.4%, short documents unchanged (deflate 9 is most of it).

Scores against HEAD rescored in the same worktree (the committed baselines
were stale for some 400 fixtures): Jaccard mean 0.8885 → 0.8886, SSIM 0.9537
→ 0.9537, largest drop 0.08pp Jaccard (shift-return-tabbed-noexpand) and 0.05pp SSIM;
largest gain 0.16pp. MuPDF glyph positions: identical text, baselines
identical, x at most 0.0016pt from HEAD in all 1209.

Precision matters more than it looks: a first version with 3-decimal `Td`
and whole-thousandth chunk moves (0.15 MB smaller) kept every glyph within
0.0105pt of HEAD yet moved 232 Jaccard scores, case48 by −2.0pp and
group-fill-inheritance by +6.3pp. MuPDF snaps glyphs to subpixel steps, and
baselines on Word's 0.25pt grid often sit exactly on a step at 150 dpi, so
0.0005pt of y or 0.006pt of x flips a whole line's pixels. Text boundary: one move,
federal_procurement_terms −0.29pp, with identical stext lines apart from
line-box right edges 0.0003pt off (a threshold in the metric, not found).

1. **One TJ per line**: `TjRun` in `src/pdf/layout.rs`. Chunks drawn as
   `Placed` (one face), `Kerned` or `Plain` in an embedded font join the open
   run; the run is flushed before anything else is written (colour, font,
   Tc, Tz, rendering mode, marked-content switch via `LinkTagger::chunk`'s
   `before_switch`, math, shadow, shear, another baseline). Each chunk starts
   where its `Td` put it (a tenth of a thousandth, a hundredth above 20pt); inside it
   the glyphs keep the moves `show_placed`/`show_kerned` made.
2. **Printable glyph codes**: `printable_code` in `src/fonts/embed.rs`, with a
   Flate `CIDToGIDMap`. A font with more than 8281 glyphs keeps Identity.
3. **3-decimal `Td`**: dropped. With one `Td` per line it saved 0.09 MB and
   moved every baseline by up to 0.0005pt (see above).
4. **Tight arrays**: `TjArray` in `src/pdf/layout.rs`, a pdf-writer
   `Primitive` that writes a TJ array without separators; `show_kerned`,
   `show_placed` and `TjRun::flush` write through it.
5. **Deflate 9** for page content, object streams and the xref stream.

**Whole-number `/W`** came with fix 1. MuPDF reads a width as an integer, so
with exact widths a merged line drifted up to 0.4pt at its end in MuPDF while
being right in Poppler. `/W` is now rounded (as Word's is) and the run's pen
uses the same numbers (`FontEntry::pen_width_1000`), so every viewer shows
what MuPDF always showed: the per-glyph moves inside a word are still
computed from the measured widths, so a chunk drifts by its glyphs' rounding
errors. In Poppler against HEAD: 207 files move more than 0.02pt, 11 more
than 0.1pt, the worst 0.5pt at the end of a long tab leader (one chunk of a
hundred dots). Removing the drift (pen-width moves inside chunks too) scored
worse, 54 fixtures down more than 0.5pp Jaccard; hypothesis, not shown: the
glyph grid was fitted to MuPDF output that includes it. Open question for the
grid work; leaders are the first place to look.

Font files were already shared at HEAD (`FontContext::font_files`): the 11
Calibri fonts of `fonts/missing_font_substitution` point at one FontFile2.

Left: `src/pdf/table.rs` (~line 1800) and list labels still write a `Td` and
`Tj` per word; tiny files grew ~200 B per font (the `CIDToGIDMap` object).
Still above Word: missing_font_substitution, case42 (images), 8 `vml-*`,
4 `picture-padding-*`.

## Verification gate

- Work in a worktree.
- Glyph positions unchanged: diff `mutool draw -F stext` before/after on
  case13 and a handful of others, within 0.01 pt.
- Visual and accessibility suites: no score moves; drops need the user's OK.
- Rerun the size table; report total MB and the count of files above Word.

## Measuring

```bash
# per-file table (after a suite run or CLI conversions into tests/output)
python3 - <<'EOF'
import os,glob
for r in sorted(glob.glob('tests/fixtures/*/*/reference.pdf')):
    g=os.path.dirname(r)[len('tests/fixtures/'):]
    o=f'tests/output/{g}/generated.pdf'
    if os.path.exists(o): print(g, os.path.getsize(r), os.path.getsize(o), sep='\t')
EOF
# one page's content stream: object from `mutool show F pages`, then
mutool show -b F <contents-obj>
```
