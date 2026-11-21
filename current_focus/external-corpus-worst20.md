# External corpus — worst 20

Handover for the 20 lowest-Jaccard documents of the neurotic_docx_bench corpus
(`~/workspace/docxside/corpora/neurotic_docx_bench`). The list was made on
6427 pairs; 662 duplicates were deleted afterwards (pixel-identical Word pages,
same folder, same revision and comment counts; list in
`_fetch/purged_duplicates.txt`), leaving 5766. Of the 20, `834c9f9f35_file_57`
went.
Initial analysis only: each cause below comes from page 1–2 side by side,
`pdftotext -bbox` positions and the document XML, not from a fix.

## How the list was made (2026-11-20, `main` at `ffc7603ce`)

```bash
python3 tools/score_external.py            # ~10 min, sandbox off
python3 tools/score_external.py --cached | head -21
```

Results: `tests/output/score_external.json`. Generated PDFs:
`tests/output/corpus/score_external/<group>/<name>/generated.pdf`.

**Every reference in this corpus is a Word for Mac local export** (Producer
"Quartz PDFContext"): All Markup view, Mac font set, 0.24 pt grid. So the
corpus is scored with `--markup`. Scored in the final view, the two tracking
groups lose deleted text and section breaks, and 56 documents could not be
scored at all (page size differs). These are desktop references: best effort,
and never tune a rule for them alone. Check a cause against an online export
(`tools/word_export.py --preset online`) before writing a rule from it.

| group | n | J | SSIM |
|---|---|---|---|
| clean | 2480 | 66.86 | 83.48 |
| tracking_without_comments | 3410 | 64.44 | 81.43 |
| with_comments_clean | 117 | 59.28 | 79.48 |
| with_comments_tracking | 420 | 57.97 | 78.18 |

## The 20

The 20 rows cover 15 distinct documents. `4ade5f9207` and `98af75f057` are the
same text, and three redline pairs appear twice with different redline ids.

| # | document | J | SSIM | pages ours/Word | cause |
|---|---|---|---|---|---|
| 1 | clean/89e56b3180 | 0.00 | 100.00 | 6/4 | H1 booklet reference |
| 2 | with_comments_clean/6ef6726c28 | 0.00 | 10.50 | 1/1 | D comment pane missing |
| 3 | tracking_without_comments/24643f4e34 | 0.33 | 6.71 | 1/4 | H2 reference is another document |
| 4 | clean/d18a18cfba | 2.57 | 7.87 | 26/19 | C vertical text (tbRl) |
| 5 | clean/b50d7a1bf1 (Kant) | 3.54 | 9.70 | 27/27 | B body top under header |
| 6 | clean/c1a1763682 | 4.16 | 27.03 | 2/2 | B body top under header |
| 7 | clean/014f5b166b | 4.27 | 14.17 | 4/3 | F1 CJK grid line pitch |
| 8 | clean/c9d39a9243 | 5.17 | 16.07 | 2/2 | F2 CJK font choice, list label |
| 9 | clean/a5e84c1e55 (CV) | 5.21 | 16.64 | 3/2 | E inline VML text box missing |
| 10 | tracking_without_comments/d0281643d2 …debc9177b1 | 5.41 | 9.97 | 1/1 | A math in revisions |
| 11 | tracking_without_comments/d0281643d2 …84263b7e47 | 5.54 | 10.66 | 1/1 | A math in revisions |
| 12 | with_comments_tracking/970e9bbbcc …2a64722aad | 5.63 | 33.55 | 1/2 | A math in revisions |
| 13 | with_comments_tracking/970e9bbbcc …90fc8e8c86 | 5.70 | 33.39 | 1/2 | A math in revisions |
| 14 | clean/b1eac08c1e (CCPR) | 5.78 | 18.48 | 4/4 | B body top under header |
| 15 | tracking_without_comments/4ba01802a7 …3c8263d2fa | 5.79 | 15.08 | 1/1 | G anchored box over body |
| 16 | tracking_without_comments/98af75f057 | 5.80 | 14.80 | 12/8 | H3 Palatino (Mac) |
| 17 | clean/4ade5f9207 | 5.80 | 14.80 | 12/8 | H3 Palatino (Mac) |
| 18 | tracking_without_comments/4ba01802a7 …362913bbc7 | 5.80 | 16.14 | 1/1 | G anchored box over body |
| 19 | clean/181e161e94_file_57 | 5.93 | 43.47 | 1/1 | H4 Helvetica (Mac) |
| 20 | clean/834c9f9f35_file_57 | 5.93 | 43.47 | 1/1 | H4 Helvetica (Mac) |

## Causes, most promising first

### A. Math inside tracked revisions — DONE (branch `math-revisions`)

Word saves a revised math run with the revision inside it,
`<m:r><w:ins|w:del><w:rPr/><m:t/></w:ins|w:del></m:r>`; we read `m:t` and
`w:rPr` only as direct children, so every revised math run was empty. Fixed:
text and properties are read through the wrapper (and through wrappers
around `m:r`); the final view drops deleted math; the markup view draws
revised math in the author colour, deleted math struck through run by run
across operator spacing, inserted math not underlined (Word). The 107 corpus
documents with revised math: mean Jaccard 36.5 → 54.3 (94 up, 4 down by at
most 2.1); the four worst-20 rows 5.4–5.7 → 33.5–54.9.

Found on the way, also on the branch:
- Display math takes only the left indent; first-line and hanging indents
  don't apply (online probe `tests/output/probes/mathind/mp_display_indent`,
  five indents, every zone within 0.3 pt of Word).
- U+3000 in math (Japanese papers pad equations with it) is an em of space;
  Cambria Math has no glyph and we drew .notdef boxes.

Open:
- Fraction bars, radical signs and overbars of revised math stay black; Word
  colours them (their `m:ctrlPr` carries the revision). Needs a colour on
  structure nodes.
- Display math on a `docGrid type="lines"` page is about 67 pt per line where
  Word takes about 37 (two 18 pt pitches): `9d67577d29` (the −2.1 above).
- Formula speech leaves deleted math out in both views.

### B. Body starts too high under the header (3 rows; 2 real)

Re-exported online (probes in `tests/output/probes/hdrB/`, gitignored):

| document | online | Mac ref | ours | verdict |
|---|---|---|---|---|
| Kant `b50d7a1bf1` | 88.1 | 102.8 | 88.1 | Mac-only (a second empty header line); not code, moved to H |
| CCPR `b1eac08c1e` | 156.5 | 156.5 | 133.0 | real: header tab wrap, below |
| `c1a1763682` | 122.3 | 122.5 | 115.8 → 122.3 | DONE: cell float offset from the text area; wrapping cell floats grow the row |

**CCPR: a tab that overflows right after a right stop on the margin.** Header
paragraphs `logo⇥(V1, 25 May 2019)⇥CCPR 19/03⇥⇥` and `⇥⇥⇥⇥` (Kopfzeile:
centre 4536, right 9072 = the text width; margins 1417; compat 15). Word
breaks *before the last unit* of the right-aligned segment, not before the
overflowing tab: "19/03" moves down with the tabs, and the tabs-only
paragraph takes 3 lines (`⇥ | ⇥ | ⇥⇥`), where we give 2. The second paragraph
is the 11.5 pt; the first takes 2 lines either way.

That carry happens only sometimes: probes with the same shapes in the body,
or in headers with 1440 margins, give Word our 2 lines. A margin sweep
(1400–1445 twips, symmetric, header) flips the line count with a 5-twip
(0.25 pt) period (1432, 1437, 1442 add a line; 1400–1414 take one line fewer),
but the same paragraph as the header's first (sweep s1417: 2 lines) or third
(`hmarg/letter_m1417`: 3 lines) also differs, so the trigger is not settled.
It looks like the device grid: the right stop on the margin compares as one
288-dpi pixel past the header's limit, depending on the margin's grid
phase. Next step: derive that comparison (glyph_grid's rounding?) before
coding the carry rule: break before the last word of the aligned segment, or
before the aligned tab itself when the segment is empty.

**Compat 14 (probe `tabs3/tabwrap14`), DONE:** a tab that overflows right
after a right/centre/decimal stop stays on the line past the margin. Still
open there: case H, a centred segment on the margin runs past it in compat 14
(`left⇥CC 19⇥Z` on one line in Word; we wrap "19").

**Cell floats, open:** a top-and-bottom wrapped picture in a cell also moves
its anchor line below it (probe `cellrow/m1*_tb`: Word's body 13pt lower
than ours).

### C. Vertical text not supported (1 row)

`d18a18cfba`: Japanese, 6 × `w:textDirection w:val="tbRl"`. Word sets the
pages in vertical columns, right to left; we lay out horizontally (26 pages
against 19). Feature work, not a rule: check `roadmap.md` before starting.

### D. Comment pane missing (1 row)

`6ef6726c28`: Word draws the pane with comment 11. We draw no pane. Comment
11's range is empty (`commentRangeStart 11`, `commentRangeEnd 11` adjacent,
then the `commentReference`), and comment 9 has a `commentRangeStart` with no
end and no reference. Likely one of the two drops the pane. Probe each alone.

### E. Inline VML text box missing, extra blank page (1 row)

`a5e84c1e55` (German CV template): page 2 in Word holds a copyright box. It is
a `w:pict` text box (`_x0000_t202`, 492.5 × 120.3 pt,
`mso-position-horizontal-relative:char`,
`mso-position-vertical-relative:line`) after a page break. We draw nothing
there and add a third, blank page. Also: the sidebar column's e-mail address
wraps in ours (`info@…gestalten.d|e`) but not in Word. The table is 5000 pct
with gridCols 3068 / 6 / 7683; compare column widths.

### F. CJK (2 rows)

- F1 `014f5b166b`: Japanese, `docGrid type="lines" linePitch="360"`, 10.5 pt
  Yu Mincho. Word fits more lines per page (ours is 4 pages, Word's 3). Some of
  our lines probably take two grid pitches where Word's take one. Compare
  baselines with `tools/pdf_lines.py`. Memory: `cjk-line-height-factor`,
  `line_height_grid`.
- F2 `c9d39a9243`: Chinese text with no eastAsia `rFonts`; the theme gives
  Jpan MS Mincho/MS Gothic and Hans 宋体. Word drew MS Mincho, Microsoft
  JhengHei and Yu Gothic; we drew SimSun. Check `w:lang w:eastAsia` against
  memory `cjk-font-substitution`. Word also puts each number label ("1.") on a
  line of its own, while ours runs inline with the text.

### G. Anchored box and bar overlap the body (2 rows)

`4ba01802a7` (contract): a `wp:anchor` text box ("Contract ACC", wrapSquare)
and a blue bar sit at the top. In Word the first paragraph starts below the
bar. In ours it overlaps the bar, and every line pitch is smaller than
Word's. Not measured yet. Author colours differ too (purple against red);
that is not code, see memory `tracked-changes-work`.

### H. Not converter bugs

- H1 `89e56b3180`: `w:bookFoldPrinting`. The Mac export imposed a booklet:
  page 1 is 396 × 612 with all text drawn off the page, and pages 2–4 are
  792 × 612 spreads. Jaccard 0 and SSIM 100 is the blank-page artefact. Needs
  an online export, or skip it.
- H2 `24643f4e34`: the reference shows "The Pot-Pourri Compendium"
  (`file_170`). That text is in no part of the docx (4 paragraphs). It is a
  mispaired reference; drop it from the corpus.
- H3 `98af75f057` / `4ade5f9207`: Palatino-Roman from the Mac. Our lines are
  taller (12 pages against 8). Memory `tracked-changes-work` already records
  "Palatino Mac hhea" as not code.
- H5 Kant `b50d7a1bf1`: the Mac export draws a second empty header line
  (body 14.9 pt lower); the online export and ours agree to 0.1 pt.
- H4 `file_57` ×2: Word used Mac Helvetica; we substitute Arial, which is what
  the online export does. Positions match; glyph shapes differ.

## Next steps

1. A follow-ups above (structure colour, grid line height).
2. B (header extent): measure CCPR and `c1a1763682` against a header-height
   model, and check one with an online export first.
3. D, E: one probe each.
4. Drop H1/H2 from the corpus, or re-export them online.
