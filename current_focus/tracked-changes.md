# Tracked changes (revision markup)

A document with tracked changes keeps each edit as a revision: `w:ins` /
`w:del` runs, `w:moveFrom` / `w:moveTo`, revised paragraph marks and table
rows, and property changes (`w:rPrChange`, `w:pPrChange`, …). The **final**
view (insertions plain, deletions dropped) is Word's "No Markup" view and what
Word's online export prints. The **markup** view is Word's "All Markup" with
revisions shown inline.

**Decided (2026-10-16):**
- **Final stays the default.** Markup is opt-in: `ConvertOptions::revisions`
  (`RevisionView::Markup`) in the library, `--markup` in the CLI.
- Fixtures with revisions get a **separate markup reference** next to the
  final one where it applies (see Testing).

## Evidence

- Word's online export prints without markup (`online-references.md` §4b);
  `scraped/bush_fires_act_comparison` (online reference, tracked changes, no
  markup) agrees. That is the final view.
- Word for Mac's local export of 300 tracked documents from an external
  corpus sample (local only) shows **All Markup inline**: inserted text
  underlined, deleted text struck through, both in the author's colour,
  deleted content kept in the layout (it wraps and paginates), a black change
  bar in the left margin beside changed lines. No revision balloons in any of
  the 300, with or without comments. `scraped/door_air_cooling_unit_spec`
  (older local reference) is the exception: deletions in balloons.
- **Inline vs balloons is a Word preference, not in the docx.** This machine's
  Word defaults to balloons ("Show revisions in balloons"), so the first probe
  export printed deletions in the margin pane. `tools/word_export.py
  --revisions inline` sets the view (`revisions mode of view of active
  window`) before saving; the markup references use it, matching the corpus.
- `w:revisionView` (settings) appears in none of them; `w:trackRevisions`
  (26 docs) only says whether new edits are tracked and changes nothing drawn.

Element census over the 300 documents:

| element | docs | elements |
|---|---|---|
| `w:ins` / `w:del` | 291 / 278 | 27620 / 30403 |
| `w:delText` / `w:delInstrText` | 275 / 82 | 25178 / 1395 |
| inserted / deleted paragraph mark (`w:pPr/w:rPr/w:ins\|del`) | 244 / 219 | |
| row `w:ins` / `w:del` (`w:trPr`) | 151 | |
| deleted picture (`w:drawing`/`w:pict` in `w:del`) | 79 | |
| `w:rPrChange` / `w:pPrChange` / `w:sectPrChange` | 144 / 160 / 176 | 746 / 965 / 176 |
| `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange` / `w:tblGridChange` | 25 / 25 / 27 / 28 | |
| `w:moveFrom` / `w:moveTo` | 9 / 14 | 169 / 179 |
| `w:cellDel` | 3 | |
| `w:cellIns`, `w:cellMerge`, `w:numberingChange` | 0 | |

Authors per document: 1 in 275, 2+ in 19 (6 have none on ins/del: property
changes only).

## Rules, as measured (Word for Mac 16, local print export, 2026-11-05)

Probe kit: `tests/output/tc/` (local; `gen.py` builds p01–p16 with
python-docx plus raw revision elements, `tcprobe.py` dumps text, colours and
rules from a PDF, `authors.py` checks author order against corpus
references). Exports: `inline/` (print preset, inline revisions), `print/`
(balloons, the machine default), `online/` (final view).

**Inline marks**
- Inserted text: the run's text in the author's colour, single underline in
  that colour at the run's normal underline position and thickness
  (0.72pt at 11pt, 0.96pt at 14pt; the underline scales with the run size).
- Deleted text: struck through, same colour and thickness, at the normal
  strike height. The text keeps its place: lines wrap and pages break with
  it (p01: byte-identical reflow between Word and us).
- The author colour **replaces** the run's own colour; the revision line
  **replaces** the run's own underline style (a double-underlined insertion
  prints one single underline, a struck deletion one strike). Highlight
  stays. Moves (`w:moveTo` / `w:moveFrom`) could not be probed: Word hangs on
  the hand-written move ranges (p05, kept in `skip/`).
- Property changes (`w:rPrChange`, `w:pPrChange`): no inline mark, only the
  change bar. Paragraph marks: an inserted or deleted mark draws no glyph,
  only the bar; a paragraph with a deleted mark stays a separate paragraph
  in the markup view and **joins the next paragraph in the final view**, the
  merged paragraph keeping the next one's properties (p04 online export:
  "Para C, deleted mark after it.Para D."). A fully deleted paragraph stays
  in the markup layout, struck through; the final view drops it.
- Section breaks in a deleted paragraph mark (p17): the **final view drops
  the break** (online export: three such sections print as one page with
  the last section's header, the text before the mark joined to the next
  paragraph); the **markup view keeps it** (inline export: three pages, each
  with its own header). Corpus documents mix Letter and A4 sections this
  way, so their final-view scores against markup references are meaningless.
- `<w:titlePg w:val="0"/>` (left behind with a `sectPrChange`) means no
  first-page header; we used to read presence as true and printed a header
  Word does not.
- Tables: inserted and deleted rows stay in the markup view with their runs
  marked; the final view drops deleted rows (p06 online export). Cell
  borders keep their colour.
- Pictures: a deleted inline picture is drawn, with a strike at the text's
  strike height across its width plus a 1.44pt line through its middle; an
  inserted one gets the underline; the bar covers the picture's height
  (p10: 603.6–720). Implemented 2026-11-05 (strike 678.80 vs 678.96).
- Deleted fields (`w:delInstrText`): the cached result struck through; a
  deleted TOC with an empty cached result prints nothing (a corpus document
  showed the heading and its tip, no entries; we regenerated the entries).
- Comments on deleted content (probes p21/p21b, 2026-11-05): a comment on
  deleted text in a paragraph or a deleted cell keeps its balloon, also with
  a dateless "Comparison" deletion; one whose range opens before a wholly
  deleted table and closes after it (between blocks), with its reference in
  a deleted paragraph, gets no balloon in the markup view either, though it
  still takes a number. Four table-heavy comments-set documents print no
  pane for that reason (6.7 → 76.4 Jaccard once modelled: deleted text
  counts as text in `comment_ranges`).
- Lists: two counters (p11b): deleted items number by the original stream,
  live items by the final one, both shown ("4.2.") when they differ. The
  label is coloured and underlined / struck like the text, the tab gap after
  a deleted label struck 0.48pt thin, a renumbered label's original number
  struck and final underlined in 0078D4 (ours identical since 2026-11-05).
- Headers and footers: marked the same way, bars at the same x (p09:
  742.56–756 beside the header's inserted line; ours identical since
  2026-11-05).
- Comments plus revisions: the page is zoomed for the comment pane as for
  comments alone; the inline marks and bars zoom with the body.

**Author colours**: one colour per author in the order the authors first
appear, from a 20-colour cycle: D13438, 0078D4, 5C2E91, 498205, CC3595,
7160E8, 038387, 6D5700, CF0F1F, 4E6AED, B146C2, 394146, 0B6A0B, CA5010,
750B1C, 5D5A58, 881798, 69797E, 005B70, 8E562E (p02 + p02b, 34 fresh
names). Word remembers an author's colour per machine, so an author seen
before keeps an earlier colour (corpus docs: the same author name has the
same colour across documents, unrelated to its order there); no converter can
reproduce that, and ink is ink for the scores.

**Inserted pictures** (p22–p22e, 2026-11-05 night): a lone picture in its own
paragraph makes a line of exactly its height (bottom on the baseline, nothing
below; compat 14 and 15, Times New Roman and Arial, single / 1.5 / triple
spacing, as already modelled). The same picture *inserted* keeps its run's
descent below it, like the underlined text it is (+2.6pt at 12pt Times New
Roman at every spacing); a short inserted picture keeps its line but sits on
the text baseline instead of the line's foot. A deleted picture adds nothing.
Implemented in `picture_line_bottom` / `inserted_picture_descent`
(`pdf/layout.rs`) and the `para.image` path of `pdf/mod.rs`, which also draws
the revision underline under a lone inserted picture (Word: 0.48pt, 1.44pt
below the baseline at 12pt).

**General layout rules found on the corpus** (not markup-specific):
- A list label that exactly fills its hanging indent keeps the text on the
  indent (p23: "(1)" in SimSun 12 under a 360/360 indent → text at 360;
  "(10)" overshoots by 6pt → the next default stop). "Fills" is judged on
  Word's 288 dpi device grid: "(iii)" in Courier New 12 is 36.25pt there, not
  36.006, and tabs past a 36pt indent (zimbabwe_broadcasting p5). An explicit
  stop the label ends on is skipped like any tab stop (indigenous_innovation
  p5: "b." in Arial 12 ends on a 1280-twip stop, the text goes to the 1440
  indent); a plain tab also skips a stop its pen lands on (12 half-width
  letters = 72pt before a 1in stop → the next stop). `list_label.rs`
  `text_hanging` / `label_device_width`.
- An exact-spaced line fits a page only with its whole box (p27: 25 lines
  of 25pt on a 648pt body, the 26th moves on although its 12pt glyphs would
  end 0.5pt above the margin); automatic spacing may hang its trailing
  leading past the margin, as before. `pdf/mod.rs` `line_fit_height`.
- Paragraphs up to three before a float's anchor wrap beside it (p26a: two
  short titles before an empty paragraph anchoring a page-tall wrapSquare
  picture 82.7pt above it both sit beside the picture, at y 0; p26b: both
  beside a 300pt picture hoisted 60pt above a text-carrying third
  paragraph). Word places the float from the paragraphs between laid out at
  full width, then re-wraps them all. The look-ahead in `pdf/mod.rs` now
  scans three plain paragraphs ahead and addresses the anchor it hands on
  to the float's block (`PageBuilder::float_anchor_for`); a float
  overhanging a margin has no strip on that side, not a negative one
  (an Erasmus cover 71.5pt into the left margin left 147pt on the right).
- VML lengths: `2in`, `cm`, `mm`, `pc`, `px` convert to points (an OLE
  preview `height:2in` fell back to its 240pt `dyaOrig` and moved a page).
- A page- or margin-anchored floating table positioned above the text
  already on its anchor's page, too tall to fit below that position, starts
  on the next page at the same position and breaks from there; the anchor's
  page keeps its text (p24f, compat 15, 70 lines then a 1.3-page table at
  tblpY 72pt: page 2 ends half empty, the table fills page 3 and part of 4,
  the anchor paragraph follows it). One that fits stays, and Word re-flows
  the page's text below it (p24a compat 14, p24g compat 15; we still draw
  such a table over the text). Compat 14 with a too-tall table keeps a row or
  two on the anchor's page and the text below them (p24c, p24d; unmodelled).
  `table.rs` `render_table`.
- Fixed-layout tables (`tblLayout fixed`) with a percentage or auto width
  take their columns from the cells' `tcW`, not the saved grid (p28a–g,
  0.2pt): 720-twip cells under a 2746…-twip grid make 36pt columns in a
  `tblW auto` table; in a 100% table pct cells take their share and dxa
  cells share the rest in proportion; shares that overrun the width starve
  columns from the right down to 12pt (Word then sets one letter per line).
  A column only spanned cells size, or a table over 100% (a fixture's 105%
  table fills its 478pt grid), keeps the saved grid. `table_layout.rs`
  `fixed_layout_columns`. Unprobed: how a spanned cell's preference splits.
- A paragraph that opens with a page break still sets the break on its first
  line, and that line must fit first; one that does not moves to a page of
  its own and breaks from there, leaving that page blank (p29a–c: 25pt exact
  break line under 23pt of room → blank page, 15pt → none). The space above
  it may run past the page bottom (erasmus_plus, nabl_lab). `pdf/mod.rs`,
  explicit break branch.
- Revised footnote lines bar like body lines (Word draws them); ours drew
  every page's footnote bars on page 1's header layer, now on their page.

**Change bars**: black, 0.72pt wide, in the left margin at
`margin_left − min(36pt, margin_left / 2)` (margins 0.5/1/1.25/2in → 18, 36,
54, 108pt; with mirrored margins the outside margin: right on odd pages at
`page_width − margin_right + min(36, margin_right / 2)`); x on a 0.24pt step
rounded down, y on a 0.48pt step rounded. One bar per revised line over its
**line box**: the font's ascent above the baseline, the rest of the line
advance below (single, 1.15, double and exact spacing agree), and on the
paragraph's first line the paragraph's own share of the gap above it
(`max(0, space_before − previous space_after)`), on its last line the space
after. Adjacent revised lines' bars join. Two columns still bar at the page's
left margin. Table rows: one bar over the row's box, borders included
(p06: 638.88–680.16 over three rows; ours abuts per row, 0.48pt lower at
the foot, since 2026-11-05).

## Implementation (branch `tracked-changes`)

- `docx::set_markup` / `docx::markup()` (thread-local, set per conversion by
  the `*_with` API functions) and `docx::author_index`.
- `runs.rs`: `collect_run_nodes` keeps `w:del` / `w:moveFrom` in markup;
  `w:delText` / `w:delInstrText` read like `w:t` / `w:instrText`;
  `mark_revision` sets colour + underline / strike + `Run::revised` from the
  innermost revision wrapper (`revision_of`) or `w:rPrChange`.
- `paragraph.rs`: `Paragraph::mark_deleted`; a revised mark or `pPrChange`
  flags all runs. `mod.rs` merges a deleted-mark paragraph into the next in
  the final view; `collect_block_nodes` keeps fully deleted paragraphs in
  markup, and a deleted mark's `sectPr` ends the section only in markup.
  `tables.rs` drops deleted rows in the final view. `numbering.rs`: deleted
  items do not advance the counter. `sections.rs`: `titlePg` by value.
- `pdf/layout.rs`: `WordChunk::revised`; `render_paragraph_lines` records
  each revised line's box in a thread-local; `pdf/mod.rs` draws the bars per
  page after layout (phase 2b').

Probe parity (ours vs Word, inline export): text and reflow identical; bars
identical where lines are the paragraph's middle lines or table rows.

Known gaps (ponytail ceilings): footnotes and text boxes get no bars in
markup mode (headers and footers do); moves untested; mirrored margins use
the right-side rule only on odd pages; a renumbered item's colour 0078D4
is the one measured with a single author (its rule with several is open).

## Testing

- Unit tests in `src/docx/runs.rs` (markup keeps deleted text, colour,
  underline/strike replacement).
- Corpus (local, 150 tracked documents without comments, Word for Mac inline
  references): see the scores section below.
- Fixtures: a fixture with revisions keeps `reference.pdf` = final view
  (online export) and may gain `reference-markup.pdf` = markup view (local
  export with `--revisions inline`, untagged). The visual suite scores the
  latter as its own entry, `<group>/<case>@markup`, from our `--markup`
  rendering (`tests/output/<group>/<case>/markup/`), with its own baseline
  and hashes. `cases/tracked-changes-inline` (probe p01: inserted and
  deleted runs, two long deletions and insertions) carries both: final 76.3,
  markup 69.1 Jaccard (2026-11-05, baselines pending).
- Existing: bush_fires (final) must not move; door_air's local reference has
  deletion balloons and stays on the final view.

## Scores

Local corpus sample, 150 tracked documents without comments, Word for Mac
inline references (`tools/corpus_score.py`, labels `tc-final` / `tc-markup`,
2026-11-05):

| view | Jaccard | SSIM | page count wrong |
|---|---|---|---|
| final (default, main) | 28.4 | 38.5 | 65 / 150 |
| markup (`--markup`) | 62.2 | 75.9 | 13 / 150 |

(Markup 61.2 before the `titlePg w:val=0` fix, which took one redline from
6 to 73.) The fixture suite is unchanged in the final view except
scraped/catalan_university_transfer_offer (+18.5 Jaccard from the deleted-row
drop) and bush_fires (+0.01, deleted-mark merge).

## Gap analysis on the published tracked sets (2026-11-05)

The competitor publishes per-document scores for 317 tracked documents
without comments, 24 with comments and tracking, and 1 with comments only
(all Word for Mac inline-markup references). Rebuilt locally from the full
corpus (`tests/output/tc/jub_sets.py`, symlink dirs in `tests/output/tc/jub/`;
`tc_gap.py <label> <group>` joins our corpus scores with theirs and prints
delta buckets plus an XML feature census of the worst documents).

| set | ours, markup (J / SSIM / pages wrong) | competitor | we win / lose |
|---|---|---|---|
| tracking, no comments (317) | 73.2 / 86.5 / 4 (2026-11-07, Microsoft PhagsPa added to the fonts; 73.0 / 86.4 / 4 after round 4, 72.9 / 86.3 / 4 after round 3, 71.7 / 85.4 / 9 in the evening, 67.9 / 82.1 / 11 in the morning) | 62.2 / 80.4 / 11 | 259 / 58 |
| comments + tracking (24) | 68.0 / 86.2 / 0 (2026-11-07 unchanged; 2026-11-06, round 4 unchanged; 2026-11-05 night; 67.9 in the evening, 55.5 / 72.5 / 0 in the morning) | 29.3 / 48.6 / 4 | 24 / 0 |

Round 3 (labels `tc-gap8` → `tc-gap9`, `tc-comments9` → `tc-comments10`):
317 set 71.7 → 72.6 → 72.9 Jaccard; eight documents +14 to +48 (deleted
page-anchored tables that now start on the next page), then four more +13
to +36 (the exact-line foot, the OLE preview, the Erasmus cover), one −1.0
(its deleted table now on Word's page, the text beside it not yet
re-flowed); page misses 9 → 6 → 4; the comments set 67.9 → 68.0.
Round 4 (labels `tc-gap9` → `tc-gap12`, `tc-comments11`): 317 set 72.89
→ 72.98; the 9-vs-17-page document 18.7 → 48.5 (16/17 pages: the
fixed-layout columns, then the break line that leaves Word's blank page),
one document −0.9 from the fixed-layout columns.
2026-11-07: Microsoft PhagsPa (6.00;O365, Office's cloud font cache) added to
`fonts/` and the assets repo: 317 set 72.98 → 73.18 (labels `tc-gap13`,
`tc-comments12`), the two 3-vs-4 siblings +19 and +40.

The competitor's current release, rescored locally with the same scorer and
Word's markup palette (`--revisions word --compress`), reads Word's own font
folders (Word.app's DFonts, Office's cloud font cache) and resolves the same
faces with or without ours (byte-identical output): 66.9 / 85.6 / 3 on the
317 set (we lead 240, it leads 66, more than 1 point; 33 ten or more behind,
9 thirty or more), 58.1 / 83.3 / 0 on the comments set (19 / 5). The table
above keeps their published scores. Its leads are mostly text on another
line or page than Word's (the press-release header family among them); its
markup colours also follow Word's per-author memory where ours cycle by
first appearance.

Fonts still missing for these sets (our fallback log, 341 documents): not on
this Mac and only subset in Word's PDFs: Californian FB (9), Berlin Sans FB,
Roboto Medium / Light, Lato SemiBold, HGPSoeiKakugothicUB.

**CSS font lists** (checked 2026-11-07, 18 documents): HTML pasted into Word
leaves CSS lists as font names, `w:rFonts w:ascii` and the fontTable both
saying `"Calibri", Arial, sans-serif` (also `"Times New Roman", Times, serif`
and `"Open Sans", sans-serif`), with fontTable altName Cambria. Word takes the
whole string as one unknown name: every such run in nine references is drawn
in Cambria (one Calibri-Italic run among them), never in Calibri, Arial or
Open Sans. Ours matches because `first_css_name` keeps the quotes, so
`"Calibri"` finds nothing and the altName wins. Stripping the quotes, or
walking the list, would break it.

**Helvetica Neue** (1 document): the only fallback where Word for Mac
differs, and on purpose: a macOS-only family with an altName counts as
missing (`src/fonts/mod.rs`), since the online export runs on Windows, which
lacks it. This corpus's Mac references embed HelveticaNeue; ours draws the
altName Arial, as Word online would.
| comments only (1) | 70.8 / 88.4 / 0 | 60.0 | |

(317 set before the page-top split fix below: 66.5 / 80.6 / 25.) 108 of the
317 are ten or more Jaccard points ahead of the competitor, 45 ten or more
behind, 17 of those thirty or more.

**Found and fixed while analysing**
- A multi-line paragraph that opens a page and outruns it was never split
  (the split path was skipped at the page top): one template of 108 manual
  line breaks, double spaced, took one page instead of five. 16 of our 25
  page-count misses; those documents gained 23–50 points. A single oversized
  line (case150's 700pt exact spacer) still stays put.
- List items whose mark alone is deleted were numbered as deleted items
  (1.1, 1.1, 1.1 instead of 1.2, 1.3, 1.4).
- `titlePg w:val="0"`, deleted section breaks (above).

**Fixed 2026-11-05 (second round, three commits on the branch)**
- Tabs after a wrapped word resolve from the new line's pen (probe
  p19_tab_wrap: "Hub<tab>30<tab>$45" at 72/108/144 after a wrap at the
  space; a tab that wraps by itself still lands on the new line's first
  stop). The inventory list 24.9 → 84.0; 317 set 67.9 → 68.2.
- A `w:pStyle` naming a missing or character style falls back to the
  default paragraph style, not docDefaults (the decree's header pStyle is
  "BodyTextIndentChar": Word prints Normal's Arial 9, our Calibri 11 + 8pt
  after pushed the body 4pt down). General rule, final view too.
- List numbering in markup: deleted items number by the original stream
  (plain + deleted), live items by the final one (plain + inserted), a
  live item whose numbers differ shows both, original first (probe
  p11b: 1, 2, 3, "4.2.", 3, 4, 5, "6.5."; an all-deleted list 1, 2, 3).
  Word strikes the original and underlines the final number in its
  numbering-change colour (0078D4 with one author); ours print plain.
  The decree 15.1 → 72.6.

**Fixed 2026-11-05 (third round: floats of split paragraphs)**
- A split paragraph drew none of its floating pictures (the split path
  returned before the float rendering) unless one narrow case matched.
  Now a float whose anchor precedes the last text hangs from the paragraph's
  top on its first page; one anchored after all text (`anchor_run`, the run
  index at the anchor) goes to the page its last line lands on, hung from
  the fragment's top there (probes p20c/p20e, Word for Mac, compat 14:
  +20pt → y 92 on page 2; corpus cover: 101.55 − 82.7 = 18.85 on page 5).
- Compat 14 keeps a negative paragraph offset only down to the page edge
  (p20c/p20d: −82.7pt from a page-top paragraph → y 0; compat 15 stops at
  the margin as before). A wrapping picture wider than the page starts at
  x 0 (p20d, the cover); one that fits keeps a negative x (cyprus_ucits).
- The final fragment's lines around its trailing float: lines whose box
  meets the picture's band sit beside it when a side is wide enough (x 249
  by a 150pt picture at 90, p20c/p20e match Word to 0.13pt), else below;
  when below is off the page the fragment keeps one line (those above the
  band, or the page's first line) and the rest move on with the float, and
  a lone anchor line sits at the page bottom above the footer (the cover:
  "Montering", an empty line and "Med forbehold" take one page each,
  "Med forbehold" at 782 vs Word's 779).
- A page break ending a split paragraph is honoured (it was dropped with
  the early return). swiss_materials loses 4 points from it: Word ends
  that paragraph exactly at page 3's bottom, we split it one line later,
  so the break now adds a page where the ignored break happened to match.
- The cover document: 8.9 → 64.9 Jaccard, 10/10 pages.

**Remaining gaps, ranked by evidence**
3. Cover-page redlines (8.9 vs 68.8, Word 10 pages, ours 8): paragraph 1
   holds a deleted full-page picture (615 × 797pt, wrapSquare, anchored
   −71.5pt from the column and −82.7pt above the paragraph) plus the
   inserted text of the whole first page in one paragraph of `w:br` lines.
   Word prints that text full-width at the top of page 1 (over the
   picture); we wrap it one letter per line in the sliver beside the
   picture ("A", "C", "B", "-" at x 72–79), so page 1 holds the whole
   paragraph and the document loses two pages. Rule to measure: a wrapped
   float that leaves less than some minimum beside it (Word: the lines
   ignore it, or the float becomes top-and-bottom) — probe a wrapSquare
   picture wider than the column with a text paragraph anchored to it.
4. ~~Verdana list (28.3 vs 75.1)~~: environment, not a lookup bug. The
   runs use the "Preformatted Text" style (DejaVu Sans Mono, fontTable
   `altName` Verdana). Word for Mac took the altName; our discovery finds
   DejaVu Sans Mono in this Mac's `~/Library/Fonts` (also matplotlib's
   copy) and uses it, so the lines wrap differently. CI has no DejaVu;
   nothing to change in the code.
5. ~~Table-heavy redlines~~ (the four 6–7-point comments-set documents):
   their comments sit on wholly deleted tables and Word prints no pane; we
   zoomed the page for one. Fixed 2026-11-05 (6.7 → 76.4 on the first);
   the deleted TOC regeneration went with it.
6. Remaining page-count misses (9, evening): the 9-vs-17 document is a
   fixed-layout table with a duplicated 12-column grid and 720-twip cells
   where Word wraps one letter per line in a narrow column (the competitor
   fails it too); the three one-page-short siblings drift 1–3pt per
   paragraph from page 1 (title −1.2, −3.5 by the first body paragraph,
   −6.0 by mid-page) until a two-line bullet fits above the footnote area
   where Word's does not: generic spacing rules, not markup. The others
   (543 deleted marks and 52 row revisions, 14 deleted pictures) unexamined.
6b. Round 3 (2026-11-05 night): the inserted-picture descent, the label
   that fills its indent and the too-tall page-anchored floating table
   (rules above, probes p22–p24, one commit each) take the 3-vs-4 sibling to
   4/4 pages (its remaining −0.9pt per page is the header font) and the
   label-filled siblings' labels to Word's x. Still open from the page-count
   misses:
   - A family of siblings' constant −1.2pt: their header title is
     Microsoft PhagsPa Bold 18 (an Office cloud font; not in `fonts/`, not on
     this Mac — Word embedded it from its cloud cache); we fall back to
     Calibri. Environment, not code (a what-if with Calibri in the header
     still gives Word 4 pages).
   - A document under a wholly deleted header (5 vs 4): Word prints an
     extra empty black 12pt line
     (docDefaults spacing after 200 / line 276 and its firstLine 720 indent)
     between the wholly deleted header and the first body paragraph that no
     `w:p` in the body or any header accounts for. What-ifs that keep it:
     no even/first header refs, no default header at all (it then follows
     the empty header), header marks live, header row live, header table
     removed, body's last paragraph removed, first paragraph's mark plain,
     sectPrChange removed; docDefaults firstLine 0 moves it to the margin.
     A minimal document with the same start (p25a–d) does not show it.
     Unexplained.
   - A labelled list (5 vs 6): after the label fix a 24pt exact line still
     fits at the foot of our page 2 (baseline 771.2 on an A4 page whose
     body ends at 769.9; Word moves it on), then an extra page break.
   - A press-release family whose header holds an inserted and a deleted
     copy of a `wpg` group (text boxes): page 1 matches Word closely (round
     4 look); the loss is page 3, where Word starts the body at the top
     margin under a different header that overlaps it and we start it below
     that header (~60pt lower). Header-extent rule, unexamined.
   - Page-anchored floating tables that fit: Word re-flows the page's text
     below them (p24a, p24g); we draw them over it.
7. Cosmetic, mostly done (bars over empty paragraphs and spacing, bridged
   strikes, coloured labels, bars in headers and footers, strike / underline
   lines on deleted / inserted pictures, one bar per revised table row,
   the label's tab gap, two-number labels in the numbering-change
   colour). Closed 2026-11-05.

## Out of scope

- Revision balloons (deletions / formatting in the margin pane).
- Word's Original view (the document as it was before the changes).
- Creating redlines (comparing two documents).
