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
  inserted one gets the underline. (We draw the bar only.)
- Deleted fields (`w:delInstrText`): the cached result struck through.
- Lists: an inserted item takes its final number; a deleted item shows a
  number without consuming one (1, inserted 2, deleted 2, 3). The label is
  coloured and underlined / struck like the text (we leave the label plain).
- Headers and footers: marked the same way, bars at the same x (we leave
  them final for now).
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
left margin. Table rows: the row's line boxes (= between its borders here).

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

Known gaps (ponytail ceilings): no space-before/after share on a paragraph's
first/last bar; strike drawn per word, not bridged across spaces; deleted
pictures get no strike lines, inserted ones no underline; list labels not
coloured; headers/footers/footnotes/text boxes stay final in markup mode;
moves untested; mirrored margins use the right-side rule only on odd pages.

## Testing

- Unit tests in `src/docx/runs.rs` (markup keeps deleted text, colour,
  underline/strike replacement).
- Corpus (local, 150 tracked documents without comments, Word for Mac inline
  references): see the scores section below.
- Fixtures: a fixture with revisions keeps `reference.pdf` = final view
  (online export) and may gain `reference-markup.pdf` = markup view (local
  export with `--revisions inline`, untagged). Harness support for scoring
  `reference-markup.pdf` as its own entry is still to do.
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
| tracking, no comments (317) | 67.9 / 82.1 / 11 | 62.2 / 80.4 / 11 | 233 / 84 before the pagination fix |
| comments + tracking (24) | 55.5 / 72.5 / 0 | 29.3 / 48.6 / 4 | 24 / 0 |
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

**Remaining gaps, ranked by evidence**
1. Tab-separated revised lines wrap differently after the first line (an
   inventory list: Word keeps "Hub 30 $45" on line two, we push "30 $45" to a
   third line, and everything below shifts: 24.9 vs 75.1). A tabs-after-wrap
   rule, not markup-specific; the ink diff is in `tests/output/tc/png-gap/`.
2. Paragraph spacing drift on pages of wholly inserted paragraphs
   (`w:pPr/w:rPr/w:ins` with heading styles, bottom borders, `w:tabs num`):
   our lines sit 4pt higher from the first line and the gap after a heading
   block differs (15.1 vs 67.0 on a two-page decree). Needs a line-by-line
   measurement against the Word export; general paragraph-spacing rules.
3. Cover-page redlines with many anchored pictures and text boxes
   (`anchor` appears in 53% of the worst 30 against 18% overall): the title
   text box lands in the body and the cover picture on page 2 (8.3 vs 68.8).
   Floating-object layout, not markup.
4. Font resolution: a Verdana-styled list prints in DejaVu Sans Mono (28.3
   vs 75.1) although `fonts/` has Verdana; check style-based `rFonts` and
   the Greek-glyph fallback choice.
5. Table-heavy redlines with dozens of inserted/deleted rows, deleted fields
   and comments (the four 6–7-point documents in the comments set, 9–11
   pages each; the competitor scores 5–6 there too).
6. Remaining page-count misses (11): a portrait/landscape document at 9
   pages against Word's 17, three siblings at one page short with row
   revisions and deleted fields, one with 543 deleted marks and 52 row
   revisions, one with 14 deleted pictures.
7. Cosmetic: no change bars in headers/footers (the marks themselves print),
   bars stop at the line box, per-word strikes, no strike/underline lines on
   deleted/inserted pictures, list labels uncoloured, two bars per table row.

## Out of scope

- Revision balloons (deletions / formatting in the margin pane).
- Word's Original view (the document as it was before the changes).
- Creating redlines (comparing two documents).
