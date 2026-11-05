# Tracked changes (revision markup)

A document with tracked changes keeps each edit as a revision: `w:ins` /
`w:del` runs, `w:moveFrom` / `w:moveTo`, revised paragraph marks and table
rows, and property changes (`w:rPrChange`, `w:pPrChange`, …). Today we draw
the **final** view (insertions plain, deletions dropped), which is Word's "No
Markup" view and what Word's online export prints.

**Decided (2026-10-16):**
- **Final stays the default.** Markup is opt-in, from the library and the CLI.
- Fixtures with revisions get a **separate markup reference** next to the
  final one where it applies (see Testing).

## Evidence so far

- Word's online export prints without markup (`online-references.md` §4b);
  `scraped/bush_fires_act_comparison` (online reference, tracked changes, no
  markup) agrees. That is the final view we draw.
- Word for Mac's local export ("Quartz PDFContext", untagged) of 300 tracked
  documents from an external corpus sample (local only) shows **All Markup
  inline**: inserted text underlined, deleted text struck through, both in the
  author's colour (`#4E6AED` = 0.306 0.416 0.929 for the single author of the
  doc inspected; underline and strike lines are filled paths in that colour),
  deleted content kept in the layout (it wraps and paginates), and a black
  change bar in the left margin beside changed lines. Deleted pictures are
  drawn as usual; only the bar marks them.
- No revision balloons in any of the 300 ("Deleted:" / "Formatted:"), not even
  the 150 that also have comments: those show the comment pane only.
  `scraped/door_air_cooling_unit_spec` (older local reference) is the
  exception: deletions in balloons. Balloon mode is a Word view setting, not
  in the docx; we follow the inline majority.
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

## Current code

- `docx/runs.rs` (`collect_run_nodes`): `w:ins` / `w:moveTo` pass through,
  `w:del` / `w:moveFrom` are skipped.
- Final-view gaps: a deleted paragraph mark doesn't merge its paragraph with
  the next one, and deleted table rows aren't dropped (to verify).
- `Run` already carries colour, underline and strikethrough, so inline markup
  needs no new drawing code.
- Public API: `convert_docx_to_pdf(input, output)` and
  `convert_docx_bytes_to_pdf(bytes, output)`, no options; the CLI
  (`src/main.rs`, clap) has none either.

## Plan

### 0. Probes (before any rule)

python-docx documents with revisions injected by ZIP post-processing,
exported with `tools/word_export.py --preset print` (local, the markup the
external sample shows; first check the preset reproduces it) and the default
online preset (final view). Measure:

1. Author colours and their order (1–8 authors); how an author gets its index
   (order of first appearance in the document, or in the revision list?).
2. Insert / delete line style (single underline, single strike?), position and
   thickness; how they combine with the run's own underline / strike / colour.
3. Move styling (`w:moveFrom` / `w:moveTo`).
4. Change bar: x position (left margin vs outside margin with mirror margins,
   columns, tables, headers/footers), thickness, vertical extent per line.
5. What each property change draws (`rPrChange`, `pPrChange`, `sectPrChange`,
   table property changes): expected only a bar.
6. Deleted / inserted paragraph marks; deleted / inserted rows; `cellDel`.
7. Deleted pictures and floating objects; deleted field codes
   (`delInstrText`: the cached result struck through?).
8. List numbering of inserted and deleted items in both views.
9. Revisions in headers, footers, footnotes and textboxes.
10. Comments + revisions together: page scale and bar placement with the pane.

Record each rule below with the probe that shows it.

### 1. Options API

```rust
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum RevisionView { #[default] Final, Markup }

#[derive(Clone, Debug, Default)]
#[non_exhaustive]
pub struct ConvertOptions { pub revisions: RevisionView }

pub fn convert_docx_to_pdf_with(input, output, &ConvertOptions) -> Result<(), Error>
pub fn convert_docx_bytes_to_pdf_with(bytes, output, &ConvertOptions) -> Result<(), Error>
```

The existing functions keep their signatures and use the defaults. CLI:
`--markup`. The view goes into parsing (`docx::parse*`), since it decides
which content exists.

### 2. Model and parsing

- `Run::revision: Option<Revision { kind: Insert | Delete | MoveFrom | MoveTo, author: u16 }>`.
- Markup view: keep `w:del` / `w:moveFrom` content, read `w:delText` and
  `w:delInstrText`; mark paragraph-mark and row revisions on `Paragraph` /
  table row.
- Final view: unchanged, plus the two gaps above (merge paragraphs across a
  deleted mark, drop deleted rows) once probes confirm Word does both.

### 3. Inline markup

Apply the author colour + underline / strike to revised runs as run
properties, so layout handles deleted text (line breaks, pagination) with no
layout changes.

### 4. Change bars

A per-line "revised" flag (a revised run, a revised paragraph mark, a
`pPrChange` / `rPrChange` on the paragraph), drawn by the measured rule in
`pdf/`.

### 5. Tables, pictures, moves, property changes

Deleted rows and pictures stay with their markup; moves as measured;
property changes draw bars only.

### 6. Comments + revisions

The comment pane exists (`pdf/comments.rs`); check page scale and bars when
both are present.

## Testing

- Unit tests (`#[cfg(test)]` in `src/docx/runs.rs` etc.): deleted run kept in
  markup, skipped in final; deleted paragraph mark merges in final; author
  index assignment.
- Fixtures: a fixture with revisions keeps `reference.pdf` = final view
  (online export, as for every fixture) and gains `reference-markup.pdf` =
  markup view (local print export, untagged: accessibility scores N/A). The
  harness renders the markup view when `reference-markup.pdf` exists and
  scores it as its own entry (`<case>@markup`) in baselines and hashes.
- New probe-derived fixtures in `cases/` for the core rules (colours, bars,
  paragraph marks, rows).
- Existing: bush_fires (final) must not move. door_air's local reference has
  markup with deletion balloons; it stays on the final view until we decide
  whether to support balloons.
- Targeted external-corpus scoring of the tracked samples after each phase
  (rendered with `--markup`), plus the full fixture suite for regressions.

## Out of scope

- Revision balloons (deletions / formatting in the margin pane).
- Word's Original view (the document as it was before the changes).
- Creating redlines (comparing two documents).
