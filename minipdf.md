# MiniPdf: where it is ahead of docxide-pdf

Reviewed 2026-09-04 against `mini-software/MiniPdf` at commit range up to 2026-09-04
(.NET engine v0.40.0, Rust crate v0.5.0). Findings come from reading the C# source in
`src/MiniPdf/` and the benchmark tooling in `tests/`, not from running the binary.

Scope note: the feature analysis below is of the .NET engine, which is the mature one.
MiniPdf also ships an independent Rust rewrite (`minipdf-rs/`, about 4,000 lines for DOCX
and PDF) that implements a subset of it. Our benchmark viewer (`tools/engine_compare.py`)
compares against the Rust crate only, because that is the implementation in our language
and the one whose progress matters to us.

MiniPdf is a heuristic-first converter. It draws all Latin text in non-embedded base-14
Helvetica, measures text with hardcoded width tables, and uses a constant line-height
factor per font family. On almost every rendering axis docxide-pdf is ahead. This file
documents the exceptions: things MiniPdf does that we do not, described in enough detail
to decide whether and how to implement them.

Corpus frequencies below come from `analyze-fixtures --grep` over our 129 real-world
DOCX fixtures.

---

## 1. Rendering features we lack

### 1.1 Table of contents generation from headings

**What MiniPdf does.** `DocxReader.cs` detects a TOC field that has no cached result.
`IsTocFieldWithoutCachedResult` walks the paragraph's runs, tracks `w:fldChar`
begin/separate/end nesting, concatenates `w:instrText`, and checks whether the instruction
starts with `TOC`. If no `w:t` text appears between `separate` and `end`, the paragraph is
replaced by a `DocxTocPlaceholder` carrying the outline range parsed from the `\o "1-3"`
switch (default 1 to 9).

After the whole body is read, `ExpandTocPlaceholders` builds the entries:

- Headings are any paragraph with `w:outlineLvl` 0 to 8, or a style id matching
  `Heading<N>`. Paragraphs whose style starts with `TOC` are skipped.
- Page numbers are estimated by walking the element list and incrementing a counter at
  every explicit page break and every non-continuous section break. Natural overflow is
  not counted, so the numbers are wrong for any document longer than its explicit breaks.
- Each entry becomes a synthetic paragraph: heading text, a tab, the page number. It uses
  a right-aligned dot-leader tab at the full text width, a 12pt left indent per level,
  11pt font, exact 24pt line spacing, style id `TOC<level>`, and 24pt space after the last
  entry.

**Our status.** `FieldCode` covers `Page`, `NumPages`, `StyleRef`, and `PageRef`. A TOC
field renders its cached result runs, which is what Word shows unless
`w:updateFields` is set in `settings.xml`. An empty TOC field renders nothing.

**Corpus.** 3 fixtures contain a TOC instruction, all with cached results. 0 fixtures set
`w:updateFields`. So this affects only programmatically generated documents.

**Assessment.** Word does not regenerate a TOC on export unless `w:updateFields` is
present. Generating one unconditionally would diverge from Word for hand-authored files
that deliberately carry an empty field. If we implement it, gate it on `w:updateFields`
or an explicit option, and compute page numbers from our real pagination in a second
pass rather than by counting breaks. The paginator extraction on the roadmap would make
that second pass cheap.

### 1.2 Image cropping (`a:srcRect`) — FIXED 2026-09-04

**What MiniPdf does.** `DocxReader.cs` reads `a:srcRect` attributes `l`, `t`, `r`, `b`
from the `a:blipFill` and divides by 100000 to get fractions of the source image. For
raster images it calls `TryCropImagePng`, which decodes with System.Drawing, cuts the
sub-rectangle, and re-encodes to PNG. For EMF and WMF it rasterizes the metafile at its
native aspect ratio to a 512px-high bitmap and then crops the bitmap. Both paths are
guarded by `Compat.IsWindows()` and silently skip cropping elsewhere.

**Our status.** Implemented 2026-09-04. `docx/images.rs` parses the crop onto
`EmbeddedImage.src_rect`; `pdf/images.rs` wraps a cropped image in a Form XObject with
BBox `[0 0 1 1]` whose content draws the source scaled and offset, so all draw sites are
untouched and EMF forms nest as-is. Unit tests cover the parser and the matrix; `case78`
is the handcrafted fixture (four crops incl. a negative one) and scores J 99.6% /
SSIM 99.4% against its Word reference. Before this, a cropped picture rendered the full
source squeezed into the frame.

**Corpus.** `analyze-fixtures --grep "a:srcRect"` matches 13 of 129 fixtures, but 11 of
those carry only an empty `<a:srcRect/>` (no crop). Real crop values appear in two:
`brazilian_logistics_study` (Jaccard 17.3%) on 3 anchored PNGs, and
`italian_evaluation_minutes` (34.0%) on one inline PNG signature:

| Fixture | Element | Frame (pt) |
|---|---|---|
| brazilian | `<a:srcRect t="4604" r="1295" b="6879"/>` | 364.5 × 146.8 |
| brazilian | `<a:srcRect t="4505" b="26576"/>` | 416.1 × 174.1 |
| brazilian | `<a:srcRect t="4711" b="5497"/>` | 452.1 × 235.2 |
| italian | `<a:srcRect t="39029" r="28683" b="25937"/>` | 149.2 × 36.0 |

So the corpus impact is two fixtures, not thirteen. The fix changed exactly the three
pages holding those images (brazilian 9 and 10, italian 7), each now matching the
reference crop visually. Scores barely moved (brazilian J 17.29 → 17.17, italian SSIM
43.83 → 43.82) because the images sit at drifted positions, where the old squeezed
rendering happened to overlap the reference ink slightly more.

**Outcome.** Cheaper for us than for MiniPdf: no re-encoding and no platform guard. The
Form XObject's unit BBox clips, and its content draws the source scaled by
`1/(1-l-r)` × `1/(1-t-b)` and offset by `-l` and `-b`, which also covers EMF forms and
negative (outward) crops. Word's reference for case78 confirmed that a negative crop
renders as blank padding inside the frame. A/B on case78 with the source changes stashed
and restored: J 75.7% / SSIM 81.3% without, 99.6% / 99.4% with. The old code still
clears the suite's absolute thresholds, so the accepted baseline is what guards this
feature against regression.

### 1.3 Content-control data binding (`w:sdt` with `w:dataBinding`)

**What MiniPdf does.** `LoadSdtContext` in `DocxReader.cs` runs once per document and
collects three things:

- Every `customXml/itemPropsN.xml` is read for its `ds:datastoreItem/@ds:itemID`, and the
  matching `customXml/itemN.xml` is loaded and keyed by that GUID.
- `docProps/core.xml` is registered under the reserved store id
  `{6C3C8BC8-F283-45AE-878A-BAB7291924A1}`, which Word uses for cover-page controls bound
  to Title, Subject, Author, and similar core properties.
- `word/glossary/document.xml` is scanned for `w:docPart` entries, and the runs of each
  part's `w:docPartBody` are kept by part name.

`UnwrapSdt` then handles each `w:sdt`. If `w:sdtPr/w:dataBinding` is present it parses
`prefixMappings` into an XML namespace manager, evaluates the `xpath` against the store,
and takes the string value. For an inline control the cached runs are replaced by a single
new run carrying the live value and the first cached run's `w:rPr`. If the bound value is
empty and `w:placeholder/w:docPart` names a glossary part, that part's runs are emitted
instead. Block-level controls fall through to plain unwrapping. The same function also
prefers `mc:Choice` over `mc:Fallback` in `mc:AlternateContent`.

**Our status.** `runs.rs` unwraps `w:sdtContent` and renders the cached runs.

**Corpus.** 0 of 129 fixtures contain `w:dataBinding`.

**Assessment.** Word does re-evaluate bound controls on open, so for a document saved by
Word the cached text already matches the store. The cases where cached and live text
differ are documents whose `customXml` or core properties were edited by a tool that did
not refresh the cached runs, which is a template-filling workflow rather than a
Word-authored document. Zero corpus hits. Low priority, but the implementation is
self-contained: a store loader in `docx/mod.rs` plus an XPath evaluator. `roxmltree` has
no XPath, so this would need a small path walker or a new dependency.

### 1.4 `w:lastRenderedPageBreak` as a keep-together hint

**What MiniPdf does.** While reading a paragraph, `DocxReader.cs` sets
`HasLastRenderedPageBreak` only when the marker appears before any visible run or image,
so mid-paragraph markers are ignored. At render time `DocxToPdfConverter.cs` checks the
flag when the cursor is not already at the top of a page. It estimates the paragraph's
flow height with `EstimateParagraphFlowHeight` (images plus wrapped text lines), and
forces a new page only if that height exceeds the remaining space. If the paragraph would
fit, the hint is treated as stale and ignored. A `SuppressNextLastRenderedPageBreak` flag
swallows the next hint after the keepNext logic has already forced a break, so heading
and body do not get separated by a duplicate break.

The effect is that paragraphs Word moved whole to the next page because of keepLines,
widow/orphan control, or table row rules are reproduced without implementing those rules.

**Our status.** Not read. We paginate from the document's own rules and the reference
PDF, which is the design choice for a fidelity project.

**Corpus.** 29 of 129 fixtures carry the marker, 629 occurrences in total.

**Assessment.** The hint is written by Word on save, so it encodes Word's real pagination
of that revision, including decisions we get wrong today. It is absent from documents
produced by other generators and stale after any edit in a non-Word tool. Two uses are
defensible for us. First, as a diagnostic: a test that compares our page-start paragraphs
against the marked ones would localize pagination drift faster than a pixel diff. Second,
as a tie-breaker when a paragraph is within one line of fitting, where our metric error
is larger than Word's decision margin. Using it as a hard break would mask layout bugs
and should stay off.

---

## 2. API and CLI surface

MiniPdf exposes more entry points than our single `convert_docx_to_pdf(&Path, &Path)`.

| Capability | MiniPdf | docxide-pdf |
|---|---|---|
| Path to path | yes | yes |
| Bytes in, bytes out | `convert_bytes_to_pdf`, `ConvertToPdf(Stream, Stream)` | no |
| Format sniffing | `detect_office_format` | no |
| Font registration by bytes | `register_font(name, bytes)` | no, `DOCXSIDE_FONTS` directory only |
| Page size override | A4, Letter, or custom points | no |
| CLI | `minipdf in.docx -o out.pdf --fonts dir --paper-size a4` | `cli` feature binary |
| Other formats | XLSX, PPTX | no |

The bytes API and font-by-bytes registration are cheap and matter for server use. The
page-size override is a convenience that breaks fidelity and does not fit our goal.

---

## 3. Testing and process

### 3.1 Score-gated automatic fix loop

`scripts/Invoke-MiniPdfContributionLoop.ps1` wraps
`.github/skills/skill-minipdf-contribution/scripts/contribution-loop.ps1`. The loop:

1. Runs the benchmark and selects the two cases with the largest visual difference.
2. For each case, a coding agent makes one root-cause change.
3. `Evaluate` re-runs the suite and accepts the change only if the case's overall score
   improves by at least 0.0001 and no other case regresses by more than 0.002.
4. Each case gets at most three attempts before it is skipped.
5. `Validate` runs the full test suite, and `Pr` attaches before/after benchmark evidence.

Our annotation workflow does the same selection by hand. A small script around
`run-tests.sh` and `latest_scores.json` that picks the worst unfixed cases and enforces
the no-regression gate would give us the same loop without changing the metric.

### 3.2 Composite score with page-count term

MiniPdf's overall score is 40% text similarity, 40% visual, 20% page-count match. Text
similarity is Python `SequenceMatcher` over PyMuPDF-extracted page text. The visual term
is 40% raw byte match, 40% ink density over a 20 by 20 grid, 20% top-strip density, and
is lenient enough that nearly everything scores above 0.95.

Our `text_boundary.rs` already compares text at page and line level and is stricter than
their text term. The page-count match was the one term we did not report. Done
2026-09-05: `visual_comparison.rs` writes `ref_pages` and `gen_pages` per case into
`tests/output/latest_scores.json`, and `tools/compact_report.py` ends its summary line
with `N/M page counts match`. Before this, a generated PDF short of pages was scored on
its common pages only and looked fine. Page counts are not baselined; the visual-hash
section already flags any case whose page set changed.

### 3.3 Optional AI difference description

`compare_pdfs.py --ai-compare` sends the reference and generated page images to GPT-4o
when the pixel score falls below a threshold and stores the model's description of the
differences in the report. We do the same interactively by reading screenshots in the
case browser. An automated version would only matter if we ran the loop in 3.1 unattended.

### 3.4 Reference generation from two engines

The DOCX classic suite uses Word COM on Windows by default and LibreOffice as a fallback.
The issue corpus uses LibreOffice. Having a second engine is not an advantage for
Word-fidelity, but their `generate_office_pdfs_docx.py` is a working Word COM driver that
could replace our manual reference step on a Windows box.

---

## 4. Reusable assets (Apache 2.0)

### 4.1 Real-world issue corpus with references

`tests/Issue_Files/docx/` holds 27 DOCX files reported by users, and
`tests/Issue_Files/reference_docx/` holds a reference PDF for each. The references were
produced by LibreOffice, not Word, so they cannot be dropped into `tests/fixtures/` as-is.
The DOCX files are still useful as inputs. Notable ones: `13_IEEE_Style_Paper`,
`14_Thesis_Chapter`, `CCU_article` (21 pages), `nthu_article` (20 pages),
`Template for MSc Thesis` (17 pages), `Class News`, `MODERN LIVING`, `SA8000 ch sample`
and `20260317_sample_CN` (Chinese), `Cooperation Agreement Template`.

**Scan on 2026-09-04** (all 27 files downloaded and run through our release CLI):

- No panics or errors. Slowest conversion 350 ms (CCU_article, 28 pages).
- Page count vs Word's `docProps/app.xml <Pages>`: 18 match. We are over on CCU_article
  (28 vs 21), nthu_article (19 vs 18), SA8000 ch sample (3 vs 2), TestIssue78 (2 vs 1),
  13_IEEE_Style_Paper (2 vs 1), under on issue26050501 (1 vs 2). TestIssue91 says 1 page
  but is clearly a 3-page contract, so the saved count is stale for tool-filled files.
  22 of 27 are Word-authored (app.xml), 3 WPS Office, 1 LibreOffice, 1 synthetic.
- `a:srcRect`: one real crop, `OSCAR WARD.docx`, `<a:srcRect r="77074" b="16438"/>` on an
  inline 128×37pt EMF. The EMF is text-only (EXTTEXTOUTW records), which `docx/emf.rs`
  skips, so today it renders as nothing. It becomes a crop test only after EMF text
  support. Three more files carry an empty `<a:srcRect/>`.
- `w:dataBinding`: `Fabrikam.docx` (70 bindings) is a genuine stale-cache case. Cached runs
  read "Fabrikam, Inc.", "Portland, OR 54321" etc. while app.xml `<Company>`, core.xml
  `<dc:creator>` and the coverPageProps customXml are all empty, so Word shows placeholders
  where we show the cached text. `20260318_issue.docx` (6 bindings) has cached == store.
- TOC: `TestIssue61.docx` is the synthetic file behind MiniPdf's TOC generation. Empty TOC
  field, `w:updateFields`, no styles.xml, no app.xml.
- Features present that our corpus barely has: `w:sdt` in 10 files, VML in 12,
  `widowControl=0` in 15, CJK `eastAsia` fonts in 26. No charts, SmartArt, math, altChunk.
- Several files are personal documents from issue reporters (cover letter, affidavit,
  support letter, contracts). Check before committing any of them as fixtures. Fabrikam is
  a Microsoft sample; the two academic articles are copyrighted papers.

### 4.2 Generated scenario list

`tests/MiniPdf.Scripts/generate_classic_docx.py` produces 180 DOCX files with python-docx,
the same generator we use. The Word reference PDFs are not committed, only PNG renderings
in `tests/MiniPdf.Benchmark/reports_docx/images/`. The scenario names are a useful idea
bank for handcrafted cases we do not have yet:

- Layout: `multi_column_layout`, `sidebar_layout`, `two_column_table_layout`,
  `calendar_layout`, `timeline_layout`, `org_chart`, `newsletter_layout`,
  `landscape_page`, `multi_section_orientation`, `cover_page_with_image`.
- Tables: `thin_border_table`, `thick_outer_border_table`, `dashed_border_table`,
  `double_border_table`, `mixed_border_styles`, `striped_table`, `checkerboard_table`,
  `heatmap_table`, `gradient_rows_table`, `rotated_text_table`, `table_merged_complex`,
  `nested_table`, `wide_table`, `multi_page_table`, `table_column_widths`.
- Text: `underline_styles`, `strikethrough_text`, `superscript_subscript`,
  `highlighted_text`, `special_characters`, `code_block_styling`, `blockquote_styling`,
  `paragraph_shading_patterns`, `bottom_border_paragraphs`, `first_line_indent`,
  `hanging_indent`, `custom_bullet_characters`, `numbered_and_bullet_mixed`.
- Documents: `resume`, `business_letter`, `memo`, `invoice_document`, `contract_template`,
  `legal_document`, `academic_paper`, `bibliography`, `glossary`, `faq_document`,
  `survey_questionnaire`, `medical_form`, `shipping_label`, `certificate_with_seal`,
  `cjk_document`, `right_to_left_text`, `multi_language_document`.

---

## 5. Things that looked like advantages but are not

- **Form checkboxes.** `AddFormCheckboxOverlay` draws a square when a table-cell line starts
  with U+25A1. It exists because Helvetica has no glyph for that character. We embed the
  document's real font and render the glyph directly.
- **Watermark repeat.** Body-anchored `behindDoc` images positioned relative to the page are
  copied onto every page. Word draws a body-anchored image on its anchor page only.
  Watermarks that repeat live in headers, which we already render per page.
- **TOC page numbers.** See 1.1. The numbers are counts of explicit breaks, not pagination.
- **CJK heuristics.** MiniPdf has about 46 font-name substring checks and a hardcoded list
  of Windows CJK font files. We resolve fonts from the actual font table and system
  discovery, which covers the same documents without per-family constants.
- **Text similarity metric.** Weaker than our text-boundary test. Only the explicit
  page-count term is new.
- **OMML math.** MiniPdf linearizes fractions, scripts, and sums into plain text. We render
  the math runs inline with the math font. Neither lays out real equations.

---

## 6. Suggested order

1. ~~`a:srcRect` cropping via clip-and-scale.~~ Done 2026-09-04, see 1.2.
2. ~~Page-count line in the compact test report.~~ Done 2026-09-05, see 3.2.
3. Worst-case selection and no-regression gate script around `run-tests.sh`.
4. Bytes-in, bytes-out API and font registration by bytes.
5. `w:lastRenderedPageBreak` as a pagination diagnostic in tests only.
6. TOC generation gated on `w:updateFields`, after the paginator extraction.
7. Content-control data binding, when a fixture needs it.
