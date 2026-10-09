# Roadmap

## Floating table positioning

- Do not reapply non-negative text-anchor offsets after a non-overlapping table has cleared a preceding float. Preserve topFromText; offsets after a separating paragraph remain (cases162–169).

## Accessibility (IN PROGRESS — started 2026-10-01)

Goal: our PDFs are accessible on their own merits; Word's export is a floor,
not the target, and where we can do better than Word we do (line numbers as
artifacts, three-level tables kept). Failures that come from the DOCX lacking
something (no title → 7.1-9, picture without `descr` → 7.3-1, headings that
skip a level → 7.4.2-1) are expected: we never invent titles or alt text or
renumber headings; the baselines absorb them. Measured by
`./tools/run-tests.sh --test accessibility` (needs `brew install verapdf poppler`;
the test skips with a notice when they are missing). Every fixture gets
**`ua_fail`** (PDF/UA-1 rules our PDF fails on its own, any increase is a
regression; a PDF claiming PDF/UA must fail none), plus three
reference-relative metrics where Word's reference is tagged, all in
`baselines.json` like Jaccard/SSIM (explained in `SCORING.md`):

- **`ua_deficit`** — veraPDF PDF/UA-1 (`-f ua1`, forced because Word writes no
  pdfuaid) rules we fail where Word passes, or where we fail a larger share of
  the rule's checks than Word (so 7.1-3 "content not tagged" can't hide behind
  Word's one stray failure). Count, lower is better, **0 = at least as good as
  Word** on the machine checks. Any increase is a regression.
- **`a11y_struct`** — 1 − edit distance over the structure-tree element types
  in reading order (`pdfinfo -struct`, role maps resolved, Span dropped,
  `Figure+alt` distinct from `Figure`).
- **`a11y_text`** — block texts in structure order (`pdfinfo -struct-text`,
  whitespace-normalised): characters of blocks that match exactly and in
  order. Catches reading order, headers/footers leaking in as content, and
  words merged because no space glyph was emitted.

Per-case detail (deficit rules with descriptions and check counts for both
PDFs) lands in `tests/output/<group>/<case>/generated.deficit.json`; analyses
are cached in `*.a11y.json` next to it (delete them after upgrading
veraPDF/Poppler). `DOCXSIDE_A11Y_GEN=libreoffice.pdf` scores another engine's
PDF with the same yardstick (no baselines written).

**What Word's references look like** (surveyed 2026-10-01): 174/222 are tagged
exports — `Document` root, P/Span/H1–H6/L/LI/Lbl/LBody/Table/THead/TBody/TR/
TH/TD/Link/Figure/Footnote/Textbox/TOC/TOCI, Word's RoleMap (Footnote/
Endnote→Note, Textbox/Header/Footer/InlineShape/Artifact→Sect, Title→H1,
Diagram→Figure, CommentAnchor→Span), `/Tabs /S` on pages, headers/footers as
`/Artifact /Pagination`. Word's own bar is low: no pdfuaid (5-1), catalog
`/Lang` always `en` (real language on Span `/Lang`), DisplayDocTitle without
a Title in 157/174, TH without `/Scope`, Figure `/Alt` copied verbatim from
`wp:docPr/@descr` (41/140 figures).

**Untagged references:** the 61 macOS print-path references were re-exported
with Word's online (tagged) preset on 2026-10-03, except 4 kept local on
purpose — cases/case63, case64, scraped/door_air_cooling_unit_spec and
fonts/missing_font_substitution: Word shows the comment pane only in a local
conversion, and these test it. They score `ua_fail` only.

**Starting point (173 scored):** struct 0 / text 0 everywhere (untagged);
ua_deficit 6–11. Every fixture: 6.2-1 MarkInfo, 7.1-3 untagged content,
7.1-8 no XMP, 7.1-10 no DisplayDocTitle, 7.1-11 no StructTreeRoot, 7.2-34 no
language. Also 7.2-2 outline language (59), 7.18.3-1 no `/Tabs` (25),
7.18.5-1 untagged links (25), 7.21.5-1 font widths ≠ glyph widths (7),
7.21.4.1-1 non-embedded base-14 fallback (3), 7.18.5-2 link `/Contents` (3),
7.21.7-1 missing ToUnicode (1), 7.21.8-1 `.notdef` referenced (1).

**Done (2026-10-01, one commit each, every one with no visual change):**
catalog `/Lang`, XMP, DisplayDocTitle, `/Tabs` · boundary space glyphs
(appended to the word's own `Tj`) · structure tree with P/H1–H6, artifacts by
default (`pdf/tagging.rs`) · object streams (`pdf/objstm.rs`, −15% size) ·
L/LI/Lbl/LBody · deterministic ToUnicode (lowest code point per CID) · tables
(THead/TBody/TR/TH/TD, TH `/Scope`, `/ColSpan`) · word boundaries at tabs and
`w:br` · Figures (alt from `docPr@descr`, decorative → artifact) · Links +
OBJR + `/Contents` · TOC/TOCI inside TOC fields · PAGEREF `\h` links ·
outlineLvl 9 = body text · built-in "heading N" levels · footnote/endnote
Notes · cell links, notes and lists.

**Done, round 2 (2026-10-01, `dad4444`..`6e50ff06`, no visual change):**
`ua_fail` scored on all 221 fixtures · textboxes → `Sect > P` hoisted after
the anchor paragraph (`Tags::hoist`) · floating pictures → hoisted Figure,
inline pictures → Figure inside their P (effects stay artifacts) · line
numbers as artifacts · nested tables inside their TD/TH · footnote/endnote
marks → Link with a GoTo to the note, holding the Note ("Footnote 3"
/Contents) · symbol-font glyph widths in `/W` (7.21.5-1) · `pdfuaid:part=1`
claimed only when title, Figure alt, heading order, embedded fonts and no
.notdef all hold (13 fixtures, all clean). Harness: `run-tests.sh` runs every
suite even when one fails; compact report shows `UaFail`.

**Done, round 3 (2026-10-01, `7587e1d9`..`b9909a03`):** caps/small caps
keep their source letters as `/ActualText` on a Span structure element (not
nested marked content, which Poppler's structure reader truncates after) ·
text-shadowed words read once · note marks drawn and linked in table cells
(the only visual change: 6 fixtures) · per-run language: inherited
`w:lang`, catalog `/Lang` = the text's dominant language (`document_lang`),
`/Lang` Spans for passages in another language. Scores unchanged except the
cell marks (erasmus_plus text 68 → 81%).

**Done, round 4 (2026-10-01, `5010019f`..`4525f72d`):** paragraph styles
inherit `w:lang` through basedOn (the write-back skipped it, so styles based
on Normal fell to docDefaults: lithuanian_public_information_law's ~157
`/Lang en-US` Spans are gone; the docDefaults-policy idea was a misdiagnosis)
· hoisted Figures/Sects follow their anchors' XML order (`anchor_seq`,
german_mezzo struct 94.3 → 100%) · Wingdings ToUnicode → real Unicode from
the font's glyph names (▪ ✔ ☺ …; Word itself is inconsistent: samtale's
reference has ☺, irish_school's keeps U+F0A8) · theme slots
`minor/majorEastAsia` and `minor/majorBidi` for Latin text (`ThemeFonts::slot`;
cs falls back to the themeFontLang `@bidi` script font) — the 3 fixtures that
use them get Word's Malgun Gothic/Arial instead of Type1 Helvetica
(east_asia_conference_form J 12 → 21%, SSIM 33 → 63%) · a font that resolves
nowhere gets Arial/Liberation Sans/Arimo/Helvetica/DejaVu Sans before Type1
(case60, multi_font) · a Unicode space the font lacks draws the font's space,
not `.notdef` (U+202F in macOS Arial 5.01, Aptos Italic). One commit each,
plus a `/simplify` pass (one field list in `resolve_based_on`).

**Done, round 5 (2026-10-03, `b761af97`..`3036c297`):** links in footnote and
endnote text kept and tagged (footnotes parsed with their part's
relationships) · warped WordArt reads its text: invisible text (rendering
mode 3) under the outlines in `Sect > P` · textbox lists → L/LI · every L
carries `/ListNumbering` (numFmt; bullets by glyph) · a text shadow's gray
copy is an artifact, so the word extracts once · list labels end with a space
glyph ("1.01SECTION" → "1.01 SECTION") · comment pane, SmartArt and WordArt
font fallbacks no longer depend on HashMap order (door_air_cooling changed
between runs). Text 95.3 → 96.6%, struct 94.8 → 95.0%, ua_fail unchanged,
+19 KB. Pixels: footnote hyperlinks now take the body's hyperlink underline
offset (0.08 em, was 0.12): czech_crisis, isla, uk_commercial move ≤0.4pt.
Element-level `/ActualText` is ignored by extraction and Poppler; prefer
content-level fixes (see `current_focus/A11Y.md`).

**Done, round 6 (2026-10-04, `40c351ff`..`cec49b91`):** measured over the 240
references now tagged (the re-exports added 67). Office Math is a Formula
with a spoken `/Alt` from Word's own rules (`docx/math_speech.rs`; pendulum
text 72 → 98%) · textboxes in table cells are `Sect > P` inside their TD/TH
(japanese_land text 86 → 100%) · header/footer links keep their annotations,
and every annotation drawn in an artifact gets a Link holding only its OBJR
(Word leaves them untagged) · OLE
objects take their alt from the VML shape, or stay artifacts without one,
and a decorative block picture's paragraph mark keeps its element ·
Symbol-font low bytes and Wingdings 3 triangles extract as Unicode · the
comment pane's label draws in the main font's bold face, so its brackets are
no longer `.notdef` (door_air_cooling, the one visual change), nor the ? of
a label without initials · a `/simplify` pass (one comment-pane font choice,
one Span/Formula slot in `LinkTagger`, shared math helpers; no score moved).
ua_deficit 2 → 0, ua_fail 510 → 504, text 96.34 → 96.51%, struct
95.83 → 95.84%, +1.2 KB. Most of the remaining text gap is Word: rows split
at page breaks become two TRs, continued footnotes two Ps, and some lists
lose their labels into LBody or entirely (see `current_focus/A11Y.md`).

**Done, round 7 (2026-10-04, `4f610d12`..`de1ebce7`):** two content losses
fixed: vertical table-cell text was untagged (japanese_interlibrary text
95 → 100%), and a continuous break to landscape put transition_to_work's
clause 108 above the top of its page (now a new page, as in Word) · pictures in table cells and textboxes are
Figures (Figure counts match Word's in 9 of 12 changed fixtures). struct
95.84 → 95.87%, text 96.51 → 96.53%, ua_deficit 0, ua_fail 504 → 510
(pictures without descr, as in Word).

**Done, round 8 (2026-10-09, `3878e02a`..`88472fcb`):** headings in
textboxes are H1–H6 (chiseldon's heading-order deficit); complex-script
letters are drawn in the cs font, Arial by default, instead of `.notdef`
(arabic_rice: 3 deficit rules → 0, its text reaches the tags); wrapNone
SmartArt floats at its anchor and its paragraph's caption is drawn
(learning_cultures). Over 340 fixtures: ua_deficit 0, ua_fail 706, all
source-limited.

Progress over the 173 tagged references: struct 0 → 94.8%, text 0 → 95.2%,
ua_deficit 1165 → 0 (every fixture fails no PDF/UA-1 rule Word passes);
LibreOffice's own tagged export scores 76% / 84% on the same yardstick. Over
all 221: ua_fail 459, 13 PDFs claim PDF/UA-1 and pass all 106 rules; what
remains is 5-1 (no claim, 208), 7.1-9 (198), 7.3-1 (29) and 7.4.2-1 (24),
all source-limited; every font is embedded. Output size ~24.0 MB (round 4
+37 KB: real fonts embedded where Type1 Helvetica was).

**Baselines accepted (round 4):** 15 fixtures' scores, 7 visual hashes. The
one drop is multi_font SSIM 42.6 → 39.2%: Copperplate Gothic Light now falls
back to Arial with real widths and fits on one line, where the
approximate-width Type1 Helvetica wrapped it like Word's real (wide) face
does; our page 1 also runs ~14pt taller than Word's, pushing its last line
(Bodoni MT) to page 2. Vendoring CopperplateGothic-Light in the assets repo is
the faithful fix. (irish_school's text drop is gone: `a11y_text` now folds
symbol glyphs, see SCORING.md.)

**How Word tags things (learned the hard way):**
- Pictures, charts and SmartArt: the paragraph's own (empty) P, then a
  `Figure` hoisted to Document level. Chart and SmartArt Figures carry **no
  extractable text** — their labels are artifacts; SmartArt's `/Alt` is the
  node texts one per line plus "(Layout Name)".
- Tables: `tblLook` firstRow → THead/TH (on when `tblLook` is absent),
  firstColumn → TH; a header-only table gets an empty TBody; a vertically
  merged cell's continuation is an empty cell (no RowSpan); no ColSpan, no
  Scope (so Word fails 7.2-42 and 7.5-1). Word also flattens some bordered
  data tables to one P per cell (who_prescribing 16×7, bush_fires 56×4,
  covid_insomnia) — the rule isn't recoverable from the DOCX features; we tag
  them as tables, which costs a few pp struct on those fixtures.
- Nested lists: the sub-list's L sits inside the parent item's LBody.
- Tabs and line breaks extract as spaces.
- TOC/TOCI only inside a real TOC field; hand-styled "toc N" paragraphs stay
  P. Each TOCI's Link is the `PAGEREF \h` page number.
- Footnotes: `P > Link > (OBJR, Span mark, Footnote > P)` — the Note sits
  inside a Link on the reference mark.

**Backlog, ordered by gap data (`tag_gaps.py` / `text_gaps.py` in the session
scratchpad; rebuild them from `tests/common/a11y.rs` if needed):**
1. Math speech reads matrices, accents and equation arrays as their contents
   in order. Anchored SmartArt with wrapping other than wrapNone is still
   laid out as if inline (no fixture has one).
2. (Decided, not a gap) slovak_eu_directive: Word has 14 TRs to our 9 because
   it starts a new TR per page a row runs onto; we keep one TR per row.
3. Link rects and outline destinations ignore the comment-pane zoom
   (`comments::page_zoom`) and vAlign (`assembly.rs`) — no fixture has links
   with either.
4. Wingdings 2 and Webdings still extract as private-use code points;
   `w:lang/@bidi` (complex-script text) ignored. (`w:softHyphen` dropped and
   `w:noBreakHyphen` → U+002D both match Word's extraction.)
5. Missing glyphs other than spaces still draw `.notdef`: Word rescues them per
   character from another font at their real width; widen the CJK rescue
   (`missing_cjk_chars`, `__cjk_fallback`) to non-CJK characters. The space
   fallback also lays U+2002/2003/2009/202F out at U+0020's width.
6. `anchor_seq` counts per parse pass: textboxes from paragraph-level
   `mc:Choice` (`collect_textboxes_from_paragraph`) sort after every run-level
   anchor. The XML position of the anchor node would give true order.
7. Test-run time: with Microsoft Defender scanning `tests/output` and a
   concurrent worktree run, the full suite took >60 min (normally ~6–10).

**Layout side findings (round 4 `/simplify`):** basedOn inheritance skips
`keep_next`, `keep_lines`, `contextual_spacing`, `page_break_before` and
`borders` (plain bool / default in `ParagraphStyle`, so "unset" can't be told
from false): a custom style based on Heading 1 loses keepNext. Also
`eastAsiaTheme="minorHAnsi"` (997 runs in the corpus) is ignored by
`resolve_east_asia_font`, and Times New Roman / Calibri / Cambria have no
metric-clone fallback (Liberation Serif, Carlito, Caladea) on Linux.

**Harness side findings (2026-10-01):** `tests/text_boundary.rs` has had no
`#[test]` since fb9373b, so the TxtBnd baselines are stale;
`engine_compare.py` `pdf_creator()` truncates Quartz producers at the escaped
paren.

## Biggest Scraped Gaps (2026-10-04, branch `gap-fixes`)

Rule evidence in the commits; open items in `current_focus/layout-accuracy.md` §5. Mean J 63.85 → 64.19.
- DONE: running heads (IF over nested STYLEREF, `\n`, forward search),
  cell-paragraph indents from the style, floating header tables, column
  breaks in one-column sections. bosch +35.0, french_sexual +22.2,
  turkish_prostate +19.5.
- DONE (radiographer round): nested-row split between lines, inherited
  header extent, table page-top gap, continued-row top border, per-cell
  margins in split rows. Mean J 64.19 → 64.77; radiographer 24.1 → 75.8.
- TODO: estonian's per-line drift; table cells through `build_paragraph`;
  one new-page helper; labels hanging outside a cell (probe Word first).
- PARKED: online glyph positioning (`online-references.md` §4a) — the
  correction phase is not derivable from exact widths; fit or more probing is
  the user's call.

## Synthetic-Case Gaps (DONE — 2026-10-04, branch `synthetic-gaps`)

Six rules from the lowest-scoring handcrafted `cases/` fixtures, each one
commit verified by a full suite run, no fixture regressed by more than 2.3 J.
`cases` mean 74.08 → **75.81** J; evidence in the commits.

- Character styles inherit along `basedOn` (case50 +24.8)
- A list marker's extra ascent adds to an auto-spaced line once, unscaled (case3, case33, 4 scraped)
- Before compat 15 a floating table's `tblpX` places its first cell's text (case40/45/46 +12–16)
- Tight wrap clears the polygon over the line's whole height; a word too wide for the gap
  beside a float goes whole to the other side (case42 +19.5)
- Glyphs inside a word are drawn with the pair kerning the word was measured with
  (TJ; case3 +13.5, russian_sports +15.5, czech_health +9.2, no regressions, +0.2% PDF size)

**Open:** case37's reference shows five black boxes where its shapes are (broken
export; re-export it). case69: Word steps the last line 15.00, we 14.83 (0.25pt grid
under `vAlign=center`).

## New-Case Accuracy Round (IN PROGRESS — 2026-10-03, branch `accuracy-oct3`)

10 scraped fixtures added with references exported by `tools/word_export.py`
(Word for Mac, unattended). Rules measured from those references or from Word
probe documents; each fix is one commit on the branch.

**Done** (suite deltas in each commit message):
- Moved text (`w:moveTo`/`w:moveFrom`), nested hyperlinks, mid-paragraph page breaks
- `w:position` raises/lowers runs and grows the line on that side only
- Super/subscript size from the face's OS/2 script size, rounded to 0.5pt
- Table style `tblCellMar` + paragraph spacing (along basedOn); cell grid snapping
  only under `adjustLineHeightInTable`; vMerge continuation cells don't number
- Row split with one line of room (14pt guard); nested tables split between rows
- Row splits charge a cell paragraph's space after; a carried-over paragraph keeps
  its space before (nabl +16 J). Merged cells draw borders row by row, closing at
  a page break. VML HRs sit 2pt above the line bottom
- Odd/even section breaks: filler page vs number bump, filler pages bare,
  per-variant header inheritance (§17.10.5) — 15 Word probes; croatian_thesis +50 J
- Autofit minimum width breaks CJK words after each ideograph; pre-2013 tables
  always outdent by the cell margin (6 fixtures +3–6 J)
- Grid + auto multiple: line = max(cells, m × pitch) (40 Word probes)
- PAGEREF prints its cached result; text boxes drop auto space-before on top (air_pollution +25 J)

**Parked:**
- ~~massachusetts: page-anchored body frames (`framePr vAnchor=page`) not implemented~~ done 2026-10-05
  (Annotation Fixes 2026-10-05, item 11)
- dutch_government: a page-anchored floating table moves the following body table
  down 2.4pt in Word (not to the float's bottom; cause unknown), and two 1pt
  `in-table` paragraphs come out 0.7pt short. Word reports "unreadable content" in
  the original fixture (valid zip; re-zipped variants open fine), so its reference
  came from a Word-repaired copy
- strategi: its original file ended in 4 stray bytes (`\r\n\r\n` after the zip's end
  record), which trips Word's repair prompt; the old reference showed the repaired layout.
  Input and reference replaced with the stripped file and its export (2026-10-03): 63.2 J
- radiographer: Word also splits *inside* a nested row (between its lines)
- Word floors auto-multiple grid lines to 0.24pt steps (19.44 vs our 19.50)
- Slash breaks: Word for Mac never breaks after `/` (probe: 138 margin crossings over 6 pair
  kinds incl. digits and a 40-char token, all wrapped whole), matching our rule. Older refs
  that do end lines on a slash (education_consultant "Partners/", romanian "septembrie/")
  presumably come from another Word build
- References of the first 10 new fixtures were staged as `<stem>_<hex>.docx`, so
  FILENAME fields print that name (massachusetts footer); fixed in the tool, refs not re-exported

## Deterministic Output (DONE — 2026-10-01, `6f64723a`)

All 226 fixtures (as of 2026-10-01) converted to identical bytes across runs
(three renders + `cmp`). **Regressed:** on 2026-10-03
scraped/door_air_cooling_unit_spec gave two different PDFs over three runs of
the same binary (J 79.85 vs 79.82); not investigated yet.
Three hash-order sources: `embed_truetype` fed `used_chars` (a `HashSet`) into the
glyph remapper; `collect_and_register_fonts` registered fonts seen outside runs
(SmartArt etc.) in `HashMap` order, shuffling F-names and font objects; and the
alpha ExtGStates were allocated and listed in `HashSet` order. Mattered beyond
reproducibility: the harness keeps a byte-identical `generated.pdf` (and its
screenshots, diffs, veraPDF results), which non-deterministic output defeated —
a `src/` touch with no output change took 2m25s, now 42s (warm 32s, cold 2m27s).
Corpus 0.7% smaller (sorted glyphs compress better); scores and conversion time
unchanged. Any new `HashMap`/`HashSet` iteration that reaches the PDF must sort.

## Note Marks in Table Cells (DONE — 2026-10-01, `a950737d`)

Footnote and endnote reference marks inside table cells were drawn empty
(erasmus_plus_staff_mobility_agreement "Seniority" vs Word's "Seniority²"):
cell layout never replaced the empty mark run with the note's number.
`RenderContext::with_note_marks` now does it for cells and body alike, and
the marks get their Link to the note. Visual output changed in the 6
fixtures with marks in cells (baselines accepted in `7879bc7f`).
Column auto-fit still measures cells without the marks (a mark's width).

## Large-corpus round (2026-10-06)

A second external corpus (~6,400 Word for Mac exports; 2,480 clean
documents scored) triaged by signal (page offsets, page drift, lost
pictures, fonts, lost text) and diagnosed per cause. Clean-corpus Jaccard
68.1 → 69.4, wrong page counts 170 → 149; fixtures unchanged except
polish_building +8.1. Details, evidence and the open queue:
`current_focus/layout-accuracy.md` §5 (open queue).

1. Legacy VML pictures (`w:pict` + `v:imagedata`) render, incl. watermarks.
2. Floats anchored in front of a page break are kept.
3. Strict OOXML namespaces are read.
4. CSS-style font lists in `w:rFonts` use the first name.
5. docDefaults space before / autospacing apply; missing docDefaults take
   Normal.dotm's.
6. A column-wide picture keeps its paragraph mark beside it.
7. contextualSpacing in headers and footers.
8. A next-page section starts on the page a page break just opened.
9. Keep-with-next chains end at pageBreakBefore; keep flags inherit via
   basedOn.

Next: eleven focus fixtures, one per open cause (`current_focus/layout-accuracy.md`
§4; in `tests/fixtures/scraped/` with baselines since 2026-10-10).

## Focus-fixture round (2026-10-07)

Ten rules from the local focus fixtures (evidence in the commits), each
A/B-tested on the corpus documents with its construct. Focus fixtures
19.4 → 44.6, scraped 63.4 → 63.7 (erasmus_plus +20.3, east_asia +29.8),
other groups unchanged. ukrainian −3.7 is fixed by item 11, estonian −7.8 by
item 12; uk_commercial_lease −10.8 is the next step
(`current_focus/layout-accuracy.md` §2.1); baselines not yet accepted.

1. hideMark: never hides the document's first mark; a hidden trailing mark
   drops its spacing; a picture in it keeps its height.
2. A line of nothing but tabs wraps at the margin.
3. `framePr yAlign="inline"` frames are in-flow paragraphs.
4. A picture-only line takes its leading from the paragraph mark.
5. Page/margin-placed header/footer floats cover their own band (85 → 2
   pages).
6. A header float over the first line pushes the header text below it.
7. Keep chains carry whole keepLines paragraphs; over-long chains start a
   page and flow.
8. A row whose first cell is keepNext stays with the next row.
9. A floating table pushes a paragraph whose first line reaches it.
10. Footnotes take the implicit tab stop at a hanging indent.
11. Justified compat-15 tab lines squeeze their spaces like untabbed ones
    (ukrainian 53.0 → 78.0, scraped 63.69 → 63.83; corpus subset 55.9 →
    66.0); uk_commercial_lease 38.3 → 27.1 from short footnote line pitch
    it uncovered.
12. Arial Narrow is Word's 2.42;O365 build (fonts/, assets, ~/Library/Fonts):
    estonian 42.5 → 69.8, indonesian 27.3 → 54.5, renewable_dispatch
    71.0 → 85.7.
13. EMF pictures draw text (fonts, Dx advances, alignment), standalone
    lines, pens, stock objects, PatBlt fills and rectangles, and map their
    rclFrame (not the ink bounds) to the picture box; a WMF with an embedded
    EMF draws the EMF. potamites 14.0 → 72.7; corpus EMF documents 49.3 →
    50.2. Clipping regions and opaque text backgrounds remain.
14. An empty anchor paragraph's line moves below a float it cannot sit
    beside (cyprus 32.3 → 79.7, 3 pages as Word; corpus subset 51.1 → 54.8).
15. hideMark hides the mark left alone after a cell's closing line break,
    with its space after (maine's extra blank line; 5 Word probes).

Still open from the focus set: Arabic layout (it has the cs font since
`40f87b52`, but needs the RTL/shaping item below), nested layout tables (maine).
## Wrapped running-head pictures (2026-10-08)

An image-only header/footer paragraph that wraps multiple inline pictures
retains its paragraph font's descent between lines and positions the visible
picture above its bottom effect extent. Case84 is a synthetic Word for Windows
reference with two wrapped pictures and a 0.75pt bottom extent. Probes at 8/10/14pt
Times New Roman and Arial, with 0/0.75/2pt extents, agree within 0.002pt; Calibri
retains a 0.14–0.26pt font-metric difference. The rule is deliberately limited to
wrapped image-only running heads; other picture paragraphs keep their existing
behaviour. The independent header-height estimate still needs per-line layout.
## Short running-head pictures (2026-10-08)

A picture shorter than its natural text line now uses that line's baseline in
headers and footers, including the bottom effect extent, as the body already
does. Both Paragraph.image and the one-picture Run.inline_image slot are
covered (an empty tab keeps the picture in runs). Synthetic case85 exercises
both slots. Fifty-four Word for Windows probes cover three fonts, three sizes,
three effect extents and both slots: Times New Roman/Arial agree within 0.002pt;
Calibri retains a font-metric difference of up to 0.26pt. On an external
letterhead, the stripe below its logo table is at 72.146pt vs Word's 72.150pt;
the 82-document corpus has no conversion/page-count or page-score regressions.
## Reflected group line connectors (2026-10-08)

Group flipH/flipV now compose through nested group transforms for linear
connectors, with positive bounding dimensions and reflected endpoint direction.
Zero-height horizontal lines retain their zero height. Other shape types retain
their existing transform: image/text/arc content mirroring is separate work.
Synthetic case86 compares two differently weighted horizontal lines in a scaled,
vertically reflected group against Word for Windows; the bounds differ by 0.02pt.
Two unit tests cover reflected zero-height lines and cancellation of nested flips.
Standalone header group lines also need the running-head connector render fix.
## Continuous section opening an empty sheet (2026-10-08)

A continuous section beginning at an otherwise empty page top now owns that
sheet's header/footer selection and uses its first-page variant. Mid-page
continuous sections keep the preceding section's running head as before.
Synthetic case87 agrees with Word for Windows on FIRST/DEFAULT/FIRST across
three pages; a mid-page control retains FIRST/DEFAULT. One external document's
last-page header is corrected without changing its three-page count; the
82-document corpus has no page-count or page-score regressions against main.
## Running-head connectors (2026-10-08)

Header/footer paragraphs now paint their DrawingML connectors, including
zero-height horizontal lines, at the paragraph anchor. The parser already
retained them, but the standalone running-head render path skipped them.
Synthetic case83 checks a horizontal and a diagonal line on two pages against
Word for Windows. On an external 82-document corpus, conversion and page counts
are unchanged; restored letterhead rules match Word's bounds exactly. No
external corpus documents are included in the repository.

## Bugs exposed by the external PRs (TODO — found 2026-10-09)

PRs #21, #29 and #33 were first declined for fixture regressions. A second
pass showed that each rule matches Word's PDF, and that main had passed those
fixtures through compensating errors. Fixed alongside them: float-zone top
(`55ed4910`), page-split floating tables (`b8ae47e5`), inherited first-line
indent under a numbering hanging indent (`20da20e3`), `lvlJc` (`43d4fc80`)
and `w:suff` (`a3681d62`). Still open; main hid each of these:

- **brazilian_logistics p10 (+1 page, J 62.5 → 46.3).** "Fonte: ALARCOM,
  (2019)." is laid out one letter per line from y≈108 (Word: one line at
  y≈346), so the paragraph is squeezed to almost no width. It was already
  wrong in main; until #21 put the lists exactly at Word's indents, main's
  35.6pt-too-narrow list indents absorbed the lost room. Since PR #36 (no
  extra mark line after wide non-OLE pictures) the page count matches Word
  again, but p10 still stacks the letters.
- **chinese_costume (+1 page, 2 → 3).** Word widens Latin–CJK boundaries
  ("5-10 分钟", "3 分钟", "MP4 封装"); we don't, so CJK lines in table cells
  carry more text than Word's and break differently. Main stayed at 2 pages
  only because its cell list text was drawn over the "（1）" labels.
- **covid_insomnia p5 right column (J −0.3).** The reference list sits one
  line low ("[10]" at 146.3pt, Word 135.3pt). #21 now wraps "[10] Kaplan…"
  exactly as Word does, which the offset turns into a score dip.
- **Page-split floating tables lose their wrap zone.** `b8ae47e5` keeps the
  one-line clearing rule when no zone survives. A full-width table that
  splits across pages would be treated as having a side strip; carry the
  table's horizontal extent past the page break if that case shows up.
- **Two thresholds for one rule.** `MIN_EMPTY_STRIP` (18.0pt, `pdf/mod.rs`,
  bracketed only between 0 and ~42pt) and PR #31's `MIN_NESTED_FLOAT_STRIP`
  (18.75pt, `pdf/table_layout.rs`, measured between 18.70 and 18.75pt) both
  decide whether an empty line fits beside a float. Probe the body case in
  Word and share one constant.
- **Nested table origin (+4.9pt).** In the PR #43/#45 fixtures
  (case204–205, case210–211) every nested cell's
  text sits 4.9pt right of Word's while the column boundaries match, so the
  nested table's left edge (cell margin outdent) is off.
- **case111 cell too narrow.** "…ора по" overflows its cell in our layout
  but fits in Word's; since PR #40 clips cell text at the cell edges the
  overflow shows as a cut-off word instead of overprinting the border.
- **PR #44 only before a continuous section.** The section-break line kept
  after a table is measured only when the next section is continuous (the
  fixtures' case). Before a new page it spilled onto a blank page
  (indigenous_innovation p15); whether Word keeps it when it fits there is
  unprobed and invisible.

## Layout accuracy round (2026-10-01)

Rules derived from Word reference PDFs (borders, text positions measured with
`mutool trace`/`stext`), one commit each. Fixture Jaccard over the round:
cases 69.9 → 74.3, scraped 39.9 → 57.7, new 38.5 → 55.9, samples 35.4 → 55.9,
hyphenation 61.2 → 63.2.

1. **compatibilityMode** is parsed (`docx::settings`). Compat 15 tables sit at
   margin + tblInd (no cell-margin outdent).
2. Header floats: a negative paragraph-relative offset counts toward the
   header's extent.
3. Table-cell list lines get the marker-ascent boost body lines already had.
4. **Border bands**: cell content sits between horizontal border bands (row =
   content + half of each band; table flow includes the outer halves). The old
   flat +0.5pt per row was Table Grid's border width.
5. Justification spreads slack over word spaces only (space gaps and text-less
   space chunks), not run joins inside a word.
6. **Justified squeeze (compat 15)**: a word stays on the line if its midpoint
   is inside the measure and the spaces can shrink to ≥ 75% (`SPACE_SQUEEZE`;
   94.6% of 14,122 measured Word line-end decisions; either condition alone
   ~90%). Older compat modes never squeeze.
7. **Per-line heights**: each body line is sized by its own runs (max top + max
   bottom, run-border pads included, math excluded); grid/exact paragraphs
   keep one box.
8. Compat 15 tables: the left border band starts at the indent (shift by half
   a band).
9. `w:kern w:val="0"` means kerning off (it overrides docDefaults).
10. Paragraph borders: the top band lies inside the paragraph like the bottom.
11. Lines span max ascent + max descent across their runs (case8: a pixel font
    with no descent beside Arial); sub/superscript offsets don't grow the box.
12. docDefaults without pPrDefault → Word's built-in 8pt after, line 278;
    an empty pPrDefault stays OOXML single/0.
13. HTML auto spacing (before/afterAutospacing) = 14pt, not at document start,
    not between items of one list (same numId), not at a cell's edges.
14. At-least trHeight bounds the row between its border bands (rows step
    trHeight + band); exact trHeight is the border-to-border pitch.
15. Negative pgMar top/bottom = absolute value, header/footer never push (§17.6.11).
16. A picture-only paragraph taller than its line gets its own mark's
    line-spacing leading (replaces the next paragraph's `after_image_boost`).
17. A body line of only spaces/breaks takes its height from a break run (it
    sizes the line it ends) or else the paragraph mark, whose font now
    inherits style → docDefaults and resolves theme fonts.
18. Leading spaces indent an inline picture as they indent a word.
19. `w:shd` `solid`/`pctNN` paint `w:color` over `w:fill` (auto black over
    auto white) for runs, paragraphs, styles and cells.
20. An empty paragraph's synthetic run takes the mark's font even when the
    mark sets no size (a Calibri mark in a Times style).
21. A tab never raises its line (12pt tabs among 11pt footer text leave the
    footer at the 11pt line); a line of only tabs still takes their size.
22. beforeAutospacing is dropped on a header/footer's first paragraph too.
23. `w:cr` is a text-wrapping break (§17.3.3.4).
24. **`w:linkStyles` without `w:attachedTemplate`**: Word reloads the styles
    from the stock Normal.dotm on open — its docDefaults (12pt, 8pt after,
    278 auto) and a Normal with no formatting of its own.
25. `doNotExpandShiftReturn`: justified lines ending in a manual break keep
    their natural width (no fixture changes; spec setting).
26. GIF and TIFF pictures are transcoded to PNG on load.
27. A word wider than the line breaks at the last character that fits.
28. URLs wrap only after a hyphen or at the margin (102 measured URL line
    ends: 2 after a `/`); the old break points after `/ ? # & = ;` are gone.
29. A float's reach is tested against the paragraph's first line (below the
    inter-paragraph gap), not its slot top.
30. A style's own `w:ind` beats the numbering it carries when the ind sits on
    the numPr style or below it (§17.7.2); an ancestor's ind still yields.
31. An empty paragraph's synthetic run inherits (so table-style sizes apply).
32. A `w:br` directly under `w:p` breaks the line (malformed, Word honors it).
33. **Keep with next** reserves the lines widow control keeps together: all of
    a ≤3-line paragraph, two of a longer one, one without widow control (the
    next paragraph is laid out to count).
34. **contextualSpacing** drops spacing only beside a same-style paragraph
    (§17.3.1.9), not whenever both neighbours carry the flag.
35. A table style's bold/italic never overrides the run's own `w:b`/`w:i`
    (direct formatting is absolute for toggle properties, §17.7.3).
36. **docGrid lines**: a grid-snapped line's glyph box (win ascent + descent)
    is centred in the cells it occupies; a Latin font keeps its line gap above
    the ascent, an East Asian font has none. 12pt TNR on an 18pt grid 13.56
    (Word 13.63), 16pt MS Gothic on two cells 23.75 (23.67), 16pt YaHei 24.37
    (24.53), 10.5pt Yu Mincho 12.69 (12.74). The baseline used to sit one
    pitch below the line top.
37. A word split over several runs wraps like a word inside one run: the run
    boundary breaks exactly where the two characters would break inside a
    run, and a word that continues across it moves to the next line whole
    (after the justified squeeze had its chance) instead of overflowing.
38. A word wider than the line breaks at the margin of the line it wraps to,
    not only when it starts an empty line.
39. Small caps are drawn at 80% (12pt text: 9.6pt capitals, one fixture with
    29 runs); the old 2pt reduction was never measured.
40. Footnotes use the same-style contextualSpacing rule (no fixture changes).
41. **A cell's first baseline** sits the font's ascent (line gap included)
    below the cell top, as in body text, not a full em (10pt TNR: 9.38).
    East Asian fonts keep the em (11pt MS Mincho: 11.0).
43. macOS system faces with only Mac Roman family names (Helvetica, Helvetica
    Neue, Optima, Apple Symbol…) are indexed in a lower tier, so the vendored
    Microsoft Symbol still wins; "Times"/"Courier" go to Times New Roman /
    Courier New first, as Word draws them.
42. **docGrid line-spacing multiples** scale the cells a line needs instead of
    being snapped: 1.5 lines of one 18pt cell is 27pt, text centred. case79
    (25 fonts and sizes on 18pt and 15.6pt grids) confirms 36 and 42: all 36
    lines within 0.4pt of Word.

Remaining gaps are mostly fonts we lack and small cumulative vertical drift
(≈1–2px) that Jaccard punishes.

Open findings (not done):
- docGrid: East Asian lines sit 0.17pt lower in Word than we put them on
  average (case79, 17 lines; Latin lines average 0.00). Multiples below 1 and
  at-least spacing on a grid are untested.
- An at-least trHeight row with top/bottom cell margins: romanian_quality's
  header row is 67.5pt in Word = trHeight 60.2 + both 3.6pt margins, though
  its content is shorter (we give it 60.2). One sample; the usual reading is
  that the height includes the margins.
- Some table rows run slightly short (slovak_eu_directive: ~0.03pt per row)
  and a vMerge/colspan test table gets a far too narrow column; fix 41 lost
  the offset that used to hide both.
- ~~Table cells don't use per-line heights yet~~ (done in `53764bee`,
  2026-10-02).
- A whitespace-only run between bold and italic runs (slovak_constitution)
  does not set the line's ascent in Word; whether it counts for anything is
  open (counting it as text made the line too low).
- covid_insomnia two-column flow regressed with the squeeze.
- Tracked changes: see "Tracked-Changes (Redline) Rendering" below.
- **Cell and footnote paragraphs bypass `paragraph::build_paragraph`
  (TODO)**: `docx/tables.rs` and `headers_footers::parse_notes_simple` build
  `Paragraph`s by hand, so cells never get `space_before_auto` /
  `space_after_auto` (rule 13's cell-edge drop never fires) and footnotes never
  get `style_id` / `contextual_spacing` (rule 40 is a no-op there; endnotes
  use the full builder). Route both through `build_paragraph` with options for
  the table-style and note defaults.
- Structural follow-ups found by a cleanup review: one per-line height model
  shared by body, cells and headers (cells have per-line heights since
  `53764bee`, headers not yet); one
  widow/orphan split helper for the body split, keep-with-next and cell splits
  (cells assume widow control on); run-boundary breaks taken from UAX #14 over
  the paragraph's joined text (a URL split across runs can still break after
  a `/` at the boundary); one table-left-x function for body, header and
  nested tables (nested tables miss the compat-15 rule).
- **fontTable altName order (done 2026-10-02, `f4ef46b8`)**: Word uses the
  altName only when the font is missing: Source Sans Pro (altName Corbel)
  and chinese_student_union's Calibri (altName DejaVu Sans) are drawn as
  requested. Requested name first now; localized names (바탕, ＭＳ 明朝)
  resolve to the same files either way. A macOS-only face still yields to an
  altName: eco_int's Helvetica (altName Arial) is Arial in Word's online
  export.
- Below compat 15 without `overrideTableStyleFontSizeAndJustification`, a
  table style's font size beats Normal's (unless it is 10pt); we always let
  Normal win.
- A nested header table (logo beside a title table) can sit 5–7pt low,
  pushing the body down.
- macOS-only faces are indexed below every other face (43). Helvetica's line
  height needs no special rule (Helvetica body lines already step as in Word);
  Courier → Courier New is Windows' substitute, unverified for Mac Word.

## Layout-accuracy merge: regressions traced, fix plan (2026-10-02)

Every fixture that scored lower after merging `layout-accuracy` than `main` did
before, traced to its commit: the CLI built at each of fixes 1–34, each
fixture scored with `page-metrics`, then vdiff before/after the culprit. The
merge itself caused none (each renders exactly as on the branch). Baselines
for all of it were accepted in `bc6d9ded`. Four real bugs below; the rest are
correct rules that exposed older errors, or metric artifacts. Handover with code
locations and the measuring method: `current_focus/merge-regressions.md`.

**Real bugs:**
1. **Character styles ignore `w:basedOn`** (`docx/styles.rs`, the
   `"character"` arm keeps only the style's own rPr). case50's MidChar
   (basedOn BaseChar: Georgia, bold) draws in Cambria, not Georgia Bold
   Italic; the narrower text wraps differently and per-line heights (fix 7)
   shorten page 2 by 1.35pt (J 58.3 → 54.4). Also used by
   carbon_farming_initiative_rule (5 runs) and traditional_skills_job_form (4).
2. **Missing fonts in LibreOffice documents.** german_mezzo's
   "Archivo;sans-serif" (fontTable altName "sans-serif", family auto) is
   **Cambria** in Word's PDF (the "D" of "Diese Frau…"); we try the altName,
   then the theme body font, and draw Arial. Its 15 empty lines are each
   ~0.26pt short, which fix 17 made visible (SSIM 78.2 → 52.2). With the font
   forced to Cambria the merged build lands within 0.07pt: J 69.0, SSIM 81.2.
   sample500kB's "Open Sans;Arial" (altName Arial) is **Segoe UI** in Word,
   Arial in ours. Word ignored the LibreOffice altName both times. The theme
   body font rule (#158 #195, 2026-09-18) cites german_mezzo → Arial, but that
   was read off the PDF's font list (its explicit Arial runs), not the Archivo
   glyphs; bosch and the Calibri-theme fixtures still support it. Cambria and
   Segoe UI (`fonts/CloudFonts`) are both available: this is a rule, not a
   missing file. **2026-10-02:** Word takes the run's "Open Sans;Arial"
   literally and finds no fontTable entry (that is "Open Sans"), so neither
   altName nor family applies; the face is Word's default for an unknown
   font. Fixture `fonts/missing_font_substitution` (42 rows, local and online
   exports) pinned the rule down: missing with no usable altName → Cambria
   (roman family or no entry) or Calibri (any other family); panose, pitch and
   the theme play no part; Helvetica → Arial online. Implemented in
   `ed5e495c` (on main); german_mezzo J 50.3 → 69.0. Open: the `;` rule, Open Sans →
   Segoe UI (`merge-regressions.md` item 5).
3. **East Asian leading stacks with another run's descent** (fix 11,
   `5e613494`). usep_handbook's checkbox lines (☐ in MS Gothic, text in
   Calibri) are 1.53pt taller each than Word's: we put Calibri's win descent
   (3.22) under MS Gothic's ascent, which already carries all of the 1.3× East
   Asian leading (13.91). Word's 21.6pt pitch is MS Gothic's full 15.6pt line
   + 6pt after. Page 5 runs +19pt by its end (SSIM 94.8 → 92.0).
4. **Superscript/subscript size and double spaces** (polish_municipal, fix 5
   exposed it). Word draws "8³⁰"'s digits at 2/3 size (4.06pt wide against
   the 8's 6.08); we use a flat 0.58 (`effective_font_size`,
   `pdf/layout.rs`). A census of the 30 fixtures with `w:vertAlign` runs
   gives 0.667 (51 spans), 0.636 = 7/11 (29), 0.65 = 6.5/10 (18), 0.632 =
   6/9.5 (7): two thirds of the size, rounded down to a half point. case10
   and environmental_law_clinic_china show 0.58/0.6 (check for an explicit
   `w:sz` before trusting them). The double space after it is 7.44pt in Word
   = 2 × (3.0 + 0.72 slack): every space character is a stretch point, where
   we add one slot for the gap (6.69pt).

**Correct rules that exposed older errors:**
- japanese_land_development_sign_form (fix 14): rows now step trHeight +
  band like Word's (27.50 vs 27.60, 26.60 vs 26.64, 27.15 vs 27.12…); the
  0.5pt shortfall per row used to cancel the table sitting 7.6pt low. That
  comes from the header block (10.5pt CJK lines pitch 17.67 in ours, 15.12 in
  Word: headers don't use per-line heights) and one content-driven row 5.2pt
  taller than Word's (520 twips: 35.64 vs 30.48). SSIM 46.7 → 26.8.
- air_pollution_permit_form (fix 13): every auto-spacing gap now matches Word
  (page-end drift −119 → +16pt), but a 13.9pt excess at the top of the
  checklist text box, unchanged by the fix, now pushes the rest of the page
  low (SSIM 30.4 → 23.8, J +2.3).
- dutch_government_budget_letter (fix 4): the old +0.5pt/row fudge hid part
  of a borderless one-row table being ~2pt short (Table Row Height Deficit).
  J 19.4 → 17.6.
- korean_japanese_conference_form, east_asia_conference_form (main's theme
  font fix + fix 4): rows now match Word within 0.015pt; main's 0.25pt excess
  per row was cancelling a 7pt shortfall above the table. Both still fit on
  one page where Word has two. SSIM −2.2 / −1.3, J +1.1 / +1.2.
  korean_japanese's reference is a macOS print-path PDF made with HY헤드라인M
  and New Gulim installed; east_asia's is a Word export with Malgun Gothic /
  Batang substitutes we have, so it is the one to measure.
- erasmus_plus_staff_mobility_agreement (#185, see the 2026-10-02 annotation
  fixes): correct endnote spacing makes the inline endnote block run past the
  bottom margin of page 3, where the body already sits 30pt low.
- case61 (fixes 4, 41): borders within 0.15pt of Word (J 30 → 51), but the
  text-to-border gap moved (SSIM 84.5 → 76.4). Its table 2 first column is
  9.6pt too wide (older).

**Metric artifacts** (positions as good or better per vdiff):
arizona_physical_education_standards (fix 41 moved cell text 0.53pt toward
Word, J −5.3; net J −0.8, SSIM +2.5), case67, case6,
turkish_ancient_religions_plan, vaccines_history_chapter, lenten_prayer_unity,
chinese_student_union_nomination_form, croatian_regulations_altchunk (page
count now matches), construction_bathroom_accessories_spec.

**Font files.** Census of every face Word drew per glyph in the 223
references against the PostScript names of all indexed font files: only
three fixtures use faces we lack. croatian_thesis_topic_approval_form is set
in Merriweather (Regular + Bold, 340 glyphs) and we fit it on one page where
Word needs two (J 23.1). **Done 2026-10-02:** Word's cloud font cache had
the static Merriweather 2.002, whose widths match the reference exactly; now
in `fonts/CloudFonts` and the assets repo (`59143d9`): 2 pages, J 34.1,
SSIM 94.9. multi_font (round 4, `96fbce99`) needs Copperplate Gothic Light
(81 glyphs; in Word's cloud catalog, not downloaded yet: pick it in Word's
font menu); its Bodoni MT is already in `fonts/CloudFonts`.
korean_japanese needs HY헤드라인M and New Gulim (55 glyphs; Hancom/Windows
fonts, low value). eco_int's 14 "Helvetica,Italic" glyphs are Arial Italic
under another name (online export; Windows maps Helvetica → Arial).

**Fix plan** (one commit each, full suite per fix, in this order):
1. Character styles inherit through `basedOn` (bug 1): resolve the chain
   after parsing, like `resolve_based_on`; unit test; check the 3 fixtures.
2. Super/subscript size = 2/3 rounded down to 0.5pt (bug 4): check the
   0.58/0.6 outliers and the raise/lower offset (0.35, `pdf/layout.rs`)
   against the same references first.
3. Each space character is a justification stretch point (bug 4): find the
   justified lines with consecutive spaces across fixtures and compare.
4. Mixed East Asian + Latin line height (bug 3): candidate line = max(Latin
   max-ascent + max-descent, each East Asian run's 1.3× box); verify on
   case79 and the CJK fixtures before changing `run_line_metrics`.
5. ~~Missing-font census (bug 2)~~ (done, `ed5e495c`: Cambria or Calibri by
   family class).
6. Per-line heights in headers (japanese_land; also the layout round's
   pending item 3).
7. Paginate inline endnotes (erasmus; the `ponytail:` note in
   `render_endnotes_inline`).
8. The conference forms' 7pt shortfall at the title → table transition,
   measured on east_asia_conference_form.
9. ~~Vendor Merriweather~~ (done 2026-10-02) and Copperplate Gothic Light.

## Annotation Fixes 2026-10-05 (one commit each, branch `annot/french-logo`)

Baseline: HEAD 4371430b. Each fix verified by a full visual run (no case below
its baseline at the end; polish_ministry dipped after 4 and recovered with 7),
each diff through `/simplify` before its commit. Rules were measured with Word
probes (`tools/word_export.py`), not fitted to the fixtures.

1. **#264 french_sexual logo at the foot of page 2** (`8ff40fad`): a
   regression from fa4db3fa. Its shorter header let a paragraph's look-ahead
   anchor the next paragraph's logo on page 1; keep-with-next then moved that
   paragraph to page 2, which still used the page-1 `pending_float_anchor`.
   The one-shot anchor now resets on page flush and column advance. J +6.0.
2. **#258 czech_wastewater checkboxes** (`dbabec2f`): legacy FORMCHECKBOX
   fields are drawn. The cell is one line of the run's font tall and wide at
   the box size (`w:size`, or the run size for `w:sizeAuto`), sitting on that
   line's descent, with the square stroked 0.75pt and 1pt / 1.5pt inside it;
   checked boxes get 0.5pt diagonals. A ballot-box placeholder run carries
   the field through line layout; table autofit and tab segments measure it
   via `word_width_for_run` too.
3. **#262 english_town_council white seams** (`d46ae4d3`): paragraph shading
   now reaches the side borders' inner edges (their 1.47/1.73pt shift had left
   a gap).
4. **#263 erasmus_plus "(if applicable)"** (`7addc6f8`): table-cell paragraphs
   now parse `w:contextualSpacing` and apply it between a cell's paragraphs.
   Probe: cells follow the body rule exactly. education_consultant J +11.2,
   estonian +7.9.
5. **Below single spacing** (`9c8dca04`): for auto m < 1 Word shrinks the
   whole line box, so the first baseline sits m × ascent below the line top.
   Body and cell paragraphs only; header/footer, textbox and footnote
   paragraphs keep the old rule (no probe, no fixture).
6. **#261 door_air comment pane** (`66f1025d`): the pane and balloons follow
   the page zoom (pane 9.1pt right of the text column and 13.3pt short of the
   279.7pt balloon area; balloons 22.8/4.3pt in, text 4.3pt in, all unzoomed).
   The balloon font (7pt) and line height (8.5pt) are still case63's zoomed
   values.
7. **At-least row heights** (`44cca044`): an at-least trHeight bounds the
   content between the cells' top and bottom margins, which sit on top of it
   like the border bands. traditional_skills J 16.0 → 48.3, estonian 29.4 →
   50.3, go_math 55.4 → 74.4, romanian +5.1, polish_ministry 29.8 → 37.6.
8. **Exact row heights** (`02387b89`): an exact trHeight grows by the bottom
   margin only. candidate_reference J 41.3 → 63.6. vAlign top/center/bottom
   in both row kinds checked against Word (within 0.15pt).
9. **#82 czech_municipal**, three causes:
   - `de28189b`: an empty paragraph stretches by its mark's `w:position`, like
     a positioned text line (Normal lowers 0.5pt: spacers 14.54, not 14.04).
     J 13.6 → 23.7.
   - `1956b952`, `0f33757e`: header lines stretch the same way, and a narrow
     wrapSquare header float no longer extends the header. Page 1 now within
     0.7pt; J → 26.1, SSIM 55 → 65.
   - `eae1ae1d`: even pages lay out around the even header (body extent takes
     the page index; physical parity). czech J → 34.6 (SSIM 89.7), samtale
     43.4 → 79.5, japanese_medical 37.1 → 56.3 with its page count matching.
10. **Row splits** (both gates in `render_table` / `find_cell_split`):
   - `3383e6b3`: an optional split needs only two lines of a long first cell
     paragraph in the room left (rows break between lines). bulgarian_road
     J 30.7 → 48.3, pages 7 → 6 matching; stem_partnerships +4.7.
   - `e7291b93`: an at-least trHeight row splits where trHeight + cell
     margins still fits (`RowLayout.split_min`); cantSplit and exact rows
     never split. croatian_grant J 30.2 → 43.1, pages 69 → 65 matching;
     education_consultant 44.6 → 53.9.
11. **Body frames** (massachusetts letterhead), rules from 8 Word probes:
   - Page/margin-anchored `framePr` with wrap notBeside/none lift out of the
     flow (`OpenFrame` in pdf/mod.rs): consecutive blocks with equal frame
     properties (hSpace included; a table joins via its cell paragraphs)
     render through the normal dispatch in the frame's own column, auto width
     = widest block; body lines then step below the frame's band. A style's
     framePr fills the attributes a paragraph's own leaves out; a missing
     vAnchor means the margin.
   - A page/margin-anchored square textbox with no room on its wrap side
     (massachusetts' secretary box: wrapText right at the right edge) blocks
     a band too. Together: J 8.3 → 77.2, SSIM 20.2 → 93.7, all 15 page
     starts matching; nothing else moved.
   - `framePr yAlign` (top/center/bottom, inside/outside) places the frame
     in its vAnchor area by its content height, and a `wrap=around` frame
     wholly beside the text area (left or right of it) lifts too, without a
     band. Frames never break across pages. Corpus 3ec631ca50 (Evonik press
     release, address block at the bottom of the right margin): 4 → 3 pages,
     J 16.7 → 72.6, SSIM 24.9 → 92.7; fixture suite pixel-identical.
12. **indonesian_school "Format 12" box** (`9b02ea49`, `be2b6d84`): an anchor
   paragraph's shading no longer covers the room it reserves below its lines
   for a top-and-bottom float, and `lnRef` outlines take the theme's
   `lnStyleLst` width (2pt there, not a fixed 1.5pt). J 26.2 → 27.3,
   japanese_land +0.2. The rest of indonesian's gap is a font file: Word
   online's Arial Narrow is 2.42 (hhea ascent 1888 = winAscent), ours and
   macOS's 2.38 (hhea 1916), so our 1.5-spaced 11pt lines step 19.0 where
   Word's step 18.75. Needs Arial Narrow 2.42 in `fonts/` and the assets
   repo (the user's call); `external_leading` then gives 18.75 unchanged.
13. **Tab stops past the margin** (`4b7c7646`), Word probes in compat 14/15:
   before compat 15 an explicit stop past the right margin keeps the tab and
   everything after it on the line (off the page if need be); from 15 a
   right/centre/decimal one clamps to the margin. A tab never parts from
   the word after it. polish_building J 25.2 → 47.5, usep_handbook +2.2.
   Open: compat-15 left stops past the margin — Word puts the tab on a line
   of its own and the text at the margin below (probe P1, P6, P7); we wrap
   to the far stop.
14. **Exact and at-least line baselines**: a Latin exact line's baseline
   sits 0.8 × its height down, whatever the font (probes: Calibri, TNR,
   Arial, Cambria, 8–24pt); a winning at-least minimum stays bottom-aligned
   at the descent (matched to 0.00pt). Table-cell paragraphs now use the
   same rule as the body (`layout::boxed_line_ascent`), as does a split
   paragraph's continuation. Open: footnote paragraphs still place exact
   and at-least lines at the plain ascent (`footnotes.rs`).
15. **Header/footer top borders** (`250a4b7c`): a top border's band (space +
   stroke) sits above the lines and counts in the header/footer height, as
   in the body. carbon_farming's footer was 1.15pt short and the body fitted
   an extra contents line per page: J 29.7 → 57.2 with 110 pages matching;
   western_australia 60.8 → 72.0, clean_energy 40.3 → 45.7.
16. **Right tab stops in the right indent** (`dbe127af`): words after a
   right/centre/decimal stop may run to that stop (the line's "reach"), so
   western_australia's index entries stop splitting: 58 → 57 pages matching.
   A 0.05pt wrap tolerance was tried and rejected (clean_energy wraps "of"
   exactly where Word does without it).
17. **Section-break paragraphs** (`4cc2bb5d`, `7ddff4b2`), 16 Word probes:
   an empty sectPr paragraph has no height, also before a continuous column
   change (the old covid 11.9pt exception was compensating a lost space
   after); a lone one opening a new-page section keeps its line; before a
   continuous section the paragraph before keeps its space after and the
   break's own after only absorbs the next space before
   (`section_break_spacing`, shared with the bookmark estimator).
   transition_to_work J 42.3 → 65.1 with 153 pages matching, covid +3.7.
18. **Picture lines and clearing breaks** (`cf0339f5`, `eac9b0e4`):
   - An image-only paragraph whose picture is no taller than the line takes
     the full line and sits the picture (effect extents included) on a
     baseline one natural line of the mark's font down (Word probe: 2.25–12pt
     pictures under Arial 12, single and 1.5 spacing, all within 0.01pt).
   - A paragraph opening with an empty clear="all" break right after a
     text-anchored floating table keeps that line beside the table.
     indigenous_innovation J 35.1 → 68.0, SSIM 45.6 → 84.8; nothing else
     moved.
   - Open: header/footer image paragraphs (`header_footer.rs`, ascent-based
     placement) and table-cell image paragraphs don't follow the short-picture
     rule yet.
19. **altChunk HTML as Word imports it** (`1f2d2605`): rules read from a Word
   resave of the fixture (Word converts the HTML to paragraphs and styles)
   and two probe documents (`word_saveas.py`-style resave plus PDF export).
   Comma-decimal CSS lengths are invalid; a missing or invalid margin is HTML
   auto spacing; a span's margins become the paragraph's spacing (auto still
   wins); single lines unless a valid line-height; a bigger paragraph
   font-size sizes only the mark; h1/h2 24/18pt bold; whitespace collapses
   across elements; td padding becomes cell margins over 0.75pt; HTML tables
   share `tables::settle_row_borders`. croatian_regulations_altchunk J 7.2 →
   45.9 (pages 1–6 line for line). The opening-auto-spacing rule now runs
   once after the body loop. Open: `ul`/`ol`/`li` are still dropped by the
   HTML converter (none in the fixture).
20. **Pre-2013 table outdent** (`90bef579`): before compat 15 a table edge
   sits the first cell's own left margin (tcMar) left of the indent, not
   the table default's. croatian → 51.4, candidate_reference 63.6 → 69.2.
   Open: nested tables use neither compat rule (`render_nested_table`);
   needs a compat 14/15 probe.

21. **Footnote separator** (`3dd1452e`), Word probes with paragraph borders
   on the separator and the note: the separator is footnotes.xml's separator
   paragraph laid out like a body paragraph (Normal without a pStyle), right
   on the first note, its space after and line-spacing extra below the text;
   the rule is that line's strikethrough (OS/2 strikeout on the 0.25pt grid,
   144pt). 23 fixtures up (case76 +5.8 J), none down.
22. **Footnote paragraphs** (`591cb5c4`): note paragraphs take their tab
   stops (`docx::resolve_tab_stops`, shared with body and cell paragraphs),
   lay out through `layout::build_lines`, size each line by its own runs (a
   larger mark raises only its line) and are measured with the mark filled
   in. uk_commercial's notes now sit within 0.4pt of Word's.
23. **Fit checks above footnotes** (`0280f618`; Word probes: a double-spaced paragraph
   with widow control off, the top margin stepping its 22nd line's lead from
   1.5pt short of the separator to 3pt into it, TNR and Calibri separators):
   a line's lead may hang past the margin but not into a footnote area
   (kept at +0.12pt, moved at +0.17/+0.62pt), now also in the
   whole-paragraph check; and a keep-with-next chain counts the footnotes
   its kept paragraphs bring. zimbabwe_gold J 29.5 → 59.5, environmental_law
   44.0 → 52.6 (page count matching), uk_commercial (with item 22) 39.1 →
   37.9. Open:
   - zimbabwe_gold p5 keeps a line whose lead reaches 1.12pt into the area
     (unexplained);
   - uk_commercial p17/18: Word continues a long note onto the next page
     (continuation separator); we never split notes;
   - notes still use the hand-built "simple" parser rather than
     `build_paragraph` (borders, shading, contextual spacing missing).

24. **Small caps** (`05b04d86`, Word probe at 8–20pt): small capitals are
   80% of the size to the nearest half point, and a word space beside a
   lowercase letter is small too. italian_project_proposal J 36.0 → 39.6.
   Open: a space at a run boundary only sees its own run.

Seen, not started (2026-10-05 x-drift survey of page 1, `word_x_diff`):
- case45: an autofit table with tblW 4000 and tcW 4680 per cell keeps Word's
  saved grid (2091/1909), we split it 100/100; its tblpXSpec="right" copy
  sits one cell margin (5.4pt) past the right margin in Word, at it in ours.
- case61: Word's autofit widens the 200-twip "narrow" grid column to fit
  its text (cols 122.5/77.6/77.8…), ours keeps closer to the grid.
- case6 and case52/53: whole columns shifted 7–11pt (autofit again).
- air_pollution: dot lines in a narrow cell fit 42 dots in Word, 41 in ours.

25. **East Asian line box from hhea** (`c7a7d86e`, online export probe:
   Yu Gothic / Yu Mincho / MS Mincho at 10.5 and 12pt, no grid): the 1.3×
   box is the hhea ascent + descent, not win — Yu Gothic steps 15.0/17.25
   (1.3 × win gave 17.6/20.1). Yu Mincho in `fonts/` and the assets repo is
   now Word's 1.92;O365 cloud build (hhea = win; Word for Mac's 1.85 carried
   Yu Gothic's hhea). japanese_land_development J 21.1 → 35.7.

Tried and reverted — **NBSP stretch before compat 15**: a probe (Arial/TNR
justified lines, compat 12/14/15, print and online presets) shows Word
widening NBSPs with the word spaces before compat 15 and keeping them at
their width in 15; croatian's "- \xa0preslika" agrees. Stretching every NBSP
like a space (diff kept in the session scratchpad; render side = TJ
adjustments after each NBSP, gaps counted in `spaces_before_each`) moved
croatian +1.2 J but russian_construction −7.4, czech_expert −5.2,
slovak_fuel −2.0, turkish_journal −1.4. In the online references an NBSP
after a one-letter word stays at its width ("k dopracování", "z celkového",
"V hlasování" in czech_expert p1), and the stretched ones get uneven shares
(1.0–2.0× a space's extra). Needs a probe of NBSPs after one-letter words
and runs of NBSPs before it can be a rule.

Open:
- Bands are checked at block start (a paragraph running into one part way
  keeps its lines), and live beside `FloatZone`; one mechanism with a
  full-width flag would be cleaner. Textbox `reserve` still uses "width ≥
  half the column" where images and the new band use the side-strip test;
  unify them with a suite run (other fixtures may move).
- Text-anchored body frames, and wrap-around ones reaching into the text
  (croatian_grant's two), still flow inline.
- Footnote separator: done (item 21). Left: endnotes.xml's separator (inline
  endnotes still draw the fixed 0.5pt rule 12pt below the body; needs a
  probe of how its space before meets the last paragraph's space after), and
  a separator run's own rPr. zimbabwe_gold still fits one more two-line
  double-spaced paragraph + footnote on page 3 than Word: its last line's box
  (double-spacing extra below the text) ends 0.6pt inside the note area.
- `content_h` in `render_paragraph_block` includes the room reserved for a
  topAndBottom float below the lines; shading now uses the lines only
  (indonesian's "Format 12" box), but the paragraph's borders still use the
  reserved height. Unverified against Word; split it into a lines height plus
  a reserved-below gap once a probe shows where Word draws that border.
- The must-split path (row taller than a page) still tests only
  `!cant_split`, so an exact row taller than a page splits; Word likely clips
  it. Unprobed, no fixture.
- Header/footer paragraphs are measured twice (`compute_header_height`, no
  line building, and `render_header_footer`); `layout::position_stretch` is a
  one-line estimate. A shared per-paragraph measure would retire the mirrors.
- #259/#260 door_air are tracked-change balloons ("Deleted: …" for each
  w:del, change bars, insertion markup); needs a revision-markup feature.
  door_air is the only fixture whose reference shows markup.
- Cell paragraphs (docx/tables.rs) still default most pPr fields that
  `build_paragraph` resolves (keep_next, widow_control, borders, shading,
  outline level, …). No fixture is affected today (czech_municipal's cell
  borders are all `nil`); a shared pPr resolver would close it.

## Annotation Fixes 2026-10-02 (one commit each, worktree `annot/wp-n`)

Baseline for the round: HEAD 48a4eba0 (the layout-accuracy merge). Each fix
verified by a full visual run against the previous fix's run; each diff then
went through `/simplify` before its commit. #243, #248 and #250 no longer
reproduced at the baseline and were marked fixed without a code change.

1. **#244 family_kinship dashed rules** (`e1bb8557`): UAX #14 never breaks
   before a hyphen (LB21), so an 87-dash rule typed as one run was an
   unbreakable word; Word wraps such a run after however many dashes fit, so
   the rules form two columns. Word's two pair-wise departures from UAX #14
   (this and the ellipsis dot-leader rule) now live in one `word_pair_rule`
   used by both the whole-text split and the run-boundary check. Only
   family_kinship moved (J +1.9, SSIM +2.5); six fixtures with `--` in prose
   changed hash by sub-point x drift (a line end 397.80 → 397.85pt).
   Residual: Word fits 47 dashes on the first line by compressing the 33
   spaces to 2.97pt (the #93 justified-space shrink), we fit 46.
2. **#242 japanese_land shaded bar**: `w:trPr/w:gridBefore` was ignored, so a
   row starting at grid column 2 was laid out from column 1 (the hatched
   `gridSpan=9` cell covered columns 1–9 instead of 2–10). `TableRow` carries
   `grid_before` and a `grid_cells()` iterator yields each cell with its grid
   column and span; the thirteen hand-rolled column walks in the parser,
   layout and renderer use it. The PDF/UA tag span of a late-starting row's
   first cell absorbs the skipped columns (as `gridAfter` already did for the
   last). japanese_land J 8.7 → 8.9, go_math (the only other gridBefore
   fixture) J 44.1 → 44.3, SSIM 65.9 → 66.3. `w:wBefore` stays unread: the
   tblGrid already defines the skipped width.
3. **#249 french_sexual_health arrow list beside a logo**: Word measures a
   paragraph's indents from a float's wrap edge as it does from the margin
   (marker at edge + 17.85, text at edge + 35.7, matched to 0.02pt). The
   both-sides per-line geometry started the right region at the edge with
   no indent and no first-line/hanging shift, so the text sat under its own
   markers. One `right_of_float` expression now serves the both-sides and
   single-side branches (the latter's width had ignored `indent_left`), the
   paragraph-level `narrow_paragraph` box uses the same width, and
   `right_region_for` applies the first line's shift like the left region.
   Only the fixture itself changed (J 22.2 → 22.2, SSIM 29.3 → 29.7); the
   remaining difference there is vertical (our lines sit 10pt lower). #245
   (italian_evaluation line fit) no longer reproduced and was marked fixed.
4. **#247 ut_koer header staff**: the header renderer drew paragraph borders
   at one text line height before laying the paragraph out, so a staff
   picture's bottom border cut through the staff. Borders are now drawn by
   each exit once the paragraph's height is known (`lines_height`, so
   multi-line bordered header paragraphs are right too), the stroke sits
   `space` below the box as in the body path, and the band (`space` +
   width) is part of the paragraph's advance and of the header/footer
   height estimate, which also gained the picture line's descent. Measured:
   ut_koer rule 46.95 → 76.95 (Word 77.78; the rest is the picture's
   `effectExtent`, unread); bush_fires footer rule 652.05 vs Word 652.20 with
   its text unmoved (its bottom-anchored footer hid the band before);
   carbon_farming header rule 98.65 → 100.40 (Word 100.53). ut_koer J +20.9
   / SSIM +8.2, bush_fires J +3.7 / SSIM +4.5, carbon_farming J +3.0, covid
   J +0.7; western_australia and croatian_grant hash only.
5. **#185 erasmus_plus endnotes**: consecutive footnotes/endnotes were packed;
   Word puts a note's last paragraph's space-after before the next note
   (erasmus endnotes 5pt from `after=100`, master_thesis footnotes 3pt from
   `after=60`, both now matched to 0.1pt per boundary). `compute_footnote_height`
   charges the note its trailing space and `render_notes_downward` advances
   past it. erasmus scores J −1.3 / SSIM −3.5 anyway: its body already sat
   30pt below Word's by page 3 (cumulative drift), so the endnote block, 30pt
   taller now, runs past the bottom margin (inline endnotes are not paginated,
   see the `ponytail:` note in `render_endnotes_inline`). master_thesis and
   croatian_grant changed hash only.
6. **Hyphen before a digit** (from #246's first divergence, "Sindh 2019-|2024"):
   UAX #14's LB25 keeps "2019-2024" whole; a scan of the 223 reference PDFs
   found 13 line ends of the form "<alnum>-" | "<digit>…" ("1(4): 108-",
   "about 3-", "11-22-", DOIs), so Word breaks after a hyphen before a digit
   as before a letter. `word_pair_rule`'s hyphen arm now covers it. Seven
   fixtures improved, none dropped: polish_council J +12.3 / SSIM +12.9,
   international_te +6.9 / +8.9, physical_therapy +6.2 / +5.9, dental_amalgam
   +1.9, hyphenation/italian +0.9, covid +0.5, education_consultant +0.3; 19
   more changed hash only. #246 itself remains: its table pagination differs
   through cumulative line-height drift on page 2.
7. **#186 alfies_arc page 1 too high**: a 1014pt-wide OLE logo strip
   (`w:object`, a paragraph-level picture) in a 468pt column leaves no room on
   its line for the paragraph mark, which Word wraps onto an 11pt line of its
   own; Word's `" "` lines for the empty paragraphs (y 194.6 / 216.0 / 249.4)
   gave the object line box away. The block-picture height adds a line when
   the picture is wider than the *column* (learning_cultures keeps the mark
   beside a column-wide picture in a right-indented paragraph, so the
   paragraph's indents do not count; ukrainian_municipal's 35pt object adds
   nothing). alfies J +3.1 / SSIM +10.0 (page 1 now 3.2pt high, likely the
   picture's descent); brazilian_logistics J +21.1 / SSIM +20.4 (its wide
   figure had the same missing line). No other fixture changed. That also
   closed #59 (brazilian p9 "white space above image": the page matches
   Word now), and #124 (english_town_council page 3 start) no longer
   reproduced after the round: page 3 starts within 0.01pt of Word, the
   fixture's residual drift is at its page 6/7 boundary (page 7 starts 33pt
   high).

Findings left for later:
- A hyphen at the start of a word (" -5") now also breaks before the digit;
  the 13 measured cases all had an alphanumeric before the hyphen. Scan the
  references for `\s-$` line ends before widening or narrowing.
- `wp:effectExtent` is not part of an inline picture's line height; ut_koer's
  0.75pt bottom extent is the residual on its header rule.
- Two adjacent `w:noBreakHyphen` (parsed to plain `-`, `docx/runs.rs`) now
  become a break opportunity; map them to U+2011 if a fixture ever shows it.
- The tblGrid-less column inference (`docx/tables.rs`, `row_widths`) does not
  add `gridBefore` columns; Word always writes a tblGrid, so untested.

## Annotation Fixes 2026-09-18 (5 fixes, one commit each)

Baseline for the round: HEAD 2c0706d, 170 tests passing. Every fix verified by
a full-suite run against the previous fix's run (`touch src/lib.rs` before each
run — a stash pop mid-run bumps mtimes and the harness silently reuses PDFs).

1. **#241 cell float z-order** (japanese_land_development): a picture anchored
   in a table-cell paragraph was always drawn before the paragraph's
   connectors/textboxes; `CellFloatingImageLayout` now carries `z_index` and
   pictures above the shapes draw after them (`draw_cell_float`). No score
   change (small region), hash change only there. Also confirmed #220 already
   fixed by the row-split work (Q3 on p6 in both).
2. **#66 list marker line height** (case33): Word sizes a list line as *marker
   ascent + text descent* (× spacing), never the marker's descent, and drops
   the first baseline by the marker's extra ascent. 11pt Symbol on 11pt Calibri
   → 16.115 (Word 16.00) instead of our 15.50. Courier New "o" sub-bullets
   (streamnet p5) and Symbol on Arial (dialysis) confirmed the marker descent
   is ignored. 27 fixtures improve (samtale SSIM +36pp, case3 J +35pp,
   usep_handbook J +31pp, case33 +21pp, romanian_quality +21pp, czech_crisis
   +9pp); streamnet SSIM −3.8pp while its Jaccard gains 13pp — an agenda
   *table* on p1 runs +0.43/+0.68pt per row (row-height drift, was cancelled
   by the wrong list pitch). Residual: Word's 0.25pt line grid (16.115 vs 16.0).
3. **#158 #195 bosch**: (a) Word substitutes the *theme body font* for a
   missing `w:family="auto"` font — bosch (theme Calibri) → Calibri,
   german_mezzo "Archivo" (theme Arial) → Arial, three more fixtures agree;
   a plain "always Calibri" broke german_mezzo (1 → 2 pages). (Superseded by
   `ed5e495c`: the measured rule is Cambria or Calibri by family class, the
   theme plays no part, and `theme_body_font` is gone.) (b) A header `framePr` with `w:wrap="notBeside"` and `w:h` pushes each
   in-flow header *line* that would overlap its band below the frame (per
   line: bosch's first two 14.75pt lines fit above the 33–139pt band, the
   third lands at 139 → body at 153.75 like Word). bosch SSIM +21pp,
   croatian_thesis SSIM +9pp.
4. **#240 case41 p6**: the #152 look-ahead only fired for image-only anchor
   paragraphs; it now accepts text-carrying ones, the anchor paragraph's own
   zone peeks the handed-forward anchor top, and the look-ahead only fires
   when the float leaves ≥48pt beside it for text (case41 p3 wraps with
   64.8pt free; brazilian p9's figure with 37.5pt beside it went 20 → 21
   pages without that gate — Word leaves its caption alone).
5. **#237 stem_partnerships p4**: table rows split between *lines* of a cell
   paragraph (2 lines kept on each side), not only between paragraphs:
   `find_cell_split` returns a `CellCursor { item, line }`; `cursor_chunks` /
   `item_chunk_height` / `chunk_space_before` share the chunk arithmetic with
   `render_partial_row` / `render_partial_cell_content`; the split gate also
   accepts a single paragraph of ≥4 lines. stem_partnerships 9 → 7 pages (= Word).

Findings left for later:
- **#93 justified space compression**: Word pulled "2251," onto the line by
  shrinking the 10 breakable spaces to 2.00pt (natural 2.75 @ 11pt TNR, i.e.
  ~73%); no `wpJustification` compat flag. We only expand (`layout.rs`
  `overflows` / `extra_per_gap.max(0.0)`). Corpus-wide effect on justified
  text; the shrink limit needs calibration before touching it.
- **#233** needs Merriweather in the assets repo (underscore 0.835em vs Arial
  0.556em); nothing to do in code. Vendored 2026-10-02 (assets `59143d9`).
- streamnet p1 agenda table rows +0.43/+0.68pt each (table cell line height).
- dental_amalgam gained +22pp J with a max-descent rule but +10pp with the
  measured ascent-only rule — worth a look at what its markers are.

## Distributed Alignment (DONE — 2026-07-27)

`w:jc="distribute"` used to fall through `parse_alignment`'s
`_ => Left` arm, so distributed paragraphs rendered left-aligned. Added
`Alignment::Distribute`: it stretches *every* line including the last, and
spreads the slack between characters via `Tc` ("Distribute All Characters
Equally", §17.18.44) rather than between word gaps. The Tc divisor differs from
CJK justify by one gap: CJK justify keeps the grid's trailing cell gap (divide
by the char count), while `distribute` ends flush at both margins whatever the
script — Japanese 均等割り付け behaves the same way — so it divides by one gap
fewer (`char_justify_gaps` in `pdf/layout.rs`, unit-tested). Getting this wrong
left case77's CJK line 31pt short of the right margin.

Kashida variants (`mediumKashida`/`highKashida`/`lowKashida`) now map to plain
Justify instead of Left; true glyph elongation needs Arabic shaping we don't have.

Zero corpus fixtures exercise any of these values (verified across all 129 DOCX,
all XML parts), so corpus scores are flat — this is correctness-only.

### What case77's reference settled (2026-07-28)

case77 now has a Word reference. It confirmed three assumptions and killed one:

- distribute stretches every line including the last — our last line lands
  within 0.3pt of Word's.
- CJK distribute ends flush at both margins — validates the one-gap-fewer
  divisor; all 10 glyphs within 0.4pt of Word.
- `mediumKashida` on Latin text renders as ordinary justify.
- **`thaiDistribute` is NOT distribute.** Word leaves a Latin thaiDistribute
  line at its natural width while stretching a `distribute` line of the same
  shape, so it now maps to Justify. Whether Thai script triggers real
  distribution is untested — no Thai fixture exists.

case77 scores J 51.5% / SSIM 70.5% / TxtBnd 84.2%. The remaining gap is two
things, neither about distribute's core geometry:

1. **Distributed Latin lines with spaces** spread across one gap too few (Word
   counts spaces as distributable characters; we don't). Interior letters up to
   34pt off, margins still flush. See the `ponytail:` note on
   `char_justify_gaps` for why this wasn't chased.
2. **Kashida changes Word's line breaking.** With identical text, Word's
   `mediumKashida` paragraph breaks earlier than its `both` paragraph
   ("…industrious" vs "…industrious beaver"), displacing a whole line. We
   reproduce `both` breaking. Needs real kashida metrics, i.e. Arabic shaping.

Also note case77's own on-page label for sample 4 claims thaiDistribute gets the
"same treatment as distribute" — baked into input.docx before the reference
disproved it. Correcting it means regenerating the DOCX and the reference.

## New-Case Triage 2026-07-03 (10 fixtures added; fixes applied 2026-07-04)

Passing: streamnet_steering (J 54%), zimbabwe_broadcasting (J 53%). Fix round results (zero regressions across 218 cases):

1. **japanese_medical (J 2.8% → 4.3%, TB 5% → 46%)** — FIXED (a) CJK numbering formats `decimalEnclosedCircle` ①②③, `decimalFullWidth`, `aiueoFullWidth` in `format_number()`; (b) docGrid cell counting now uses sTypo metrics (`grid_snapped_line_h` in `pdf/layout.rs`) — win+lineGap (Yu Mincho 1.787) overshot the 18pt pitch and doubled every grid line (4 pages → 3, page 1 now matches ref). Remaining: sub-line drift through table rows (tables don't grid-snap; cell line heights slightly exceed Word's), one line still spills page 1 → 2.
2. **croatian_thesis (J 18.8% flat, SSIM 50.8% → 41.8%)** — FIXED the font bug: fontTable altName `SignPainter-HouseScript` (Word-for-Mac artifact for fonts missing on the authoring machine) is now rejected; falls to family fallback (Arial). Reference embeds real Merriweather — install it (Google font) or wait for Bundled Fallback Fonts for the rest. SSIM dip = more ink slightly misaligned; structure now correct. Ref page 2 is blank (trailing paragraph) — we emit 1 page.
3. **indigenous_innovation (J 12.2%, TB 0.7% — flat)** — FIXED indent precedence: numbering-level `w:ind` now outranks style ind when only some attrs are direct (§17.9.27, `docx/paragraph.rs`); recital numbers/text now align with ref. Score flat: 20 pages of justified-Arial wrap drift dominates.
4. **french_sexual_health (TB 75% → 82.7%)** — FIXED zone clobbering: a wrap float entirely outside the text column (margin QR code) no longer replaces an active in-column float zone (`pdf/mod.rs`). Body now wraps beside the top-right image. Remaining: ~2-line vertical offset (drift class).
5. **stiavnicke_bane (J 15.8% → 76.3%, SSIM 93%, TB 100%)** — FIXED: leading spaces on a line whose left region is blocked by a float now carry into the right region as an indent instead of triggering blank-line + x=0 (`pdf/layout.rs` build_lines).
6. **ut_koer (J 13.7% — open)** — tab-stop two-column contact header interleaved with a wrapTight header image: needs per-line segment layout with tab stops spanning the float's excluded band (`build_tabbed_line` has no float-zone awareness). Right column currently gets left-column content.
7. **physical_therapy (J 12.6%, TB 91.7% — open)** — no missing feature; constant small dx/dy glyph offset. Vertical Drift Investigation class.
8. **candidate_reference (J 21.3% — open, passing)** — minor table row-height drift (Table Row Height Deficit class).

## Hyphenation (PARKED — Word's online converter doesn't hyphenate)

`w:autoHyphenation` and `w:suppressAutoHyphens` are not parsed (the unused parse was removed in `3ead74c4`); `w:lang` is parsed on runs. However, **Word's online PDF converter does not perform syllable-break hyphenation** for any language, even when `autoHyphenation` is set. Every line-ending hyphen in reference PDFs comes from pre-existing hyphens in compound words (verified across 8 language-specific fixtures and all scraped fixtures).

The `hyphenation` crate (Knuth-Liang algorithm) was tested with 8 languages but its dictionaries don't match Word's — enabling it caused 40-50pp regressions across all languages because line breaks diverged. Removed for now.

**Prerequisites to revisit:** Reference PDFs generated by desktop Word (not the online converter) with proofing tools installed. The online converter ignores `autoHyphenation` entirely.

Test fixtures in `tests/fixtures/hyphenation/` (8 languages with Wikipedia text, `autoHyphenation` enabled) are ready for comparison when desktop-Word references become available.

## Image Cropping `a:srcRect` (DONE — 2026-09-04)

Cropped pictures used to render the full source squeezed into the frame. Now
`parse_src_rect` (`docx/images.rs`) reads l/t/r/b as 1/100000 fractions onto
`EmbeddedImage.src_rect` (None for absent, empty, all-zero or nothing-visible crops;
negative outward crops kept). It is applied through `apply_pic_props`, the one shared
step for outline, effects, clip geometry and crop that replaced three copy-pasted blocks
at the inline, anchored and paragraph-image parse sites. `embed_single_image`
(`pdf/images.rs`) wraps a cropped image in a Form XObject with BBox `[0 0 1 1]` drawing
the source through `crop_matrix` = `[1/(1-l-r), 0, 0, 1/(1-t-b), -l·sx, -b·sy]`; the inner
image lives only in the form's resources, not on the pages. The bbox clips, so negative
crops become blank padding for free and none of the eight draw sites changed; EMF forms
nest as-is. Downscaling and soft-edge radii see the effective source extent
(frame / visible fraction), so a heavily cropped photo keeps its resolution.

Verification: 9 unit tests (parser on the four real corpus elements plus edge cases,
matrix on identity/quarter/half/negative). Full suite: 220 fixtures, hashes changed only
on brazilian_logistics_study pages 9–10 and italian_evaluation_minutes page 7 — exactly
the pages holding the 4 real crops in the corpus (11 other fixtures have an empty
`<a:srcRect/>`, unaffected). All three pages match the reference crop visually. Scores
moved within noise (brazilian J 17.29 → 17.17, italian SSIM 43.83 → 43.82): the images
sit at drifted y positions, so the old squeezed image overlapped reference ink by accident.
`cases/case78` (generate.py + input.docx, grid PNG cropped four ways incl. negative) has
its Word reference and scores J 99.6% / SSIM 99.4%. It settled the open question: Word
renders a negative `a:srcRect` as blank padding inside the frame, exactly what the
unit-bbox form gives us for free.

Known ceilings (`ponytail:` note in `embed_single_image`): soft-edge and reflection masks
are still built on the uncropped source; SmartArt pictures use their own draw path. Also
seen while verifying: italian page 7's third signature is a bitmap EMF (`image3.emf`) that
`docx/emf.rs` rendered as nothing — fixed 2026-09-15 (annotation #228): `emf_to_raster`
wraps a lone EMR_STRETCHDIBITS DIB as a BMP, and inline pictures honour `a:xfrm@rot`
(quarter turns swap the layout box, `EmbeddedImage::layout_size`). See `minipdf.md` §1.2 and
§4.1 for the MiniPdf comparison and their corpus scan.

## Engine Comparison Findings (2026-09-04, `tools/engine_compare.py`)

CI (2026-10-04): the site is built by 4 shard jobs (`--shard i/4`, own cache
each) and a `publish` job that joins them (`--merge`); a cold shard takes ~30 min
where one runner took ~110 min. Competitor versions left the cache key (each PDF
is stamped with its engine's version instead), so a release no longer cold-starts
everything. Scoring is 3/4 of the thread time, mostly one veraPDF JVM start per
competitor PDF (~1,400 per cold run, 0.7–1 s each on a laptop, more on the runner);
if runs grow again, batch those into one `verapdf` call per shard (its JSON report
carries one job per file) via a `page-metrics` pre-pass that writes the `.a11y.json` caches.

Across 207 fixtures vs the Word reference we lead LibreOffice on mean Jaccard
(48.2 vs 39.5) and roughly tie on SSIM and text-boundary; MiniPdf is far behind
on all three. But LibreOffice beats us by 45+ points averaged over J/SSIM/TB on
a cluster of cases, i.e. we get the *structure* wrong, not just glyph placement:

| Case | ours J/S/TB | LibreOffice J/S/TB |
|---|---|---|
| cases/case13 (205 pp) | 8/21/1 | 46/90/100 |
| scraped/brazilian_logistics_study | 17/30/16 | 54/81/93 |
| scraped/russian_sports_ranking_decree | 10/20/42 | 43/89/100 |
| scraped/candidate_reference_check_form | 21/48/23 | 81/96/62 |
| scraped/family_kinship_lesson_plan | 23/39/32 | 58/87/94 |
| scraped/slovak_pedagogical_practice_agreement | 14/38/14 | 24/86/100 |
| scraped/school_meal_assistance_faq | 6/17/35 | 19/83/100 |

TB near 0 with LO at 100 means our line breaks or pagination diverge from page
one onward. (Snapshot of 2026-09-04. As of 2026-10-03 ours J/SSIM: brazilian
62/78, russian_sports 65/79, candidate_reference 30/52, family_kinship 91/97,
slovak_pedagogical 35/87, school_meal 71/91; case13 is in SKIPLIST.) Open
`comparison/index.html`, pick the case, and use the overlay to see where the
flow first departs.

**Page-count term in the compact report (DONE — 2026-09-05).** `run-tests.sh` now ends
its summary with `N/M page counts match`; per-case `ref_pages`/`gen_pages` sit in
`tests/output/latest_scores.json`. The visual test scores only the common pages, so a
generated PDF short of pages used to look fine. Not baselined (see `minipdf.md` §3.2).
Found while doing it: `tests/text_boundary.rs` has had no `#[test]` since fb9373b
(2026-09-03), so the TxtBnd values in `tests/baselines.json` are frozen; whether that
was intended is unverified.

**Other Rust converters checked (2026-10-02)**, same scorer, 220 fixtures
(cases/scraped/samples), ours 62.4 mean Jaccard on the same run:
- `dxpdf` 0.8.1 (Skia, MIT, ~110k lines, spec-first with Microsoft's
  [MS-OI29500] implementer notes cited at each Word deviation): 19.8 mean /
  10.6 median, 8 strict-parse failures on real-world files (duplicate XML
  children, `10923f` read as a float), 44 page-count mismatches, ~0.9 s per
  document (ours ~0.04 s). No fonts flag: it reads the host font system, so it was scored
  with a copy of `fonts/` in `~/Library/Fonts`. **That folder also changes our
  own output** (user fonts outrank the vendored ones; 80 of 220 fixtures moved,
  some 88 → 5) — never leave it in place, and score ours without it.
- `libreoffice-pure` 0.5.8: a document-generation toolkit, not a layout engine
  (≈6k lines for DOCX import + layout, standard-14 Helvetica only): 2.3 mean.
- `libreoffice_convert_rust` 0.1.0: a wrapper that shells out to `soffice`;
  its output is LibreOffice's.

**The published site understated every Times New Roman document (fixed
2026-10-03).** Site vs MacBook, same binary: identical on 109 of 208 cases,
55 cases 20+ points lower on the site, 52 of them set in Times New Roman
(case48 65.7 → 9.7, slovak_misdemeanor 88.3 → 4.6). The references were made
with Apple's Times New Roman 5.01 (hhea lineGap 87, 13.80pt per 12pt line);
the Ubuntu runner only had Word's bundled 7.00 (lineGap 0, 13.29pt), so every
line sat half a point too tight. Reproduced in a Linux container to the
decimal, and restored by adding Apple's faces. CI's font set now carries
Apple's four Times New Roman faces instead of Word's, so CI and the laptop
draw the same font; every engine on the site was affected the same way.

## Picture Effects (PARTIALLY DONE)

### Wide inline pictures and paragraph marks

Ordinary inline pictures wider than the text column do not reserve a second paragraph-mark line (cases176–179, local Word references). Keep the historical extra-line rule only for actual Office OLE previews, identified by the Office OLEObject child, rather than all w:object images. Static VML pictures without an OLEObject are ordinary pictures. The existing alfies_arc_adult_safeguarding_policy OLE preview remains unchanged.

**Done:** Smooth outer shadow (rasterized Gaussian blur mask via SMask), soft edge (edge-fade SMask on image), glow (centered blur), inner shadow (inverted blur mask), reflection (flipped image with gradient SMask). All use the same rasterized mask + SMask XObject infrastructure. Test fixtures: case56 (shadow variations), case57 (2D effects), case58 (3D effects — deferred).

**Remaining (deferred — no real-world fixtures use these):**
- **3D effects** (`a:scene3d`, `a:sp3d`) — bevel, metal frame, perspective rotation. Would require 3D lighting simulation. case58 has test fixtures ready.
- **Preset shadows** (`a:prstShdw`) — 20 built-in shadow presets. Need mapping table from preset names to parameters.
- **Theme color resolution in effects** — inner shadow with `a:schemeClr` falls back to black instead of resolving the theme accent color.
- **Header/footer pictures** — their effect XObjects are embedded but never drawn, so shadows and glows on header/footer images are missing (spring-cleaning finding 3).
- **Shapes and text boxes** — effects are parsed for pictures only (`parse_pic_effects`).

Also done: `a:lum` brightness/contrast (`da43eadf`).

## CJK Rendering Polish (TODO — MEDIUM IMPACT)

Core CJK support is implemented: CIDFont/Identity-H/ToUnicode encoding, Word-compatible substitution by fontTable charset and family (one platform-independent list, `cjk_fallback_fonts`), per-character font fallback at render time, script-based run splitting via `w:rFonts @eastAsia`, and vertical text in table cells. Remaining spacing/positioning issues:

1. **`w:firstLineChars`** (DONE — `15a1bc3f`, 2026-04-13) — `leftChars`, `hangingChars` and `firstLineChars` are parsed (`docx/mod.rs`).
2. **Vertical text centering** — `render_vertical_cjk_cell` uses a simplistic height calculation (chars x font_size) that doesn't account for paragraph spacing, causing vertical misalignment in merged cells.
3. **East Asian line height is 1.3× the font's Windows metrics** (DONE
   2026-09-14, was "Fallback line-height fidelity", annotation #8) — the ~1.73
   ratio seen in references is Malgun Gothic's 1.33 × 1.3. Word lays out any
   East Asian font (has CJK/Hangul/kana glyphs) at 1.3 × (winAscent+winDescent),
   no hhea lineGap, extra leading above the glyphs; exact-height boxes still
   bottom-align at winDescent; the docGrid counts cells with the same height
   (16pt YaHei on an 18pt grid → 2 cells, 10.5pt Yu Mincho → 1). Verified to
   0.1pt on line pitch in east_asia_conference_form, chinese_student,
   taiwanese_education, japanese_medical, tokyo_welfare. A run of nothing but
   spaces keeps the plain metrics so it cannot raise a Latin line
   (destination_loyalty: a lone MS Mincho space in a Times New Roman line);
   empty paragraph marks, tabs and the blank line after a break keep the real
   metrics (japanese_interlibrary_loan loses 20pp SSIM otherwise).
   `embed::compute_line_metrics`, `layout::run_line_metrics`.
   Open: where the extra leading sits. All-above matches exact boxes, but the
   first auto-spaced baselines in east_asia_conference_form (Batang 20pt after
   an empty paragraph) and japanese_land_development (heading 5, Yu Gothic
   10.5pt) land ~3pt higher than all-above predicts, closer to a half-above
   split. Needs a clean no-grid, no-header, text-first measurement.
4. **Linux CJK fallback list was Noto-only** (DONE 2026-09-14) — the Linux
   lists in `fonts/mod.rs` and `pdf/fonts.rs` named only Noto Sans CJK, which
   the CI runner lacks, so glyphs in runs whose font was missing were dropped
   outright (east_asia_conference_form on gh-pages: title shrank to "2024 발표",
   text boundary 0%). Locally the fixture looked fine only because
   `engine_compare.py` ran the binary without `DOCXSIDE_FONTS` (the
   `.cargo/config.toml` env applies only under cargo) and macOS system fonts
   covered the gap. Now one platform-independent list (`cjk_fallback_fonts`)
   leads with the vendored Word fonts, Apple/Noto faces trail, and the compare
   script passes `DOCXSIDE_FONTS`. The per-character rescue font
   (`cjk_rescue_fonts`) is ranked by glyph coverage of the missing characters.
5. **Theme `a:ea typeface=""`** — `asciiTheme/hAnsiTheme="minorEastAsia"`
   (47 table-label runs in east_asia_conference_form) resolves to an empty
   name and registers Helvetica. Word resolves the empty typeface through
   `a:font script="Hang"/"Jpan"` by the run's East Asian language (→ 맑은 고딕
   here); we should do the same.
6. **Missing CJK fonts substitute by charset + family** (DONE 2026-09-14) —
   the reference shows Word turned HY헤드라인M and 새굴림 (fontTable
   `charset=81`, `family=roman`) into Batang, rescuing kanji Batang lacks with
   MS Mincho. `classify_cjk_script` now reads the fontTable charset (0x80 JA,
   0x81 KO, 0x86 SC, 0x88 TC) before name hints, then Hangul/kana in the name or
   text; `cjk_fallback_fonts` orders serif vs sans by `w:family`. The fixture's
   font set now matches the reference except Helvetica for item 5.

## Bundled Fallback Fonts (TODO — MEDIUM IMPACT)

We rely entirely on installed fonts (tests and CI use the private assets repo's `fonts/` via `DOCXSIDE_FONTS`); a font that resolves nowhere gets Arial, Liberation Sans, Arimo, Helvetica or DejaVu Sans, then Type1 Helvetica as a last resort. This produces inconsistent output across environments (servers, Docker, CI). Should bundle metric-compatible open fonts behind a feature flag:
- **Carlito** — metric-compatible with Calibri (the most common Word font)
- **Caladea** — metric-compatible with Cambria
- **Liberation Sans/Serif/Mono** — metric-compatible with Arial/Times New Roman/Courier New

Metric compatibility means identical advance widths, so layout stays correct even with substitution. Ensures consistent output without requiring specific system fonts.

## Paginator Extraction (TODO — MEDIUM IMPACT, HIGH ARCHITECTURAL VALUE)

The `render()` function in `pdf/mod.rs` mixes pagination with rendering. widowControl, keepNext, keepLines, and tblHeader are already implemented inline, but extracting a dedicated pagination pass would:
1. **Clean up widow/orphan / keep-* logic** — currently embedded in the render loop with complex state tracking. A separate pass would be cleaner and more correct.
2. **Enable look-back wrapping** — paragraphs before a floating image anchor can't wrap beside it because the float zone isn't set until the anchor renders. Requires two-pass layout.
3. **Enable post-pagination field resolution** — PAGE/NUMPAGES fields could be resolved after layout instead of during rendering.

Architecture: a `Paginator` takes the document model and produces `Vec<Page>` where each `Page` contains positioned elements. The PDF renderer then simply draws them. This is a significant refactor but would simplify the render loop and enable features that require look-ahead/look-back.

## Performance: Image Pipeline + Test Harness (TODO — LOW IMPACT, found 2026-10-08)

GPU acceleration was considered and rejected for the converter: the work is branchy and serial (zip, XML, style resolution, line breaking, table fit, font subsetting, PDF writing), output is vector so nothing is rasterized, and typical documents convert in 10–450 ms (release) — less than GPU device start-up. It would also add heavy platform-specific deps. The real wins are on the CPU:

1. **Image recompression is the only converter hotspot.** `cases/case59` (28 MB, four 2.5–9 MB JPEGs) takes ~5.1 s, almost all in `src/pdf/images.rs` `embed_image_xobject` (decode → Lanczos3 `resize_exact` → JPEG re-encode). Options, cheapest first:
   - Decode/resize images in parallel (rayon or `std::thread::scope`) inside `embed_all_images`, then write the XObjects sequentially to keep object IDs deterministic.
   - Faster resampling: `fast_image_resize` (SIMD), or a cheaper filter (Triangle/CatmullRom) for large downscale ratios where Lanczos3 is visually indistinguishable.
   - Scaled JPEG decode (1/2, 1/4, 1/8 via zune-jpeg / jpeg-decoder) when the target is much smaller than the source.
   - Pass original JPEG bytes through as DCTDecode when no downscale/crop/recolour is needed.
   Verify with the visual suite and the deterministic-output check; case59 is the benchmark.
2. **Test harness metrics.** The SSIM with ±8px spatial search in `tests/common/mod.rs` (`ssim_score`) is the one genuinely GPU-shaped workload (uniform per-pixel math; ~17 CPU-s for a 205-page fixture). A compute-shader (wgpu) port is possible, but try first: rayon over the window search within a page, integral images for the local means/variances, and SIMD. Only worth doing if test turnaround becomes a bottleneck.

## Vertical Drift Investigation (TODO — HIGH IMPACT)

**Root cause identified: glyph advance width precision.** Thorough investigation (April 2025) proved the drift is NOT from line height errors — line heights match Word exactly. The drift comes from our character advance widths being ~0.003pt/char wider than Word's at 12pt, causing ~1 fewer character per line on borderline lines. Over 48+ pages, this compounds into 1 extra page.

Evidence:
- Character-level comparison on case4 (Calibri 12pt): by char 89, our x-position is +0.27pt ahead of Word's (0.003pt/char average drift)
- Our widths match the font file exactly (verified via fontTools), but Word's widths are systematically narrower
- Removing hhea lineGap from line_h_ratio was tested and disproven — caused massive regressions with no benefit
- ceil() rounding of line heights was tested and disproven — too aggressive, destroyed all scores

**Disproven hypotheses:**
1. Line height formula (hhea lineGap inclusion) — disproven: removing it causes 80+ regressions
2. Line height rounding (ceil to whole points) — disproven: too aggressive, 90% regressions
3. Margin calculation error — disproven: our margins match DOCX spec exactly
4. Image paragraph height rounding — fixed in prior work
5. Table trailing spacing — disproven in prior work
6. Line-break tolerance (0.07–0.75pt) — tested April 2026: fragile, can't distinguish bias from genuine overflow. Any tolerance >0.07pt regresses Cambria-based cases (case11)
7. Global width correction factor — tested April 2026: helps Calibri/TNR but overcorrects Cambria. Magnitude varies by font and by font size.
8. Per-font width correction factor — tested April 2026: Calibri/TNR=0.99985, Arial=0.9999, others=1.0 gives 3 improvements, 0 regressions. Safe but captures only ~30% of needed correction. Can't go further because correction is size-dependent.

**Root cause confirmed:** Word's DirectWrite engine applies proprietary grid-fitting corrections that vary per glyph AND per font size (signs flip between sizes). These corrections are not in font data and can't be reproduced by FreeType or rustybuzz. The signed bias varies by font: Calibri +0.007pt/char, Arial +0.003pt/char, Cambria ~0pt at 12pt.

**Next steps — data-driven width correction (April 2026):**

The correction varies per glyph (some positive, some negative — not a uniform scale factor) and is likely ppem-level rounding from DirectWrite hinting. Plus, 6/9 inter-glyph adjustments in Word's TJ output aren't in font data at all. Rule-based reverse-engineering has hit a wall — data-driven learning is the natural next step.

**Phase 1 — Data collection pipeline:**
Build a synthetic DOCX generator producing controlled text for width extraction:
1. Single-glyph sheets: one character repeated per line (e.g., "TTTTTT...") at a specific font + size. TJ positions in Word's PDF give the exact advance width Word uses.
2. Bigram sheets: pairs like "THTHTH..." to capture inter-glyph adjustments (the proprietary DirectWrite corrections).
3. Font × size matrix: Calibri, TNR, Arial, Aptos, Cambria at sizes 8–24pt in 1pt steps.
Pipeline: `generate_width_sheets.py → .docx → Word conversion → extract_widths.py → width_corrections.json`

**Phase 2 — Analysis (formula or model?):**
Before any ML, test whether corrections follow a discoverable formula:
- ppem rounding: `round(advance * ppem / UPM) / ppem * fontSize` where `ppem = fontSize * 96 / 72`
- hdmx table: TNR has device-specific metrics — check if they match Word's widths
- Linear correction per font: maybe each font just needs a single scale factor per size
If a formula fits → implement directly. No model needed.

**Phase 3 — Correction table or model (if no formula):**
- Option A — Lookup table: `{font, ppem, glyph_id} → width_correction_pts`. ~20K entries × 4 bytes = 80KB. Simplest, most accurate.
- Option B — Small regression model: input = `[glyph_advance_units, lsb, rsb, ppem, font_class]`, output = `width_correction_pts`. 2-layer MLP (~1K params). Generalizes to unseen fonts.
- Option C — Per-font scale function: learn `correction(ppem) → scale_factor` per font. Small polynomial per font.

**Phase 4 — Kerning corrections (stretch goal):**
Same pipeline for bigrams. Input space is `glyphs²` but only ~500 common pairs matter. Sparse lookup table.

**Previous next-step ideas (status updated April 2026):**
- ~~Add a small configurable "text width tolerance"~~ — tested, fragile, regresses low-bias fonts
- ~~Test ppem-based rounding at various DPIs~~ — tested in March 2026 (see kerning_and_shaping.md), none match
- Create more diagnostic fixtures with different fonts/sizes — still valid, needed for Phase 1 data collection
- **Interim safe win:** ship per-font factor (Calibri/TNR=0.99985, Arial=0.9999, others=1.0) for 3 clean improvements while data pipeline is built
- **Analysis tooling:** `tools/experiments/width_analysis.py` extracts per-char signed width errors from reference PDFs

**Blocked annotations (triage 2026-07-03):** four open annotations diagnosed as
this drift class flipping a soft page break — no targeted per-case fix exists;
they should clear when width/height fidelity improves (matching triage notes
appended in `annotations.json`):
- **#59 brazilian_logistics_study p9** — pure width-drift: extra wrapped lines
  by page 8 spill ~4 blank spacer paragraphs above the Figura 2 caption (~82pt).
- **#82 czech_municipal_grant_form p2** — page-1 line/row heights ~28pt short;
  the intro paragraph Word overflows to page 2 (`lastRenderedPageBreak` on it)
  fits our page 1, so page-2 content sits ~26pt high. Row-height deficit class
  (see next section) as much as width drift.
- **#124 english_town_council_report p3** — 11pt TOC rows ~1.4pt/row short; the
  16-empty-paragraph stack straddling pages 2–3 fits our page 2 entirely, so the
  page-3 bordered box starts flush at the top margin (~25pt high, ~10pt of it
  from page-top space_before suppression once the box lands there).
- **#8 east_asia_conference_form p1** — different sub-class: the Korean fonts'
  CJK fallback has line ratio ~1.27 vs ~1.73 in the reference, so every
  `atLeast` row collapses to its trHeight while Word grows them (~116pt lost
  across one table). Belongs with Bundled Fallback Fonts / CJK metrics, not
  Latin width correction.

## Table Row Height Deficit (LARGELY DONE — 2026-10-02)

Border bands (layout fixes 4 and 14) and per-line cell heights (`53764bee`)
closed most of this; case51 now has Word's 2 pages (J 61.3, SSIM 85.6).
Residuals are in the layout round's open findings (slovak_eu ~0.03pt/row,
romanian's at-least row). Original notes:

Our table rows run ~0.5–0.9pt shorter than Word's, compounding down a page of
stacked tables (case51: −6.5pt accumulated over 3 tables, measured via stext
anchor diffing). Consequence: content that Word pushes to the next page can
stay on ours. case51's reference has a blank page 2 (Word's implicit final
paragraph mark spills after the doc-ending table at 710.9pt); ours ends 8.2pt
higher so the mark fits and no page 2 is emitted — this alone costs case51
~22pp SSIM (missing page scores 0). The implicit final-¶ model and the
end-of-cell-mark suppression after nested tables are already in (2026-07);
only the per-row height accounting remains.

## Paragraph Border Groups (DONE — 2026-09-09, annotation #224)

Word joins adjacent paragraphs into one border group only when their `w:pBdr`
*and* indentation (left/right/hanging/firstLine) are identical; inside a group
no bottom/top rule or padding is drawn at the joins (only `between`).
`joins_border_group` in `pdf/helpers.rs` replaces the earlier "collapse only if
a paragraph is empty" heuristic from #122 — that heuristic was misreading
samtale p2 items 12/13, whose rules survive because item 13 has a direct
`w:ind left=1128` vs the numbering level's 1131 (3 twips → separate group).
samtale p1: the br-only spacer above "Medarbeiderens navn" no longer draws its
own rule; Din leder → name-line spacing 55.57 → 52.32pt (Word 52.56). Jaccard
-1.9pp because the -2.85pt drift accumulated above (br-only paragraph -1.22pt,
three bullets -0.53pt each) was previously masked by the bogus 3.25pt border
space. Corpus has no other adjacent identical-border/identical-indent
non-empty pair, so the other 14 pBdr fixtures are unchanged.

## Header Multi-Float Wrap (DONE — 2026-07-02, annotation #212)

`hdr_fz` is now a Vec of zones; all wrapping floats (same-paragraph + earlier
paragraphs) constrain the text bounds together, and paragraph indents are
measured from the column edge with float bounds clipping (Word semantics).
`parse_object_floating_image` honors `w10:wrap type="square|tight|through|
topAndBottom"`. Letterhead center now within ~5pt of reference. HR `o:hrpct`
width now uses the indent-adjusted paragraph box (2026-10-03).

## Annotation Fixes 2026-07-03 round 2 (#121 #133 #167 #190 #219 — DONE)

- **#133 / #190 table row splitting**: the `row_h > page_content_h * 0.5` gate in
  `table.rs` blocked Word-style row splits. Word splits any non-cantSplit row that
  overflows the page remainder — EXCEPT rows with an explicit `trHeight` (exact or
  atLeast), which always migrate whole (verified: arizona/traditional all-trHeight
  tables never split in Word; isla/master_thesis no-trHeight rows do). New gate:
  no trHeight + multi-item cells + first chunk fits + `available_h > 50pt` sliver
  guard (our lines run a few pt short of Word's, so near-boundary rows see phantom
  space — victorian p8 had 43pt where Word had 6pt; lower once line-height
  fidelity improves). isla +5.8pp TxtBnd, master_thesis +24.1pp TxtBnd; collateral:
  carbon_farming +38pp TxtBnd, stem_partnership +9pp, english_town_council +7.2pp.
- **#167 floats in vAlign-centered cells**: the cell's vAlign centering offset was
  baked into the anchor base handed to `render_cell_floating_shapes` /
  cell floating images. Word anchors paragraph-relative floats to the cell content
  top. `render_cell_content` now takes `valign_off` and adds it back for float
  anchors only. The 50cm arrow in japanese_land_development_sign_form now spans
  table-bottom → ground-hatch exactly (scores flat — tiny ink area).
- **#219 exact line-rule baseline**: baselines were placed `font_size *
  ascender_ratio` below slot top; `ascender_ratio` folds in hhea lineGap, so a
  big-lineGap CJK substitute (Hiragino for 方正小标宋简体) pushed descenders out of
  the fixed `lineRule="exact"` box into the table border below. Word bottom-aligns
  the exact box: baseline = box bottom − winDescent (identity: `line_h_ratio −
  ascender_ratio`). `exact_baseline_base` in `render_paragraph_block` (both
  baseline sites). chinese_student_union +4.4pp SSIM/+2.3 Jaccard; polish_archery
  +8pp, auditor_regulatory +5.9pp Jaccard.
- **#121 trailing-break mark line**: the empty line a trailing `<w:br/>` leaves
  was sized with the break run's font (samtale: 26pt br), but it holds only the
  paragraph mark — Word sizes it by the mark's rPr (12pt here). Per-line loop in
  `render_paragraph_block` now uses `paragraph_mark_font_size` for the final
  break-created empty line when known (break char still sizes the line it
  terminates; intermediate br-created lines keep the break size). samtale +57.7pp
  TxtBnd / +10.9pp SSIM / +6.5pp Jaccard, german_mezzo_soprano +2.2pp.

## Annotation Fixes 2026-09-16 (#229 #230 #200 #232 #152 #238 — DONE)

Five fixes, one commit each, every run diffed against a snapshot of the
pre-round suite (221 fixtures, 203/221 page counts match); no fixture lost more
than 0.1pp, 15 improved.

- **#229 picture brightness/contrast** (italian_evaluation_minutes p7 "stamp"):
  the signature scan's `a:lum bright/contrast=30%` was ignored; Word's +30/+30
  washes the pale stamp inside the crop window to white. `parse_lum` +
  `apply_lum` (LibreOffice's DrawingML mapping: contrast scales about mid-grey,
  brightness offsets). Only italian and mongolian_human_rights_law changed.
- **#230 inline picture baseline** (italian signatures under their names): Word
  sits the picture bottom on the baseline and the picture top on the paragraph
  top; the picture line advances by picture height + descent of the runs with
  visible glyphs + multiple-spacing leading of the text run's font (the picture
  run's own `w:sz` never counts; a picture-only line has neither). Measured on
  italian (+2.2), english_town (+2.6 at 1.15), family_kinship (+7.6 at 1.5) and
  old_blue_truck (+0). `pdf/layout.rs`: `inline_image_line_extra`,
  `inline_line_advance`, `picture_line_bottom`; `render_paragraph_lines` takes
  the paragraph (ascent, descent). 12 fixtures improved (case16 J +5.7 / SSIM
  +16.4, russian_chess +2.6, croatian_thesis +2.6, english_town +1.6,
  polish_tender +1.3, ut_koer +1.1, usep +1.0). `after_image_boost` now only
  applies after block pictures (`para.image`), whose height is still bare.
- **#200 #232 floating tables** (croatian_grant p4-5 green box,
  indigenous_innovation p1-2 DEFINED TERM table): a `vertAnchor="text"` table
  with `tblpY ≥ 0` paginates like an inline table (rows split/migrate per the
  trHeight rule); only a negative-tblpY table (pendulum) moves whole with its
  anchor. `render_table`: `flows_inline` / `keep_with_anchor`. indigenous J +6.0
  / SSIM +7.7.
- **#152 pre-anchor wrap** (case41 p3): the paragraph before an image-only
  anchor paragraph wraps around the float positioned from its full-width
  layout; Word leaves the float there while the paragraph grows. Look-ahead
  installs the zone (top raised by the paragraph gap for its own geometry) and
  hands the anchor to the next paragraph via `pending_float_anchor`; replaces
  the "narrow the last line if the picture is under half the column" heuristic.
  Only paragraph-relative floats (Offset / AlignTop): `resolve_fi_y_top` puts an
  AlignTop-relative-to-paragraph float at the *page* top, which narrowed
  indonesian_benchmarking p6 until the look-ahead computed the top itself.
  case41 J +3.9 / SSIM +4.2.
- **#238 compressPunctuation** (taiwanese heading's lone 決): with
  `w:characterSpacingControl compressPunctuation` Word trims the full-width
  closing marks already on the line, evenly and by at most ¼ em, to keep one
  more character (marks on the taiwanese page advance 12–16pt at 16pt, never
  less). `docx/settings.rs` → `Document::compress_punctuation` →
  `RenderContext` → `CjkLayout` → `compress_punctuation()` in `pdf/layout.rs`.
  Gated: 191 fixtures say doNotCompress, only 6 compress. taiwanese J +3.8 /
  SSIM +9.1, tokyo_welfare +3.6 / +8.2, japanese_land_development −0.1 SSIM.

Left open with triage (2026-09-16; #66, #220, #158/#195, #237, #240, #241
closed in the 2026-09-18 round above): #233 (Merriweather not vendored), #186
(alfies p1 13.9pt high: OLE object paragraph + 24pt empty marks, not
investigated), #93 (Word compresses justified inter-word spaces to fit one
more word — see the 2026-09-18 notes), #185/#239 (vague), #8/#59/#82/#124
(systemic drift).

Follow-ups: `resolve_fi_y_top` should treat AlignTop-relative-to-paragraph as
the anchor top (then the look-ahead can drop its paragraph-relative filter);
opening brackets are not compressed; table cells still drop run-level inline
pictures (`EMPTY_INLINE_IMAGE_MAP`), so the picture-line rule does not reach
them; `Picture Effects` above can list `a:lum` as done. From the `/simplify`
review of this round: (1) one picture model — `docx/paragraph.rs` hoists a lone
inline picture into `Paragraph::image` (bare height + `after_image_boost` on
the next paragraph) while two or more stay in runs (`inline_line_advance`);
removing the hoist would give every picture the measured rule (~40
`para.image` renderer references); (2) a shared `decode_raster` so `a:lum`,
crop and soft-edge are applied once instead of per format branch, and reach
`embed_reflection`; (3) the floating-table keep-together could be geometric
(`fp.y > saved`) rather than reading raw `tblpY`, which would also decide
page/margin-anchored tables — no fixture evidence yet; (4) the look-ahead's
full-width lines could be reused when its zone reaches no line.

## Annotation Fixes 2026-09-15 (#231 #228 #236 #235 #223 #234 — DONE)

Five localized rendering bugs, one commit each; every fix changed only its own
fixture's visual hash, no score regressions across the 221 scored fixtures.

- **#231 EMF clip path** (indigenous footer logo as a black block): `pdf/emf.rs`
  skipped EMR_SELECTCLIPPATH, so the clip rectangle stayed an open PDF path and
  merged into the next FILLPATH. Now `W n`/`W* n` per the fill rule; ABORTPATH → `n`.
- **#228 bitmap EMF + inline rotation** (italian p7 signature missing): a lone
  EMR_STRETCHDIBITS is wrapped as BMP by `docx/emf.rs::emf_to_raster` (WMF-style);
  inline pictures keep `a:xfrm@rot`, draw turned about the frame centre and occupy
  the rotated bounding box (`EmbeddedImage::layout_size`) — Word gives a -90°
  56×108pt frame a 56pt line. `para.image` block pictures (a lone picture in its own
  paragraph) still ignore rotation. #230 stays open: Word puts the text baseline at
  an inline picture's bottom, we centre the picture on the text (`img_bottom = y +
  font_size - line_max_img_h` in `pdf/layout.rs`).
- **#236 / #235 table border inheritance** (croatian_grant_guidelines): inline
  `w:tblBorders` replaced the style's set wholesale; now merged per side
  (`merge_table_borders`), so Table Grid's insideH/insideV survive. Rule confirmed
  for #235: at a page split each row draws its own top/bottom border, which for
  inner rows is insideH — a table with insideH=nil shows no line at the split.
- **#223 spAutoFit in table cells** (japanese_land_development arrow hidden):
  `render_simple_textbox` ignored `AutoFit::Shape` and painted the white box at
  Word's 110.6pt default height. Height computation shared via `textbox_height`.
- **#234 footnote laid out as a table** (auditor_regulatory_report_template):
  `parse_notes_simple` read only `w:p` children. Table rows are flattened to one
  paragraph each, cells joined by a space (ponytail note in `headers_footers.rs`;
  real column geometry needs Block support in `Footnote`).

Triage notes for the annotations left open: #233 (Merriweather not vendored — font
availability, not code), #237 (row split is paragraph-granular; needs line-level
`find_cell_split`), #232 (floating `tblpPr` table pushed whole to the next page instead
of breaking), #229 (`a:srcRect` crop is parsed but the italian stamp still shows —
check the inline draw path), #238/#158/#195 (font/width class).

Follow-ups from the `/simplify` review of this round (not done): the EMF translator
still leaves immediate-mode segments (MoveTo/LineTo outside BeginPath) and mid-path
SaveDc/RestoreDc unhandled — a `path_pending` flag emitting `n` would generalise the
#231 fix; mixed vector+bitmap EMFs need the DIB placed as an image XObject inside the
form (bitmap-only EMFs take the raster path today); block pictures (`para.image`, five
draw copies across `pdf/mod.rs`, `table.rs`, `header_footer.rs`, `textbox_render.rs`)
want one shared `draw_embedded_image` that applies rotation + effects; body/header
flow still reserves `height_pt` for TopAndBottom textboxes while rendering uses
`textbox_height`; `TableStyleDef` parses no `pPr`/`tblCellMar`, so `has_tbl_style`
is a proxy for "style defines borders"; `Footnote { paragraphs }` should become
blocks so footnote tables keep their geometry.

## Annotation Fixes 2026-09-09 (#225 #226 — DONE)

- **#225 / #226 split-row borders**: `render_partial_row` drew the cell's top
  border only on the first fragment and the bottom border only on the last, and
  stretched non-final fragments to the body bottom (`fill_to_bottom_y`, added
  2026-04 without fixture evidence). Word closes every fragment as a complete
  box: master_thesis p2 ref bottom border at y=181.5 (fragment = 3-line
  paragraph + 1 empty paragraph, ends at an item boundary with ~9pt of body
  space left unused), p3 top border at the top margin (y=771); slovak_eu shows
  the same. Now every fragment draws all four borders and ends after its last
  fitted item. slovak_eu +5.6pp SSIM, isla +1.4pp, master_thesis +0.25pp; 10
  split-row fixtures scored, no regressions. Remaining on master_thesis p2: our
  body bottom sits ~19pt lower than Word's (footnote area: our separator 3pt
  below body bottom vs Word ~10.5pt; footnote text 15pt lower; last footnote
  line ends at the margin with no space-after), so we fit two extra empty
  paragraphs (26.9pt) and the bottom border lands at y=157 instead of 181.5.

## Annotation Fixes 2026-07-03 (#114 #118 #193 #214 #218 — DONE)

- **#114 ellipsis line breaks**: UAX #14 allows a break after U+2024/25/26 before
  digits, splitting TOC dot-leader tokens like `Preparation………45`. Word keeps
  them unbreakable; `split_preserving_spaces` now filters those break positions
  unless followed by whitespace (unit test in layout.rs).
- **#118 leading after tall inline image**: Word lays the line following a tall
  inline image one full line height below the image bottom (leading above the
  text). `after_image_boost` in `render_paragraph_block` extends the following
  paragraph's first baseline offset and block height by the missing leading
  (skipped for empty/grid-snapped/image paragraphs). brazilian_logistics p4 gap
  1.9pt → ~9pt (Word: 9.2pt).
- **#193 oversized list labels**: the ±1pt guard in `label_boosted_line_h` is
  gone (the "handled separately" path it referenced never existed) and the new
  `label_boosted_baseline_offset` drops the first baseline to the label's
  ascent — a 20pt number label on 10pt text now sizes the first line like Word.
  samtale +2.9pp SSIM, +40pp text-boundary; case16 +6.5pp, family_kinship +5.5pp SSIM.
- **#214 / #218**: see their sections (vAlign center, clear="all").

## Run Properties

### `w:emboss` / `w:imprint` / `w:shadow` / legacy `w:outline` (DONE — `a07f232b`, 2026-06-19)

Parsed in `runs.rs` (`legacy_text_shadow`), drawn as an offset grey copy
(an approximation, see the `ponytail:` note in `layout.rs`); `w:outline` maps to
a text outline with no fill.

### `w:shd` on runs (DONE — `07a7198a`, 2026-04-20)

`parse_run_shd` (`docx/mod.rs`); layout fix 19 covers `solid`/`pctNN` patterns.

## Paragraph / Layout Features

### Final descent of automatically wrapped running-head pictures

Keep the paragraph-mark descent between picture lines, but omit its duplicate contribution after the final line when the image runs and mark use the same size (cases180–187, local Word references). Both measured height and render cursor use the same rule. Exact spacing, explicit breaks and differently sized marks retain their existing paths. Depends on the wrapped-picture rendering and measurement fixes. Tall final pictures and three-picture layouts retain separate residual discrepancies; this change does not claim to resolve them.

`w:jc="distribute"` (see "Distributed Alignment"), `w:gutter` (`552c09cc`, incl.
`gutterAtTop`/`rtlGutter`), `w:pgBorders` and sectPr `w:vAlign` (`79b7f850`) are
done.

### Column regions (MOSTLY DONE — 2026-10-05, branch `columns`; case80 47.8 → 65.4, case81 71.0)

Measured on case80/81/82 and ~30 Word probes, implemented in `pdf/mod.rs`:
- Balancing: a region ending at a continuous section break, without a column
  break in it, is balanced on its (last) page. Word starts at the content
  height over the column count and adds a line until it fits, a line needing
  all its leading (not the shortest fit: 14/14/12 where 14/13/13 fits). The
  content counts a page-top heading's dropped space before.
- The region ends at its deepest column; the last column's final space after
  counts, clipped to the balancing height. What follows starts there.
- A mid-page region's columns start below the pending space after; the first
  paragraph still opens with max(space after, its space before).
- `w:sep`: 0.75pt, region top to deepest column on each page (trailing space
  after included), up to the last column holding text.
- Overflow at an empty mid-page column goes to the next page. A paragraph
  continues column by column and page by page, keeping widow control.
- A paragraph ending in a column break puts its mark on a line at the top of
  the next column.
- An empty section-break paragraph has zero height; mid-page gap = prev after
  + max(0, next before − break after). (Replaced the column-change line rule.)
- An autofit table in a newspaper column is squeezed in proportion to each
  grid column's room above its longest word.
- Footnotes in a multi-column section sit at the foot of the citing column,
  in its width; only that column shrinks. A note whose style is missing or
  sets no spacing takes the document defaults.

Open:
- Footnote continuation across columns (case82 J 21.3): Word lays the note
  area out in the columns and may continue a note under the next column so
  more body text fits (p1: footnote 1 runs 3 lines under column 1, 2 under
  column 2; in another probe the whole note stayed in column 1). Needs a rule
  for how much of the note must stay with its reference.
- case80's 4-column region: Word's separator runs one line (15pt) below the
  space after of column 1's last paragraph; unexplained.
- The table probe's region ends 1.6pt off.
- Each balancing trial lays the whole region out again (a few times).

### `w:mirrorMargins` (TODO — MEDIUM IMPACT)

Parsed in `settings.rs` into `Document::mirror_margins` (used only for odd/even
section-break filler pages), never applied to margins. Fix: swap left/right
margins (and the gutter side) on even pages.

### `w:textAlignment` (TODO — LOW IMPACT)

Vertical alignment of runs within a line (top/center/baseline/bottom/auto). Only superscript/subscript are handled; the paragraph-level `w:textAlignment` property for mixed-size runs is not.

### RTL / BiDi (TODO — HIGH EFFORT, MEDIUM IMPACT)

`w:bidi` (paragraph-level) and `w:rtl` (run-level) right-to-left support is completely absent. Requires implementing the Unicode BiDi algorithm (UAX #9) for correct visual reordering. Architecturally complex — affects line building, text rendering, and alignment.

## Table Features

### Empty anchor line at the top of a float

Clear a blocking float before consuming a textless line whose box intersects its top edge. A 0.05pt text offset otherwise lets that line pass above the table and loses its height when the next paragraph clears it (cases170–175). Keep the existing side-strip policy. Covered by main's `40275abe` (a floating table pushes a paragraph whose first line reaches it); PR #35's separate empty-line check changed no fixture and was dropped, its fixtures kept.

### Cell paragraph `indent_right` in render pass (DONE — 2026-07-02, annotations #215/#217)

`table.rs` computed the render-time `text_w` without subtracting `para.indent_right`
while the wrap width in `table_layout.rs` did — centered cell text shifted right by
`indent_right/2` and justified text overshot the cell border. Both spots now match
the layout width (romanian_quality_evaluation_strategy SWOT headings).

### `w:vAlign="center"` text sits ~3pt high (DONE — 2026-07-03, annotation #214)

Root cause: baselines sit `font_size` below each line top, so a fallback font
with big leading (Hiragino Sans GB for 仿宋_GB2312: lineGap 0.5em) dangles that
leading below the ink of the last line, and centering the full block rode the
ink high. `cell_content_h_for_valign` now drops the last line's unused bottom
leading — but only when the font is a metric-changing substitution
(`FontEntry.is_substituted`): with the document's real font (Yu Mincho in
japanese_land_development_sign_form) the full-line-box centering already
matches Word, and subtracting regressed it −2.9pp. chinese_student_union +2.6pp SSIM.

### `w:tblLook` / `w:tblStylePr` (DONE)

Table conditional formatting (firstRow, lastRow, firstCol, lastCol, banded rows/cols): flags parsed in `tables.rs`, overrides from `styles.rs`, accumulated in one struct since `f9d8268d`.

### Table auto-fit vs `tblW` (NO IMPACT — corpus check 2026-05)

Our `auto_fit_columns` uses `gridCol` widths from `tblGrid`, ignoring the specified `tblW` when `type="dxa"`. Word treats `tblW` as the authoritative total width and scales/caps columns to fit. This causes tables to render at full page width when python-docx (or other generators) emit oversized `gridCol` values alongside a smaller `tblW`.

**Verified empty in current corpus**: a sweep of all `tests/fixtures/scraped/*` documents found zero tables where `gridCol` total exceeds the `tblW` value (tolerance 100 twips). The bug is real per OOXML, but no fixture triggers it — implementing this clamp moves zero scores. Park until a real-world fixture exhibits the mismatch.

### Percent-based widths: `tcW`/`tblW` `type="pct"` (PARTIALLY DONE 2026-06)

`twips_attr` reads `w:w` as twips regardless of the `w:type` attribute. For `type="pct"` the value is in fiftieths of a percent (5000 = 100%). **Implemented**: `Table.width_pct` is parsed from `tblW type="pct"` and `apply_pct_width` scales columns to pct × content width — but ONLY for tables whose `tblGrid` is missing (grid inferred from row `tcW` values, which preserves pct proportions). When a real tblGrid exists, Word renders the grid widths as-is even when the pct width disagrees (observed: arizona 115%, zimbabwe 100% vs grid at 102% of content — scaling them regressed scores). **Remaining**: `tcW type="pct"` is still mis-read as twips for per-cell preferred widths; harmless today because grid widths dominate, but would matter for Word's full preferred-width algorithm (§17.18.87).

## Document Features

### Footnote pagination of split paragraphs (DONE — 2026-09-09, annotation #221)

Word puts a footnote in the footnote area of the page where its reference mark is laid out, and a body line fits on a page only together with the footnotes it references. The split-paragraph path in `render_paragraph_block` used to reserve every footnote of the paragraph on the first page and register them all on the continuation page after the flush (hole on page N, notes on page N+1, lines broken early). Now `WordChunk.footnote_id` records which `TextLine` carries which reference, `per_line_footnote_extra` charges each footnote to its line when computing `lines_that_fit`, and the first part's footnotes are registered before `advance_column_or_page`. Word also keeps one line for an empty footnote paragraph (sized by the paragraph mark) — `compute_footnote_height`/`render_notes_downward` count it. environmental_law_clinic_china Jaccard 0.085 → 0.210, russian_volunteerism_essay 0.247 → 0.670, czech_crisis_measure_notice 0.372 → 0.419.

**Remaining deviation**: when a reference line fits but its footnote does not, Word splits the footnote across pages with a continuation separator; we push the line to the next page instead (no overlap, rarely hit). Table rows (`table.rs` `row_fn_extra`) still reserve per row, which is right because rows are atomic.

### Endnotes (DONE — `703b4497`, 2026-05-12)

Endnotes flow at the document end (`render_endnotes_inline`). Remaining: they
are not paginated (merge fix plan item 7) and `endnotePr pos="sectEnd"`
(section-end placement) is not supported.

### Moved text `w:moveTo` (TODO — found 2026-10-03)

`collect_run_nodes` (`runs.rs`) has no arm for `w:moveTo`, so moved-in text is
dropped; Word's final view shows it. Deleted paragraph marks
(`w:pPr/w:rPr/w:del`) don't merge paragraphs either.

### Additional Field Codes (TODO — LOW IMPACT)

Only PAGE, NUMPAGES, STYLEREF, and PAGEREF field codes are supported. Others (DATE, TIME, AUTHOR, FILENAME, IF, MERGEFIELD, SEQ, etc.) are silently dropped — only the cached display text is used. For static PDF export this is usually acceptable since Word pre-computes the display text, but dynamic fields (DATE, PAGE in headers) may show stale values.

## Anchored Shapes: Canvas/Group + Z-Order (PARTIALLY DONE — 2026-06)

**Done (2026-06):**
- **Drawing canvas (`wpc:wpc`) and shape groups (`wpg:wgp`/`wpg:grpSp`)** — flattened at parse
  time in `src/docx/group.rs`: composes `off/ext/chOff/chExt` child-space transforms recursively,
  emits leaf `wps:wsp` (textbox or connector), and `pic:pic` as independently positioned shapes.
  Fixes isla_language_lesson_plan venn diagram + grouped boxes (+2.6pp Jaccard). Fixtures with
  groups: isla, arizona_physical_education_standards (header), ukrainian_municipal_heating_resolution.
- **`a:noFill` overrides style-ref fill** — explicit noFill no longer falls through to the
  `fillRef` theme fill (was rendering noFill ellipses as solid accent-color shapes).
- **Style `lnRef` strokes on textbox shapes** — shapes without explicit `a:ln` color now get the
  shape-style stroke (previously only connectors did).
- **Z-order via `relativeHeight`** — `Textbox.z_index`/`ConnectorShape.z_index` parsed from
  `wp:anchor`; non-behindDoc textboxes and connectors render into per-shape buffers deferred to
  page flush, painted above the page text layer sorted by z (Word stacks floating shapes across
  paragraphs). Fixes lenten_prayer_unity white link on purple band; connectors must interleave
  with shapes by z or letter strokes drawn over gradient circles disappear
  (vaccines_history_chapter T/Y/B).
- **Connector presets stay connectors** — `parse_wsp_shape` declines line/straightConnector1/arc
  presets without text so they reach the connector parser (preset-geometry path loses
  flipH/flipV and arc sweeps; regressed vaccines_history letters when lnRef strokes made the
  textbox parse succeed).

**Done since:** floating images take part in z-order (`5682de3c`, 2026-06-17);
text in preset shapes uses the shape's text rectangle (`1d41a24f`, `shape_text_rect`).

**Remaining:**
- **behindDoc shapes from later paragraphs** can still paint over earlier paragraphs' text
  (needs pre-pass/paginator).
- **Group flips/rotation** — group-level flipH/flipV and rot are ignored (rare); leaf connector
  flips work.
- **Canvas/group inside paragraph-level mc:AlternateContent** — only the run-level path
  flattens groups; `collect_textboxes_from_paragraph` still grabs the first wsp.

## Floating Image Positioning (TODO — MEDIUM IMPACT)

Floating images (`wp:anchor`) with large `posOffset` values can render off-page. Word appears to clamp or reflow these positions, but we render at the raw coordinates. Observed in `learning_cultures_dissertation` (rId14: column-relative offset 4702029 EMU = 370pt, placing a 334pt-wide image past the 612pt page edge). A naive right-edge clamp was tested but regressed `stem_partnerships_guide` — a more nuanced approach is needed (possibly only clamping when the image would be entirely off-page, or respecting wrap constraints).

Additionally, truncated/corrupt PNG images in DOCX files cause the `image` crate to fail with "unexpected end of file". Currently falls back to a 1x1 placeholder via `decode_png_raw` (using the `png` crate directly). Word renders these partially — investigate partial PNG decoding to match. Observed in `learning_cultures_dissertation` image1.png (216KB file, 2205 bytes short of complete IDAT data, no IEND chunk).

## `w:smallCaps` Rendering Accuracy (DONE — verified 2026-05)

`smallcaps_segments()` in `src/pdf/layout.rs` applies the per-character rule: only originally-lowercase characters are uppercased and rendered at 80% of the size (layout fix 39); originally-uppercase characters render at full size. Unit tests in the same file cover mixed/upper/lower/non-letter cases.

## SmartArt Remaining Work

Basic fallback rendering via pre-flattened `dsp:drawing` shape trees is done, with full geometry engine support (all 187 preset shapes). Remaining:

1. **Group shapes** (MEDIUM EFFORT) — `dsp:grpSp` groups with nested transforms. Need recursive parsing.
2. **Connector shapes** (MEDIUM EFFORT) — `dsp:cxnSp` connectors between shapes (arrows, lines).
3. ~~**Image shapes**~~ (DONE) — `a:blipFill` image fills parsed from diagram-specific relationships, rendered with cover-fill scaling and shape clipping.
4. **Full layout engine** (VERY HIGH EFFORT) — implement the constraint-based layout algorithm that interprets ~200 XML layout recipes. Only needed for files that lack the `dsp:drawing` fallback. Not planned for the near term.

## Charts Remaining Work

All 8 chart types are supported (bar, line, pie, area, doughnut, radar, scatter, bubble). Remaining:

- **3D charts**: `c:bar3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` — not parsed (`c:pie3DChart` is drawn as a flat pie)
- **Stock charts**: `c:stockChart` — not parsed
- **Combo charts**: two chart types overlaid on the same plot area — only the first is drawn
- **Data labels**: not parsed or rendered
- **Chart title**: not parsed or rendered
- **Secondary axes**: not handled (only the first `catAx`/`valAx`)
- **Chart label positioning**: axis labels still have small offsets vs Word (real font widths are used when the font is present, `62eed970`; `text_width_approx` is only the fallback).
- **Legend placement fine-tuning**: small positional offsets vs Word. Centering formula and spacing need per-chart-type calibration.

Done: stacked/percent-stacked bars (`7c2b9238`), theme minor font for labels.

## Tracked-Changes (Redline) Rendering (TODO — HIGH IMPACT for documents with revisions)

A document with tracked changes ("redline") keeps each edit as a revision:
`w:ins` / `w:del` runs, `w:moveFrom` / `w:moveTo`, and property changes
(`w:rPrChange`, `w:pPrChange`, …). Word's PDF export shows them marked up;
we render the final text (insertions plain, deletions dropped), which is
Word's "No Markup" view. Comments already render as Word does (scaled page,
balloon pane, `pdf/comments.rs`).

What Word's markup export looks like (seen in reference PDFs):
- Inserted text in the author's colour, underlined.
- Deleted text in the author's colour, struck through, and still laid out:
  it takes space, so lines wrap and pages break differently from the final
  text.
- A change bar in the outside margin next to every changed line.
- One colour per author (measure the palette and its order from references).
- With comments present the page is scaled for the balloon pane, and
  deletions / formatting changes can move into balloons.

Measured on 100 tracked-changes documents against Word's exports: Jaccard
≈ 25, SSIM ≈ 35, page count wrong on 36 — because we draw the final text.

Plan, in order:
1. Inline markup for `w:ins` / `w:del` runs: colour + underline /
   strikethrough, deleted runs kept in layout (`docx/runs.rs` skips `w:del`
   today).
2. Change bars beside changed lines.
3. Per-author colours.
4. Paragraph-level revisions (`w:ins` / `w:del` on paragraph marks and whole
   `w:p` at body level), moves (`w:moveFrom` / `w:moveTo`).
5. Property changes (`w:rPrChange`, `w:pPrChange`, `w:sectPrChange`,
   `w:tblPrChange`): no inline mark, only balloons in Word's balloon view.
6. Deletion / formatting balloons when the document also has comments.

Keep the final view selectable (Word's "No Markup"); which one is the
default is open. Creating redlines (comparing two documents into a
tracked-changes .docx) is a separate tool and out of scope.

## WordArt Remaining Work (LOW IMPACT)

Levels 1-4 are done (flat rendering, text effects, envelope warping, text-on-a-path).

**Level 5 — Legacy VML enhancement (TODO):** VML fill types (gradient/pattern), VML shadow, VML shapetype-to-prstTxWarp mapping. Basic flat rendering already done in Level 1.

## Image Drop Shadow Quality (TODO — LOW IMPACT)

Mostly superseded by "Picture Effects": shadows are a rasterized blur SMask
(`box_blur_3pass`, `embed_shadow`) drawn by `color::draw_image_shadow` at every
site. What remains: header/footer pictures, whose effect XObjects are never
drawn, fall back to the pre-blended rectangle.

## Bullet Line-Height Drift on macOS (DONE — superseded)

Fixed by annotation #66 (2026-09-18): `label_boosted_line_h` now takes the
marker's ascent plus the text's descent, and the vendored Microsoft Symbol
(`fonts/symbol.ttf`) outranks the macOS one (layout fix 43). case33 J 78.5,
SSIM 95.5. Original notes:

case33 annotation #66: bulleted list paragraphs drift ~0.5pt LOWER per bullet vs the
Word reference (text above the list aligns perfectly; drift starts at the first bullet
and accumulates). Root cause precisely identified:

`label_boosted_line_h()` (`src/pdf/mod.rs`) boosts a bullet paragraph's line height to
`max(text_line_h, label_line_h)`, where `label_line_h` uses the bullet label font's
`line_h_ratio` (commit 8691a0d — Word includes the numbering label font in the
tallest-font-on-the-line calc; this fixed under-spacing, +1.5pp case33 / +11.8pp
polish_archery). The bullet font is **Symbol** (`w:numFmt="bullet"`, `w:rFonts ascii="Symbol"`).
On macOS we resolve `/System/Library/Fonts/Symbol.ttf`, whose `usWinAscent=1694`,
`usWinDescent=612` (upm 2048) give `line_h_ratio = 1.126` — anomalously tall (the win
descent is ~0.30em). The Windows Symbol font Word actually used yields ~1.08, so we
over-boost by ~0.046×fs ≈ 0.5pt per bullet.

This is the same class as the bundled-fonts gap: a precise fix needs authentic Windows
Symbol metrics, not the divergent macOS substitute. A hardcoded canonical ratio was
considered but rejected — it overfits and risks regressing `polish_archery_range`
(near threshold at 29.85% Jaccard), and the boost is a deliberately-tuned tradeoff.
Revisit alongside bundled fallback fonts (ship metric-stable Symbol metrics).

## Partially Implemented

- **Tab stops** — left/center/right/decimal work (decimal precision not verified); bar tabs act as left stops and draw no line; `middleDot`/`heavy` leaders draw nothing.
- **Underlines** — only single and double; dotted/dash/wave/thick draw single, underline color is ignored.

## Floating Image Wrapping — Remaining

- **wrapSquare height reserve gated on side-strip width (DONE — 2026-07-02)**: the anchor-paragraph reserve in `render_paragraph_block` now fires only when no usable side strip remains (`MIN_EMPTY_STRIP` = 18pt). With a real strip (brazilian_logistics_study, ~42pt) empty spacer paragraphs absorb through the float's span and the next real paragraph is displaced to `fz.bottom_y` — the old 48pt threshold double-counted the image height there. With no strip (sample500kB, image width == column width) Word stacks everything below, which the reserve reproduces. Note Word actually puts the anchor's own line box below the float too (ref gap 67.7pt vs our 51.8pt on sample500kB p4) — a first displacement attempt lost inter-paragraph gaps; revisit with the paginator.
- **`w:br type="textWrapping" clear="all"` (DONE — 2026-07-02, annotation #111; refined 2026-07-03, annotation #218)**: parsed into `Paragraph.clears_floats`; block loop drops the cursor to the float-zone bottom after such a paragraph. 2026-07-03: the cursor now drops one line height *below* the float bottom — the line following the break (the break paragraph's mark line) still occupies its full line height there, matching Word's ~16pt gap on indonesian_benchmarking_guide p7. Approximation: clear applies after the whole paragraph, not mid-paragraph (fine when the break is alone in its ¶, the common Word idiom).
- **Multiple floats per paragraph (PARTIALLY DONE — 2026-06)**: When one paragraph anchors 2+ wrapping floats (e.g. a logo on each side of a centered title, `pendulum_mechanics_oscillation_lab`), per-line geometry now subtracts every float's exclusion span and places text in the widest gap. Limitation: the page-level `float_zone` for *subsequent* paragraphs still tracks only the first float, so a following paragraph that overlaps only the second float won't wrap around it.
- **Remaining y-shift (page 2 only)**: Word places page 2's image (180x144pt) 14.8pt higher than all other images, despite identical `posOffset=0`. Pages 1,3,4,5,7 match perfectly (delta <0.02pt). Pages 2 and 6 (both cy=1828800/144pt) are the outliers. Likely Word snapping to grid/text boundaries based on image dimensions.
- **Look-back wrapping (DONE — #152, 2026-09-16)**: the paragraph directly before a float's anchor wraps beside it (`pending_float_anchor`); further back still needs a paginator or two-pass layout.
- **Image in text paragraph (DONE — #240, 2026-09-18)**: case41 page 6 now J 93.4.
- **Tight vs Through distinction**: Both currently use convex-hull polygon scanline. For Through wrapping, text should fill polygon concavities. Requires returning per-line interval segments instead of hull bounds. Rare in practice.
- **Word-break precision**: BothSides wrapping produces correct structure but slightly different word breaks from Word, causing ~2pp Jaccard differences on case41.
- **Polygon wrap text distribution**: Case42 (wrapTight + BothSides + complex 53-vertex polygon around Mario) scores ~46% Jaccard. Zone overlap detection is correct but line breaks differ from Word — likely font metric differences for Times New Roman causing different left/right text distribution. Text near concave polygon areas (Mario's arm) appears visually close to the image despite respecting the 9pt distL margin.

## Code Structure

### Duplication & extraction sweep (DONE — 2026-06-21)

Whole-repo over-engineering/duplication audit applied — see `extraction-audit.md`
for the full findings. All 30 verified survivors landed across 11 commits with
zero rendering regressions (208/208 scores unchanged). Highlights:

- **Shared parse helpers** in `docx/mod.rs`: `parse_on_off` (ST_OnOff), `parse_pt`
  (VML/CSS lengths), `is_wml` (namespace predicate, since replaced by roxmltree's
  `has_tag_name` in PR #9), `merge_tab_stops`; plus
  `styles::{parse_font_size, parse_char_spacing, rfonts_ascii_name}` and
  `color::{resolve_dml_color reuse, parse_line_stroke}` now shared instead of
  re-inlined across runs/styles/paragraph/numbering/sections/tables/wordart/etc.
- **`FontEntry::encode`** replaces 5 copies of the char→gid/WinAnsi dispatch.
- **PDF emission helpers** in `pdf/`: `color::box_blur_3pass`, `helpers::draw_circle`,
  `images::{write_jpeg_xobject, write_gray_mask_xobject, write_solid_color_with_gray_mask}`,
  `table::render_table_rows` (nested + header/footer shared the same loop).
- **`render_chart` split** (`pdf/charts.rs`): 714→470 lines; the data-rendering
  match moved verbatim into `draw_chart_series(PlotRect, …)`.
- Dead code removed (`parse_tab_stops`, `FontEntry::char_width_1000_with_fallback`,
  EMF `color_at`, two `SampledBoundary` methods).

Deliberately NOT touched (see audit "leave alone"): the long-but-cohesive
god-functions (`render_paragraph_block`, `parse_table_node`, `render()`), the
generated geometry data tables, and the 3 intentionally-distinct path-command enums.

### Spring cleaning (DONE — PR #9, merged 2026-10-02)

Dead code, shared helpers, one `w:rPr` parser, rustfmt; see
`current_focus/spring_cleaning.md`. `render()` in `pdf/mod.rs` is now ~850
lines; the long one is `render_paragraph_block` (~1650 lines, 20 args).
`PageBuilder` and `LayoutState` exist. Next rounds ("Not done" in that file):
parameter bundling (local branch `worktree-agent-a69249f7b055f3538`),
`PageBuilder` → `Vec<FinishedPage>`, the 4× textbox layout, 5× picture
effects, the triplicated `table.rs` row renderers; the paginator extraction
above remains the bigger goal.

### Image-embedding cleanups (LOW IMPACT — deferred from textbox-image work)

Small consistency / efficiency wins in `pdf/images.rs` that were considered but skipped to avoid scope creep when adding textbox-internal image rendering:

- **Global Arc→pdf-name registry to dedupe XObjects across maps.** Same image data used in body + textbox (or table cell + textbox) is currently embedded as two separate PDF XObjects because each map (`inline_image_pdf_names`, `table_cell_image_names`, `textbox_image_names`, …) keys independently by `Arc::as_ptr`. A single global registry would let the second site reuse the first XObject. Wasteful in theory, accepted limitation in practice.
- **`build_paragraph_lines` / `build_tabbed_line` should accept `&HashMap<usize, &str>`.** The current `&HashMap<usize, String>` signature forces every caller (body, header/footer, table cell, textbox) to `.clone()` pdf names into a fresh per-paragraph map. Borrowing would eliminate the clones, but ripples through `pdf/layout.rs` and every caller.
- **Pair `image_names` + `effect_names` into a struct.** Every embedder (`hf_*`, `table_*`, `textbox_*`) threads the two maps as separate `&mut HashMap<…>` parameters. Pairing them would shrink signatures throughout `pdf/images.rs` and `pdf/textbox_render.rs`, but only worth doing alongside the dedup registry above (otherwise diverges from the established style without enough payoff).

## Performance

### Known Bottlenecks

- **Repeated WinAnsi conversion** — same text is converted in line-building, rendering, and table auto-fit. Pre-compute once and store in `WordChunk`.
- **String allocations** — the allocating `font_key()` is still used at 9 sites (an allocation-free `font_key_buf` exists); `WordChunk` clones font name strings per word. Use indices or interning.
- (Font scanning is memory-mapped and disk-cached, so the old "double font reads" item is moot.)

### Parallelism (rayon)

Not started: rayon is only a dev-dependency (tests and `page-metrics`); the library is single-threaded. Candidates:

- Font directory scanning — embarrassingly parallel, biggest win
- Font metric computation — parse face, compute widths per font independently
- Paragraph line wrapping — independent per paragraph once font metrics are ready
- ZIP decompression + XML parsing — read all entries into memory, parse in parallel

### Other

- Memory usage for large DOCX files with many images

## Scraped Fixture Status

As of 2026-10-02, 39 of 138 scraped fixtures fall below the Jaccard 20.5% or SSIM 75% labels. Run `./tools/target/debug/analyze-fixtures --failing` for the current breakdown.

## Test Harness: Surface Conversion Panics Loudly (TODO — HIGH PRIORITY, found 2026-06)

A library panic went unnoticed for an unknown number of runs: `scraped/construction_bathroom_accessories_spec` panicked in `cell_span_width` on every conversion, but the suite still reported "134 passed" with exit code 0. Three gaps compounded:

1. `tests/visual_comparison.rs` catches per-case panics and emits `[SKIP] <case>: conversion panicked` — visible only in `--verbose` output; the case silently gets no score, so the compact report's "N scored, N unchanged" looks green.
2. `run-tests.sh` greps `thread.*panicked` into a "Panics:" section, but the exit code stays 0 — nothing fails.
3. Conversion worker threads are unnamed, so panic messages show `thread '<unnamed>' panicked at src/...` with no case attribution — diagnosing required a separate verbose run.

Fixes:
- `run-tests.sh`: exit non-zero when the Panics section is non-empty.
- Harness/compact report: count panicked cases as failures and list them by name (`PANIC: scraped/construction_bat..`) in the compact output.
- Name conversion threads after the case (`std::thread::Builder::new().name(case.clone())`) so panic messages self-identify.

## Test Corpus Expansion

case50–53 (deep style inheritance, nested tables, stacked bar charts, extreme
chart data) have their reference PDFs. Only case58 (3D effects) and case65 still
lack one.

## Single-row nested AutoFit reserved fields

Keep explicit grid/cell preferences for a single-row nested auto-width table when it fits its parent. Empty number/date slots must not collapse to the current text width. Restrict the change to explicit grids and unmerged preferred-width cells; inferred, overflowing and multi-row tables keep their existing AutoFit path. case88 (empty slots) and case89 (filled slots) reproduce the issue with synthetic documents and Word reference PDFs.
## Explicit widths in single-row nested tables

An explicitly sized single-row nested table should use the fixed-width base when its cell preferences and grid totals agree and fit the parent. Preserve min-content redistribution for narrow marker cells instead of collapsing all columns to text. Synthetic case90/case91 cover empty and populated fields. Global nested-table origin/float positioning is a separate issue.
### Empty continuous sections

Preserve the line and spacing of an empty section-break paragraph when it is the first block of its section. This also applies to continuous sections between adjacent section breaks. Keep zero-height handling for empty break paragraphs following content within the same section.
### Saved asymmetric AutoFit grids with nested tables

Keep a saved non-uniform AutoFit grid containing directly nested tables rather than rebuilding it from stale tcW preferences and proportionally squeezing the host column. Uniform grids retain content-based sizing; content minimums and available-width limits still apply.
### Empty cell paragraphs with anchored drawings

Rendering now consumes the paragraph-mark line height already counted by row layout when an otherwise empty paragraph contains floating content. This keeps following text in header cells aligned with cells containing plain empty paragraphs.
- Preserve asymmetric saved AutoFit grids when uniform oversized cell preferences would erase them; signature-title wrapping covered by cases109–112 (including nearly equal saved columns).
- Honor tblOverlap=never for colliding floating tables, retaining the preceding aligned float zone for following text (cases113–117). Empty wrapping frames after this table remain outside the body flow; ordinary empty paragraphs and line breaks remain in flow.
- Reserve the actual wrapped-line height of image-only running heads with multiple inline pictures (cases121–123); fitting picture lines retain their existing height.
- Keep non-wrapping body pictures anchored to the margins of the current sheet across a mid-page continuous section margin change; new-page and unchanged-margin controls retain their placement (cases124–126).
- Honor preceding paragraph spacing and vertical placement of single-row text-anchored nested floating tables in top-aligned cells with no following visible paragraph content. Preserve floating extent in row height, field visibility and page continuation (cases131–138). Clamp negative offsets for the first table at the top of its cell. Horizontal positioning and nested float overlap/wrapping remain separate work.

- Resolve automatically left-placed nested text floats against the occupied cell area. Empty anchor marks stay beside a float when the free right strip is at least 18.75 pt, or clear below it; following intersecting floats move below and preserve their text distance. Synthetic Word transition controls at 18.75/18.70 pt, different mark sizes, border/overlap controls and an inline predecessor (cases139–151). Reanchor dependent collisions after a page break in drawing, fragment height and break selection. Controls use fixed grids and plain markers to isolate placement from pending AutoFit/field/tab fixes. This builds on the nested vertical-flow change; explicit/page/column horizontal anchors remain separate.

- Preserve leftFromText when a non-overlapping body table is pushed below a float, both during pre-layout clearance and geometric collision. Compat 15 offset positions are clamped to the text-area edge plus that distance; larger explicit offsets, page anchors and compat 14 keep their prior positions (cases152–161). This builds on the non-overlap stacking change.
- Honor negative paragraph indents in cell minimum-content widths, preventing an AutoFit nested registration table from collapsing reserved fields (cases118–120). Positive-indent sizing and the separate zero-indent reserved-slot fallback are unchanged.

- Fix line overlap with floating table zones: check line-box intersection at the top edge, preventing short anchor paragraphs from printing over a table (existing cases152–161).
- Keep empty auto-height wrapping frames outside body flow and preserve both lines of an empty clearing paragraph below a full-width floating table; narrow-table and no-frame controls included (cases127–130).
### List label and following tab geometry

Use the same first-line text calculation for body paragraphs and table cells. First-line indents move the label as well; following text advances to the next available tab after the label instead of overlapping it.
