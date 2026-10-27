# Roadmap

Open work only. Finished rounds, their rules and evidence live in git history
(`git log -p roadmap.md`), the commit messages and `current_focus/`.

## Open fixture leads

**New fixtures round 2 (2026-10-23):**
- **swiss_materials (33.3)**: Word keeps a 3-line list item at the foot of
  page 2 (last baseline 783.5, its 2pt descender past the 785.3 margin) where
  we break the page before it; our fit test is 0.46pt over. Letting the last
  line's descent hang (`last_line_lead = line_h - ascent`) regressed 27
  fixtures and was reverted; Word's exact tolerance (its 0.25pt line grid
  makes the item 3 x 11.0 = exactly the margin) is unknown. The rest of the
  gap is the page-3 reflow; its `›` bullet sits 3pt low.
- **bulgarian_farmland (71.5)**: the letterhead's first line is not wrapped
  beside the emblem, which is anchored to the *next* paragraph with a negative
  offset (header float zones only start at their anchor).
- **japanese_hr_consultation (39.9)**: table rows are ~1pt short each
  (HG丸ｺﾞｼｯｸM-PRO 9pt bold cells on a 17.5pt lines grid: Word rows
  18.25/19.0, ours 18.0), so the form drifts 9pt by the table's end; needs a
  row-height probe.
- spanish_constitutional_court (59.1, 18 pages), catalan, lithuanian,
  czech_council (59-67): not inspected.

**New fixtures round 1 (2026-10-22):**
- **advocacy_initiative_newsletter_2013 (36.7)**: our nested picture table is
  1.4pt narrower than Word's grid, so the stories wrap a word earlier or later
  and pages 2-3 end 24pt off; the standard `o:hr` draws 1.5pt, Word 0.5. Word
  starts the text ~91pt right of the picture table's edge for both 1755tw
  grids, we start at 89.5-90 (table width + rightFromText).
- **cote_dazur (68.3)**: body text does not wrap around header floats (Word
  does, per line); only tables are pushed.

**Bugs exposed by the external PRs (2026-10-09):**
- **brazilian_logistics p10**: "Fonte: ALARCOM, (2019)." is laid out one
  letter per line from y≈108 (Word: one line at y≈346); the paragraph is
  squeezed to almost no width.
- **chinese_costume (+1 page)**: Word widens Latin–CJK boundaries ("5-10 分钟",
  "MP4 封装"); we don't, so CJK lines in table cells break differently.
- **covid_insomnia p5 right column**: the reference list sits one line low
  ("[10]" at 146.3pt, Word 135.3pt). Its two-column flow also regressed with
  the justified squeeze.
- **Two thresholds for one rule**: `MIN_EMPTY_STRIP` (18.0pt, `pdf/mod.rs`,
  bracketed only between 0 and ~42pt) and `MIN_NESTED_FLOAT_STRIP` (18.75pt,
  `pdf/table_layout.rs`, measured between 18.70 and 18.75pt) both decide
  whether an empty line fits beside a float. Probe the body case and share
  one constant.
- **case111**: "…ора по" overflows its cell (now clipped) where Word fits it;
  re-check after `4c47ec52`.
- **PR #44 only before a continuous section**: the section-break line kept
  after a table is measured only when the next section is continuous; before
  a new page Word's behaviour is unprobed.

**WordCraft head-to-head leftovers (2026-10-13):**
- **Nested floating tables in a cell** (case145/148): without a paragraph
  between two floating tables Word stacks the second below the first
  (`tblOverlap="never"`, cell 13.75pt taller); we let them overlap. Cell
  bottom borders start 0.7pt left of the text in Word (tblCellMar left=0).
- **case113/114**: "Body marker" 1.3pt left of Word's.
- **Documents without styles.xml**: our built-in defaults match desktop Word,
  not online (case135's cell line: desktop 633.22, us 633.20, online 633.70).
  Probe Word online's defaults for a docx without styles.xml.

**Older fixture leads:**
- dutch_government: a page-anchored floating table moves the following body
  table down 2.4pt in Word (cause unknown); two 1pt `in-table` paragraphs come
  out 0.7pt short; a borderless one-row table ~2pt short. Its reference came
  from a Word-repaired copy.
- estonian: per-line drift.
- maine: nested layout tables.
- zimbabwe_gold p5 keeps a line whose lead reaches 1.12pt into the footnote
  area (unexplained); page 3 fits one more two-line double-spaced paragraph +
  footnote than Word.
- streamnet p1 agenda table rows +0.43/+0.68pt each (cell line height).
- dental_amalgam: gained +22pp J with a max-descent marker rule but +10pp with
  the measured ascent-only rule; check what its markers are.
- ut_koer (2026-07, unverified since): a tab-stop two-column contact header
  beside a wrapTight header image needs tab layout aware of the float band
  (`build_tabbed_line` has no float-zone awareness).
- multi_font: needs Copperplate Gothic Light (in Word's cloud catalog; pick it
  in Word's font menu to download); page 1 also runs ~14pt taller than
  Word's. korean_japanese needs HY헤드라인M and New Gulim (low value).
- case37's reference shows five black boxes where its shapes are: re-export.
- case58 (3D effects) and case65 lack a reference.
- The first 10 fixtures of the 2026-10-03 round were exported as
  `<stem>_<hex>.docx`, so FILENAME fields print that name (massachusetts
  footer); tool fixed, refs not re-exported.
- case77's on-page label claims thaiDistribute gets distribute's treatment;
  Word disproved it. Fixing it means regenerating DOCX and reference.

Open queue from the large-corpus and focus-fixture rounds:
`current_focus/layout-accuracy.md` §5.

## Line and paragraph layout

- **0.25pt baseline grid**: Word quantizes baselines to 0.25pt online (0.24
  local). Residuals it explains: greensboro (Arial 11pt at 1.5 lines: ours
  18.97, Word 19.0), romanian_candidate_declaration (borders drift 0.1-0.3pt
  per row, SSIM 91.5), case69 (last line 15.00 vs 14.83 under
  `vAlign=center`), case33's list lines (16.115 vs 16.0).
- **Each space character is a justification stretch point**: polish_municipal's
  double space is 7.44pt in Word = 2 × (3.0 + 0.72 slack); we add one slot for
  the gap. Find justified lines with consecutive spaces across fixtures first.
- **Mixed East Asian + Latin line height**: usep_handbook's checkbox lines (☐
  in MS Gothic, text in Calibri) are 1.53pt taller than Word's: we put
  Calibri's win descent under MS Gothic's ascent, which already carries the
  1.3× leading. Candidate: max(Latin ascent + descent, each East Asian run's
  1.3× box); verify on case79 and the CJK fixtures.
- **Per-line heights in headers** (japanese_land: 10.5pt CJK header lines pitch
  17.67 in ours, 15.12 in Word). Header/footer paragraphs are also measured
  twice (`compute_header_height` without line building, and
  `render_header_footer`); `layout::position_stretch` is a one-line estimate.
- **Title → table transition** in the conference forms: a 7pt shortfall above
  the table (measure on east_asia_conference_form; both still fit one page
  where Word has two).
- **NBSP stretch before compat 15** (tried, reverted): Word widens NBSPs with
  word spaces before compat 15, but not after one-letter words ("k
  dopracování" in czech_expert) and with uneven shares. Stretching all moved
  croatian +1.2 J, russian_construction −7.4, czech_expert −5.2. Needs a probe
  of NBSPs after one-letter words and runs of NBSPs.
- docGrid: East Asian lines sit 0.17pt lower in Word on average (case79, 17
  lines). Multiples below 1 and at-least spacing on a grid are untested.
- A whitespace-only run between bold and italic runs (slovak_constitution)
  does not set the line's ascent in Word; whether it counts at all is open.
- Small caps: a space at a run boundary only sees its own run.
- Compat-15 left tab stops past the margin: Word puts the tab on a line of its
  own and the text at the margin below (probes P1, P6, P7); we wrap to the far
  stop.
- Tab stops: decimal precision unverified; bar tabs act as left stops and draw
  no line; `middleDot`/`heavy` leaders draw nothing.
- Underlines: dotted/dash/wave/thick draw single; underline colour ignored.
- `w:textAlignment`: only top is laid out; center, baseline, bottom and auto
  behave alike.
- `w:mirrorMargins`: parsed (used only for odd/even filler pages), never
  applied. Swap left/right margins (and the gutter side) on even pages.
- Below compat 15 without `overrideTableStyleFontSizeAndJustification`, a
  table style's font size beats Normal's (unless it is 10pt); we always let
  Normal win.
- Below single spacing (auto m < 1): header/footer, textbox and footnote
  paragraphs keep the old rule (no probe, no fixture).
- Line breaking: a hyphen at the start of a word (" -5") also breaks before
  the digit; scan the references for `\s-$` line ends. Two adjacent
  `w:noBreakHyphen` (parsed to `-`) become a break opportunity; map them to
  U+2011 if a fixture shows it.
- Distributed alignment: Latin lines with spaces spread across one gap too few
  (Word counts spaces as distributable characters; see the `ponytail:` note on
  `char_justify_gaps`). Kashida justification changes Word's line breaks
  (needs kashida metrics). thaiDistribute on Thai text is untested.
- `eastAsiaTheme="minorHAnsi"` (997 runs in the corpus) is ignored by
  `resolve_east_asia_font`.
- Missing fonts: names with `;` (Word takes "Open Sans;Arial" literally), and
  "Open Sans" → Segoe UI (`current_focus/merge-regressions.md` item 5).
  Courier → Courier New is Windows' substitute, unverified for Mac Word.

## Tables

- **Autofit vs Word** (2026-10-05 x-drift survey): case45 (tblW 4000, tcW 4680
  per cell) keeps Word's saved grid 2091/1909, we split 100/100, and its
  tblpXSpec="right" copy sits one cell margin past the right margin in Word;
  case61's 200-twip column is widened to fit its text by Word, and its table 2
  first column is 9.6pt too wide in ours; case6 and case52/53 whole columns
  shifted 7–11pt; air_pollution's dot lines fit 42 dots in Word, 41 in ours.
- Some rows run slightly short (slovak_eu_directive ~0.03pt per row); a
  vMerge/colspan test table gets a far too narrow column.
- Nested tables use neither compat rule for the outdent
  (`render_nested_table`); needs a compat 14/15 probe. A nested header table
  (logo beside a title table) can sit 5–7pt low.
- Single-row nested tables: global nested-table origin/float positioning,
  horizontal positioning and explicit/page/column horizontal anchors of
  nested floats remain.
- The must-split path (row taller than a page) tests only `!cant_split`, so an
  exact row taller than a page splits; Word likely clips it. Unprobed.
- Column auto-fit measures cells without footnote/endnote marks.
- `tcW type="pct"` is read as twips (harmless while grid widths dominate).
  `tblW` vs oversized `gridCol` is unclamped (no fixture triggers it).
- Table cells drop run-level inline pictures (`EMPTY_INLINE_IMAGE_MAP`), and
  cell image paragraphs don't follow the short-picture rule.
- Vertical CJK cell text: `render_vertical_cjk_cell` sizes by chars ×
  font_size, ignoring paragraph spacing (misaligned in merged cells).
- Labels hanging outside a cell (probe Word first).

## Floats, frames and wrapping

- Bands are checked at block start (a paragraph running into one part way
  keeps its lines) and live beside `FloatZone`; one mechanism with a
  full-width flag would be cleaner. Textbox `reserve` still uses "width ≥
  half the column" where images and bands use the side-strip test.
- Text-anchored body frames, and wrap-around frames reaching into the text
  (croatian_grant's two), still flow inline.
- `content_h` in `render_paragraph_block` includes the room reserved for a
  topAndBottom float below the lines; borders still use it (shading doesn't).
  Unverified against Word.
- Body/header flow reserves `height_pt` for TopAndBottom textboxes while
  rendering uses `textbox_height`.
- Word puts the anchor's own line box below a no-strip float (sample500kB p4:
  Word gap 67.7pt, ours 51.8pt).
- Multiple floats per paragraph: the page-level zone for *following*
  paragraphs tracks only the first float.
- pendulum: Word places the 144pt-tall images on pages 2 and 6 14.8pt higher
  than the others despite identical `posOffset=0`.
- Look-back wrapping beyond the paragraph directly before an anchor needs a
  paginator or two-pass layout. `resolve_fi_y_top` should treat
  AlignTop-relative-to-paragraph as the anchor top (then the look-ahead can
  drop its paragraph-relative filter).
- Tight vs Through: both use the convex hull; Through should fill concavities.
  case42's 53-vertex polygon wrap breaks lines differently from Word.
- Off-page `posOffset` (learning_cultures rId14: 334pt-wide image placed past
  the page edge): Word clamps or reflows; a naive right-edge clamp regressed
  stem_partnerships.
- Truncated PNGs fall back to a 1x1 placeholder; Word renders them partially
  (learning_cultures image1.png).
- behindDoc shapes from later paragraphs can paint over earlier text (needs a
  pre-pass). Group-level flipH/flipV/rot are ignored except for linear
  connectors; image/text/arc mirroring in groups is open. Canvas/group inside
  paragraph-level `mc:AlternateContent` grabs only the first wsp.
- Wrapped running-head pictures: tall final pictures and three-picture
  layouts keep residual discrepancies.
- `wp:effectExtent` is not part of an inline picture's line height in body
  text (ut_koer's 0.75pt residual).
- Anchored SmartArt with wrapping other than wrapNone is laid out as inline
  (no fixture).

## Columns and sections

- Footnote continuation across columns (case82 J 21.3): Word may continue a
  note under the next column so more body text fits; needs a rule for how
  much must stay with its reference.
- case80's 4-column region: Word's separator runs one line (15pt) below
  column 1's last space after; unexplained. The table probe's region ends
  1.6pt off. Each balancing trial lays the region out again.

## Footnotes and endnotes

- Footnotes that don't fit: Word splits a note across pages with a
  continuation separator (uk_commercial p17/18); we move the reference line.
- Notes use the hand-built `parse_notes_simple`, not `build_paragraph`
  (borders, shading, `style_id`, contextual spacing missing); note tables are
  flattened to one paragraph per row (`Footnote { paragraphs }` should become
  blocks). Footnote paragraphs place exact and at-least lines at the plain
  ascent (`footnotes.rs`).
- Inline endnotes are not paginated (`render_endnotes_inline` ponytail;
  erasmus' endnote block runs past page 3's margin); `endnotePr
  pos="sectEnd"` unsupported; endnotes.xml's separator not used (fixed 0.5pt
  rule 12pt below the body); a separator run's own rPr ignored.

## Tracked changes and comments

- Revision markup (insert/delete colours, change bars, per-author colours,
  moves, property-change balloons) is on branch `tracked-changes`, not merged;
  final view stays the default. door_air (#259/#260) is the fixture whose
  reference shows markup. Deleted paragraph marks (`w:pPr/w:rPr/w:del`) don't
  merge paragraphs.
- Comment balloons: font (7pt) and line height (8.5pt) are case63's zoomed
  values.

## Pictures, shapes, charts

- Picture effects: 3D (`a:scene3d`, `a:sp3d`; case58), preset shadows
  (`a:prstShdw`), theme colours in effects (inner shadow `a:schemeClr` falls
  back to black), header/footer pictures (effect XObjects embedded but never
  drawn), effects on shapes and text boxes (pictures only). Soft-edge and
  reflection masks are built on the uncropped source; SmartArt pictures use
  their own draw path. Block pictures (`para.image`) ignore rotation.
- EMF: clipping regions and opaque text backgrounds; immediate-mode segments
  and mid-path SaveDc/RestoreDc; mixed vector+bitmap EMFs.
- SmartArt: `dsp:grpSp` groups, `dsp:cxnSp` connectors; the layout engine for
  files without the `dsp:drawing` fallback (not planned).
- Charts: 3D charts (pie3D drawn flat), stock, combo (only the first type),
  data labels, titles, secondary axes; small axis-label and legend offsets.
- WordArt level 5: VML gradient/pattern fills, VML shadow, shapetype →
  prstTxWarp mapping.

## RTL / BiDi

GPOS mark anchors (harakat off-centre); mirrored brackets extract as the
mirrored character; Arabic in a face without Arabic glyphs draws `.notdef`
(Word rescues from another font); no joining across runs that differ in
formatting; lines with tabs or float regions keep logical order; tab stops
and `bidiVisual` not mirrored; Arabic letters in a plain run whose font has
Arabic glyphs should use that font (Word: Courier New), we use the cs font.

## Office Math

Each needs a Word online probe first (corpus refs are print exports); rules in
`current_focus/office-math.md`.
1. `m:wrapRight`: continuation lines of a wrapped display equation flush
   right.
2. Display ∫ limit positions and `m:borderBox` padding.
3. Inline radicals over tall contents: Word's surd sits 0.3–0.6pt lower than
   the half-surplus rule at every size.
4. An `m:nor` superscript sits 0.6pt closer to its base than Word's (e^max).
5. Placeholders: an inline n-ary with placeholder limits stands 2.6pt too
   low; a function with an empty argument is 2pt short; a sup over an empty
   sub sits 0.5pt low.

## Accessibility

Backlog (details in `current_focus/A11Y.md`):
1. Math speech reads matrices, accents and equation arrays as their contents
   in order.
2. Link rects and outline destinations ignore the comment-pane zoom and
   vAlign (no fixture has links with either).
3. Wingdings 2 and Webdings extract as private-use code points; `w:lang/@bidi`
   ignored.
4. Missing glyphs other than spaces draw `.notdef`: Word rescues them per
   character from another font; widen the CJK rescue (`missing_cjk_chars`) to
   non-CJK characters. The space fallback lays U+2002/2003/2009/202F out at
   U+0020's width.
5. `anchor_seq` counts per parse pass: textboxes from paragraph-level
   `mc:Choice` sort after every run-level anchor.

## Other features

- Field codes: only PAGE, NUMPAGES, STYLEREF, PAGEREF and IF are evaluated;
  MACROBUTTON shows its display text or picture; the rest print their cached
  result.
- Body PAGE/NUMPAGES (issue #140) come from a first render pass: a PAGE in a
  paragraph that spans two pages shows the first page's number, and fields in
  table cells keep their cached result.
- Line numbering ignores `w:restart`: the counter always runs on, never
  restarting per page or section.
- altChunk HTML: `ul`/`ol`/`li` are dropped by the converter.
- `tblGrid`-less column inference (`row_widths`) doesn't add `gridBefore`
  columns (Word always writes a tblGrid).

## Parked

- **Online glyph positioning / width drift** (`current_focus/online-references.md`
  §4a): Word's per-glyph corrections are not derivable from exact widths; an
  empirical 0.07% in-run shrink gives +0.6 J. Fit or more probing is the
  user's call.
- **Hyphenation**: Word's online converter never hyphenates; the 8
  `hyphenation/` fixtures wait for desktop-Word references with proofing
  tools.

## Fonts

- Bundle metric-compatible open fonts behind a feature flag (Liberation
  Sans/Serif/Mono, Carlito, Caladea) for servers and Docker; today only
  installed fonts and the assets repo count (Carlito/Caladea are used when
  installed).

## Architecture and code structure

- **Paginator extraction**: a pass producing `Vec<Page>` of positioned
  elements, so widow/keep logic leaves the render loop, look-back wrapping
  becomes possible and PAGE/NUMPAGES resolve after layout.
- Cell and footnote paragraphs bypass `paragraph::build_paragraph`
  (`docx/tables.rs`, `parse_notes_simple`): cells default most pPr fields
  (keep_next, widow_control, borders, shading, outline level) and never get
  auto spacing; one shared pPr resolver would close it.
- One per-line height model shared by body, cells and headers; one
  widow/orphan split helper for body split, keep-with-next and cell splits;
  run-boundary breaks from UAX #14 over the paragraph's joined text (a URL
  split across runs can still break after `/`); one table-left-x function
  for body, header and nested tables; one new-page helper.
- One picture model: `docx/paragraph.rs` hoists a lone inline picture into
  `Paragraph::image` while two or more stay in runs; one shared
  `draw_embedded_image` for the five block-picture draw copies (rotation +
  effects); a shared `decode_raster` so `a:lum`, crop and soft-edge apply
  once and reach `embed_reflection`.
- Floating-table keep-together could be geometric (`fp.y > saved`) instead of
  reading raw `tblpY`.
- Spring-cleaning next rounds (`current_focus/spring_cleaning.md`): parameter
  bundling, `PageBuilder` → `Vec<FinishedPage>`, the 4× textbox layout, 5×
  picture effects, the triplicated `table.rs` row renderers.
- `pdf/images.rs`: one global image registry to dedupe XObjects across maps;
  `build_paragraph_lines` taking `&HashMap<usize, &str>`; pair `image_names` +
  `effect_names`.

## Performance

- Image recompression is the converter hotspot (case59: ~5.1 s in
  `embed_image_xobject`). Cheapest first: decode/resize in parallel then write
  XObjects in order; faster resampling (`fast_image_resize` or a cheaper
  filter for large downscales); scaled JPEG decode; pass JPEG bytes through as
  DCTDecode when nothing changes.
- Repeated WinAnsi conversion (line building, rendering, autofit);
  allocating `font_key()` at 9 sites; `WordChunk` clones font names.
- Library is single-threaded: font scanning, metrics, per-paragraph wrapping
  and XML parsing are candidates for parallelism.
- Column balancing re-lays the region per trial.

## Test harness and tooling

- **Surface conversion panics** (high priority): `run-tests.sh` prints a
  Panics section but exits 0; panicked cases get no score and look green;
  conversion threads are unnamed, so panics don't name the case. Exit
  non-zero, list panicked cases by name, name threads after the case.
- Deterministic output regressed once: scraped/door_air_cooling_unit_spec gave
  two different PDFs over three runs (2026-10-03); not investigated.
- Engine-comparison site: if runs grow, batch veraPDF into one call per shard
  via a `page-metrics` pre-pass. `engine_compare.py` `pdf_creator()` truncates
  Quartz producers at the escaped paren.
- SSIM with ±8px search (~17 CPU-s for a 205-page fixture): rayon over the
  window search, integral images, SIMD, if turnaround becomes a bottleneck.
