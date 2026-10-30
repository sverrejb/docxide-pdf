# Roadmap

Open work only. Finished rounds, their rules and evidence live in git history
(`git log -p roadmap.md`), the commit messages and `current_focus/`.

## Open fixture leads

**torture_test round (2026-10-29/30, branch `torture-test`):** a two-page
fixture built only from features absent from `src/` (kit with generator in
`tests/output/torture_test/`; Word online reference). 2.8% J / 26.7% SSIM
(3 pages against 2) → 67.2% J / 81.2% SSIM (2 pages); the picture-line rule
for charts also lifted case52/case31/case29/case53 by 3–8pp. Done on the
branch: linked text boxes (whole paragraphs; Word splits between lines),
drop caps, ruby, fitText/tcFitText, bidiVisual, suppressLineNumbers,
wordWrap=0, numbers on lines without text, picture grayscale/duotone/
alphaModFix, picture-filled inline shapes (tiled), shape glow/reflection
(vector approximations), chart title/frame/value labels/data table. Open:
- **3-D charts** (`c:bar3DChart`, `c:view3D`) draw flat. Word rasterises the
  plot and its back wall sits (dx, dy) behind the front face, which shrinks
  the front: rotX 15/rotY 20/depth 100 on a 432pt frame gives front x
  159–459.8 and dx≈21.5, dy≈16.7; rotY 40 dx≈38.8 (dy same); depthPercent 200
  dx≈36.3, dy≈27.8; rotX 30 dx≈19.4, dy≈27.9; a 300pt frame dx≈15.1. Probes
  and `measure.py` in `tests/output/tt/probe/` (worktree). Raster extents
  include margins: measure the wall from the raster pixels before fitting.
  Word also picks axis max 25 for a 24 maximum where we pick 30.
- Inline `wps` shapes with text, charts and SmartArt are block boxes at the
  paragraph start, not line objects (a picture-filled shape is now a
  picture); an inline-object run kind would fix all three.
- The empty line before a page-break run gets no margin number ("19").
- Value labels only on vertical clustered bars; titles not on pie/radar;
  shape soft edge absent; reflection mirrors the fill, not the text.

**Weakest-fixtures round (2026-10-27, branch `weak-fixtures`):**
- **compressPunctuation needs probes**: chinese_costume's left-aligned cell
  keeps "…使用。" on its line with "，" and the final "。" each half an em
  narrower (4.5pt at 9pt), yet wraps "学" rather than squeeze the five marks
  before it; taiwanese_education_fraud_ruling (mostly left-aligned) squeezes
  for ordinary characters. Three rules failed on the suite: hanging the mark
  past the margin, compressing only for a word ending in a mark (taiwanese
  -23), and that only outside justified lines (taiwanese -23). Probe Word
  online with left/justified paragraphs, a plain character versus a closing
  mark at the line end, and a quarter versus half em. This is
  chinese_costume's extra page.
- **japanese_land_development (35.6)**: Word puts no gap between "、" and a
  digit ("2、3", 10.5pt step), which our autospace rule now matches. But
  Word's text after "　2　" plus a plain half-width space starts 2.6pt
  (a quarter em) later than ours. The old rule's wrong gap after "、" hid
  this, so the note wraps one character late. Probe where Word's gap sits:
  "　" + space, or space + digit.
- **sample500kB (12.8)**: Word online draws "DejaVu Sans" in Sylfaen and
  "Open Sans" in Segoe UI, by name: charset 00/01, family roman/swiss/auto and
  removing Open Sans' altName Arial change nothing (probes 2026-10-27). Needs
  Sylfaen and Segoe UI in the assets; DejaVu Sans in ~/Library/Fonts wins
  locally anyway. Its page-4 picture now sits at its anchor paragraph's top
  as in Word, but lower than Word's because our text above runs seven lines
  longer.
- **uk_commercial_lease (27.6, 78 pages vs 79)**: on page 64 we fit 4 lines
  of 2.1.6 above the notes where Word fits 2 (our separator 736.83, Word's
  737.23), so a blank page 67 goes missing; our fit test is more lenient
  there. Page 11's notes block is 0.67pt short from per-line differences of
  0.05-0.25pt (the 0.25pt grid below); the "bullet lines 9.2 vs 9.75" note
  does not reproduce (Word 9.25, ours 9.20).
- **air_pollution (45.5)**: its accents are stored decomposed ("t" + U+030C);
  Word draws the precomposed glyph ("ť", 0.53 em in Times New Roman Bold
  Italic) and maps that one glyph back to the decomposed pair in ToUnicode.
  Composing at parse time gained +1.0 J but lost 47 letters from our
  structure tree; an ActualText span instead garbled the extracted text and
  added 3 PDF/UA failures (both reverted). Needs a glyph-to-several-chars
  ToUnicode entry. The dot leader after "Katastrálne územie: " (a tab
  after a w:br) is missing; leaders start at x 68.8, Word 66.0.
- **japanese_hr (41.4)**: sub-rows are 19.0pt where the merged left cell's
  continuation carries an invisible sz=12 border, 18.25 elsewhere (ours
  18.0); Word puts the wrapNone box's empty anchor paragraph below the table
  (all six 10pt paragraphs follow it); `w:spacing w:val="-10"` in 7pt Meiryo
  draws -1.0pt per character in Word (-1 Tc), ours -0.5;
  `balanceSingleByteDoubleByteWidth` is unimplemented (Meiryo spaces 0.507
  em, ours 0.332; 10 fixtures set it); a paragraph bottom border ignores
  `w:ind w:right="-903"`; the banner title sits 8.8pt high.
- **swiss_materials (32.6)**: besides the page-2 break below, the VML
  watermark draws ~24% heavier and ~4pt left of Word's; the "Helvetica 55
  Roman" bold headings (altName Bahnschrift Light) draw regular-weight in
  Word; page 6's table bottom border is solid grey in Word, dotted in ours.

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

- **compressPunctuation** (issue #139, probed through Word online 2026-10-27,
  rules in `pdf/layout.rs` `PUNCT_SQUEEZE`/`PUNCT_PULL_FREE`): every mark on
  the line gives the same cut; compat 15 squeezes only justified paragraphs,
  ½ em per mark, and pulls the word only while 1.2% of the measure is free
  before it; compat 14 squeezes any alignment, 0.4 em, pulls at ½ em free
  including the autoSpaceDE gap. Japanese halves a closing mark before
  another mark regardless of fit. Open: zh-TW centred marks lose the cut on
  both sides (`locl` glyphs; taiwanese_education 二、 sits 1pt left of ours);
  autoSpaceDE gaps are lost at run boundaries unless the runs merge
  (chinese_costume "5-10 分钟").
- **0.25pt baseline grid**: Word quantizes baselines to 0.25pt online (0.24
  local). Residuals it explains: greensboro (Arial 11pt at 1.5 lines: ours
  18.97, Word 19.0), romanian_candidate_declaration (borders drift 0.1-0.3pt
  per row, SSIM 91.5), case69 (last line 15.00 vs 14.83 under
  `vAlign=center`), case33's list lines (16.115 vs 16.0), swiss_materials'
  page-2 break (0.46pt over: our pitches 11.47/10.92, Word 11.50/11.00),
  uk_commercial page 11's notes block (0.67pt short).
- **Mixed Latin/East Asian justification** (issue #142, Word online probes
  2026-10-28, `MixedJustify` in `pdf/layout.rs`): the slack spreads over every
  space and the gap after every East Asian character alike, each space +1/9
  em, Latin letters never stretch, the line's last character has no gap.
  Open: a Latin letter glued to a CJK character ("abc一二三def") measures the
  gap plus ~5pt (autoSpaceDE and something else) on both sides; Word's
  squeezed spaces come in 0.25pt steps and it let the repro's final period
  hang 2.9pt where we squeeze the spaces fully; `Distribute` keeps the old
  per-character spread.
- **`w:ind` character units** (issue #143, probes 2026-10-28,
  `apply_indent_chars`): chars beat twips at any level, an explicit 0 switches
  them off, one character is the first run's size plus the section grid's
  `charSpace`/4096, left + hanging add up in chars mode. Numbering-level
  `*Chars` are not read; headers and footnotes use no grid delta.
- **Trailing punctuation**: Word never lets a Latin mark hang past the margin
  (`overflowPunct` is not what it does). In a zh/ja/ko document whose East
  Asian font is a CJK face, a justified line's word-share limit grows by half
  a final mark (`punct_half_out`; the 25% space cap never widens). Probed with
  `.` only; `, ; : ! ?` assumed. Issue #142's "introduction." still wraps: with
  Word's glyph advances the line sits 0.4pt inside the cap, with ours 0.4pt
  outside — the glyph device-grid work decides it.
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

- **Online glyph positioning** (merged 2026-10-30, `src/pdf/glyph_grid.rs`,
  `current_focus/fundamentals.md` §2, probes 1–5): a 288 dpi device grid,
  device advances rounding up from 0.46 px (half-up for OpenType-kerned
  runs), a one-glyph look-ahead ±1 px correction, 1024 px chunks, the PDF
  writer's per-glyph unit rounding restarting at rsid run boundaries;
  justified spaces, plain-path kerning and letter spacing rounded apart; line
  starts rounded per part (frame or margin, left indent, first line). Visual
  suite J +4.34, SSIM +2.89; output +3.7 %. Left exact: RTL, East Asian
  text, text scaling, fallback glyphs, centred and right-aligned starts.
  Regressions against online references, not yet investigated unless noted
  (Jaccard points):
  - case152–154, case166, case167 (−3.2, one small case five times); our
    text order differs from Word's.
  - case149 (−4.4): glyphs closer to Word than before (mean |dx| 0.087 →
    0.058pt); a near-empty page, so likely score noise.
  - nested-float-align-m15-center-* (−2.4) and -right-* (−1.7): the line
    beside the inner float now starts ~0.47pt right of Word (was ~0.3).
  - cell-label-double-level-left-para-center-text1 (−2.8),
    underlined-space-compat15-* (−2.1), air_pollution_permit_form (−2.4).
  - About −1: traditional_skills_job_form (glyph |dx| 0.195 → 0.245pt),
    ut_koer_annual_meeting_minutes, frame-float-clearance-*, case69,
    case119, case142, min-width-run-whole.
  The 12 Quartz-produced references (case63/64, door_air_cooling_unit_spec,
  missing_font_substitution, 8 vml-*) lose up to 9.6: they place glyphs at
  exact positions; best-effort, not tuned for.

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
- A nested table's pictures now floor the host's natural width and the nested
  percentage share (issue #142). Open: in a fixed 396pt host Word draws the
  40% nested table's picture page-centred (control-02-fixed-outer) where we
  keep the nested table at the cell's left; the nested width/position in a
  fixed host is unprobed.
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
- SVG pictures (`asvg:svgBlip`) draw as vector Form XObjects (shapes, paths,
  solid paints, strokes, dashes, transforms, `.class` stylesheets; the viewBox
  is stretched to the frame like Word does). Anything else keeps the PNG
  fallback: `text`, `use`, `image` (indigenous_innovation_sandbox_agreement's
  icons), gradients, opacity, clip paths, filters, nested `svg`, the 130 less
  common colour names. Word online rasterises SVGs at 2× their pixel size;
  desktop Word emits the vectors with a ~0.7px inset in a 161×81px box for a
  160×80 drawing (unmodelled, no online reference does it).
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
4. The space fallback lays U+2002/2003/2009/202F out at U+0020's width.
5. `anchor_seq` counts per parse pass: textboxes from paragraph-level
   `mc:Choice` sort after every run-level anchor.

## Other features

- Field codes: only PAGE, NUMPAGES, STYLEREF, PAGEREF and IF are evaluated;
  MACROBUTTON shows its display text or picture; the rest print their cached
  result.
- Body PAGE/NUMPAGES (issue #140) come from a first render pass: a PAGE in a
  paragraph that spans two pages shows the first page's number, and fields in
  table cells keep their cached result. Body STYLEREF (issue #142) is resolved
  at parse time: nearest preceding paragraph in the style, else the first
  following one; `IF` over STYLEREF in the body still prints its cache.
- Pictures: 1-bit sources stay 1 bpc DeviceGray like Word's export, but
  8-bit indexed sources are written as RGB Flate where Word keeps Indexed, and
  every Flate image gets `/Interpolate true` where Word sets it only on JPEGs
  (unmeasured).
- Line numbering ignores `w:restart`: the counter always runs on, never
  restarting per page or section.
- altChunk HTML: `ul`/`ol`/`li` are dropped by the converter.
- `tblGrid`-less column inference (`row_widths`) doesn't add `gridBefore`
  columns (Word always writes a tblGrid).

## Parked

- **Hyphenation**: Word's online converter never hyphenates; the 8
  `hyphenation/` fixtures wait for desktop-Word references with proofing
  tools.

## Fonts

- An at-least line whose minimum wins bottom-aligns at the deepest descent
  of all its faces in Word (Arial 20pt at-least 30 with ☑: baseline 0.78pt
  higher, Segoe UI Symbol's deeper descent); `boxed_line_ascent` only looks
  at the paragraph's tallest run, and per-line sizing keeps the paragraph's
  baseline when the pitch doesn't change. Same gap for mixed-font lines.
- ※ ‥ ‧ ‵ in a Latin run come from MS Gothic in Word, the East Asian slot
  rather than a fallback; we rescue them from Segoe UI Symbol.
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
