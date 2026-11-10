# Roadmap

Open work only. Finished rounds, their rules and evidence live in git history
(`git log -p roadmap.md`), the commit messages and `current_focus/`.

## Open fixture leads

**Regressions accepted with the PRs #154–#158 baselines (2026-11-09):** 827
scores accepted (739 J up, 16 down, 39 new fixtures). The drops of 1pt J or
more, with the cause where known. "Positions better" means the fixture now
hits more of Word's baselines and rules within 0.05pt than before, but the
lines fall differently on the 150 dpi pixel grid.
- hyphen-opening-guillemet-nonbreaking-hyphen 84.6 → 70.8: our Times New
  Roman 5.01 has no U+2011, so the line falls back to Cambria Math and grows;
  Word's TNR has it. Environment font, not a rule.
- toc-section-page-labels 80.6 → 73.0, case63 55.7 → 52.6 (SSIM 74.8),
  door_air_cooling_unit_spec 75.2 → 72.2: desktop Word / Quartz references,
  which use the 0.24pt local grid and no writer stretch. Not tuned for.
- japanese_interlibrary_loan 71.8 → 66.6: Word's cell rows are 0.25pt
  multiples (27.5, 28, 39.25), ours 27.45, 27.85, 39.5; everything below
  y≈170 sits one grid step low. Does Word quantize row heights?
- symbol-fallback-per-character 79.2 → 77.3: from the third line down we
  sit one grid step (0.25pt) low. Word draws some fallback glyphs with plain
  Cambria, which our chain (Arial → Cambria Math → Segoe UI Symbol) never
  uses, so those lines get another height.
- case146 83.8 → 77.1: positions better (3/7 vs 1/7); the rule at y 187
  is 0.24pt high, a layout error that predates the grid work.
- Positions better: nested-first-row-prefix300/330-cant1 99.9 → 95.3 (6/6
  exact), text-box-table-compat14 94.4 → 90.8, autofit-cell-widths 88.4 →
  85.6 (27/27 exact), japanese_medical_consultation_form 64.9 → 62.2,
  isla_language_lesson_plan 69.5 → 67.4, polish_municipal_letter 77.8 → 76.2.

**Unaccepted drops (2026-11-10, main at c6981122e, from commits before PR
#161; baselines left at the old scores, cause not bisected):**
- massachusetts_community_sanitation J 90.9 → 88.9, SSIM 97.4 → 96.9.
- maine_criminal_history_record J 74.8 → 88.0 but SSIM 94.3 → 92.5.

**Engine-gap list (CI comparison 2026-11-09, `git show
origin/gh-pages:report.json`):** 48 fixtures where another engine beats us
by ≥3pp J (120 on 2026-11-07), ours vs the best engine:
- ≥20pp: empty-margin-frame-long-prefix 47.5 (office2pdf 88.9),
  italian_academic_cv_form 52.9 (LibreOffice 85.6), vaccines_history_chapter
  46.4 (LO 78.2), slovak_pedagogical_practice 27.2 (LO 57.5), case82 22.6
  (dxpdf 49.7), case60 75.0 (office2pdf 97.6),
  nested-floats-spaced-text-bottom-cell 42.8 (LO 63.2).
- 10–17pp: candidate_reference_check_form 73.8, case145 58.2, case140 56.2,
  line-anchor-overrides 76.6, connector-paragraph-anchor-plain-0/-240 83.0,
  case64 72.0, hyphen-opening-guillemet-nonbreaking-hyphen 70.8,
  period-after-space-filler-period-normal/small-compat15 75.8/79.4,
  czech_village_budget_commentary 62.1,
  romanian_candidate_declaration_form_f2 65.5, centered-list-marker 72.0.
- 4–8pp: case63 52.6, case37 23.2 (skipped, see below), case121/122 65.0,
  seq-fields 78.1, nested-percent-2500-* 89.8, case151 58.3,
  shifted-text-position0-mixed1-* 72.2, numeric-suffix-final-five /
  -short-final / -split-final 48–49, case114 47.4, nested-percent-5000-*
  93.9, displaced-anchor-narrow-compat12-overlap0-* 85.0,
  go_math_grade4_guide 78.1 (double-underline was on this list at 55.8; it
  is 80.9 since the 2026-11-09 double-underline rule).
- 3pp: case152–154/166/167 66.8, vml-group-picture-washout 96.8,
  traditional_skills_job_form 52.6, case59 93.1.

**Benchmark-gap round (2026-11-08/09):** the external benchmark's published
per-document scores against ours on its clean corpus (2478 documents): ours
66.8 J, theirs 74.4; two thirds of the 60 worst are the references' platform
(Word for Mac fonts and hhea metrics, Mac glyph rounding) and stay as they
are. Its tracking sets (3393 + 410 + 116 documents) have no baseline run of
ours yet. Probe generators, the per-document join (`gap.py`) and the paired
rescore script live in the main checkout's untracked `tests/output/jub/`.
Open, diagnosed:
- Floating tables keep the old 20% wrap threshold: probe `tblwrap_d*` says
  the band including bottomFromText counts for them too (as it now does for
  pictures and text boxes), and case45–47 gain 3–15 points under that rule,
  but case40's line 1.6pt into a right table's margin runs full width in
  Word (−19). Find what separates them before changing tables.
- A table style with a `w:pPr` (even `jc` alone) and no `rPr` makes Word
  draw Normal-styled cell text at a size that is neither Normal's nor
  docDefaults' in any consistent way (probe matrix `probe_tablesize2/3`:
  docDefaults 11 / Normal 12 → 11, 10/12 → 12, 14/12 → 14, 11/14 → 14; a
  table-style rPr sz 10 → 12, sz 14 → 14). Normal's rFonts always apply. A
  Russian landscape table shows Word at 11 where ours is 12.
- An empty "No Spacing" paragraph under a 10pt Verdana heading is 12.2pt in
  Word (a 10pt line) although its mark inherits docDefaults' 11pt; ours 14.0.
  A probe of the same structure gives 11pt, so some other property of that
  document decides.
- A calendar's table rows come out 1.5pt shorter than Word's per row
  (Century Gothic cells, 4–5 lines each); a header table's shaded row 5pt
  taller. Measure against an online export before chasing (Mac metrics).

Leads from the 2026-11-07/08 triage round, still open:
- isla_language_lesson_plan (67.4): its o:hr paragraph below a header table
  lands 0.4pt low; Word draws the rule ~1.4pt lower relative to the table
  than we do, while body probes (`tests/output/koer/hr/` on branch
  koer-overlap, 2026-11-07) put our rule 0.45pt too LOW and show a rule-only
  body line as tall as a text line. Word also draws o:hrstd rules ~0.75pt
  thick, we 0.5.
- `empty-margin-frame-long-prefix` (47.5): the frame's anchor is on page 2;
  `empty_frame_table_bands` only reserves the box on the first page. Word
  reserves it on the page where the anchor lands (2026-11-08), so the fix
  needs to know which page the preceding table ends on before layout.
- Double underline geometry is done (`layout.rs` `double_underline`,
  2026-11-09, 170/172 online probe points exact); unchecked: the 2026-11-07
  note that Word also underlines a line's trailing space.
- nested-floats-spaced-text-bottom-cell (42.8): the same two registration
  tables in a vAlign=bottom cell. Bottom cells now float like centred ones,
  but Word sets the 12 spacer marks beside the floats (topic 2 lines higher)
  and the second float 9pt right (its leftFromText); ungating
  `spacers_clear_nested_registration` (gated to top-aligned cells) for
  bottom cells overlaps the floats (tried 2026-11-08).
- nested-floats-spaced-text-short-spacers (56.6): Word starts text in the
  narrow strip beside the cell's floating tables; same special-case tangle
  as the bottom-cell variant above (looked at 2026-11-08, not changed).
- Header inline picture at 1.5 spacing (czech_village_budget, 0.6pt body
  drift): Word's picture line is 1.13pt taller than ours.
- candidate_reference_check_form: rows drift 0.6pt then 1.6pt; Word paints
  shaded cells in strips (its tcMar bands), not one rect.
- slovak_pedagogical_practice (27.2, body +2.9pt): removing its header
  o:hr run moves Word's body down 1.75pt and ours not at all; the run is
  Cambria 13 under a Korinna BT (missing) mark, so the mark font's
  substitute probably sets the empty line; a further 1.2pt residual.
- vaccines_history_chapter: page 1 fits one more line than Word (−14pt
  after); 2.6pt lost in the anchored-textbox timeline above the body, and the
  16pt subtitle is 11.5pt low.
- italian_academic_cv_form: the last row fits by 0.26pt in ours, misses by
  ~0.3pt in Word (0.44pt drift above + row height); no single rule.
- romanian_candidate_declaration_form_f2: header 3.6pt high.
- case60: bullets inside SmartArt shapes.
- case140/case148 nested-table cell rules sit 0.36pt high and start 0.7pt
  right of Word's (rules present, not missing).
- Not worth chasing: sparse probe fixtures whose text is within 0.5pt and
  whose gap is 1–2px (connector-paragraph-anchor, picture-mark-leading-*,
  nested-percent-*, dutch_council_member_resignation).

**torture_test** (two-page fixture built from features absent from `src/`,
kit with generator in `tests/output/torture_test/`, Word online reference).
Linked text boxes move whole paragraphs between boxes where Word splits
between lines. Open:
- **Charts** (3-D, 2-D layout, Office 2016 chartEx): status, probe kit and
  next steps in `current_focus/3d_charts.md`. Open, in order: the
  perspective box-width rule at 4 rows over 8 categories and Word's
  off-centre aim; standard grouping with several depth rows; 3-D pie
  perspective, best-fit labels and legend gap; surface, stock, pie-of-pie
  and bar-of-pie not drawn. Fixtures `cases/charts_3d` and `cases/charts_ex`
  guard what is drawn.
- A `keepNext` paragraph before a chart paragraph is stranded at the page
  bottom; Word carries it to the next page with the chart.
- Inline `wps` shapes with text, charts and SmartArt are block boxes at the
  paragraph start, not line objects (a picture-filled shape is now a
  picture); an inline-object run kind would fix all three.
- The empty line before a page-break run gets no margin number ("19").
- Value labels only on vertical clustered bars; titles not on pie/radar;
  shape soft edge absent; reflection mirrors the fill, not the text.

**Weakest fixtures (2026-10-27 round, scores as of 2026-11-09):**
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
- **tokyo_welfare (57.4)**: `snapToChars` (whole grid cells, 12pt text takes
  two) is not modelled; only the `linesAndChars` pitch is. Under the char
  grid, a compat-14 left-aligned list line wraps 認 with 6pt free where our
  squeeze of （ ） 、 keeps it, yet Word does squeeze the 「す。）」 line
  above (both kept under a plain `lines` grid). Turning the pull off for
  the grid broke the second line; probe what separates them.
- **japanese_land_development (44.7)**: Word puts no gap between "、" and a
  digit ("2、3", 10.5pt step), which our autospace rule matches. But Word's
  text after "　2　" plus a plain half-width space starts 2.6pt (a quarter
  em) later than ours, so the note wraps one character late. Probe where
  Word's gap sits: "　" + space, or space + digit.
- **sample500kB (13.0)**: Word online draws "DejaVu Sans" in Sylfaen and
  "Open Sans" in Segoe UI, by name: charset 00/01, family roman/swiss/auto and
  removing Open Sans' altName Arial change nothing (probes 2026-10-27). Needs
  Sylfaen and Segoe UI in the assets; DejaVu Sans in ~/Library/Fonts wins
  locally anyway. Word's cloud cache has both (FontCache/4/CloudFonts:
  Sylfaen 5.08, Segoe UI 5.67 regular/bold/italic/bold-italic); vendoring them
  alone moves no fixture (A/B 2026-11-08). FOLLOW-UP: decide whether to treat
  fonts Word online lacks (DejaVu, Open Sans) as missing even when installed,
  and find Word's rule for which names fall to Sylfaen (also "Helvetica
  Neue"). Its page-4 picture sits lower than Word's because our text above
  runs seven lines longer.
- **uk_commercial_lease (34.6)**: on page 64 we fit 4 lines of 2.1.6 above
  the notes where Word fits 2 (our separator 736.83, Word's 737.23), so a
  blank page 67 goes missing; our fit test is more lenient there. Page 11's
  notes block is 0.67pt short from per-line differences of 0.05-0.25pt; the
  "bullet lines 9.2 vs 9.75" note does not reproduce (Word 9.25, ours 9.20).
- **air_pollution (64.7)**: its accents are stored decomposed ("t" + U+030C);
  Word draws the precomposed glyph ("ť", 0.53 em in Times New Roman Bold
  Italic) and maps that one glyph back to the decomposed pair in ToUnicode.
  Composing at parse time gained +1.0 J but lost 47 letters from our
  structure tree; an ActualText span instead garbled the extracted text and
  added 3 PDF/UA failures (both reverted). Needs a glyph-to-several-chars
  ToUnicode entry. The dot leader after "Katastrálne územie: " (a tab
  after a w:br) is missing; leaders start at x 68.8, Word 66.0.
- **japanese_hr (69.0)**: sub-rows are 19.0pt where the merged left cell's
  continuation carries an invisible sz=12 border, 18.25 elsewhere (ours
  18.0); Word puts the wrapNone box's empty anchor paragraph below the table
  (all six 10pt paragraphs follow it); `w:spacing w:val="-10"` in 7pt Meiryo
  draws -1.0pt per character in Word (-1 Tc), ours -0.5;
  `balanceSingleByteDoubleByteWidth` doubles `w:spacing` only for
  all-East-Asian body runs, and the half-width space it widens is not done
  (Meiryo spaces 0.507 em, ours 0.332; 10 fixtures set it); a paragraph
  bottom border ignores `w:ind w:right="-903"`; the banner title sits 8.8pt
  high.
- **swiss_materials (58.9)**: the VML watermark draws ~24% heavier and ~4pt
  left of Word's; the "Helvetica 55 Roman" bold headings (altName
  Bahnschrift Light) draw regular-weight in Word; page 6's table bottom
  border is solid grey in Word, dotted in ours; page 1's `›` bullet sits
  3pt low.
- **advocacy_initiative_newsletter_2013 (36.7)**: our nested picture table is
  1.4pt narrower than Word's grid, so the stories wrap a word earlier or later
  and pages 2-3 end 24pt off; the standard `o:hr` draws 1.5pt, Word 0.5. Word
  starts the text ~91pt right of the picture table's edge for both 1755tw
  grids, we start at 89.5-90 (table width + rightFromText).
- **cote_dazur (78.0)**: body text does not wrap around header floats (Word
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
- **PR #44 only before a continuous section**: the section-break line kept
  after a table is measured only when the next section is continuous; before
  a new page Word's behaviour is unprobed.

**WordCraft head-to-head leftovers (2026-10-13):**
- **Nested floating tables in a cell** (case145/148): without a paragraph
  between two floating tables Word stacks the second below the first
  (`tblOverlap="never"`, cell 13.75pt taller); we let them overlap. Cell
  bottom borders start 0.7pt left of the text in Word (tblCellMar left=0).
- **case113/114**: "Body marker" 1.3pt left of Word's. Word probe
  2026-11-08 (borders added): the floating table spans 71.8–431.4, the
  text beside it starts at 432.42 (unbordered 432.93), ours at the table
  edge; the offset follows the table width (60pt narrower → 60pt left). The
  same ~1pt shows in advocacy's story text beside its picture tables. Find
  the rule (device grid? border band?) before adding a gap.
- **Documents without styles.xml**: our built-in defaults
  (`NORMAL_TEMPLATE_DEFAULTS`) match desktop Word, not online (case135's
  cell line: desktop 633.22, us 633.20, online 633.70). Probe Word online's
  defaults for a docx without styles.xml.

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
- ut_koer: whether compat 14 header floats also skip the full-width push in
  `float_text_push` (unprobed; the per-line wrap skip is done).
- multi_font: needs Copperplate Gothic Light (in Word's cloud catalog; pick it
  in Word's font menu to download; not on this Mac as of 2026-11-08, Word's
  DFonts only ships the Bold); page 1 also runs ~14pt taller than Word's.
- case37 (in SKIPLIST): its reference shows five black boxes where its
  shapes are, and a fresh export does too (2026-11-08): the fixture's
  SmartArt is hand-built, with an empty data model and an empty layoutDef
  beside a full drawing cache. Word rebuilds from data and layout (linking
  the cache via dsp:dataModelExt changes nothing), so it draws its default
  five unstyled blocks. Rebuild the fixture from a real Word SmartArt;
  re-exporting won't help.
- case58 (3D effects) and case65 lack a reference.
- case77's on-page label claims thaiDistribute gets distribute's treatment;
  Word disproved it. Fixing it means regenerating DOCX and reference.

Open queue from the large-corpus and focus-fixture rounds:
`current_focus/layout-accuracy.md` §5.

## Line and paragraph layout

- **compressPunctuation** (issue #139, rules in `pdf/layout.rs`
  `PUNCT_SQUEEZE`/`PUNCT_PULL_FREE`): zh-TW centred marks lose the cut on
  both sides (`locl` glyphs; taiwanese_education 二、 sits 1pt left of ours);
  autoSpaceDE gaps are lost at run boundaries unless the runs merge
  (chinese_costume "5-10 分钟").
- **0.25pt baseline grid** (Word quantizes baselines to 0.25pt online, 0.24
  local; body, header, footer and cell baselines and table/paragraph rules
  follow it since 2026-11-08, `current_focus/fundamentals.md` §2). Open:
  text-box and footnote text get neither grid nor stretch
  (text-box-table-anchor: the text-box paragraphs around the tables sit
  0.14–0.17pt high, `spAutoFit` 413.23 vs Word 413.37, `3pt borders` 445.23
  vs 445.40, `after` 449.23 vs 449.40, while the cells inside the same boxes
  match within 0.02pt); table rules ignore w:vAlign's shift
  (`pdf/table.rs`); japanese_interlibrary_loan's
  cell rows are 0.25pt multiples in Word (27.5, 28, 39.25) but 27.45, 27.85,
  39.5 in ours, so the layout may quantize row heights. Residuals to
  re-measure now the grid is in: greensboro (Arial 11pt at 1.5 lines: ours
  18.97, Word 19.0), romanian_candidate_declaration (borders drift 0.1-0.3pt
  per row), case69 (last line 15.00 vs 14.83 under `vAlign=center`), case33's
  list lines (16.115 vs 16.0), uk_commercial page 11's notes block (0.67pt
  short).
- **Mixed Latin/East Asian justification** (issue #142, `MixedJustify` in
  `pdf/layout.rs`): a Latin letter glued to a CJK character ("abc一二三def")
  measures the gap plus ~5pt (autoSpaceDE and something else) on both sides;
  Word's squeezed spaces come in 0.25pt steps and it let the repro's final
  period hang 2.9pt where we squeeze the spaces fully; `Distribute` keeps
  the old per-character spread.
- **`w:ind` character units** (issue #143, `apply_indent_chars`):
  numbering-level `*Chars` are not read (`docx/numbering.rs` reads only the
  twip values); headers and footnotes use no grid delta.
- **Paragraph borders** (issue #145): renewable_dispatch's last blue "CBIR
  Process" heading bar is not drawn at all (white text on white); table
  cell borders keep their own dash arrays and draw wave and thin/thick pairs
  as one line (`pdf::table::draw_border` has no Wave or ThinThick arm).
- **Trailing punctuation**: Word never lets a Latin mark hang past the margin
  (`overflowPunct` is not what it does). The CJK half-mark allowance
  (`punct_half_out`) was probed with `.` only; `, ; : ! ?` assumed. Issue
  #142's "introduction." still wraps: with Word's glyph advances the line
  sits 0.4pt inside the cap, with ours 0.4pt outside.
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
  no line (no `Bar` variant in `TabAlignment`).
- Underlines: dotted/dash/wave/thick draw single; underline colour ignored
  (`docx/styles.rs` reads only none and double).
- `w:textAlignment`: top is laid out and center only centres list markers;
  baseline, bottom and auto parse as auto.
- `w:mirrorMargins`: parsed (used for odd/even filler pages and the change-bar
  side), never applied to the margins. Swap left/right margins (and the
  gutter side) on even pages.
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
  `char_justify_gaps`). Kashida on Arabic text is unprobed (the Latin
  kashida reserve is done). thaiDistribute on Thai text is untested.
- `eastAsiaTheme="minorHAnsi"` (997 runs in the corpus) is ignored by
  `resolve_east_asia_font`. Resolving every slot (tried 2026-10-31 and again
  on top of the cs-font writer restart, 2026-11-03) changes 3–4 fixtures
  without moving glyph x towards Word; no fixture needs it yet.
- Writer restarts (probe 4, `current_focus/fundamentals.md` item 1): Word
  also restarts where `noProof` (not parsed) or `szCs` (parsed, not compared)
  differ, and at ~490-character internal boundaries no model can see.
- Missing fonts: names with `;` (Word takes "Open Sans;Arial" literally, we
  split), and "Open Sans" → Segoe UI (`current_focus/merge-regressions.md`
  item 5). Courier → Courier New is Windows' substitute, unverified for Mac
  Word.
- **Online glyph positioning** (`src/pdf/glyph_grid.rs`,
  `current_focus/fundamentals.md` §2, probes 1–5). Left exact: RTL, East
  Asian text, text scaling, fallback glyphs, centred and right-aligned
  starts. Regressions from the 2026-10-30 merge still below their pre-merge
  score (Jaccard points): case152–154, case166, case167 (−3.2, one small
  case five times; our text order differs from Word's);
  nested-float-align-m15-center-* (−2.4) and -right-* (−1.7): the line
  beside the inner float starts ~0.47pt right of Word (was ~0.3); case142
  (−1). The 12 Quartz-produced references (case63/64,
  door_air_cooling_unit_spec, missing_font_substitution, 8 vml-*) lose up to
  9.6: they place glyphs at exact positions; best-effort, not tuned for.

## Tables

- **Autofit vs Word** (2026-10-05 x-drift survey): case45 (tblW 4000, tcW 4680
  per cell) keeps Word's saved grid 2091/1909, we split 100/100; case61's
  200-twip column is widened to fit its text by Word (its "merged" cells
  carry no gridSpan), and its table 2 first column is 9.6pt too wide in
  ours; case6 and case52/53 whole columns shifted 7–11pt; air_pollution's
  dot lines fit 42 dots in Word, 41 in ours. Re-measure after the
  preferred-cell-width work (PR #149).
- **Over-wide AutoFit cells** (2026-11-08, case100 probes): squeezed by room
  above each column's longest word. erasmus_plus's three slightly-over tables
  (cells ~446pt on 439pt) come out up to 1.6pt per column off Word's grid,
  which spares some columns that rule would cut (−1.1 J); federal_procurement's
  merged-cell table is within 1.5pt.
- Some rows run slightly short (slovak_eu_directive ~0.03pt per row); a
  vMerge/colspan test table gets a far too narrow column.
- Nested tables use neither compat rule for the outdent
  (`render_nested_table`); needs a compat 14/15 probe. A nested header table
  (logo beside a title table) can sit 5–7pt low.
- Nested table in a fixed host: in a fixed 396pt host Word draws the 40%
  nested table's picture page-centred (control-02-fixed-outer) where we keep
  the nested table at the cell's left; the nested width/position in a fixed
  host is unprobed.
- Single-row nested tables: global nested-table origin/float positioning,
  horizontal positioning and explicit/page/column horizontal anchors of
  nested floats remain.
- The must-split path (row taller than a page) tests only `!cant_split`, so an
  exact row taller than a page splits; Word likely clips it. Unprobed.
- Column auto-fit measures cells without footnote/endnote marks (the
  reference run's text is empty).
- `tblW` vs oversized `gridCol` is unclamped (no fixture triggers it).
- Cell image paragraphs don't follow the short-picture rule (unverified
  since the cell picture-line changes of 2026-11-07/08).
- Vertical CJK cell text: `render_vertical_cjk_cell` sizes by chars ×
  font_size, ignoring paragraph spacing (misaligned in merged cells).
- Labels hanging outside a cell (probe Word first).

## Floats, frames and wrapping

- Bands (`Band`, full-width floats) live beside `FloatZone`; one mechanism
  with a full-width flag would be cleaner.
- The page pre-scan estimates where the page ends from plain line counts and
  stops at the first table; an anchor that lands on the next page anyway
  keeps its band on this one too.
- Paragraph-relative topAndBottom pictures below their anchor's text still
  reserve in the anchor (textboxes and groups are bands; no
  `hangs_below_anchor` check for pictures).
- Group `positionV` alignment (top/center/bottom) is lost like the
  horizontal one was: children get `offset_or_zero`.
- Text-anchored body frames without an explicit offset (`lifted_frame` lifts
  only offset frames), and wrap-around frames reaching into the text
  (croatian_grant's two), still flow inline.
- `content_h` in `render_paragraph_block` includes the room reserved for a
  topAndBottom float below the lines; borders still use it (shading doesn't).
  Unverified against Word.
- Header flow (and `textbox_span`) reserves `height_pt` for TopAndBottom
  textboxes while the body uses `textbox_height`.
- Word puts the anchor's own line box below a no-strip float (sample500kB p4:
  Word gap 67.7pt, ours 51.8pt).
- pendulum: Word places the 144pt-tall images on pages 2 and 6 14.8pt higher
  than the others despite identical `posOffset=0`.
- Look-back wrapping beyond the paragraph directly before an anchor needs a
  paginator or two-pass layout.
- Tight vs Through: both wrap to the polygon's per-scanline outer extent
  (`poly_scanline`); Through should fill concavities. case42's 53-vertex
  polygon wrap breaks lines differently from Word.
- Off-page `posOffset` (learning_cultures rId14: 334pt-wide image placed past
  the page edge): Word clamps or reflows; a naive right-edge clamp regressed
  stem_partnerships.
- Group-level flipH/flipV/rot are ignored except for linear connectors;
  image/text/arc mirroring in groups is open. Canvas/group inside
  paragraph-level `mc:AlternateContent` grabs only the first wsp.
- Wrapped running-head pictures: tall final pictures and three-picture
  layouts keep residual discrepancies.
- `wp:effectExtent` is not part of an inline picture's line height in body
  text (ut_koer's 0.75pt residual); only headers set
  `inline_image_extra_height`.
- Anchored SmartArt with wrapping other than wrapNone is laid out as inline
  (no fixture).

## Columns and sections

- Footnote continuation across columns (case82 J 22.6): Word may continue a
  note under the next column so more body text fits; needs a rule for how
  much must stay with its reference.
- case80's 4-column region: Word's separator runs one line (15pt) below
  column 1's last space after; unexplained. The table probe's region ends
  1.6pt off. Each balancing trial lays the region out again.

## Footnotes and endnotes

- Footnotes that don't fit: Word splits a note across pages with a
  continuation separator (uk_commercial p17/18); we move the reference line.
- frame-clip-page-break (J 60.9, SSIM 72.2): body and frame match Word; the
  footnote separator rule sits 10.4pt too high (ours 330.30, Word 340.73 from
  the top). The docx has no styles.xml and the separator paragraph has no
  pPr, so its line height falls back to defaults; suspect ours differ from
  Word's (unverified). The near-empty page magnifies the miss in the scores.
- Footnote paragraphs other than numbered ones are hand-built in
  `parse_footnotes` (endnotes already use `build_paragraph`): borders,
  shading, `style_id`, contextual spacing missing; note tables are flattened
  to one paragraph per row (`Footnote { paragraphs }` should become blocks).
  Footnote paragraphs place exact and at-least lines at the plain ascent
  (`footnotes.rs`).
- Inline endnotes are not paginated (`render_endnotes_inline` ponytail;
  erasmus' endnote block runs past page 3's margin); `endnotePr
  pos="sectEnd"` unsupported.

## Tracked changes and comments

- Revision markup (Word's inline All Markup: author colours, underline /
  strike, deleted text kept in the layout, change bars, moves) is opt-in:
  `ConvertOptions::revisions(RevisionView::Markup)` / `DOCXIDE_REVISIONS=markup`
  / CLI `--markup`; the final view stays the default. Rules, probes, the
  external tracked-corpus scores and the ranked remaining gaps are in
  `current_focus/tracked-changes.md`. door_air (#259/#260) is the fixture
  whose reference shows markup, in balloons (out of scope).
- Next steps, in order: (1) the page's text re-flowing below a
  page-anchored floating table that fits (p24a/p24g; needs the page re-laid
  once the float lands, our render is single-pass); (2) a header that
  overlaps the body on later pages (a press-release family, ~60pt); (3) an
  auto table whose cells' `tcW` sum to twice the column (Word centres it off
  both page edges); (4) text boxes' change bars (unprobed); (5) a moved list
  label is struck once, not twice.

## Pictures, shapes, charts

- Picture effects: 3D (`a:sp3d` unread; `a:scene3d` only sets the camera
  rotation; case58), preset shadows (`a:prstShdw`), theme colours in effects
  (`schemeClr` falls back to black in `parse_dml_color_alpha`), floating
  header/footer pictures draw no effects and block ones lack inner shadow
  and reflection, shadow / soft edge / inner shadow on shapes and text boxes
  (glow and reflection are done). Soft-edge and reflection masks are built
  on the uncropped source; SmartArt pictures use their own draw path.
- EMF: clipping regions and opaque text backgrounds; immediate-mode segments
  and mid-path SaveDc/RestoreDc; mixed vector+bitmap EMFs (the bitmap is
  dropped).
- SVG pictures (`asvg:svgBlip`) draw as vector Form XObjects (shapes, paths,
  solid paints, strokes, dashes, transforms, `.class` stylesheets; the viewBox
  is stretched to the frame like Word does). Anything else keeps the PNG
  fallback: `text`, `use`, `image` (indigenous_innovation_sandbox_agreement's
  icons), gradients, opacity, clip paths, filters, nested `svg`, the 130 less
  common colour names. Word online rasterises SVGs at 2× their pixel size;
  desktop Word emits the vectors with a ~0.7px inset in a 161×81px box for a
  160×80 drawing (unmodelled, no online reference does it).
- SmartArt: `dsp:grpSp` groups, `dsp:cxnSp` connectors (only direct `dsp:sp`
  children of `spTree` are read); the layout engine for files without the
  `dsp:drawing` fallback (not planned).
- Charts: surface, stock, pie-of-pie/bar-of-pie, combo (only the first
  type), secondary axes; see `current_focus/3d_charts.md`.
- WordArt level 5: VML gradient/pattern fills, VML shadow, shapetype →
  prstTxWarp mapping.

## RTL / BiDi

GPOS mark anchors (harakat off-centre); mirrored brackets: glyphs are drawn
left to right as Word does (2026-11-09, so MuPDF and pdfinfo read Word's
text), but Word maps the mirrored glyph back to the logical character
through a second font subset where we use ActualText; no joining across
runs that differ in formatting; lines with tabs keep logical order
(`build_tabbed_line` never reorders); tab stops not mirrored; Arabic
letters in a plain run whose font
has Arabic glyphs should use that font (Word: Courier New), we use the cs
font. Arabic in a face without Arabic glyphs probably rescues per character
since d3368996f; unconfirmed without an Arabic fixture.

## Office Math

Each needs a Word online probe first (corpus refs are print exports); rules in
`current_focus/office-math.md`.
1. `m:wrapRight`: continuation lines of a wrapped display equation flush
   right.
2. Inline radicals over tall contents: Word's surd sits 0.3–0.6pt lower than
   the half-surplus rule at every size.
3. An `m:nor` superscript sits 0.6pt closer to its base than Word's (e^max).
4. Placeholders: an inline n-ary with placeholder limits stands 2.6pt too
   low; a function with an empty argument is 2pt short; a sup over an empty
   sub sits 0.5pt low.

## Accessibility

Backlog (details in `current_focus/A11Y.md`):
1. Math speech reads matrices and accents as their contents in order.
2. Link rects and outline destinations ignore the comment-pane zoom and
   vAlign (no fixture has links with either).
3. Wingdings 2 and Webdings extract as private-use code points.
4. The space fallback lays U+2002/2003/2009/202F out at U+0020's width.
5. `anchor_seq` counts per parse pass: textboxes from paragraph-level
   `mc:Choice` sort after every run-level anchor.

## Other features

- Field codes: SEQ fields in text boxes, headers and footers keep their cache
  (`resolve_seq_fields` walks body paragraphs only). TOC regeneration
  ignores `\p` separators, `\b` scopes and TC fields, and indents a contents
  table "toc N" by 11pt × (N − 1) when the style is missing (unprobed); a
  table's cells count as its first page, so a split table's later rows
  point at the wrong page; a Word-written TOC spanning paragraphs is still
  drawn as cached.
- Issue #147 follow-up: an inline MathType OLE equation sits with its
  bottom on the baseline like any picture, but Word lifts it by the baseline
  offset in the equation's WMF preview (MathType's comment record), so the
  "(1)" labels of `rust.docx` sit 12pt below Word's. The MathType-generated
  WMF previews of its `both.docx` draw nothing in our output, where Word
  paints each as a picture (a bitmap record our WMF vector path skips).
- Issue #145 follow-ups: a table style's `tblStylePr` (firstRow…) `vAlign`
  is not read (base `tcPr` and `jc` are); table styles are still walked per
  table (`style_chain`) instead of flattened at load like paragraph styles;
  the row-split predictor (`first_chunks_fit`) and `find_cell_split` encode
  the first chunk twice; list labels in headers, text boxes and footnotes do
  not follow a centred or right-aligned first line as body and cell labels
  do.
- Body PAGE/NUMPAGES (issue #140): a PAGE in a paragraph that spans two pages
  shows the first page's number, and fields in table cells keep their cached
  result. `IF` over STYLEREF in the body prints its cache.
- Pictures: 8-bit indexed PNG sources are written as RGB Flate where Word
  keeps them Indexed (`decode_png_raw` returns None for Indexed).
- altChunk HTML: `ul`/`ol`/`li` are dropped by the converter.
- `tblGrid`-less column inference (`row_widths`) doesn't add `gridBefore`
  columns (Word always writes a tblGrid).

## Parked

- **Hyphenation** (issue #39): Word's online converter never hyphenates; the
  8 `hyphenation/` fixtures wait for desktop-Word references with proofing
  tools. When it lands it is a `ConvertOptions` field (`hyphenation`, default
  off; `DOCXIDE_HYPHENATION=1`, `--hyphenate`) like `resample_pictures`
  (issue #138), honouring `w:autoHyphenation` only when switched on.
- **Embedded-font references** (issue #139, 2026-10-30): Word online can use
  fonts embedded in the docx (case8's Feb 2026 reference draws Press Start 2P
  as `___WRD_EMBED_SUB_1437`), but nothing embedded survives
  `tools/word_export.py --preset online`: 楷体_GB2312 full or subset, under
  either name, and even case8's Press Start bytes in another document all
  fell back to YaHei; case8 itself now hangs Word (its fontTable rels list
  rId20 four times). `--preset print` does use them. Next step: export
  `tests/output/i139/gb2312font/embed4/issue139_gb2312.docx` through Word on
  the web to tell the docx apart from the Mac export path.

## Fonts

- An at-least line whose minimum wins bottom-aligns at the deepest descent
  of all its faces in Word (Arial 20pt at-least 30 with ☑: baseline 0.78pt
  higher, Segoe UI Symbol's deeper descent); `boxed_line_ascent` only looks
  at the paragraph's tallest run, and per-line sizing keeps the paragraph's
  baseline when the pitch doesn't change. Same gap for mixed-font lines.
- With only a regular face embedded, our bold run falls back to the charset
  substitute (Microsoft YaHei in issue #139) instead of emboldening the
  embedded face (the embedded lookup needs an exact name/bold/italic key);
  untested against Word.
- ※ ‥ ‧ ‵ in a Latin run come from MS Gothic in Word, the East Asian slot
  rather than a fallback; we rescue them from Segoe UI Symbol
  (`SYMBOL_FALLBACK_FONTS` order; U+2027 is not in the hint list).
- Bahnschrift, by name or as a missing font's altName, is measured with
  fitted metrics and drawn in Calibri like Word online (fixture
  bahnschrift-heading-metrics). Open from the same probes (kit in
  `.worktrees/issue-148/tests/output/swiss-probes/`, online export):
  - The constants (`BAHNSCHRIFT_ASCENT` 0.9965, `BAHNSCHRIFT_LINE_H` 1.2004)
    are fitted from the PDFs, not read from the font. Office's cloud copy of
    Bahnschrift Light (`~/Library/Group
    Containers/UBF8T346G9.Office/FontCache/4/CloudFonts/Bahnschrift/`) has win
    2036/422 on 2048 = ascent 0.9941, line 1.2002 em. The 0.0024 em extra
    ascent is the same excess Calibri shows (0.9521 in the font, 0.954
    measured), so store the font's values and let the usual rule add it.
  - "Cannot export the variable font" is unproven: the cloud copies are static
    fonts. The behaviour is measured; the reason is a guess, and the code
    comment in `register_font` repeats it.
  - "Helvetica 55 Roman" with swiss's fontTable entry but no altName is still
    Bahnschrift-measured Calibri in Word; we draw Arial (the "Helv" rule). What
    makes Word pick Bahnschrift there is unknown (a "Zqx Sans" entry with
    Bahnschrift's panose gets plain Calibri).
  - A missing font Word draws in Calibri puts its baseline 0.9715 em down
    (hhea ascent + lineGap), not real Calibri's 0.954; the line stays 1.221 em.
    We use 0.952.
  - "Zqx Sans" with no fontTable entry draws Calibri in Word (python-docx's
    template, Calibri theme), where `family_fallback` says Cambria; check
    whether the theme's minor font decides.
- `w:afterLines` alone is ignored (we fall back to the style's `after`); Word
  counts a line as 12pt (afterLines 80 = 9.6pt, same as after 192).
- Bundle metric-compatible open fonts behind a feature flag (Liberation
  Sans/Serif/Mono, Carlito, Caladea) for servers and Docker; today only
  installed fonts and the assets repo count (Carlito/Caladea are used when
  installed).

## Architecture and code structure

- **Paginator extraction**: a pass producing `Vec<Page>` of positioned
  elements, so widow/keep logic leaves the render loop, look-back wrapping
  becomes possible and PAGE/NUMPAGES resolve after layout.
- Cell and footnote paragraphs bypass `paragraph::build_paragraph`
  (`docx/tables.rs`, the hand-built branch of `parse_footnotes`): cells
  default paragraph shading and outline level and never get `mark_revised`
  (no change bar for a cell paragraph whose mark alone is revised;
  `paragraph.rs` sets it, `tables.rs` does not); one shared pPr resolver
  would close it. The markup view's `item_revision` and the two list-label
  fields are wired into both builders.
- One per-line height model shared by body, cells and headers (`pdf/mod.rs`
  alone computes a paragraph's line height six ways, `plain_line_h` with the
  `unsized_line_metrics` fallback and `para_line_h` without, and counts
  lines two ways: `paragraph_line_count` vs `line_count`; merging them moves
  output); one widow/orphan split helper for body split, keep-with-next and
  cell splits; run-boundary breaks from UAX #14 over the paragraph's joined
  text (`breaks_between` tests a two-character pair, so a URL split across
  runs can still break after `/`); one table-left-x function for body,
  header and nested tables; one new-page helper.
- One picture model: `docx/paragraph.rs` hoists a lone inline picture into
  `Paragraph::image` while two or more stay in runs; one shared
  `draw_embedded_image` for the five block-picture draw copies (rotation +
  effects; mod.rs, header_footer.rs, layout.rs, positioning.rs, table.rs); a
  shared `decode_raster` so `a:lum`, crop and soft-edge apply once and reach
  `embed_reflection`.
- Floating-table keep-together could be geometric (`fp.y > saved`) instead of
  reading raw `tblpY` (`text_anchor_offset`).
- Spring-cleaning next rounds (`current_focus/spring_cleaning.md`, whose
  "Not done" list is still accurate while its branch notes are stale):
  parameter bundling (`render_paragraph_lines` and `assemble_pdf_pages` take
  21 parameters each), `PageBuilder` (14 parallel `all_*` Vecs) →
  `Vec<FinishedPage>`, the 4× textbox layout, 5× picture effects, the
  triplicated `table.rs` row renderers.
- `pdf/images.rs`: the seven per-kind name maps remain beside the shared
  `ImageKey` XObjects; `build_paragraph_lines` taking `&HashMap<usize,
  &str>`; pair `image_names` + `effect_names`. `--no-resample` threads
  `resample: bool` through 6 functions and ~13 calls; an `ImageEnv { fonts,
  options }` passed instead puts the calls back on one line and makes the
  next image option (interpolation, JPEG quality) a field, not another
  parameter.

### Restructurings left by the v0.20.2 simplify review (2026-11-07)

Each is a larger change; every entry says what it buys and what it risks.
All are meant to keep output identical unless the entry says otherwise.

- **Change bars out of thread-local state** (`pdf/layout.rs` `CHANGE_BARS` /
  `BAR_EXTENSION`). Today `mod.rs` sets and clears the extension around each
  paragraph (3 set sites, 2 clears), bars are flushed by `take_change_bars`
  per page in `flush_page`, at render start, after footnotes and after
  headers/footers, and `table.rs` brackets each row with `change_bars_len` /
  `collapse_change_bars_from` (3 copies). Forgetting a pair fails silently
  (bars on the wrong page or none). Change: a `ChangeBars { bars, extension }`
  on `PageBuilder`, or `&mut` into `render_paragraph_lines` (10 callers). The
  ponytail note's own trigger ("until headers want bars") has fired:
  75e8b967b added them. Medium size, markup view only, low risk; do it
  before the next markup feature.
- **Split paragraphs share the paragraph ending** (`pdf/mod.rs`, the split
  path returns early around line 6358). It copies page-break-after handling
  and leading/trailing float rendering, and does not run the shared ending's
  `column_break_after` (line ~7133): check whether a split paragraph with a
  column break after loses it. `FragmentPlan` (Beside/Below/NextPage/
  PageBottom) re-does wrapping by hand instead of installing the trailing
  float as a `float_zone`. Change: let the split path fall through to the
  common ending and route the trailing float through the float-zone
  machinery. Large, touches the hottest pagination code; worth it because
  every new end-of-paragraph rule otherwise needs a second copy. Expect a
  few fixtures to move if the hand-rolled wrap differs from the zone's.
- **One page-end estimator and per-section bands** (`pdf/mod.rs`
  `paragraph_bands`, `prescan_page_bands`, `empty_frame_table_bands`, the
  `splits_for_sure` local in the 3-paragraph float look-ahead). Four
  separate guesses at where a page ends, and `paragraph_bands` (which lays
  out each blocking textbox) runs per page in the pre-scan and again in
  `step_below_bands`. Change: compute bands once per single-column section
  as a per-block `Vec` and one `estimate_para_height(p, col_w)` that the
  band collectors and the look-ahead read; the look-ahead's "three
  paragraphs" limit becomes "anchor estimated on this page". Large; the
  estimator merge moves output (the two line counters differ), the band
  cache alone does not. Part of the Paginator extraction above.
- **One float-blocks-band rule** (`pdf/mod.rs` `textbox_blocks_band` vs
  `positioning.rs` `image_blocks_band`; `hangs_below_anchor` exists for
  textboxes only). The two already differ: `side_room(wrap_text)` for
  textboxes, `no_strip_beside`'s `left.max(right)` for pictures. Change: one
  predicate over a shared float-box view (x, w, wrap type, wrap text,
  v_rel). Medium; decide first which side-room rule is Word's, since
  unifying picks one.
- **One vertical region resolver** (`positioning.rs` `resolve_fi_y_top`,
  `header_footer.rs` `resolve_tb_y_top`, which already builds a local
  region tuple, and two header copies at ~719 (pictures) and ~2013 (text
  boxes)). TopMargin/BottomMargin (837485a50) was patched into all four.
  Change: `VRelativeFrom::region(sp) -> (top, bottom)` used by every
  non-paragraph anchor. Small to medium; `(H+h)/2` vs `H-(H-h)/2` can differ
  in the last f32 bit, so check the visual hashes.
- **`render_floating_images` by `(index, &fi)`** (`pdf/mod.rs` `float_subset`).
  The split path clones its floats and re-keys two name maps only because
  the renderer keys XObject names by position. Change: take index/float
  pairs or a predicate. Small to medium, output identical.
- **Per-line layout state with one reset** (`pdf/layout.rs` line builders).
  Each new-line site resets ~6 fields by hand (`current_x`, `line_reach`,
  `is_first_line`, `tab_line_metrics`, …) and the tab-wrap site a different
  set; `tab_stop_pos = None` now sits at three sites. Change: a `LineState`
  struct with `new_line()`. Medium, output identical; prevents the next
  forgotten field.
- **Typed cell width** (`docx/tables.rs` `tcW`, `model/table.rs`). `width`
  holds pct values misread as twips ("25%" → 0.005) next to `width_pct`,
  which the layout uses; `apply_pct_width`'s proportional scaling relies on
  the misread values. Change: parse `tcW` into `Dxa | Pct | Auto`. Large and
  output-moving (cells written as `NN%` start reading right); worth it when
  pct-width tables next misbehave.
- **chartEx as one chart-type variant** (`model/chart.rs`, `pdf/charts.rs`,
  `pdf/fonts.rs`). chartEx is detected by `ex_text.is_some()`, the six chartEx
  variants are `unreachable!()` in the 2-D renderer, and histograms are typed
  as `Bar` (`docx/chartex.rs`; they never reach the bar renderer). Change:
  `ChartType::Ex { kind, text }` dispatched like Pie and Radar; `fonts.rs`
  then reads the title font and waterfall legend names from the model
  instead of re-deriving them. Medium, output identical.
- **chartEx legend through `render_chart_legend`** (`pdf/charts_ex.rs`
  ~185–275 is a second legend engine: own key/gap constants, row centring,
  right column; it imports only `render_swatch`). Change: let
  `render_chart_legend` take the text style and key/entry gaps, give 2-D
  charts a 10pt-black style. Large-ish; worth it the next time either
  legend changes.
- **List label as styled segments** (`pdf/list_label.rs`, `Paragraph`'s
  `list_label_revision` / `list_label_renumbered`). The markup view draws the
  label's strike/underline and the renumbered "4.2." by special-case code.
  Change: segments of (text, strike, underline, colour). Large for a
  markup-only gain; only if labels grow more styling.
- **Bahnschrift as a table row** (`fonts/mod.rs` `register_font`: the rule
  sits in five places — primary filter, altName reject, Helv skip,
  hard-coded "Calibri", metric override). Change: one
  `[(prefix, drawn_face, ascent, line_h)]` table; see Fonts for the open
  trigger and constants questions, which should be settled first.

## Performance

- Image recompression is the converter hotspot (case59: ~5.1 s in
  `embed_image_xobject`). Cheapest first: decode/resize in parallel then write
  XObjects in order (rayon is a dev-dependency only today); faster resampling
  (`fast_image_resize` or a cheaper filter than Lanczos3 for large
  downscales); scaled JPEG decode. Passing JPEG bytes through as DCTDecode
  would be a fidelity trade-off: Word always re-encodes, and `images.rs`
  follows it on purpose.
- Repeated WinAnsi conversion (line building, rendering, autofit);
  allocating `font_key()` at 9 sites; `WordChunk` clones font names.
- Library is single-threaded: font scanning, metrics, per-paragraph wrapping
  and XML parsing are candidates for parallelism.
- Column balancing re-lays the region per trial (`trial_at`).

## Test harness and tooling

- Deterministic output regressed once: scraped/door_air_cooling_unit_spec gave
  two different PDFs over three runs (2026-10-03); not investigated.
- Engine-comparison site: veraPDF already runs once per fixture; if runs
  grow, batch it per shard via a `page-metrics` pre-pass.
  `engine_compare.py` `pdf_creator()` truncates Quartz producers at the
  first `)`.
- SSIM with ±8px search (~17 CPU-s for a 205-page fixture): rayon over the
  window search, integral images, SIMD, if turnaround becomes a bottleneck.
- Text boundary (`tests/common/text_boundary.rs`): since 2026-11-09 our
  fragments join the line of the nearest Word fragment within 3pt and lines
  pair by longest common subsequence; open: a page shifted by reflow above
  (sample500kB p1, 18pt) lands on whichever Word lines sit at its tops. 40
  fixtures (mostly cases/ accepted visual-only, e.g. through-picture-*,
  ul-trail-space-*, identical-parts-and-fonts) have no text_boundary/a11y
  baselines.
