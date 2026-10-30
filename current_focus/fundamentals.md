# Fundamentals review — current focus

Are the core rules derived from Word, or fitted until a fixture matched?
Started 2026-10-28 after the Word probing workflow (`tools/word_export.py`)
made every such question answerable in an afternoon. Branch `fundamentals`
(worktree `.worktrees/fundamentals`). Probe kits live in
`tests/output/fundamentals/<probe>/` (gitignored scratch, like
`tests/output/probe_adv/`); each finding is recorded here with its kit.

## 1. Verdict of the initial code review (2026-10-28, no code changes)

Sound (probe-derived, evidence in the code comments): the line model
(`fonts/embed.rs` line metrics, external leading by the OS/2 definition,
East Asian 1.3× leading, exact/at-least boxes, grid snapping), missing-font
substitution, small caps, strikeout, punctuation squeeze and pull, widow and
orphan counting, cell margin defaults, the list-marker line rule.

Fitted on one or two documents, or never probed, ranked by expected gain:

1. **Horizontal glyph placement** (every online reference). Layout widths
   are proven exact; Word draws each source run on a 0.25pt grid, ~0.07%
   narrower than exact, face-dependent, and resets at run boundaries that
   `docx/runs.rs merge_compatible_runs` erases before layout. A fitted
   in-run shrink (+0.6 J mean) is parked. Section 2 below.
2. **Vertical 0.25pt grid.** Word's online export rounds every baseline to
   0.25pt (with a phase); ours are exact floats. uk_commercial's note block
   is 0.67pt short from 0.05–0.25pt per-line residues, swiss_materials misses
   a page bottom by 0.46pt. `snap_baseline_288` never reached main.
3. **Float-wrap thresholds**, each from one or two documents or unsourced:
   18pt minimum strip for empty paragraphs (`pdf/mod.rs` `MIN_EMPTY_STRIP`),
   48pt "text can sit beside" (`pdf/mod.rs`, the paragraph look-ahead), 72pt
   minimum wrap width (`min_wrap_w`), 20% line-overlap rule (annotation #120).
4. **Run-level constants never probed:** superscript +0.35 em and subscript
   −0.14 em (`pdf/layout.rs vert_y_offset`), highlight box −0.2 to +1.15 em,
   0.65 super/sub size fallback. OS/2 carries superscript offsets.
5. **Table autofit** `natural × 0.9` for nested tables
   (`pdf/table_layout.rs`), and `apply_pct_width` keeping grid proportions
   where Word re-splits by content (150:425 came out 200:350).

Structural, no probe needed:

- Table cells and footnotes build paragraphs by hand
  (`docx/tables.rs`, `headers_footers::parse_notes_simple`, "tuned against
  the existing corpus"); rules added to `paragraph::build_paragraph` miss
  them. Also layout-accuracy.md §5 item 9.
- Cell heights subtract "dangling leading" only for substituted fonts
  (`pdf/table.rs`): a correction layered on a fallback.
- `compute_bookmark_positions` guesses 14pt per table paragraph; only an
  empty PAGEREF result sees it.
- CLAUDE.md's baseline formula ("slot_top − font_size, aligns cap tops")
  no longer describes the code (ascender-based).

## 2. Probe 1: how Word places glyphs inside a run

### What is already known (`current_focus/online-references.md` §4a)

- Layout widths (line breaking, run starts) are the exact font widths.
- Inside a run the PDF draws each glyph step on a 0.25pt grid (288 dpi): a
  constant step (the exact width rounded to 0.25) with ±0.25 corrections.
- The corrections do not follow the drift of the exact cumulative position
  (uncorrected steps occur at drifts from −0.6 to +0.35pt).
- Over 114k glyphs, in-run positions average ~0.07% narrower than exact:
  TNR −1.0‰, Calibri −0.8, Cambria −0.6, Arial −0.4, Aptos −0.2, Arial Bold 0.
- The last glyph of a word absorbs a correction; word starts drift by a
  font-dependent amount (TNR 12 −0.27pt, Arial 12 −0.12, Aptos ≈ 0).

### Hypotheses the probe separates

- **H-A, cumulative rounding of exact positions** (prefix-invariant, mean
  step = exact). Already contradicted by the drift data; the probe
  re-checks it cheaply.
- **H-B, GDI-style per-run distribution:** each glyph has a device width
  (exact rounded to 0.25pt); the run's total device error against the
  exact total is spread as ±1px corrections at even intervals through the
  run. Predicts correction positions that depend on the run length and the
  3/4 periods seen for TNR 12 "e" (25 × 21.3px: 8px of error over 25).
- **H-C, truncated widths:** device positions accumulate the glyph widths
  truncated to whole 1/1000 em (the PDF /Widths integers), snapped to the
  grid. Predicts a per-glyph shrink equal to the width's fractional part
  (TNR 12 "e" 443.85 → 443: −0.19%; "o" 500.0: 0; "i" 277.8: −0.29%) and
  a face dependence that matches the ‰ list above (Arial's widths are
  near-integers in 1/1000 em, TNR's are not).
- **H-D, hinted device widths** at 288 dpi distributed per run.

### Probe document (`tests/output/fundamentals/glyph_grid/mk.py`)

Letter, 1" margins, one run per paragraph, 5 faces × 10/12pt:

- A. long single-glyph runs (e o i s m t n a; 45–80 copies): the mean
  step per glyph to ±0.003pt decides H-C against H-A/H-B.
- B. "e" × 5, 10, 20, 40, 70 in separate paragraphs: prefix invariance
  decides cumulative (H-A/H-C) against per-run distribution (H-B).
- C. alternations "eo", "ei", "ie", "em": how mixed widths share the
  corrections.
- D. "eeeeeeee" × 8 with single spaces in one run: whether a word end
  resets the accumulation, and whether the space carries it.

Export: `python3 tools/word_export.py tests/output/fundamentals/glyph_grid/glyph_grid.docx`
(default online preset). Positions: `tests/output/probe_adv/pdfglyphs.py`
(simple-font TJ walker), analysis `an.py` in the kit.

### Findings (2026-10-28): the rule, derived

Word's online export places glyphs on a 288 dpi device grid (1 px = 0.25 pt)
by a cumulative, chunked, one-glyph-look-ahead rule. `sim.py` reproduces all
8,824 glyph steps of the probe except the 155 of two glyphs whose device
width is not the nearest pixel (item 2). On real references (`validate.py`,
line by line): mandated_reporter p1 18 wrong advances of 3,211,
dutch_council p1 4 of 361 (all four in the two glyph-width exceptions below).
H-A (plain rounding) and H-B (per-run distribution) are dead: the B rows are
prefix-invariant and the mean step is below the exact width. H-C (truncated
widths) was wrong too: the shrink is the PDF writer's (item 5).

1. **Ideal positions** are the exact font advances accumulated across the
   whole line. Run boundaries play no part: a run's `Tm` is the writer's
   running pen, not a layout position (a 26-glyph run starts 0.85 px off
   the exact cumulative, its single leading space is its own segment and
   advances by the device width). The line's first glyph anchors the grid.
2. **Device advance** of a glyph = its exact advance in px rounded half up.
   Exceptions seen: Cambria 12 pt "o" (25.48 → 26), Aptos 12 pt "i"
   (11.46 → 12), Calibri 11.5 pt "E" (22.45 → 23) and "I" (11.59 → 11),
   consistently for every copy: the hinted advance at that ppem. FreeType's
   v35 and v40 interpreters reproduce none of them (`hint288.py`), the
   fonts carry no `hdmx` record for 40/46/48 ppem, and the rounded
   `/Widths` value does not explain them either. Roughly 1 glyph in 25.
3. **Correction**: glyph k keeps its device advance unless the glyph after
   next, placed with device advances, would land 2 px or more from the
   half-up rounded ideal (both measured from the chunk start); then k's
   advance moves 1 px toward the ideal. The run end stands in for the glyph
   after next when k+1 is the last glyph. (A plain error threshold fits
   only with T = 1.5 − 2|δ| per glyph; this rule is why.) Ties: half-up
   rounding of the ideal makes the two directions asymmetric, and the
   probe's exact-tie glyphs (Aptos 10 "a" = 21.25 px, Cambria 12 "s" =
   20.625) confirm it.
4. **Chunks**: a glyph whose rounded ideal end lies 1024 px (256 pt) or
   more past the chunk's first glyph starts a new chunk at round(ideal);
   the error is measured from that start. 1024 is sharp: 1000 or 1040
   quadruples the mismatches.
5. **PDF writer**: `/Widths` = round(exact × 1000). A run whose device
   advance is below W is tightened with a negative `Tc`; otherwise each
   advance is written as W − adj with adj = ceil(W − D) + [D > W] (D the
   device advance in 1/1000 em), so the drawn advance is one unit
   (size/1000 pt) short whenever D > W. Over a line this is the
   "~0.07 % in-run shrink" of online-references.md §4a: a rendering
   artifact of the writer, face-dependent only through how often a glyph's
   nearest pixel lies above its rounded width. Not a layout effect.

Also measured (`half_sizes.pdf`, mk2.py/an2.py): 10, 10.5, 11, 11.5 and
12 pt give the same picture; the "X after 25 copies" method of §4a measures
a device position, so its ±0.005 pt was really ±0.015 pt per glyph. The
device model itself is the sharper instrument: a quantised ideal width
(twips, 0.025 pt) would change the correction cadence, and it does not.

Not covered by this probe: justified lines (czech_census p1 fails 146 of
647 — 12 paragraphs are `w:jc both`, so the ideal space widths differ),
kerning pairs (`w:kern`), character spacing (`Tc`), tabs, RTL. Each needs
its own probe before the rule is applied there.

### What it means for the engine

- Line breaking keeps exact widths. Correct as is.
- Drawing: per line, run this rule over the glyph sequence and emit the
  device positions (TJ) instead of exact advances, in
  `src/pdf/layout.rs` (`show_kerned` / `show_placed`). Expected gain: the
  ~0.38 pt per-line word drift against every online reference (§4a), the
  largest known residual. Whether to also mimic item 5 (the writer's unit
  loss) is a choice: it is an artifact, but the references carry it.
- The hinted-advance exceptions need a TrueType hinting VM to reproduce
  (skrifa would be a new dependency; FreeType does not match anyway).
  Nearest pixel first; measure what the exceptions cost.
- `merge_compatible_runs` is fine: run boundaries do not matter to the rule.
  (They do to the writer; see the corrections below.)

Kit: `tests/output/fundamentals/glyph_grid/` — `mk.py`/`glyph_grid.docx`
(A/B/C/D rows), `mk2.py`/`half_sizes.docx`, `an.py` (per-row steps),
`sim.py` (the rule vs the probe), `validate.py <ref.pdf> fonts [page]
[-v|-d|-j|-t]` (the rule vs a real reference, line by line), `widths.py`
(the `/Widths` Word wrote), `hint288.py` (FreeType hinted advances).

### Corrections and implementation (2026-10-30)

Checked against every online fixture reference (pages 1–2, 321k in-run
steps on unstretched lines; kit scripts `thresh*.py`, `boundary.py`,
`aggb.py`, `xcompare.py`), not only the probe. Three items above were wrong
or incomplete:

- **Item 2, device advance: rounds up from 0.46 px, not 0.5.** Grouping the
  clear-cut glyph classes by the fraction of their exact px advance, 235/512
  (0.4590) goes down and 236/512 (0.4609) goes up in every face and size
  (Cambria 11 "e" 21.463 → 22, Arial 11 "a" 24.471 → 25, Calibri 11 "A"
  23.459 → 23). Wrong advances: 12.9 % → 6.6 % (Cambria 11 26.9 → 0.5 %,
  Aptos 12 31 → 1.5 %, Arial 11 42 → 10 %). The probe's "hinted" Cambria 12
  "o" and Aptos 12 "i" are this rule; the probe now fails 26 of 10,010
  steps. The rounded ideal in the correction (item 3) and the chunk starts
  stay half-up (8.3 % if they take the 0.46 too). A scale factor fits worse
  (k = 1.002: 7.4 %). Not hinting either: the vendored Cambria has no
  `fpgm`/`prep`. The genuine hinted exceptions left are few (Times New Roman
  12 "F" 26.70 → 26, "G" 34.66 → 34; Arial 11 "g"/"q" stay down).
- **Item 5, writer.** Fitted on observed device steps the writer writes, for
  each advance, adj = round(W − D) half away from zero when W > D, one more
  when W < D, nothing when |W − D| < 0.5 (`models_seg2.py`). Small gaps
  (|W − D| ≤ 3) are sometimes left uncorrected in Cambria 11, but that dead
  band fits worse once D is predicted (0.083 vs 0.066 pt), so it is left out.
- **Item 1 / run boundaries.** The writer restarts (new `Tm`) at Word's own
  run boundaries, and those are where adjacent source runs differ in
  `w:rsidR` or `w:rsidRPr` (probe 3, `tests/output/fundamentals/runsplit/`:
  per-word runs with no or equal rsids give one `Tm` per line, alternating
  ones a `Tm` per word, even mid-word). So `merge_compatible_runs` must
  keep those boundaries: `Run::source_breaks`/`Run::rsid`. A new `Tm` lands
  0.05–0.2 pt right of the device position, growing along the line; no
  quantisation of the advances reproduces it (`tm4.py`), so the engine
  starts it on the device grid.

Pooled in-run error against the references (`aggb.py`, glyphs anchored at
Word's run starts): exact advances 0.146 pt, device grid 0.112, device grid
plus writer 0.066; better than exact in every face/size bucket.

Implemented in `src/pdf/glyph_grid.rs` (the rule) and `device_grid` in
`src/pdf/layout.rs` (which chunks, sequences of chunks set edge to edge or a
plain space apart, run starts). Applies to lines without justification,
kerning, letter spacing, text scaling, RTL or East Asian text. Line starts
take Word's +0.025 pt phase.

### Justified lines (probe 2, `tests/output/fundamentals/justify/`)

Five faces × left/justified/per-word runs/double spaces/long words/"eee"
words. Word draws a justified line as one `Tm` (with per-word runs too).
The rule is probe 1's with two additions (`just10.py`: 70 wrong of 5,777
steps, most of them the step into the line's trailing space):

- the ideal positions are the exact advances with each space widened by
  an even share e of the slack, e taken from the layout (measure minus
  natural width over the spaces). Solving e from the drawn last glyph
  instead (`just2..9.py`, 20–27 % wrong) is what hid the rule: that glyph
  carries the writer's drift;
- a stretched space's device advance is its own (0.46 rule) plus e rounded
  half-up apart: Arial 10 11.11 + 11.47 px → 11 + 11 = 22, not 23.

Corrections then fall where the look-ahead rule puts them, spaces included.
Implemented: `GridGlyph::stretch`, and `device_grid` joins chunks across
plain spaces plus the line's `extra_per_gap`. 24 justified fixtures: J
66.97 → 68.89, SSIM 87.13 → 89.05 (czech_expert_witness_law J +10.6).

### Kerned text (probe 4, `tests/output/fundamentals/kern/`)

Five faces, prose with kerning pairs, "AVAV…", "To To …", and the prose
unkerned as control (`kan.py`, `kan3.py`; pair values from GPOS or the
legacy `kern` table by `kernpairs.py`). Word has two paths for kerned text,
switched by the `enableOpenTypeFeatures` compat setting, not by the
compatibility mode (the probe exported at modes 12, 14 and 15 with the
setting all go one way, the mode-15 probe without it the other):

- **With it**: kerned runs are written with a CID font (glyph ids,
  Identity-H), Word's shaping path. The kern joins the first glyph's
  advance and the whole advance rounds **half-up**, for every glyph of the
  run, spaces included (Aptos 12 "o" 26.48 → 26, Cambria 11 "e" 21.46 →
  21): prose 3.5 %, "To" rows 2.8 % wrong; 0.46-rounding on the same lines
  12 %.
- **Without it** (learning_cultures_dissertation, `kern_noot.pdf`): plain
  WinAnsi fonts, kerned all the same (TJ "A)47(T"); the glyph keeps its
  0.46-rounded advance and the kern is rounded half-up on its own and
  added, like a justified space's stretch: 3.5 % / 3.8 % wrong (combined
  rounding 17–20 %).

Unkerned rows keep the 0.46 rule on either path. Writer: the same rule
(TNR "AV": W − D = 138.7 → TJ 139). Text outside WinAnsi (Cyrillic in
underlined-space-compat14) is written with a CID font even unkerned; its
rounding is not probed.

Implemented: `GridGlyph::shaped` / `GridGlyph::apart`, `TextLine::
shaped_kerning` from `DocumentSettings::open_type_features`. 28
`w:kern="2"` fixtures: J 66.07 → 67.37, SSIM 85.49 → 87.15, none worse
(case1 J +6.3, case18 +8.2).

### Other scripts and letter spacing (probe 5, `tests/output/fundamentals/script_tc/`)

Five faces, unkerned: Cyrillic, Greek and Polish prose (CID fonts, outside
WinAnsi), and Latin prose with `w:spacing` +20, −10 and +3 twips (`an.py`).

- Text outside WinAnsi keeps the 0.46 rule on the CID path (Cyrillic 4 %,
  Greek 2.6 %, Polish 1.5 % wrong; half-up is worse or equal): the half-up
  rounding belongs to OpenType kerning, not to CID output.
- Letter spacing is written as TJ moves (no `Tc`) and rounds apart like the
  justification stretch: glyph 0.46-rounded plus the spacing rounded
  half-up (+0.15 pt: 11 % wrong against 19–23 % combined; whole-pixel
  spacings fit every variant). Spaces carry the spacing too, so a spaced
  line is one sequence: restarting at every word (our spaces were not
  "plain") reset the writer's drift that Word keeps, up to 1 pt by the
  line end. Probe: glyph error 0.165 → 0.029 pt, J 72.5 → 81.4.

## 3. Queue after probe 1

1. Text scaling (`w:w`) on the grid.
2. Items 2–5 of §1, in that order; each gets its own probe kit and section.
