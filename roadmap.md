# Roadmap

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

`w:autoHyphenation` and `w:suppressAutoHyphens` are parsed from DOCX. `w:lang` is parsed on runs. However, **Word's online PDF converter does not perform syllable-break hyphenation** for any language, even when `autoHyphenation` is set. Every line-ending hyphen in reference PDFs comes from pre-existing hyphens in compound words (verified across 8 language-specific fixtures and all scraped fixtures).

The `hyphenation` crate (Knuth-Liang algorithm) was tested with 8 languages but its dictionaries don't match Word's — enabling it caused 40-50pp regressions across all languages because line breaks diverged. Removed for now.

**Prerequisites to revisit:** Reference PDFs generated by desktop Word (not the online converter) with proofing tools installed. The online converter ignores `autoHyphenation` entirely.

Test fixtures in `tests/fixtures/hyphenation/` (8 languages with Wikipedia text, `autoHyphenation` enabled) are ready for comparison when desktop-Word references become available.

## Image Cropping `a:srcRect` (TODO — LOW-MEDIUM IMPACT, found 2026-09-04, planned)

Not parsed. Cropped pictures render the full source squeezed into the frame, so both the
visible region and aspect ratio are wrong. Only 1 of 129 corpus fixtures has real crop
values: brazilian_logistics_study (J 17.3%), 3 anchored PNGs cropped 4.5–26.6% top/bottom.
12 other fixtures carry an empty `<a:srcRect/>`, which must stay a no-op.

Plan: parse l/t/r/b (1/100000 fractions) onto `EmbeddedImage.src_rect`; in
`embed_single_image` wrap the image XObject in a Form XObject with BBox `[0 0 1 1]` whose
content draws the inner image with `cm = [1/(1-l-r), 0, 0, 1/(1-t-b), -l/(1-l-r), -b/(1-t-b)]`.
The bbox clips, so negative crops (blank padding) work for free and every draw site
(inline, floating, table, header/footer, textbox) is untouched. Verify with unit tests on
the parser and matrix, a handcrafted case (needs Word reference), and an unchanged
`latest_hashes.json` for every fixture except brazilian_logistics_study. Found while
reviewing MiniPdf — see `minipdf.md` for that and the other items it has that we lack (TOC
generation, sdt data binding, `lastRenderedPageBreak` hint, score-gated fix loop).

## Picture Effects (PARTIALLY DONE)

**Done:** Smooth outer shadow (rasterized Gaussian blur mask via SMask), soft edge (edge-fade SMask on image), glow (centered blur), inner shadow (inverted blur mask), reflection (flipped image with gradient SMask). All use the same rasterized mask + SMask XObject infrastructure. Test fixtures: case56 (shadow variations), case57 (2D effects), case58 (3D effects — deferred).

**Remaining (deferred — no real-world fixtures use these):**
- **3D effects** (`a:scene3d`, `a:sp3d`) — bevel, metal frame, perspective rotation. Would require 3D lighting simulation. case58 has test fixtures ready.
- **Preset shadows** (`a:prstShdw`) — 20 built-in shadow presets. Need mapping table from preset names to parameters.
- **Theme color resolution in effects** — inner shadow with `a:schemeClr` falls back to black instead of resolving the theme accent color.

## CJK Rendering Polish (TODO — MEDIUM IMPACT)

Core CJK support is implemented: CIDFont/Identity-H/ToUnicode encoding, platform-specific font fallback chains (Hiragino/Noto/Yu Gothic), per-character font fallback at render time, script-based run splitting via `w:rFonts @eastAsia`, and vertical text rendering. CJK fixtures render readable output but score low (4-9% Jaccard) due to spacing/positioning precision issues:

1. **`w:firstLineChars`** (MEDIUM) — character-based indent (e.g. `firstLineChars="100"` = 1 character width). Not parsed; we only handle `w:firstLine` (twip-based). In practice, twip fallback is always present alongside firstLineChars.
2. **Vertical text centering** — `render_vertical_cjk_cell` uses a simplistic height calculation (chars x font_size) that doesn't account for paragraph spacing, causing vertical misalignment in merged cells.
3. **Fallback line-height fidelity** (MEDIUM, annotation #8) — Korean fonts
   (함초롬바탕, HY헤드라인M, 굴림) route to a fallback whose line ratio is ~1.27
   vs ~1.73 for the fonts Word used in the reference. Every `lineRule="auto"`
   line and every `atLeast` table row comes out short (east_asia_conference_form:
   ~116pt lost over one table, flipping a page break). Fix alongside Bundled
   Fallback Fonts: pick/ship CJK fallbacks with matching vertical metrics, or
   apply a per-script line-height compensation.

## Bundled Fallback Fonts (TODO — MEDIUM IMPACT)

We rely entirely on system fonts and fall back to Helvetica Type1 as a last resort. This produces inconsistent output across environments (servers, Docker, CI). Should bundle metric-compatible open fonts behind a feature flag:
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

## Table Row Height Deficit (TODO — MEDIUM IMPACT, discovered 2026-07)

Our table rows run ~0.5–0.9pt shorter than Word's, compounding down a page of
stacked tables (case51: −6.5pt accumulated over 3 tables, measured via stext
anchor diffing). Consequence: content that Word pushes to the next page can
stay on ours. case51's reference has a blank page 2 (Word's implicit final
paragraph mark spills after the doc-ending table at 710.9pt); ours ends 8.2pt
higher so the mark fits and no page 2 is emitted — this alone costs case51
~22pp SSIM (missing page scores 0). The implicit final-¶ model and the
end-of-cell-mark suppression after nested tables are already in (2026-07);
only the per-row height accounting remains.

## Header Multi-Float Wrap (DONE — 2026-07-02, annotation #212)

`hdr_fz` is now a Vec of zones; all wrapping floats (same-paragraph + earlier
paragraphs) constrain the text bounds together, and paragraph indents are
measured from the column edge with float bounds clipping (Word semantics).
`parse_object_floating_image` honors `w10:wrap type="square|tight|through|
topAndBottom"`. Letterhead center now within ~5pt of reference. Remaining:
HR `o:hrpct` width should use the indent-adjusted paragraph box.

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

## Unimplemented Run Properties

### `w:emboss` / `w:imprint` / `w:shadow` (TODO — MEDIUM IMPACT)

636 hits across fixtures, 1 failing fixture. These are WML text effects (mutually exclusive per spec):
- **`w:emboss`** — raised/embossed appearance (highlight color on top-left, shadow on bottom-right)
- **`w:imprint`** — engraved/debossed appearance (inverse of emboss)
- **`w:shadow`** — drop shadow on text (offset copy in shadow color)

Not parsed, not rendered. Trivially implementable: parsing is `wml_bool`, rendering is offset/color-shift drawing passes.

### `w:outline` (legacy) (TODO — LOW IMPACT)

The legacy WML `w:rPr/w:outline` element (hollow text, no fill) is not parsed. We handle the modern `w14:textOutline` but not the pre-Word 2010 equivalent.

### `w:shd` on runs (TODO — LOW IMPACT)

Run-level shading (`w:rPr/w:shd`) is not parsed. We handle paragraph-level and cell-level `w:shd` but not run-level. Different from `w:highlight` (named colors) — `w:shd` supports arbitrary hex fill colors and patterns.

## Unimplemented Paragraph / Layout Features

### `w:jc val="distribute"` (TODO — MEDIUM IMPACT)

Distribute alignment (equal spacing including edges, different from justify). Currently silently treated as left-align — should at minimum fall back to justify.

### `w:mirrorMargins` (TODO — MEDIUM IMPACT)

Parsed from `word/settings.xml` and stored in `DocumentSettings.mirror_margins`, but **never applied to layout**. Fix: swap `margin_left`/`margin_right` on even-numbered pages.

### `w:gutter` (TODO — LOW IMPACT)

Gutter margin (`w:pgMar @gutter`) is not parsed. Adds extra space on the binding side for printed documents.

### `w:pgBorders` (TODO — LOW IMPACT)

Page borders (decorative borders around entire page) are not parsed or rendered. Defined in `w:sectPr/w:pgBorders` with per-side border definitions.

### `w:vAlign` on `sectPr` (TODO — LOW IMPACT)

Vertical alignment of text on the page (top/center/bottom/both). Not parsed from section properties. Mainly affects title pages and short documents.

### `w:textAlignment` (TODO — LOW IMPACT)

Vertical alignment of runs within a line (top/center/baseline/bottom/auto). Only superscript/subscript are handled; the paragraph-level `w:textAlignment` property for mixed-size runs is not.

### RTL / BiDi (TODO — HIGH EFFORT, MEDIUM IMPACT)

`w:bidi` (paragraph-level) and `w:rtl` (run-level) right-to-left support is completely absent. Requires implementing the Unicode BiDi algorithm (UAX #9) for correct visual reordering. Architecturally complex — affects line building, text rendering, and alignment.

## Unimplemented Table Features

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

### `w:tblLook` / `w:tblStylePr` (TODO — MEDIUM IMPACT)

Table conditional formatting (firstRow, lastRow, firstCol, lastCol, banded rows/cols). The table style is resolved for default borders but conditional formatting overrides (bold headers, alternating row shading, etc.) are not applied.

### Table auto-fit vs `tblW` (NO IMPACT — corpus check 2026-05)

Our `auto_fit_columns` uses `gridCol` widths from `tblGrid`, ignoring the specified `tblW` when `type="dxa"`. Word treats `tblW` as the authoritative total width and scales/caps columns to fit. This causes tables to render at full page width when python-docx (or other generators) emit oversized `gridCol` values alongside a smaller `tblW`.

**Verified empty in current corpus**: a sweep of all `tests/fixtures/scraped/*` and `tests/fixtures/new/*` documents found zero tables where `gridCol` total exceeds the `tblW` value (tolerance 100 twips). The bug is real per OOXML, but no fixture triggers it — implementing this clamp moves zero scores. Park until a real-world fixture exhibits the mismatch.

### Percent-based widths: `tcW`/`tblW` `type="pct"` (PARTIALLY DONE 2026-06)

`twips_attr` reads `w:w` as twips regardless of the `w:type` attribute. For `type="pct"` the value is in fiftieths of a percent (5000 = 100%). **Implemented**: `Table.width_pct` is parsed from `tblW type="pct"` and `apply_pct_width` scales columns to pct × content width — but ONLY for tables whose `tblGrid` is missing (grid inferred from row `tcW` values, which preserves pct proportions). When a real tblGrid exists, Word renders the grid widths as-is even when the pct width disagrees (observed: arizona 115%, zimbabwe 100% vs grid at 102% of content — scaling them regressed scores). **Remaining**: `tcW type="pct"` is still mis-read as twips for per-cell preferred widths; harmless today because grid widths dominate, but would matter for Word's full preferred-width algorithm (§17.18.87).

## Unimplemented Document Features

### Endnotes (TODO — MEDIUM IMPACT)

`w:endnoteReference` is completely unimplemented. Footnotes already work — the plumbing (reference parsing, content parsing, rendering at page bottom) exists and could be adapted. Endnotes collect at the end of a section or document rather than at the page bottom.

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

**Remaining:**
- **Floating images don't participate in z-order** — they still paint inline at their anchor
  paragraph; e.g. lenten's white bird icon is covered by the purple band (icon z=251658243 >
  band 251658241). Same deferral treatment as textboxes/connectors would fix it.
- **behindDoc shapes from later paragraphs** can still paint over earlier paragraphs' text
  (needs pre-pass/paginator).
- **Group flips/rotation** — group-level flipH/flipV and rot are ignored (rare); leaf connector
  flips work.
- **Canvas/group inside paragraph-level mc:AlternateContent** — only the run-level path
  flattens groups; `collect_textboxes_from_paragraph` still grabs the first wsp.
- **Text in preset shapes placement** (case35 annotations) — text inside rightArrow etc. is
  positioned with plain rect insets, not the shape's text rectangle.

## Floating Image Positioning (TODO — MEDIUM IMPACT)

Floating images (`wp:anchor`) with large `posOffset` values can render off-page. Word appears to clamp or reflow these positions, but we render at the raw coordinates. Observed in `learning_cultures_dissertation` (rId14: column-relative offset 4702029 EMU = 370pt, placing a 334pt-wide image past the 612pt page edge). A naive right-edge clamp was tested but regressed `stem_partnerships_guide` — a more nuanced approach is needed (possibly only clamping when the image would be entirely off-page, or respecting wrap constraints).

Additionally, truncated/corrupt PNG images in DOCX files cause the `image` crate to fail with "unexpected end of file". Currently falls back to a 1x1 placeholder via `decode_png_raw` (using the `png` crate directly). Word renders these partially — investigate partial PNG decoding to match. Observed in `learning_cultures_dissertation` image1.png (216KB file, 2205 bytes short of complete IDAT data, no IEND chunk).

## `w:smallCaps` Rendering Accuracy (DONE — verified 2026-05)

`smallcaps_segments()` in `src/pdf/layout.rs:349` already applies the per-character rule correctly: only originally-lowercase characters are uppercased and rendered at `font_size - 2pt`; originally-uppercase characters render at full size. Unit tests at `src/pdf/layout.rs:1824+` cover mixed/upper/lower/non-letter cases.

## SmartArt Remaining Work

Basic fallback rendering via pre-flattened `dsp:drawing` shape trees is done, with full geometry engine support (all 187 preset shapes). Remaining:

1. **Group shapes** (MEDIUM EFFORT) — `dsp:grpSp` groups with nested transforms. Need recursive parsing.
2. **Connector shapes** (MEDIUM EFFORT) — `dsp:cxnSp` connectors between shapes (arrows, lines).
3. ~~**Image shapes**~~ (DONE) — `a:blipFill` image fills parsed from diagram-specific relationships, rendered with cover-fill scaling and shape clipping.
4. **Full layout engine** (VERY HIGH EFFORT) — implement the constraint-based layout algorithm that interprets ~200 XML layout recipes. Only needed for files that lack the `dsp:drawing` fallback. Not planned for the near term.

## Charts Remaining Work

All 8 chart types are supported (bar, line, pie, area, doughnut, radar, scatter, bubble). Remaining:

- **3D charts**: `c:bar3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` — not parsed
- **Stock charts**: `c:stockChart` — not parsed
- **Combo charts**: two chart types overlaid on the same plot area — not handled
- **Stacked bar rendering**: parsed but rendering treats as clustered
- **Data labels**: not parsed or rendered
- **Chart title**: not parsed or rendered
- **Secondary axes**: not handled
- **Chart label positioning**: axis labels still have small offsets vs Word. `text_width_approx` (len x fs x 0.5) is crude — real font metrics would help.
- **Legend placement fine-tuning**: small positional offsets vs Word. Centering formula and spacing need per-chart-type calibration.
- **Font selection in chart labels**: picks arbitrary font from seen_fonts, not theme font

## Track Changes Remaining Work

Final mode (insertions included, deletions removed) is done. Remaining:

- **Markup mode** — rendering deletions with red strikethrough, insertions with red underline (for documents exported with markup visible)
- **Paragraph-level changes** — `w:ins`/`w:del` wrapping entire `w:p` elements at `w:body` level
- **Property changes** — `w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tblPrChange` (formatting revisions)

## WordArt Remaining Work (LOW IMPACT)

Levels 1-4 are done (flat rendering, text effects, envelope warping, text-on-a-path).

**Level 5 — Legacy VML enhancement (TODO):** VML fill types (gradient/pattern), VML shadow, VML shapetype-to-prstTxWarp mapping. Basic flat rendering already done in Level 1.

## Image Drop Shadow Quality (TODO — LOW IMPACT)

Basic drop shadow rendering is implemented (`a:effectLst/a:outerShdw`): offset, color, alpha, and directional soft edge via layered transparent rectangles. Current limitations:
- **No real gaussian blur** — approximated with 10 stepped layers, visible banding at close zoom
- **Fallback paths lack alpha** — inline images in text lines, floating images, table/header images use pre-blended solid color instead of PDF ExtGState transparency (only body-level paragraph images get proper alpha)

## Bullet Line-Height Drift on macOS (TODO — font-metric blocked, found 2026-06)

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

- **Line spacing** — Auto and Exact work. AtLeast parsed but may not enforce minimum correctly.
- **Tab stops** — basic left/center/right tabs work but decimal alignment has precision issues.
- **Panose font matching** — fontTable.xml contains panose classification bytes; could use for more precise font substitution.

## Floating Image Wrapping — Remaining

- **wrapSquare height reserve gated on side-strip width (DONE — 2026-07-02)**: the anchor-paragraph reserve in `render_paragraph_block` now fires only when no usable side strip remains (`MIN_EMPTY_STRIP` = 18pt). With a real strip (brazilian_logistics_study, ~42pt) empty spacer paragraphs absorb through the float's span and the next real paragraph is displaced to `fz.bottom_y` — the old 48pt threshold double-counted the image height there. With no strip (sample500kB, image width == column width) Word stacks everything below, which the reserve reproduces. Note Word actually puts the anchor's own line box below the float too (ref gap 67.7pt vs our 51.8pt on sample500kB p4) — a first displacement attempt lost inter-paragraph gaps; revisit with the paginator.
- **`w:br type="textWrapping" clear="all"` (DONE — 2026-07-02, annotation #111; refined 2026-07-03, annotation #218)**: parsed into `Paragraph.clears_floats`; block loop drops the cursor to the float-zone bottom after such a paragraph. 2026-07-03: the cursor now drops one line height *below* the float bottom — the line following the break (the break paragraph's mark line) still occupies its full line height there, matching Word's ~16pt gap on indonesian_benchmarking_guide p7. Approximation: clear applies after the whole paragraph, not mid-paragraph (fine when the break is alone in its ¶, the common Word idiom).
- **Multiple floats per paragraph (PARTIALLY DONE — 2026-06)**: When one paragraph anchors 2+ wrapping floats (e.g. a logo on each side of a centered title, `pendulum_mechanics_oscillation_lab`), per-line geometry now subtracts every float's exclusion span and places text in the widest gap. Limitation: the page-level `float_zone` for *subsequent* paragraphs still tracks only the first float, so a following paragraph that overlaps only the second float won't wrap around it.
- **Remaining y-shift (page 2 only)**: Word places page 2's image (180x144pt) 14.8pt higher than all other images, despite identical `posOffset=0`. Pages 1,3,4,5,7 match perfectly (delta <0.02pt). Pages 2 and 6 (both cy=1828800/144pt) are the outliers. Likely Word snapping to grid/text boundaries based on image dimensions.
- **Look-back wrapping (TODO — MEDIUM IMPACT)**: Paragraphs BEFORE the image anchor cannot wrap beside the image because the float zone isn't set until the anchor paragraph renders. In Word, text from preceding paragraphs also wraps (e.g. case41 page 3 — the first paragraph's lower lines should wrap beside the centered image). Requires either a paginator or a two-pass layout with look-back.
- **Image in text paragraph**: Case41 page 6 — last line of text paragraph overlaps the image. Look-ahead only fires for the NEXT block, not same-paragraph floats.
- **Tight vs Through distinction**: Both currently use convex-hull polygon scanline. For Through wrapping, text should fill polygon concavities. Requires returning per-line interval segments instead of hull bounds. Rare in practice.
- **Word-break precision**: BothSides wrapping produces correct structure but slightly different word breaks from Word, causing ~2pp Jaccard differences on case41.
- **Polygon wrap text distribution**: Case42 (wrapTight + BothSides + complex 53-vertex polygon around Mario) scores ~46% Jaccard. Zone overlap detection is correct but line breaks differ from Word — likely font metric differences for Times New Roman causing different left/right text distribution. Text near concave polygon areas (Mario's arm) appears visually close to the image despite respecting the 9pt distL margin.

## Code Structure

### Duplication & extraction sweep (DONE — 2026-06-21)

Whole-repo over-engineering/duplication audit applied — see `extraction-audit.md`
for the full findings. All 30 verified survivors landed across 11 commits with
zero rendering regressions (208/208 scores unchanged). Highlights:

- **Shared parse helpers** in `docx/mod.rs`: `parse_on_off` (ST_OnOff), `parse_pt`
  (VML/CSS lengths), `is_wml` (namespace predicate), `merge_tab_stops`; plus
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

### Refactor `pdf/mod.rs` `render()` (see "Paginator Extraction")

The `render()` function in `pdf/mod.rs` is ~2400 lines with many closures and shared mutable state. The right fix is the paginator extraction described above. (Smaller in-file extractions are already done: `embed_single_image` is a free fn in `pdf/images.rs`; `label_for_paragraph` lives in `pdf/list_label.rs`.)

### Image-embedding cleanups (LOW IMPACT — deferred from textbox-image work)

Small consistency / efficiency wins in `pdf/images.rs` that were considered but skipped to avoid scope creep when adding textbox-internal image rendering:

- **Global Arc→pdf-name registry to dedupe XObjects across maps.** Same image data used in body + textbox (or table cell + textbox) is currently embedded as two separate PDF XObjects because each map (`inline_image_pdf_names`, `table_cell_image_names`, `textbox_image_names`, …) keys independently by `Arc::as_ptr`. A single global registry would let the second site reuse the first XObject. Wasteful in theory, accepted limitation in practice.
- **`build_paragraph_lines` / `build_tabbed_line` should accept `&HashMap<usize, &str>`.** The current `&HashMap<usize, String>` signature forces every caller (body, header/footer, table cell, textbox) to `.clone()` pdf names into a fresh per-paragraph map. Borrowing would eliminate the clones, but ripples through `pdf/layout.rs` and every caller.
- **Pair `image_names` + `effect_names` into a struct.** Every embedder (`hf_*`, `table_*`, `textbox_*`) threads the two maps as separate `&mut HashMap<…>` parameters. Pairing them would shrink signatures throughout `pdf/images.rs` and `pdf/textbox_render.rs`, but only worth doing alongside the dedup registry above (otherwise diverges from the established style without enough payoff).

## Performance

### Known Bottlenecks

- **Double font reads** — scan reads each font file for indexing, then `register_font` reads again for embedding. Keep the data from the first read.
- **Repeated WinAnsi conversion** — same text is converted in line-building, rendering, and table auto-fit. Pre-compute once and store in `WordChunk`.
- **String allocations** — `font_key()` allocates on every call; `WordChunk` clones font name strings per word. Use indices or interning.

### Parallelism (rayon)

- Font directory scanning — embarrassingly parallel, biggest win
- Font metric computation — parse face, compute widths per font independently
- Paragraph line wrapping — independent per paragraph once font metrics are ready
- ZIP decompression + XML parsing — read all entries into memory, parse in parallel

### Other

- Compress font file streams with FlateDecode (currently uncompressed)
- Memory usage for large DOCX files with many images

## Scraped Fixture Status

32 passing, 16 failing, 0 skipped out of 48 scraped fixtures. Breakdown of 16 failures by dominant issue:
- **text/layout only**: 8 fixtures
- **anchored images**: 4 fixtures
- **floating tables**: 3 fixtures
- **structured doc tags**: 2 fixtures (SDT content is extracted but wrapping may cause layout shifts)

Run `./tools/target/debug/analyze-fixtures --failing` for current breakdown.

## Test Harness: Surface Conversion Panics Loudly (TODO — HIGH PRIORITY, found 2026-06)

A library panic went unnoticed for an unknown number of runs: `new/construction_bathroom_accessories_spec` panicked in `cell_span_width` on every conversion, but the suite still reported "134 passed" with exit code 0. Three gaps compounded:

1. `tests/visual_comparison.rs` catches per-case panics and emits `[SKIP] <case>: conversion panicked` — visible only in `--verbose` output; the case silently gets no score, so the compact report's "N scored, N unchanged" looks green.
2. `run-tests.sh` greps `thread.*panicked` into a "Panics:" section, but the exit code stays 0 — nothing fails.
3. Conversion worker threads are unnamed, so panic messages show `thread '<unnamed>' panicked at src/...` with no case attribution — diagnosing required a separate verbose run.

Fixes:
- `run-tests.sh`: exit non-zero when the Panics section is non-empty.
- Harness/compact report: count panicked cases as failures and list them by name (`PANIC: new/construction_bat..`) in the compact output.
- Name conversion threads after the case (`std::thread::Builder::new().name(case.clone())`) so panic messages self-identify.

## Test Corpus Expansion

- Deep style inheritance (3+ level chains with run vs style vs paragraph conflicts) — **case50** (awaiting reference PDF)
- Nested tables (tables inside table cells, 2-level and 3-level nesting) — **case51** (awaiting reference PDF)
- Stacked bar chart rendering (stacked + percentStacked, vertical + horizontal) — **case52** (awaiting reference PDF)
- Charts with extreme data (50 categories, small/large/mixed-range values) — **case53** (awaiting reference PDF)
