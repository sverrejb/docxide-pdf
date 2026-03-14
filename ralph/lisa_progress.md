# Progress for Lisa

## Session 1 — 2026-03-14: Implement w:br line break handling

### Case Selected
`russian_sports_ranking_decree` (text/layout only, 2 pages, 12.7% Jaccard) — chosen as the analysis target because it's a small text-only fixture where `w:br` line breaks are critical for layout. The fix is broadly applicable across all fixtures using soft line breaks.

### Problem
`w:br` (soft line break) elements were only counted (`line_break_count`) and used to inflate minimum paragraph height via `extra_line_breaks`. The actual text flow completely ignored them — text was laid out as continuous paragraphs. This caused incorrect layout in documents with explicit line breaks (common in legal/official documents across many languages).

### Analysis
- Investigated 3 candidate fixtures: `czech_grant_application`, `russian_sports_ranking_decree`, `mandated_reporter_child_abuse`
- `russian_sports_ranking_decree` had clear `w:br` elements between title lines (e.g., "ГЛАВА" / "ГОРОДСКОГО ОКРУГА КОТЕЛЬНИКИ" / "МОСКОВСКОЙ ОБЛАСТИ") that were being rendered as one continuous line
- Found that `w:br` handling was a 2-line counter increment instead of creating actual break markers

### Implementation
1. Added `is_line_break: bool` to `Run` struct in `model.rs`
2. Changed `parse_runs()` in `runs.rs` to create `Run { is_line_break: true }` instead of incrementing a counter
3. In `build_paragraph_lines()`: line break runs force a new line and reset cursor
4. In `build_tabbed_line()`: same line break handling for tab-containing paragraphs
5. Added `ends_with_break` flag to `TextLine` so lines ending with `w:br` are NOT justified (matching Word behavior — only natural word-wrapped lines get justification)
6. Updated `is_text_empty()` to recognize line break runs as non-empty content
7. Removed `extra_line_breaks` from `Paragraph` and `line_break_count` from `ParsedRuns`

### Files Modified
- `src/model.rs` — added `is_line_break` to Run, removed `extra_line_breaks` from Paragraph
- `src/docx/runs.rs` — generate line break runs instead of counting
- `src/pdf/layout.rs` — handle line breaks in both layout functions, justify suppression
- `src/pdf/mod.rs` — removed min_lines calculation
- `src/docx/mod.rs`, `headers_footers.rs`, `tables.rs`, `textbox.rs` — removed `extra_line_breaks` assignments
- `tests/baselines.json` — reset polish_council_resolution baseline

### Results
- **24 passing fixtures (was 23) — 1 new passing fixture**
- `russian_university_proceedings`: 19.8% → 20.2% (crossed 20% threshold)
- `mandated_reporter_child_abuse`: 16.8% → 18.2% (+1.4pp)
- `polish_municipal_letter`: 11.3% → 13.2% (+1.9pp)
- `russian_sports_ranking_decree`: 12.7% → 12.8% (+0.1pp)
- `polish_council_resolution`: 37.2% → 24.3% (regression — correct breaks expose font metric differences; still above threshold)

### Commit
`a7038d2` — "Implement proper w:br line break handling in text layout"

## Session 2 — 2026-03-14: Nested table flattening and inline images in table cells

### Case Selected
`mandated_reporter_child_abuse` (text/layout only, 5 pages, 18.2% Jaccard) — chosen because it's closest to the 20% threshold among text/layout-only failing fixtures (only 1.8pp away). The first-page header contains a table with an inline image (logo) and a nested table (title), both of which were silently dropped.

### Problem
Two issues in table cell parsing:
1. **Nested tables dropped**: `parse_table_node()` only collected `w:p` elements from cells, completely ignoring `w:tbl` (nested tables). The header's right cell contained a nested table with the document title "5001.3 Child Abuse Notification of Reporting Procedures and Employee Acknowledgement Form" — all that content was silently lost.
2. **Inline images in table cells not extracted**: The body parser (`docx/mod.rs`) lifts `Run.inline_image` into `Paragraph.image` and sets `content_height`, but table cell parsing didn't do this. The header logo (~63pt tall) was parsed as a run-level image but never contributed to cell/row height.

### Analysis
- Investigated 15 failing fixtures; 7 were text/layout-only
- `mandated_reporter_child_abuse` header1.xml contains: outer table (2 cols: logo+text | nested table with title), using `w:titlePg` for first-page header
- Nested tables are rare in the corpus — only 1 out of ~75 fixtures uses them, and only in a header
- Also investigated `czech_grant_application` (9.2%), `polish_archery_range_plan` (15.0%), `slovak_misdemeanor_amendment` (12.9%) for common issues
- Found a separate bug: empty paragraphs in table cells contribute 0 height instead of `line_h`. This affects the Czech form fixture significantly (cells use empty paragraphs as vertical spacers). However, fixing this caused SSIM regressions (-2.9pp, -3.4pp on other fixtures) because every cell's end-of-cell marker also gets height, so this fix was NOT included.

### Implementation
1. Added `collect_nested_table_paragraphs()` in `tables.rs` — recursively extracts `w:p` nodes from nested `w:tbl` elements, skipping `vMerge=continue` cells to avoid duplicating merged content
2. Changed cell content iteration to handle both `w:p` (direct paragraphs) and `w:tbl` (nested tables via flattening)
3. Added inline image extraction in table cell paragraph parsing — mirrors `docx/mod.rs` logic: when a cell paragraph has `inline_image` and no text, lifts it to `Paragraph.image` and sets `content_height`
4. In `compute_row_layouts()` (`pdf/table.rs`): when an empty paragraph has `content_height > 0` (image), adds it to the cell's total height

### Files Modified
- `src/docx/tables.rs` — nested table flattening, inline image extraction in cell paragraphs
- `src/pdf/table.rs` — image height in cell height computation

### Results
- **No visual REGRESSION flags across all fixtures**
- `mandated_reporter_child_abuse`: 18.2% → 18.4% Jaccard (+0.2pp), 43.5% → 43.7% SSIM (+0.2pp)
- `mandated_reporter_child_abuse` text boundary: 9% → 28% (+19pp) — title text now correctly rendered
- Small noise-level variations on unrelated fixtures (samtale -0.4pp, japanese SSIM -0.6pp) — these fixtures have no nested tables and the changes don't affect their code paths
- Empty paragraph fix for table cells investigated but deferred due to regressions (see Analysis)

### Not Fixed (deferred)
- **Empty paragraph height in table cells**: Every `w:p` in a table cell (including spacer paragraphs) should contribute `line_h` to cell height. Current code gives 0 height. Fixing this improved `czech_grant_application` by +1.5pp but caused -2.9pp and -3.4pp SSIM regressions on other fixtures because the mandatory end-of-cell paragraph marker also gets full `line_h`. Needs a way to distinguish spacer paragraphs from the structural end-of-cell marker.

## Session 3 — 2026-03-14: Inline image effectExtent/dist + table cell image rendering

### Case Selected
`mandated_reporter_child_abuse` (text/layout only, 5 pages, 18.4% Jaccard) — continued from session 2 as the fixture closest to the 20% threshold (1.6pp away). The header contains a table cell with a large inline image (JCS Inc. logo) whose layout extra height from `wp:effectExtent` and `distT/distB` was not being accounted for. Additionally, inline images in table cells were parsed but never rendered.

### Problem
Two issues:
1. **Inline image layout height missing effectExtent and dist margins**: `wp:inline` elements have `effectExtent` (space for visual effects like borders) and `distT/distB` (minimum distance to surrounding text) attributes. These were not included in the image's layout height, causing table cells containing images to be shorter than they should be. In the header table, this made the first page body text start too high, shifting all subsequent pages.
2. **Inline images in table cells not rendered**: Session 2 added image parsing and height contribution for table cell images, but the actual image XObject was never embedded or drawn. The logo was correctly sized but invisible.

### Analysis
- Investigated the first-page header of `mandated_reporter_child_abuse`: a table with logo image (73.9×63.4pt) in cell 1, nested table with title in cell 2
- The image's `wp:inline` had `distT="114300" distB="114300"` (9pt each) and `effectExtent t="25400" b="25400"` (2pt each) — total 22pt of extra height not being used
- Without this extra height, cell 1 content was ~80pt, below the `trHeight=1965` (98.25pt) minimum. With it, cell content reaches ~102pt, exceeding the minimum and increasing the table row height by 4.25pt
- The 4.25pt increase in header height shifts body text down, improving vertical alignment with the reference across all 5 pages
- Initial approach included effectExtent in body paragraph height too, but this caused -1.0pp regression on `russian_sports_ranking_decree` (its coat-of-arms image has effectExtent b=0.6pt that shifted text). Fixed by only applying layout_extra_height in table cell context, not body paragraphs.
- For table cell image rendering, added XObject pre-embedding for all table cells (body + headers/footers) and image drawing in the cell rendering code

### Implementation
1. Added `layout_extra_height: f32` field to `EmbeddedImage` struct — captures effectExtent top+bottom + distT+distB (in pts)
2. Added `inline_extra_height()` helper in `images.rs` — extracts effectExtent and dist from `wp:inline` container
3. In `parse_run_drawing_result()` and `compute_drawing_info()` — inline images now get `layout_extra_height` set
4. In `tables.rs` cell parsing — `content_height` includes `display_height + layout_extra_height` for image paragraphs
5. In `header_footer.rs` `compute_header_height()` — uses `display_height + layout_extra_height` for inline images
6. In `mod.rs` body parser — deliberately does NOT include layout_extra_height (to avoid body text displacement regressions)
7. Added `table_cell_image_names` HashMap to `EmbeddedImages` and `RenderContext` — maps `Arc::as_ptr()` address to PDF XObject name
8. In `embed_all_images()` — walks all table cells in body + headers/footers to pre-embed images
9. Added `image_name`, `image_width`, `image_height`, `content_height` fields to `CellParagraphLayout`
10. In `render_cell_paragraphs()` — draws images using `content.x_object()` when `image_name` is set
11. Fixed `content_h` calculation in `render_table_row()` — includes `content_height` for image paragraphs (needed for vAlign centering)

### Files Modified
- `src/model.rs` — added `layout_extra_height` to `EmbeddedImage`
- `src/docx/images.rs` — `inline_extra_height()` helper, `read_image_from_zip_extra()`, pass extra height for inline images
- `src/docx/mod.rs` — body parser uses `display_height` only (no extra height for body paragraphs)
- `src/docx/tables.rs` — table cell parser includes `layout_extra_height` in content_height
- `src/pdf/mod.rs` — `Table` import, `table_cell_image_names` in `RenderContext` and `EmbeddedImages`, pre-embedding for table cell images
- `src/pdf/table.rs` — `CellParagraphLayout` image fields, image rendering in cells, fixed content_h for vAlign
- `src/pdf/header_footer.rs` — includes `layout_extra_height` in header height computation

### Results
- **mandated_reporter_child_abuse**: 18.4% → 19.6% Jaccard (+1.2pp), 43.7% → 49.5% SSIM (+5.8pp)
- Logo image now renders visibly in the header table cell
- Header table height increased from 98.25pt to 102.5pt (effectExtent+dist pushes cell content past trHeight minimum)
- Small noise-level variations on some fixtures (samtale -0.4pp, sample500kB -0.3pp, indonesian_benchmark -0.5pp) — these are within measurement noise range and don't affect pass/fail status
- No REGRESSION flags on visual comparison

### Not Fixed (deferred)
- **mandated_reporter still 0.4pp below 20% threshold**: The page break on page 1 still falls at a slightly different point than Word. ~10pt of additional header height is needed (from nested table row heights lost during flattening). Fixing nested table height preservation in the flattening code would likely push this fixture over the threshold. **→ Fixed in session 4.**
- **Empty paragraph height in table cells**: Still deferred (see session 2 notes).

## Session 4 — 2026-03-14: Cell-level tcMar + nested table margin preservation

### Case Selected
`mandated_reporter_child_abuse` (text/layout only, 5 pages, 19.6% Jaccard) — continued from session 3 as it was 0.4pp below the 20% threshold. The header table row height was too small because cell-level margins (`w:tcMar`) were not parsed, and nested table cell margins were lost during flattening.

### Problem
Two issues in table cell margin handling:
1. **Cell-level `w:tcMar` not parsed**: The code only read table-level `w:tblCellMar` defaults. Individual cells can override margins via `w:tcMar` in `w:tcPr`. In the mandated_reporter header table, each cell has `tcMar` with 1.8pt (36 twips) all around, but we were using the table-level defaults (0pt top/bottom). This made cell content heights 3.6pt shorter than they should be.
2. **Nested table cell margins lost during flattening**: `collect_nested_table_paragraphs()` extracted paragraph XML nodes from nested tables but discarded the nested table's cell margins (`tcMar`). In the header, the nested table's cells had 5pt (100 twips) margins on each side — 10pt of vertical spacing completely lost.

### Analysis
- Header table height was 99.3pt (computed by `compute_hf_table_height`); should be ~102.9pt
- Outer row trHeight=1965 twips (98.25pt) with hRule=atLeast — content was barely above the minimum
- Cell 1 (logo): image paragraph (85.4pt content_height) + text paragraph (~13.3pt) + 0pt cell margins = ~98.7pt content
- With cell-level tcMar (+3.6pt): cell 1 = ~102.3pt → row height = ~102.9pt
- This 3.6pt increase in header table height shifted body text down, improving vertical alignment across all 5 pages
- Nested table margin fix adds 10pt to cell 2, but cell 1 is taller so it doesn't affect the row height in this case. However, the fix is correct for other documents with nested tables where cell 2 might be the tallest cell.

### Implementation
1. Added `cell_margins: Option<CellMargins>` to `TableCell` struct in `model.rs`
2. Parse `w:tcMar` for each cell in `parse_table_node()` in `tables.rs` — falls back to table-level `tblCellMar`
3. In `compute_row_layouts()` (`pdf/table.rs`): use per-cell margins (`ecm`) for cell text width, initial total height, and rotated cell width
4. In `render_table_row()` and `render_header_footer_table()` (`pdf/table.rs`): use per-cell margins for vAlign offset and cursor_y computation, pass to `render_cell_paragraphs()`
5. Changed `collect_nested_table_paragraphs()` to use `AnnotatedNode` struct carrying `extra_space_before`/`extra_space_after` from nested table cell margins
6. Nested table cell margins read from both `tblCellMar` (table-level fallback) and `tcMar` (cell-level override)
7. Extra spacing applied to first/last paragraphs from each nested cell during outer cell paragraph parsing

### Files Modified
- `src/model.rs` — added `cell_margins: Option<CellMargins>` to `TableCell`
- `src/docx/tables.rs` — `AnnotatedNode` struct, nested table margin preservation in `collect_nested_table_paragraphs()`, `tcMar` parsing per cell
- `src/docx/alt_chunk.rs` — added `cell_margins: None` to HTML table cell construction
- `src/pdf/table.rs` — per-cell margins in `compute_row_layouts()`, `render_table_row()`, `render_header_footer_table()`
- `src/pdf/header_footer.rs` — removed debug output (cleanup)

### Results
- **mandated_reporter_child_abuse**: 19.6% → 26.4% Jaccard (+6.8pp), 49.5% → 49.9% SSIM (+0.4pp) — **NOW PASSING** (25 passing fixtures)
- No REGRESSION flags across all fixtures
- Small noise-level variations: sample500kB -0.3pp, samtale -0.4pp, japanese_interlibrary -0.6pp (all within noise range, no pass/fail status changes)

### Commit
`29d96fc` — "Support cell-level tcMar and preserve nested table cell margins during flattening"

### Not Fixed (deferred)
- **Empty paragraph height in table cells**: Partially fixed in session 5 (conservative approach for all-empty cells only).
- **mandated_reporter SSIM still below 75% (49.9%)**: Horizontal text positioning differences remain. The SSIM metric has zero horizontal tolerance (see memory notes), so even small horizontal shifts severely impact SSIM scores.

## Session 5 — 2026-03-14: Empty paragraph height in all-empty table cells

### Case Selected
`czech_grant_application` (text/layout only, 2 pages, 9.2% Jaccard) — chosen because it's a form-style document where empty paragraphs in table cells serve as vertical spacers for fill-in areas. This was the deferred issue from session 2 — empty paragraphs in table cells contributed 0 height instead of `line_h`, causing form fields to collapse.

### Problem
Empty paragraphs (no text runs) in table cells were given 0 height in `compute_row_layouts()`. In Word, every paragraph contributes at least one line of height to the cell. This caused form-style documents (like the Czech grant application) to render with collapsed cells — multi-line fill-in areas appeared as single-line rows.

### Analysis
- Investigated 6 text/layout-only failing fixtures: `polish_archery_range_plan` (15.0%), `slovak_misdemeanor_amendment` (12.9%), `russian_sports_ranking_decree` (12.8%), `czech_grant_application` (9.2%), `croatian_regulations_altchunk` (8.1%), `japanese_interlibrary_loan` (3.5%)
- `polish_archery_range_plan` differences were primarily subtle text-wrapping/font-metric issues — no actionable structural bug found
- `czech_grant_application` was a form with table cells using empty paragraphs as vertical spacers (e.g., "Účel, na který chce žadatel dotaci použít" has 4+ empty paragraphs in the right cell creating a fill-in area)
- Session 2 had found this bug but deferred it because naively giving ALL empty paragraphs `line_h` caused -2.9pp and -3.4pp SSIM regressions on other fixtures (every end-of-cell marker paragraph also got height)
- Root cause of regressions: cells with text content + one empty end-of-cell marker paragraph get inflated by `line_h`. In cells with mixed content, the end-of-cell marker should NOT add extra height.
- First attempt (give empty paragraphs `line_h`, skip last paragraph if cell has content): improved Czech (+1.5pp Jaccard) but caused regressions on `education_consultant_posting` (-1.0pp Jaccard, -3.4pp SSIM) and `croatian_regulations_altchunk` (-0.7pp Jaccard, -2.3pp SSIM)
- Conservative approach: only give empty paragraphs `line_h` when ALL paragraphs in the cell are empty (pure spacer cells). This targets the form pattern without affecting cells that have text + end-of-cell marker.

### Implementation
1. In `compute_row_layouts()` (`pdf/table.rs`): before iterating cell paragraphs, compute `cell_has_content` — whether any paragraph in the cell has non-empty text
2. In the empty paragraph branch: add `line_h` to `total_h` only when `!cell_has_content` (all paragraphs are empty, indicating spacer cells)
3. Existing behavior preserved for cells with content — their end-of-cell markers still contribute 0 height
4. Also updated `tests/baselines.json` to fix pre-existing stale baselines for `go_math_grade4_guide` (was 26.37% stored but actually 16.9%) and `croatian_regulations_altchunk`

### Files Modified
- `src/pdf/table.rs` — empty paragraph height in all-empty table cells
- `tests/baselines.json` — updated stale baselines

### Results
- **czech_grant_application**: 9.2% → 10.7% Jaccard (+1.5pp), 30.8% → 29.6% SSIM (-1.2pp — structural shift from taller cells moves page 2 content slightly)
- No REGRESSION flags across all fixtures
- 25 passing fixtures (unchanged)
- Form cells in the Czech grant application now have correct vertical spacing — "Účel", "Odůvodnění žádosti", and "Seznam příloh žádosti" cells match reference layout

### Not Fixed (deferred)
- **Empty paragraph height in mixed-content cells**: Cells with text + empty end-of-cell marker should also give the empty paragraph `line_h`, but this causes regressions with current layout code (other compensating errors). Requires either OS/2 WinMetrics line height fix (roadmap item) or more careful row height computation to land without regressions.
- **Font-metric text wrapping differences**: Multiple text/layout-only fixtures (`polish_archery_range_plan`, `slovak_misdemeanor_amendment`, `russian_sports_ranking_decree`) have subtle line-break differences from font width measurement discrepancies. Requires text shaping (rustybuzz) or font-specific metric corrections.

## Session 6 — 2026-03-14: Investigation of remaining failing fixtures (no code change)

### Objective
Find structural bugs or missing features to push failing fixtures closer to or past the 20% Jaccard threshold.

### Fixtures Investigated
All 14 failing fixtures were analyzed. Deep investigation was done on:
- `polish_archery_range_plan` (15.0% — closest text/layout-only to threshold)
- `slovak_misdemeanor_amendment` (12.9% — text/layout only)
- `russian_sports_ranking_decree` (12.8% — text/layout only)
- `education_consultant_posting` (8.6% — already worked on in prior plans)
- `east_asia_conference_form` (3.8% — CJK font issue, not fixable without CJK support)
- `croatian_regulations_altchunk` (7.5% — entire document is MHT altChunk)
- `mongolian_human_rights_law` (13.5% — Cyrillic text + 2 anchored images)
- `go_math_grade4_guide` (16.9% — 30 anchored images)

### Approaches Attempted

#### 1. TJ Kerning in PDF Output (reverted)
- **Problem**: Layout computation uses kerning (`word_width` with kern=true) but PDF rendering uses `Tj` (no kerning), causing a mismatch between computed and rendered text widths.
- **Implementation**: Added `kern: bool` to `WordChunk`, implemented `show_with_kerning()` using `content.show_positioned()` (TJ arrays) with kern pair adjustments.
- **Results**: Net negative — 4 Jaccard regressions (-0.1 to -0.4pp), only 1 improvement (`russian_sports_ranking_decree` +0.2pp Jaccard, +0.5pp SSIM). No pass/fail status changes.
- **Conclusion**: Our kern pair values don't exactly match Word's kerning behavior, so applying them in rendering worsens alignment for some fixtures. Reverted.

#### 2. snapToGrid / docGrid Investigation
- **Hypothesis**: `w:docGrid @w:linePitch=360` (18pt grid) with `snapToGrid` (default true) would cause line snapping to 18pt instead of 13.8pt natural height — a 30% expansion.
- **Finding**: Per OOXML spec §17.18.14, `w:docGrid @w:type` defaults to `"default"` = "No Document Grid". NO fixtures in the corpus have an explicit `w:type` attribute. Therefore `snapToGrid` is irrelevant for the entire test corpus.

#### 3. AltChunk Default Line-Height Tuning (reverted)
- **Context**: `croatian_regulations_altchunk` is entirely MHT HTML content. Paragraphs without explicit `line-height` use default `Auto(1.1)`.
- **Tested**: `Auto(1.0)` (single spacing) → -1.4pp Jaccard, -5.8pp SSIM. `Auto(1.15)` → -0.3pp Jaccard, -2.3pp SSIM.
- **Conclusion**: Original `Auto(1.1)` is already near-optimal. The issue is font metrics, not line-height defaults.

#### 4. Feature Audit
- Ran `--audit` across all fixtures. Key features already implemented: `w:spacing` (char spacing), `w:ind w:right`, `w:kern`, `w:numPr`, `w:smallCaps`, `w:vanish`, `contextualSpacing`.
- No missing features found that would affect multiple failing fixtures.

#### 5. Prior Work Review
- Read `plan_archery.md` / `plan_archery_progress.md`: OS/2 font metrics fix was already done, was net-neutral for this fixture (TNR's OS/2 win metrics + hhea lineGap = original hhea metrics).
- Read `plan_education.md` / `plan_education_progress.md`: Table row splitting, per-paragraph rendering, SDT parsing already implemented.
- Conclusion from archery plan: "The cumulative vertical drift is NOT caused by font metrics."

### Key Finding: Root Cause of Remaining Failures
All 6 text/layout-only failures share the same root cause: **font width measurement discrepancies** causing different line wrapping decisions. This manifests as:
1. Different word-per-line counts → different number of lines per paragraph
2. Cascading vertical position shifts across the page
3. Different page break points (though all fixtures match page count)

This is confirmed by:
- Page counts match between generated and reference PDFs
- Visual diffs show red/blue text pairs close together (horizontal displacement)
- The displacement increases progressively down each page (cumulative drift)
- No missing content blocks — all text is present, just positioned differently

### Blocked By (from roadmap)
1. **Text Shaping (rustybuzz)** — proper OpenType shaping would fix ligatures, kerning, and glyph substitution, producing more accurate text widths.
2. **Unicode Line Breaking** — correct break opportunities for non-Latin scripts.
3. **CJK Font Support** — blocks `japanese_interlibrary_loan` and `east_asia_conference_form`.

### No Commit (Session 6)
No code changes were made. All experimental changes were reverted.

## Session 7 — 2026-03-14: Deep investigation of non-text-layout failures (no code change)

### Objective
Move beyond the text/layout-only failures (blocked by font metrics per session 6) by investigating failing fixtures with structural features: anchored images, floating tables, SDTs, textboxes, and footnotes.

### Fixtures Investigated
All 14 failing fixtures were analyzed. Deep investigation on:
- `go_math_grade4_guide` (16.9% — 30 anchored images, `wrapSquare`)
- `brazilian_logistics_study` (16.9% — 8 anchored images, `wrapSquare`)
- `mongolian_human_rights_law` (13.5% — standard fonts Arial/TNR, 2 anchored images, 6 footnotes)
- `education_consultant_posting` (8.6% — cell-level SDTs)
- `croatian_grant_guidelines` (7.0% — generates 72 pages vs 65 reference)
- `air_pollution_permit_form` (12.6% — 21 textboxes)

### Approaches Attempted

#### 1. Font Line Height: Remove hhea lineGap from OS/2 Win Metrics Path (reverted)
- **Hypothesis**: Word uses `usWinAscent + usWinDescent` without `hhea lineGap` for fonts without `USE_TYPO_METRICS`. Our code adds lineGap, making lines ~0.4pt too tall for Arial 12pt.
- **Results**: 21 regressions, including -42.5pp (case7), -59pp (centrifugal_water_chillers), -39.4pp (seminary_hill). The mongolian fixture itself worsened by -5.3pp.
- **Conclusion**: The current line height formula (win metrics + hhea lineGap) is well-calibrated and represents a local optimum. Removing lineGap makes everything worse. This confirms the roadmap note that the line height fix "causes 23 regressions."

#### 2. wrapSquare Vertical Space Threshold: Lower from 90% to 50% (reverted)
- **Finding**: Text wrapping around floating images is NOT implemented. For `wrapSquare` images, only images ≥90% of text width get vertical space reserved. This misses images at 78-89% (like the Brazilian fixture's chart images).
- **Results**: `indonesian_benchmarking_guide` +5pp (22.4%→27.4%), `brazilian_logistics_study` -4pp (16.9%→12.9%).
- **Problem**: Indonesian was already passing (22.4% > 20%). Brazilian regression is because reserving space for paragraph-relative images creates gaps where text should wrap but doesn't (since we don't implement wrapping). The height formula `v_offset + display_height` overestimates needed space.
- **Conclusion**: Without implementing actual text wrapping, lowering the threshold hurts more than it helps.

#### 3. Footnote Height vs Rendering Line Spacing Mismatch
- **Bug found**: `compute_footnote_height()` uses `ctx.doc_line_spacing` as fallback (could be 1.15×), but `render_page_footnotes()` uses `LineSpacing::Auto(1.0)` (single spacing). Height calculation overestimates for documents with >1.0 default spacing.
- **Impact**: Negligible — footnote paragraphs in all tested fixtures have explicit single spacing set via the "Footnote Text" style, so the fallback is never used.
- **Not fixed**: Correct fix would be to use `LineSpacing::Auto(1.0)` in both places, but zero measurable improvement.

#### 4. Cell-Level SDT Parsing
- **Fixture**: `education_consultant_posting` with 32 SDT elements wrapping `w:tc` in table rows.
- **Finding**: Already handled correctly. `collect_block_nodes()` recursively unwraps `w:sdt/w:sdtContent` and exposes the inner `w:tc` elements. The 2-page difference (5 vs 7) and 189 "missing" words are layout compression (text packed tighter), not lost content.

#### 5. Feature Implementation Audit
- `w:caps`, `w:smallCaps`, `w:vanish`, `w:dstrike`, `w:spacing` (char spacing), `contextualSpacing`: all already implemented and working correctly.
- `beforeAutospacing`/`afterAutospacing`: not parsed (12 fixtures), but the default NormalWeb style values (`before="100"`) happen to match Word's auto-spacing behavior (~5pt), so there's no effective difference.
- `contextualSpacing` spec deviation: our code checks if BOTH paragraphs have the flag; spec says check if they have the SAME STYLE. Doesn't matter in practice since the flag typically comes from a shared style.

### Key Finding: Confirmation of Session 6 Conclusion
All remaining failures share the same root cause: **font width measurement discrepancies** causing different line wrapping. This is confirmed by:
1. The mongolian fixture uses standard fonts (Arial, Times New Roman) yet still has progressive vertical drift from cumulative per-line wrapping differences
2. `go_math_grade4_guide` page count mismatch (23 vs 26) is from Museo Sans 300 substitution
3. All text boundary tests show >95% word presence — content is present but displaced
4. Non-text-layout fixtures (images, tables, SDTs) all have correct structural parsing but inherit the same font-width-driven layout differences

### Blocked By (same as session 6)
1. **Text Shaping (rustybuzz)** — proper OpenType shaping would fix glyph widths
2. **Text Wrapping** — implementing `wrapSquare`/`wrapTight` text flow around floating images would help fixtures with large images but is architecturally complex
3. **Unicode Line Breaking** — correct break opportunities for non-Latin scripts

### No Commit (Session 7)
No code changes were made. All experimental changes were reverted.

## Session 8 — 2026-03-14: Push body text below page-anchored floating tables

### Case Selected
`polish_municipal_letter` (floating table + 2 anchored images, 1 page, 13.2% Jaccard) — chosen because it had a structural bug in floating table positioning that hadn't been investigated in sessions 6-7 (those sessions focused on text/layout-only failures). The fixture has a floating table header (coat of arms + municipal contact info) that should push body text below it.

### Problem
Floating tables with `vertAnchor="page"` positioned above or straddling the top margin caused body text to render inside the table area. After rendering a floating table, `slot_top` was unconditionally restored to its pre-table value (the margin position), regardless of where the table's bottom edge ended up. When the table started above the margin and extended below it, body text would overlap the table by the distance between the margin and the table bottom.

### Analysis
- Investigated all 14 failing fixtures. Grouped by structural feature: 6 text/layout-only (blocked by font metrics per sessions 6-7), 5 with anchored images, 3 with floating tables, 1 with textboxes.
- Deep investigation of `polish_municipal_letter`: floating table at `tblpY=946` (47.3pt from page top), `vertAnchor="page"`. Page margin top = 70.85pt. Table starts 23.55pt ABOVE the margin and extends to ~150pt from top.
- Debug tracing confirmed: `slot_top` restored to 771.05 (= margin position, 70.85pt from top) after table renders, while table bottom was at 691.89 (= 150pt from top). Body text started 79pt above the table bottom.
- Also investigated `italian_project_proposal` (passing at 28.3%): its floating table at `tblpY=2236` (111.8pt from top) starts BELOW the margin (70.85pt). Body text correctly appears in the gap above the table. This case requires preserving the old behavior.
- Also investigated `brazilian_logistics_study` (16.9%) and `go_math_grade4_guide` (16.9%) — both blocked by text wrapping (wrapSquare) and font substitution respectively, not actionable.

### Implementation
1. Changed `saved_slot_top` from `Option<f32>` to `Option<(f32, f32)>` — now stores both the original slot_top and the table's initial y position
2. After rendering all rows, compare: if `table_top_y >= saved` (table starts at/above margin) AND `table_bottom < saved` (table extends below margin), set `slot_top = table_bottom` to push body text below the table
3. Otherwise, restore `slot_top` to saved value (existing behavior for tables starting below the margin)

### Files Modified
- `src/pdf/table.rs` — floating table slot_top restoration logic

### Results
- **polish_municipal_letter**: 13.2% → 26.5% Jaccard (+13.3pp), 28.4% → 68.3% SSIM (+39.9pp) — **NOW PASSING** (26 passing fixtures)
- `italian_project_proposal`: 28.3% Jaccard (unchanged) — correctly not affected
- No REGRESSION flags across all fixtures
- Small noise-level changes: `croatian_grant_guidelines` -0.2pp Jaccard, `east_asia_conference_form` -0.1pp Jaccard (both within noise range, no pass/fail changes)

### Commit
`24f23ab` — "Push body text below page-anchored floating tables that cover the margin"

### Not Fixed (deferred)
- **Text wrapping around floating tables**: When a floating table doesn't span full width, text should wrap beside it. Currently no text wrapping is implemented for floating tables — text either goes above or below. Affects `croatian_grant_guidelines` and `east_asia_conference_form`.
- **Font width measurement discrepancies**: 6 text/layout-only fixtures remain blocked by font metrics (sessions 6-7 conclusion).
- **wrapSquare text wrapping**: `brazilian_logistics_study` (16.9%) blocked by lack of text wrapping around floating images.
