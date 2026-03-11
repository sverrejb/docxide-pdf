# Cleanup Progress

## Completed

### Tier 1, Task 1: Split `render()` phases 1/4/7 into standalone functions
- Extracted `collect_and_register_fonts()` — font collection & registration (~240 lines)
- Extracted `embed_all_images()` — image XObject creation (~170 lines) with `EmbeddedImages` struct
- Extracted `assemble_pdf_pages()` — gradient defs, page objects, resource dicts (~220 lines)
- Moved `para_runs_with_textboxes` and `collect_paras` to module-level functions
- Converted `embed_image` closure to standalone `embed_single_image` function
- `render()` reduced from ~1870 to ~1268 lines
- Merged `t_collect` into `t_fonts` timing (simplified log message)
- Verified: builds clean, no test regressions (case1: 59.1% Jaccard, 85.8% SSIM)

### Tier 1, Task 2: Introduce `PageBuilder` struct to encapsulate mutable render state and unify page-flush logic
- Created `PageBuilder` struct in `src/pdf/mod.rs` encapsulating 18 fields:
  - Current page state: `content`, `links`, `footnote_ids`, `alpha_states`, `gradient_specs`
  - Cross-page state: `styleref_running`, `styleref_page_first`
  - Layout position: `slot_top`, `is_first_page_of_section`
  - 9 accumulated-pages vectors (`all_contents`, `all_links`, etc.)
- Added `flush_page(sect_idx)` method — single source of truth for page breaks (replaces 8 lines × 6 sites)
- Added `push_blank_page(sect_idx)` for odd/even page alignment
- Replaced 6 duplicated flush sites in `render()` + 1 in table.rs with method calls
- Updated `render_table` signature from 15 parameters to 8 (`&mut PageBuilder` replaces 7 individual refs)
- Removed 4 padding loops that were compensating for table.rs's partial flush
- Fixed `LinkAnnotation` visibility (`pub(super)` → `pub(crate)`) to match PageBuilder exposure
- Verified: builds clean, zero test regressions

### Tier 2, Task 3: Extract shared docx parsing helpers
- Added `resolve_theme_color_key()` in `docx/mod.rs` — maps OOXML scheme names (dk1/lt1/tx1/bg1 etc.) to theme element names. Replaced 4 identical 11-line match blocks in `textbox.rs`.
- Added `parse_paragraph_spacing()` in `docx/mod.rs` — extracts space_before/after/line_spacing from ppr node with style fallback. Returns `(Option<f32>, Option<f32>, Option<LineSpacing>)` so callers apply their own defaults. Replaced duplicated extraction in `mod.rs`, `textbox.rs`, `tables.rs`, `headers_footers.rs` (5 call sites total).
- Added `extract_indents()` in `docx/mod.rs` — extracts left/right/hanging/firstLine from `w:ind` node with bidi fallback (start→left, end→right). Replaced duplicated extraction in `mod.rs`, `textbox.rs`, `tables.rs`, `styles.rs` (4 call sites).
- **Bug fix**: `textbox.rs` indent extraction now uses bidi fallback (`w:start`/`w:end`) via the shared helper, matching the behavior in all other modules.
- Removed unused `parse_line_spacing` import from `textbox.rs`, `tables.rs`, `headers_footers.rs`.
- Removed unused `twips_attr` import from `textbox.rs`, `headers_footers.rs`.
- Verified: builds clean, all tests pass, zero regressions.

### Tier 2, Task 4: Introduce `RenderContext` struct bundling `seen_fonts` + related read-only state
- Created `RenderContext<'a>` struct in `src/pdf/mod.rs` with two fields:
  - `fonts: &'a HashMap<String, FontEntry>` — immutable font registry
  - `doc_line_spacing: LineSpacing` — document-level default line spacing
- Updated 11 function signatures across 3 submodules to take `&RenderContext` instead of separate `(seen_fonts, doc_line_spacing)` parameters:
  - `header_footer.rs`: `compute_header_height`, `effective_slot_top`, `compute_effective_margin_bottom`, `render_header_footer` (13→12 params)
  - `table.rs`: `render_table` (8→7 params), `compute_row_layouts` (5→4 params), `compute_hf_table_height` (3→2 params), `render_header_footer_table` (9→8 params)
  - `footnotes.rs`: `compute_footnote_height` (4→3 params), `render_page_footnotes` (9→8 params)
  - `mod.rs`: `render_single_textbox` (11→10 params)
- All internal `seen_fonts` accesses changed to `ctx.fonts`, `doc_line_spacing` to `ctx.doc_line_spacing`
- Functions taking only `seen_fonts` (layout.rs, charts.rs, smartart.rs) receive `ctx.fonts` at call sites — no signature changes needed
- Removed unused `FontEntry` import from `header_footer.rs`, unused `LineSpacing` import from `table.rs`
- Verified: builds clean, all tests pass, zero regressions (case1: 59.1% Jaccard, 85.8% SSIM)

### Tier 2, Task 5: Unify `build_paragraph_lines` / `build_tabbed_line` shared core + `WordChunk` constructors
- Added `WordChunk::text()` constructor (8 params) — replaces 2 copy-pasted 17-field struct literals for text word chunks in `build_paragraph_lines` and `build_tabbed_line`
- Added `WordChunk::image()` constructor (5 params) — replaces 2 copy-pasted 17-field struct literals for inline image chunks in both functions
- Replaced 3 manual `TextLine { chunks, total_width }` construction sites in `build_tabbed_line` with existing `finish_line()` helper
- Leader chunk (1 site) left as manual construction — unique enough that a dedicated constructor would be over-engineering
- Did NOT merge the two layout functions into one: they have fundamentally different space-tracking approaches (`split_preserving_spaces` + `pending_space_w` vs `split_whitespace` + `prev_ws` + tab-stop resolution). Merging would hurt readability without meaningful benefit.
- Net reduction: ~65 lines removed from layout.rs
- Verified: builds clean, all tests pass, zero regressions

### Tier 3, Task 6: Fix `Error::Pdf("Missing w:body")` → `Error::InvalidDocx`
- Changed `Error::Pdf("Missing w:body".into())` to `Error::InvalidDocx("Missing w:body".into())` in `docx/mod.rs:412`
- This is a parse-time error (missing XML element in DOCX), not a PDF rendering error — `InvalidDocx` is the correct variant
- Verified: builds clean, no new warnings

### Tier 3, Task 7: Convert `h_relative_from`/`v_relative_from` to enums
- Added `HRelativeFrom` enum (Page, Margin, Column) and `VRelativeFrom` enum (Page, Margin, TopMargin, Paragraph) to `model.rs`
- Changed `FloatingImage.h_relative_from` and `FloatingImage.v_relative_from` from `&'static str` to the new enums
- Changed `Textbox.h_relative_from` and `Textbox.v_relative_from` from `&'static str` to the new enums
- Updated `parse_anchor_position()` in `docx/images.rs` to return enum values instead of string literals
- Updated VML textbox parsing in `docx/textbox.rs` to use enum values
- Updated all match sites in `pdf/mod.rs` (`resolve_fi_x`, `resolve_fi_y_top`, `render_single_textbox`) and `pdf/header_footer.rs` (textbox + floating image positioning)
- Benefit: exhaustive match catches typos at compile time; no more wildcard `_ =>` fallbacks hiding unhandled variants
- Verified: builds clean, all tests pass, zero regressions (case1: 59.1% Jaccard, 85.8% SSIM)

### Tier 3, Task 8: Fix `analyze_fixtures.rs` stale skip list parsing
- Rewrote `load_skip_list()` in `tools/src/bin/analyze_fixtures.rs` to read `tests/fixtures/SKIPLIST` file directly
- Old implementation parsed `tests/common/mod.rs` source code looking for a deleted `SKIP_FIXTURES` constant — always returned empty list
- New implementation matches the format used by `tests/common/mod.rs:load_skiplist()`: reads SKIPLIST file, strips comments and blank lines
- `baselines_key()` already matches `common::display_name()` logic — no change needed
- `load_baselines()` uses its own JSON parsing but produces correct results — acceptable for a standalone tool binary
- Verified: builds clean

### Tier 3, Task 9: Clean up `#[allow(dead_code)]` on chart types
- Removed blanket `#[allow(dead_code)]` from `ChartType` enum and `ChartAxis` struct in `model.rs`
- Added field-level `#[allow(dead_code)]` on specific unused fields:
  - `ChartType::Bar { stacked }` — parsed but not yet used in rendering (stacked bar not implemented)
  - `ChartAxis.delete` — parsed but not yet used in rendering
- Also narrowed existing/new dead code annotations in `model.rs`:
  - `Textbox.margin_bottom` — restored field-level allow (was blanket before)
  - `Textbox.dist_top` — added field-level allow (previously unnoticed dead field)
  - `SmartArtDiagram.display_width` / `display_height` — added field-level allows (set during parsing but rendering uses individual shape coordinates)
- Net effect: reduced model.rs warnings from 2 to 0; total crate warnings from 8 to 6
- Verified: builds clean, remaining 6 warnings are pre-existing in geometry/ and settings.rs (outside scope)

## All Tasks Complete
