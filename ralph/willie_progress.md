# Willie Progress — Code Simplification

## Processed Files

### 1. `src/docx/alt_chunk.rs`
- Consolidated imports: added `CellVAlign`, `LineSpacing`, `TextDirection`, `VMerge` to import block, replacing 6 verbose `crate::model::` qualified paths throughout the file
- Removed 5 section-divider comments (`// --- MHT / MIME envelope ---`, etc.) that described "what" not "why"
- Removed 1 redundant "what" comment on the default line-height else branch
- All tests pass, no regressions

### 2. `src/model/mod.rs`
- Removed re-export "what" comment
- Removed redundant unit comments on `ParagraphBorder` fields (names already encode units via `_pt` suffix)
- Removed `// points` on `ColumnDef` fields
- Condensed `num_level_tab_stop` doc comment to single line
- Tightened `HorizontalRule.is_standard` doc comment
- Kept `Run.color` comment (restored: semantic "why" — explains `None` = DOCX "automatic")
- Removed `Run.text_scale` percentage comment (default impl shows 100.0)
- All tests pass, no regressions

### 3. `src/model/drawing.rs`
- Fixed inconsistent import: added `VerticalPosition` to `use super::` block, replacing inline `super::VerticalPosition`
- Removed struct-level doc on `ImageShadow` (restated type name)
- Removed redundant field comments on `EmbeddedImage` (`// points` on display_width/height, stroke field XML-path comments)
- Tightened `layout_extra_height` and `layout_extra_top` doc comments
- Tightened `wrap_polygon` doc comment
- Restored `ShapeGeometry` doc comment (scope info: 187 presets + custom geometry)
- Restored `ImageShadow.offset_y` screen-coords note (coordinate system matters for rendering)
- All tests pass, no regressions

### 4. `src/model/table.rs`
- Removed `// points` on `Table.col_widths` and `TableCell.width`
- Condensed `all_paragraphs` doc comment — kept non-obvious "includes nested tables" note
- All tests pass, no regressions

### 5. `src/model/chart.rs`
- Skipped — already clean, nothing to simplify

### 6. `src/docx/mod.rs`
- Consolidated imports: added `BorderStyle`, `CellBorder`, `DocGridType`, `ParagraphBorder` to import block, replacing 18 verbose `crate::model::` qualified paths throughout the file
- Functions affected: `parse_one_border`, `parse_cell_border`, `parse_cell_border_with_fallback`, `parse_cell_border_left`, `parse_cell_border_right`, `parse_zip` (fallback SectionProperties)
- No comments changed (existing comments are all "why" comments — spec refs, workaround explanations)
- All tests pass, no regressions

### 7. `src/docx/styles.rs`
- Consolidated imports: added `Read`, `Seek` (std::io), `ParagraphBorders` (crate::model), and `extract_indents` (super) to import block
- Replaced 2 verbose `std::io::Read + std::io::Seek` trait bounds with short `Read + Seek`
- Replaced 1 `crate::model::ParagraphBorders` qualified path with `ParagraphBorders`
- Replaced 2 `super::extract_indents` calls with `extract_indents`
- Replaced 2 `super::parse_hex_color` calls with `parse_hex_color` (already imported)
- Removed redundant `.as_str()` on `&String` auto-deref
- Removed 2 unnecessary closure type annotations (`|s: &str|`, `|n: &&String|`)
- Removed explicit `HashMap<String, String>` type annotation (inferred from usage)
- All tests pass, no regressions

### 8. `src/docx/runs.rs`
- Consolidated imports: added `Seek` (std::io), `parse_text_fill`, `parse_text_glow`, `parse_text_outline`, `parse_text_shadow` (super::wordart) to import block
- Replaced verbose `std::io::Seek` trait bound with short `Seek` in `parse_runs` signature
- Replaced 4 `wordart::parse_text_*` qualified calls with direct `parse_text_*` calls
- Removed `wordart::self` import (no longer needed after direct function imports)
- Optimized `mc_choice_or_fallback()`: single-pass iteration with early return on Choice (was iterating children twice)
- Removed unnecessary closure type annotation (`|prev: &Run|` → `|prev|`) in `merge_compatible_runs`
- Removed 2 redundant "what" comments in `merge_compatible_runs` (`// Both must be plain text runs`, `// All visual properties must match`)
- All tests pass, no regressions

### 9. `src/docx/numbering.rs`
- Consolidated imports: added `Read`, `Seek` (std::io) to import block, replacing inline `std::io::Read + std::io::Seek` in `parse_numbering` signature
- Deduplicated `rFonts` font extraction: `bullet_font` and `label_font` used identical 6-line lookup; extracted into single `rpr_font` variable, cloned for both fields
- Deduplicated `wml(lvl, "pPr")` lookup: was called twice (once for `ind`, once for `tabs`); extracted into single `ppr` variable
- Removed 5 unnecessary explicit `HashMap` type annotations on local variables (types inferred from usage and return type)
- Removed 3 redundant "what" doc comments on struct fields (`suff` on LevelDef and ListLabelInfo, `label_font` on LevelDef)
- All tests pass, no regressions

### 10. `src/docx/images.rs`
- Consolidated imports: added `Seek` (std::io), `Textbox`, `ConnectorShape` to import block
- Replaced 4 verbose `std::io::Seek` trait bounds with short `Seek`
- Replaced 3 `crate::model::Textbox` qualified paths with `Textbox`
- Replaced 1 `crate::model::ConnectorShape` qualified path with `ConnectorShape`
- Replaced `crate::model::VRelativeFrom::Paragraph` and `crate::model::WrapType::TopAndBottom` with short forms (already imported)
- Extracted `PIC_NS` module-level constant, replacing 2 inline `pic_ns` locals in `parse_pic_outline` and `parse_pic_shadow`
- Extracted `find_pic_sp_pr` helper to deduplicate identical pic:pic → pic:spPr lookup shared by `parse_pic_outline` and `parse_pic_shadow`
- Refactored both functions to accept `Option<Node>` directly; updated 3 call sites to call `find_pic_sp_pr` once and pass result to both
- All tests pass, no regressions

### 11. `src/docx/charts.rs`
- Consolidated imports: added `Seek` to `std::io` import block, replacing inline `std::io::Seek` in `parse_chart_from_zip` signature
- Extracted `chart_children` helper to deduplicate 3 identical inline filter patterns (filtering children by name + CHART_NS namespace) in `collect_indexed_pts`, `collect_series`, and `parse_chart_space`
- Replaced inline `.children().find(|n| ...)` in `extract_plot_border` with existing `chart_child` helper
- Removed redundant "what" comment (`// noFill means no line`) in `extract_line_color`
- Removed unnecessary local type annotation on `val: f32` in `extract_fill_alpha` (moved to turbofish on `.parse::<f32>()`)
- All tests pass, no regressions

### 12. `src/docx/textbox.rs`
- Consolidated imports: added `FormulaOp`, `PathFill` (crate::geometry) and `AutoFit`, `TextWarp`, `VerticalPosition` (crate::model) to top-level import block
- Replaced 5 verbose `crate::model::*` qualified paths with short forms (`AutoFit::None`, `TextWarp`, `VerticalPosition::Offset`)
- Removed 2 local `use crate::geometry::*` statements inside `parse_custom_geometry` (now handled by top-level imports)
- Removed unused `LineSpacing` import
- Removed 2 redundant "what" comments (DrawingML path and VML fallback path descriptions that restate code structure)
- Kept "VML WordArt uses v:textpath instead of v:textbox" comment (explains *why* the else branch tries textpath)
- Replaced `textbox_node.is_none()` check + `.unwrap()` anti-pattern in `parse_textbox_from_vml` with `let-else` binding
- All tests pass, no regressions

### 13. `src/docx/tables.rs`
- Consolidated imports: added `Seek` to `std::io` import block, replacing inline `std::io::Seek` in `parse_table_node` signature
- Extracted `parse_table_borders_def` helper to deduplicate identical `TableBordersDef` construction used for inline table borders and row-level border exceptions
- Combined 3 separate iterations over `runs` (2 mutation loops for conditional bold/color + 2 `.any()` checks) into a single pass
- Removed 2 unnecessary `f32` type annotations on `indent_first_line` and `indent_right` (inferred from usage)
- All tests pass, no regressions

### 14. `src/docx/smartart.rs`
- Consolidated imports: added `Seek` to `std::io` import block, replacing inline `std::io::Seek` in `parse_smartart_shapes` signature
- Extracted `is_ns` helper to deduplicate 5 repeated tag-name + namespace check patterns across `has_dml`, `parse_smartart_drawing`, and `parse_dsp_text`
- Removed unnecessary `_f32` suffix on `0.0_f32` literal (type inferred from context)
- Removed unnecessary `Option<[u8; 3]>` type annotation on `text_color` in `parse_dsp_text` (inferred from return type)
- All tests pass, no regressions

### 15. `src/docx/wordart.rs`
- Consolidated imports: merged two separate `crate::model::` import blocks into one
- Added `find_child` to `super::` imports, created local `find_w14` helper (mirrors `find_dml` pattern) wrapping `find_child(node, name, W14_NS)`
- Replaced 11 inline `n.tag_name().name() == "X" && n.tag_name().namespace() == Some(W14_NS)` closures with `find_w14()` calls across `parse_text_outline`, `parse_text_fill`, `parse_text_shadow`, `parse_text_glow`, `resolve_w14_color`, `find_w14_solid_fill_color`, `parse_w14_gradient`
- Replaced 1 inline DML namespace child-find in `resolve_w14_color` with `find_child(node, "srgbClr", DML_NS)`
- Simplified `find_w14_solid_fill_color`: replaced nested if-let chains (6 lines) with chained `.and_then()` / `.or_else()` (3 lines)
- Simplified `parse_text_fill`: replaced nested if-let for solidFill with `.and_then(resolve_w14_color)`
- Simplified `parse_text_shadow` alpha parsing: replaced 8-line nested `.find()` chain with 4-line `find_w14` chain
- Removed `// --- Private helpers ---` section-divider comment
- All tests pass, no regressions

### 16. `src/docx/sections.rs`
- Consolidated imports: added `Seek` to `std::io` import block, replacing inline `std::io::Seek` in `parse_section_properties` signature
- All tests pass, no regressions

### 17. `src/docx/headers_footers.rs`
- Consolidated imports: added `Seek` to `std::io` import block, replacing inline `std::io::Seek` in `parse_header_footer_xml` and `parse_footnotes` signatures
- Removed unnecessary `f32` type suffix on indent initialization literals (`0.0f32` → `0.0`)
- Removed extra blank line between `is_wml_element` and `resolve_alignment`
- All tests pass, no regressions

### 18. `src/docx/settings.rs`
- Consolidated imports: added `Seek` to `std::io` import block, replacing inline `std::io::Seek` in `parse_settings` signature
- File already very clean (62 lines) — no other simplifications needed
- All tests pass, no regressions

### 19. `src/fonts/mod.rs`
- Consolidated imports: added `PathBuf` (std::path) and `Instant` (std::time) to import block
- Replaced 2 verbose `std::path::PathBuf` qualified paths with `PathBuf` (in `ResolvedFont` and `FontEntry` structs)
- Replaced `std::time::Instant::now()` with `Instant::now()`
- Condensed duplicate `FontMetrics` doc comment (two overlapping descriptions → one concise version)
- Removed unnecessary `f32` type annotation on `w` in `word_width`
- Deduplicated CJK fallback loop: extracted `try_cjk_fallback` closure accepting `&mut dyn FnMut`, replacing 2 identical 6-line loops
- Cached `lookup_font_table` result into `table_entry`, eliminating redundant second lookup
- All tests pass, no regressions

### 20. `src/fonts/discovery.rs`
- Consolidated imports: added `Instant` (std::time), `env`, `fs` (std) to import block
- Replaced 5 verbose `std::env::var` calls with `env::var`
- Replaced 2 `std::fs::read_dir` and 1 `std::fs::File` with `fs::read_dir` / `fs::File`
- Replaced `std::time::Instant::now()` with `Instant::now()`
- Combined two separate iterations over `families` (index insertion + CachedFace construction) into a single loop
- Removed redundant "what" comment ("Collect directory listing: subdirs to recurse, font files to process")
- Removed unnecessary `Vec<PathBuf>` type annotation on `dirs` (inferred from usage)
- Removed unnecessary `HashSet<PathBuf>` type annotation on `visited_dirs` (inferred from usage)
- All tests pass, no regressions

### 21. `src/fonts/embed.rs`
- Consolidated imports: merged two `ttf_parser::` imports and two `super::` imports into single blocks
- Removed unnecessary `Vec<f32>` type annotation on `widths_1000` (inferred from iterator chain)
- Optimized `extract_kern_pairs`: pre-collected filtered horizontal kern subtables before the O(n²) glyph pair loop (was re-filtering on every iteration)
- All tests pass, no regressions

### 22. `src/fonts/encoding.rs`
- Pre-allocated `to_winansi_bytes()`: replaced `chars().map().filter().collect()` with `Vec::with_capacity(s.len())` + explicit loop, avoiding default-capacity Vec resizing
- Pre-allocated `encode_as_gids()`: replaced `chars().flat_map().collect()` with `Vec::with_capacity(text.len() * 2)` + explicit loop — this function is called in hot text-rendering paths
- No comments changed (all existing comments are useful doc comments or data-table annotations)
- File already very clean (104 lines, well-structured match tables) — no import consolidation or deduplication needed
- All tests pass, no regressions

### 23. `src/fonts/cache.rs`
- Consolidated imports: added `Path` (std::path), `UNIX_EPOCH` (std::time), `env`, `fs` (std) to import block
- Replaced 4 verbose `std::env::var` calls with `env::var`
- Replaced 3 `std::fs::*` calls (`read_to_string`, `create_dir_all`, `write`) with `fs::*`
- Replaced `std::fs::metadata` with `fs::metadata`
- Replaced `std::path::Path` with `Path` in `dir_mtime` signature
- Replaced `std::time::UNIX_EPOCH` with `UNIX_EPOCH`
- File already very clean (134 lines, well-structured TSV cache) — no other simplifications needed
- All tests pass, no regressions

### 24. `src/geometry/mod.rs`
- Consolidated imports: added `CustomGeometry` (crate::model) to top-level import block
- Replaced 1 verbose `crate::model::CustomGeometry` qualified path with `CustomGeometry` in `evaluate_custom` signature
- File already very clean (97 lines of production code + 170 lines of well-structured tests) — no other simplifications needed
- All comments are "why" comments (coordinate system conventions, EMU scaling rationale) — none removed
- No deduplication opportunities: `evaluate_def` and `evaluate_custom` share structure but differ meaningfully (preset vs custom adjustments/guides, text_rect handling)
- All tests pass, no regressions

## Not Yet Processed
- src/error.rs (skipped — already minimal, nothing to simplify)
- src/main.rs (skipped — already minimal, nothing to simplify)
- src/lib.rs (skipped — already minimal, nothing to simplify)
- src/docx/embedded_fonts.rs (skipped — already clean, nothing to simplify)
- src/geometry/definitions.rs
- src/geometry/formulas.rs
- src/geometry/path.rs
- src/geometry/text_warp_definitions.rs
- src/pdf/mod.rs
- src/pdf/layout.rs
- src/pdf/table.rs
- src/pdf/table_layout.rs
- src/pdf/charts.rs
- src/pdf/charts_radial.rs
- src/pdf/chart_legend.rs
- src/pdf/header_footer.rs
- src/pdf/footnotes.rs
- src/pdf/smartart.rs
- src/pdf/wordart.rs
- src/pdf/images.rs
- src/pdf/positioning.rs
- src/pdf/color.rs
- src/pdf/helpers.rs
- src/pdf/fonts.rs
- src/pdf/list_label.rs
- src/pdf/assembly.rs
- src/pdf/textbox_render.rs
