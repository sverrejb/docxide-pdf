## Phase 1: Module size & complexity scan 

### Lines per file (top modules, excluding geometry/definitions.rs which is 9876 lines of lookup data)
| File | Lines |
|------|-------|
| src/pdf/mod.rs | 2365 |
| src/pdf/charts.rs | 1031 |
| src/docx/textbox.rs | 983 |
| src/pdf/layout.rs | 958 |
| src/docx/styles.rs | 739 |
| src/docx/alt_chunk.rs | 719 |
| src/docx/mod.rs | 667 |
| src/docx/runs.rs | 657 |
| src/model.rs | 620 |
| src/pdf/table.rs | 572 |
| src/geometry/formulas.rs | 519 |
| src/pdf/header_footer.rs | 494 |
| src/geometry/path.rs | 478 |
| src/docx/images.rs | 378 |
| src/docx/charts.rs | 352 |
| src/fonts/mod.rs | 323 |

### API surface (notable)
| File | pub | pub(super) | pub(crate) |
|------|-----|------------|------------|
| model.rs | 278 | 0 | 0 |
| docx/styles.rs | 0 | 90 | 0 |
| pdf/layout.rs | 0 | 30 | 0 |
| fonts/mod.rs | 0 | 0 | 27 |
| docx/mod.rs | 2 | 16 | 2 |
| docx/textbox.rs | 0 | 17 | 0 |
| docx/images.rs | 0 | 13 | 0 |
| docx/numbering.rs | 0 | 12 | 0 |
| docx/runs.rs | 0 | 11 | 0 |
| geometry/definitions.rs | 11 | 0 | 0 |

### Largest functions (split candidates)
| Lines | Location | Function |
|-------|----------|----------|
| 1870 | src/pdf/mod.rs:463 | `render()` — THE major monolith |
| 320 | src/docx/mod.rs:348 | `parse_zip()` |
| 192 | src/geometry/definitions.rs:28 | `lookup()` (data table, not logic) |
| 134 | src/fonts/discovery.rs:110 | `scan_font_dirs()` |
| 98 | src/docx/styles.rs:642 | `resolve_based_on()` |
| 85 | src/docx/textbox.rs:436 | `parse_path_commands()` |
| 79 | src/pdf/table.rs:19 | `auto_fit_columns()` |
| 75 | src/geometry/formulas.rs:143 | `eval_op()` |
| 74 | src/pdf/charts.rs:91 | `draw_marker()` |
| 72 | src/docx/smartart.rs:74 | `parse_dsp_shape()` |
| 68 | src/docx/charts.rs:75 | `parse_series()` |
| 65 | src/docx/runs.rs:109 | `split_run_by_script()` |
| 65 | src/docx/textbox.rs:370 | `parse_custom_geometry()` |

### Key observations
1. **`render()` at 1870 lines is the #1 split candidate** — it's a single function doing all PDF rendering
2. **`parse_zip()` at 320 lines** is the parse orchestrator — moderately large but coherent
3. **`model.rs` has 278 pub items** — massive IR surface, but expected for a bridge between parser and renderer
4. **`docx/styles.rs` has 90 pub(super) items** — mostly struct fields for style types, not overly concerning
5. **`geometry/definitions.rs` at 9876 lines** is a data lookup table (preset shape definitions), not logic — not a refactor target

## Phase 1: Dependency graph 

### Cross-layer isolation
- **docx/ → pdf/: NONE** ✓ Parsing never references rendering
- **pdf/ → docx/: NONE** ✓ Rendering never references parsing
- **All communication via model.rs** — clean two-layer architecture: DOCX → Model → PDF

### Internal dependency graph (within docx/)
```
docx/mod.rs (coordinator)
  ├── styles.rs ← model (Alignment, CellBorder, LineSpacing, TabStop)
  ├── numbering.rs ← standalone (no cross-module deps)
  ├── runs.rs ← images, numbering, styles, textbox (hub module)
  ├── images.rs ← charts, smartart, textbox (branching detection)
  ├── charts.rs ← standalone (minimal deps)
  ├── textbox.rs ← runs, styles, images (⚠ mutual with runs)
  ├── tables.rs ← runs, styles, numbering
  ├── sections.rs ← headers_footers, styles, relationships
  ├── headers_footers.rs ← runs, tables, styles, numbering
  ├── embedded_fonts.rs ← relationships
  ├── alt_chunk.rs ← standalone (minimal deps)
  ├── smartart.rs ← textbox, styles
  └── settings.rs ← standalone
```

### Internal dependency graph (within pdf/)
```
pdf/mod.rs (coordinator, render() 1870-line monolith)
  ├── layout.rs ← standalone (no super:: deps, foundation layer)
  ├── table.rs ← layout, header_footer
  ├── charts.rs ← chart_legend, charts_radial
  ├── charts_radial.rs ← chart_legend, charts (⚠ mutual with charts)
  ├── chart_legend.rs ← charts (drawing primitives)
  ├── header_footer.rs ← layout, table
  ├── footnotes.rs ← layout
  └── smartart.rs ← geometry crate, charts
```

### Managed circular dependencies
1. **docx/textbox.rs ↔ docx/runs.rs** — textbox calls parse_runs, runs calls parse_textbox_from_vml
2. **pdf/charts.rs ↔ pdf/charts_radial.rs ↔ pdf/chart_legend.rs** — mutual helper sharing
- Risk: Low (logically separated, no infinite loops)

### model.rs usage heat map (types used per module)
| Module | # Types | Heaviest types |
|--------|---------|----------------|
| pdf/mod.rs | 18 | Document, Block, Paragraph, Run, FloatingImage, SectionProperties |
| docx/textbox.rs | 14 | Paragraph, ConnectorShape, CustomGeometry, ShapeFill, Textbox |
| docx/tables.rs | 12 | Table, TableCell, TableRow, Paragraph, CellBorders |
| pdf/header_footer.rs | 9 | HeaderFooter, Paragraph, Run, FieldCode, SectionProperties |
| docx/runs.rs | 8 | Run, FloatingImage, InlineChart, Textbox, FieldCode |
| docx/images.rs | 8 | FloatingImage, EmbeddedImage, InlineChart, SmartArtDiagram |
| docx/charts.rs | 8 | Chart, ChartSeries, ChartAxis, ChartType, InlineChart |
| pdf/layout.rs | 5 | Run, Alignment, TabStop, VertAlign |
| docx/styles.rs | 4 | Alignment, CellBorder, LineSpacing, TabStop |
| pdf/charts.rs | 4 | ChartType, InlineChart, MarkerSymbol |

### Key observations
1. **Architecture is clean** — strict DOCX → Model → PDF data flow, no backward references
2. **pdf-writer only in pdf/**, roxmltree only in docx/, ttf-parser only in fonts/ — good crate isolation
3. **docx/runs.rs is the parsing hub** — most modules feed into run parsing
4. **pdf/layout.rs is independent** — no super:: deps, serves as rendering foundation
5. **model.rs is the coupling point** — 278 pub items, but this is expected for an IR bridging two layers
6. **No surprising cross-cutting deps** — geometry used only by pdf/smartart.rs (expected for shape rendering)

## Phase 2: Qualitative Code Reading — Hot Paths 

### `src/pdf/mod.rs` — render() monolith (2365 lines) ⚠️ PRIMARY REFACTOR TARGET

The 1870-line `render()` function (lines 463–2332) has clear internal phases:
1. **Font collection & embedding** (~100 lines) — collects runs from all sections/headers/footers
2. **Character set collection** (~70 lines) — builds used-chars-per-font for subsetting
3. **Font registration** (~30 lines) — registers each unique font key
4. **Image embedding** (~130 lines) — embeds images as PDF XObjects
5. **Section/page layout loop** (~800 lines) — THE CORE: paragraph/table rendering with page overflow
6. **Post-processing** (~200 lines) — column separators, footnotes, headers/footers
7. **PDF assembly** (~200 lines) — gradients, page objects, resource dictionaries

**Key issues:**
- ~30 mutable state variables thread through the entire function (`slot_top`, `prev_space_after`, `current_content`, `current_page_links`, `current_page_footnote_ids`, `current_alpha_states`, `page_gradient_specs`, `styleref_running`, `styleref_page_first`, etc.)
- Page-flush logic is duplicated in 3+ places (section breaks, page overflow, column overflow) — each must update ~8 parallel vectors
- **Extractable phases**: Font collection (phase 1), image embedding (phase 4), and PDF assembly (phase 7) have minimal coupling to the core layout loop and could become standalone functions
- The core layout loop (phase 5) is harder to split because of the shared mutable state, but paragraph rendering (~200 lines inside the `Block::Paragraph` match arm) could potentially be a method on a `PageBuilder` struct that encapsulates the mutable state

### `src/docx/mod.rs` — parse orchestrator (667 lines) ✓ ACCEPTABLE

- `parse_zip()` (320 lines) is the only large function. Structure is clear:
  1. Parse sub-parts (styles, numbering, relationships, fonts, footnotes) — 10 lines
  2. Read and parse document.xml — 5 lines
  3. Iterate block nodes, assembling `Block::Table` or `Block::Paragraph` — 250 lines
  4. Build final section from body-level `sectPr` — 30 lines
- Paragraph assembly (lines 394–616, ~220 lines) follows a consistent pattern: read inline `pPr` property → style fallback → default. This is repeated ~15 times for different properties. Tedious but regular and readable.
- XML utility functions at top (~270 lines: `wml`, `wml_attr`, `wml_bool`, `twips_attr`, `parse_hex_color`, `parse_paragraph_borders`, etc.) are clean, well-focused helpers used by all submodules.
- **Verdict**: Moderately large but coherent. The paragraph assembly block could be extracted but the current form is readable. Not a priority refactor target.

### `src/pdf/layout.rs` — text layout (958 lines) ⚠️ DUPLICATION

Three main functions:
- `build_paragraph_lines()` (137 lines) — word wrapping for normal paragraphs
- `build_tabbed_line()` (280 lines) — word wrapping for tab-containing paragraphs
- `render_paragraph_lines()` (288 lines) — renders pre-built lines to PDF content stream

**Key issues:**
- **Significant duplication** between `build_paragraph_lines` and `build_tabbed_line`: word width calculation, line-wrapping logic, and `WordChunk` construction (~20 fields each) are nearly identical. The tabbed version adds tab-stop resolution and leader fill, but the core word layout is copy-pasted.
- `WordChunk` construction appears in 4 places (normal text, tabbed text, inline images in normal, inline images in tabbed) with identical field patterns — a constructor would reduce repetition.
- `render_paragraph_lines` mixes font state management, text output, decoration drawing (underlines, strikethrough, highlights), and inline image rendering. This is somewhat expected for a render function, and the internal structure is organized with clear sections.
- **No mixed high-level concerns** — all functions are about text layout/rendering. The module is cohesive.
- **Verdict**: Cohesive module with duplication. Extract shared word-layout logic into a common helper; add `WordChunk` constructor.

### `src/model.rs` — intermediate representation (620 lines) ✓ CLEAN

- 278 pub items — mostly struct fields. This is proportional to DOCX feature breadth.
- Types are well-organized by domain: alignment/tabs → headers/footnotes → spacing → sections → fonts → document → images → shapes → textboxes → borders → paragraphs → runs → tables → charts.
- **Format-neutral**: Types use points (f32), RGB colors ([u8; 3]), and domain enums. No DOCX XML artifacts or PDF-specific constructs leak through.
- `Paragraph` has 22 fields, `Run` has 22 fields — large but each maps to a real DOCX property that affects rendering. Both have sensible `Default` impls.
- **Minor concern**: `h_relative_from: &'static str` and `v_relative_from: &'static str` in `FloatingImage` and `Textbox` use string literals ("page", "margin", "column") instead of enums. Type safety improvement opportunity.
- **`#[allow(dead_code)]`** on `ChartType` and `ChartAxis` — some variants/fields not yet used in rendering.
- **Verdict**: Clean IR. Size is appropriate for the domain. No split needed.

### `src/fonts/mod.rs` — font handling (323 lines) ✓ CLEAN

- `register_font()` is the main entry point — clear fallback chain: embedded fonts → filesystem discovery → font table alt names → family-class fallback → Helvetica.
- `FontEntry` is a focused struct with clean methods (`char_width_1000`, `word_width`, `space_width`, `kern_1000`).
- `font_key_buf` is performance-conscious (reuses buffer to avoid allocation on hot path).
- Submodules (`cache`, `discovery`, `embed`, `encoding`) are well-separated by concern.
- No unexpected dependencies — uses only `pdf_writer`, `crate::model`, and its own submodules.
- **Verdict**: Clean, well-bounded module. No concerns.

### Summary of Hot Path Findings
| Module | Status | Action Needed |
|--------|--------|---------------|
| pdf/mod.rs render() | ⚠️ Major monolith | Extract phases 1, 4, 7; consider PageBuilder struct for core loop state |
| pdf/layout.rs | ⚠️ Duplication | Extract shared word-layout logic; add WordChunk constructor |
| docx/mod.rs | ✓ Acceptable | Low priority — paragraph assembly could be extracted but is readable |
| model.rs | ✓ Clean | Minor: string-based relative_from → enums |
| fonts/mod.rs | ✓ Clean | No action needed |

## Phase 2: Qualitative Code Reading — Boundaries 

### `src/lib.rs` — public API (43 lines) ✓ MINIMAL & CLEAN

- Only 2 public functions: `convert_docx_to_pdf(input, path)` and `convert_docx_bytes_to_pdf(bytes, path)`.
- 1 public re-export: `Error`.
- Private helper `render_and_write()` handles timing instrumentation and file write — clean separation.
- All modules are private (`mod docx/pdf/fonts/model/geometry/error`) — only `Error` is re-exported.
- **Minor observation**: `render_and_write` forces `.pdf` extension via `path.as_ref().with_extension("pdf")` — silently overwrites user-provided extension. This is pragmatic but could surprise callers.
- **Verdict**: Excellent API surface. Minimal, clean, hard to misuse.

### `src/error.rs` — error handling (42 lines) ✓ CLEAN

- 5 variants: `InvalidDocx(String)`, `Zip(ZipError)`, `Xml(roxmltree::Error)`, `Pdf(String)`, `Io(io::Error)`.
- Implements `Display`, `std::error::Error`, and `From` for Zip/Xml/Io — idiomatic Rust error pattern.
- `Error::InvalidDocx` used 3 times in `docx/mod.rs` — guards against non-ZIP files and missing `word/document.xml`.
- `Error::Pdf` used once — "Missing w:body" in `docx/mod.rs` (arguably should be `InvalidDocx` since it's a parse-time error, not a PDF rendering error).
- **No error usage in pdf/ module** — `render()` returns `Result<Vec<u8>, Error>` but the function body never produces `Err`. All rendering failures are silently handled (missing fonts → fallback, missing images → skip, etc.). This is a deliberate resilience choice — render what you can, never crash.
- **No `thiserror` dependency** — manual `Display`/`From` impls. Fine at this size; consider `thiserror` if error variants grow.
- **Verdict**: Clean, minimal error handling. The `Error::Pdf("Missing w:body")` is a minor misclassification. The "never fail rendering" approach is intentional and appropriate for the project's goals.

### Summary of Boundary Findings
| Module | Status | Action Needed |
|--------|--------|---------------|
| lib.rs | ✓ Minimal & clean | None (2 pub fns, 1 re-export) |
| error.rs | ✓ Clean | Minor: "Missing w:body" should be InvalidDocx, not Pdf |

## Phase 3: Pattern Detection 

### Code Smells — Self-Assessment Markers
- **Zero TODO/FIXME/HACK/workaround markers** — codebase has no self-flagged technical debt

### Long Match/If-Else Chains
| Location | Arms | Dispatches On | Refactoring Opportunity |
|----------|------|---------------|------------------------|
| docx/runs.rs:456-595 | 13+ | XML element tag names in run nodes | Dispatch table mapping names to handlers |
| pdf/charts.rs:92-164 | 12 | MarkerSymbol enum variants | Trait dispatch or lookup table for shape drawing |
| pdf/mod.rs:97-119,123-144 | 6+ (nested) | Floating image h/v positioning | Extract to PositionResolver helper |
| docx/textbox.rs:442-517 | 8 | DrawingML path commands | Visitor pattern (minor) |

**Verdict**: No extreme chains requiring urgent trait-dispatch refactoring. The runs.rs element dispatch (13 arms) is the largest but is standard XML parsing — a dispatch table would help readability but isn't critical.

### Duplicated Logic Between Modules

**1. Theme color scheme mapping** — 9-line identical match block in textbox.rs and styles.rs mapping dk1/lt1/dk2/lt2/tx1/bg1 etc. Extract to shared `resolve_theme_color()` helper.

**2. Spacing property extraction** — 6-line chain (`ppr → wml("spacing") → twips_attr("before"/"after") + line_spacing`) repeated in **4 modules**: docx/mod.rs, textbox.rs, tables.rs, headers_footers.rs. Extract to `parse_paragraph_spacing()`.

**3. Indent property extraction with bidi fallback** — 4-7 lines repeated in mod.rs, tables.rs, styles.rs. textbox.rs diverges (lacks bidi handling — potential bug). Extract to `extract_all_indents()`.

**4. WordChunk construction** (layout.rs) — 17-19 field struct literal appears 5 times. Text chunk and image chunk patterns each repeat twice (build_paragraph_lines vs build_tabbed_line). Use constructor/builder.

**5. Line wrapping logic** (layout.rs) — 5-6 line pattern (check overflow → push line → reset cursor) appears 4 times. Extract to helper.

### God Structs
| Struct | Fields | Location | Concern |
|--------|--------|----------|---------|
| Paragraph | 28 | model.rs:310-339 | Largest; mixes content, formatting, layout, and floating elements |
| Run | 23 | model.rs:377-403 | Text + all formatting in one flat struct |
| SectionProperties | 20 | model.rs:65-85 | Page geometry + headers + grid + columns + break type |
| Textbox | 18 | model.rs:272-292 | Positioning + content + styling |

**Note**: While large, these structs map 1:1 to real DOCX properties. Splitting them would improve ergonomics but isn't urgent — the property breadth is inherent to the Word format.

### IR Assessment (model.rs)

**Format neutrality**: Mostly clean. Uses points (f32), RGB colors, domain enums. No PDF-specific constructs leak in.

**DOCX-specific terminology leakage**:
- `h_relative_from`/`v_relative_from` use `&'static str` ("page", "margin", "column") — should be enums (flagged in Phase 2)
- `contextual_spacing`, `keep_next`, `keep_lines`, `page_break_before` — Word-specific names, but semantically correct for the domain
- `grid_span`, `v_merge` (VMerge::Restart/Continue) — Word table model vocabulary
- `east_asia_font_name` — Word's dual-font CJK model

**Verdict**: The IR has mild DOCX vocabulary leakage but this is acceptable given the project's single-format scope (DOCX→PDF). If multi-format input is ever planned, these would need generalization. Not a current concern.

**Data shuttle types**: Run, Paragraph, and Textbox are constructed in docx/ and consumed in pdf/ without modification — this is the expected IR pattern, not a smell.

### Render() State Threading (pdf/mod.rs)
- **16 mutable local variables** threaded through nested loops and function calls
- `seen_fonts: &HashMap<String, FontEntry>` passed **113+ times** across 4 modules (mod.rs, table.rs, layout.rs, header_footer.rs) through 3+ call levels — prime candidate for a `RenderContext` struct
- Page-break/flush logic duplicated in 6+ locations, each managing 8+ parallel vectors with `std::mem::take()/push()` — fragile, error-prone
- `render_table()` takes 15 parameters (table.rs:210-225)

### Summary of Phase 3 Findings
| Category | Priority | Finding |
|----------|----------|---------|
| Duplication | Medium | Spacing/indent/color parsing repeated across 4 docx submodules |
| Duplication | Medium | WordChunk + line wrapping duplicated in layout.rs (build_paragraph_lines ↔ build_tabbed_line) |
| God function | High | render() threads 16 mutable vars; page-flush in 6+ places |
| State threading | Medium | `seen_fonts` passed through 3+ levels — context struct candidate |
| IR quality | Low | Mild DOCX vocabulary leakage; acceptable for single-format scope |
| Code smells | None | Zero TODO/FIXME/HACK markers; no extreme dispatch chains |

## Phase 4: Test & Tooling Review 

### Test Organization — All 6 Test Files ✓ WELL-ORGANIZED

**`tests/visual_comparison.rs`** (648 lines) — Clean two-phase pipeline
- Helper functions well-separated: `pdf_page_count`, `screenshot_pdf`, `compare_and_diff`, `ssim_score`, `save_side_by_side`
- `prepare_fixture()` handles convert+screenshot; `score_fixture()` handles comparison+metrics
- `OnceLock` shares prepared fixtures between `visual_comparison` and `ssim_comparison` (the latter is a no-op stub for `cargo test ssim` backwards compat)
- Regression detection via baselines.json with configurable slack (2%)
- Minor: ~30 lines of commented-out timing breakdown code

**`tests/text_boundary.rs`** (356 lines) — Clean analysis pipeline
- Uses mutool stext for structured text extraction
- `extract_page_lines()` clusters lines by y-position (±8pt) to handle super/subscripts
- `normalize_leaders()` strips tab-leader dots for content-focused comparison
- Matching logic: first word + last word per line (tolerates mid-line reformatting)

**`tests/font_validation.rs`** (621 lines) — Comprehensive, intentionally independent
- Implements its own DOCX font extraction (theme resolution, style inheritance, basedOn chains)
- Compares expected fonts (DOCX) vs actual fonts (PDF via mutool)
- Intentionally doesn't use the main crate's parsing — validates from the outside
- Only asserts on `cases/` group failures, not scraped fixtures

**`tests/image_count.rs`** (172 lines), **`tests/page_geometry.rs`** (122 lines), **`tests/file_size.rs`** (122 lines) — Clean, focused single-purpose tests

**`tests/common/mod.rs`** (235 lines) — Well-factored shared infrastructure
- `discover_fixtures()`: SKIPLIST + env var filtering (DOCXIDE_CASE, DOCXSIDE_GROUP)
- `ensure_generated_pdf()`: incremental rebuild (skips if PDF newer than DOCX + src/)
- `read_baselines()`/`update_baselines()`: JSON-based best-ever score tracking
- `log_csv()`: append-only CSV history per metric
- `delta_str()`: human-readable score change formatting

### Tools/ Duplication Analysis

| Tool | Duplication | Severity | Verdict |
|------|-------------|----------|---------|
| jaccard.rs | `is_ink()` + Jaccard computation from visual_comparison.rs | Low | Intentional — standalone CLI, uses simpler per-pixel API |
| case_diff.rs | `is_ink()` + Jaccard + mutool rendering from tests | Low | Intentional — standalone CLI for quick single-case checks |
| case_browser.rs | luma/overlay color mapping from compare_and_diff | Low | Intentional — egui GUI app, needs independence |
| analyze_fixtures.rs | baselines loading, skip list, display_name | ⚠️ Medium | **Stale coupling risk**: `load_skip_list()` parses old SKIP_FIXTURES format (now replaced by SKIPLIST file); `baselines_key()` reimplements display_name |
| docx_fonts.rs | Theme/style/rPr parsing from docx/styles.rs | Low | Intentional — diagnostic tool, raw DOCX inspection |
| docx_inspect.rs | None | — | Standalone zip viewer |
| generate_shapes.rs | None | — | Code generator |

**Key finding**: `analyze_fixtures.rs` has the most concerning duplication:
- `load_skip_list()` parses `tests/common/mod.rs` source code looking for a `SKIP_FIXTURES` constant that no longer exists — skip list is now in `tests/fixtures/SKIPLIST` file
- `baselines_key()` reimplements `common::display_name()` logic — could diverge
- `load_baselines()` uses its own JSON parsing instead of sharing with test infrastructure

### Fixture Organization ✓ CLEAN

- **3 groups**: `cases/` (30 handcrafted), `scraped/` (real-world corpus), `samples/` (additional)
- **SKIPLIST**: Top-level file with comments explaining why each fixture is skipped
- **Output structure**: `tests/output/<group>/<case>/` with `generated.pdf`, `reference/`, `generated/`, `diff/` subdirs
- **baselines.json**: Per-fixture best-ever scores (Jaccard, SSIM, text_boundary)
- **CSV logs**: Append-only history in `tests/output/` (results.csv, ssim_results.csv, etc.)

### Summary of Phase 4 Findings
| Area | Status | Action Needed |
|------|--------|---------------|
| Test files (6) | ✓ Well-organized | No structural issues |
| tests/common/ | ✓ Good shared infra | None |
| tools/ duplication | Low risk | Intentional for CLI independence; jaccard/case_diff/case_browser OK |
| analyze_fixtures.rs | ⚠️ Stale | Fix load_skip_list() to read SKIPLIST file; align baselines_key() with common::display_name() |
| Fixture structure | ✓ Clean | 3 groups, SKIPLIST, baselines.json, CSV history |

## Phase 5: Synthesis 

### 1. Split Candidates — modules doing too many things

| Target | Priority | What to split | Effort |
|--------|----------|---------------|--------|
| **pdf/mod.rs `render()` — 1870 lines** | **HIGH** | Extract into phases: (a) `collect_fonts()` — font collection & embedding (~170 lines, minimal coupling), (b) `embed_images()` — image XObject creation (~130 lines, minimal coupling), (c) `assemble_pdf()` — gradient defs, page objects, resource dicts (~200 lines, minimal coupling), (d) Introduce `PageBuilder` struct to encapsulate the ~16 mutable state variables and page-flush logic (eliminates 6+ duplicated flush sites) | Large — the core layout loop (phase 5, ~800 lines) is tightly coupled to mutable state; phases 1/4/7 are straightforward extractions |
| **pdf/layout.rs `build_paragraph_lines` ↔ `build_tabbed_line`** | **MEDIUM** | Unify shared word-layout logic (word width calc, line-wrapping, WordChunk construction) into a common core, with tab-stop resolution as an extension. Add `WordChunk::new()` / `WordChunk::image()` constructors to replace 5 copy-pasted struct literals | Medium — the two functions share ~60% of their logic |

### 2. Merge Candidates — tiny modules always used together

**None identified.** All modules have clear, distinct responsibilities. Even the smallest modules (settings.rs, embedded_fonts.rs, sections.rs) serve independent purposes and are not always co-invoked. The module granularity is appropriate.

### 3. Abstraction Opportunities — repeated patterns that could be unified

| Pattern | Where repeated | Proposed abstraction | Priority |
|---------|---------------|---------------------|----------|
| **Theme color resolution** | `docx/textbox.rs`, `docx/styles.rs` — identical 9-line match on dk1/lt1/dk2/lt2/tx1/bg1 | Extract `resolve_theme_color(scheme_name, theme) -> Option<[u8; 3]>` into `docx/mod.rs` as a shared helper | Medium |
| **Paragraph spacing extraction** | `docx/mod.rs`, `textbox.rs`, `tables.rs`, `headers_footers.rs` — same 6-line chain | Extract `parse_paragraph_spacing(ppr, styles) -> (Option<f32>, Option<f32>, Option<LineSpacing>)` | Medium |
| **Indent extraction with bidi** | `docx/mod.rs`, `tables.rs`, `styles.rs`; `textbox.rs` diverges (lacks bidi — potential bug) | Extract `extract_all_indents(ind_node, is_bidi) -> (f32, f32, f32, f32)` — also fixes textbox.rs bidi gap | Medium |
| **WordChunk construction** | `pdf/layout.rs` — 5 struct literal sites | `WordChunk::new(text, ...)` and `WordChunk::image(...)` constructors | Low |
| **Line-wrap overflow check** | `pdf/layout.rs` — 4 sites | `LineBuilder::check_overflow()` helper or similar | Low |
| **`seen_fonts` parameter threading** | `pdf/mod.rs`, `table.rs`, `layout.rs`, `header_footer.rs` — passed 113+ times through 3+ call levels | Bundle into `RenderContext { seen_fonts, ... }` struct, reducing parameter counts (e.g., `render_table()` from 15 params) | Medium |

### 4. API Boundary Issues — leaky abstractions, over-exposed internals

| Issue | Location | Severity | Fix |
|-------|----------|----------|-----|
| **`Error::Pdf("Missing w:body")`** | `docx/mod.rs` | Low | Should be `Error::InvalidDocx` — it's a parse-time error, not a PDF rendering error |
| **`model.rs` 278 pub items** | `src/model.rs` | Low | Expected for IR scope. No action needed — all items are consumed by pdf/ and produced by docx/, which is the design intent. Reducing visibility would require pub(crate) on all fields |
| **`analyze_fixtures.rs` stale skip list parsing** | `tools/analyze_fixtures.rs` | Medium | `load_skip_list()` parses source code for a deleted `SKIP_FIXTURES` constant — must be updated to read `tests/fixtures/SKIPLIST` file |
| **`render_and_write` forces .pdf extension** | `src/lib.rs` | Low | Silently overwrites caller-provided extension. Pragmatic but could surprise library users. Document or make configurable if the API is ever stabilized |

### 5. IR Improvements — model.rs changes that would simplify both sides

| Improvement | Impact | Effort |
|-------------|--------|--------|
| **`h_relative_from` / `v_relative_from` → enums** | Replace `&'static str` ("page", "margin", "column") with `enum RelativeFrom { Page, Margin, Column, Character, Paragraph, ... }` in `FloatingImage` and `Textbox`. Enables exhaustive match on the pdf/ side instead of string comparison; catches typos at compile time | Low — ~20 lines changed across model.rs, docx/images.rs, docx/textbox.rs, pdf/mod.rs |
| **Remove `#[allow(dead_code)]` on `ChartType`/`ChartAxis`** | Clean up once all chart variants are implemented in rendering; or narrow the allow to specific unused variants rather than blanket suppression | Low — bookkeeping |

### Prioritized Action List

**Tier 1 — High impact, addresses the biggest pain point:**
1. ~~**Split `render()` phases 1/4/7** into standalone functions (font collection, image embedding, PDF assembly) — reduces the monolith from ~1870 to ~1370 lines with minimal risk~~ ✅ DONE — extracted `collect_and_register_fonts()`, `embed_all_images()`, `assemble_pdf_pages()`; render() reduced to ~1268 lines
2. ~~**Introduce `PageBuilder` struct** to encapsulate mutable render state and unify page-flush logic — eliminates the fragile 6+ duplicated flush sites~~ ✅ DONE — created PageBuilder with flush_page()/push_blank_page() methods; replaced 6 flush sites + 1 in table.rs; render_table reduced from 15 to 8 params; removed 4 padding loops

**Tier 2 — Medium impact, reduces duplication:**
3. ~~**Extract shared docx parsing helpers**: `resolve_theme_color()`, `parse_paragraph_spacing()`, `extract_all_indents()` — each deduplicates 4+ call sites and fixes the textbox.rs bidi gap~~ ✅ DONE — added `resolve_theme_color_key()`, `parse_paragraph_spacing()`, `extract_indents()` to docx/mod.rs; replaced 4+5+4 duplicated sites; fixed textbox.rs bidi gap
4. ~~**Introduce `RenderContext` struct** bundling `seen_fonts` + related read-only state — reduces `render_table()` from 15 params, cleans up call signatures across pdf/~~ ✅ DONE — created `RenderContext<'a>` with `fonts` + `doc_line_spacing`; updated 11 function signatures across 4 files; reduced params by 1 in each function taking both values
5. ~~**Unify `build_paragraph_lines` / `build_tabbed_line`** shared core + `WordChunk` constructors~~ ✅ DONE — added `WordChunk::text()` and `WordChunk::image()` constructors replacing 4 copy-pasted 17-field struct literals; replaced 3 manual `TextLine` constructions with `finish_line()`; ~65 lines removed

**Tier 3 — Low impact, polish:**
6. ~~Fix `Error::Pdf("Missing w:body")` → `Error::InvalidDocx`~~ ✅ DONE — changed to `Error::InvalidDocx` in docx/mod.rs:412
7. ~~Convert `h_relative_from`/`v_relative_from` to enums~~ ✅ DONE — added `HRelativeFrom` (Page/Margin/Column) and `VRelativeFrom` (Page/Margin/TopMargin/Paragraph) enums in model.rs; updated FloatingImage + Textbox struct fields; updated all parsing (images.rs, textbox.rs) and matching (pdf/mod.rs, pdf/header_footer.rs) sites; eliminates string comparison in favor of exhaustive match
~~Fix `analyze_fixtures.rs` stale skip list parsing~~ ✅ DONE — rewrote `load_skip_list()` to read `tests/fixtures/SKIPLIST` directly instead of parsing deleted `SKIP_FIXTURES` constant from source code
~~9. Clean up `#[allow(dead_code)]` on chart types~~ ✅ DONE — removed blanket `#[allow(dead_code)]` on `ChartType` enum and `ChartAxis` struct; narrowed to field-level annotations on specific unused fields (`stacked` in `Bar` variant, `delete` in `ChartAxis`); also narrowed existing `Textbox.margin_bottom` allow and added field-level allows for `Textbox.dist_top`, `SmartArtDiagram.display_width`, `SmartArtDiagram.display_height`
