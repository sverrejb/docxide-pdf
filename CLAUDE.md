# CLAUDE.md

docxide-pdf is a Rust library that converts DOCX to PDF and aims to match
Microsoft Word's PDF export (the online export: tagged, Windows fonts and
metrics) as closely as possible. It is tested against real DOCX files with
Word-generated reference PDFs by visual comparison.

## Commands

```bash
cargo build | cargo check | cargo fmt | cargo clippy
./tools/run-tests.sh                              # ALWAYS this, never raw cargo test
./tools/run-tests.sh --test visual_comparison     # one suite (a module of tests/suite/main.rs)
./tools/run-tests.sh --case case5                 # one fixture, full directory name
./tools/run-tests.sh --test accessibility         # PDF/UA-1 + structure parity (brew install verapdf poppler)
./tools/run-tests.sh --verbose                    # full cargo output, debugging only
cd tools && cargo build                           # analysis tools → tools/target/debug/
```

Analysis tools (`--help` for flags): `analyze-fixtures` (fixture features and
scores: `--failing --fonts --audit --grep`), `docx-inspect` (ZIP entries, pretty
XML, `--grep`), `case-browser` (also `cargo compare`: reference vs generated GUI
with annotations), `accept-baselines` (`--dry-run`, substring filter),
`page-metrics` (release build; the suite's scorer as a CLI).

Python helpers in `tools/` (each has a docstring):
- `engine_compare.py [--open] [--case NAME|glob]... [--html-only] [--shard i/n]`
  — Word | ours | LibreOffice | MiniPdf | rdocx | office2pdf | jubarte | dxpdf |
  WordCraft side by side → `comparison/index.html`; `deploy_comparison.sh`
  publishes it; CI does both on every push to main (`.github/workflows/comparison.yml`).
- `word_export.py <docx|dir>... [--preset online|print]` — unattended Word for
  Mac export (default online = tagged, Windows metrics); `--check` preflight;
  needs the sandbox off.
- Accuracy work: `score_snapshot.sh`, `compare_scores.py`, `ab_view.py`,
  `pdf_lines.py` (rules and baselines of a page), `line_diff.py` (where reflow
  starts), `word_x_diff.py` (per-word x drift), `ink_diff.py` (blue ref-only,
  red ours-only), `docx_edit.py` (literal edit of one part for what-ifs),
  `convert_scan.py` (hangs/crashes), `corpus_score.py` / `corpus_compare.py` /
  `corpus_view.py` (external docx+pdf corpora).

## Planning documents

- `roadmap.md` — open work only, by area. Consult it before starting; update it
  when finishing or discovering things. Finished work lives in git history.
- `current_focus/` — one handover document per active area (tracked changes,
  charts, accessibility, fundamentals, …). Read the one for the area you work
  in; keep it current.
- Specs: query `mcp__local-rag__query_documents` (OOXML spec and PDF spec are
  ingested) before guessing WordprocessingML or PDF semantics.

## Architecture

```
src/
  lib.rs       — public API: convert_docx_to_pdf / convert_docx_bytes_to_pdf / convert_docx_bytes_to_pdf_bytes,
                 each with a `_with(…, &ConvertOptions)` variant (plain ones use ConvertOptions::from_env()).
                 A new option gets three spellings: field + builder method, DOCXIDE_* var in from_vars, CLI flag
                 in main.rs; defaults = Word's online export. Explicit options never read the environment.
  main.rs      — CLI (feature `cli`); error.rs — Error enum
  model/       — IR: mod.rs (Document, Section, Paragraph, Run, fonts, settings), table.rs, drawing.rs
                 (shape geometry, fills, textboxes), chart.rs, math.rs
  docx/        — parsing: mod.rs (relationships, orchestrator), xml.rs (namespaces, element finders, unit
                 and attribute readers), doc_fields.rs (TOC/SEQ/STYLEREF passes over the body), paragraph.rs,
                 runs.rs, field_code.rs (field instruction parsing), styles.rs (style inheritance), theme.rs
                 (theme fonts and fills), numbering.rs, tables.rs, table_borders.rs (border conflicts, vMerge),
                 sections.rs, headers_footers.rs (also footnotes), settings.rs, images.rs (DrawingML pictures,
                 anchors), media.rs (image decoding, blips), vml_object.rs (legacy w:object pictures),
                 textbox.rs (wps shapes), vml_shape.rs (VML shapes and text boxes), charts.rs, chartex.rs,
                 group.rs (wpc canvases / wpg groups flattened), smartart.rs, wordart.rs, svg.rs, emf.rs,
                 wmf.rs, color.rs, comments.rs, embedded_fonts.rs, alt_chunk.rs (HTML), math.rs,
                 math_speech.rs, arabic.rs (presentation-form shaping)
  pdf/         — layout + output: mod.rs (RenderContext, render_pass), page_builder.rs (PageBuilder,
                 LayoutState, anchors, page numbers), paragraph.rs (render_paragraph_block), bands.rs
                 (bands, pre-flow pagination, frames), prepass.rs (document prepass, bookmarks),
                 line_metrics.rs (mark/break line sizing), layout.rs (line breaking, justification, tabs),
                 line_sizing.rs (line heights, baseline grid), line_render.rs (drawing lines: glyphs,
                 decorations, line numbers), link_tagger.rs (links and inline tags), positioning.rs (float
                 placement, float zones, wrap geometry), glyph_grid.rs (Word's device-grid glyph
                 placement), table_layout.rs (row layouts) + table_columns.rs (column widths) +
                 table_split.rs (row splitting) + table.rs (drawing) + cell_draw.rs (borders, shading,
                 cell primitives), header_footer.rs, hf_fields.rs (PAGE/NUMPAGES/IF results), footnotes.rs,
                 list_label.rs, textbox_render.rs, images.rs (XObjects, effects), shape_path.rs (shape
                 paths, clipped pictures), svg.rs, emf.rs, charts.rs / charts_3d.rs / charts_radar.rs /
                 charts_format.rs / charts_radial.rs / charts_ex.rs / chart_legend.rs, smartart.rs,
                 wordart.rs, bidi.rs, comments.rs (comment pane), fonts.rs (per-document font set),
                 tagging.rs (PDF/UA structure tree), assembly.rs (pages, borders, document objects),
                 objstm.rs (object streams), color.rs, helpers.rs (border geometry)
  fonts/       — mod.rs (FontEntry, metrics, font keys), resolve.rs (resolution, substitution, fallback,
                 registration), discovery.rs (cross-platform search), cache.rs (TSV index cache), embed.rs
                 (subsetting, kern/GPOS, line metrics), encoding.rs (WinAnsi)
  geometry/    — OOXML preset shapes (definitions.rs generated by tools/src/bin/generate_shapes.rs),
                 text warps, guide formulas, path evaluation
tests/
  suite/main.rs       — the one integration-test binary; modules visual_comparison, accessibility,
                        text_boundary, font_validation, image_count, page_geometry, file_size,
                        libreoffice_comparison (DOCXIDE_LO_COMPARE=1), speed (#[ignore])
  common/             — shared harness: fixtures, conversion, baselines, per-case score caches,
                        metrics.rs (rasterization, Jaccard/SSIM), a11y.rs, text_boundary.rs
  baselines.json, visual_hashes.json — accepted scores and page hashes (committed)
  fixtures/<group>/<case>/input.docx + reference.pdf; groups cases/ (handcrafted and probe fixtures),
                        fonts/, hyphenation/, scraped/ (real-world); SKIPLIST
  output/<group>/<case>/ — generated.pdf, PNGs, diffs, per-case caches (gitignored)
```

## Testing

- Scores: **Jaccard** on ink pixels (luma < 200; threshold 20.5%) and **SSIM**
  with ±8 px vertical search (threshold 75%), both in
  `tests/suite/visual_comparison.rs`. Accessibility: `ua_deficit` (veraPDF
  PDF/UA-1 rules we fail worse than Word), `a11y_struct` / `a11y_text` (tag
  sequence and block text vs Word's tree), `a11y_missing` (DOCX letters and
  digits absent from our tree; internal only, never in README or site).
- Baselines: `accept-baselines [--dry-run] [case-substring...]` promotes
  `tests/output/latest_scores.json` and page hashes into the two committed
  files after a full visual suite. Accept **improvements** yourself, in the
  same push as the change that earned them; any score **drop** needs the
  user's explicit approval: show the `--dry-run` output and wait. A flag the
  tool does not know is ignored and runs a full accept, so always dry-run
  first.
- Per-case caches under `tests/output/` are reused while newer than their
  inputs: after changing a metric, `rm tests/output/*/*/visual.json
  tests/output/*/*/text_boundary.json`.
- Harness variables: `DOCXIDE_CASE` (exact fixture name), `DOCXIDE_GROUP`,
  `DOCXIDE_IMAGES=1` (side-by-side PNGs), `DOCXIDE_LO_COMPARE`,
  `DOCXIDE_A11Y_GEN`, `LIBREOFFICE_PATH`; `DOCX_BIN` for engine_compare.py.
- Fixture generation: python-docx; custom numbering via ZIP post-processing,
  numIds 100+. References come from `word_export.py --preset online` (the
  four untagged comment-pane references case63/64, door_air and
  missing_font_substitution stay local on purpose).

## Environment variables (library)

`DOCXIDE_FONTS` (extra font dirs, `:`-separated, highest priority;
`.cargo/config.toml` points it at the gitignored `fonts/` mirror of the
private sverrejb/docxide-pdf-assets repo, so `cargo run`/`cargo test` see it
but a bare binary needs `DOCXIDE_FONTS=$PWD/fonts`), `DOCXIDE_NO_FONT_CACHE`,
`DOCXIDE_DEBUG_WRAP` (draw wrap zones), `DOCXIDE_REVISIONS=markup` (All
Markup view; default is the final text), `DOCXIDE_NO_RESAMPLE` (keep picture
pixels instead of the export's 192 ppi). CLI flags: `--fonts`,
`--no-font-cache`, `--debug-wrap`, `--markup`, `--no-resample`.

## Working alongside other sessions

Several agent sessions often run on this repo at once. **Before your first
edit**, check: `git status` (changes that are not yours) and `pgrep -fl
"claude|cargo"` (sandbox off). If another session is active, work in a
worktree, never in the main checkout:

```bash
git worktree add .worktrees/<topic> -b <topic>
ln -s ../../fonts .worktrees/<topic>/fonts     # fonts/ is gitignored; without it every score collapses to Arial
cd .worktrees/<topic> && ./tools/run-tests.sh
git push origin <topic>:main                   # when done; leaves the shared checkout alone
```

Two suites in one checkout share `tests/output/` and corrupt each other's
reports. Never `git checkout`/`switch` branches in the main checkout. `touch
src/lib.rs` before a run after any interrupted suite (the harness reuses
`generated.pdf` by mtime). `run-tests.sh`, git writes and worktree writes
need the sandbox off.

## Word layout notes

- Defaults come from `word/styles.xml` `docDefaults` (modern Word: Aptos 12pt,
  `w:spacing w:after="160" w:line="278"`), never from constants.
- First baseline (PDF y from bottom): `page_height − margin_top − space_before
  − font_size`; aligns cap tops of different faces. `w:docGrid @linePitch`
  (twips) is the baseline pitch for grid-snapped text. Word quantizes
  baselines to 0.25pt online (0.24pt in local exports).
- East Asian line height, CJK substitution, glyph device grid, justification,
  punctuation and float rules are recorded in memory and `current_focus/`;
  derive every rule from Word's PDFs, never fit per-case constants.
- `tests/output/annotations.json` holds click-to-annotate notes from the case
  browser (case, page, Reference/Generated, PDF points x from left / y from
  bottom, `fixed` flag). Read the unfixed ones for the case you work on; the
  browser's **Fix** button runs a headless `claude -p` per note in
  `.worktrees/fix-<id>` (branch `fix/ann-<id>`, log `tests/output/fix-<id>.log`).
