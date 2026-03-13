# Progress for ralph/plan_samtale.md

## Fix 1A: Parse label rPr properties in `src/docx/numbering.rs` — DONE
- Added `label_font_size: Option<f32>`, `label_bold: bool`, `label_color: Option<[u8; 3]>` fields to `LevelDef`
- Created `ListLabelInfo` struct to replace the `(f32, f32, String, Option<String>)` tuple return type
- Parsed `w:sz`, `w:b`, `w:color` from `w:lvl/w:rPr` in the abstractNum level parsing
- Changed `parse_list_info` return type to `ListLabelInfo` and updated all early returns + final return
- Added `parse_hex_color` and `wml_bool` to imports
- Note: callers (mod.rs, tables.rs, textbox.rs) now have type errors — will be fixed in tasks C, D, E

## Fix 1B: Add label properties to Paragraph in `src/model.rs` — DONE
- Added `list_label_font_size: Option<f32>`, `list_label_bold: bool`, `list_label_color: Option<[u8; 3]>` fields to `Paragraph` struct (after `list_label_font`)
- Added corresponding defaults in `Default` impl (`None`, `false`, `None`)

## Fix 1C: Pass label properties to Paragraph in `src/docx/mod.rs` — DONE
- Added `ListLabelInfo` to the import from `numbering`
- Destructured `parse_list_info` result as `ListLabelInfo` struct instead of tuple
- Added `list_label_font_size`, `list_label_bold`, `list_label_color` fields to the `Paragraph` construction
- Note: tables.rs and textbox.rs still have type errors — will be fixed in tasks D and E

## Fix 1D: Pass label properties to Paragraph in `src/docx/tables.rs` — DONE
- Added `ListLabelInfo` to the import from `numbering`
- Destructured `parse_list_info` result as `ListLabelInfo` struct instead of tuple
- Added `list_label_font_size`, `list_label_bold`, `list_label_color` fields to the `Paragraph` construction
- Note: textbox.rs still has type error — will be fixed in task E

## Fix 1E: Pass label properties to Paragraph in `src/docx/textbox.rs` — DONE
- Added `ListLabelInfo` to the import from `numbering`
- Destructured `parse_list_info` result as `ListLabelInfo` struct instead of tuple
- Added `list_label_font_size`, `list_label_bold`, `list_label_color` fields to the `Paragraph` construction
- All callers now updated — no more type errors from the ListLabelInfo change

## Fix 1F: Font registration for bold label in `src/pdf/mod.rs` — DONE
- Modified the label font key construction in the font registration loop (line ~765-777)
- When `list_label_font` is set: appends `/B` suffix if `list_label_bold` is true
- When falling back to first run: creates a temporary `Run` with `bold: list_label_bold || run.bold` to get the correct font key
- This ensures bold label fonts get registered and subsetted correctly

## Fix 1G: Update `label_for_paragraph` in `src/pdf/mod.rs` — DONE
- Rewrote `label_for_paragraph` to build font key with `/B` suffix when `list_label_bold` is true
- When `list_label_font` is set: appends `/B` suffix (matching font registration in Fix 1F)
- When falling back to first run: creates a `Run` with `bold: list_label_bold || run.bold` and uses `font_key()`
- Falls back to `("", vec![])` if font not found in `seen_fonts`
- Note: `label_for_run` is now unused (dead code) — will be cleaned up in simplification

## Fix 1H: Label rendering (3 sites) in `src/pdf/mod.rs` — DONE
- Updated all three label rendering sites to use `list_label_color` and `list_label_font_size`
- **Main rendering (~line 2011)**: color uses `para.list_label_color.or(first_run_color)`, font size uses `para.list_label_font_size.unwrap_or(font_size)`
- **Split paragraph rendering (~line 1701)**: same pattern
- **Textbox rendering (~line 258)**: same pattern, using `tp.list_label_font_size.unwrap_or(tb_fs)`
- Color reset now checks `label_color.is_some()` instead of `first_run_color.is_some()`

## Fix 1I: First-line height for label font size in `src/pdf/mod.rs` — DONE
- Modified the `else` branch of `content_h` calculation (~line 1582) for normal text paragraphs
- When `para.list_label_font_size` is set and larger than `font_size`, the first line uses `resolve_line_h(effective_ls, label_fs, tallest_lhr)` for a taller first line
- For multi-line paragraphs: `first_line_h + (num_lines - 1) * line_h`
- This is the critical fix for column distribution — with 13 questions having 20pt labels but 10pt body text, the compounded height error was shifting column break points

## Fix 1J: Label rendering in `src/pdf/header_footer.rs` — DONE
- Updated label rendering in header/footer textbox code (line ~285-304) to use `list_label_color` and `list_label_font_size`
- Color uses `tp.list_label_color.or_else(|| tp.runs.first().and_then(|r| r.color))` — same pattern as other sites
- Font size uses `tp.list_label_font_size.unwrap_or(tb_fs)`
- Color reset now checks `label_color.is_some()` instead of checking first run color directly

## Fix 2: `w:sym` element parsing in `src/docx/runs.rs` — DONE
- Added `"sym"` match arm before the `_ => {}` catch-all in the run child element loop
- Flushes pending text first so the symbol gets its own Run with the symbol font
- Reads `w:font` attribute for font name (falls back to current run format font)
- Parses `w:char` as hex → u32 → char (PUA codepoint, e.g. U+F04A for Wingdings smiley)
- Creates a Run with the symbol font and character, inheriting formatting from current run
- The PUA character maps through the symbol font's cmap to the correct glyph
