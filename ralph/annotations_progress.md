# Progress for Lisa

## 2026-03-16: Use document's defaultTabStop setting for tab interval

**Case**: `czech_tree_cutting_permit` (57.7% → 72.5% Jaccard, passing)

**Problem**: The layout code used a hardcoded `DEFAULT_TAB_INTERVAL = 36.0pt` (720 twips = 0.5 inches, the US default) for computing default tab stop positions when no explicit tab stops are defined. However, the document's `word/settings.xml` specifies `<w:defaultTabStop w:val="708"/>` (35.4pt = 1.25cm, the European/metric default). This 0.6pt-per-tab discrepancy accumulated across tab-heavy content, shifting text horizontally and causing significant pixel misalignment on the single-page form.

The `defaultTabStop` value was already parsed from settings.xml in `settings.rs` and stored in `DocumentSettings`, but it was never passed through to the `Document` model or the renderer — it was effectively dead code.

**Fix**:
1. `src/model.rs`: Added `default_tab_stop: f32` field to the `Document` struct
2. `src/docx/mod.rs`: Pass `settings.default_tab_stop` when constructing the Document
3. `src/pdf/mod.rs`: Added `default_tab_stop` to `RenderContext`, threaded through all `build_tabbed_line` call sites
4. `src/pdf/layout.rs`: Updated `find_next_tab_stop` and `build_tabbed_line` to accept and use the document's tab interval instead of the hardcoded constant
5. `src/pdf/header_footer.rs`: Updated `build_lines` helper to pass tab stop through

**Result**: Zero REGRESSION flags across all fixtures. Czech tree cutting permit improved +14.8pp Jaccard (57.7→72.5%), +18.6pp SSIM (78.0→96.6%). Fix affects all 48 scraped fixtures with `defaultTabStop` settings — 16 use 708 twips (35.4pt), 2 use 1296 twips (64.8pt), and others use various non-720 values. The 720-twips fixtures (US default) are unaffected since they match the previous hardcoded value.

## 2026-03-16: Fix numbered paragraph indentation (style indent as minimum)

**Case**: `italian_evaluation_minutes` (33.9% Jaccard, passing)

**Annotation**: "Dashes should be indented so they line up with the 'Alla' above" and "Numbers here should be indented so they align with 'Argomenti' above" (page 1).

**Problem**: Numbered paragraphs with direct `w:ind w:left="0" w:firstLine="0"` had their paragraph style's indent completely overridden to zero. In Word, when direct `w:ind` overrides a numbering definition's indent, the paragraph style's indent is preserved as a minimum base. Additionally, when `w:firstLine` is specified in direct formatting, the numbering definition's `w:hanging` should be cleared (since firstLine and hanging are mutually exclusive per OOXML spec).

**Fix** (`src/docx/mod.rs`):
1. For numbered paragraphs with direct `w:ind`, use `max(direct_left, style_left)` instead of just `direct_left`
2. When direct `w:ind` specifies `w:firstLine`, clear the numbering definition's hanging indent

**Result**: Zero new regressions across all fixtures. The dash and numbered list items in the Italian case are now correctly indented at 14.15pt (matching the "Rientrocorpodeltesto" / Body Text Indent style's 283 twips). Overall Jaccard/SSIM scores unchanged (fix affects only a few paragraphs in a 7-page document), but the annotated rendering issue is visually corrected.

## 2026-03-16: Fix vertical alignment in vertically merged table cells

**Case**: `turkish_ancient_religions_plan` (21.4% → 23.1% Jaccard, passing)

**Annotation**: "Text in this column is not vertically centered" (page 1, left column of course schedule table).

**Problem**: When table cells span multiple rows via `w:vMerge`, the `vAlign="center"` calculation only used the current row's height instead of the total merged height. The `merge_spans` data (sum of Continue row heights) was already computed and used for border drawing, but was not applied to the vertical alignment offset calculation. This caused text in vertically merged cells to appear at the top instead of being centered.

**Fix** (`src/pdf/table.rs`):
Applied in three table rendering functions (`render_table_row`, `render_nested_table`, `render_header_footer_table`):
1. Track `cell_grid_col` before incrementing the grid column counter
2. Look up `merge_extra` from `merge_spans` for `vMerge::Restart` cells
3. Use `effective_h = row_h + merge_extra` instead of `row_h` when computing available space for `valign_offset`

**Result**: Zero REGRESSION flags across all fixtures. Turkish case improved +1.7pp Jaccard, +2.4pp SSIM. Fix applies to all 9 fixtures containing `w:vMerge` elements. Small noise-level variations (<0.5pp) on a few unrelated fixtures (no vMerge in those documents).

## 2026-03-16: Fix numbering indent for w:start attribute (logical left indent)

**Case**: `german_mezzo_soprano_bio` (51.2% → 51.3% Jaccard, passing)

**Annotation**: "Incorrect indentation on these bullet points." (page 1, lower section with bullet list under "Who is Huhs?").

**Problem**: The numbering definition indent parser in `numbering.rs` only read `w:ind w:left` for the left indent, but this DOCX (generated by LibreOffice) uses the OOXML logical attribute `w:ind w:start="720"` instead. Per the spec (§17.3.1.12), `w:start` is the logical start-edge indent equivalent to `w:left` for LTR text. The paragraph indent parser in `mod.rs` already handled `w:start` as a fallback for `w:left`, but the numbering parser did not. Result: bullet point indentation was 0pt instead of the correct 36pt (720 twips).

**Fix** (`src/docx/numbering.rs`):
Changed the indent parsing from:
```rust
let indent_left = ind.and_then(|n| twips_attr(n, "left")).unwrap_or(0.0);
```
to:
```rust
let indent_left = ind
    .and_then(|n| twips_attr(n, "start").or_else(|| twips_attr(n, "left")))
    .unwrap_or(0.0);
```
This mirrors the existing pattern in `mod.rs:307` for paragraph indents.

**Result**: Zero REGRESSION flags across all fixtures. German case improved +0.1pp Jaccard, +0.2pp SSIM. Fix affects all documents using `w:start` in numbering definitions (common in LibreOffice-generated DOCX files). Small noise-level variations (≤0.4pp) on a few unrelated fixtures.

## 2026-03-16: Fix tabs inside field results (TOC first entry rendering)

**Case**: `croatian_grant_guidelines` (8.5% Jaccard, failing — dominant issue is floating tables)

**Annotation**: "This first line of the TOC is not rendered right. There is no space between the '1!', the Title and the page number, and the dotted line is missing as well (......)" (page 1).

**Problem**: The first TOC entry rendered as "1Opće informacije4" with no tab spacing or dot leaders. The root cause was in `parse_runs` (`src/docx/runs.rs`): `<w:tab/>` elements inside field result sections were unconditionally skipped by the guard `!in_field`. The first TOC entry is special because it contains the TOC field code (`fldChar begin` + `instrText "TOC..."` + `fldChar separate`), so `in_field` is true for the entire paragraph. Subsequent TOC entries are separate `<w:p>` elements where `in_field` resets to false. Text in field results was already correctly handled (line 531: `in_field_result && !is_dynamic_field(...)`) but tabs used a stricter guard.

Additionally, in `src/pdf/layout.rs`, the tab leader lookup in `build_tabbed_line` re-searched the tab stops array instead of using the already-resolved `stop` variable, which could find the wrong tab stop after line wrapping.

**Fix**:
1. `src/docx/runs.rs`: Changed tab guard from `!in_field` to `!in_field || (in_field_result && !is_dynamic_field(&field_instr))`, matching the text handling condition
2. `src/pdf/layout.rs`: Use the resolved tab stop's leader directly instead of re-searching the tab stops array

**Result**: Zero REGRESSION flags across all fixtures. The first TOC entry now renders with proper tab spacing and dot leaders. Overall Jaccard score unchanged (8.5%) because the dominant issue for this 65-page document is floating tables. Small noise-level variations (≤0.4pp) on a few unrelated fixtures. Fix applies to all documents with tab characters inside non-dynamic field results (common in TOC fields).

## 2026-03-16: Fix missing glyphs in STYLEREF header values (non-breaking space)

**Case**: `bush_fires_act_comparison` (42.2% → 42.4% Jaccard, passing)

**Annotation**: "We are rendering some empty squares here that are not supposed to be here." (page 3, header area) and '"[Heading" should be aligned with the "[2" above.' (page 3, body text).

**Problem**: STYLEREF fields in headers/footers resolve dynamically at render time using text from body paragraphs. The body text in this legal document contains non-breaking spaces (U+00A0) — e.g., "Bush Fires Act\u{a0}1954". The font character collection for STYLEREF fields used a hardcoded set of ASCII characters (`'0'..='9'`, `'A'..='Z'`, `'a'..='z'`, and a few punctuation marks), which did NOT include U+00A0 or other non-ASCII characters. When the STYLEREF resolved to text with non-breaking spaces, those characters were missing from the font subset, causing .notdef glyphs (empty squares) to render and the spacing between words to collapse.

**Fix** (`src/pdf/mod.rs`):
Replaced the hardcoded STYLEREF character set with dynamic collection from the document body. Before the header/footer character collection loop, scan all body paragraphs and collect characters from:
1. All runs in paragraphs that have a `style_id` (paragraph-level STYLEREF targets)
2. All runs that have a `char_style_id` (character-level STYLEREF targets)

This ensures every character that could appear in a STYLEREF value is included in the font subset, regardless of encoding.

**Result**: Zero REGRESSION flags across all fixtures. Bush fires case improved +0.2pp Jaccard, +0.2pp SSIM. Header now correctly renders "Bush Fires Act 1954" with proper non-breaking space instead of collapsed/missing glyph. Fix applies to all documents using STYLEREF fields with non-ASCII characters in the referenced text (common in legal documents with non-breaking spaces).

## 2026-03-16: Fix inline page break before text treated as page-break-after

**Case**: `transition_to_work_deed` (25.6% → 26.1% Jaccard, passing)

**Annotation**: "\"Reader's guide to this deed\" should come on page 2" (page 1).

**Problem**: When `w:br w:type="page"` appears in a run before any text content in the same paragraph, the break should cause text after it to render on the next page. For example:
```xml
<w:p>
  <w:r><w:br w:type="page"/></w:r>
  <w:r><w:t>Reader's Guide to this Deed</w:t></w:r>
</w:p>
```
Our code always set `has_page_break_after = true`, which rendered the entire paragraph (including "Reader's Guide") on the current page, then triggered a page break. The text appeared on page 1 instead of page 2.

**Fix** (`src/docx/runs.rs`):
When `w:br type="page"` is encountered and no text content has been emitted yet (`runs.is_empty() && pending_text.is_empty()`), set a `page_break_before_content` flag instead of `has_page_break_after`. This flag is then OR'd into `has_page_break_before`, causing the paragraph's text to render on the new page.

**Result**: Zero REGRESSION flags across all fixtures. Three fixtures improved:
- `transition_to_work_deed`: +0.5pp Jaccard (25.6→26.1%), +0.5pp SSIM (37.4→37.9%)
- `federal_procurement_terms`: +2.1pp Jaccard (51.7→53.8%), +6.7pp SSIM (67.1→73.8%)
- `go_math_grade4_guide`: +1.3pp Jaccard (26.4→27.7%), +3.2pp SSIM (46.1→49.3%)
Fix applies to all 9 fixtures containing `w:br w:type="page"` inline breaks.

## 2026-03-16: Fix empty paragraph height in table cell rendering

**Case**: `italian_project_proposal` (28.3% → 28.5% Jaccard, 57.0% → 57.6% SSIM, passing)

**Annotation**: "\"Progretto 1\" is too high up compared to reference. Also should not be broken into two lines." (page 1, left column of form table).

**Problem**: In vertically merged table cells, empty paragraphs (paragraphs with no text runs) were not consuming their line height during rendering, even though the layout computation correctly accounted for it. The merged cell in this document has 7 empty paragraphs before the "progetto 1" text, which should push it down ~165pt. But during rendering in `render_cell_content`, the cursor advance for empty paragraphs was `para.space_before + para.lines.len() * para.line_h`. Since empty paragraphs have `lines = vec![]` (set during `compute_row_layouts`), `lines.len() = 0` and only `space_before` was subtracted — losing ~109pt of vertical space. This caused text to appear near the top of the merged cell instead of at the correct position 2/3 down.

The layout computation in `compute_row_layouts` correctly handled this: for empty text with `content_height = 0`, it added `line_h` to `total_h`. But the rendering code did not match.

**Fix** (`src/pdf/table.rs`):
1. Added `para_block_height()` helper that returns the correct height for a paragraph's content block, matching the layout computation: `line_h` for empty paragraphs (no lines, no content_height), `content_height` when present, or `lines.len() * line_h` for non-empty.
2. Applied consistently across 6 locations: `render_cell_content` (empty para branch), content_h recalculations for vAlign in `render_table_row` and `render_nested_table`, `find_cell_split`, and `render_partial_row`.

**Result**: Zero REGRESSION flags across all fixtures. Multiple fixtures improved:
- `italian_project_proposal`: +0.2pp Jaccard (28.3→28.5%), +0.6pp SSIM (57.0→57.6%)
- `traditional_skills_job_form`: +9.2pp SSIM (37.6→46.8%)
- `parish_housing_data_profile`: +5.9pp SSIM (55.8→61.7%)
- `education_consultant_posting`: +0.9pp SSIM (19.2→20.1%)
Fix affects all documents with empty paragraphs in table cells, which is extremely common in form-style documents. Small noise-level variations (≤0.4pp) on a few unrelated fixtures.

## 2026-03-16: Handle w:noBreakHyphen run element

**Case**: `bush_fires_act_comparison` (42.4% → 42.7% Jaccard, passing)

**Problem**: The `<w:noBreakHyphen/>` run element was not handled in the run parser (`src/docx/runs.rs`). This element represents a non-breaking hyphen — visually a hyphen character that prevents line breaking at that position. Without handling it, the hyphen was silently dropped, causing missing characters in the rendered text. The `bush_fires_act_comparison` document (a legal comparison of Bush Fires Act amendments) contains 53 instances of `w:noBreakHyphen`, and `centrifugal_water_chillers` contains 9.

**Fix** (`src/docx/runs.rs`):
Added a `"noBreakHyphen"` match arm in the run element parser that appends a regular hyphen character (`'-'`) to the pending text. Used ASCII hyphen-minus (U+002D) rather than Unicode non-breaking hyphen (U+2011) since U+2011 is not in the WinAnsi encoding used for PDF text output, and our line breaker only splits on whitespace so the non-breaking property is already implicit.

**Result**: Zero REGRESSION flags across all fixtures. Two fixtures improved:
- `bush_fires_act_comparison`: +0.3pp Jaccard (42.4→42.7%), +0.6pp SSIM (81.9→82.5%)
- `centrifugal_water_chillers`: +1.0pp Jaccard (79.2→80.2%), +0.8pp SSIM (93.2→94.0%)
Small noise-level variations (≤0.4pp) on a few unrelated fixtures.

## 2026-03-16: Apply paragraph indentation in header/footer rendering

**Case**: `lithuanian_ethics_law` (36.6% → 40.2% Jaccard, passing)

**Problem**: Header/footer paragraphs with `w:ind` attributes (left indent, right indent, firstLine, hanging) had their indentation completely ignored. The header/footer parser in `src/docx/headers_footers.rs` never extracted `w:ind` from paragraph properties — the `Paragraph` struct was always constructed with `..Paragraph::default()` which sets all indent fields to 0. Additionally, the header/footer renderer in `src/pdf/header_footer.rs` always passed 0 for indent parameters to `build_lines()` and `render_paragraph_lines()`.

In the `lithuanian_ethics_law` document, the header contains two paragraphs ("Projekto Nr." and "lyginamasis variantas") with `w:ind w:firstLine="6237"` (311.85pt first-line indent), which should position them at the right side of the header area. Without indent support, they rendered at the left margin.

**Fix**:
1. `src/docx/headers_footers.rs`: Added indent parsing via `extract_indents()` for header/footer paragraphs, with style fallback when no direct `w:ind` is specified. Set `indent_left`, `indent_right`, `indent_hanging`, and `indent_first_line` on the Paragraph struct.
2. `src/pdf/header_footer.rs`: Updated `build_lines()` to accept `indent_left` and `text_hanging` parameters. Compute `para_text_x`, `para_text_width`, and `text_hanging` from paragraph indent fields in the main paragraph rendering path. Pass `text_hanging` to `render_paragraph_lines()`. Also updated textbox paragraph rendering within headers/footers to use indent-adjusted widths.

**Result**: Zero REGRESSION flags across all fixtures. Lithuanian ethics law improved +3.6pp Jaccard (36.6→40.2%), +3.6pp SSIM (55.1→58.7%). Fix affects all documents with indented paragraphs in headers/footers. Small noise-level variations (≤0.4pp) on a few unrelated fixtures.
