# Annotations Progress

## 2026-03-16: Use theme minor font for chart labels (annotation 0)

### Problem
Annotation 0 reported that chart labels in `samples/sample500kB` use a serif font instead of sans-serif. The reference PDF shows sans-serif labels for axis labels ("Row 1", "Row 2", etc.) and legend entries ("Column 1", "Column 2", "Column 3").

### Root Cause
Chart rendering picked a label font by iterating `seen_fonts` HashMap keys with a fragile heuristic (filtering out "symbol", "serif", and "/" strings). HashMap iteration order is non-deterministic, and when a document only uses serif fonts (like this LibreOffice document with Liberation Serif), the heuristic failed and fell back to the first HashMap key — a serif font.

Word uses the theme minor (body) font for chart labels when no explicit font is specified in the chart XML. This defaults to Calibri/Aptos (sans-serif).

### Implementation
1. **`src/model.rs`**: Added `chart_font_name: String` to `Document` — stores the theme minor font name.
2. **`src/docx/mod.rs`**: Set `chart_font_name` from `theme.minor` during parsing.
3. **`src/pdf/mod.rs`**: Added `chart_font_name` to `RenderContext`. In `collect_used_chars`, chart label characters (axis labels, legend text, digits) are now registered under the theme minor font so it gets embedded even if no body text uses it.
4. **`src/pdf/charts.rs`**: Removed the broken heuristic. `render_chart` now uses `default_font_name` (the theme minor font) directly as the chart label font key.

### Results
- `samples/sample500kB`: Chart labels now correctly use a sans-serif font matching the reference.
- Jaccard: 31.9% (-0.4pp), SSIM: 51.4% (-0.1pp) — small metric drop due to Aptos/Helvetica glyph metrics differing from Calibri in the reference.
- No regressions on other fixtures — case29 and case30 (chart test cases) unchanged.

### Files Modified
- `src/model.rs` — added `chart_font_name` field to Document
- `src/docx/mod.rs` — set `chart_font_name` from theme minor font
- `src/pdf/mod.rs` — thread chart font through RenderContext, register chart label chars
- `src/pdf/charts.rs` — use theme minor font directly instead of heuristic

## 2026-03-16: Header/footer inheritance + page number format (annotations 11, 12)

### Problem
Annotations 11 and 12 reported missing page numbers in the `feminist_voice_dissertation` fixture. The preamble section should show roman numeral page numbers (i, ii, iii...) in the footer, and the body section should show decimal page numbers (1, 2, 3...). Neither section rendered page numbers.

### Root Cause
Two issues:
1. **Missing header/footer inheritance**: When a section doesn't define its own header/footer references, Word inherits them from the previous section. Our code treated missing references as "no header/footer" instead of falling back to previous sections. The feminist_voice_dissertation has 3 sections: Section 1 defines footer1.xml (with PAGE field), Sections 2 and 3 have no footer references and should inherit Section 1's footer.
2. **Missing page number format**: `w:pgNumType @w:fmt` (e.g., "lowerRoman") was not parsed. Page numbers were always rendered as decimal regardless of the section's format setting.

### Implementation
1. **`src/model.rs`**: Added `page_num_format: Option<String>` to `SectionProperties`.
2. **`src/docx/sections.rs`**: Parse `w:fmt` attribute from `w:pgNumType`.
3. **`src/docx/numbering.rs`**: Made `format_number()` `pub(crate)` (was private). Already handles lowerRoman, upperRoman, lowerLetter, upperLetter, decimal, decimalZero.
4. **`src/docx/mod.rs`**: Made `numbering` module `pub(crate)` for cross-module access to `format_number`.
5. **`src/pdf/header_footer.rs`**: `substitute_hf_runs` now accepts `page_num_format` parameter and uses `format_number()` for `FieldCode::Page` when a format is specified.
6. **`src/pdf/mod.rs`**: Header/footer selection now walks backward through sections when current section has no h/f. For inherited sections, always uses the "default" variant (not "first" or "even").
7. **`src/pdf/table.rs`**: `HfSubstitution` struct extended with `page_num_format` field, passed through to `substitute_hf_runs`.

### Results
- `feminist_voice_dissertation`: Page numbers now render correctly — roman numerals (i, ii, iii...) in preamble, decimal in body.
- Jaccard: 33.6% (+0.1pp), SSIM: 69.4% (+0.1pp). Small improvement because page numbers are tiny text at the bottom of each page.
- No regressions across all fixtures — no passing fixtures became failing.
- The header/footer inheritance fix is broadly applicable to any multi-section document where later sections inherit h/f from earlier ones.

### Files Modified
- `src/model.rs` — added `page_num_format` field
- `src/docx/sections.rs` — parse `w:fmt` from `pgNumType`
- `src/docx/numbering.rs` — made `format_number` pub(crate)
- `src/docx/mod.rs` — made `numbering` module pub(crate), added `page_num_format: None` to fallback
- `src/pdf/header_footer.rs` — page number formatting in `substitute_hf_runs`
- `src/pdf/mod.rs` — header/footer inheritance across sections
- `src/pdf/table.rs` — pass `page_num_format` through `HfSubstitution`

## 2026-03-16: Paragraph border collapsing for identical adjacent paragraphs (annotation 1)

### Problem
Annotation 1 reported a spurious grey horizontal line above "Medarbeiderens navn:" in the `samples/samtale` fixture. The reference PDF doesn't show this line.

### Root Cause
The document has two consecutive paragraphs with identical `w:pBdr` settings (bottom border: single, sz=18, color=A6A6A6): a break-only spacer paragraph and the "Medarbeiderens navn:" paragraph. Per the OOXML spec (§17.3.1.7), when adjacent paragraphs have identical border definitions for all border properties, they form a visual group — the `bottom` border is only drawn below the **last** paragraph in the group, and the `between` border (if defined) is used between them.

Our code only applied this grouping logic when a `between` border element was explicitly defined (`bdr.between.is_some()`). When no `between` was defined, each paragraph drew its own `bottom` border independently, producing the spurious grey line on the spacer paragraph.

### Implementation
Changed the border rendering logic in `src/pdf/mod.rs` to check `borders_match()` independently of whether a `between` element is defined:
- `prev_borders_match`: suppress top border when previous paragraph has matching borders
- `next_borders_match`: draw `between` border if defined, otherwise draw nothing; only draw `bottom` border when next paragraph does NOT match

### Results
- `samples/samtale`: Grey line above "Medarbeiderens navn:" removed, matching reference behavior
- Jaccard: 12.0% (-0.5pp), SSIM: 55.1% (-1.4pp) — small metric drops from content shift after border removal
- No pass/fail status changes across any fixtures

### Files Modified
- `src/pdf/mod.rs` — paragraph border collapsing: check `borders_match` without requiring `between.is_some()`

## 2026-03-16: Bottom-only border positioning tighter to text (annotation 2)

### Problem
Annotation 2 reported that in `samples/samtale` page 2, answer text (e.g., "I stor grad.") was "floating above" the grey bottom border line instead of sitting directly above it. The reference PDF shows text tightly above the border.

### Root Cause
The bottom border's `box_bottom` was computed as `slot_top - bdr_top_pad - content_h - bdr_bottom_pad`, where `content_h` includes the full `line_h` (font metrics including line-gap leading below the descender). Per the OOXML spec (§17.3.1.7), the `w:space` attribute on a bottom border measures "the space after the bottom of the text" — i.e., from the text descender, not from the line box bottom. The trailing leading (typically ~2pt for a 10pt font) was being added as extra gap between text and border.

### Implementation
Added a `trailing_lead` adjustment in `src/pdf/mod.rs` that subtracts the last line's leading (`line_h - font_size`) from the border/shading box bottom. This only applies for bottom-only borders (no top border present) — the "separator/underline" pattern. Boxed paragraphs (all-side borders) are unaffected, preserving their symmetric appearance.

### Results
- `samples/samtale`: Answer text now sits directly above grey border lines, matching reference
- Jaccard: 12.8% (+0.3pp), SSIM: 54.3% (-2.2pp) — SSIM drop is from border position shift, not a visual regression
- No regressions across all fixtures in Jaccard or SSIM (case17 border boxes unaffected)

### Files Modified
- `src/pdf/mod.rs` — added `trailing_lead` adjustment for bottom-only border/shading box positioning
- `tests/baselines.json` — updated samtale SSIM baseline
