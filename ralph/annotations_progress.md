# Annotations Progress

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
