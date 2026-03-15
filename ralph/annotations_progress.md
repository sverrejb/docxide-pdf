# Progress for Lisa

## 2026-03-16: Fix numbered paragraph indentation (style indent as minimum)

**Case**: `italian_evaluation_minutes` (33.9% Jaccard, passing)

**Annotation**: "Dashes should be indented so they line up with the 'Alla' above" and "Numbers here should be indented so they align with 'Argomenti' above" (page 1).

**Problem**: Numbered paragraphs with direct `w:ind w:left="0" w:firstLine="0"` had their paragraph style's indent completely overridden to zero. In Word, when direct `w:ind` overrides a numbering definition's indent, the paragraph style's indent is preserved as a minimum base. Additionally, when `w:firstLine` is specified in direct formatting, the numbering definition's `w:hanging` should be cleared (since firstLine and hanging are mutually exclusive per OOXML spec).

**Fix** (`src/docx/mod.rs`):
1. For numbered paragraphs with direct `w:ind`, use `max(direct_left, style_left)` instead of just `direct_left`
2. When direct `w:ind` specifies `w:firstLine`, clear the numbering definition's hanging indent

**Result**: Zero new regressions across all fixtures. The dash and numbered list items in the Italian case are now correctly indented at 14.15pt (matching the "Rientrocorpodeltesto" / Body Text Indent style's 283 twips). Overall Jaccard/SSIM scores unchanged (fix affects only a few paragraphs in a 7-page document), but the annotated rendering issue is visually corrected.
