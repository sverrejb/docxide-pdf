# Lisa Progress

## Completed: Annotation #49 — Links overflow page width (croatian_grant_guidelines)

**Status**: Already fixed. URLs wrap correctly within margins via `unicode_linebreak` segments.

## Completed: Annotations #50, #51 — Footer issues (croatian_grant_guidelines)

**Status**: Already fixed. Footers render correctly on all pages, matching reference positions.

## Completed: Annotation #54 — Table cells wrong (go_math_grade4_guide)

**Status**: Fixed. Inferred gridSpan from cell width when not explicit. Jaccard +0.2pp, no regressions. Committed as 928d052.

## Completed: Annotation #56 — Text wrapping around image (indonesian_benchmarking_guide)

**Status**: Fixed. Lowered the wrap-to-TopAndBottom threshold from 90% to 50% of text_width for Square/Tight/Through floating images. Images wider than half the text area now displace text vertically (TopAndBottom) instead of allowing text to wrap beside them. Jaccard -4.1pp (40.5% → 36.4%) due to cascading page breaks — structurally correct, text now flows below image matching reference.
