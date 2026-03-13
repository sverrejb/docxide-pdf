# Progress for ralph/plan_deed_fix.md

## Fix 1: numStyleLink/styleLink resolution — DONE

Changes in `src/docx/numbering.rs`:
1. Added `#[derive(Clone)]` to `LevelDef` struct
2. Added `num_style_link` and `style_link_target` HashMaps to track link attributes during abstractNum parsing
3. After the main parsing loop, resolve numStyleLink → styleLink chains by copying level definitions from the target abstractNum to the source when the source has no levels

Results: Jaccard 6.3% → 25.6% (+19.3pp), SSIM 18.2% → 37.7% (+19.5pp). No regressions.

## Fix 2: Suppress Hyperlink char style for anchor-only hyperlinks — DONE

Changes in `src/docx/runs.rs`:
1. Extended `collect_run_nodes()` output tuple to `(Node, Option<String>, bool)` — third element is `is_anchor_hyperlink` flag
2. For `w:hyperlink` nodes: detect anchor-only hyperlinks (`w:anchor` present, `r:id` absent) and set flag to `true`
3. In the main run loop, suppress character style lookup when `is_anchor_hyperlink` is true — this prevents the Hyperlink style (blue + underline) from being applied to TOC internal links

Results: No score change (25.6% Jaccard, 37.6% SSIM) — TOC color is a subtle visual difference. No regressions across full suite.
