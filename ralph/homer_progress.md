# Progress for Homer — implement features to make Maggie's cases render right

Homer reads Marge's gap detail (`ralph/marge_progress.md`) and Maggie's case
ledger (`ralph/maggie_progress.md`), then works ONE case per iteration: implement
the rendering feature the case isolates, score it against its Word `reference.pdf`,
and commit on improvement. It works only on Maggie's FIXTURE-PENDING-REF cases.

## Cases

Status values:
- **DONE** — crosses thresholds (Jaccard ≥20% / SSIM ≥75%), committed, removed from SKIPLIST.
- **IMPROVED** — score went up but still below threshold (committed, left in SKIPLIST).
- **BLOCKED** — can't fix now; note why and what was learned.
- **IN-PROGRESS** — started, not finished.

| Case | Gap / feature | Status | Score (Jaccard / SSIM) | Notes |
|---|---|---|---|---|
| case66 | Run text-effect toggles `w:outline`/`w:shadow`/`w:emboss`/`w:imprint` + `w:effect` (§17.3.2.23/.31/.13/.18) | IMPROVED | 63.1/60.4 → 74.1/70.6 | Parsed all four bare toggles via `wml_bool` in `src/docx/runs.rs`; `w:effect` (animated shimmer) intentionally ignored — no print form. `outline` → hollow glyph (`TextOutline` stroke + `TextFill::NoFill`, hairline width); `shadow` → light gray drop shadow down-right (offset = 0.035·fontSize); `emboss`/`imprint` → prominent gray drop shadow at 0.7× that offset. Render path in `src/pdf/layout.rs` draws an offset gray glyph copy behind the face (`WordChunk.text_shadow`). Jaccard clears the 20% threshold by 54pp; SSIM lands 4.4pp short of 75%. Confirmed empirically (face-thickening via FillStroke *lowered* SSIM to 68%, so the reference's extra ink is the shadow, not a heavier face) that single-offset shadow ≈ the ceiling — Word's exact multi-pass face/highlight/shadow antialiasing for these legacy effects isn't pixel-replicable cheaply. Full suite: no regressions (197 unchanged). |
| case67 | `w:customXml` transparent wrapper at block / inline-run / cell / row placements (§17.5.1.1/.3/.5/.6) | DONE | 25.0/26.1 → 71.8/91.3 (TxtBnd 0→100) | `w:customXml` is a transparent content wrapper whose children ARE the content (no `sdtContent` nesting, unlike `w:sdt`). Two edits: (1) `collect_block_nodes` in `src/docx/mod.rs` now recurses into `w:customXml` children — one fix covers block (body), row (tbl→tr) and cell (tr→tc) placements since tables already descend via `collect_block_nodes`; (2) `collect_run_nodes` in `src/docx/runs.rs` adds `customXml` to the `ins`/`smartTag` passthrough arm for the inline-run placement. Previously this content was silently dropped. Crosses both thresholds. Full suite: no regressions (197 unchanged). |
