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
