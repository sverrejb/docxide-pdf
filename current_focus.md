# Current Focus Areas

## ~~1. Table conditional formatting (`w:tblLook` + `w:tblStylePr`)~~ — ALREADY DONE
Already fully implemented: all 13 condition types parsed, `w:tblLook` (named attrs + hex bitmask), priority-ordered application. Verified working on `stem_partnerships_guide` — purple header rows, banding, styled text all render correctly. Low scores on those fixtures are from text drift, not missing table styling.

## ~~2. Floating tables (`w:tblpPr`)~~ — ALREADY IMPLEMENTED
Already fully implemented: `w:tblpPr` parsing (all attributes), position computation (page/margin/column/text anchors, tblpX/tblpY/tblpXSpec offsets), float zone registration, and text wrapping (narrow paragraphs beside or push below). Verified working on case32/40/45/46 handcrafted test cases. The 3 lowest-scoring fixtures remain low due to unrelated issues:
- **croatian_grant_guidelines (7.9%)**: Text layout height inflation — our single-row floating table computes at 558pt vs Word's ~350pt, causing it to span pages and cascade 68 vs 65 page drift. Root cause is general text layout accuracy (line breaking, font metrics), not floating table code.
- **east_asia_conference_form (8.7%)**: CJK font substitution — Korean/Japanese fonts (HY헤드라인M, 맑은 고딕, etc.) unavailable on test system.
- **korean_japanese_conference_form (10.1%)**: Same CJK font issue as east_asia.

## ~~3. Run-level shading (`w:rPr/w:shd`)~~ — ALREADY IMPLEMENTED
Parsed in `src/docx/runs.rs:352` via `parse_run_shd`, stored as `Run.shading`, rendered as a background rectangle in `src/pdf/layout.rs:1359` (drawn before highlight, after layout). Spec'd in `src/model/mod.rs:309`.

## 4. Structured Document Tags (SDT) layout — 2 failing fixtures
`w:sdtContent` content is extracted but the wrapping SDT elements can affect layout. `education_consultant_posting` (13.4% Jaccard, 10 SDT occurrences) is one of the lowest-scoring fixtures, and annotations note "text outside table" issues that may stem from SDT processing. Improving how SDT containers interact with table/paragraph flow could unlock significant gains.

## ~~5. `w:smallCaps` accuracy fix~~ — ALREADY CORRECT
`src/pdf/layout.rs:349` `smallcaps_segments()` applies the spec-correct per-character rule (lowercase → uppercase glyph at `font_size - 2pt`, uppercase stays at `font_size`). Unit tests at `src/pdf/layout.rs:1824+`. Listed in roadmap as TODO but the code is fine.

## 6. `w:dstrike` — ALREADY IMPLEMENTED
Double strikethrough renders two parallel lines at `src/pdf/layout.rs:1595-1601`.

## 7. Wide `wrapSquare` floating images push subsequent paragraphs into useless slivers (NEW — discovered 2026-05)
`brazilian_logistics_study` (16.8% Jaccard, 4 unfixed annotations) has 415pt-wide `wrapSquare` images on a ~500pt column. The float-zone wrapping in `src/pdf/mod.rs:1214` uses a 72pt minimum-side threshold; with 415pt images this leaves ~80pt on one side, just above the threshold, so we wrap "Fonte:" caption text into an 80pt sliver beside the image instead of pushing it below. Word treats this as `wrapTopAndBottom`. The threshold needs tuning (likely a fractional rule like ≥35% of column width) but a previous naive clamp regressed `stem_partnerships_guide`, so requires care.

## ~~8. Blank page emitted after `<w:br type="page">` in section-break paragraph~~ — FIXED 2026-05-12
Root cause: `Paragraph.page_break_before` conflated two distinct sources — the idempotent `<w:pageBreakBefore/>` style property and the unconditional `<w:br w:type="page"/>` run-level break. The renderer's `at_top` skip (correct for the property) silently dropped explicit breaks that immediately followed a section break, collapsing two page advances into one. Now tracked separately as `Paragraph.page_break_before_explicit` (parsed from `page_break_before_content` in `parse_runs`), and the renderer flushes whenever the break is explicit even if already at the page top. Lifted `victorian_universal_design_policy` from 0.1671→0.3525 Jaccard with zero regressions across the corpus.
