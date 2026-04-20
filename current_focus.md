# Current Focus Areas

## ~~1. Table conditional formatting (`w:tblLook` + `w:tblStylePr`)~~ — ALREADY DONE
Already fully implemented: all 13 condition types parsed, `w:tblLook` (named attrs + hex bitmask), priority-ordered application. Verified working on `stem_partnerships_guide` — purple header rows, banding, styled text all render correctly. Low scores on those fixtures are from text drift, not missing table styling.

## ~~2. Floating tables (`w:tblpPr`)~~ — ALREADY IMPLEMENTED
Already fully implemented: `w:tblpPr` parsing (all attributes), position computation (page/margin/column/text anchors, tblpX/tblpY/tblpXSpec offsets), float zone registration, and text wrapping (narrow paragraphs beside or push below). Verified working on case32/40/45/46 handcrafted test cases. The 3 lowest-scoring fixtures remain low due to unrelated issues:
- **croatian_grant_guidelines (7.9%)**: Text layout height inflation — our single-row floating table computes at 558pt vs Word's ~350pt, causing it to span pages and cascade 68 vs 65 page drift. Root cause is general text layout accuracy (line breaking, font metrics), not floating table code.
- **east_asia_conference_form (8.7%)**: CJK font substitution — Korean/Japanese fonts (HY헤드라인M, 맑은 고딕, etc.) unavailable on test system.
- **korean_japanese_conference_form (10.1%)**: Same CJK font issue as east_asia.

## 3. Run-level shading (`w:rPr/w:shd`) — not yet audited but referenced in roadmap
Run-level background shading (distinct from `w:highlight`) supports arbitrary hex colors and patterns. We handle paragraph and cell shading but skip run-level. Several scraped fixtures with forms and structured documents use colored text backgrounds for emphasis that we currently miss entirely.

## 4. Structured Document Tags (SDT) layout — 2 failing fixtures
`w:sdtContent` content is extracted but the wrapping SDT elements can affect layout. `education_consultant_posting` (13.4% Jaccard, 10 SDT occurrences) is one of the lowest-scoring fixtures, and annotations note "text outside table" issues that may stem from SDT processing. Improving how SDT containers interact with table/paragraph flow could unlock significant gains.

## 5. `w:smallCaps` accuracy fix — 10 fixtures, 5 failing
Currently we uppercase ALL text and shrink ALL of it by 2pt. The spec says only lowercase letters get the size reduction; uppercase stays at original size. This causes visible rendering errors (wrong character sizes throughout smallCaps runs). 68 XML hits across 10 fixtures, 5 of which are failing. A targeted fix — render lowercase chars as uppercase at `font_size - 2pt`, keep uppercase at `font_size` — should be straightforward.
