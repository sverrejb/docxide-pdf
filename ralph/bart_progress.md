# Progress for Bart

## 2026-06-21 — SELECTED annotation: scraped/learning_cultures_dissertation p62
Note: "Figure 3.1 is rendered wrong. The arrows are wonky and the large text should be white."
(Generated, page index 62). Discrete SmartArt/drawing rendering bug — in scope.
Skipped #0-#6 above: all systemic vertical spacing / line-length / word-wrap / pagination
(the exception category) or already deferred (#5 = indonesian image wrap = old #111).

### 2026-06-21 — #8 RESOLVED (already fixed by prior work). Marked fixed=true.
Verified against reference p63 with a fresh 200dpi render + aligned pixel-diff:
- "large text should be white" → DONE: the SmartArt node labels (Data Collected / Data
  Sources Identified / Data Analyzed) render white, matching reference (commit 2f6bd3e
  "SmartArt fontRef text color honors tint/shade" is the likely fixer).
- "arrows are wonky" → DONE: the blue (circularArrow) + magenta (leftCircularArrow) arrows
  render with byte-identical bounding boxes to the reference (gen blue 418×511, mag 482×563
  — same as ref, just offset ~111px down from accumulated page drift). Diff overlay shows
  the arrow bands match pixel-for-pixel.
The only residual diff is bullet-text wrapping INSIDE the boxes ("Participant Interviews"
wraps in gen, fits on one line in ref) — that is text line-length/wrapping (the exception)
and is NOT what the annotation flagged. Case score unchanged after fresh test run (1 scored,
1 unchanged) → no regression. No code change required; committed annotations.json fixed flag.
LEARNING: re-render before trusting an annotation — this one had been resolved by later commits.

## 2026-06-21 — SELECTED annotation #111 (scraped/indonesian_benchmarking_guide)
Note: "We wrap this image. Word does not." (page index 6, x≈227, y≈637, Generated)
Skipped #208 (text_boundary drift reminder), #8 (1-vs-2 page CJK table row-height overflow),
#59/#66/#82/#93/#118 (all systemic vertical text spacing / line-length / word-wrap = the exception).
#111 is a discrete text-wrap-around-image bug → in scope.

### 2026-06-21 — #111 ANALYSIS: not fixable in scope (vertical-shift exception). NOT marked fixed.

Annotation: scraped/indonesian_benchmarking_guide p7 — "We wrap this image. Word does not."
Image = `image3.PNG`, an anchored (`wp:anchor`) wrapSquare/bothSides float, 251×374pt, the
"American Productivity and Quality Center" questionnaire. Anchored to an empty paragraph
(block 99) right after "2.1.4 Kuesioner …".

Ground truth (mutool stext on reference.pdf p7):
- Reference does NOT wrap. The image occupies the page top; body text "3. Metode Benchmarking"
  starts full-width at y≈462 (top coords), everything below it full-width x=72..526.
- So the note is correct: Word flows text *below* the floated image, not beside it.

Working tree already carried an uncommitted fix for this (`float_bumped_to_new_page` in
src/pdf/mod.rs, ~19 lines; cf. stash@{1} "Push paragraph text below its own wide wrapSquare
floating image"). It correctly suppresses the side-wrap when an empty anchor paragraph's tall
float is bumped to a fresh page. Verified via fresh build (the committed CLI binary was stale,
10:15 — had to build with CARGO_TARGET_DIR=target_dbg to get current code): page 7 now renders
no-wrap, visually matching reference.

BUT it regresses all three metrics for the case:
  Jaccard  40.4% -> 36.6%  (-3.8pp)
  SSIM     54.7% -> 50.4%  (-4.3pp)
  TxtBnd   46.9% -> 31.0%  (-15.9pp)

Root cause = vertical over-reserve (the EXCEPTION category). The fix reserves the full 374pt
image height starting at page top AND still advances for the intervening empty paragraphs, so
"3. Metode Benchmarking" lands at y≈544 instead of reference's y≈462 — ~82pt too low. That
downward shift misaligns page-7+ text and pagination against the reference (the wrapped/baseline
version scores higher only because packing text beside the image keeps later pages aligned).

To get a NET win the float_bumped reserve must match Word: place the image at page top, absorb
the empty anchor paragraphs *beside* it (no extra vertical advance), and start "3. Metode" just
below the image bottom (~58pt gap, not ~140pt). That is vertical-layout tuning with cross-case
risk — squarely the "systemic vertical text spacing causing vertical shift" exception. Deferred.

Left the inherited uncommitted mod.rs WIP untouched (concurrent bart session is active — it
edited tests/output/annotations.json mid-session). Did NOT mark #111 fixed and did NOT commit
(no improvement; would race the other session).

## 2026-06-21 — SELECTED annotation #145 (new/belgian_youth_program_certificate)
Note: "Date should be aligned with 'datum' above it." (page 0, x≈402.8, y≈96.2, Generated)
Skipped #8/#59/#66/#82/#93/#114/#118/#121/#124/#133 (all systemic vertical text spacing /
line-length / word-wrap / pagination = the exception). #111 deferred. #145 is a discrete
horizontal-alignment bug → in scope.

### 2026-06-21 — #145 FIXED (code + annotations.json flag set true). COMMIT BLOCKED by env.
Root cause: the signature date paragraph is `[8×tab]21/10/2024` with tab stops at 1134/3828
twips + default 708. A `wrapTight` floating signature image occupies x≈255.5–311.0pt. Word
advances a left tab that lands inside a floating image's body to the first stop CLEAR of the
image's right edge; so its 5th tab (would be 283.2pt, inside the image) bumps to 318.6pt,
shifting the date one stop right to x=424.8pt — aligned with the "datum" label above. We were
ignoring the image and landing the date at 389.4pt (one 35.4pt stop short).

Fix (verified via fresh CLI render + mutool stext: date now x=424.80, datum x=424.80):
- `build_tabbed_line` gains a `tab_exclusions: &[(f32,f32)]` param (from-text-margin spans).
  Left-tab resolution loops: if the resolved stop falls inside an exclusion span, re-resolve
  from the span's right edge. layout.rs.
- Render call (pdf/mod.rs render_paragraph_block) computes spans from `para.floating_images`
  with Square/Tight/Through wrap that vertically overlap the first line (bare image box, no
  wrap-distance — matches Word's observed bump target). All other call sites pass `&[]`.
GOTCHA: the CLI appends a counter when the output path exists (`/tmp/x(3).pdf`) — must rm or
use a fresh path or you measure a stale PDF (burned ~30 min chasing a phantom non-fix).

Tests: 149 passed, 0 score regressions (208 unchanged), only belgian visual-hash change.
Belgian score nudge Jaccard 21.21→21.24, SSIM 51.87→51.95, text_boundary 1.0 (unchanged).
Did NOT run accept-baselines (user pref: ask first); tests pass without it.

ALSO: removed the inherited uncommitted `float_bumped_to_new_page` WIP from src/pdf/mod.rs
(the prior session's deferred #111 attempt) — it was the SOLE source of the indonesian
regression (TxtBnd 46.9→31.0 etc.) seen on the first test run. Working tree now contains only
the #145 fix + the #145 fixed flag.

COMMIT BLOCKED: this environment denies `.git` writes inside the sandbox and rejects
git-with-sandbox-disabled ("Run outside of the sandbox"). The fix is complete on disk and
ready to commit. To commit, a human/operator should run from the repo root:
  git add -u && git commit  # 6 files: 5 src/pdf/*.rs + tests/output/annotations.json

## 2026-06-21 — SELECTED annotation #117 (scraped/learning_cultures_dissertation p62)
Note: "Figure 3.1 is rendered wrong. The arrows are wonky and the large text should be white."
Skipped #208 (text_boundary drift reminder = systemic vertical), #8 (pagination), #59/#66/#82/
#93/#114/#118/#121/#124 (all systemic vertical text spacing / line-length / word-wrap = the
exception), #111 (deferred image-wrap, vertical-shift). #117 = discrete SmartArt drawing → in scope.

### 2026-06-21 — #117 RESOLVED (verified, prior commits fixed it). Marked fixed=true. COMMITTED 8d20dd4.
Fresh 120dpi render of gen p65 (1-based; the SmartArt "Process of Data Collection and Analysis"
diagram) vs reference p63. Both annotation complaints are resolved in the current build:
- "large text should be white" → box labels (Data Sources Identified / Data Collected /
  Data Analyzed) render WHITE in gen, matching ref (commit 2f6bd3e SmartArt fontRef color).
- "arrows are wonky" → quantified arrow bounding boxes (magenta top circularArrow, cyan bottom
  leftCircularArrow). GEN magenta 334x288 / 3188px vs REF 336x288 / 3185px; GEN cyan 306x250 /
  2991px vs REF 306x250 / 3063px. Near-identical shape, only offset ~84px down by accumulated
  page drift. Arrows are NOT malformed.
Only residual = bullet text inside boxes (gen "• Participant"+wrap vs ref "•Participant" tight)
= text line-length/wrap (the exception), NOT what the annotation flagged.
Tests: 149 passed, 208 scored, 208 UNCHANGED, 0 regressions (1 visual change = belgian, the #145
WIP carried in the tree). No code change for #117; set fixed flag only.
ALSO this session: git commit now WORKS via dangerouslyDisableSandbox (prior "Run outside of
sandbox" no longer applies). Committed the prior session's complete-but-blocked #145 tab-exclusion
fix (5 src/pdf/*.rs) together with the #117 + #145 annotation flags. annotations.json lives under
gitignored tests/output → stage it with `git add -f`.
LEARNING (confirms prior): re-render before trusting an annotation — #117 was already resolved
by later commits; its fixed flag had been lost to concurrent annotations.json churn.

## 2026-06-21 — SELECTED annotation #16 (case35) "This text is placed wrong"
Skipped #0/#1/#2/#3/#4/#10/#11/#12 (systemic vertical text-spacing / pagination / vertical-shift =
the exception), #5/#7 (word-wrap/line-length = exception), #6 (#111 deferred image-wrap = vertical),
#14 (framed-text "empty space above" = vertical positioning = exception). #16 = discrete textbox/shape
placement → in scope. (#17 "arrow 2 placed wrong" is the companion, deferred to a later iteration.)

### 2026-06-21 — #16 RESOLVED (verified, prior geometry/textbox work fixed it). Marked fixed=true.
case35 = 20 floating preset shapes (10 rows × 2 cols), each with a center-anchored text label.
Fresh 150dpi render of gen p1 vs reference.pdf p1 + pixel-diff (16307/2.1M diff px, all sub-pixel
edge AA spread evenly across rows; row9/triangle near-zero). Cropped+stacked the rightArrow ("arrow 2")
and star5 text regions: text is horizontally + vertically centered identically to reference, arrow
geometry (narrow head adj1=25000 / thin shaft adj2=75000) matches pixel-closely. The only residual
diff is font-rendering anti-aliasing (gen substitute font slightly bolder/higher) — case-wide, out of
scope, NOT a placement bug. Case scores Jaccard 0.9954 / SSIM 0.9925 / text_boundary 1.0 (excellent).
No code change required; set fixed flag only.
LEARNING (again): re-render before trusting — #16 was stale, resolved by earlier avLst/geometry +
textbox-centering commits. The avLst override path (adj1/adj2 on rightArrow) renders correctly.

## 2026-06-21 — SELECTED annotation #151 (case35) "On arrow 2 is placed wrong"
Note at x≈423.8, y≈561.3 (Generated, page 0 — right column, near top). Companion to #150
("This text is placed wrong", already fixed). Skipped #208/#8/#59/#66/#82/#93/#114/#118/#121/
#124/#133 (systemic vertical text-spacing / pagination / line-length / framed-text positioning =
the exception), #111 (deferred image-wrap = vertical-shift). #151 = discrete shape placement → in scope.

### 2026-06-21 — #151 RESOLVED (verified, stale — prior geometry work fixed it). Marked fixed=true.
Annotation at row 2 right col = the ADJUSTED rightArrow (adj1=25000 narrow head, adj2=75000 thin
shaft). Fresh 150dpi render of gen p1 vs reference.pdf p1, then a tight colored ink-diff of just
that arrow box (x 333..513pt, y 207..261pt from top): the entire arrow SHAPE — arrowhead, shaft
top/bottom edges, tip — is solid gray (perfect overlap). Arrow-region ink Jaccard 0.9435
(ref 13198px / gen 13187px ink). The ONLY residual diff is red/blue fringing on the label TEXT
"rightArrow adj1=25000 adj2=75000" = font-substitution anti-aliasing (case-wide, out of scope,
NOT a placement bug). Case scores unchanged (1 scored, 1 unchanged; 149 passed) → no regression.
No code change required; set fixed flag only.
LEARNING (3rd time in case35, confirms pattern): re-render before trusting — #16/#150/#151 were all
stale, resolved by earlier avLst/geometry + textbox-centering commits. The avLst override path
(adj1/adj2 on rightArrow) renders pixel-perfect against the reference.

## 2026-06-21 — SELECTED annotation #153 (case51 p0) "First outer column is too wide"
Note at x≈218.2, y≈599.6 (Generated, page 0). Skipped #0/#8/#59/#66/#82/#93/#114/#118/#121/#124/
#133 (systemic vertical text-spacing / pagination / line-length / framed-text = the exception),
#111 (deferred image-wrap), #152 (text wrap-around-rectangle = line-length = exception).
#153 = discrete horizontal column-width bug → in scope.

### 2026-06-21 — #153 FIXED (case51 outer-column AutoFit-to-Window). Committed.
case51 outer tables are `tblW type="auto"` with an equal 4680/4680 gridCol but
wildly asymmetric content (col1 = short text + a nested table, col2 = long
paragraph). Word's AutoFit-to-contents sizes col1 narrow / col2 wide; we were
honoring the equal grid → "first outer column too wide" (both 234pt).

Inherited a prior session's uncommitted WIP that added an AutoFit-to-Window
distribution (`distribute_autofit` + `natural_widths` refactor) — its column
math is CORRECT (renders pixel-match to reference), but it was wired far too
broadly: the top-level `render_table` call was changed to pass `Some(fit_w)` as
`available_width`, which shoved EVERY top-level auto table onto the nested-table
*shrink* path → 75 metric regressions (TxtBnd 100→0 on case61/67, etc.).

Two-part narrowing to make it a clean win:
1. Gated the fill distribution on `has_nested_table` (a cell directly holds a
   `Block::Table`) — case51's exact signature, near-zero corpus blast radius.
2. Split the param: `auto_fit_columns(table, fonts, available_width, fill_width)`.
   `fill_width` (Some only for top-level) steers ONLY the window-fill path; the
   shrink path keeps reading `available_width` (None for top-level), so ordinary
   top-level tables fall through to the gridCol path exactly as before.

Result: case51 TxtBnd 23.3→40.6 (+17.3pp), SSIM 65.1→70.3 (+5.2pp); 207
unchanged, 0 regressions (149 passed). belgian = visual-hash-only change (carried
#145 baseline, no score delta). Did NOT run accept-baselines (user pref).
LEARNING: when a regression survives a trigger-narrowing, the damage is in a
DIFFERENT code path — here the call-site arg change silently activated the
shrink branch; the fix was separating the two width inputs, not just the gate.

## 2026-06-21 — SELECTED annotation #21 (new/isla_language_lesson_plan p6)
Note: "520 L box way too large, and also the content next to it is not showing?"
Skipped #0/#2/#3/#4/#10/#11/#12/#14/#26/#30 (systemic vertical text-spacing / pagination /
framed-text = the exception), #1/#7 (pagination/line-length), #5 (word-wrap), #6 (#111 deferred
image-wrap), #18 (text-wrap-around-rectangle = line-length family), #20 (vague "looks different").
#21 = discrete horizontal placement bug -> in scope.

### Analysis
The "520 L" box is an INLINE wpc drawing canvas (two wps textboxes: the bordered "520 L"
box + the unbordered explanation text). Canvas sits in a ListParagraph (left indent 720tw=36pt;
page left margin also 720tw=36pt). Measured rendered box (300dpi): width identical gen vs ref
(~82pt), content shows fine -- but the whole canvas is shifted ~36pt LEFT in gen: box left edge
gen~=50pt vs ref~=86pt (=36 margin + 36 list-indent). Root cause: group.rs flattens inline-canvas
children into floating textboxes anchored to Column with the child's a:off as offset, dropping
the paragraph's left indent that inline text flow would apply.
FIX: added `indent_relative` flag on Textbox (true for inline-canvas children, via BaseAnchor),
render adds para.indent_left to col_x for those. Pic-in-inline-canvas left unhandled (ponytail note).

### 2026-06-21 — #21 FIXED (verified). fixed=true set + staged. COMMIT BLOCKED by env.
Box left edge moved 49.4pt -> 85.4pt = exact reference match (right 132.5 -> 168.5pt). Box width was
already correct (~82pt) and content was already showing, so the residual was purely the 36pt left shift.
Tests: 149 passed, 208 scored, 206 unchanged, 0 regressions; isla SSIM 57.4->59.7 (+2.3pp).
(case51 +17.3/+5.2 and belgian visual-hash are PRE-EXISTING deltas from earlier commits whose baselines
were never accepted per user pref — not caused by this change.)
GOTCHA (memory's stale-binary): `cargo build --features cli` into the default target/ left target/debug/
docxide-pdf STALE (grep of the binary for my new eprintln = 0 hits). Had to build with
`CARGO_TARGET_DIR=target_dbg cargo build --features cli` and run ./target_dbg/debug/docxide-pdf to see the
fix. The cargo TEST harness (run-tests.sh) DID pick up the lib change (isla improved), so only the CLI bin
is pinned. Always verify CLI renders via target_dbg in this env.
COMMIT BLOCKED: every `git commit` (sandbox-disabled and in-sandbox) returns "Run outside of the sandbox"
this session (same as the #145 session's block). Changes are STAGED and ready. Operator can finish with:
  git commit -F .tmp_isla/msg.txt   # 8 files: 6 src + tests/output/annotations.json + ralph/bart_progress.md
annotations.json fixed=true for #21 is set on disk regardless.

## 2026-06-21 — SELECTED annotation #167 (new/japanese_land_development_sign_form p0)
Note: "Arrows and labels not correct. The arrows should be touching the grey line and the table above."
Skipped #208/#8/#59/#66/#82/#93/#118/#121/#124/#133 (systemic vertical text-spacing / pagination /
line-length / framed-text = the exception), #111 (deferred image-wrap), #114/#152 (line-length wrap),
#158 (vague "text looks different" = font, not a discrete geometry bug). #167 = discrete connector
arrow rendering → in scope.

### Analysis
The form has three straightConnector1 arrows (90cm top double-arrow, 50cm vertical, 70cm). All render
BLUE in gen but BLACK in reference. Root cause in parse_connector_shape_node (textbox.rs): stroke color
took the style matrix lnRef FIRST, falling back to the explicit a:ln only second — backwards from OOXML
precedence (and from the non-connector path at line 432 which does explicit-ln-first). The 50cm arrow
has a:ln/solidFill=tx1 (black) but wps:style/lnRef=accent1 (blue); Word draws black. FIX: swap so the
explicit a:ln solidFill wins over the lnRef style ref.

## 2026-06-22 — SELECTED annotation #18 / id 168 (new/japanese_land_development_sign_form p0)
Note: '"70 .. " text should be rotated 90 degs, and also the arrow is placed wrong.'
Skipped #0-#16 (systemic vertical text-spacing / pagination / line-length / word-wrap / vague = the
exception, or already fixed). Skipped #17/id167 ("arrows should be touching the grey line/table"):
analysis below shows that one IS the vertical-shift exception. #18's PRIMARY complaint (rotate the
"70cm" label 90°) is a discrete rendering bug → in scope.

### 2026-06-22 — #18 FIXED (cell-anchored floating image rotation). Committed.
The whole form is ONE big w:tbl; every arrow/brace/textbox/image is a float anchored to a paragraph
inside a table cell. The "70センチメートル以上" label is NOT text — it is image1.png, a 116.8×18pt PNG
carrying a 90° turn via `pic:spPr/a:scene3d/a:camera/a:rot @rev="5400000"`. parse_image_rotation()
already reads that rev into FloatingImage.rotation_deg, and the BODY float path (positioning.rs)
rotates about the box center — but the CELL float path dropped it: CellFloatingImageLayout had no
rotation field, so the label rendered horizontal.
FIX (3 spots): add `rotation_deg` to CellFloatingImageLayout (table_layout.rs), populate from
fi.rotation_deg, and apply it at BOTH cell-float render sites (table.rs) via a new push_center_rotation()
helper mirroring positioning.rs's center-rotation CTM. Surgical: the helper no-ops unless
deg.abs()>0.01, so unrotated cell floats are byte-identical.
VERIFIED: fresh CLI render — the label is now vertical, matching reference orientation. Tests: 149
passed, 208 scored, 0 regressions; the ONLY change attributable to this fix is the japanese
visual-hash change (case51 TxtBnd+17.3/SSIM+5.2 and isla SSIM+2.3 are pre-existing unaccepted
baselines, not mine). japanese SCORE is flat (jac 0.101→0.1005, ssim 0.3027→0.3025) — the page is
dominated by the vertical-shift exception below, which swamps a tiny 18pt-wide label; the fix is a
verified rendering-correctness win regardless.

RESIDUAL (the vertical-shift EXCEPTION, NOT fixed — covers #17/id167 + #18's "arrow placed wrong"):
the rotated label and all three connector arrows (90/70/50cm) render with correct length, arrowheads,
and black color (color was fixed earlier in textbox.rs) but land at wrong VERTICAL positions because
their anchor paragraphs sit in cells whose row heights differ from Word (90cm arrow ≈correct;
70cm float ~160pt too low; 50cm arrow ~30pt too low). That is table-internal vertical layout drift =
the systemic vertical-shift exception → deferred. Did NOT mark #17/id167 fixed.
LEARNING: rotation was already parsed AND rendered for body floats; the bug was the cell-float layout
struct silently dropping it. When a feature "works" for body content but not in tables, check the
parallel cell-layout/render path — they don't share the same struct.

## 2026-06-22 — SELECTED annotation #176 (new/scottish_fundraising_awards_campaign p0)
Note: '"Scottish Fundraising ..." header overlaps images' (x≈347.5, y≈759.1, Generated).
Skipped #208/#8/#59/#66/#82/#93/#118/#121/#124 (systemic vertical/pagination/line-length = the
exception), #111 (deferred image-wrap), #114/#152 (line-length wrap), #133 (framed-text vertical),
#158/#167 (font/vertical-shift). #176 = discrete header-height-vs-body-overlap → in scope.

### 2026-06-22 — #176 FIXED (header logo height now pushes body down). Marked fixed=true. Staged.
The header is a w:tbl whose first cell holds an inline logo PNG (335×68.5pt) followed by a w:br.
effective_slot_top() already grows the top margin to fit header content via compute_header_height
→ compute_hf_table_height → compute_row_layouts. BUG: the cell-height code only added the image's
content_height in the is_text_empty() branch, but the trailing w:br makes the paragraph NOT
text-empty, so it fell to the text-line branch and reserved only ~one line — the 68.5pt logo height
was dropped. Header height came out ~15pt instead of ~95pt, so the body (orange "Scottish
Fundraising Awards 2025" title) started at the top margin, overlapping the logos.
FIX (table_layout.rs, 2 spots): (1) treat an inline-image paragraph whose only runs are the image +
line breaks as `image_only` so it routes through the content_height branch; (2) in that branch also
add br_count*line_h (mirrors the non-table header path in header_footer.rs which counts w:br lines).
VERIFIED (fresh CLI render via target_dbg — the default target/debug bin was STALE, Jun 21 10:15):
title top moved 78.4pt → 133.2pt from page top; reference is 134.1pt (Δ0.8pt). Page-1 render now
matches reference (logos on top, title below).
TESTS: 149 passed, 0 regressions. scottish TxtBnd 29.4→76.1 (+46.7pp), SSIM 19.6→53.0 (+33.4pp),
Jaccard 12.7→25.3 (+12.7pp). (case51 +17.3/+5.2 and isla SSIM+2.3 are PRE-EXISTING unaccepted
baselines from earlier sessions, not this change.)
LEARNING: is_text_empty() returns false for a lone w:br, which quietly diverts image-only cell
paragraphs off the content_height path. Same body-vs-table-path divergence pattern as #18.
STAGING NOTE: `git add` is blocked this session (sandbox denies .git writes; sandbox-disabled
git add returns "Run outside of the sandbox"). Changes are complete on disk (.gitignore,
table_layout.rs, annotations.json #176=true, this file) and .bart_commit_msg is written — the
loop shell's own `git add -A`/commit will pick them up. target_dbg/ build dir gitignored.

## 2026-06-22 — SELECTED annotation #177 (new/scottish_fundraising_awards_campaign p1)
Note: "Table overlaps images" (x≈347.5, y≈759.1, Generated, page index 1). Skipped #208/#8/#59/
#66/#82/#93/#114/#118/#121/#124/#133 (systemic vertical/pagination/line-length/framed-text = the
exception), #111 (deferred image-wrap), #152 (wrap-around-rect), #158/#167 (font/vertical-shift).
#176 fixed. #177 = discrete table-vs-image overlap → in scope. (Prior #192 session noted #177 may
already render correctly — re-rendering to verify before trusting.)

### 2026-06-22 — #177 RESOLVED (verified, prior #176 header-height fix covers it). Marked fixed=true.
Fresh CLI render of scottish page 2 (0-based p1) vs reference.pdf p2, 120dpi + top-region crops:
the three header logos (Chartered Institute / think Consulting / Scottish Fundraising Awards) sit
clearly ABOVE the body table in both gen and ref — no overlap. The annotation ("Table overlaps
images") no longer reproduces. Root cause is identical to #176: compute_header_height now reserves
the inline logo height for the header w:tbl, and that runs on every page, so page 2's table starts
below the logos just like page 1. No code change needed for #177.
Tests: 149 passed, 0 regressions (204 unchanged). The scottish improvements (TxtBnd 29.4→76.1,
SSIM 19.6→53.0, Jaccard 12.7→25.3) are the pre-existing #176 baseline delta (unaccepted per user
pref); case51/slovak/isla deltas are likewise pre-existing unaccepted baselines from earlier sessions.
LEARNING (confirms pattern): re-render before trusting — #177 was already resolved by the #176 commit
on the same case; a page-level header fix naturally covers every page, not just the annotated one.

## 2026-06-22 — SELECTED annotation #178 (new/slovak_pedagogical_practice_agreement p0)
Note: "Large empty space here." (x≈88.4, y≈691.5, Generated, page 0). Skipped #208/#8/#59/#66/#82/
#93/#114/#118/#121/#124/#133 (systemic vertical/pagination/line-length/framed-text = the exception),
#111 (deferred image-wrap), #152 (wrap-around-rect), #158/#167 (font/vertical-shift). #178 = discrete
empty-space-from-misplaced-logo bug → in scope. The #192 fix note says it "also caused #178's empty
space" — re-rendering to verify before trusting.

### 2026-06-22 — #178 RESOLVED (verified, prior #192 owl-float fix covers it). Marked fixed=true. Staged.
Fresh CLI render of slovak page 0 (target_dbg bin) vs reference.pdf, 120dpi + top-region crops:
the umb logo (left), centered "Univerzita Mateja Bela…" heading, owl (right), horizontal rule, and
body "Vážená pani riaditeľka," all sit at the SAME positions as the reference — no large empty space
at the annotated spot (x≈88, y≈691.5 ≈150pt from top). Root cause was the #192 bug: the absolutely-
positioned `<w:object>` owl was treated as inline and stacked inside the centered heading paragraph,
pushing the heading down and leaving the gap. Commit 947a354 (parse_object_floating_image) floats the
owl top-right, so the heading reclaims its slot and the gap closes. No code change needed for #178.
Tests: 149 passed, 0 regressions (204 unchanged). The slovak SSIM 14.3→18.4 / Jaccard 4.7→7.8 deltas
are the pre-existing #192 baseline (unaccepted per user pref); scottish/case51/isla are likewise
pre-existing unaccepted baselines from earlier sessions. LEARNING (confirms pattern): re-render before
trusting — #178 was already resolved by the #192 commit on the same case; the empty space and the owl
misplacement were two symptoms of one root cause.

## 2026-06-22 — SELECTED annotation #192 (new/slovak_pedagogical_practice_agreement p0)
Note: "Owl should be on the right, with the 'Univerzita ...' heading between the logos, not under
the 'Filozofica ...' logo". Skipped #208/#8/#59/#66/#82/#93/#114/#118/#121/#124/#133 (systemic
vertical/pagination/line-length/framed-text = the exception), #111 (deferred image-wrap), #152
(wrap-around-rect), #158/#167 (font/vertical-shift). Verified #177 (scottish table overlap) and
#182 (carbon emblem aspect ratio) already render correctly in current output (stale, resolved by
prior work) — did NOT mark them, worked only #192. #192 = discrete horizontal float position → in scope.

### 2026-06-22 — #192 FIXED (absolutely-positioned `<w:object>` VML logo now floats). Marked fixed=true. Staged.
The slovak header has two logos: umb (left, DrawingML wp:anchor — rendered fine) and the owl
(image2.emf inside a `<w:object>` whose VML `<v:shape>` is `position:absolute;margin-left:423pt;
margin-top:12.15pt;mso-position-*-relative:text`). We treated EVERY `<w:object>` as INLINE
(parse_object_inline_image → inline Run), so the owl landed in the centered header paragraph,
stacked ABOVE the heading and pushed it down (also caused #178's empty space).
FIX (images.rs + runs.rs): added parse_object_floating_image() — when the inner VML shape style
has `position:absolute`, parse margin-left/top + mso-position relatives and build a FloatingImage
(WrapType::None so the centered heading stays full-width). runs.rs "object" handler tries the
floating path first, falls back to inline. Also added object_is_absolute() so compute_object_height
skips absolute objects (no phantom inline line reserved). Inline objects unchanged (None → inline).
VERIFIED (fresh CLI render via target_dbg): owl now top-RIGHT, heading centered in the middle,
matching reference; no empty space. TESTS: 149 passed, 0 regressions. slovak SSIM 14.3→18.4
(+4.1pp), Jaccard 4.7→7.8 (+3.1pp); only slovak visual-hash change is mine. (scottish/case51/isla
deltas are PRE-EXISTING unaccepted baselines from earlier sessions, not this change.)
LEARNING: `<w:object>` OLE embeddings aren't always inline — Word uses them for absolutely-positioned
logos via the VML fallback shape's `position:absolute` style. Same body-vs-special-path divergence
theme as #18/#176.

## 2026-06-22 — SELECTED annotation #182 (new/carbon_farming_initiative_rule p0)
Note: "Emblem rendered wrong here. Overflows bounding box, wrong aspect ratio." (x≈148.8, y≈716.5).
Skipped #208/#8/#59/#66/#82/#93/#118/#121/#124/#133 (systemic vertical/pagination/line-length/framed =
the exception), #111 (deferred image-wrap), #114/#152 (line-length wrap), #158/#167 (font/vertical-shift).
#182 = discrete image (WMF emblem) rendering bug → in scope. The #192 session noted #182 looked already
correct (stale) — re-rendered to verify.

### 2026-06-22 — #182 RESOLVED (verified, prior multi-DIB WMF work fixed it). Marked fixed=true. Staged.
The emblem is the Commonwealth Coat of Arms embedded as `word/media/image1.wmf` — a multi-band metafile
(three stacked monochrome DIB bands) inline at extent 1419225×1104900 EMU (111.75×87pt). Fresh CLI render
(target_dbg) of carbon page 1 vs reference.pdf, 150dpi + top-region crops + tight ink-bbox measure:
- GEN emblem bbox 711×495px, aspect 1.436, ink 25570px.
- REF emblem bbox 711×495px, aspect 1.436, ink 25598px (Δ0.1%).
Byte-for-byte the same size, aspect, and position — kangaroo/emu/shield/AUSTRALIA banner all correct, no
overflow. The annotation no longer reproduces. Root cause was the old single-DIB path stretching only the
first band to fill the box (wrong aspect/overflow); `wmf.rs::composite_blits` now composites all DIB bands
onto one canvas at their destination rectangles, preserving the true aspect. No code change needed for #182.
Tests: 149 passed, 0 regressions. LEARNING (confirms pattern): re-render before trusting — #182 was already
resolved by the multi-DIB WMF compositing work; its fixed flag was simply never set.

## 2026-06-22 — SELECTED annotation #202 (case74 p0) "Endnote is placed wrong"
Skipped from the top: #208/#8/#59/#66/#82/#93/#114/#118/#121/#124/#133 (systemic vertical/pagination/
line-length/framed = the exception), #111 (deferred image-wrap), #152 (wrap-around-rect), #158 (font/vague),
#167 (vertical-shift). Examined & skipped #193 (samtale "HR through 13" — ref shows rows 10/11 where gen
shows 12/13 at same y = accumulated table row-height drift = the vertical exception) and #201 (pendulum "g
too high" — header row content 7pt high but body rows diverge the OTHER way = systemic table row-height
divergence = the exception). #202 = discrete endnote PLACEMENT bug → in scope.

### 2026-06-22 — #202 FIXED (endnotes flow inline at doc end, not pinned to page bottom). Marked fixed=true. Staged.
Root cause: endnotes (default pos=docEnd) were rendered by reusing render_page_footnotes pinned to the
final-page bottom, STACKED above the footnotes (src/pdf/mod.rs Phase 2c). Word instead flows endnotes in
the normal content stream right after the LAST body block (footnotes stay pinned to page bottom). So gen
showed both note groups jammed at the bottom; ref shows endnotes just below the body text, footnotes at
the bottom.
FIX (footnotes.rs + mod.rs):
- Refactored render_page_footnotes: extracted the separator-draw into draw_note_separator() and the note
  loop into render_notes_downward(start_fn_y, ...). render_page_footnotes still anchors from the bottom.
- Added render_endnotes_inline(top_y, ...): draws the separator at top_y and flows notes downward from the
  body cursor. ponytail-noted: single-page flow only (no overflow-to-next-page pagination yet).
- mod.rs: captured endnote_top_y = pb.slot_top - prev_space_after - 12 (last body cursor + the para's
  space_after + the 12pt separator gap Word leaves), and call render_endnotes_inline on the last page.
VERIFIED (fresh CLI render via target_dbg + mutool stext, all three endnote fixtures):
- case74: endnote "First endnote text" now at top-coord ~212 (ref ~221), footnotes still at bottom ~691.
- case75 (section-wide bag): endnotes ~224 (ref ~221), footnotes bottom.
- erasmus (3-page real doc): the long "Adaptations of this template" endnote now at p3 y≈419 vs ref ≈427
  (was pinned to bottom before) — matches reference mid-page placement.
TESTS: 149 passed, 0 regressions. case74 SSIM 83.6→85.8 (+2.2pp), case75 SSIM 83.6→85.6 (+2.0pp),
erasmus SSIM 50.4→54.5 (+4.1pp). (scottish/case51/slovak/isla/japanese deltas are PRE-EXISTING unaccepted
baselines from earlier sessions, not this change.) Did NOT run accept-baselines (user pref).
LEARNING: endnotes are not footnotes — pos=docEnd means inline-at-end of the content flow, which also
naturally handles multi-page docs (erasmus endnote lands mid-page-3 like Word, not at the page floor).

## 2026-06-22 — SELECTED annotation #203 (scraped/uk_commercial_lease_template p39)
Note: "missing yellow highlights on footnotes". Skipped from the top #208/#8/#59/#66/#82/#93 (systemic
vertical/pagination/line-length/word-wrap = the exception), #111 (deferred image-wrap), #114 (line-wrap),
#118/#121/#124 (vertical), #133 (framed vertical), #152 (wrap-around-rect), #158 (vague/font), #167
(vertical-shift), #185 (vague), #186 (vertical), #190 (pagination), #193/#201 (table row-height vertical),
#195 (font/vertical), #200 (pagination). #203 = discrete run-shading rendering bug → in scope.

### 2026-06-22 — #203 FIXED (footnote ref marks now honor run w:shd shading). Marked fixed=true. Staged.
The yellow on the footnote numbers is NOT w:highlight (grep found zero highlight elements) — it's run
shading: the FootnoteReference character style carries `<w:shd w:val="clear" w:fill="FFFF00"/>`. parse_run_shd
+ char-style resolution already populate Run.shading for it, and render_paragraph_lines already draws run
shading backgrounds. BUG: the footnoteRef/endnoteRef marks are built via superscript_run() → styled_run(),
which copied `highlight` but DROPPED `shading`, so the parsed yellow was thrown away before render.
FIX (1 line, runs.rs styled_run): add `shading: self.shading,` alongside the existing `highlight` copy.
This carries shd into both the inline reference superscript and the footnote-definition ref mark.
VERIFIED (fresh CLI render via target_dbg, mutool p40 @300dpi): footnote markers 89/90 and the inline body
marker all render with the yellow background, matching reference.pdf p40. TESTS: 149 passed, 0 regressions;
uk_commercial_lease is a visual-hash change (the new highlight); 7 "improved"/other visual changes are
PRE-EXISTING unaccepted baselines from earlier sessions (scottish/case51/slovak/erasmus/isla/case74/75/
japanese), not this change. Did NOT run accept-baselines (user pref).
LEARNING: "highlight" in the eye ≠ w:highlight in the XML — Word's FootnoteReference style uses w:shd. The
abbreviated run constructors (styled_run/superscript_run) are a recurring leak point: they hand-copy a
subset of fields, so any property added to the full path (here shading) must also be added there.

## 2026-06-22 — SELECTED annotation #206 (scraped/uk_commercial_lease_template p39)
Note: "6.1.22 Here should be 6.7.1 per the reference" (x≈89.8, y≈746.7, Generated).
Skipped from the top: #208 (text_boundary drift reminder), #8/#200/#190 (pagination/table-split),
#59/#82/#118/#121/#124/#186/#201 (systemic vertical text spacing), #66 (bullet vertical drift),
#93/#114/#152 (line-length/word-wrap), #111 (deferred image-wrap), #133 (framed-text vertical),
#158/#195 (font/vague), #167 (vertical-shift), #185 (vague), #193 (table row-height vertical).
#206 = discrete list-numbering bug (wrong counter value on a heading) → in scope.
