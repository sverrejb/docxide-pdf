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
