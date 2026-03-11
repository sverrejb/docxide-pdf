# Progress for plan_vaccine_fix.md

## Issue 1: SmartArt Timeline Labels Not Rendering ✅
**Status: Complete**

**Actual Root Cause:** The plan's hypothesis about `sa_font_entry` being None was incorrect — when None, `text_width` uses an approximation (not 0) and `show_text` still renders. The actual root cause was a **text encoding mismatch**: `charts::show_text()` always used `to_winansi_bytes()` (WinAnsi encoding), but the font picked from `seen_fonts` is a CID-encoded TrueType font with `char_to_gid` mapping. The PDF expected glyph IDs but received WinAnsi-encoded bytes, causing all SmartArt text to render as garbage characters.

**Fix Applied:**
1. Added `show_text_encoded()` function in `src/pdf/charts.rs` that accepts an optional `FontEntry` and uses `encode_as_gids()` when the font has `char_to_gid`, falling back to `to_winansi_bytes()` otherwise
2. Updated `src/pdf/smartart.rs` to call `show_text_encoded()` instead of `show_text()`, passing the font entry

**Files Modified:**
- `src/pdf/charts.rs` — added `show_text_encoded()` function
- `src/pdf/smartart.rs` — switched to `show_text_encoded()` for correct text encoding

**Result:** SmartArt text labels now render correctly. All 10 date/description labels appear in the output. Zero regressions across full test suite. Visual comparison scores for this case stayed similar (41.1% Jaccard, 48.5% SSIM) because the 5pt text contributes minimally to pixel-level metrics, but the text is visibly present in the generated PDF.

## Issue 2: Circle Gradient Fills (fillRef idx="3") ✅
**Status: Complete**

**Root Cause:** Confirmed as described in the plan. `parse_style_fill()` ignored the `idx` value and always returned a flat solid color. When `fillRef idx="3"`, it should look up the theme's 3rd fill style from `<a:fillStyleLst>`, which is a `gradFill` with 3 gradient stops using `phClr` (placeholder color) with color transforms (satMod, lumMod, tint, shade).

**Fix Applied:**
1. Added `ColorTransforms`, `ThemeGradientStop`, and `ThemeFillStyle` types in `src/docx/styles.rs`
2. Extended `ThemeFonts` with `fill_styles: Vec<ThemeFillStyle>` field
3. Added `fillStyleLst` parsing in `parse_theme()` — extracts solid fills and gradient fills with per-stop color transforms
4. Added HSL conversion helpers (`rgb_to_hsl`, `hsl_to_rgb`) and `apply_color_transforms()` in `src/docx/textbox.rs` to handle `tint`, `shade`, `satMod`, `lumMod`, `lumOff` transforms
5. Changed `parse_style_fill()` return type from `Option<[u8; 3]>` to `Option<ShapeFill>` — when `idx >= 1`, looks up the corresponding theme fill style and resolves `phClr` to the actual scheme color with per-stop transforms
6. Updated caller in `parse_textbox_from_wsp` to use new return type directly

**Files Modified:**
- `src/docx/styles.rs` — new types, `fillStyleLst` parsing in `parse_theme()`
- `src/docx/textbox.rs` — HSL helpers, `apply_color_transforms()`, updated `parse_style_fill()` and caller

**Result:** Circles now get `ShapeFill::LinearGradient` with 3 stops (light→medium→dark blue at 90° vertical) instead of flat solid accent1. Zero regressions across full test suite. Vaccine case scores: 41.1% Jaccard (-0.1pp, noise), 48.6% SSIM (unchanged).

## Issue 3: "B" Letter Shape (Arc Rendering) ✅
**Status: Complete**

**Actual Root Cause:** Three bugs in `render_arc()`:

1. **Sweep angle not computed**: OOXML arcs sweep from adj1 (start) to adj2 (end). When adj1 > adj2, the sweep wraps around: `swAng = adj2 - adj1 + 360°`. The old code simply used `(start_deg - end_deg)` as the sweep, giving ~87° instead of the correct ~273° for the "B" arcs.

2. **Wrong angle mapping**: The old code added 180° to convert OOXML angles to PDF coordinates (`180 + angle`). The correct mapping is negation (`-angle`). OOXML uses standard trig angles (0°=right, counterclockwise positive) in a y-down display; for PDF y-up, negating all angles preserves the visual appearance.

3. **Bézier alpha too small**: The cubic Bézier approximation used `(step/4).cos()/(step/4).sin()` which computes `tan(step/8)`. The correct formula uses `step/2`, giving `tan(step/4)` — the standard kappa = (4/3)·tan(θ/4). The old formula produced control points ~half as far from the arc endpoints, making curves too flat.

**Fix Applied:**
1. Compute OOXML sweep angle correctly: `sweep = end_deg - start_deg; if sweep <= 0: sweep += 360`
2. Use negation for PDF angle mapping: `math_start = -(start_deg + rotation_deg)`; `total = -sweep_deg`
3. Fixed Bézier alpha: `step/4.0` → `step/2.0`

**Files Modified:**
- `src/pdf/mod.rs` — rewrote `render_arc()` function (lines 333-386)

**Result:** Arc shapes now render with correct sweep angles, positions, and curvature. The "B" letter arcs (273° sweep each) render as large elliptical arcs instead of small ~87° segments. Zero regressions across full test suite. Vaccine case scores: 41.1% Jaccard (unchanged), 48.7% SSIM (+0.1pp).

## Issue 4: Underline Split on Heading ✅
**Status: Complete**

**Root Cause:** Confirmed as described in the plan. Each text chunk with `underline: true` generated its own decoration rectangle at `(x, ul_y, chunk.width, thick)`. When a single underlined run is split into multiple chunks (e.g., per-word for spacing/kerning), each word gets its own underline segment with visible gaps between them.

**Fix Applied:**
1. In `src/pdf/layout.rs`, modified the underline decoration logic to merge consecutive underline rectangles that share the same y-position (within 0.01pt tolerance), thickness, and color
2. When a new underline chunk is adjacent to the previous one, the previous decoration's width is extended to cover `x + chunk.width` instead of creating a new entry
3. This mirrors the existing link annotation merging pattern already used for hyperlinks (lines 822-829)

**Files Modified:**
- `src/pdf/layout.rs` — underline decoration merging in the chunk rendering loop

**Result:** Consecutive underlined text chunks now produce a single continuous underline rectangle instead of per-word segments. Zero regressions across full test suite. Vaccine case scores: 41.1% Jaccard (unchanged), 48.7% SSIM (unchanged).

## Issue 5: Spacing Above "The History of the Vaccine" Heading ✅
**Status: Complete**

**Actual Root Cause:** The plan's hypothesis about miscalculated paragraph height was partially correct but missed the key factor. The first paragraph contains 14 floating anchor shapes, including a `wrapTopAndBottom` gradient rectangle (Rectangle 1, behindDoc=1) and a `wrapSquare` figure caption textbox (Text Box 4). The `wrapTopAndBottom` shape correctly expanded `content_h` to 175pt (posV=-41pt + height=180pt + distB=36pt). However, the `wrapSquare` Text Box 4 (posV=127.7pt, height=144pt, bottom=271.7pt) was NOT being accounted for in the content height calculation. Since `wrapSquare` shapes affect text flow by pushing text around them, body text following the paragraph should not start within the textbox's vertical extent. The heading appeared ~85pt too high because it started at 175pt + 24pt (space_before) = 199pt from margin, while the reference placed it after the wrapSquare zone at ~296pt.

Key discovery during investigation: the "Chapter 1" and subtitle text visible on the gradient are inside Rectangle 1's `wps:txbx/w:txbxContent` (textbox content rendered as part of the shape), NOT separate body paragraphs. The parser correctly produces only 2 blocks for the first page section: the empty shapes paragraph and the Heading1 paragraph.

**Fix Applied:**
1. In `src/pdf/mod.rs`, expanded the textbox content_h loop to also include `WrapType::Square` textboxes (previously only `WrapType::TopAndBottom` was checked)
2. For `wrapSquare` textboxes, the paragraph's content_h is expanded to `max(content_h, v_offset + height + dist_bottom)`, matching the existing TopAndBottom behavior

**Files Modified:**
- `src/pdf/mod.rs` — added `WrapType::Square` to the textbox content_h reservation logic

**Result:** The heading now starts below the wrapSquare textbox's extent, matching the reference position much more closely. Page count now matches reference (5 = 5, was 4 vs 5 before). Zero regressions in other fixtures. Vaccine case scores: 39.4% Jaccard (-1.7pp), 44.0% SSIM (-4.7pp) — slight metric regression due to content shifting ~97pt lower on page 1, but the heading position visually matches the reference much better and the correct page count is a structural improvement.

## Issue 6: Large Gradient Rectangle Color ✅
**Status: Complete — No changes needed**

**Investigation:** Visual comparison of generated vs reference page 1 confirms the gradient rectangle is rendering correctly:
- Direction: top-to-bottom (90° angle from `a:lin ang="5400000"`) — correct
- Start color: blue (#4472C4, accent1 at pos=0) — correct
- End color: orange (#ED7D31, accent2 at pos=100000) — correct
- The middle transition area blends naturally from blue through warm tones to orange, matching the reference

The `parse_gradient_fill()` function in `src/docx/textbox.rs` correctly parses the explicit `a:gradFill` with two stops and the 90° linear angle. The PDF axial shading pattern in `src/pdf/mod.rs` renders it faithfully. The differences visible in the diff image are from layout/positioning (addressed by Issues 1-5), not gradient color rendering.

**Files Modified:** None
**Result:** No code changes required. The gradient was already rendering acceptably as predicted by the plan.
