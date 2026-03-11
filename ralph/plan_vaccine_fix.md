# Plan: Fix vaccines_history_chapter Case

## Current State
- **Fixture**: `tests/fixtures/scraped/vaccines_history_chapter/input.docx`
- **Page 1** has the main visual issues; pages 2-5 are body text and look reasonable.

## Issues Identified (Priority Order)

### Issue 1: SmartArt Timeline Labels Not Rendering ✅ COMPLETED
**Severity: High** — The SmartArt timeline (horizontal process diagram) has 10 text label shapes with dates and descriptions positioned above/below the arrow. None of these labels appear in our output.

**Root Cause Analysis:**
- The SmartArt drawing XML (`word/diagrams/drawing1.xml`) has 20 `<dsp:sp>` shapes:
  - 1 arrow shape (`notchedRightArrow`, solid fill)
  - 10 text label shapes (`rect`, `noFill`, with `<dsp:txBody>` containing `<a:p>/<a:r>/<a:t>` elements)
  - 9 dot shapes (`ellipse`, solid fill, no text)
- Text labels have `sz="500"` → 5pt font, positioned at various y-coordinates above/below the arrow
- The parsing code in `src/docx/smartart.rs:126` correctly keeps shapes where `text.is_empty()` is false
- The rendering code in `src/pdf/smartart.rs:106` should render text when `font_size > 0.0`
- **Suspected cause**: Font entry lookup failure. `render_smartart` picks a font from `seen_fonts` (the already-registered fonts HashMap), but if the SmartArt is rendered before any body text fonts are registered, `seen_fonts` could be empty, making `sa_font_entry` be `None`. When `sa_font_entry` is None, `charts::text_width` returns 0 and `charts::show_text` may skip rendering.

**Fix approach:**
1. Add debug logging / test to confirm whether text shapes are parsed and whether `sa_font_entry` is None at render time
2. Ensure SmartArt text rendering has a fallback font (e.g., the document's default body font) even if `seen_fonts` is empty
3. If font lookup is fine, investigate coordinate mapping (SmartArt internal EMU space vs display extent)
4. Consider whether SmartArt `<dsp:txXfrm>` should be used for text positioning instead of the shape's `<a:xfrm>` (the txXfrm provides a separate transform specifically for text placement)

**Files to modify:**
- `src/pdf/smartart.rs` — font fallback, text rendering logic
- `src/docx/smartart.rs` — potentially parse `txXfrm` separately for text position

---

### Issue 2: Circle Gradient Fills (fillRef idx="3") ✅ COMPLETED
**Severity: High** — The three large circles (Oval 31, etc.) containing the "T", "Y", "B" letters have flat solid blue fills instead of the gradient fills shown in the reference.

**Root Cause:**
- The circles use `wps:style` with `<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>`
- `fillRef idx="3"` means "use the 3rd fill style from the theme's `<a:fillStyleLst>`"
- The theme's 3rd fill style is a `<a:gradFill>` with three gradient stops using `phClr` (placeholder color) with color transforms (satMod, lumMod, tint, shade)
- Our `parse_style_fill()` in `src/docx/textbox.rs:237-269` ignores the `idx` value and just resolves the scheme color as a flat solid color — it never checks the theme fill style list
- Result: circles get solid `accent1` color instead of the accent1-based gradient

**Fix approach:**
1. In `parse_style_fill()` (or a new function), when `idx >= 2`, look up the corresponding fill style from the theme's `fillStyleLst`
2. If the theme fill style is `gradFill`, resolve each gradient stop's `phClr` to the actual scheme color specified in the `fillRef`, then apply color transforms (lumMod, satMod, tint, shade)
3. Return a `ShapeFill::LinearGradient` instead of `ShapeFill::Solid`
4. This requires the `ThemeFonts` struct to also carry the theme fill styles (currently it only carries colors and font names)

**Files to modify:**
- `src/docx/styles.rs` — parse and store theme fill style list in `ThemeFonts`
- `src/docx/textbox.rs` — update `parse_style_fill()` to check `idx` against theme fill styles
- `src/model.rs` — possibly no changes (ShapeFill::LinearGradient already exists)

**Complexity:** Medium-High. Requires parsing the theme's fill style matrix and applying color transforms (lumMod, satMod, tint, shade) to the base scheme color.

---

### Issue 3: "B" Letter Shape (Arc Rendering) ✅ COMPLETED
**Severity: Medium** — The "B" letter in the third circle is formed by two `arc` shapes with adjustment values and rotation, plus a vertical `line` connector. The arcs don't form a proper "B" shape in our output.

**Root Cause:**
- The arcs have adjustment values: `adj1="val 12807064"` `adj2="val 7608109"` (first arc) and `adj1="val 12635688"` `adj2="val 7044883"` (second arc)
- They also have rotation: `rot="367053"` (first) and `rot="1104882"` (second) in 60,000ths of a degree
- Our arc renderer in `src/pdf/mod.rs:333-380` (render_arc) computes the arc path, but the angle calculations may be incorrect for these specific adjustment values
- The OOXML arc preset's adj1/adj2 represent start and end sweep angles in 60,000ths of a degree. adj1=12807064 → 213.5°, adj2=7608109 → 126.8°
- The arc should sweep from ~127° to ~214° (the right side of an ellipse = the bumps of the "B")
- The rotation is applied on top of this
- Possible issue: the arc renderer's angle mapping between OOXML conventions (clockwise, 0° at top/right) and PDF conventions (counterclockwise, 0° at right) may have an error with these specific values

**Fix approach:**
1. Write a test case with the specific arc parameters from this document
2. Verify the angle conversion math in `render_arc`
3. Compare the rendered arc path against the VML fallback path data (available in the document XML) — the VML contains explicit coordinate paths that can serve as ground truth
4. Fix the angle conversion if needed

**Files to modify:**
- `src/pdf/mod.rs` — `render_arc()` function

---

### Issue 4: Underline Split on Heading ✅ COMPLETED
**Severity: Low-Medium** — "The History of the Vaccine" heading has `w:u w:val="single"` underline that appears as segments instead of one continuous line.

**Root Cause:**
- The heading text is a single run: `<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>The History of the Vaccine</w:t></w:r>`
- Our underline rendering draws a thin rectangle per word/segment. If the text layout splits into multiple segments (e.g., due to word spacing or kerning), each segment gets its own underline rectangle with gaps between them
- The reference shows one continuous underline spanning the full text width

**Fix approach:**
1. When rendering underlines, instead of drawing per-segment rectangles, accumulate the total extent of consecutive underlined text and draw a single rectangle from the leftmost x to the rightmost x+width
2. This may require changes in the text layout pass to track the underline span across words

**Files to modify:**
- `src/pdf/layout.rs` — underline rendering logic

---

### Issue 5: Spacing Above "The History of the Vaccine" Heading ✅ COMPLETED
**Severity: Low-Medium** — The vertical distance from the circles/gradient area to the "The History of the Vaccine" heading appears too small in our output compared to the reference.

**Root Cause:**
- The heading uses style `Heading1` which has `w:spacing w:before="480"` (480 twips = 24pt of space before)
- The paragraph above the heading contains many floating anchor shapes. The effective height of that paragraph may not account for the total extent of its floating shapes
- The gradient rectangle is a `wrapTopAndBottom` anchor with large offsets (`positionV: -520700 EMU = -41pt` from margin), so it extends well above the text area
- The circles are `wrapNone` anchors, so they don't affect text flow
- The text box captions have `wrapSquare` which pushes text down
- The total vertical space consumed by the first paragraph (with all its floating shapes) may be miscalculated

**Fix approach:**
1. Investigate how the first paragraph's height is calculated when it contains `wrapTopAndBottom` shapes
2. `wrapTopAndBottom` shapes should push subsequent content below their bottom edge + distB
3. Verify that the gradient rectangle's effective bottom edge is correctly calculated (it's a large rect from margin top -41pt, height=180pt, so bottom at ~139pt from margin top)
4. The heading's `w:before="480"` should be applied relative to the effective bottom of the preceding content

**Files to modify:**
- `src/pdf/mod.rs` — paragraph spacing calculation with floating shapes
- `src/pdf/layout.rs` — paragraph vertical positioning

---

### Issue 6: Large Gradient Rectangle Color (lower priority) ✅ COMPLETED
**Severity: Low** — The large background rectangle at page top uses `a:gradFill` with two stops: `accent1` (blue, pos=0) → `accent2` (orange, pos=100000), linear at 90° angle. The gradient rendering looks close but the middle pink/warm area in the reference might not match.

**Root Cause:**
- This rectangle uses an explicit `a:gradFill` in its spPr (not a theme fillRef), so `parse_gradient_fill()` already parses it
- The gradient stops resolve to the theme colors `accent1` (#4472C4 blue) and `accent2` (#ED7D31 orange)
- The reference shows the gradient transitioning through a pink/warm intermediate color, which is the natural blend of blue → orange
- Our gradient rendering may be close but could differ in how PDF linear gradients interpolate vs how Word renders them

**Fix approach:**
- Lower priority; likely looks acceptable already
- If needed, verify the gradient angle and color stop positions

---

## Suggested Implementation Order

1. **Issue 1 (SmartArt labels)** — Highest visual impact, likely a relatively simple fix (font lookup or coordinate issue)
2. **Issue 2 (Gradient fills)** — Significant visual impact, moderate complexity (theme fill style parsing)
3. **Issue 3 (Arc "B" shape)** — Moderate impact, contained fix (angle math)
4. **Issue 4 (Underline continuity)** — Low-medium impact, benefits other cases too
5. **Issue 5 (Spacing)** — Low-medium impact, complex root cause analysis
6. **Issue 6 (Gradient blending)** — Low impact, likely already acceptable

## Key Files Summary
- `src/docx/smartart.rs` — SmartArt shape parsing (text, txXfrm)
- `src/pdf/smartart.rs` — SmartArt rendering (font lookup, text positioning)
- `src/docx/textbox.rs` — `parse_style_fill()` for theme fillRef → gradient
- `src/docx/styles.rs` — Theme fill style list storage
- `src/pdf/mod.rs` — `render_arc()`, connector rendering, spacing
- `src/pdf/layout.rs` — Underline rendering, paragraph spacing
