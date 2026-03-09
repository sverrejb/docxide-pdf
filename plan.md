# Plan: Improve vaccines_history_chapter fixture

## Context

The `vaccines_history_chapter` scraped fixture scores **0.0% Jaccard, 0.0% SSIM** (5 ref pages, 4 generated). Page 1 of the reference has a decorative header: a blue-to-orange gradient banner with white "Chapter 1" text, a red subtitle, 3 filled blue circles, connecting lines/arcs, and a SmartArt vaccine timeline. The generated output shows none of this — just plain body text.

### Root causes (confirmed via code + XML analysis)

1. **Gradient fill not parsed → white text invisible (CRITICAL)**: The main textbox ("Chapter 1" + subtitle) is inside a rect with `a:gradFill` (accent1 blue → accent2 orange). `parse_solid_fill()` only handles `a:solidFill`, so `fill_color = None`. The textbox IS returned (has text paragraphs), but Title style sets `w:color FFFFFF` (white). White text on white page = invisible.

2. **Style-based fills not resolved**: The 3 blue circles (ellipses) and other shapes get fill from `wps:style/a:fillRef idx="3" schemeClr="accent1"`, not from explicit `a:solidFill`. The parser doesn't check `wps:style`, so `fill_color = None` and these shapes return None (no text + no fill → dropped).

3. **Non-textbox shapes silently dropped**: `parse_run_drawing()` in `images.rs` only recognizes textboxes, images, and charts. Lines, arcs, and other geometric shapes → None.

4. **SmartArt unimplemented**: The vaccine timeline diagram (`dgm:relIds`) is completely ignored.

5. **Page count mismatch (4 vs 5)**: Header decoration occupies ~250pt on page 1 in reference. Without it, text reflows into fewer pages. Will self-correct as elements are rendered.

---

## Step 1: Parse gradient fills as solid-fill approximation

**Goal**: Make the "Chapter 1" textbox background visible so white text becomes readable.

The gradient in this document:
```xml
<a:gradFill><a:gsLst>
  <a:gs pos="0"><a:schemeClr val="accent1"/></a:gs>   <!-- blue #4472C4 -->
  <a:gs pos="100000"><a:schemeClr val="accent2"/></a:gs> <!-- orange -->
</a:gsLst><a:lin ang="5400000"/></a:gradFill>
```

**`src/docx/textbox.rs`**:
- Add `parse_gradient_fill_as_solid(sp_pr, theme) -> Option<[u8; 3]>` — find `a:gradFill/a:gsLst/a:gs` (first stop), resolve `a:schemeClr` or `a:srgbClr` via existing `resolve_scheme_color`.
- Modify line 238 fill resolution to try gradient after solid:
  ```rust
  let fill_color = sp_pr.and_then(|sp| {
      parse_solid_fill(sp, theme)
          .or_else(|| parse_gradient_fill_as_solid(sp, theme))
  });
  ```

**Expected impact**: +10-15pp Jaccard. Blue rect behind white text = text visible.

**Verify**: `DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture`, check page_001 comparison image.

---

## Step 2: Parse `a:fillRef` style-based fills

**Goal**: Shapes using theme style fills (3 blue circles, filled rects) become visible as colored rectangles.

The ellipses have no explicit fill on `wps:spPr`. Their fill comes from:
```xml
<wps:style>
  <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>
</wps:style>
```

**`src/docx/textbox.rs`**:
- Add `parse_style_fill(wsp, theme) -> Option<[u8; 3]>` — find `wps:style` → `a:fillRef` → resolve `schemeClr` child to RGB.
- Chain into fill resolution after gradient:
  ```rust
  let fill_color = sp_pr.and_then(|sp| {
      parse_solid_fill(sp, theme)
          .or_else(|| parse_gradient_fill_as_solid(sp, theme))
  }).or_else(|| parse_style_fill(wsp, theme));
  ```

This makes the line 259 check (`paragraphs.is_empty() && (has_no_fill || fill_color.is_none())`) pass for filled shapes, so they get returned as textboxes with empty paragraphs + fill color. The rendering code already handles this case (draws fill rect, skips text).

**Expected impact**: +3-5pp Jaccard. Circles appear as blue rectangles at correct positions.

**Verify**: Same test command. Check that 3 blue rectangles appear where the circles should be.

---

## Step 3: True PDF gradient fills

**Goal**: Replace solid-fill approximation with actual linear gradient rendering for visual fidelity.

### Step 3a: Model change — `ShapeFill` enum

**`src/model.rs`**: Add enum and update Textbox:
```rust
pub enum ShapeFill {
    Solid([u8; 3]),
    LinearGradient { stops: Vec<([u8; 3], f32)>, angle_deg: f32 },
}
```
Change `Textbox.fill_color: Option<[u8; 3]>` → `Textbox.fill: Option<ShapeFill>`.

**Touch points** (mechanical migration):
- `src/docx/textbox.rs` — `WspResult.fill_color` → `fill`, both DrawingML and VML parsing
- `src/docx/images.rs` — textbox construction from WspResult
- `src/pdf/mod.rs` — textbox fill rendering
- `src/pdf/header_footer.rs` — header/footer textbox rendering (if separate)

### Step 3b: Full gradient parsing

**`src/docx/textbox.rs`**: Replace `parse_gradient_fill_as_solid` with `parse_gradient_fill(sp_pr, theme) -> Option<ShapeFill>`:
- Extract all stops from `a:gsLst/a:gs` (pos attribute = 0-100000, normalize to 0.0-1.0)
- Extract angle from `a:lin @ang` (in 60000ths of a degree → divide by 60000)
- Return `ShapeFill::LinearGradient { stops, angle_deg }`

### Step 3c: PDF gradient rendering

**`src/pdf/mod.rs`**: In textbox fill rendering (~line 1244):
- `Solid` → existing rect fill behavior
- `LinearGradient` → write PDF shading pattern:
  1. Write `ExponentialFunction` (Type 2): `domain [0,1]`, `c0 = start_color`, `c1 = end_color`, `n = 1.0`
  2. Write `ShadingPattern` (Type 2 axial): `coords` based on angle + bounding box, `function` ref
  3. In content stream: `save_state`, clip to rect, set pattern color space (`/Pattern cs /P1 scn`), `fill`, `restore_state`
  4. Add pattern to page resources `/Pattern` dict

Follow the existing `alpha_gs_refs` pattern: collect gradient specs during rendering, write PDF objects after render loop.

OOXML angle conversion: `ang=5400000` = 90° = top-to-bottom. In PDF coords (origin bottom-left): coords = `(cx, y+h, cx, y)`.

**Expected impact**: +2-3pp Jaccard. Blue-to-orange gradient matches reference closely.

---

## Step 4: Basic ellipse geometry rendering

**Goal**: Render circles as actual circles instead of rectangles.

### Step 4a: Shape type in model

**`src/model.rs`**: Add to `Textbox`:
```rust
pub enum ShapeType { Rect, Ellipse }
// default: Rect
```

### Step 4b: Parse preset geometry

**`src/docx/textbox.rs`**: Parse `wps:spPr/a:prstGeom @prst`. Map `"ellipse"` → `Ellipse`, all others → `Rect`.

### Step 4c: Ellipse rendering

**`src/pdf/mod.rs`**: For `Ellipse` fill, draw 4 cubic Bezier curves (control point factor `0.5522847`) instead of a rectangle. Same fill logic applies (solid or gradient).

**Expected impact**: +1-2pp Jaccard.

---

## NOT in scope

- **SmartArt diagrams** — `word/diagrams/drawing1.xml` has pre-rendered shapes but parsing is a separate project
- **Line/arc stroke rendering** — decorative connecting elements, low ROI
- **Text wrapping** — shapes use `wrapNone`, not needed here

## Verification

```bash
# Run just this fixture
DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture

# Check for regressions
cargo test -- --nocapture 2>&1 | grep "REGRESSION"

# Full score overview
./tools/target/debug/analyze-fixtures
```

After each step: (1) no regressions, (2) visual inspection of `tests/output/scraped/vaccines_history_chapter/comparison/page_001.png`.

## Critical files

- `src/docx/textbox.rs` — fill parsing (lines 156-189, 221-271)
- `src/model.rs` — `Textbox` struct (line 166)
- `src/pdf/mod.rs` — textbox fill rendering (~line 1244)
- `src/docx/images.rs` — `parse_run_drawing` textbox construction (~line 206)
