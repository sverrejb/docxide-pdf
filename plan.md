# Fix SmartArt Timeline Rendering

## Context

The `vaccines_history_chapter` fixture contains a SmartArt horizontal timeline (arrow with date circles and text labels). The current output is visually wrong compared to the Word reference:

- **Reference**: Light blue notched right arrow, blue circles with white borders on the arrow, small text labels above/below
- **Generated**: Flat light blue rectangle (instead of arrow), circles with no borders, tiny text barely visible

The SmartArt parsing already extracts shapes from `word/diagrams/drawing1.xml` via `parse_smartart_drawing()` in `images.rs`. The issues are in shape type recognition and rendering.

The diagram's internal EMU coordinates match the display extent (`wp:extent cx=6146800 cy=1496060`), so no scaling is needed.

## Issues (ordered by visual impact)

1. **`notchedRightArrow` rendered as rectangle** — `ShapeType` enum only has `Rect`/`Ellipse`; all other presets fall back to `Rect`
2. **No stroke rendering** — circles should have 1pt white borders (`a:ln w="12700"` with `schemeClr lt1`), but `SmartArtShape` has no stroke fields
3. **Text blocked on filled shapes** — rendering condition `shape.fill.is_none()` prevents text over fills (doesn't affect this fixture since text shapes use `noFill`, but is a general bug)
4. **Text color hardcoded black** — should use parsed color from XML

## Implementation

### Step 1: Model changes (`src/model.rs`)

**1a.** ~~COMPLETED~~ Add `NotchedRightArrow` variant to `ShapeType` enum (line 167-171):
```rust
pub enum ShapeType {
    #[default]
    Rect,
    Ellipse,
    NotchedRightArrow,
}
```

**1b.** ~~COMPLETED~~ Add fields to `SmartArtShape` (line 173-182):
```rust
pub stroke_color: Option<[u8; 3]>,
pub stroke_width: f32,
pub text_color: Option<[u8; 3]>,
```

### Step 2: Parsing changes (`src/docx/images.rs`)

**2a.** ~~COMPLETED~~ Recognize `notchedRightArrow` in `parse_dsp_shape()` (line 464):
```rust
"notchedRightArrow" => ShapeType::NotchedRightArrow,
```

**2b.** ~~COMPLETED~~ Parse stroke from `a:ln` in `parse_dsp_shape()` (after line 448, before text):
- Find `a:ln` child of `sp_pr`
- Check for `a:noFill` child (means no stroke)
- Otherwise call `parse_solid_fill(ln_node, theme)` — reuses existing function since `a:ln > a:solidFill` has the same structure
- Parse width from `a:ln/@w` attribute (EMU / 12700 → points, default 0.75pt)

**2c.** ~~COMPLETED~~ Parse text color in `parse_dsp_text()`:
- Change signature to accept `theme: &ThemeFonts` and return `(String, f32, Option<[u8; 3]>)`
- In the `rPr` handling block, call `parse_solid_fill(rpr, theme)` for text color
- Update call site at line 450

**2d.** ~~COMPLETED~~ Update `SmartArtShape` construction (line 469) with new fields.

### Step 3: Rendering changes (`src/pdf/mod.rs`)

**3a.** ~~COMPLETED~~ Add `NotchedRightArrow` arm to `draw_shape_path()` (line 44-63):

Geometry for default adjust values (adj1=50000, adj2=50000):
- `ss = min(w, h)` (shape scale unit)
- `arrow_dx = ss * 0.5` (arrowhead width)
- `arrow_start = w - arrow_dx`
- `shaft_inset = h * 0.25` (25% from top/bottom)
- `notch_depth = ss * 0.25` (notch V-depth)

8-point polygon clockwise from top-left:
1. `(notch_depth, y + h - shaft_inset)` — top-left of shaft
2. `(arrow_start, y + h - shaft_inset)` — top-right of shaft
3. `(arrow_start, y + h)` — top of arrowhead
4. `(w, y + h/2)` — arrow tip
5. `(arrow_start, y)` — bottom of arrowhead
6. `(arrow_start, y + shaft_inset)` — bottom-right of shaft
7. `(notch_depth, y + shaft_inset)` — bottom-left of shaft
8. `(0, y + h/2)` — notch vertex
Close path.

**3b.** ~~COMPLETED~~ Rewrite SmartArt rendering loop (lines 1473-1521):

Replace the current fill-only + text-without-fill logic with:
- **Fill + stroke combined**: For shapes with both fill and stroke, use `fill_nonzero_and_stroke()` (already used in `charts.rs:634`)
- **Fill only**: Use existing `fill_nonzero()`
- **Stroke only**: Use `stroke()` with `set_stroke_rgb` and `set_line_width` (patterns from `table.rs`, `charts.rs`)
- **Text on any shape**: Remove the `shape.fill.is_none()` guard — render text regardless of fill
- **Text color**: Use `shape.text_color` if present, else black

## Files to modify

| File | Changes |
|------|---------|
| `src/model.rs` | Add `NotchedRightArrow` to `ShapeType`; add `stroke_color`, `stroke_width`, `text_color` to `SmartArtShape` |
| `src/docx/images.rs` | Recognize `notchedRightArrow` preset; parse `a:ln` stroke; parse text color in `parse_dsp_text` |
| `src/pdf/mod.rs` | Add arrow path to `draw_shape_path()`; rewrite SmartArt render loop for fill+stroke+text |

## Existing code to reuse

- `parse_solid_fill()` (`src/docx/textbox.rs:173`) — already imported in `images.rs:12`, works for stroke colors
- `fill_nonzero_and_stroke()` — already used in `src/pdf/charts.rs:634`
- `set_stroke_rgb()`, `set_line_width()`, `stroke()` — used throughout `src/pdf/table.rs`, `charts.rs`
- `draw_shape_path()` — existing shape dispatch, just needs new arm

## Verification

1. `cargo build` — compilation check
2. `DOCXIDE_CASE=vaccines_history_chapter cargo test -- --nocapture` — check improved scores
3. View `tests/output/scraped/vaccines_history_chapter/generated/page_001.png` — arrow shape, circle borders, text visible
4. Compare with `tests/output/scraped/vaccines_history_chapter/reference/page_001.png`
5. `cargo test -- --nocapture` — full suite, check for "REGRESSION in:" lines
