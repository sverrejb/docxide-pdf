# Plan: DrawingML Preset Geometry Support

## Context

We currently hardcode 3 shapes (Rect, Ellipse, NotchedRightArrow). All other preset shapes fall back to Rect. OOXML defines **187 preset shapes** via a formula-based geometry system (`presetShapeDefinitions.xml`). Only 11 distinct shapes appear in our test fixtures today, but SmartArt, textboxes, and standalone shapes all need this. Building the full formula interpreter (~17 operators) unlocks all 187 shapes at once and also enables `a:custGeom` support.

## Approach: Full formula interpreter from the start

Hand-coding individual shapes doesn't scale and the hand-coded shapes become dead code once the interpreter exists. The formula language is small (17 operators, ~35 built-in variables) and well-specified.

## New module: `src/geometry/`

Pure-math module shared between parsing (text rect computation) and rendering (path generation).

```
src/geometry/
  mod.rs          -- public API: evaluate_preset(), evaluate_custom()
  formulas.rs     -- guide formula evaluator (17 operators + built-ins)
  definitions.rs  -- generated: 187 preset shape definitions as Rust const data
  path.rs         -- PathCommand types, arcTo→cubic + quadBez→cubic conversion
```

## Model changes (`src/model.rs`)

Replace `ShapeType` enum with:

```rust
pub struct ShapeGeometry {
    pub preset: Option<String>,              // e.g. "roundRect"
    pub adjustments: Vec<(String, i64)>,     // overrides from a:avLst
    pub custom_paths: Option<CustomGeometry>, // for a:custGeom
}
```

Update `SmartArtShape.shape_type` and `Textbox.shape_type` fields to `ShapeGeometry`.

## Geometry engine design

### `path.rs` — Intermediate path representation

```rust
enum PathCommand { MoveTo, LineTo, CubicTo, ArcTo, Close }
enum ResolvedCommand { MoveTo, LineTo, CubicTo, Close }  // arcTo resolved to cubics
```

- Extract existing arc→cubic math from `pdf/mod.rs:267-320`
- Add quadBez→cubic conversion (standard formula)
- OOXML `arcTo(wR, hR, stAng, swAng)` is relative from current point — different from the absolute arc in connectors

### `formulas.rs` — 17-operator evaluator

| Op | Semantics |
|----|-----------|
| `val x` | literal |
| `*/ x y z` | (x*y)/z |
| `+- x y z` | x+y-z |
| `+/ x y z` | (x+y)/z |
| `?: x y z` | if x>0 then y else z |
| `abs`, `sqrt`, `min`, `max` | standard |
| `pin x y z` | clamp y to [x,z] |
| `sin/cos/tan x y` | x * trig(y) where y in 60000ths deg |
| `at2 x y` | atan2(y,x) → 60000ths deg |
| `cat2/sat2 x y z` | compound trig |
| `mod x y z` | sqrt(x²+y²+z²) (vector magnitude) |

Built-in variables computed from (w, h): `w`, `h`, `l`(=0), `t`(=0), `r`(=w), `b`(=h), `ss`, `ls`, `hc`, `vc`, `wd2..wd12`, `hd2..hd10`, `ssd2..ssd32`, `cd2/cd4/cd8/3cd4/3cd8`.

Use i64 for guide computation (matches OOXML integer coordinates), f64 for final resolved paths, f32 only at PDF boundary.

### `definitions.rs` — Generated from spec XML

A code generator tool (`tools/generate-shapes/`) parses `presetShapeDefinitions.xml` (from ECMA-376 Annex D, MIT-licensed via OfficeDev/Open-XML-SDK) and emits Rust const data:

```rust
pub fn lookup(name: &str) -> Option<&'static PresetDef> {
    match name { "rect" => Some(&RECT), "roundRect" => Some(&ROUND_RECT), /* ... */ _ => None }
}
```

### Public API (`mod.rs`)

```rust
pub fn evaluate_preset(name: &str, w: f64, h: f64, adj_overrides: &[(String, i64)]) -> Option<EvaluatedShape>
pub fn evaluate_custom(custom: &CustomGeometry, w: f64, h: f64) -> EvaluatedShape
```

Y-axis inversion (OOXML y-down → PDF y-up) happens during evaluation: `pdf_y = h - ooxml_y`.

## Parsing changes

### `src/docx/smartart.rs` (lines 130-142)
Replace `match prst { "ellipse" => ..., _ => Rect }` with `parse_shape_geometry(sp_pr)` that extracts preset name + `a:avLst` overrides.

### `src/docx/textbox.rs` (lines 381-394)
Same — replace limited prst match with full geometry extraction. Also add `a:custGeom` parsing (pathLst + gdLst + avLst).

## Rendering changes

### `src/pdf/smartart.rs` — `draw_shape_path`
Replace 3-branch ShapeType match with geometry engine dispatch. The function signature changes from `shape: ShapeType` to `geometry: &ShapeGeometry`. Calls `evaluate_preset()` or `evaluate_custom()`, then emits resolved path commands to the PDF content stream.

### `src/pdf/mod.rs` — `render_shape_fill` (line 47-86)
Signature changes from `shape: ShapeType` to `geometry: &ShapeGeometry`. All callers updated.

## Files to modify

| File | Change |
|------|--------|
| `src/model.rs` | Add `ShapeGeometry`, `CustomGeometry`; replace `ShapeType` usage |
| `src/geometry/mod.rs` | **NEW** — public API |
| `src/geometry/formulas.rs` | **NEW** — formula evaluator |
| `src/geometry/path.rs` | **NEW** — path types, arc/quad conversion |
| `src/geometry/definitions.rs` | **NEW** — generated preset data |
| `src/docx/smartart.rs` | Update `parse_dsp_shape` to produce `ShapeGeometry` |
| `src/docx/textbox.rs` | Update `parse_textbox_from_wsp` for full geometry; add `a:custGeom` |
| `src/pdf/smartart.rs` | Replace `draw_shape_path` with geometry engine |
| `src/pdf/mod.rs` | Update `render_shape_fill` signature |
| `tools/generate-shapes/` | **NEW** — codegen tool for definitions.rs |

## Implementation phases

### Phase 1: Geometry engine core
1. ~~Create `src/geometry/` with types, formula evaluator, path commands~~ DONE
2. ~~Unit test all 17 operators + built-in variables~~ DONE
3. ~~Extract arc→cubic from `pdf/mod.rs:267-320`, add quadBez→cubic~~ DONE (completed in Step 1)
4. ~~Hand-code definitions for the 11 shapes in fixtures: rect, ellipse, roundRect, corner, triangle, notchedRightArrow, leftCircularArrow, circularArrow, line, arc, straightConnector1~~ DONE
5. ~~Unit tests comparing evaluated rect/ellipse/notchedRightArrow against current hand-coded paths~~ DONE

### Phase 2: Model migration + rendering integration
1. ~~Add `ShapeGeometry` to model.rs~~ DONE
2. ~~Update SmartArt + textbox parsing to produce `ShapeGeometry`~~ DONE
3. ~~Update `draw_shape_path` and `render_shape_fill` to use geometry engine~~ DONE
4. ~~Run full test suite — verify zero regressions (`cargo test -- --nocapture`, check for "REGRESSION" lines)~~ DONE

### Phase 3: Code generator + all 187 shapes
1. ~~Write `tools/generate-shapes/` that parses `presetShapeDefinitions.xml` → `definitions.rs`~~ DONE
2. ~~Generate all 187 definitions, replace hand-coded Phase 1 definitions~~ DONE
3. ~~Remove old `ShapeType` enum entirely~~ DONE

### Phase 4: Test fixtures
1. ~~**case34 — Preset shape gallery**: Grid of ~20 common preset shapes with solid fills and strokes. Shapes: roundRect, diamond, chevron, pentagon, hexagon, triangle, rightArrow, heart, star5, cloud, donut, flowChartProcess, flowChartDecision, flowChartTerminator, plus, trapezoid, parallelogram, frame, leftArrow, upDownArrow. Generated with python-docx + ZIP post-processing to inject `a:prstGeom` elements.~~ DONE
2. ~~**case35 — Adjusted shapes**: Same shapes but with non-default adjustment values (e.g. roundRect with larger corner radius, star5 with different point depth). Tests that the formula evaluator handles overrides.~~ DONE
3. ~~**case36 — Custom geometry**: Shapes defined with `a:custGeom` paths (moveTo/lnTo/cubicBezTo/arcTo/close), with guide formulas (gdLst) and adjustment defaults (avLst). 6 shapes: star (lnTo), heart (cubicBezTo), arrow (guides), wave (mixed), cross (guides), rounded rect (arcTo).~~ DONE
4. ~~**case37 — SmartArt with diverse shapes**: A SmartArt diagram that exercises multiple preset types (builds on existing SmartArt parsing). Could use a Process or Hierarchy layout that produces roundRects, arrows, connectors.~~ DONE
5. Generate reference PDFs via Word for each, run visual comparison. **PARTIAL** — case34 (94.7% Jaccard, 80.4% SSIM) and case36 (96.7% Jaccard, 91.4% SSIM) done. **BLOCKED**: case35 and case37 need reference.pdf from MS Word.

### Phase 5: Cleanup
1. Unify connector arc rendering to use geometry engine's arc conversion
2. Update roadmap.md and CLAUDE.md
3. Consider caching evaluated shapes (same preset + dimensions = same paths)

## Verification

1. `cargo test -- --nocapture` — zero regressions in existing fixtures
2. `./tools/target/debug/analyze-fixtures` — scores stable or improved
3. New fixtures (case34-37) pass Jaccard threshold
4. Unit tests in `src/geometry/` cover all 17 operators, built-in variables, arc conversion, and path evaluation for key shapes
