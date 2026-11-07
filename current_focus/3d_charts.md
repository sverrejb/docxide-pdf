# Charts: 3-D, 2-D layout and Office 2016 (chartEx)

Charts are drawn as vectors in `src/pdf/charts.rs` (cartesian),
`charts_radial.rs` (pie, doughnut), `chart_legend.rs` and `charts_ex.rs`
(Office 2016 types). Every rule below was measured from Word online exports of
probe documents; each constant in the code names the probe it came from.

## Probe kit

Gitignored, in the main checkout: `tests/output/torture_test/probe3d/`.

- `in*/` hold probe .docx files, `out*/` Word's PDFs, `ours*/` ours.
  - `in`, `in2`: one probe per 3-D type (`persp3d`, `line3d`, `area3d` are the
    rAngAx 0 ones); `in3`: standard grouping with several depth rows;
    `in4`: 3-D pie variants (height, perspective, rotX/rotY); `in5`: 2-D
    layout.
  - `inx`: chartEx with PowerPoint's 133% chartStyle; `inx9`: Word's own
    9/14pt style plus pareto (`mkx9.py`); `inx9b`: no legend, no title, right
    legend (`mkx9b.py`).
- `tools/`: `plotrect2.py` (plot rect from gridlines), `front.py` (3-D front
  face), `barpx.py` (faces from Word's raster), `piemeasure.py`, `cmp.py`.
  Run from a worktree root; they call `tools/pdf_lines.py`.
- `cx/`: `m.py` (text, fills, strokes, y from the top), `cmp.py <set/name>`
  (ours vs Word, run from probe3d), `render.sh`.
- Export: `python3 tools/word_export.py <dir> --out <dir>`, sandbox off, Word
  with no open documents, one batch at a time. A chartEx part only opens when
  its chartStyle and chartColorStyle parts come with it.

## Fixtures

- `cases/charts_3d`: 12 probes (bar depth/rotX variants, cylinder, cone,
  pyramid, stacked, horizontal and horizontal stacked bars, four 3-D pies).
  Left out until drawn right: rAngAx 0 (`persp3d`, `line3d`, `area3d`) and
  several depth rows (`standard3d`, `in3/std*`).
- `cases/charts_ex`: the seven `inx9` types plus the four `inx9b` variants.
- Built by `probe3d/mkfix.py <dir>` from the probes, exported with Word.
  The torture_test fixture has one more 3-D bar chart; case29/30/31/52/53
  cover 2-D charts.

## Done

- **3-D bars and columns** (box, cylinder, cone, pyramid, ToMax cuts;
  clustered, stacked, standard). A bar is a square prism `depthPercent` bar
  widths deep, the floor adds `gapDepth` per row, depth z projects to
  (z sin rotY, z sin rotX). The 3-D plot is the 2-D one inset
  5.0/6.35/2.85/1.77pt; the front shrinks until the box fits, centred. Faces
  shade 1.0/0.76/0.63.
- **3-D pies**: ellipse a·sin rotX tall on a rim 0.261·a·hPercent·cos rotX.
- **3-D lines and areas**, drawn in the right-angle (oblique) view.
- **2-D layout**: plot margins, legend placement and spacing, pie layout,
  markers (rules in `charts.rs`/`chart_legend.rs`).
- **Office 2016 charts** (2026-11-06): waterfall, histogram, pareto, box &
  whisker, funnel, treemap, sunburst, all within 0.6pt of Word.
  - Text comes from the chartStyle part: sizes rounded to half points, theme
    minor font (titles too), greys #595959 / #404040 for data labels.
  - Value axis: 5% headroom, then the smallest 1-2-5 unit with ≤ 10 steps.
  - Treemap rows grow while the mean tile aspect falls (not plain squarify).
  - Sunburst hole = ring width; labels run along the radius, flipped so they
    never read right to left.
  - No fixture has a chartEx part; only the probes check it.

Fixture scores after the chart rounds: case53 74.5% J, case31 78.4%,
case30 93.1%, case29 89.7%, case52 85.6%, torture_test 78.8%.

## Next steps, in order

1. **True perspective (`rAngAx` 0).** Word's default for 3-D line and area:
   it rotates the scene and the floor recedes; we draw the oblique view
   (ponytail comment above the 3-D line/area code in `charts.rs`). Fit the
   camera (rotX, rotY, perspective angle) to the raster corners of
   `in2/persp3d`, `line3d`, `area3d`. Word draws the plot as a raster, so
   use `tools/barpx.py`-style corner finding. 2–4 hours.
2. **Standard grouping, several depth rows** (`in3/std*`). Word's front is
   103pt tall for 2 rows, 96 for 3, whatever the width; we keep the one-row
   fit (≤ 12pt off; 43pt wide when the series axis is hidden). Find the rule
   that gives those heights.
3. **3-D pie details.**
   - Word's default 30° perspective draws the pie ~4% larger (we are
     orthographic; ponytail comment in `charts_radial.rs`).
   - Best-fit label placement (we put labels at 70% of the radius).
   - Legend entries 7.2pt apart in Word, 9.2 in ours.
   - rotY/firstSliceAng and explosion.
4. **Types not drawn yet**: surface (2-D and 3-D), stock, pie-of-pie,
   bar-of-pie. Build probes for each first (a `mk*.py` writing the chart XML,
   then a Word export), then measure.
5. **chartEx leftovers**: pareto plot top 1.2pt lower in Word (cause
   unknown, not fitted); top/left legends unprobed; element `txPr` overrides
   and rich titles not read; region maps not drawn (ponytail comment in
   `docx/chartex.rs`).

Also open: a `keepNext` paragraph followed by a chart paragraph stays at
the page bottom in ours; Word moves it to the next page with the chart (why
the chart fixtures carry no captions).

Also open (2-D): stacked 3-D areas draw a row per series; combo charts draw
only the first type; no secondary axes; value labels only on vertical
clustered bars; no titles on pie/radar.
