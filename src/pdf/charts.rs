use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, to_winansi_bytes};
use crate::model::{
    BarGrouping, BarShape, Chart, ChartType, InlineChart, LegendPosition, MarkerSymbol, View3D,
};

use super::chart_legend::{
    LegendItem, LegendPlacement, SwatchStyle, render_chart_legend, render_swatch,
};
use super::charts_ex;
use super::charts_radial;
use super::color::{fill_rgb, set_fill_alpha, stroke_rgb};
use super::helpers::draw_circle;

fn ceil_nice(val: f32) -> f32 {
    if val <= 0.0 {
        return 1.0;
    }
    let mag = 10.0f32.powf(val.log10().floor());
    let norm = val / mag;
    let nice = if norm <= 1.5 {
        1.0
    } else if norm <= 3.0 {
        2.0
    } else if norm <= 7.0 {
        5.0
    } else {
        10.0
    };
    nice * mag
}

fn nice_tick_step(max_val: f32, target_ticks: usize) -> f32 {
    let raw_step = max_val / target_ticks as f32;
    ceil_nice(raw_step)
}

fn text_width_approx(text: &str, font_size: f32) -> f32 {
    text.len() as f32 * font_size * 0.5
}

pub(super) fn text_width(text: &str, font_size: f32, font: Option<&FontEntry>) -> f32 {
    match font {
        Some(f) => f.word_width(text, font_size, false),
        None => text_width_approx(text, font_size),
    }
}

fn format_tick_label(val: f32, step: f32, format_code: Option<&str>) -> String {
    match format_code {
        Some(fmt) => apply_number_format(val, fmt),
        None => {
            if step.fract() == 0.0 {
                format!("{}", val as i64)
            } else {
                format!("{:.1}", val)
            }
        }
    }
}

fn apply_number_format(val: f32, fmt: &str) -> String {
    let has_comma = fmt.contains(',');
    let has_percent = fmt.contains('%');
    let effective_val = if has_percent { val * 100.0 } else { val };

    let decimals = if let Some(dot_pos) = fmt.rfind('.') {
        fmt[dot_pos + 1..]
            .chars()
            .take_while(|&c| c == '0' || c == '#')
            .count()
    } else {
        0
    };

    let formatted = format!("{:.prec$}", effective_val as f64, prec = decimals);

    let result = if has_comma {
        insert_thousands_separator(&formatted)
    } else {
        formatted
    };

    if has_percent {
        format!("{}%", result)
    } else {
        result
    }
}

fn insert_thousands_separator(s: &str) -> String {
    let (int_part, dec_part) = match s.find('.') {
        Some(pos) => (&s[..pos], Some(&s[pos..])),
        None => (s, None),
    };
    let negative = int_part.starts_with('-');
    let digits = if negative { &int_part[1..] } else { int_part };
    let mut result = String::new();
    for (i, ch) in digits.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(ch);
    }
    let mut out: String = result.chars().rev().collect();
    if negative {
        out.insert(0, '-');
    }
    if let Some(d) = dec_part {
        out.push_str(d);
    }
    out
}

fn wrap_label(text: &str, max_width: f32, font_size: f32, font: Option<&FontEntry>) -> Vec<String> {
    let full_w = text_width(text, font_size, font);
    if full_w <= max_width || !text.contains(' ') {
        return vec![text.to_string()];
    }
    let words: Vec<&str> = text.split_whitespace().collect();
    let mut lines: Vec<String> = Vec::new();
    let mut current_line = String::new();
    for word in &words {
        let test = if current_line.is_empty() {
            word.to_string()
        } else {
            format!("{} {}", current_line, word)
        };
        if text_width(&test, font_size, font) > max_width && !current_line.is_empty() {
            lines.push(current_line);
            current_line = word.to_string();
        } else {
            current_line = test;
        }
    }
    if !current_line.is_empty() {
        lines.push(current_line);
    }
    lines
}

fn set_color(content: &mut Content, color: Option<[u8; 3]>) {
    if let Some(c) = color {
        fill_rgb(content, c);
    }
}

fn set_stroke_color(content: &mut Content, color: Option<[u8; 3]>) {
    if let Some(c) = color {
        stroke_rgb(content, c);
    }
}

static DEFAULT_AXIS_COLOR: [u8; 3] = [179, 179, 179];

const DEFAULT_MARKER_CYCLE: [MarkerSymbol; 3] = [
    MarkerSymbol::Diamond,
    MarkerSymbol::Square,
    MarkerSymbol::Triangle,
];

fn resolve_marker(marker: Option<MarkerSymbol>, series_index: usize) -> MarkerSymbol {
    marker.unwrap_or(DEFAULT_MARKER_CYCLE[series_index % DEFAULT_MARKER_CYCLE.len()])
}

/// A series' marker as Word draws it: `c:size` (default 5) wide plus its
/// outline, which is the marker's line, else the series' line, else
/// `default_lw` wide; filled with the marker's own colour, else the theme
/// accent for the series (case30's orange line carries accent2's red
/// squares, 7.25pt with the default 2.25pt line). Returns the fill and the
/// radius `draw_marker` takes for that symbol.
fn marker_look(
    series: &crate::model::ChartSeries,
    si: usize,
    accents: &[[u8; 3]],
    symbol: MarkerSymbol,
    default_lw: f32,
) -> ([u8; 3], f32) {
    let outline = series
        .marker_line_width
        .or(series.line_width)
        .unwrap_or(default_lw);
    let size = series.marker_size.unwrap_or(5.0) + outline;
    let color = series.marker_color.unwrap_or(accents[si % accents.len()]);
    let r = if symbol == MarkerSymbol::Square {
        size / 1.6
    } else {
        size / 2.0
    };
    (color, r)
}

/// Word's default series line on line and radar charts.
const DEFAULT_LINE_W: f32 = 2.25;

pub(super) fn draw_marker(content: &mut Content, symbol: MarkerSymbol, x: f32, y: f32, r: f32) {
    match symbol {
        MarkerSymbol::None => {}
        MarkerSymbol::Circle | MarkerSymbol::Dot => {
            draw_circle(content, x, y, r);
            content.fill_nonzero();
        }
        MarkerSymbol::Square => {
            content.rect(x - r * 0.8, y - r * 0.8, r * 1.6, r * 1.6);
            content.fill_nonzero();
        }
        MarkerSymbol::Diamond => {
            content.move_to(x, y + r);
            content.line_to(x + r, y);
            content.line_to(x, y - r);
            content.line_to(x - r, y);
            content.close_path();
            content.fill_nonzero();
        }
        MarkerSymbol::Triangle => {
            content.move_to(x, y + r);
            content.line_to(x + r, y - r * 0.7);
            content.line_to(x - r, y - r * 0.7);
            content.close_path();
            content.fill_nonzero();
        }
        MarkerSymbol::Plus => {
            let t = r * 0.25;
            content.rect(x - r, y - t, r * 2.0, t * 2.0);
            content.fill_nonzero();
            content.rect(x - t, y - r, t * 2.0, r * 2.0);
            content.fill_nonzero();
        }
        MarkerSymbol::X => {
            let d = r * 0.7;
            let t = r * 0.25;
            // Two rotated rectangles approximated as thin diamonds
            content.move_to(x, y + d + t);
            content.line_to(x + t, y + d);
            content.line_to(x + d + t, y);
            content.line_to(x + d, y - t);
            content.line_to(x, y - d - t);
            content.line_to(x - t, y - d);
            content.line_to(x - d - t, y);
            content.line_to(x - d, y + t);
            content.close_path();
            content.fill_nonzero();
        }
        MarkerSymbol::Star => {
            // 5-point star
            for i in 0..5 {
                let angle_outer =
                    std::f32::consts::FRAC_PI_2 + i as f32 * std::f32::consts::PI * 2.0 / 5.0;
                let angle_inner = angle_outer + std::f32::consts::PI / 5.0;
                let ox = x + r * angle_outer.cos();
                let oy = y + r * angle_outer.sin();
                let ix = x + r * 0.4 * angle_inner.cos();
                let iy = y + r * 0.4 * angle_inner.sin();
                if i == 0 {
                    content.move_to(ox, oy);
                } else {
                    content.line_to(ox, oy);
                }
                content.line_to(ix, iy);
            }
            content.close_path();
            content.fill_nonzero();
        }
        MarkerSymbol::Dash => {
            content.rect(x - r, y - r * 0.2, r * 2.0, r * 0.4);
            content.fill_nonzero();
        }
    }
}

// Pie chart colors when series data points don't have individual colors
pub(super) const DEFAULT_PIE_COLORS: &[[u8; 3]] = &[
    [68, 114, 196],
    [237, 125, 49],
    [165, 165, 165],
    [255, 192, 0],
    [91, 155, 213],
    [112, 173, 71],
    [38, 68, 120],
    [158, 72, 14],
];

pub(super) fn resolve_accent_colors(accent_colors: &[[u8; 3]]) -> &[[u8; 3]] {
    if accent_colors.is_empty() {
        DEFAULT_PIE_COLORS
    } else {
        accent_colors
    }
}

pub(super) fn show_text(
    content: &mut Content,
    font_key: &str,
    font_size: f32,
    x: f32,
    y: f32,
    text: &str,
    font_entry: Option<&FontEntry>,
) {
    let matrix = [1.0, 0.0, 0.0, 1.0, x, y];
    show_text_matrix(content, font_key, font_size, matrix, text, font_entry);
}

/// `show_text` placed by a whole text matrix, for turned text.
pub(super) fn show_text_matrix(
    content: &mut Content,
    font_key: &str,
    font_size: f32,
    matrix: [f32; 6],
    text: &str,
    font_entry: Option<&FontEntry>,
) {
    let bytes = match font_entry {
        Some(e) => e.encode(text),
        None => to_winansi_bytes(text),
    };
    content
        .begin_text()
        .set_font(Name(font_key.as_bytes()), font_size)
        .set_text_matrix(matrix)
        .show(Str(&bytes))
        .end_text();
}

fn compute_axis_range(max_val: f32, target_ticks: usize, headroom_threshold: f32) -> (f32, f32) {
    let ts = nice_tick_step(max_val, target_ticks);
    let mut am = (max_val / ts).ceil() * ts;
    if max_val >= am * headroom_threshold {
        am += ts;
    }
    (ts, am)
}

/// A category's slot along the category axis and the width of one bar in it.
fn bar_slot(c: &Chart, cat_extent: f32, num_categories: usize) -> (f32, f32) {
    let group = cat_extent / num_categories as f32;
    let slots = match c.chart_type {
        ChartType::Bar {
            grouping: BarGrouping::Clustered,
            ..
        } => c.series.len() as f32,
        _ => 1.0,
    };
    (group, group / (slots + c.gap_width_pct / 100.0))
}

/// A 3-D column chart's depth axis. With right-angle axes Word keeps the
/// front face a rectangle and draws a point `z` deep at (z sin rotY,
/// z sin rotX) from where it would sit on the front plane (five Word probes
/// varying rotX, rotY, depthPercent and the chart width).
/// Depth rows on a 3-D chart's floor: one per series when they stand one
/// behind another (3-D lines and areas always do), else one.
/// ponytail: stacked 3-D areas still get a row per series.
fn depth_rows(c: &Chart) -> usize {
    match c.chart_type {
        ChartType::Bar {
            grouping: BarGrouping::Standard,
            ..
        }
        | ChartType::Line
        | ChartType::Area => c.series.len().max(1),
        _ => 1,
    }
}

/// Word's 3-D plot is its 2-D one inset this far left, right, top and
/// bottom (probes with and without a data table).
const INSET_3D: (f32, f32, f32, f32) = (5.0, 6.35, 2.85, 1.77);

impl View3D {
    /// How deep one row of bars sits on the floor, its gap included.
    fn row_pitch(self, thick: f32) -> f32 {
        thick * self.depth * (1.0 + self.gap_depth)
    }

    fn project(self, z: f32) -> (f32, f32) {
        (
            z * self.rot_y.to_radians().sin(),
            z * self.rot_x.to_radians().sin(),
        )
    }

    /// The front face in a 3-D chart's plot area and where the floor's back
    /// edge sits from its front edge. Word's 3-D plot is its 2-D one inset
    /// 5.0/6.35pt left/right and 2.85/1.77pt top/bottom (probes with and
    /// without a data table); the front shrinks until the box with its depth
    /// fits, centred both ways, beside `reserve` points kept on the right
    /// for series names.
    /// ponytail: with several depth rows Word's front no longer keeps the
    /// plot's proportions (standard3d/std3 probes: 103pt tall for 2 rows, 96
    /// for 3, whatever the width; the box fills the width, 390pt, or 345pt
    /// beside series names); we keep the one-row rule, within ~12pt of it.
    fn front(
        self,
        c: &Chart,
        plot: PlotRect,
        num_categories: usize,
        reserve: f32,
    ) -> (PlotRect, f32) {
        let (left, right, top, bottom) = INSET_3D;
        let inner = PlotRect {
            x: plot.x + left,
            y: plot.y + bottom,
            w: plot.w - (left + right),
            h: plot.h - (top + bottom),
        };
        let horizontal = matches!(
            c.chart_type,
            ChartType::Bar {
                horizontal: true,
                ..
            }
        );
        let cat_extent = if horizontal { inner.h } else { inner.w };
        let depth = self.floor_depth(c, cat_extent, num_categories);
        let (dx, dy) = self.project(depth);
        let avail_w = inner.w - reserve;
        let s = (inner.h / (inner.h + dy)).min(avail_w / (inner.w + dx));
        let front = PlotRect {
            x: inner.x + (avail_w - s * (inner.w + dx)) / 2.0,
            y: inner.y + (inner.h - s * (inner.h + dy)) / 2.0,
            w: s * inner.w,
            h: s * inner.h,
        };
        (front, s * depth)
    }

    /// The floor's depth: a bar's depth plus the gap between depth rows, per
    /// row, with the categories spread over `cat_extent`.
    fn floor_depth(self, c: &Chart, cat_extent: f32, num_categories: usize) -> f32 {
        let thick = bar_slot(c, cat_extent, num_categories).1;
        self.row_pitch(thick) * depth_rows(c) as f32
    }

    /// The world's front face, the floor's depth and how they reach the page.
    fn space(
        self,
        c: &Chart,
        plot: PlotRect,
        num_categories: usize,
        reserve: f32,
    ) -> (PlotRect, f32, Space) {
        let (front, depth, camera) = if self.perspective.is_none() {
            let (front, depth) = self.front(c, plot, num_categories, reserve);
            (front, depth, None)
        } else {
            // Word's box is as wide against its height as the plot, or
            // 2n/(n + 2) of that for n categories over fewer depth rows,
            // more than one (probe3d inp/, inb/, inr/, inr2/: 1.00 ±0.005 for
            // one row or as many rows as categories; 4 categories 1.33, 6
            // 1.50, 8 1.60, each ±0.005).
            // ponytail: 4 rows over 8 categories measured 1.33, not 1.6
            // (inr2/s4c8); the rule behind that is open.
            let rows = depth_rows(c);
            let n = num_categories as f32;
            let widen = if rows > 1 && rows < num_categories {
                2.0 * n / (n + 2.0)
            } else {
                1.0
            };
            let front = PlotRect {
                w: plot.w * widen,
                ..plot
            };
            let depth = self.floor_depth(c, front.w, num_categories);
            (front, depth, Some(Camera::fit(plot, front, depth, self)))
        };
        let space = Space { view: self, camera };
        (front, depth, space)
    }
}

/// Where a 3-D chart's world lands on the page. World x and y are a front
/// face's, in page points like a 2-D plot's; z runs back from the front,
/// drawn obliquely with right-angle axes, else through the camera.
#[derive(Clone, Copy)]
struct Space {
    view: View3D,
    camera: Option<Camera>,
}

impl Space {
    fn at(self, x: f32, y: f32, z: f32) -> (f32, f32) {
        match self.camera {
            Some(camera) => camera.at(x, y, z),
            None => {
                let (dx, dy) = self.view.project(z);
                (x + dx, y + dy)
            }
        }
    }
}

/// A world point on the page; a 2-D chart has no depth.
fn on_page(space: Option<Space>, x: f32, y: f32, z: f32) -> (f32, f32) {
    space.map_or((x, y), |s| s.at(x, y, z))
}

/// Word's camera for `rAngAx` 0, fitted to the edges of its rasters of
/// probes varying rotX, rotY, perspective and the plot's size (probe3d
/// line3d, persp3d and inp/): it orbits the box's centre by rotY, then tilts
/// down by rotX, its axis through the plot's centre, with a vertical field of
/// view of perspective/2 degrees spanning the plot's height.
#[derive(Clone, Copy)]
struct Camera {
    centre: [f32; 3],
    /// (sin, cos) of rotX and rotY.
    tilt: (f32, f32),
    turn: (f32, f32),
    focal: f32,
    /// The box centre's distance from the eye and its offset off the axis,
    /// world units, in the camera's frame.
    dist: f32,
    offset: (f32, f32),
    /// The page point on the camera's axis.
    eye: (f32, f32),
}

impl Camera {
    /// The world box `front` × `depth` seen in `plot`: Word stands it off
    /// until its eight corners fill the plot inset 5.7pt each side and 2.3
    /// above and below, centred (the right-angle insets 5.0/6.35/2.85/1.77pt,
    /// split evenly).
    /// ponytail: the corners land 0.7–1.5pt from Word's at the default view
    /// and up to 3pt at rotX 0 or rotY 40 and perspective 60 (inp/ probes,
    /// mean); its true fitting rule is not recovered.
    fn fit(plot: PlotRect, front: PlotRect, depth: f32, view: View3D) -> Self {
        let (left, right, top, bottom) = INSET_3D;
        let (margin_x, margin_y) = ((left + right) / 2.0, (top + bottom) / 2.0);
        let rad = |deg: f32| deg.to_radians().sin_cos();
        let eye = (plot.x + plot.w / 2.0, plot.y + plot.h / 2.0);
        let mut camera = Camera {
            centre: [
                front.x + front.w / 2.0,
                front.y + front.h / 2.0,
                depth / 2.0,
            ],
            tilt: rad(view.rot_x),
            turn: rad(view.rot_y),
            focal: plot.h / 2.0 / (view.perspective.unwrap_or(30.0) / 4.0).to_radians().tan(),
            dist: 6.0 * front.h,
            offset: (0.0, 0.0),
            eye,
        };
        let corners: [[f32; 3]; 8] = std::array::from_fn(|i| {
            let pick = |bit: usize, lo: f32, len: f32| lo + if i & bit != 0 { len } else { 0.0 };
            [
                pick(1, front.x, front.w),
                pick(2, front.y, front.h),
                pick(4, 0.0, depth),
            ]
        });
        for _ in 0..50 {
            let (x0, x1, y0, y1) = corners.iter().map(|p| camera.at(p[0], p[1], p[2])).fold(
                (f32::MAX, f32::MIN, f32::MAX, f32::MIN),
                |(a, b, c, d), p| (a.min(p.0), b.max(p.0), c.min(p.1), d.max(p.1)),
            );
            let k = camera.focal / camera.dist;
            camera.offset.0 += (eye.0 - (x0 + x1) / 2.0) / k;
            camera.offset.1 += (eye.1 - (y0 + y1) / 2.0) / k;
            camera.dist *=
                ((x1 - x0) / (plot.w - 2.0 * margin_x)).max((y1 - y0) / (plot.h - 2.0 * margin_y));
        }
        camera
    }

    fn at(self, x: f32, y: f32, z: f32) -> (f32, f32) {
        let (px, py, pz) = (x - self.centre[0], y - self.centre[1], z - self.centre[2]);
        let ((sy, cy), (sx, cx)) = (self.turn, self.tilt);
        let (vx, vz) = (px * cy + pz * sy, pz * cy - px * sy);
        let (vy, vz) = (py * cx + vz * sx, vz * cx - py * sx);
        let k = self.focal / (self.dist + vz);
        (
            self.eye.0 + k * (vx + self.offset.0),
            self.eye.1 + k * (vy + self.offset.1),
        )
    }
}

/// 3-D lines and areas: each series a depth row (as standard-grouped bars),
/// a line a ribbon one bar depth deep through the category centres, an area
/// a slab under it. Fronts take the series colour, ribbons and slab tops
/// 0.82× and a slab's end 0.63× (line3d/area3d probes).
fn draw_ribbons(
    c: &Chart,
    content: &mut Content,
    rect: PlotRect,
    num_categories: usize,
    axis_max: f32,
    space: Space,
    area: bool,
) {
    let view = space.view;
    let (group, thick) = bar_slot(c, rect.w, num_categories);
    let row = thick * view.depth;
    let shade = |color: Option<[u8; 3]>, k: f32| shade(color.unwrap_or([0, 0, 0]), k);
    for (si, series) in c.series.iter().enumerate().rev() {
        let z = si as f32 * view.row_pitch(thick) + row * view.gap_depth / 2.0;
        let pts: Vec<(f32, f32)> = series
            .values
            .iter()
            .take(num_categories)
            .enumerate()
            .map(|(ci, &v)| {
                (
                    rect.x + (ci as f32 + 0.5) * group,
                    rect.y + v / axis_max * rect.h,
                )
            })
            .collect();
        let (Some(&first), Some(&last)) = (pts.first(), pts.last()) else {
            continue;
        };
        let quad = |content: &mut Content, a: (f32, f32), b: (f32, f32)| {
            polygon(
                content,
                &[
                    space.at(a.0, a.1, z),
                    space.at(b.0, b.1, z),
                    space.at(b.0, b.1, z + row),
                    space.at(a.0, a.1, z + row),
                ],
            );
            content.fill_nonzero();
        };
        fill_rgb(content, shade(series.color, 0.82));
        for pair in pts.windows(2) {
            quad(content, pair[0], pair[1]);
        }
        if area {
            fill_rgb(content, shade(series.color, 0.63));
            quad(content, last, (last.0, rect.y));
            let front: Vec<(f32, f32)> = std::iter::once((first.0, rect.y))
                .chain(pts.iter().copied())
                .chain(std::iter::once((last.0, rect.y)))
                .map(|p| space.at(p.0, p.1, z))
                .collect();
            fill_rgb(content, shade(series.color, 1.0));
            polygon(content, &front);
            content.fill_nonzero();
        }
    }
}

/// A 3-D bar standing on its front face's bottom edge (x, y), `w` wide, `h`
/// tall and from z0 to z1 deep in the world `map` puts on the page, its
/// bottom and top sections scaled by `taper` about the bar's axis (1 for a
/// box or cylinder, 0 at a cone's or pyramid's tip). Faces take the
/// series colour shaded as Word's lighting draws them, sampled at every
/// probed rotation: box front 1.0, right 0.63, top 0.76; pyramid front 1.06,
/// right 0.69; round solids run 0.72 → 1.07 (at 45%) → 0.80 left to right.
fn draw_bar_3d(
    content: &mut Content,
    color: Option<[u8; 3]>,
    shape: BarShape,
    map: impl Fn(f32, f32, f32) -> (f32, f32),
    (x, y, w, h): (f32, f32, f32, f32),
    (z0, z1): (f32, f32),
    (bottom, top): (f32, f32),
) {
    let base = color.unwrap_or([0, 0, 0]);
    let fill = |content: &mut Content, k: f32, pts: &[(f32, f32)]| {
        fill_rgb(content, shade(base, k));
        polygon(content, pts);
        content.fill_nonzero();
    };
    // Where a point of the bottom's outline at (u across, v deep), both
    // 0..1, sits at height `z` on a section scaled about the axis.
    let at = |u: f32, v: f32, z: f32| {
        let s = if z > 0.0 { top } else { bottom };
        let (u, v) = (0.5 + (u - 0.5) * s, 0.5 + (v - 0.5) * s);
        map(x + u * w, y + z, z0 + v * (z1 - z0))
    };
    match shape {
        BarShape::Box | BarShape::Pyramid | BarShape::PyramidToMax => {
            let (front, side) = if shape == BarShape::Box {
                (1.0, 0.63)
            } else {
                (1.06, 0.69)
            };
            fill(
                content,
                side,
                &[
                    at(1.0, 0.0, 0.0),
                    at(1.0, 1.0, 0.0),
                    at(1.0, 1.0, h),
                    at(1.0, 0.0, h),
                ],
            );
            if top > 0.0 {
                fill(
                    content,
                    0.76,
                    &[
                        at(0.0, 0.0, h),
                        at(1.0, 0.0, h),
                        at(1.0, 1.0, h),
                        at(0.0, 1.0, h),
                    ],
                );
            }
            fill(
                content,
                front,
                &[
                    at(0.0, 0.0, 0.0),
                    at(1.0, 0.0, 0.0),
                    at(1.0, 0.0, h),
                    at(0.0, 0.0, h),
                ],
            );
        }
        BarShape::Cylinder | BarShape::Cone | BarShape::ConeToMax => {
            const STEPS: usize = 48;
            let ring = |z: f32| -> Vec<(f32, f32)> {
                (0..STEPS)
                    .map(|i| {
                        let t = i as f32 / STEPS as f32 * std::f32::consts::TAU;
                        at(0.5 + 0.5 * t.cos(), 0.5 + 0.5 * t.sin(), z)
                    })
                    .collect()
            };
            let cap = ring(h);
            let mut body = ring(0.0);
            body.extend(&cap);
            let body = convex_hull(body);
            let (x0, x1) = body
                .iter()
                .fold((f32::MAX, f32::MIN), |(a, b), p| (a.min(p.0), b.max(p.0)));
            let (y0, y1) = body
                .iter()
                .fold((f32::MAX, f32::MIN), |(a, b), p| (a.min(p.1), b.max(p.1)));
            fill_strips(content, &body, (x0, y0, x1, y1), 24, base, |t| {
                if t < 0.45 {
                    0.72 + 0.35 * t / 0.45
                } else {
                    1.07 - 0.27 * (t - 0.45) / 0.55
                }
            });
            if top > 0.0 {
                fill(content, 0.76, &cap);
            }
        }
    }
}

/// `base` scaled by `k`, per channel.
pub(super) fn shade(base: [u8; 3], k: f32) -> [u8; 3] {
    base.map(|v| (v as f32 * k).round().min(255.0) as u8)
}

/// Shades the inside of `outline` in `n` vertical strips across
/// `(x0, y0, x1, y1)`, strip `t` (0..1, at its centre) in `base` scaled by `k(t)`.
pub(super) fn fill_strips(
    content: &mut Content,
    outline: &[(f32, f32)],
    (x0, y0, x1, y1): (f32, f32, f32, f32),
    n: usize,
    base: [u8; 3],
    k: impl Fn(f32) -> f32,
) {
    content.save_state();
    polygon(content, outline);
    content.clip_nonzero();
    content.end_path();
    let sw = (x1 - x0) / n as f32;
    for i in 0..n {
        let t = (i as f32 + 0.5) / n as f32;
        fill_rgb(content, shade(base, k(t)));
        // Each strip runs to the right edge so the next covers the join.
        content.rect(x0 + i as f32 * sw, y0, x1 - x0 - i as f32 * sw, y1 - y0);
        content.fill_nonzero();
    }
    content.restore_state();
}

pub(super) fn polygon(content: &mut Content, pts: &[(f32, f32)]) {
    if !pts.is_empty() {
        polyline(content, pts.iter().copied());
        content.close_path();
    }
}

/// An open path through `pts`.
fn polyline(content: &mut Content, pts: impl IntoIterator<Item = (f32, f32)>) {
    let mut pts = pts.into_iter();
    if let Some((x, y)) = pts.next() {
        content.move_to(x, y);
        for (x, y) in pts {
            content.line_to(x, y);
        }
    }
}

/// Andrew's monotone chain, counter-clockwise.
fn convex_hull(mut pts: Vec<(f32, f32)>) -> Vec<(f32, f32)> {
    pts.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let cross = |o: (f32, f32), a: (f32, f32), b: (f32, f32)| {
        (a.0 - o.0) * (b.1 - o.1) - (a.1 - o.1) * (b.0 - o.0)
    };
    let mut hull: Vec<(f32, f32)> = Vec::with_capacity(pts.len() + 1);
    for pass in [false, true] {
        let start = hull.len();
        let iter: Box<dyn Iterator<Item = &(f32, f32)>> = if pass {
            Box::new(pts.iter().rev())
        } else {
            Box::new(pts.iter())
        };
        for &p in iter {
            while hull.len() >= start + 2
                && cross(hull[hull.len() - 2], hull[hull.len() - 1], p) <= 0.0
            {
                hull.pop();
            }
            hull.push(p);
        }
        hull.pop();
    }
    hull
}

#[derive(Clone, Copy)]
struct PlotRect {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
}

/// Render a chart's data series (bars/line/area/scatter/bubble) into the plot
/// rectangle. Pie/doughnut/radar use dedicated renderers and never reach here.
/// Split out of render_chart, which keeps axis/gridline/legend orchestration.
fn draw_chart_series(
    c: &Chart,
    content: &mut Content,
    rect: PlotRect,
    num_categories: usize,
    axis_max: f32,
    x_axis_max: f32,
    space: Option<Space>,
    alpha_states: &mut HashSet<u8>,
) -> Vec<ValueLabel> {
    let at = |x: f32, y: f32, z: f32| on_page(space, x, y, z);
    let mut labels = Vec::new();
    let PlotRect {
        x: plot_x,
        y: plot_y,
        w: plot_w,
        h: plot_h,
    } = rect;
    match c.chart_type {
        ChartType::Bar {
            horizontal,
            grouping,
        } => {
            // Categories run along one side of the plot, values along the other.
            let (cat_origin, cat_extent, val_origin, val_extent) = if horizontal {
                (plot_y, plot_h, plot_x, plot_w)
            } else {
                (plot_x, plot_w, plot_y, plot_h)
            };
            let stacked = matches!(grouping, BarGrouping::Stacked | BarGrouping::PercentStacked);
            let is_percent_stacked = grouping == BarGrouping::PercentStacked;
            let (group, thick) = bar_slot(c, cat_extent, num_categories);
            let gap = c.gap_width_pct / 100.0 * thick;
            // A 3-D bar is `depth` bar widths deep and its front face sits half
            // the depth rows' gap into the floor.
            let z_front = space.map_or(0.0, |s| thick * s.view.depth * s.view.gap_depth / 2.0);
            // `stack` is the whole column's length: a cone or pyramid narrows
            // over the stack, a `ToMax` one over the value axis.
            let bar = |content: &mut Content,
                       color: Option<[u8; 3]>,
                       cat_pos: f32,
                       val_pos: f32,
                       len: f32,
                       stack: f32,
                       z: f32| {
                if let Some(space) = space {
                    let (shape, deep) = (space.view.shape, thick * space.view.depth);
                    let span = match shape {
                        BarShape::Cone | BarShape::Pyramid => stack,
                        BarShape::ConeToMax | BarShape::PyramidToMax => val_extent,
                        BarShape::Box | BarShape::Cylinder => f32::INFINITY,
                    };
                    let lo = val_pos - val_origin;
                    let taper = (1.0 - lo / span, 1.0 - (lo + len) / span);
                    let map = |x, y, z| space.at(x, y, z);
                    let zs = (z, z + deep);
                    if !horizontal {
                        let rect = (cat_pos, val_pos, thick, len);
                        draw_bar_3d(content, color, shape, map, rect, zs, taper);
                    } else if shape == BarShape::Box {
                        // A box lights by screen direction: top 0.76, end 0.63.
                        let rect = (val_pos, cat_pos, len, thick);
                        draw_bar_3d(content, color, shape, map, rect, zs, taper);
                    } else {
                        // A lying solid is a standing one mirrored about y = x.
                        content.save_state();
                        content.transform([0.0, 1.0, 1.0, 0.0, 0.0, 0.0]);
                        let rect = (cat_pos, val_pos, thick, len);
                        let mirrored = |x: f32, y: f32, z: f32| {
                            let (a, b) = space.at(y, x, z);
                            (b, a)
                        };
                        draw_bar_3d(content, color, shape, mirrored, rect, zs, taper);
                        content.restore_state();
                    }
                    return;
                }
                set_color(content, color);
                if horizontal {
                    content.rect(val_pos, cat_pos, len, thick);
                } else {
                    content.rect(cat_pos, val_pos, thick, len);
                }
                content.fill_nonzero();
            };

            // A label past a bar's end: after a lying bar, above a column.
            let end_label = |pos: f32, len: f32, value: f32, z: f32| {
                let (end, across) = (val_origin + len, pos + thick / 2.0);
                if horizontal {
                    ValueLabel {
                        at: at(end, across, z),
                        value,
                        place: LabelPlace::After,
                    }
                } else {
                    ValueLabel {
                        at: at(across, end, z),
                        value,
                        place: LabelPlace::Above,
                    }
                }
            };

            if grouping == BarGrouping::Standard {
                // Each series a row further back, the first in front: draw
                // the back row first.
                let pitch = space.map_or(0.0, |s| s.view.row_pitch(thick));
                for (si, series) in c.series.iter().enumerate().rev() {
                    let z = z_front + pitch * si as f32;
                    for ci in 0..num_categories {
                        let val = series.values.get(ci).copied().unwrap_or(0.0);
                        let len = (val / axis_max) * val_extent;
                        let pos = cat_origin + ci as f32 * group + gap / 2.0;
                        bar(content, series.color, pos, val_origin, len, len, z);
                        if c.show_values {
                            labels.push(end_label(pos, len, val, z));
                        }
                    }
                }
                return labels;
            }

            for ci in 0..num_categories {
                let group_pos = cat_origin + ci as f32 * group + gap / 2.0;
                if stacked {
                    let cat_total = if is_percent_stacked {
                        c.series
                            .iter()
                            .map(|s| s.values.get(ci).copied().unwrap_or(0.0))
                            .sum::<f32>()
                    } else {
                        0.0
                    };
                    let lens: Vec<f32> = c
                        .series
                        .iter()
                        .map(|series| {
                            let raw = series.values.get(ci).copied().unwrap_or(0.0);
                            let val = if is_percent_stacked && cat_total > 0.0 {
                                (raw / cat_total) * 100.0
                            } else {
                                raw
                            };
                            (val / axis_max) * val_extent
                        })
                        .collect();
                    let stack: f32 = lens.iter().sum();
                    let mut cumulative = 0.0;
                    for (series, len) in c.series.iter().zip(lens) {
                        let pos = val_origin + cumulative;
                        bar(content, series.color, group_pos, pos, len, stack, z_front);
                        if c.show_values {
                            let (along, across) = (pos + len / 2.0, group_pos + thick / 2.0);
                            labels.push(ValueLabel {
                                at: if horizontal {
                                    at(along, across, z_front)
                                } else {
                                    at(across, along, z_front)
                                },
                                value: series.values.get(ci).copied().unwrap_or(0.0),
                                place: LabelPlace::Center,
                            });
                        }
                        cumulative += len;
                    }
                } else {
                    for (si, series) in c.series.iter().enumerate() {
                        let val = series.values.get(ci).copied().unwrap_or(0.0);
                        let len = (val / axis_max) * val_extent;
                        let pos = group_pos + si as f32 * thick;
                        bar(content, series.color, pos, val_origin, len, len, z_front);
                        if c.show_values {
                            labels.push(end_label(pos, len, val, z_front));
                        }
                    }
                }
            }
        }
        ChartType::Line | ChartType::Area if let Some(space) = space => {
            let area = matches!(c.chart_type, ChartType::Area);
            draw_ribbons(c, content, rect, num_categories, axis_max, space, area);
        }
        ChartType::Line => {
            let accents = resolve_accent_colors(&c.accent_colors);
            let group_w = plot_w / num_categories as f32;
            for (si, series) in c.series.iter().enumerate() {
                content.set_line_width(series.line_width.unwrap_or(DEFAULT_LINE_W));
                set_stroke_color(content, series.color);
                let pts: Vec<(f32, f32)> = series
                    .values
                    .iter()
                    .enumerate()
                    .map(|(ci, &val)| {
                        (
                            plot_x + (ci as f32 + 0.5) * group_w,
                            plot_y + (val / axis_max) * plot_h,
                        )
                    })
                    .collect();
                if let Some(&(x0, y0)) = pts.first() {
                    content.move_to(x0, y0);
                    if pts.len() >= 2 {
                        for i in 0..pts.len() - 1 {
                            let prev = if i > 0 {
                                pts[i - 1]
                            } else {
                                (2.0 * pts[0].0 - pts[1].0, 2.0 * pts[0].1 - pts[1].1)
                            };
                            let next = if i + 2 < pts.len() {
                                pts[i + 2]
                            } else {
                                let n = pts.len() - 1;
                                (2.0 * pts[n].0 - pts[n - 1].0, 2.0 * pts[n].1 - pts[n - 1].1)
                            };
                            let cp1x = pts[i].0 + (pts[i + 1].0 - prev.0) / 6.0;
                            let cp1y = pts[i].1 + (pts[i + 1].1 - prev.1) / 6.0;
                            let cp2x = pts[i + 1].0 - (next.0 - pts[i].0) / 6.0;
                            let cp2y = pts[i + 1].1 - (next.1 - pts[i].1) / 6.0;
                            content.cubic_to(cp1x, cp1y, cp2x, cp2y, pts[i + 1].0, pts[i + 1].1);
                        }
                    }
                }
                content.stroke();

                let sym = resolve_marker(series.marker, si);
                let (color, marker_r) = marker_look(series, si, accents, sym, DEFAULT_LINE_W);
                fill_rgb(content, color);
                for (ci, &val) in series.values.iter().enumerate() {
                    let lx = plot_x + (ci as f32 + 0.5) * group_w;
                    let ly = plot_y + (val / axis_max) * plot_h;
                    draw_marker(content, sym, lx, ly, marker_r);
                }
            }
        }
        ChartType::Area => {
            let cat_w = plot_w / (num_categories - 1).max(1) as f32;
            for series in &c.series {
                let pts: Vec<(f32, f32)> = series
                    .values
                    .iter()
                    .enumerate()
                    .map(|(ci, &val)| {
                        (
                            plot_x + ci as f32 * cat_w,
                            plot_y + (val / axis_max) * plot_h,
                        )
                    })
                    .collect();

                set_color(content, series.color);
                content.move_to(plot_x, plot_y);
                for &(lx, ly) in &pts {
                    content.line_to(lx, ly);
                }
                if let Some(&(last_x, _)) = pts.last() {
                    content.line_to(last_x, plot_y);
                }
                content.close_path();
                content.fill_nonzero();

                set_stroke_color(content, series.color);
                content.set_line_width(1.5);
                for (i, &(lx, ly)) in pts.iter().enumerate() {
                    if i == 0 {
                        content.move_to(lx, ly);
                    } else {
                        content.line_to(lx, ly);
                    }
                }
                content.stroke();
            }
        }
        ChartType::Scatter => {
            let accents = resolve_accent_colors(&c.accent_colors);
            for (si, series) in c.series.iter().enumerate() {
                let sym = resolve_marker(series.marker, si);
                // A scatter marker's outline defaults to a hairline 0.75pt.
                let (color, marker_r) = marker_look(series, si, accents, sym, 0.75);
                fill_rgb(content, color);
                if let Some(ref x_vals) = series.x_values {
                    for (&xv, &yv) in x_vals.iter().zip(series.values.iter()) {
                        let px = plot_x + (xv / x_axis_max) * plot_w;
                        let py = plot_y + (yv / axis_max) * plot_h;
                        draw_marker(content, sym, px, py, marker_r);
                    }
                }
            }
        }
        ChartType::Bubble => {
            let max_size = c
                .series
                .iter()
                .filter_map(|s| s.bubble_sizes.as_ref())
                .flat_map(|bs| bs.iter().copied())
                .fold(0.0f32, f32::max);
            let min_r = 4.0;
            let max_r = 22.0;

            for series in &c.series {
                set_color(content, series.color);
                if let Some(alpha) = series.fill_alpha {
                    set_fill_alpha(content, alpha, alpha_states);
                }
                if let Some(color) = series.color {
                    stroke_rgb(content, color);
                }
                content.set_line_width(1.0);
                if let (Some(x_vals), Some(bsizes)) = (&series.x_values, &series.bubble_sizes) {
                    for ((&xv, &yv), &bs) in
                        x_vals.iter().zip(series.values.iter()).zip(bsizes.iter())
                    {
                        let px = plot_x + (xv / x_axis_max) * plot_w;
                        let py = plot_y + (yv / axis_max) * plot_h;
                        let r = if max_size > 0.0 {
                            min_r + (bs / max_size).sqrt() * (max_r - min_r)
                        } else {
                            min_r
                        };
                        draw_circle(content, px, py, r);
                        content.fill_nonzero_and_stroke();
                    }
                }
                if series.fill_alpha.is_some() {
                    set_fill_alpha(content, 1.0, alpha_states);
                }
            }
        }
        ChartType::Pie
        | ChartType::Doughnut { .. }
        | ChartType::Radar
        | ChartType::Waterfall { .. }
        | ChartType::Funnel
        | ChartType::Treemap
        | ChartType::Sunburst
        | ChartType::BoxWhisker { .. }
        | ChartType::Pareto => unreachable!(),
    }
    labels
}

/// A value label's baseline above its bar's front face, 2-D and 3-D alike
/// (Word probes: 6.63pt and 6.65pt).
const VALUE_LABEL_RISE: f32 = 6.64;
/// A label beside a bar's front face starts this far past its end, and a
/// label on a bar's centre line sits its baseline this far below it
/// (hbar3d and stacked3d probes: 6.2–6.7pt and 3.45–3.9pt over 8 labels).
const VALUE_LABEL_GAP: f32 = 6.45;
pub(super) const VALUE_LABEL_SINK: f32 = 3.66;

/// A bar's value label, placed by Word's defaults for its bar: above a
/// column's front face, after a bar's end, centred in a stacked segment.
/// ponytail: line, area and pie points do not show their values yet.
struct ValueLabel {
    at: (f32, f32),
    value: f32,
    place: LabelPlace,
}

enum LabelPlace {
    Above,
    After,
    Center,
}

/// A chart title's size and place: torture_test's 18pt title sits on a
/// baseline 24.6pt under the chart's top, in a 30pt band above the plot.
const TITLE_PT: f32 = 18.0;
const TITLE_BASELINE: f32 = 24.6;
const TITLE_BAND: f32 = 30.0;

/// The band a chart's title takes above its plot.
pub(super) fn title_band(c: &Chart) -> f32 {
    if c.title.is_some() { TITLE_BAND } else { 0.0 }
}
/// The plot area's distance from the chart's top and right edges, and the
/// band a one-row legend takes under it. Word, 252pt-high probes: 10.55
/// top (11.0 on case53's 216pt charts, so not a share of the height), 11.0
/// right (10–12.6 on case29/53), legend 23.6.
const CHART_PAD_TOP: f32 = 10.7;
const CHART_PAD_RIGHT: f32 = 11.0;
const LEGEND_BAND: f32 = 23.6;
const LEGEND_PAD_RIGHT: f32 = 9.9;
/// A legend key's width and the space after it.
const LEGEND_KEY_W: f32 = 8.0;
/// A title's rich text without run properties of its own falls back to
/// DrawingML's default face, not the theme's: Word draws it in Arial where
/// every other label takes the minor font (torture_test, probe exports).
pub(super) const TITLE_FONT: &str = "Arial";

/// A data label's text: the value as the number format shows it, or as
/// written when the format is General.
pub(super) fn data_label(val: f32, format_code: Option<&str>) -> String {
    match format_code.filter(|f| *f != "General") {
        Some(fmt) => apply_number_format(val, fmt),
        None => format!("{val}"),
    }
}

/// `c:dTable` under a column chart's plot: the category row, then a row per
/// series with its legend key, name and values, ruled like Word's grid
/// (torture_test at 10pt: 15.1pt rows, text 12.8pt under each row's top).
fn draw_data_table(
    content: &mut Content,
    c: &Chart,
    plot: PlotRect,
    num_categories: usize,
    header_w: f32,
    font_key: &str,
    font_size: f32,
    font: Option<&FontEntry>,
    val_fmt: Option<&str>,
) {
    const KEY: f32 = 5.27;
    const PAD: f32 = 3.4;
    let row_h = table_row_h(font_size);
    let rows = 1 + c.series.len();
    let bottom = plot.y - rows as f32 * row_h;
    let left = plot.x - header_w;
    let right = plot.x + plot.w;
    let col_w = plot.w / num_categories.max(1) as f32;
    let baseline = |row: usize| plot.y - row as f32 * row_h - 1.28 * font_size;
    let centered = |content: &mut Content, ci: usize, row: usize, text: &str| {
        let tw = text_width(text, font_size, font);
        let cx = plot.x + (ci as f32 + 0.5) * col_w;
        let (x, y) = (cx - tw / 2.0, baseline(row));
        show_text(content, font_key, font_size, x, y, text, font);
    };

    content.set_line_width(0.5);
    stroke_rgb(content, DEFAULT_AXIS_COLOR);
    for row in 0..=rows {
        let ly = plot.y - row as f32 * row_h;
        content.move_to(if row == 0 { plot.x } else { left }, ly);
        content.line_to(right, ly);
    }
    for ci in 0..=num_categories {
        let gx = plot.x + ci as f32 * col_w;
        content.move_to(gx, plot.y);
        content.line_to(gx, bottom);
    }
    content.move_to(left, plot.y - row_h);
    content.line_to(left, bottom);
    content.stroke();

    content.set_fill_gray(0.0);
    if let Some(cat_axis) = &c.cat_axis {
        for (ci, label) in cat_axis.labels.iter().enumerate() {
            centered(content, ci, 0, label);
        }
    }
    for (si, series) in c.series.iter().enumerate() {
        let row = si + 1;
        let key_y = plot.y - (row + 1) as f32 * row_h + (row_h - KEY) / 2.0;
        let color = series.color.unwrap_or([0, 0, 0]);
        render_swatch(content, &SwatchStyle::Rect, color, left + PAD, key_y, KEY);
        content.set_fill_gray(0.0);
        let (x, y) = (left + PAD + KEY + 2.3, baseline(row));
        show_text(content, font_key, font_size, x, y, &series.label, font);
        for (ci, val) in series.values.iter().enumerate().take(num_categories) {
            centered(content, ci, row, &data_label(*val, val_fmt));
        }
    }
}

/// A data table row's height at the table's font size.
fn table_row_h(font_size: f32) -> f32 {
    1.51 * font_size
}

pub(super) fn render_chart(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    seen_fonts: &HashMap<String, FontEntry>,
    default_font_name: &str,
    alpha_states: &mut HashSet<u8>,
) {
    let c = &chart.chart;
    let w = chart.display_width;
    let h = chart.display_height;

    let font_size = 10.0;
    let label_font = seen_fonts.get(default_font_name);
    let has_font = label_font.is_some();
    let label_font_key = label_font
        .map(|e| e.pdf_name.as_str())
        .unwrap_or(default_font_name);

    let num_categories = c
        .cat_axis
        .as_ref()
        .map(|ax| ax.labels.len())
        .filter(|&n| n > 0)
        .unwrap_or_else(|| c.series.first().map(|s| s.values.len()).unwrap_or(0));
    let num_series = c.series.len();
    if num_series == 0 {
        return;
    }
    let is_scatter_like = matches!(c.chart_type, ChartType::Scatter | ChartType::Bubble);
    if num_categories == 0 && !is_scatter_like {
        return;
    }

    // The chart area's outline, whatever the chart type.
    if let Some(color) = c.frame {
        content.save_state();
        content.set_line_width(0.75);
        stroke_rgb(content, color);
        content.rect(x, y - h, w, h);
        content.stroke();
        content.restore_state();
    }

    if c.ex_text.is_some() {
        charts_ex::render(chart, content, x, y, label_font_key, label_font);
        return;
    }

    if has_font && let Some(title) = &c.title {
        let (key, font) = match seen_fonts.get(TITLE_FONT) {
            Some(e) => (e.pdf_name.as_str(), Some(e)),
            None => (label_font_key, label_font),
        };
        let tw = text_width(title, TITLE_PT, font);
        let (tx, ty) = (x + (w - tw) / 2.0, y - TITLE_BASELINE);
        show_text(content, key, TITLE_PT, tx, ty, title, font);
    }

    match c.chart_type {
        ChartType::Pie => {
            charts_radial::render_radial(chart, content, x, y, label_font_key, label_font, None);
            return;
        }
        ChartType::Doughnut { hole_size_pct } => {
            charts_radial::render_radial(
                chart,
                content,
                x,
                y,
                label_font_key,
                label_font,
                Some(hole_size_pct),
            );
            return;
        }
        ChartType::Radar => {
            render_radar(chart, content, x, y, has_font, label_font_key, label_font);
            return;
        }
        _ => {}
    }

    let horizontal = matches!(
        c.chart_type,
        ChartType::Bar {
            horizontal: true,
            ..
        }
    );
    let bar_grouping = match c.chart_type {
        ChartType::Bar { grouping, .. } => grouping,
        _ => BarGrouping::Clustered,
    };
    let is_stacked = bar_grouping == BarGrouping::Stacked;
    let is_percent_stacked = bar_grouping == BarGrouping::PercentStacked;
    let depth = c.view_3d;

    let legend_pos = c.legend.as_ref().map(|l| l.position);
    let has_legend = legend_pos.is_some();
    let legend_on_right = legend_pos == Some(LegendPosition::Right);
    let legend_on_bottom = legend_pos == Some(LegendPosition::Bottom);

    let max_val = if is_stacked {
        (0..num_categories)
            .map(|ci| {
                c.series
                    .iter()
                    .map(|s| s.values.get(ci).copied().unwrap_or(0.0))
                    .sum::<f32>()
            })
            .fold(0.0f32, f32::max)
    } else {
        c.series
            .iter()
            .flat_map(|s| s.values.iter())
            .copied()
            .fold(0.0f32, f32::max)
    };

    let (tick_step, axis_max) = if is_percent_stacked {
        (10.0, 100.0)
    } else if matches!(c.chart_type, ChartType::Scatter) {
        compute_axis_range(max_val, 10, 0.98)
    } else if depth.is_some() {
        // ponytail: one probe — Word's 3-D columns end the axis at 25 for a
        // 24 maximum, without the 2-D headroom step.
        compute_axis_range(max_val, 5, f32::INFINITY)
    } else {
        compute_axis_range(max_val, 5, 0.9)
    };
    if axis_max <= 0.0 {
        return;
    }

    let (x_axis_max, x_tick_step) = if is_scatter_like {
        let x_max = c
            .series
            .iter()
            .filter_map(|s| s.x_values.as_ref())
            .flat_map(|xv| xv.iter())
            .copied()
            .fold(0.0f32, f32::max);
        let (xts, xam) = compute_axis_range(x_max, 5, 0.9);
        if xam <= 0.0 {
            return;
        }
        (xam, xts)
    } else {
        (0.0, 0.0)
    };

    let val_fmt = c.val_format_code.as_deref();
    let max_tick_label = if is_percent_stacked {
        "100%".to_string()
    } else {
        format_tick_label(axis_max, tick_step, val_fmt)
    };
    let val_label_w = text_width(&max_tick_label, font_size, label_font) + 15.0;

    let is_point_chart = matches!(
        c.chart_type,
        ChartType::Line | ChartType::Area | ChartType::Scatter | ChartType::Bubble
    );
    // Word's data table runs under the plot: a row of categories, then one
    // per series headed by its legend key and name in a column left of it.
    let table_rows = if c.data_table && !horizontal && !is_scatter_like {
        1 + num_series
    } else {
        0
    };
    let table_header_w = if table_rows > 0 {
        c.series
            .iter()
            .map(|s| text_width(&s.label, font_size, label_font) + 13.5)
            .fold(0.0f32, f32::max)
    } else {
        0.0
    };
    // The axis down the left takes its widest label plus 15pt: values on a
    // column chart, categories on a bar chart (h2d_bar/h2d_bar_long probes).
    let cat_label_w = || {
        c.cat_axis.as_ref().map_or(0.0, |ax| {
            ax.labels
                .iter()
                .map(|l| text_width(l, font_size, label_font))
                .fold(0.0f32, f32::max)
        }) + 15.0
    };
    // A data table's header column (series names behind their keys) starts
    // 17pt in from the chart's left edge and is 13.5pt wider than its
    // widest name (f_2d probe).
    let margin_left = if !horizontal {
        val_label_w
    } else {
        cat_label_w()
    }
    .max(if table_rows > 0 {
        table_header_w + 17.0
    } else {
        0.0
    });
    let legend_label_w = c
        .series
        .iter()
        .map(|s| text_width(&s.label, 10.0, label_font))
        .fold(0.0f32, f32::max);
    // A right legend ends 9.9pt from the chart's edge, its key 8pt before
    // its names (20pt for a line through the marker), and stands 15.6pt
    // clear of the plot (v2d_legend_r probe, case30's line chart). A value
    // axis along the bottom hangs half its last label past the plot.
    let line_key = matches!(c.chart_type, ChartType::Line) && depth.is_none();
    let key_w = LEGEND_KEY_W + if line_key { 12.0 } else { 0.0 };
    let legend_right_x = x + w - LEGEND_PAD_RIGHT - legend_label_w - key_w;
    let x_overhang = if is_scatter_like {
        text_width(
            &format_tick_label(x_axis_max, x_tick_step, None),
            font_size,
            label_font,
        ) / 2.0
    } else {
        0.0
    };
    let margin_right = if has_legend && legend_on_right {
        LEGEND_PAD_RIGHT + legend_label_w + key_w + 15.6 + x_overhang
    } else if is_point_chart {
        (w * 0.06).max(20.0)
    } else if horizontal {
        // The last value label overhangs the plot's right end by half.
        CHART_PAD_RIGHT + (val_label_w - 15.0) / 2.0
    } else {
        CHART_PAD_RIGHT
    };
    let margin_top = CHART_PAD_TOP + title_band(c);

    // Estimate plot_w for label wrapping and thinning calculations
    let plot_w_est = w - margin_left - margin_right;
    let line_h = font_size + 2.0;

    let max_wrap_lines = if !horizontal && !is_scatter_like && num_categories > 0 {
        if let Some(ref cat_axis) = c.cat_axis {
            let slot_w = plot_w_est / num_categories as f32;
            cat_axis
                .labels
                .iter()
                .map(|l| wrap_label(l, slot_w * 0.9, font_size, label_font).len())
                .max()
                .unwrap_or(1)
        } else {
            1
        }
    } else {
        1
    };

    let cat_label_h = line_h * max_wrap_lines as f32 + 4.0;
    let margin_bottom = if table_rows > 0 {
        // Word, f_2d probe: the table ends 29.05pt above the chart's bottom.
        let legend_band = if has_legend && legend_on_bottom {
            28.6
        } else {
            8.0
        };
        legend_band + table_rows as f32 * table_row_h(font_size)
    } else if has_legend && legend_on_bottom {
        cat_label_h + 8.0 + LEGEND_BAND
    } else {
        cat_label_h + 8.0
    };

    let series_names_w = depth
        .is_some_and(|d| d.series_axis && depth_rows(c) > 1)
        .then(|| {
            c.series
                .iter()
                .map(|s| text_width(&s.label, font_size, label_font))
                .fold(0.0f32, f32::max)
        });
    // With rAngAx 0 Word's plot runs to the chart's right pad, short of the
    // series names by their width and 1.2pt (line3d against noser and
    // persp3d probes).
    let perspective = depth.is_some_and(|d| d.perspective.is_some());
    let margin_right = if perspective && !(has_legend && legend_on_right) {
        CHART_PAD_RIGHT + series_names_w.map_or(0.0, |w| w + 1.2)
    } else {
        margin_right
    };
    let plot_2d = PlotRect {
        x: x + margin_left,
        y: y - h + margin_bottom,
        w: w - margin_left - margin_right,
        h: h - margin_top - margin_bottom,
    };
    // Series names along the floor's right edge start 12pt past it.
    let reserve = series_names_w.map_or(0.0, |w| w + 12.0);
    let (plot, floor_depth, space) = match depth {
        Some(d) => {
            let (front, floor_depth, space) = d.space(c, plot_2d, num_categories, reserve);
            (front, floor_depth, Some(space))
        }
        None => (plot_2d, 0.0, None),
    };
    let at = |x: f32, y: f32, z: f32| on_page(space, x, y, z);
    let stroke_world = |content: &mut Content, pts: &[(f32, f32, f32)]| {
        polyline(content, pts.iter().map(|&(x, y, z)| at(x, y, z)));
        content.stroke();
    };
    let PlotRect {
        x: plot_x,
        y: plot_y,
        w: plot_w,
        h: plot_h,
    } = plot;

    content.save_state();

    // Gridlines and value-axis tick marks
    if let Some(ref val_axis) = c.val_axis {
        content.set_line_width(0.5);
        let num_ticks = (axis_max / tick_step).round() as usize;
        let tick_len = 4.0;

        stroke_rgb(
            content,
            val_axis.gridline_color.unwrap_or(DEFAULT_AXIS_COLOR),
        );
        for i in 0..=num_ticks {
            let frac = (i as f32 * tick_step) / axis_max;
            // In 3-D along the left wall (a lying bar's floor) to the back,
            // then across the back.
            let (a, b) = if !horizontal {
                let gy = plot_y + frac * plot_h;
                ((plot_x, gy), (plot_x + plot_w, gy))
            } else {
                let gx = plot_x + frac * plot_w;
                ((gx, plot_y), (gx, plot_y + plot_h))
            };
            let line = [
                (a.0, a.1, 0.0),
                (a.0, a.1, floor_depth),
                (b.0, b.1, floor_depth),
            ];
            if depth.is_some() {
                stroke_world(content, &line);
            } else {
                stroke_world(content, &[line[0], line[2]]);
            }
        }

        stroke_rgb(content, val_axis.line_color.unwrap_or(DEFAULT_AXIS_COLOR));
        for i in 0..=num_ticks {
            let frac = (i as f32 * tick_step) / axis_max;
            if !horizontal {
                let (ax, ay) = at(plot_x, plot_y + frac * plot_h, 0.0);
                content.move_to(ax - tick_len, ay);
                content.line_to(ax, ay);
            } else {
                let gx = plot_x + frac * plot_w;
                content.move_to(gx, plot_y - tick_len);
                content.line_to(gx, plot_y);
            }
            content.stroke();
        }
    }

    // X-axis tick marks for category/scatter axes
    {
        let axis_color = c
            .cat_axis
            .as_ref()
            .and_then(|a| a.line_color)
            .unwrap_or(DEFAULT_AXIS_COLOR);
        stroke_rgb(content, axis_color);
        content.set_line_width(0.5);
        let tick_len = 4.0;
        if is_scatter_like {
            let num_ticks = (x_axis_max / x_tick_step).round() as usize;
            for i in 0..=num_ticks {
                let frac = (i as f32 * x_tick_step) / x_axis_max;
                let gx = plot_x + frac * plot_w;
                content.move_to(gx, plot_y - tick_len);
                content.line_to(gx, plot_y);
                content.stroke();
            }
        } else if !horizontal {
            for ci in 0..=num_categories {
                let gx = plot_x + (ci as f32 / num_categories as f32) * plot_w;
                let (ax, ay) = at(gx, plot_y, 0.0);
                content.move_to(ax, ay - tick_len);
                content.line_to(ax, ay);
                content.stroke();
            }
        }
    }

    // Vertical gridlines for scatter/bubble
    if is_scatter_like && let Some(color) = c.cat_axis.as_ref().and_then(|a| a.gridline_color) {
        content.set_line_width(0.5);
        stroke_rgb(content, color);
        let num_ticks = (x_axis_max / x_tick_step).round() as usize;
        for i in 0..=num_ticks {
            let frac = (i as f32 * x_tick_step) / x_axis_max;
            let gx = plot_x + frac * plot_w;
            content.move_to(gx, plot_y);
            content.line_to(gx, plot_y + plot_h);
            content.stroke();
        }
    }

    let value_labels = draw_chart_series(
        c,
        content,
        plot,
        num_categories,
        axis_max,
        x_axis_max,
        space,
        alpha_states,
    );

    // Plot area border
    content.set_line_width(0.75);
    if let Some(border) = c.plot_border_color {
        stroke_rgb(content, border);
        content.rect(plot_x, plot_y, plot_w, plot_h);
        content.stroke();
    } else {
        let axis_color = c
            .val_axis
            .as_ref()
            .and_then(|a| a.line_color)
            .unwrap_or(DEFAULT_AXIS_COLOR);
        stroke_rgb(content, axis_color);
        // The category wall's free edge recedes from the far end of its axis.
        let top = (plot_x, plot_y + plot_h);
        let end = (plot_x + plot_w, plot_y);
        for (far, recedes) in [(top, horizontal), (end, !horizontal)] {
            let edge = [
                (plot_x, plot_y, 0.0),
                (far.0, far.1, 0.0),
                (far.0, far.1, floor_depth),
            ];
            let n = if recedes && depth.is_some() { 3 } else { 2 };
            stroke_world(content, &edge[..n]);
        }
    }

    // Axis labels
    if has_font {
        content.set_fill_gray(0.0);

        // Value axis tick labels
        let num_ticks = (axis_max / tick_step).round() as usize;
        for i in 0..=num_ticks {
            let val = i as f32 * tick_step;
            let label = if is_percent_stacked {
                format!("{}%", val as i32)
            } else {
                format_tick_label(val, tick_step, val_fmt)
            };
            let tw = text_width(&label, font_size, label_font);
            let frac = val / axis_max;

            if !horizontal {
                let (ax, ay) = at(plot_x, plot_y + frac * plot_h, 0.0);
                let (lx, ly) = (ax - tw - 9.0, ay - font_size * 0.3);
                show_text(
                    content,
                    label_font_key,
                    font_size,
                    lx,
                    ly,
                    &label,
                    label_font,
                );
            } else {
                let lx = plot_x + frac * plot_w - tw / 2.0;
                let ly = plot_y - font_size - 8.0;
                show_text(
                    content,
                    label_font_key,
                    font_size,
                    lx,
                    ly,
                    &label,
                    label_font,
                );
            }
        }

        // X-axis tick labels for scatter/bubble
        if is_scatter_like {
            let num_x_ticks = (x_axis_max / x_tick_step).round() as usize;
            for i in 0..=num_x_ticks {
                let val = i as f32 * x_tick_step;
                let label = format_tick_label(val, x_tick_step, None);
                let tw = text_width(&label, font_size, label_font);
                let frac = val / x_axis_max;
                let lx = plot_x + frac * plot_w - tw / 2.0;
                let ly = plot_y - font_size - 8.0;
                show_text(
                    content,
                    label_font_key,
                    font_size,
                    lx,
                    ly,
                    &label,
                    label_font,
                );
            }
        }

        for value_label in &value_labels {
            let label = data_label(value_label.value, val_fmt);
            let tw = text_width(&label, font_size, label_font);
            let (x, y) = value_label.at;
            let (lx, ly) = match value_label.place {
                LabelPlace::Above => (x - tw / 2.0, y + VALUE_LABEL_RISE),
                LabelPlace::After => (x + VALUE_LABEL_GAP, y - VALUE_LABEL_SINK),
                LabelPlace::Center => (x - tw / 2.0, y - VALUE_LABEL_SINK),
            };
            show_text(
                content,
                label_font_key,
                font_size,
                lx,
                ly,
                &label,
                label_font,
            );
        }

        // Series names beside the floor's right edge, one per depth row
        // (standard3d probe: 12pt right of a row's middle, baseline 3.9pt
        // below it; in perspective 8.25pt right and centred on it as the
        // value labels are, line3d probe).
        if series_names_w.is_some() {
            let (gap, sink) = if perspective {
                (8.25, font_size * 0.3)
            } else {
                (12.0, 3.9)
            };
            let rows = depth_rows(c) as f32;
            let (corner_x, corner_y) = if horizontal {
                (plot_x, plot_y + plot_h)
            } else {
                (plot_x + plot_w, plot_y)
            };
            for (si, series) in c.series.iter().enumerate() {
                let (ax, ay) = at(corner_x, corner_y, (si as f32 + 0.5) / rows * floor_depth);
                let (lx, ly) = (ax + gap, ay - sink);
                show_text(
                    content,
                    label_font_key,
                    font_size,
                    lx,
                    ly,
                    &series.label,
                    label_font,
                );
            }
        }

        if table_rows > 0 {
            draw_data_table(
                content,
                c,
                plot,
                num_categories,
                table_header_w,
                label_font_key,
                font_size,
                label_font,
                val_fmt,
            );
        }

        // Category axis labels with auto-skip and multi-line wrapping
        if !is_scatter_like
            && table_rows == 0
            && let Some(ref cat_axis) = c.cat_axis
        {
            let label_skip = if !horizontal {
                let slot_w = plot_w / num_categories.max(1) as f32;
                // Measure the widest wrapped line to decide thinning.
                // Labels that wrap at spaces are allowed to slightly exceed the slot.
                let can_wrap = cat_axis.labels.iter().any(|l| l.contains(' '));
                let max_wrapped_line_w = cat_axis
                    .labels
                    .iter()
                    .flat_map(|l| wrap_label(l, slot_w, font_size, label_font))
                    .map(|line| text_width(&line, font_size, label_font))
                    .fold(0.0f32, f32::max);
                // Word is tolerant of slight overflow for wrapped labels
                let threshold = if can_wrap { slot_w * 1.5 } else { slot_w };
                if max_wrapped_line_w > threshold {
                    (max_wrapped_line_w / slot_w).ceil() as usize
                } else {
                    1
                }
            } else {
                1
            };

            for (ci, label) in cat_axis.labels.iter().enumerate() {
                if ci % label_skip != 0 {
                    continue;
                }
                if !horizontal {
                    let group_w = plot_w / num_categories as f32;
                    let slot_w = group_w * label_skip as f32;
                    let (cx_center, edge_y) = at(plot_x + (ci as f32 + 0.5) * group_w, plot_y, 0.0);
                    // In perspective the baseline sits 16.6pt under the
                    // front edge's projected middle (line3d probe).
                    let drop = if perspective { 16.6 } else { font_size + 8.0 };
                    let lines = wrap_label(label, slot_w * 0.9, font_size, label_font);
                    for (li, line) in lines.iter().enumerate() {
                        let tw = text_width(line, font_size, label_font);
                        let lx = cx_center - tw / 2.0;
                        let ly = edge_y - drop - (li as f32 * line_h);
                        show_text(content, label_font_key, font_size, lx, ly, line, label_font);
                    }
                } else {
                    let tw = text_width(label, font_size, label_font);
                    let group_h = plot_h / num_categories as f32;
                    let (ax, ay) = at(plot_x, plot_y + ci as f32 * group_h + group_h / 2.0, 0.0);
                    let (cx, cy) = (ax - tw - 9.0, ay - font_size * 0.3);
                    show_text(
                        content,
                        label_font_key,
                        font_size,
                        cx,
                        cy,
                        label,
                        label_font,
                    );
                }
            }
        }

        // Legend
        if let Some(ref legend) = c.legend {
            // A 3-D line's legend shows its ribbon as a plain swatch.
            let is_line = matches!(
                c.chart_type,
                ChartType::Line | ChartType::Scatter | ChartType::Bubble | ChartType::Radar
            ) && depth.is_none();
            let is_bubble = matches!(c.chart_type, ChartType::Bubble);
            // Word lists 2-D stacked columns top series first, but its 3-D
            // stacked columns in series order (stacked3d probe).
            let reverse_legend = match c.chart_type {
                ChartType::Bar {
                    horizontal: true,
                    grouping: BarGrouping::Clustered,
                } => true,
                ChartType::Bar {
                    horizontal: false,
                    grouping: BarGrouping::Stacked | BarGrouping::PercentStacked,
                } => depth.is_none(),
                _ => false,
            };
            let series_order: Vec<(usize, &crate::model::ChartSeries)> = if reverse_legend {
                c.series.iter().enumerate().rev().collect()
            } else {
                c.series.iter().enumerate().collect()
            };
            let items: Vec<LegendItem> = series_order
                .iter()
                .map(|&(si, series)| {
                    let swatch = if line_key {
                        SwatchStyle::LineMarker(resolve_marker(series.marker, si))
                    } else if is_line {
                        SwatchStyle::Marker(if is_bubble {
                            MarkerSymbol::Circle
                        } else {
                            resolve_marker(series.marker, si)
                        })
                    } else {
                        SwatchStyle::Rect
                    };
                    LegendItem {
                        label: &series.label,
                        color: series.color.unwrap_or([0, 0, 0]),
                        swatch,
                    }
                })
                .collect();
            let placement = match legend.position {
                // Centred on the chart below its title; names 17.63pt apart.
                LegendPosition::Right => LegendPlacement::Right {
                    x: legend_right_x,
                    center_y: y - (h + title_band(c)) / 2.0 - 4.67,
                },
                _ => LegendPlacement::bottom(x, y, w, h),
            };
            render_chart_legend(
                content,
                &items,
                placement,
                label_font_key,
                label_font,
                5.5,
                17.63,
            );
        }
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}

fn render_radar(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    has_font: bool,
    label_font_key: &str,
    label_font: Option<&FontEntry>,
) {
    let c = &chart.chart;
    let w = chart.display_width;
    let h = chart.display_height;
    let font_size = 10.0;

    let num_categories = c
        .cat_axis
        .as_ref()
        .map(|ax| ax.labels.len())
        .unwrap_or_else(|| c.series.first().map(|s| s.values.len()).unwrap_or(0));
    if num_categories == 0 || c.series.is_empty() {
        return;
    }

    let legend_pos = c.legend.as_ref().map(|l| l.position);
    let has_legend = legend_pos.is_some();
    let legend_on_right = legend_pos == Some(LegendPosition::Right);
    let legend_on_bottom = legend_pos == Some(LegendPosition::Bottom);

    let max_val = c
        .series
        .iter()
        .flat_map(|s| s.values.iter())
        .copied()
        .fold(0.0f32, f32::max);
    let (tick_step, axis_max) = compute_axis_range(max_val, 4, 0.98);
    if axis_max <= 0.0 {
        return;
    }

    let margin_right = if has_legend && legend_on_right {
        w * 0.22
    } else {
        w * 0.08
    };
    let margin_bottom = if has_legend && legend_on_bottom {
        h * 0.18
    } else {
        h * 0.08
    };
    let margin_top = h * 0.08;
    let margin_left = w * 0.08;

    let avail_w = w - margin_left - margin_right;
    let avail_h = h - margin_top - margin_bottom;
    let radius = avail_w.min(avail_h) / 2.0 - 4.0;
    let cx = x + margin_left + avail_w / 2.0;
    let cy = y - margin_top - avail_h / 2.0;

    let angle_step = 2.0 * std::f32::consts::PI / num_categories as f32;

    // Angle for category i: start at top (pi/2) and go clockwise
    let cat_angle = |ci: usize| -> f32 { std::f32::consts::FRAC_PI_2 - ci as f32 * angle_step };

    content.save_state();

    // Concentric polygon gridlines
    let num_ticks = (axis_max / tick_step).round() as usize;
    let grid_color = c
        .val_axis
        .as_ref()
        .and_then(|a| a.gridline_color)
        .unwrap_or(DEFAULT_AXIS_COLOR);
    content.set_line_width(0.5);
    stroke_rgb(content, grid_color);

    for ti in 1..=num_ticks {
        let r = radius * (ti as f32 / num_ticks as f32);
        for ci in 0..num_categories {
            let a = cat_angle(ci);
            let px = cx + r * a.cos();
            let py = cy + r * a.sin();
            if ci == 0 {
                content.move_to(px, py);
            } else {
                content.line_to(px, py);
            }
        }
        content.close_path();
        content.stroke();
    }

    // Radial spokes
    for ci in 0..num_categories {
        let a = cat_angle(ci);
        content.move_to(cx, cy);
        content.line_to(cx + radius * a.cos(), cy + radius * a.sin());
        content.stroke();
    }

    // Series data polygons
    let accent_colors = resolve_accent_colors(&c.accent_colors);

    for (si, series) in c.series.iter().enumerate() {
        let color = series
            .color
            .unwrap_or(accent_colors[si % accent_colors.len()]);
        content.set_line_width(series.line_width.unwrap_or(DEFAULT_LINE_W));
        stroke_rgb(content, color);
        for ci in 0..num_categories {
            let val = series.values.get(ci).copied().unwrap_or(0.0);
            let r = (val / axis_max) * radius;
            let a = cat_angle(ci);
            let px = cx + r * a.cos();
            let py = cy + r * a.sin();
            if ci == 0 {
                content.move_to(px, py);
            } else {
                content.line_to(px, py);
            }
        }
        content.close_path();
        content.stroke();

        // Markers
        let sym = resolve_marker(series.marker, si);
        let (marker_color, marker_r) = marker_look(series, si, accent_colors, sym, DEFAULT_LINE_W);
        fill_rgb(content, marker_color);
        for ci in 0..num_categories {
            let val = series.values.get(ci).copied().unwrap_or(0.0);
            let r = (val / axis_max) * radius;
            let a = cat_angle(ci);
            draw_marker(content, sym, cx + r * a.cos(), cy + r * a.sin(), marker_r);
        }
    }

    if has_font {
        content.set_fill_gray(0.0);

        // Category labels at perimeter — position based on spoke angle
        if let Some(ref cat_axis) = c.cat_axis {
            let label_r = radius + 4.0;
            for (ci, label) in cat_axis.labels.iter().enumerate() {
                let a = cat_angle(ci);
                let cos_a = a.cos();
                let sin_a = a.sin();
                let lx = cx + label_r * cos_a;
                let ly = cy + label_r * sin_a;
                let tw = text_width(label, font_size, label_font);
                // Horizontal: smoothly right-align on left side, left-align on right
                let tx = lx - tw * (1.0 - cos_a) / 2.0;
                // Vertical: top labels sit above vertex, bottom labels hang below
                let ty = if sin_a > 0.5 {
                    ly + font_size * 0.25
                } else if sin_a < -0.3 {
                    ly - font_size * 0.15
                } else {
                    ly - font_size * 0.3
                };
                show_text(
                    content,
                    label_font_key,
                    font_size,
                    tx,
                    ty,
                    label,
                    label_font,
                );
            }
        }

        // Value labels along the 12 o'clock spoke (including "0" at center)
        let val_gap = 12.0;
        {
            let tw = text_width("0", font_size, label_font);
            show_text(
                content,
                label_font_key,
                font_size,
                cx - tw - val_gap,
                cy - font_size * 0.3,
                "0",
                label_font,
            );
        }
        for ti in 1..=num_ticks {
            let val = ti as f32 * tick_step;
            let label = format_tick_label(val, tick_step, None);
            let r = radius * (ti as f32 / num_ticks as f32);
            let ly = cy + r - font_size * 0.3;
            let tw = text_width(&label, font_size, label_font);
            show_text(
                content,
                label_font_key,
                font_size,
                cx - tw - val_gap,
                ly,
                &label,
                label_font,
            );
        }

        // Legend
        if has_legend {
            let items: Vec<LegendItem> = c
                .series
                .iter()
                .enumerate()
                .map(|(si, series)| {
                    let color = series
                        .color
                        .unwrap_or(accent_colors[si % accent_colors.len()]);
                    LegendItem {
                        label: &series.label,
                        color,
                        swatch: SwatchStyle::LineMarker(resolve_marker(series.marker, si)),
                    }
                })
                .collect();
            let placement = if legend_on_right {
                LegendPlacement::Right {
                    x: x + w - margin_right + 10.0,
                    center_y: cy,
                }
            } else {
                LegendPlacement::bottom(x, y, w, h)
            };
            render_chart_legend(
                content,
                &items,
                placement,
                label_font_key,
                label_font,
                5.5,
                18.0,
            );
        }
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_ceil_nice_zero_and_negative() {
        assert_eq!(ceil_nice(0.0), 1.0);
        assert_eq!(ceil_nice(-5.0), 1.0);
    }

    #[test]
    fn test_ceil_nice_values() {
        assert_eq!(ceil_nice(0.3), 0.2); // mag=0.1, norm=3.0, nice=2.0
        assert_eq!(ceil_nice(0.8), 1.0); // mag=0.1, norm=8.0, nice=10.0
        assert_eq!(ceil_nice(1.5), 1.0); // mag=1.0, norm=1.5, nice=1.0
        assert_eq!(ceil_nice(25.0), 20.0); // mag=10, norm=2.5, nice=2.0
        assert_eq!(ceil_nice(70.0), 50.0); // mag=10, norm=7.0, nice=5.0 (<=7.0)
        assert_eq!(ceil_nice(15.0), 10.0); // mag=10, norm=1.5, nice=1.0
        assert_eq!(ceil_nice(800.0), 1000.0);
        assert_eq!(ceil_nice(3500.0), 5000.0);
    }

    #[test]
    fn test_nice_tick_step() {
        assert_eq!(nice_tick_step(100.0, 5), 20.0);
        assert_eq!(nice_tick_step(1000.0, 5), 200.0);
        assert_eq!(nice_tick_step(80.0, 5), 20.0);
    }

    #[test]
    fn test_compute_axis_range_with_headroom() {
        let (step, max) = compute_axis_range(100.0, 5, 0.9);
        assert_eq!(step, 20.0);
        assert_eq!(max, 120.0);
    }

    #[test]
    fn test_perspective_camera_meets_word_corners() {
        // probe3d line3d: Word's raster of the plot and the box corners found
        // on it, page points from the bottom.
        let plot = PlotRect {
            x: 98.5,
            y: 792.0 - 137.5 - 163.06,
            w: 350.06,
            h: 163.06,
        };
        let front = PlotRect {
            w: plot.w * 4.0 / 3.0,
            ..plot
        };
        let depth = front.w / 2.0;
        let view = View3D {
            rot_x: 15.0,
            rot_y: 20.0,
            depth: 1.0,
            gap_depth: 1.5,
            height: 1.0,
            shape: BarShape::Box,
            series_axis: true,
            perspective: Some(30.0),
        };
        let camera = Camera::fit(plot, front, depth, view);
        let (x1, y1) = (front.x + front.w, front.y + front.h);
        for ((x, y, z), word) in [
            ((front.x, front.y, 0.0), (116.49, 531.94)),
            ((front.x, y1, 0.0), (109.52, 633.22)),
            ((front.x, y1, depth), (182.11, 652.56)),
            ((x1, front.y, 0.0), (411.96, 494.04)),
            ((x1, front.y, depth), (431.32, 543.17)),
        ] {
            let (px, py) = camera.at(x, y, z);
            assert!(
                (px - word.0).hypot(py - word.1) < 2.0,
                "({px}, {py}) vs Word's {word:?}"
            );
        }
    }

    #[test]
    fn test_compute_axis_range_tight() {
        let (step, max) = compute_axis_range(85.0, 10, 0.98);
        assert_eq!(step, 10.0);
        assert_eq!(max, 90.0);
    }

    #[test]
    fn test_format_tick_label_integer_step() {
        assert_eq!(format_tick_label(100.0, 20.0, None), "100");
        assert_eq!(format_tick_label(0.0, 10.0, None), "0");
    }

    #[test]
    fn test_format_tick_label_fractional_step() {
        assert_eq!(format_tick_label(1.5, 0.5, None), "1.5");
    }

    #[test]
    fn test_format_tick_label_with_format_code() {
        assert_eq!(format_tick_label(1234.5, 100.0, Some("#,##0")), "1,234");
        assert_eq!(format_tick_label(0.15, 0.1, Some("0.0%")), "15.0%");
    }

    #[test]
    fn test_apply_number_format_basic() {
        assert_eq!(apply_number_format(1234.56, "0.00"), "1234.56");
        assert_eq!(apply_number_format(1234.0, "0"), "1234");
    }

    #[test]
    fn test_apply_number_format_comma() {
        assert_eq!(apply_number_format(1234567.0, "#,##0"), "1,234,567");
        assert_eq!(apply_number_format(999.0, "#,##0"), "999");
    }

    #[test]
    fn test_apply_number_format_percent() {
        assert_eq!(apply_number_format(0.15, "0%"), "15%");
        assert_eq!(apply_number_format(1.0, "0%"), "100%");
    }

    #[test]
    fn test_insert_thousands_separator() {
        assert_eq!(insert_thousands_separator("1234"), "1,234");
        assert_eq!(insert_thousands_separator("1234567"), "1,234,567");
        assert_eq!(insert_thousands_separator("999"), "999");
        assert_eq!(insert_thousands_separator("1234.56"), "1,234.56");
        assert_eq!(insert_thousands_separator("-1234"), "-1,234");
    }

    #[test]
    fn test_text_width_approx() {
        assert_eq!(text_width_approx("Hello", 12.0), 30.0);
        assert_eq!(text_width_approx("", 12.0), 0.0);
        assert_eq!(text_width_approx("A", 10.0), 5.0);
    }
}
