//! 3-D charts: the oblique and perspective projections, 3-D bars and ribbons.

use pdf_writer::Content;

use crate::model::{BarGrouping, BarShape, Chart, ChartType, View3D};

use super::charts::{PlotRect, bar_slot, polygon};
use super::color::fill_rgb;

/// A 3-D column chart's depth axis. With right-angle axes Word keeps the
/// front face a rectangle and draws a point `z` deep at (z sin rotY,
/// z sin rotX) from where it would sit on the front plane (five Word probes
/// varying rotX, rotY, depthPercent and the chart width).
/// Depth rows on a 3-D chart's floor: one per series when they stand one
/// behind another (3-D lines and areas always do), else one.
/// ponytail: stacked 3-D areas still get a row per series.
pub(super) fn depth_rows(c: &Chart) -> usize {
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
    pub(super) fn row_pitch(self, thick: f32) -> f32 {
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
    pub(super) fn space(
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
pub(super) struct Space {
    pub(super) view: View3D,
    camera: Option<Camera>,
}

impl Space {
    pub(super) fn at(self, x: f32, y: f32, z: f32) -> (f32, f32) {
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
pub(super) fn on_page(space: Option<Space>, x: f32, y: f32, z: f32) -> (f32, f32) {
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
pub(super) fn draw_ribbons(
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
pub(super) fn draw_bar_3d(
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
