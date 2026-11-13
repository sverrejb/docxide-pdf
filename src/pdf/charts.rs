use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, to_winansi_bytes};
use crate::model::{
    BarGrouping, BarShape, Chart, ChartType, InlineChart, LegendPosition, MarkerSymbol,
};

use super::chart_legend::{
    LegendItem, LegendPlacement, SwatchStyle, render_chart_legend, render_swatch,
};
use super::charts_3d::{Space, depth_rows, draw_bar_3d, draw_ribbons, on_page};
use super::charts_ex;
use super::charts_format::{
    apply_number_format, compute_axis_range, format_tick_label, text_width, wrap_label,
};
use super::charts_radar::render_radar;
use super::charts_radial;
use super::color::{fill_rgb, set_fill_alpha, stroke_rgb};
use super::helpers::draw_circle;

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

pub(super) static DEFAULT_AXIS_COLOR: [u8; 3] = [179, 179, 179];

const DEFAULT_MARKER_CYCLE: [MarkerSymbol; 3] = [
    MarkerSymbol::Diamond,
    MarkerSymbol::Square,
    MarkerSymbol::Triangle,
];

pub(super) fn resolve_marker(marker: Option<MarkerSymbol>, series_index: usize) -> MarkerSymbol {
    marker.unwrap_or(DEFAULT_MARKER_CYCLE[series_index % DEFAULT_MARKER_CYCLE.len()])
}

/// A series' marker as Word draws it: `c:size` (default 5) wide plus its
/// outline, which is the marker's line, else the series' line, else
/// `default_lw` wide; filled with the marker's own colour, else the theme
/// accent for the series (case30's orange line carries accent2's red
/// squares, 7.25pt with the default 2.25pt line). Returns the fill and the
/// radius `draw_marker` takes for that symbol.
pub(super) fn marker_look(
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
pub(super) const DEFAULT_LINE_W: f32 = 2.25;

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

/// A category's slot along the category axis and the width of one bar in it.
pub(super) fn bar_slot(c: &Chart, cat_extent: f32, num_categories: usize) -> (f32, f32) {
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

pub(super) fn polygon(content: &mut Content, pts: &[(f32, f32)]) {
    if !pts.is_empty() {
        polyline(content, pts.iter().copied());
        content.close_path();
    }
}

/// An open path through `pts`.
pub(super) fn polyline(content: &mut Content, pts: impl IntoIterator<Item = (f32, f32)>) {
    let mut pts = pts.into_iter();
    if let Some((x, y)) = pts.next() {
        content.move_to(x, y);
        for (x, y) in pts {
            content.line_to(x, y);
        }
    }
}

#[derive(Clone, Copy)]
pub(super) struct PlotRect {
    pub(super) x: f32,
    pub(super) y: f32,
    pub(super) w: f32,
    pub(super) h: f32,
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
                for &(lx, ly) in &pts {
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
                polyline(content, pts.iter().copied());
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
    let max_series_label_w = c
        .series
        .iter()
        .map(|s| text_width(&s.label, font_size, label_font))
        .fold(0.0f32, f32::max);
    let table_header_w = if table_rows > 0 {
        max_series_label_w + 13.5
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
    // A right legend ends 9.9pt from the chart's edge, its key 8pt before
    // its names (20pt for a line through the marker), and stands 15.6pt
    // clear of the plot (v2d_legend_r probe, case30's line chart). A value
    // axis along the bottom hangs half its last label past the plot.
    let line_key = matches!(c.chart_type, ChartType::Line) && depth.is_none();
    let key_w = LEGEND_KEY_W + if line_key { 12.0 } else { 0.0 };
    let legend_right_x = x + w - LEGEND_PAD_RIGHT - max_series_label_w - key_w;
    let x_overhang = if is_scatter_like {
        text_width(
            &format_tick_label(x_axis_max, x_tick_step, None),
            font_size,
            label_font,
        ) / 2.0
    } else {
        0.0
    };
    let margin_right = if legend_on_right {
        LEGEND_PAD_RIGHT + max_series_label_w + key_w + 15.6 + x_overhang
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
        let legend_band = if legend_on_bottom { 28.6 } else { 8.0 };
        legend_band + table_rows as f32 * table_row_h(font_size)
    } else if legend_on_bottom {
        cat_label_h + 8.0 + LEGEND_BAND
    } else {
        cat_label_h + 8.0
    };

    let series_names_w = depth
        .is_some_and(|d| d.series_axis && depth_rows(c) > 1)
        .then_some(max_series_label_w);
    // With rAngAx 0 Word's plot runs to the chart's right pad, short of the
    // series names by their width and 1.2pt (line3d against noser and
    // persp3d probes).
    let perspective = depth.is_some_and(|d| d.perspective.is_some());
    let margin_right = if perspective && !legend_on_right {
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

            let (lx, ly) = if !horizontal {
                let (ax, ay) = at(plot_x, plot_y + frac * plot_h, 0.0);
                (ax - tw - 9.0, ay - font_size * 0.3)
            } else {
                (plot_x + frac * plot_w - tw / 2.0, plot_y - font_size - 8.0)
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
