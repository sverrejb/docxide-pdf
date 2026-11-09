use std::f32::consts::{FRAC_PI_2, TAU};

use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::model::{InlineChart, LegendPosition, View3D};

use super::chart_legend::{LegendItem, LegendPlacement, SwatchStyle, render_chart_legend};
use super::charts::{
    VALUE_LABEL_SINK, data_label, fill_strips, polygon, resolve_accent_colors, show_text,
    text_width, title_band,
};
use super::color::fill_rgb;

struct RadialLayout<'a> {
    cx: f32,
    cy: f32,
    radius: f32,
    legend_x: f32,
    colors: &'a [[u8; 3]],
    has_legend: bool,
    legend_on_right: bool,
    labels: &'a [String],
    /// The series sum each slice's sweep is a share of.
    total: f32,
}

const ARC_SEGMENTS: usize = 64;

pub(super) fn arc_segment_count(sweep: f32) -> usize {
    ((ARC_SEGMENTS as f32 * sweep / TAU).ceil() as usize).max(2)
}

/// `n` line segments from `start` clockwise through `sweep` (a negative
/// sweep runs the arc back anticlockwise).
pub(super) fn emit_arc(
    content: &mut Content,
    cx: f32,
    cy: f32,
    radius: f32,
    start: f32,
    sweep: f32,
    n: usize,
) {
    for s in 0..=n {
        let a = start - (s as f32 / n as f32) * sweep;
        content.line_to(cx + radius * a.cos(), cy + radius * a.sin());
    }
}

fn setup_radial_chart<'a>(
    chart: &'a InlineChart,
    x: f32,
    y: f32,
    label_font: Option<&FontEntry>,
) -> Option<RadialLayout<'a>> {
    let c = &chart.chart;
    let w = chart.display_width;
    let h = chart.display_height;

    let series = c.series.first()?;
    let total: f32 = series.values.iter().sum();
    if total <= 0.0 {
        return None;
    }

    let has_legend = c.legend.is_some();
    let legend_on_right = c
        .legend
        .as_ref()
        .is_some_and(|l| l.position == LegendPosition::Right);

    let labels: &[String] = c
        .cat_axis
        .as_ref()
        .map(|ax| ax.labels.as_slice())
        .unwrap_or_default();

    let colors = resolve_accent_colors(&c.accent_colors);

    // Word's pie keeps 11pt from the chart's edges (a 30pt title band on
    // top) and centres in what the legend leaves: a right legend ends 9.9pt
    // from the edge, its key 8pt before its names, 15.6pt clear of the pie's
    // area (case30, case31).
    let max_label_w = labels
        .iter()
        .map(|l| text_width(l, 10.0, label_font))
        .fold(0.0f32, f32::max);
    let legend_x = x + w - 9.9 - max_label_w - 8.0;
    let left = x + 11.0;
    let right = if legend_on_right {
        legend_x - 15.6
    } else {
        x + w - 11.0
    };
    let top = y - 11.0 - title_band(c);
    let bottom = y - h
        + 11.0
        + if has_legend && !legend_on_right {
            23.6
        } else {
            0.0
        };
    let (cx, cy) = ((left + right) / 2.0, (top + bottom) / 2.0);
    let radius = ((right - left).min(top - bottom) / 2.0).max(0.0);

    Some(RadialLayout {
        cx,
        cy,
        radius,
        legend_x,
        colors,
        has_legend,
        legend_on_right,
        labels,
        total,
    })
}

fn render_radial_legend(
    content: &mut Content,
    chart: &InlineChart,
    layout: &RadialLayout,
    (x, y): (f32, f32),
    label_font_key: &str,
    label_font: Option<&FontEntry>,
) {
    if !layout.has_legend || layout.labels.is_empty() {
        return;
    }
    let (w, h) = (chart.display_width, chart.display_height);
    let items: Vec<LegendItem> = layout
        .labels
        .iter()
        .enumerate()
        .map(|(i, label)| LegendItem {
            label: label.as_str(),
            color: layout.colors[i % layout.colors.len()],
            swatch: SwatchStyle::Rect,
        })
        .collect();
    let placement = if layout.legend_on_right {
        // As for other charts: centred below the title, names on a baseline
        // 11.9pt above the bottom when the legend runs along it.
        LegendPlacement::Right {
            x: layout.legend_x,
            center_y: y - (h + title_band(&chart.chart)) / 2.0 - 4.67,
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
        5.274,
        17.6,
    );
}

/// A pie, or with `hole_pct` (the doughnut hole as a percentage of the
/// radius) a doughnut.
pub(super) fn render_radial(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    label_font_key: &str,
    label_font: Option<&FontEntry>,
    hole_pct: Option<f32>,
) {
    let Some(layout) = setup_radial_chart(chart, x, y, label_font) else {
        return;
    };
    let c = &chart.chart;

    content.save_state();
    if let Some(view) = c.view_3d.filter(|_| hole_pct.is_none()) {
        draw_pie_3d(
            content,
            chart,
            &layout,
            view,
            x,
            y,
            label_font_key,
            label_font,
        );
    } else {
        draw_pie_2d(content, chart, &layout, hole_pct);
    }
    if label_font.is_some() {
        render_radial_legend(content, chart, &layout, (x, y), label_font_key, label_font);
    }
    content.set_fill_gray(0.0);
    content.restore_state();
}

fn draw_pie_2d(
    content: &mut Content,
    chart: &InlineChart,
    layout: &RadialLayout,
    hole_pct: Option<f32>,
) {
    let c = &chart.chart;
    let outer_r = layout.radius;
    let inner_r = hole_pct.map(|pct| outer_r * (pct / 100.0));

    let mut angle = FRAC_PI_2;
    let values = &c.series[0].values;
    let total = layout.total;

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * TAU;
        let n = arc_segment_count(sweep);
        fill_rgb(content, layout.colors[i % layout.colors.len()]);

        if inner_r.is_some() {
            content.move_to(
                layout.cx + outer_r * angle.cos(),
                layout.cy + outer_r * angle.sin(),
            );
        } else {
            content.move_to(layout.cx, layout.cy);
        }
        // Outer arc clockwise (decreasing angle), then the inner arc back.
        emit_arc(content, layout.cx, layout.cy, outer_r, angle, sweep, n);
        if let Some(inner_r) = inner_r {
            emit_arc(
                content,
                layout.cx,
                layout.cy,
                inner_r,
                angle - sweep,
                -sweep,
                n,
            );
        }
        content.close_path();
        content.fill_nonzero();

        angle -= sweep;
    }
}

/// A 3-D pie seen from `rotX` above: the top an ellipse `sin rotX` as tall as
/// it is wide, standing on a rim 0.261 of the radius thick (times
/// `hPercent`) foreshortened by `cos rotX`, the whole fitted between Word's
/// insets (pie3d probes: 54pt below the chart's top with a title, 49.6pt
/// above its bottom with a legend there). Top faces take the slice colour,
/// the rim darkens left to right from 0.87× to 0.40× at 72% of the width.
/// ponytail: orthographic; Word's default 30° perspective draws the pie
/// ~4% larger with a fuller near half. Value labels sit at 70% of the
/// radius on each slice's middle, where Word's best fit varies by slice.
#[allow(clippy::too_many_arguments)]
fn draw_pie_3d(
    content: &mut Content,
    chart: &InlineChart,
    layout: &RadialLayout,
    view: View3D,
    x: f32,
    y: f32,
    label_font_key: &str,
    label_font: Option<&FontEntry>,
) {
    let c = &chart.chart;
    let (w, h) = (chart.display_width, chart.display_height);
    let values = &c.series[0].values;
    let total = layout.total;
    let (sin_x, cos_x) = (view.rot_x.to_radians().sin(), view.rot_x.to_radians().cos());
    let rim_k = 0.261 * view.height * cos_x;
    let top = y - 24.05 - title_band(c);
    let bottom = y - h
        + if layout.has_legend && !layout.legend_on_right {
            49.6
        } else {
            24.05
        };
    let avail_w = (layout.legend_x - x) - 2.0 * w * 0.05;
    let a = (avail_w / 2.0).min((top - bottom) / (2.0 * sin_x + rim_k));
    let (b, rim) = (a * sin_x, a * rim_k);
    let cx = layout.cx;
    let cy = (top + bottom) / 2.0 + (rim / 2.0);
    // A point on the top ellipse, `t` radians clockwise from 12 o'clock.
    let at = |t: f32, r: f32| (cx + r * a * t.sin(), cy + r * b * t.cos());
    let slices: Vec<(f32, f32)> = values
        .iter()
        .scan(0.0f32, |start, &v| {
            let sweep = v / total * TAU;
            let s = *start;
            *start += sweep;
            Some((s, s + sweep))
        })
        .collect();
    let color = |i: usize| layout.colors[i % layout.colors.len()];

    // The rim: the front half of the outline (from 3 o'clock round through
    // 6 to 9), per slice, shaded in vertical strips.
    for (i, &(t0, t1)) in slices.iter().enumerate() {
        let (lo, hi) = (t0.max(FRAC_PI_2), t1.min(3.0 * FRAC_PI_2));
        if lo >= hi {
            continue;
        }
        let n = arc_segment_count(hi - lo);
        let edge: Vec<(f32, f32)> = (0..=n)
            .map(|k| at(lo + (hi - lo) * k as f32 / n as f32, 1.0))
            .collect();
        let mut band = edge.clone();
        band.extend(edge.iter().rev().map(|&(px, py)| (px, py - rim)));
        let strip_box = (cx - a, cy - b - rim - 1.0, cx + a, cy + b + 1.0);
        fill_strips(content, &band, strip_box, 40, color(i), |t| {
            (0.87 - 0.655 * t).max(0.40)
        });
    }

    for (i, &(t0, t1)) in slices.iter().enumerate() {
        fill_rgb(content, color(i));
        let n = arc_segment_count(t1 - t0);
        let fan: Vec<(f32, f32)> = std::iter::once((cx, cy))
            .chain((0..=n).map(|k| at(t0 + (t1 - t0) * k as f32 / n as f32, 1.0)))
            .collect();
        polygon(content, &fan);
        content.fill_nonzero();
    }

    if c.show_values && label_font.is_some() {
        content.set_fill_gray(0.0);
        for (&(t0, t1), &v) in slices.iter().zip(values) {
            let text = data_label(v, None);
            let tw = text_width(&text, 10.0, label_font);
            let (px, py) = at((t0 + t1) / 2.0, 0.7);
            show_text(
                content,
                label_font_key,
                10.0,
                px - tw / 2.0,
                py - VALUE_LABEL_SINK,
                &text,
                label_font,
            );
        }
    }
}
