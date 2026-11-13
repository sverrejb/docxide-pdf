//! Radar charts: spokes, rings, series polygons and their labels.

use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::model::{InlineChart, LegendPosition};

use super::chart_legend::{LegendItem, LegendPlacement, SwatchStyle, render_chart_legend};
use super::charts::{
    DEFAULT_AXIS_COLOR, DEFAULT_LINE_W, draw_marker, marker_look, polygon, resolve_accent_colors,
    resolve_marker, show_text,
};
use super::charts_format::{compute_axis_range, format_tick_label, text_width};
use super::color::{fill_rgb, stroke_rgb};

pub(super) fn render_radar(
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

    let margin_right = if legend_on_right { w * 0.22 } else { w * 0.08 };
    let margin_bottom = if legend_on_bottom { h * 0.18 } else { h * 0.08 };
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
    // The point `r` out along category `ci`'s spoke.
    let at = |ci: usize, r: f32| {
        let a = cat_angle(ci);
        (cx + r * a.cos(), cy + r * a.sin())
    };

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
        let ring: Vec<(f32, f32)> = (0..num_categories).map(|ci| at(ci, r)).collect();
        polygon(content, &ring);
        content.stroke();
    }

    // Radial spokes
    for ci in 0..num_categories {
        content.move_to(cx, cy);
        let (px, py) = at(ci, radius);
        content.line_to(px, py);
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
        let pts: Vec<(f32, f32)> = (0..num_categories)
            .map(|ci| {
                let val = series.values.get(ci).copied().unwrap_or(0.0);
                at(ci, (val / axis_max) * radius)
            })
            .collect();
        polygon(content, &pts);
        content.stroke();

        // Markers
        let sym = resolve_marker(series.marker, si);
        let (marker_color, marker_r) = marker_look(series, si, accent_colors, sym, DEFAULT_LINE_W);
        fill_rgb(content, marker_color);
        for &(px, py) in &pts {
            draw_marker(content, sym, px, py, marker_r);
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
