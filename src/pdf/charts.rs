use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, encode_as_gids, to_winansi_bytes};
use crate::model::{ChartType, InlineChart, LegendPosition, MarkerSymbol};

use super::chart_legend::{LegendItem, LegendPlacement, SwatchStyle, render_chart_legend};
use super::charts_radial;

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

pub(super) fn text_width_approx(text: &str, font_size: f32) -> f32 {
    text.len() as f32 * font_size * 0.5
}

pub(super) fn text_width(text: &str, font_size: f32, font: Option<&FontEntry>) -> f32 {
    match font {
        Some(f) => f.word_width(text, font_size, false),
        None => text_width_approx(text, font_size),
    }
}

fn format_tick_label(val: f32, step: f32) -> String {
    if step.fract() == 0.0 {
        format!("{}", val as i32)
    } else {
        format!("{:.1}", val)
    }
}

pub(super) fn fill_rgb(content: &mut Content, color: [u8; 3]) {
    content.set_fill_rgb(
        color[0] as f32 / 255.0,
        color[1] as f32 / 255.0,
        color[2] as f32 / 255.0,
    );
}

pub(super) fn stroke_rgb(content: &mut Content, color: [u8; 3]) {
    content.set_stroke_rgb(
        color[0] as f32 / 255.0,
        color[1] as f32 / 255.0,
        color[2] as f32 / 255.0,
    );
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

const DEFAULT_MARKER_CYCLE: [MarkerSymbol; 3] = [
    MarkerSymbol::Diamond,
    MarkerSymbol::Square,
    MarkerSymbol::Triangle,
];

fn resolve_marker(marker: Option<MarkerSymbol>, series_index: usize) -> MarkerSymbol {
    marker.unwrap_or(DEFAULT_MARKER_CYCLE[series_index % DEFAULT_MARKER_CYCLE.len()])
}

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

fn draw_circle(content: &mut Content, cx: f32, cy: f32, r: f32) {
    // Approximate circle with 4 cubic Bezier curves
    let k = r * 0.5522848;
    content.move_to(cx + r, cy);
    content.cubic_to(cx + r, cy + k, cx + k, cy + r, cx, cy + r);
    content.cubic_to(cx - k, cy + r, cx - r, cy + k, cx - r, cy);
    content.cubic_to(cx - r, cy - k, cx - k, cy - r, cx, cy - r);
    content.cubic_to(cx + k, cy - r, cx + r, cy - k, cx + r, cy);
    content.close_path();
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
) {
    show_text_encoded(content, font_key, font_size, x, y, text, None);
}

pub(super) fn show_text_encoded(
    content: &mut Content,
    font_key: &str,
    font_size: f32,
    x: f32,
    y: f32,
    text: &str,
    font_entry: Option<&FontEntry>,
) {
    let bytes = match font_entry.and_then(|e| e.char_to_gid.as_ref()) {
        Some(map) => encode_as_gids(text, map),
        None => to_winansi_bytes(text),
    };
    content
        .begin_text()
        .set_font(Name(font_key.as_bytes()), font_size)
        .next_line(x, y)
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
    let label_font_key = seen_fonts
        .keys()
        .find(|k| {
            let lower = k.to_lowercase();
            !lower.contains("symbol") && !lower.contains("serif") && !lower.contains("/")
        })
        .map(|s| s.as_str())
        .unwrap_or(default_font_name);
    let has_font = seen_fonts.contains_key(label_font_key);
    let label_font = seen_fonts.get(label_font_key);

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

    match c.chart_type {
        ChartType::Pie => {
            charts_radial::render_pie(chart, content, x, y, has_font, label_font_key, label_font);
            return;
        }
        ChartType::Doughnut { hole_size_pct } => {
            charts_radial::render_doughnut(
                chart,
                content,
                x,
                y,
                has_font,
                label_font_key,
                hole_size_pct,
                label_font,
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

    let has_legend = c.legend.is_some();
    let legend_on_right = c
        .legend
        .as_ref()
        .is_some_and(|l| l.position == LegendPosition::Right);
    let legend_on_bottom = c
        .legend
        .as_ref()
        .is_some_and(|l| l.position == LegendPosition::Bottom);

    let max_val = c
        .series
        .iter()
        .flat_map(|s| s.values.iter())
        .copied()
        .fold(0.0f32, f32::max);

    let (tick_step, axis_max) = if matches!(c.chart_type, ChartType::Scatter) {
        // Scatter Y-axis: target ~10 ticks, only add headroom if data hits ceiling
        compute_axis_range(max_val, 10, 0.98)
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

    let max_tick_label = format_tick_label(axis_max, tick_step);
    let val_label_w = text_width(&max_tick_label, font_size, label_font) + 15.0;
    let cat_label_h = font_size + 6.0;

    let is_point_chart = matches!(
        c.chart_type,
        ChartType::Line | ChartType::Area | ChartType::Scatter | ChartType::Bubble
    );
    let margin_left = if !horizontal { val_label_w } else { w * 0.12 };
    let margin_right = if has_legend && legend_on_right {
        let legend_swatch = 5.5;
        let legend_gap = 8.0;
        let max_label_w = c
            .series
            .iter()
            .map(|s| text_width(&s.label, 10.0, label_font))
            .fold(0.0f32, f32::max);
        let computed = legend_gap + legend_swatch + 4.0 + max_label_w + 8.0;
        computed.max(w * 0.12)
    } else if is_point_chart {
        (w * 0.06).max(20.0)
    } else {
        8.0
    };
    let margin_top = h * 0.05;
    let margin_bottom = if has_legend && legend_on_bottom {
        h * 0.22
    } else {
        cat_label_h + 8.0
    };

    let plot_x = x + margin_left;
    let plot_y = y - h + margin_bottom;
    let plot_w = w - margin_left - margin_right;
    let plot_h = h - margin_top - margin_bottom;

    content.save_state();

    // Gridlines
    if let Some(ref val_axis) = c.val_axis {
        let color = val_axis.gridline_color.unwrap_or([179, 179, 179]);
        content.set_line_width(0.5);
        stroke_rgb(content, color);
        let num_ticks = (axis_max / tick_step).round() as usize;
        for i in 0..=num_ticks {
            let frac = (i as f32 * tick_step) / axis_max;
            if !horizontal {
                let gy = plot_y + frac * plot_h;
                content.move_to(plot_x, gy);
                content.line_to(plot_x + plot_w, gy);
            } else {
                let gx = plot_x + frac * plot_w;
                content.move_to(gx, plot_y);
                content.line_to(gx, plot_y + plot_h);
            }
            content.stroke();
        }
    }

    // Y-axis tick marks
    if let Some(ref val_axis) = c.val_axis {
        let color = val_axis.line_color.unwrap_or([179, 179, 179]);
        stroke_rgb(content, color);
        content.set_line_width(0.5);
        let num_ticks = (axis_max / tick_step).round() as usize;
        let tick_len = 4.0;
        for i in 0..=num_ticks {
            let frac = (i as f32 * tick_step) / axis_max;
            if !horizontal {
                let gy = plot_y + frac * plot_h;
                content.move_to(plot_x - tick_len, gy);
                content.line_to(plot_x, gy);
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
            .unwrap_or([179, 179, 179]);
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
                content.move_to(gx, plot_y - tick_len);
                content.line_to(gx, plot_y);
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

    // Data rendering
    match c.chart_type {
        ChartType::Bar {
            horizontal: false, ..
        } => {
            let gap_ratio = c.gap_width_pct / 100.0;
            let group_w = plot_w / num_categories as f32;
            let bar_w = group_w / (num_series as f32 + gap_ratio);
            let gap = gap_ratio * bar_w;

            for ci in 0..num_categories {
                let group_x = plot_x + ci as f32 * group_w + gap / 2.0;
                for (si, series) in c.series.iter().enumerate() {
                    let val = series.values.get(ci).copied().unwrap_or(0.0);
                    let bar_h = (val / axis_max) * plot_h;
                    let bx = group_x + si as f32 * bar_w;
                    set_color(content, series.color);
                    content.rect(bx, plot_y, bar_w, bar_h);
                    content.fill_nonzero();
                }
            }
        }
        ChartType::Bar {
            horizontal: true, ..
        } => {
            let gap_ratio = c.gap_width_pct / 100.0;
            let group_h = plot_h / num_categories as f32;
            let bar_h = group_h / (num_series as f32 + gap_ratio);
            let gap = gap_ratio * bar_h;

            for ci in 0..num_categories {
                let group_y = plot_y + ci as f32 * group_h + gap / 2.0;
                for (si, series) in c.series.iter().enumerate() {
                    let val = series.values.get(ci).copied().unwrap_or(0.0);
                    let bw = (val / axis_max) * plot_w;
                    let by = group_y + si as f32 * bar_h;
                    set_color(content, series.color);
                    content.rect(plot_x, by, bw, bar_h);
                    content.fill_nonzero();
                }
            }
        }
        ChartType::Line => {
            content.set_line_width(2.0);
            for (si, series) in c.series.iter().enumerate() {
                set_stroke_color(content, series.color);
                let cat_w = plot_w / (num_categories - 1).max(1) as f32;
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

                set_color(content, series.color);
                let marker_r = 3.5;
                let sym = resolve_marker(series.marker, si);
                for (ci, &val) in series.values.iter().enumerate() {
                    let lx = plot_x + ci as f32 * cat_w;
                    let ly = plot_y + (val / axis_max) * plot_h;
                    draw_marker(content, sym, lx, ly, marker_r);
                }
            }
        }
        ChartType::Area => {
            let cat_w = plot_w / (num_categories - 1).max(1) as f32;
            for series in &c.series {
                set_color(content, series.color);
                content.move_to(plot_x, plot_y);
                for (ci, &val) in series.values.iter().enumerate() {
                    let lx = plot_x + ci as f32 * cat_w;
                    let ly = plot_y + (val / axis_max) * plot_h;
                    content.line_to(lx, ly);
                }
                let last_x = plot_x + (num_categories - 1) as f32 * cat_w;
                content.line_to(last_x, plot_y);
                content.close_path();
                content.fill_nonzero();

                set_stroke_color(content, series.color);
                content.set_line_width(1.5);
                for (ci, &val) in series.values.iter().enumerate() {
                    let lx = plot_x + ci as f32 * cat_w;
                    let ly = plot_y + (val / axis_max) * plot_h;
                    if ci == 0 {
                        content.move_to(lx, ly);
                    } else {
                        content.line_to(lx, ly);
                    }
                }
                content.stroke();
            }
        }
        ChartType::Scatter => {
            for (si, series) in c.series.iter().enumerate() {
                set_color(content, series.color);
                let marker_r = 4.0;
                let sym = resolve_marker(series.marker, si);
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
                    let pct = (alpha * 100.0).round() as u8;
                    let gs_name = format!("GSa{pct}");
                    content.set_parameters(Name(gs_name.as_bytes()));
                    alpha_states.insert(pct);
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
                    content.set_parameters(Name(b"GSa100"));
                    alpha_states.insert(100);
                }
            }
        }
        ChartType::Pie | ChartType::Doughnut { .. } | ChartType::Radar => unreachable!(),
    }

    // Plot area border
    if let Some(border) = c.plot_border_color {
        content.set_line_width(0.75);
        stroke_rgb(content, border);
        content.rect(plot_x, plot_y, plot_w, plot_h);
        content.stroke();
    } else {
        let axis_color = c
            .val_axis
            .as_ref()
            .and_then(|a| a.line_color)
            .unwrap_or([179, 179, 179]);
        content.set_line_width(0.75);
        stroke_rgb(content, axis_color);
        content.move_to(plot_x, plot_y);
        content.line_to(plot_x, plot_y + plot_h);
        content.stroke();
        content.move_to(plot_x, plot_y);
        content.line_to(plot_x + plot_w, plot_y);
        content.stroke();
    }

    // Axis labels
    if has_font {
        content.set_fill_gray(0.0);

        // Value axis tick labels
        let num_ticks = (axis_max / tick_step).round() as usize;
        for i in 0..=num_ticks {
            let val = i as f32 * tick_step;
            let label = format_tick_label(val, tick_step);
            let tw = text_width(&label, font_size, label_font);
            let frac = val / axis_max;

            if !horizontal {
                let ly = plot_y + frac * plot_h - font_size * 0.3;
                let lx = plot_x - tw - 9.0;
                show_text(content, label_font_key, font_size, lx, ly, &label);
            } else {
                let lx = plot_x + frac * plot_w - tw / 2.0;
                let ly = plot_y - font_size - 8.0;
                show_text(content, label_font_key, font_size, lx, ly, &label);
            }
        }

        // X-axis tick labels for scatter/bubble
        if is_scatter_like {
            let num_x_ticks = (x_axis_max / x_tick_step).round() as usize;
            for i in 0..=num_x_ticks {
                let val = i as f32 * x_tick_step;
                let label = format_tick_label(val, x_tick_step);
                let tw = text_width(&label, font_size, label_font);
                let frac = val / x_axis_max;
                let lx = plot_x + frac * plot_w - tw / 2.0;
                let ly = plot_y - font_size - 8.0;
                show_text(content, label_font_key, font_size, lx, ly, &label);
            }
        }

        // Category axis labels
        let is_point_chart = matches!(c.chart_type, ChartType::Line | ChartType::Area);
        if !is_scatter_like && let Some(ref cat_axis) = c.cat_axis {
            for (ci, label) in cat_axis.labels.iter().enumerate() {
                let tw = text_width(label, font_size, label_font);
                if !horizontal {
                    let cx = if is_point_chart && num_categories > 1 {
                        let cat_w = plot_w / (num_categories - 1) as f32;
                        plot_x + ci as f32 * cat_w - tw / 2.0
                    } else {
                        let group_w = plot_w / num_categories as f32;
                        plot_x + ci as f32 * group_w + group_w / 2.0 - tw / 2.0
                    };
                    let cy = plot_y - font_size - 8.0;
                    show_text(content, label_font_key, font_size, cx, cy, label);
                } else {
                    let group_h = plot_h / num_categories as f32;
                    let cy = plot_y + ci as f32 * group_h + group_h / 2.0 - font_size * 0.3;
                    let cx = plot_x - tw - 9.0;
                    show_text(content, label_font_key, font_size, cx, cy, label);
                }
            }
        }

        // Legend
        if let Some(ref legend) = c.legend {
            let is_line = matches!(
                c.chart_type,
                ChartType::Line | ChartType::Scatter | ChartType::Bubble | ChartType::Radar
            );
            let is_bubble = matches!(c.chart_type, ChartType::Bubble);
            let reverse_legend = matches!(
                c.chart_type,
                ChartType::Bar {
                    horizontal: true,
                    ..
                }
            );
            let series_order: Vec<(usize, &crate::model::ChartSeries)> = if reverse_legend {
                c.series.iter().enumerate().rev().collect()
            } else {
                c.series.iter().enumerate().collect()
            };
            let items: Vec<LegendItem> = series_order
                .iter()
                .map(|&(si, series)| {
                    let swatch = if is_line {
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
                LegendPosition::Right => LegendPlacement::Right {
                    x: plot_x + plot_w + 5.0,
                    center_y: plot_y + plot_h / 2.0,
                },
                _ => LegendPlacement::Bottom {
                    center_x: plot_x + plot_w / 2.0,
                    y: y - h + 4.0,
                },
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

    let has_legend = c.legend.is_some();
    let legend_on_right = c
        .legend
        .as_ref()
        .is_some_and(|l| l.position == LegendPosition::Right);
    let legend_on_bottom = c
        .legend
        .as_ref()
        .is_some_and(|l| l.position == LegendPosition::Bottom);

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
        .unwrap_or([179, 179, 179]);
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

    content.set_line_width(2.0);
    for (si, series) in c.series.iter().enumerate() {
        let color = series
            .color
            .unwrap_or(accent_colors[si % accent_colors.len()]);
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
        set_color(content, Some(color));
        let sym = resolve_marker(series.marker, si);
        for ci in 0..num_categories {
            let val = series.values.get(ci).copied().unwrap_or(0.0);
            let r = (val / axis_max) * radius;
            let a = cat_angle(ci);
            draw_marker(content, sym, cx + r * a.cos(), cy + r * a.sin(), 4.0);
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
                show_text(content, label_font_key, font_size, tx, ty, label);
            }
        }

        // Value labels along the 12 o'clock spoke (including "0" at center)
        let val_gap = 12.0;
        {
            let label = "0".to_string();
            let tw = text_width(&label, font_size, label_font);
            show_text(
                content,
                label_font_key,
                font_size,
                cx - tw - val_gap,
                cy - font_size * 0.3,
                &label,
            );
        }
        for ti in 1..=num_ticks {
            let val = ti as f32 * tick_step;
            let label = format_tick_label(val, tick_step);
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
                LegendPlacement::Bottom {
                    center_x: cx,
                    y: y - h + 12.0,
                }
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
