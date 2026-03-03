use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, to_winansi_bytes};
use crate::model::{ChartType, InlineChart, LegendPosition, MarkerSymbol};

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

fn set_color(content: &mut Content, color: Option<[u8; 3]>) {
    if let Some([r, g, b]) = color {
        content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
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

fn draw_marker(content: &mut Content, symbol: MarkerSymbol, x: f32, y: f32, r: f32) {
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
                let angle_outer = std::f32::consts::FRAC_PI_2 + i as f32 * std::f32::consts::PI * 2.0 / 5.0;
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

fn set_stroke_color(content: &mut Content, color: Option<[u8; 3]>) {
    if let Some([r, g, b]) = color {
        content.set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
    }
}

// Pie chart colors when series data points don't have individual colors
const DEFAULT_PIE_COLORS: &[[u8; 3]] = &[
    [68, 114, 196],
    [237, 125, 49],
    [165, 165, 165],
    [255, 192, 0],
    [91, 155, 213],
    [112, 173, 71],
    [38, 68, 120],
    [158, 72, 14],
];

pub(super) fn render_chart(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    seen_fonts: &HashMap<String, FontEntry>,
    default_font_name: &str,
) {
    let c = &chart.chart;
    let w = chart.display_width;
    let h = chart.display_height;

    let font_size = 8.0;
    let label_font_key = seen_fonts
        .keys()
        .find(|k| {
            let lower = k.to_lowercase();
            !lower.contains("symbol")
                && !lower.contains("serif")
                && !lower.contains("/")
        })
        .map(|s| s.as_str())
        .unwrap_or(default_font_name);
    let has_font = seen_fonts.contains_key(label_font_key);

    let num_categories = c
        .cat_axis
        .as_ref()
        .map(|ax| ax.labels.len())
        .filter(|&n| n > 0)
        .unwrap_or_else(|| {
            c.series
                .first()
                .map(|s| s.values.len())
                .unwrap_or(0)
        });
    let num_series = c.series.len();
    if num_series == 0 {
        return;
    }
    let is_scatter_like_early = matches!(c.chart_type, ChartType::Scatter | ChartType::Bubble);
    if num_categories == 0 && !is_scatter_like_early {
        return;
    }

    match c.chart_type {
        ChartType::Pie => {
            render_pie(chart, content, x, y, has_font, label_font_key, font_size);
            return;
        }
        ChartType::Doughnut { hole_size_pct } => {
            render_doughnut(chart, content, x, y, has_font, label_font_key, font_size, hole_size_pct);
            return;
        }
        ChartType::Radar => {
            render_radar(chart, content, x, y, has_font, label_font_key, font_size);
            return;
        }
        _ => {}
    }

    let horizontal = matches!(c.chart_type, ChartType::Bar { horizontal: true, .. });
    let is_scatter_like = matches!(c.chart_type, ChartType::Scatter | ChartType::Bubble);

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
        let ts = nice_tick_step(max_val, 10);
        let mut am = (max_val / ts).ceil() * ts;
        if max_val > am * 0.98 {
            am += ts;
        }
        (ts, am)
    } else {
        let ts = nice_tick_step(max_val, 5);
        let mut am = (max_val / ts).ceil() * ts;
        if max_val >= am * 0.9 {
            am += ts;
        }
        (ts, am)
    };
    if axis_max <= 0.0 {
        return;
    }

    // For scatter/bubble, compute x-axis max independently
    let (x_axis_max, x_tick_step) = if is_scatter_like {
        let x_max = c
            .series
            .iter()
            .filter_map(|s| s.x_values.as_ref())
            .flat_map(|xv| xv.iter())
            .copied()
            .fold(0.0f32, f32::max);
        let xts = nice_tick_step(x_max, 5);
        let mut xam = (x_max / xts).ceil() * xts;
        if x_max >= xam * 0.9 {
            xam += xts;
        }
        if xam <= 0.0 {
            return;
        }
        (xam, xts)
    } else {
        (0.0, 0.0)
    };

    // Compute margins based on content
    let max_tick_label = if tick_step.fract() == 0.0 {
        format!("{}", axis_max as i32)
    } else {
        format!("{:.1}", axis_max)
    };
    let val_label_w = text_width_approx(&max_tick_label, font_size) + 18.0;
    let cat_label_h = font_size + 6.0;

    let is_point_chart = matches!(c.chart_type, ChartType::Line | ChartType::Area | ChartType::Scatter | ChartType::Bubble);
    let margin_left = if !horizontal { val_label_w } else { w * 0.12 };
    let margin_right = if has_legend && legend_on_right {
        // Compute right margin from actual legend content width
        let legend_swatch = 5.5;
        let legend_gap = 21.0;
        let max_label_w = c
            .series
            .iter()
            .map(|s| text_width_approx(&s.label, 10.0))
            .fold(0.0f32, f32::max);
        let computed = legend_gap + legend_swatch + 4.0 + max_label_w + 8.0;
        if is_point_chart { computed.max(w * 0.12) } else { w * 0.22 }
    } else if is_point_chart {
        (w * 0.06).max(20.0)
    } else {
        8.0
    };
    let margin_top = h * 0.05;
    let margin_bottom = if has_legend && legend_on_bottom {
        h * 0.22
    } else {
        cat_label_h + 10.0
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
        content.set_stroke_rgb(
            color[0] as f32 / 255.0,
            color[1] as f32 / 255.0,
            color[2] as f32 / 255.0,
        );
        let num_ticks = (axis_max / tick_step).round() as usize;
        for i in 0..=num_ticks {
            let val = i as f32 * tick_step;
            let frac = val / axis_max;
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

    // Vertical gridlines for scatter/bubble (only if x-axis has majorGridlines)
    if is_scatter_like {
        if let Some(color) = c.cat_axis.as_ref().and_then(|a| a.gridline_color) {
            content.set_line_width(0.5);
            content.set_stroke_rgb(
                color[0] as f32 / 255.0,
                color[1] as f32 / 255.0,
                color[2] as f32 / 255.0,
            );
            let num_ticks = (x_axis_max / x_tick_step).round() as usize;
            for i in 0..=num_ticks {
                let val = i as f32 * x_tick_step;
                let frac = val / x_axis_max;
                let gx = plot_x + frac * plot_w;
                content.move_to(gx, plot_y);
                content.line_to(gx, plot_y + plot_h);
                content.stroke();
            }
        }
    }

    // Data rendering
    match c.chart_type {
        ChartType::Bar { horizontal: false, .. } => {
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
        ChartType::Bar { horizontal: true, .. } => {
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
                        (plot_x + ci as f32 * cat_w, plot_y + (val / axis_max) * plot_h)
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
                    let cat_w = plot_w / (num_categories - 1).max(1) as f32;
                    let lx = plot_x + ci as f32 * cat_w;
                    let ly = plot_y + (val / axis_max) * plot_h;
                    draw_marker(content, sym, lx, ly, marker_r);
                }
            }
        }
        ChartType::Area => {
            for series in &c.series {
                set_color(content, series.color);
                content.move_to(plot_x, plot_y);
                for (ci, &val) in series.values.iter().enumerate() {
                    let cat_w = plot_w / (num_categories - 1).max(1) as f32;
                    let lx = plot_x + ci as f32 * cat_w;
                    let ly = plot_y + (val / axis_max) * plot_h;
                    content.line_to(lx, ly);
                }
                let last_x = plot_x + (num_categories - 1) as f32 * (plot_w / (num_categories - 1).max(1) as f32);
                content.line_to(last_x, plot_y);
                content.close_path();
                content.fill_nonzero();

                // Stroke the top edge
                set_stroke_color(content, series.color);
                content.set_line_width(1.5);
                let mut first = true;
                for (ci, &val) in series.values.iter().enumerate() {
                    let cat_w = plot_w / (num_categories - 1).max(1) as f32;
                    let lx = plot_x + ci as f32 * cat_w;
                    let ly = plot_y + (val / axis_max) * plot_h;
                    if first {
                        content.move_to(lx, ly);
                        first = false;
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
            let all_sizes: Vec<f32> = c
                .series
                .iter()
                .filter_map(|s| s.bubble_sizes.as_ref())
                .flat_map(|bs| bs.iter().copied())
                .collect();
            let max_size = all_sizes.iter().copied().fold(0.0f32, f32::max);
            let min_r = 3.0;
            let max_r = 15.0;

            for series in &c.series {
                set_color(content, series.color);
                if let (Some(x_vals), Some(bsizes)) =
                    (&series.x_values, &series.bubble_sizes)
                {
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
                        content.fill_nonzero();
                    }
                }
            }
        }
        ChartType::Pie | ChartType::Doughnut { .. } | ChartType::Radar => unreachable!(),
    }

    // Plot area border
    if let Some(border) = c.plot_border_color {
        content.set_line_width(0.75);
        content.set_stroke_rgb(
            border[0] as f32 / 255.0,
            border[1] as f32 / 255.0,
            border[2] as f32 / 255.0,
        );
        content.rect(plot_x, plot_y, plot_w, plot_h);
        content.stroke();
    } else {
        let axis_color = c
            .val_axis
            .as_ref()
            .and_then(|a| a.line_color)
            .unwrap_or([179, 179, 179]);
        content.set_line_width(0.75);
        content.set_stroke_rgb(
            axis_color[0] as f32 / 255.0,
            axis_color[1] as f32 / 255.0,
            axis_color[2] as f32 / 255.0,
        );
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
            let label = if tick_step.fract() == 0.0 {
                format!("{}", val as i32)
            } else {
                format!("{:.1}", val)
            };
            let bytes = to_winansi_bytes(&label);
            let tw = text_width_approx(&label, font_size);
            let frac = val / axis_max;

            if !horizontal {
                let ly = plot_y + frac * plot_h - font_size * 0.3;
                let lx = plot_x - tw - 4.0;
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), font_size)
                    .next_line(lx, ly)
                    .show(Str(&bytes))
                    .end_text();
            } else {
                let lx = plot_x + frac * plot_w - tw / 2.0;
                let ly = plot_y - font_size - 3.0;
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), font_size)
                    .next_line(lx, ly)
                    .show(Str(&bytes))
                    .end_text();
            }
        }

        // X-axis tick labels for scatter/bubble
        if is_scatter_like {
            let num_x_ticks = (x_axis_max / x_tick_step).round() as usize;
            for i in 0..=num_x_ticks {
                let val = i as f32 * x_tick_step;
                let label = if x_tick_step.fract() == 0.0 {
                    format!("{}", val as i32)
                } else {
                    format!("{:.1}", val)
                };
                let bytes = to_winansi_bytes(&label);
                let tw = text_width_approx(&label, font_size);
                let frac = val / x_axis_max;
                let lx = plot_x + frac * plot_w - tw / 2.0;
                let ly = plot_y - font_size - 3.0;
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), font_size)
                    .next_line(lx, ly)
                    .show(Str(&bytes))
                    .end_text();
            }
        }

        // Category axis labels
        let is_point_chart = matches!(c.chart_type, ChartType::Line | ChartType::Area);
        if !is_scatter_like {
        if let Some(ref cat_axis) = c.cat_axis {
            for (ci, label) in cat_axis.labels.iter().enumerate() {
                let bytes = to_winansi_bytes(label);
                let tw = text_width_approx(label, font_size);
                if !horizontal {
                    let cx = if is_point_chart && num_categories > 1 {
                        let cat_w = plot_w / (num_categories - 1) as f32;
                        plot_x + ci as f32 * cat_w - tw / 2.0
                    } else {
                        let group_w = plot_w / num_categories as f32;
                        plot_x + ci as f32 * group_w + group_w / 2.0 - tw / 2.0
                    };
                    let cy = plot_y - font_size - 3.0;
                    content
                        .begin_text()
                        .set_font(Name(label_font_key.as_bytes()), font_size)
                        .next_line(cx, cy)
                        .show(Str(&bytes))
                        .end_text();
                } else {
                    let group_h = plot_h / num_categories as f32;
                    let cy =
                        plot_y + ci as f32 * group_h + group_h / 2.0 - font_size * 0.3;
                    let cx = plot_x - tw - 4.0;
                    content
                        .begin_text()
                        .set_font(Name(label_font_key.as_bytes()), font_size)
                        .next_line(cx, cy)
                        .show(Str(&bytes))
                        .end_text();
                }
            }
        }
        }

        // Legend
        render_legend(c, content, label_font_key, num_series, plot_x, plot_y, plot_w, plot_h, y, h);
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}

fn render_pie(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    has_font: bool,
    label_font_key: &str,
    _font_size: f32,
) {
    let c = &chart.chart;
    let w = chart.display_width;
    let h = chart.display_height;

    // Pie uses first series data points as slices
    let series = &c.series[0];
    let values = &series.values;
    let total: f32 = values.iter().sum();
    if total <= 0.0 {
        return;
    }

    let has_legend = c.legend.is_some();
    let legend_on_right = c
        .legend
        .as_ref()
        .is_some_and(|l| l.position == LegendPosition::Right);

    // For pie charts, category labels become the legend entries
    let labels: Vec<&str> = c
        .cat_axis
        .as_ref()
        .map(|ax| ax.labels.iter().map(|s| s.as_str()).collect())
        .unwrap_or_default();

    let margin = w * 0.05;
    let avail_h = h - margin * 2.0;
    let radius = avail_h / 2.0;

    let (cx, legend_x) = if has_legend && legend_on_right {
        // Legend occupies roughly the right 14.3% of the chart box
        let legend_area = w * 0.143;
        let legend_x = x + w - legend_area;
        // Pie is centered in the remaining space
        let pie_area_w = w - legend_area;
        let cx = x + pie_area_w / 2.0;
        (cx, legend_x)
    } else {
        let cx = x + w / 2.0;
        (cx, x + w)
    };
    let cy = y - h / 2.0;

    content.save_state();

    let mut angle = std::f32::consts::FRAC_PI_2; // start at top (90 degrees)
    let segments = 64; // segments per full circle for smooth arcs

    let pie_colors: &[[u8; 3]] = if c.accent_colors.is_empty() {
        DEFAULT_PIE_COLORS
    } else {
        &c.accent_colors
    };

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * 2.0 * std::f32::consts::PI;
        let color = pie_colors[i % pie_colors.len()];
        content.set_fill_rgb(
            color[0] as f32 / 255.0,
            color[1] as f32 / 255.0,
            color[2] as f32 / 255.0,
        );

        // Draw pie slice as a filled polygon: center -> arc -> center
        content.move_to(cx, cy);
        let n_seg = ((segments as f32 * sweep / (2.0 * std::f32::consts::PI)).ceil() as usize).max(2);
        for s in 0..=n_seg {
            let a = angle - (s as f32 / n_seg as f32) * sweep;
            let px = cx + radius * a.cos();
            let py = cy + radius * a.sin();
            content.line_to(px, py);
        }
        content.close_path();
        content.fill_nonzero();

        angle -= sweep;
    }

    // Legend for pie chart
    if has_font && has_legend && !labels.is_empty() {
        let legend_fs = 10.0;
        let swatch = 5.274;
        let spacing = 2.5;
        let line_h = 17.6;

        if legend_on_right {
            let lx = legend_x;
            let num_items = labels.len();
            let ly_start = cy + (num_items as f32 - 1.0) / 2.0 * line_h;
            for (i, label) in labels.iter().enumerate() {
                let ly = ly_start - i as f32 * line_h;
                let color = pie_colors[i % pie_colors.len()];
                content.set_fill_rgb(
                    color[0] as f32 / 255.0,
                    color[1] as f32 / 255.0,
                    color[2] as f32 / 255.0,
                );
                content.rect(lx, ly, swatch, swatch);
                content.fill_nonzero();

                content.set_fill_gray(0.0);
                let bytes = to_winansi_bytes(label);
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), legend_fs)
                    .next_line(lx + swatch + spacing, ly - 0.3)
                    .show(Str(&bytes))
                    .end_text();
            }
        } else {
            let total_w: f32 = labels
                .iter()
                .map(|l| swatch + spacing + text_width_approx(l, legend_fs) + 12.0)
                .sum();
            let mut lx = cx - total_w / 2.0;
            let ly = y - h + 4.0;
            for (i, label) in labels.iter().enumerate() {
                let color = pie_colors[i % pie_colors.len()];
                content.set_fill_rgb(
                    color[0] as f32 / 255.0,
                    color[1] as f32 / 255.0,
                    color[2] as f32 / 255.0,
                );
                content.rect(lx, ly, swatch, swatch);
                content.fill_nonzero();

                content.set_fill_gray(0.0);
                let bytes = to_winansi_bytes(label);
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), legend_fs)
                    .next_line(lx + swatch + spacing, ly + 1.0)
                    .show(Str(&bytes))
                    .end_text();
                lx += swatch + spacing + text_width_approx(label, legend_fs) + 12.0;
            }
        }
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}

fn render_doughnut(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    has_font: bool,
    label_font_key: &str,
    font_size: f32,
    hole_size_pct: f32,
) {
    let c = &chart.chart;
    let w = chart.display_width;
    let h = chart.display_height;

    let series = &c.series[0];
    let values = &series.values;
    let total: f32 = values.iter().sum();
    if total <= 0.0 {
        return;
    }

    let has_legend = c.legend.is_some();
    let legend_on_right = c
        .legend
        .as_ref()
        .is_some_and(|l| l.position == LegendPosition::Right);

    let labels: Vec<&str> = c
        .cat_axis
        .as_ref()
        .map(|ax| ax.labels.iter().map(|s| s.as_str()).collect())
        .unwrap_or_default();

    let margin = w * 0.05;
    let avail_h = h - margin * 2.0;
    let outer_r = avail_h / 2.0;
    let inner_r = outer_r * (hole_size_pct / 100.0);

    let (cx, legend_x) = if has_legend && legend_on_right {
        let legend_area = w * 0.143;
        let legend_x = x + w - legend_area;
        let pie_area_w = w - legend_area;
        let cx = x + pie_area_w / 2.0;
        (cx, legend_x)
    } else {
        let cx = x + w / 2.0;
        (cx, x + w)
    };
    let cy = y - h / 2.0;

    content.save_state();

    let mut angle = std::f32::consts::FRAC_PI_2;
    let segments = 64;

    let pie_colors: &[[u8; 3]] = if c.accent_colors.is_empty() {
        DEFAULT_PIE_COLORS
    } else {
        &c.accent_colors
    };

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * 2.0 * std::f32::consts::PI;
        let color = pie_colors[i % pie_colors.len()];
        content.set_fill_rgb(
            color[0] as f32 / 255.0,
            color[1] as f32 / 255.0,
            color[2] as f32 / 255.0,
        );

        let n_seg = ((segments as f32 * sweep / (2.0 * std::f32::consts::PI)).ceil() as usize).max(2);

        // Outer arc (clockwise = decreasing angle)
        let a_start = angle;
        let a_end = angle - sweep;
        let first_a = a_start;
        content.move_to(cx + outer_r * first_a.cos(), cy + outer_r * first_a.sin());
        for s in 1..=n_seg {
            let a = a_start - (s as f32 / n_seg as f32) * sweep;
            content.line_to(cx + outer_r * a.cos(), cy + outer_r * a.sin());
        }
        // Line to inner arc at end angle
        content.line_to(cx + inner_r * a_end.cos(), cy + inner_r * a_end.sin());
        // Inner arc backwards (increasing angle)
        for s in 1..=n_seg {
            let a = a_end + (s as f32 / n_seg as f32) * sweep;
            content.line_to(cx + inner_r * a.cos(), cy + inner_r * a.sin());
        }
        content.close_path();
        content.fill_nonzero();

        angle -= sweep;
    }

    // Reuse pie legend
    if has_font && has_legend && !labels.is_empty() {
        let legend_fs = 10.0;
        let swatch = 5.274;
        let spacing = 2.5;
        let line_h = 17.6;

        if legend_on_right {
            let lx = legend_x;
            let num_items = labels.len();
            let ly_start = cy + (num_items as f32 - 1.0) / 2.0 * line_h;
            for (i, label) in labels.iter().enumerate() {
                let ly = ly_start - i as f32 * line_h;
                let color = pie_colors[i % pie_colors.len()];
                content.set_fill_rgb(
                    color[0] as f32 / 255.0,
                    color[1] as f32 / 255.0,
                    color[2] as f32 / 255.0,
                );
                content.rect(lx, ly, swatch, swatch);
                content.fill_nonzero();

                content.set_fill_gray(0.0);
                let bytes = to_winansi_bytes(label);
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), font_size)
                    .next_line(lx + swatch + spacing, ly - 0.3)
                    .show(Str(&bytes))
                    .end_text();
            }
        } else {
            let total_w: f32 = labels
                .iter()
                .map(|l| swatch + spacing + text_width_approx(l, legend_fs) + 12.0)
                .sum();
            let mut lx = cx - total_w / 2.0;
            let ly = y - h + 4.0;
            for (i, label) in labels.iter().enumerate() {
                let color = pie_colors[i % pie_colors.len()];
                content.set_fill_rgb(
                    color[0] as f32 / 255.0,
                    color[1] as f32 / 255.0,
                    color[2] as f32 / 255.0,
                );
                content.rect(lx, ly, swatch, swatch);
                content.fill_nonzero();

                content.set_fill_gray(0.0);
                let bytes = to_winansi_bytes(label);
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), legend_fs)
                    .next_line(lx + swatch + spacing, ly + 1.0)
                    .show(Str(&bytes))
                    .end_text();
                lx += swatch + spacing + text_width_approx(label, legend_fs) + 12.0;
            }
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
    font_size: f32,
) {
    let c = &chart.chart;
    let w = chart.display_width;
    let h = chart.display_height;

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
    let tick_step = nice_tick_step(max_val, 4);
    let mut axis_max = (max_val / tick_step).ceil() * tick_step;
    if max_val > axis_max * 0.9 {
        axis_max += tick_step;
    }
    if axis_max <= 0.0 {
        return;
    }

    let margin_right = if has_legend && legend_on_right { w * 0.22 } else { w * 0.08 };
    let margin_bottom = if has_legend && legend_on_bottom { h * 0.18 } else { h * 0.08 };
    let margin_top = h * 0.08;
    let margin_left = w * 0.08;

    let avail_w = w - margin_left - margin_right;
    let avail_h = h - margin_top - margin_bottom;
    let radius = avail_w.min(avail_h) / 2.0;
    let cx = x + margin_left + avail_w / 2.0;
    let cy = y - margin_top - avail_h / 2.0;

    let angle_step = 2.0 * std::f32::consts::PI / num_categories as f32;

    // Angle for category i: start at top (pi/2) and go clockwise
    let cat_angle = |ci: usize| -> f32 {
        std::f32::consts::FRAC_PI_2 - ci as f32 * angle_step
    };

    content.save_state();

    // Concentric polygon gridlines
    let num_ticks = (axis_max / tick_step).round() as usize;
    let grid_color = c
        .val_axis
        .as_ref()
        .and_then(|a| a.gridline_color)
        .unwrap_or([179, 179, 179]);
    content.set_line_width(0.5);
    content.set_stroke_rgb(
        grid_color[0] as f32 / 255.0,
        grid_color[1] as f32 / 255.0,
        grid_color[2] as f32 / 255.0,
    );

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
    let accent_colors: &[[u8; 3]] = if c.accent_colors.is_empty() {
        DEFAULT_PIE_COLORS
    } else {
        &c.accent_colors
    };

    content.set_line_width(2.0);
    for (si, series) in c.series.iter().enumerate() {
        let color = series.color.unwrap_or(accent_colors[si % accent_colors.len()]);
        content.set_stroke_rgb(
            color[0] as f32 / 255.0,
            color[1] as f32 / 255.0,
            color[2] as f32 / 255.0,
        );
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
            let px = cx + r * a.cos();
            let py = cy + r * a.sin();
            draw_marker(content, sym, px, py, 3.0);
        }
    }

    if has_font {
        content.set_fill_gray(0.0);

        // Category labels at perimeter
        if let Some(ref cat_axis) = c.cat_axis {
            let label_r = radius + 8.0;
            for (ci, label) in cat_axis.labels.iter().enumerate() {
                let a = cat_angle(ci);
                let lx = cx + label_r * a.cos();
                let ly = cy + label_r * a.sin();
                let bytes = to_winansi_bytes(label);
                let tw = text_width_approx(label, font_size);
                // Center label around the point
                let text_x = lx - tw / 2.0;
                let text_y = ly - font_size * 0.3;
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), font_size)
                    .next_line(text_x, text_y)
                    .show(Str(&bytes))
                    .end_text();
            }
        }

        // Value labels along the 12 o'clock spoke
        for ti in 1..=num_ticks {
            let val = ti as f32 * tick_step;
            let label = if tick_step.fract() == 0.0 {
                format!("{}", val as i32)
            } else {
                format!("{:.1}", val)
            };
            let r = radius * (ti as f32 / num_ticks as f32);
            let ly = cy + r - font_size * 0.3;
            let bytes = to_winansi_bytes(&label);
            let tw = text_width_approx(&label, font_size);
            content
                .begin_text()
                .set_font(Name(label_font_key.as_bytes()), font_size)
                .next_line(cx - tw - 2.0, ly)
                .show(Str(&bytes))
                .end_text();
        }

        // Legend
        let num_series = c.series.len();
        if has_legend {
            let legend_fs = 10.0;
            let swatch = 5.5;
            let spacing = 2.5;
            let line_h = 18.0;

            if legend_on_right {
                let lx = x + w - margin_right + 10.0;
                let block_h = swatch + (num_series as f32 - 1.0) * line_h;
                let ly_start = cy + block_h / 2.0 - swatch + 5.0;
                for (si, series) in c.series.iter().enumerate() {
                    let ly = ly_start - si as f32 * line_h;
                    let color = series.color.unwrap_or(accent_colors[si % accent_colors.len()]);
                    set_color(content, Some(color));
                    let sym = resolve_marker(series.marker, si);
                    draw_marker(content, sym, lx + swatch / 2.0, ly + swatch / 2.0, swatch / 2.0);
                    content.set_fill_gray(0.0);
                    let bytes = to_winansi_bytes(&series.label);
                    content
                        .begin_text()
                        .set_font(Name(label_font_key.as_bytes()), legend_fs)
                        .next_line(lx + swatch + spacing, ly - 0.3)
                        .show(Str(&bytes))
                        .end_text();
                }
            } else {
                let total_w: f32 = c
                    .series
                    .iter()
                    .map(|s| swatch + spacing + text_width_approx(&s.label, legend_fs) + 12.0)
                    .sum();
                let mut lx = cx - total_w / 2.0;
                let ly = y - h + 4.0;
                for (si, series) in c.series.iter().enumerate() {
                    let color = series.color.unwrap_or(accent_colors[si % accent_colors.len()]);
                    set_color(content, Some(color));
                    let sym = resolve_marker(series.marker, si);
                    draw_marker(content, sym, lx + swatch / 2.0, ly + swatch / 2.0, swatch / 2.0);
                    content.set_fill_gray(0.0);
                    let bytes = to_winansi_bytes(&series.label);
                    content
                        .begin_text()
                        .set_font(Name(label_font_key.as_bytes()), legend_fs)
                        .next_line(lx + swatch + spacing, ly + 1.0)
                        .show(Str(&bytes))
                        .end_text();
                    lx += swatch + spacing + text_width_approx(&series.label, legend_fs) + 12.0;
                }
            }
        }
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}

fn render_legend(
    c: &crate::model::Chart,
    content: &mut Content,
    label_font_key: &str,
    num_series: usize,
    plot_x: f32,
    plot_y: f32,
    plot_w: f32,
    plot_h: f32,
    y: f32,
    h: f32,
) {
    let Some(ref legend) = c.legend else { return };
    let legend_fs = 10.0;
    let swatch = 5.5;
    let spacing = 2.5;
    let line_h = 18.0;

    let is_line = matches!(c.chart_type, ChartType::Line | ChartType::Scatter | ChartType::Bubble | ChartType::Radar);

    match legend.position {
        LegendPosition::Right => {
            let lx = plot_x + plot_w + 21.0;
            let block_h = swatch + (num_series as f32 - 1.0) * line_h;
            let ly_start = plot_y + plot_h / 2.0 + block_h / 2.0 - swatch + 5.0;
            for (si, series) in c.series.iter().enumerate() {
                let ly = ly_start - si as f32 * line_h;
                set_color(content, series.color);
                if is_line {
                    let sym = resolve_marker(series.marker, si);
                    draw_marker(content, sym, lx + swatch / 2.0, ly + swatch / 2.0, swatch / 2.0);
                } else {
                    content.rect(lx, ly, swatch, swatch);
                    content.fill_nonzero();
                }

                content.set_fill_gray(0.0);
                let bytes = to_winansi_bytes(&series.label);
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), legend_fs)
                    .next_line(lx + swatch + spacing, ly - 0.3)
                    .show(Str(&bytes))
                    .end_text();
            }
        }
        _ => {
            let total_w: f32 = c
                .series
                .iter()
                .map(|s| swatch + spacing + text_width_approx(&s.label, legend_fs) + 12.0)
                .sum();
            let mut lx = plot_x + (plot_w - total_w) / 2.0;
            let ly = y - h + 4.0;
            for (si, series) in c.series.iter().enumerate() {
                set_color(content, series.color);
                if is_line {
                    let sym = resolve_marker(series.marker, si);
                    draw_marker(content, sym, lx + swatch / 2.0, ly + swatch / 2.0, swatch / 2.0);
                } else {
                    content.rect(lx, ly, swatch, swatch);
                    content.fill_nonzero();
                }

                content.set_fill_gray(0.0);
                let bytes = to_winansi_bytes(&series.label);
                content
                    .begin_text()
                    .set_font(Name(label_font_key.as_bytes()), legend_fs)
                    .next_line(lx + swatch + spacing, ly + 1.0)
                    .show(Str(&bytes))
                    .end_text();
                lx += swatch + spacing + text_width_approx(&series.label, legend_fs) + 12.0;
            }
        }
    }
}
