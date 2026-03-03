use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, to_winansi_bytes};
use crate::model::{ChartType, InlineChart, LegendPosition};

fn ceil_nice(val: f32) -> f32 {
    if val <= 0.0 {
        return 1.0;
    }
    let mag = 10.0f32.powf(val.log10().floor());
    let norm = val / mag;
    let nice = if norm <= 1.0 {
        1.0
    } else if norm <= 2.0 {
        2.0
    } else if norm <= 5.0 {
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
    // Prefer a sans-serif font for chart labels (Word uses Calibri/Aptos)
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
        .unwrap_or_else(|| {
            c.series
                .first()
                .map(|s| s.values.len())
                .unwrap_or(0)
        });
    let num_series = c.series.len();
    if num_categories == 0 || num_series == 0 {
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

    let tick_step = nice_tick_step(max_val, 5);
    let axis_max = (max_val / tick_step).ceil() * tick_step;
    if axis_max <= 0.0 {
        return;
    }

    let ChartType::Bar { horizontal, .. } = c.chart_type;

    // Compute margins based on content
    let max_tick_label = if tick_step.fract() == 0.0 {
        format!("{}", axis_max as i32)
    } else {
        format!("{:.1}", axis_max)
    };
    let val_label_w = text_width_approx(&max_tick_label, font_size) + 6.0;

    let cat_label_h = font_size + 6.0;

    let margin_left = if !horizontal { val_label_w } else { w * 0.12 };
    let margin_right = if has_legend && legend_on_right {
        w * 0.22
    } else {
        4.0
    };
    let margin_top = h * 0.05;
    let margin_bottom = if has_legend && legend_on_bottom {
        h * 0.22
    } else {
        cat_label_h + 4.0
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

    // Bars — gapWidth is percentage of individual bar width
    let gap_ratio = c.gap_width_pct / 100.0;
    if !horizontal {
        let group_w = plot_w / num_categories as f32;
        let bar_w = group_w / (num_series as f32 + gap_ratio);
        let gap = gap_ratio * bar_w;

        for (ci, _) in (0..num_categories).enumerate() {
            let group_x = plot_x + ci as f32 * group_w + gap / 2.0;
            for (si, series) in c.series.iter().enumerate() {
                let val = series.values.get(ci).copied().unwrap_or(0.0);
                let bar_h = (val / axis_max) * plot_h;
                let bx = group_x + si as f32 * bar_w;

                if let Some([r, g, b]) = series.color {
                    content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                }
                content.rect(bx, plot_y, bar_w, bar_h);
                content.fill_nonzero();
            }
        }
    } else {
        let group_h = plot_h / num_categories as f32;
        let bar_h = group_h / (num_series as f32 + gap_ratio);
        let gap = gap_ratio * bar_h;

        for ci in 0..num_categories {
            let group_y = plot_y + (num_categories - 1 - ci) as f32 * group_h + gap / 2.0;
            for (si, series) in c.series.iter().enumerate() {
                let val = series.values.get(ci).copied().unwrap_or(0.0);
                let bw = (val / axis_max) * plot_w;
                let by = group_y + si as f32 * bar_h;

                if let Some([r, g, b]) = series.color {
                    content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                }
                content.rect(plot_x, by, bw, bar_h);
                content.fill_nonzero();
            }
        }
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
        // Fallback: just left + bottom axes
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

    // Axis labels (only if we have the font)
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

        // Category axis labels
        if let Some(ref cat_axis) = c.cat_axis {
            for (ci, label) in cat_axis.labels.iter().enumerate() {
                let bytes = to_winansi_bytes(label);
                let tw = text_width_approx(label, font_size);
                if !horizontal {
                    let group_w = plot_w / num_categories as f32;
                    let cx = plot_x + ci as f32 * group_w + group_w / 2.0 - tw / 2.0;
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
                        plot_y + (num_categories - 1 - ci) as f32 * group_h + group_h / 2.0 - font_size * 0.3;
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

        // Legend
        if let Some(ref legend) = c.legend {
            let legend_fs = 10.0;
            let swatch = 5.5;
            let spacing = 2.5;
            let line_h = 18.0;

            match legend.position {
                LegendPosition::Right => {
                    let lx = plot_x + plot_w + 21.0;
                    let block_h = swatch + (num_series as f32 - 1.0) * line_h;
                    let ly_start = plot_y + plot_h / 2.0 + block_h / 2.0 - swatch + 5.0;
                    for (si, series) in c.series.iter().enumerate() {
                        let ly = ly_start - si as f32 * line_h;
                        if let Some([r, g, b]) = series.color {
                            content.set_fill_rgb(
                                r as f32 / 255.0,
                                g as f32 / 255.0,
                                b as f32 / 255.0,
                            );
                        }
                        content.rect(lx, ly, swatch, swatch);
                        content.fill_nonzero();

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
                    for series in &c.series {
                        if let Some([r, g, b]) = series.color {
                            content.set_fill_rgb(
                                r as f32 / 255.0,
                                g as f32 / 255.0,
                                b as f32 / 255.0,
                            );
                        }
                        content.rect(lx, ly, swatch, swatch);
                        content.fill_nonzero();

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
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}
