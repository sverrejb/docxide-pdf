use std::f32::consts::{FRAC_PI_2, TAU};

use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::model::{InlineChart, LegendPosition};

use super::chart_legend::{LegendItem, LegendPlacement, SwatchStyle, render_chart_legend};
use super::charts::{fill_rgb, resolve_accent_colors, text_width};

struct RadialLayout {
    cx: f32,
    cy: f32,
    legend_x: f32,
    colors: Vec<[u8; 3]>,
    has_legend: bool,
    legend_on_right: bool,
    labels: Vec<String>,
}

const ARC_SEGMENTS: usize = 64;

fn arc_segment_count(sweep: f32) -> usize {
    ((ARC_SEGMENTS as f32 * sweep / TAU).ceil() as usize).max(2)
}

fn emit_arc(content: &mut Content, cx: f32, cy: f32, radius: f32, start: f32, sweep: f32) {
    let n = arc_segment_count(sweep);
    for s in 0..=n {
        let a = start - (s as f32 / n as f32) * sweep;
        content.line_to(cx + radius * a.cos(), cy + radius * a.sin());
    }
}

fn emit_arc_reverse(content: &mut Content, cx: f32, cy: f32, radius: f32, end: f32, sweep: f32) {
    let n = arc_segment_count(sweep);
    for s in 0..=n {
        let a = end + (s as f32 / n as f32) * sweep;
        content.line_to(cx + radius * a.cos(), cy + radius * a.sin());
    }
}

fn setup_radial_chart(
    chart: &InlineChart,
    x: f32,
    y: f32,
    label_font: Option<&FontEntry>,
) -> Option<RadialLayout> {
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

    let labels: Vec<String> = c
        .cat_axis
        .as_ref()
        .map(|ax| ax.labels.clone())
        .unwrap_or_default();

    let colors = resolve_accent_colors(&c.accent_colors).to_vec();

    let (cx, legend_x) = if has_legend && legend_on_right {
        let legend_fs = 10.0;
        let swatch = 5.274;
        let spacing = 2.5;
        let max_label_w = labels
            .iter()
            .map(|l| text_width(l, legend_fs, label_font))
            .fold(0.0f32, f32::max);
        let legend_area = (swatch + spacing + max_label_w + 10.0).max(w * 0.143);
        let pie_area_w = w - legend_area;
        (x + pie_area_w / 2.0, x + w - legend_area)
    } else {
        (x + w / 2.0, x + w)
    };
    let cy = y - h / 2.0;

    Some(RadialLayout {
        cx,
        cy,
        legend_x,
        colors,
        has_legend,
        legend_on_right,
        labels,
    })
}

fn render_radial_legend(
    content: &mut Content,
    layout: &RadialLayout,
    label_font_key: &str,
    y: f32,
    h: f32,
    label_font: Option<&FontEntry>,
) {
    if !layout.has_legend || layout.labels.is_empty() {
        return;
    }
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
        LegendPlacement::Right {
            x: layout.legend_x,
            center_y: layout.cy,
        }
    } else {
        LegendPlacement::Bottom {
            center_x: layout.cx,
            y: y - h + 4.0,
        }
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

pub(super) fn render_pie(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    has_font: bool,
    label_font_key: &str,
    label_font: Option<&FontEntry>,
) {
    let Some(layout) = setup_radial_chart(chart, x, y, label_font) else {
        return;
    };
    let c = &chart.chart;
    let h = chart.display_height;
    let margin = chart.display_width * 0.05;
    let radius = (h - margin * 2.0) / 2.0;

    content.save_state();

    let mut angle = FRAC_PI_2;
    let values = &c.series[0].values;
    let total: f32 = values.iter().sum();

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * TAU;
        fill_rgb(content, layout.colors[i % layout.colors.len()]);

        content.move_to(layout.cx, layout.cy);
        emit_arc(content, layout.cx, layout.cy, radius, angle, sweep);
        content.close_path();
        content.fill_nonzero();

        angle -= sweep;
    }

    if has_font {
        render_radial_legend(content, &layout, label_font_key, y, h, label_font);
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}

pub(super) fn render_doughnut(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    has_font: bool,
    label_font_key: &str,
    hole_size_pct: f32,
    label_font: Option<&FontEntry>,
) {
    let Some(layout) = setup_radial_chart(chart, x, y, label_font) else {
        return;
    };
    let c = &chart.chart;
    let h = chart.display_height;
    let margin = chart.display_width * 0.05;
    let outer_r = (h - margin * 2.0) / 2.0;
    let inner_r = outer_r * (hole_size_pct / 100.0);

    content.save_state();

    let mut angle = FRAC_PI_2;
    let values = &c.series[0].values;
    let total: f32 = values.iter().sum();

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * TAU;
        fill_rgb(content, layout.colors[i % layout.colors.len()]);

        // Outer arc (clockwise = decreasing angle)
        content.move_to(
            layout.cx + outer_r * angle.cos(),
            layout.cy + outer_r * angle.sin(),
        );
        emit_arc(content, layout.cx, layout.cy, outer_r, angle, sweep);
        // Inner arc backwards (increasing angle)
        emit_arc_reverse(content, layout.cx, layout.cy, inner_r, angle - sweep, sweep);
        content.close_path();
        content.fill_nonzero();

        angle -= sweep;
    }

    if has_font {
        render_radial_legend(content, &layout, label_font_key, y, h, label_font);
    }

    content.set_fill_gray(0.0);
    content.restore_state();
}
