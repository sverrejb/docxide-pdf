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

    let mut angle = std::f32::consts::FRAC_PI_2;
    let segments = 64;
    let values = &c.series[0].values;
    let total: f32 = values.iter().sum();

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * 2.0 * std::f32::consts::PI;
        fill_rgb(content, layout.colors[i % layout.colors.len()]);

        content.move_to(layout.cx, layout.cy);
        let n_seg =
            ((segments as f32 * sweep / (2.0 * std::f32::consts::PI)).ceil() as usize).max(2);
        for s in 0..=n_seg {
            let a = angle - (s as f32 / n_seg as f32) * sweep;
            content.line_to(layout.cx + radius * a.cos(), layout.cy + radius * a.sin());
        }
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

    let mut angle = std::f32::consts::FRAC_PI_2;
    let segments = 64;
    let values = &c.series[0].values;
    let total: f32 = values.iter().sum();

    for (i, &val) in values.iter().enumerate() {
        let sweep = (val / total) * 2.0 * std::f32::consts::PI;
        fill_rgb(content, layout.colors[i % layout.colors.len()]);

        let n_seg =
            ((segments as f32 * sweep / (2.0 * std::f32::consts::PI)).ceil() as usize).max(2);

        // Outer arc (clockwise = decreasing angle)
        let a_start = angle;
        let a_end = angle - sweep;
        content.move_to(
            layout.cx + outer_r * a_start.cos(),
            layout.cy + outer_r * a_start.sin(),
        );
        for s in 1..=n_seg {
            let a = a_start - (s as f32 / n_seg as f32) * sweep;
            content.line_to(layout.cx + outer_r * a.cos(), layout.cy + outer_r * a.sin());
        }
        // Inner arc backwards (increasing angle)
        content.line_to(
            layout.cx + inner_r * a_end.cos(),
            layout.cy + inner_r * a_end.sin(),
        );
        for s in 1..=n_seg {
            let a = a_end + (s as f32 / n_seg as f32) * sweep;
            content.line_to(layout.cx + inner_r * a.cos(), layout.cy + inner_r * a.sin());
        }
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
