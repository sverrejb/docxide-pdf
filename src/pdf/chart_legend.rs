use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::model::MarkerSymbol;

use super::charts::{draw_marker, fill_rgb, show_text, stroke_rgb, text_width};

pub(super) enum SwatchStyle {
    Rect,
    Marker(MarkerSymbol),
    LineMarker(MarkerSymbol),
}

pub(super) struct LegendItem<'a> {
    pub label: &'a str,
    pub color: [u8; 3],
    pub swatch: SwatchStyle,
}

pub(super) enum LegendPlacement {
    Right { x: f32, center_y: f32 },
    Bottom { center_x: f32, y: f32 },
}

pub(super) fn render_chart_legend(
    content: &mut Content,
    items: &[LegendItem],
    placement: LegendPlacement,
    label_font_key: &str,
    label_font: Option<&FontEntry>,
    swatch_size: f32,
    line_h: f32,
) {
    if items.is_empty() {
        return;
    }

    let legend_fs = 10.0;
    let spacing = 2.5;

    let item_width = |item: &LegendItem| -> f32 {
        let ext = line_extension(&item.swatch);
        swatch_size + ext * 2.0 + spacing + text_width(item.label, legend_fs, label_font) + 12.0
    };

    match placement {
        LegendPlacement::Right { x: lx, center_y } => {
            let block_h = swatch_size + (items.len() as f32 - 1.0) * line_h;
            let ly_start = center_y + block_h / 2.0 - swatch_size + 5.0;
            for (i, item) in items.iter().enumerate() {
                let ly = ly_start - i as f32 * line_h;
                render_swatch(content, &item.swatch, item.color, lx, ly, swatch_size);
                let text_x = lx + swatch_size + spacing + line_extension(&item.swatch);
                content.set_fill_gray(0.0);
                show_text(
                    content,
                    label_font_key,
                    legend_fs,
                    text_x,
                    ly - 0.3,
                    item.label,
                );
            }
        }
        LegendPlacement::Bottom { center_x, y: ly } => {
            let total_w: f32 = items.iter().map(|item| item_width(item)).sum();
            let mut lx = center_x - total_w / 2.0;
            for item in items {
                render_swatch(content, &item.swatch, item.color, lx, ly, swatch_size);
                let ext = line_extension(&item.swatch);
                let text_x = lx + swatch_size + spacing + ext;
                content.set_fill_gray(0.0);
                show_text(
                    content,
                    label_font_key,
                    legend_fs,
                    text_x,
                    ly + 1.0,
                    item.label,
                );
                lx += item_width(item);
            }
        }
    }
}

fn line_extension(swatch: &SwatchStyle) -> f32 {
    match swatch {
        SwatchStyle::LineMarker(_) => 12.0,
        _ => 0.0,
    }
}

fn render_swatch(
    content: &mut Content,
    swatch: &SwatchStyle,
    color: [u8; 3],
    x: f32,
    y: f32,
    size: f32,
) {
    let center_x = x + size / 2.0;
    let center_y = y + size / 2.0;
    let radius = size / 2.0;

    match swatch {
        SwatchStyle::Rect => {
            fill_rgb(content, color);
            content.rect(x, y, size, size);
            content.fill_nonzero();
        }
        SwatchStyle::Marker(sym) => {
            fill_rgb(content, color);
            draw_marker(content, *sym, center_x, center_y, radius);
        }
        SwatchStyle::LineMarker(sym) => {
            let ext = line_extension(swatch);
            stroke_rgb(content, color);
            content.set_line_width(1.5);
            content.move_to(center_x - ext, center_y);
            content.line_to(center_x + ext, center_y);
            content.stroke();
            fill_rgb(content, color);
            draw_marker(content, *sym, center_x, center_y, radius);
        }
    }
}
