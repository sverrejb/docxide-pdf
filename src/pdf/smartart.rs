use std::collections::HashMap;

use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::geometry::{self, ResolvedCommand};
use crate::model::{ShapeGeometry, SmartArtDiagram};

use super::charts;

pub(super) fn draw_shape_path(content: &mut Content, x: f32, y: f32, w: f32, h: f32, shape: &ShapeGeometry) {
    let evaluated = evaluate_shape_geometry(shape, w as f64, h as f64);
    match evaluated {
        Some(eval) => emit_evaluated_paths(content, x, y, &eval),
        None => { content.rect(x, y, w, h); }
    }
}

fn evaluate_shape_geometry(shape: &ShapeGeometry, w: f64, h: f64) -> Option<geometry::EvaluatedShape> {
    if let Some(ref custom) = shape.custom {
        Some(geometry::evaluate_custom(custom, w, h, &shape.adjustments))
    } else if let Some(ref preset) = shape.preset {
        geometry::evaluate_preset(preset, w, h, &shape.adjustments)
    } else {
        None
    }
}

fn emit_evaluated_paths(content: &mut Content, x: f32, y: f32, shape: &geometry::EvaluatedShape) {
    for path in &shape.paths {
        for cmd in &path.commands {
            match cmd {
                ResolvedCommand::MoveTo(px, py) => {
                    content.move_to(x + *px as f32, y + *py as f32);
                }
                ResolvedCommand::LineTo(px, py) => {
                    content.line_to(x + *px as f32, y + *py as f32);
                }
                ResolvedCommand::CubicTo { x1, y1, x2, y2, x: px, y: py } => {
                    content.cubic_to(
                        x + *x1 as f32, y + *y1 as f32,
                        x + *x2 as f32, y + *y2 as f32,
                        x + *px as f32, y + *py as f32,
                    );
                }
                ResolvedCommand::Close => {
                    content.close_path();
                }
            }
        }
    }
}

pub(super) fn render_smartart(
    content: &mut Content,
    diagram: &SmartArtDiagram,
    diag_x: f32,
    diag_y: f32,
    seen_fonts: &HashMap<String, FontEntry>,
) {
    let sa_font_entry = seen_fonts
        .values()
        .find(|e| {
            let lower = e.pdf_name.to_lowercase();
            !lower.contains("symbol")
        })
        .or_else(|| seen_fonts.values().next());
    let sa_font_pdf_name = sa_font_entry
        .map(|e| e.pdf_name.as_str())
        .unwrap_or("F1");

    for shape in &diagram.shapes {
        let sx = diag_x + shape.x;
        let sy = diag_y - shape.y - shape.height;
        let has_fill = shape.fill.is_some();
        let has_stroke = shape.stroke_color.is_some() && shape.stroke_width > 0.0;

        if has_fill || has_stroke {
            content.save_state();
            if let Some([r, g, b]) = shape.fill {
                content.set_fill_rgb(
                    r as f32 / 255.0,
                    g as f32 / 255.0,
                    b as f32 / 255.0,
                );
            }
            if let Some([r, g, b]) = shape.stroke_color {
                content.set_line_width(shape.stroke_width);
                content.set_stroke_rgb(
                    r as f32 / 255.0,
                    g as f32 / 255.0,
                    b as f32 / 255.0,
                );
            }
            draw_shape_path(content, sx, sy, shape.width, shape.height, &shape.shape_type);
            if has_fill && has_stroke {
                content.fill_nonzero_and_stroke();
            } else if has_fill {
                content.fill_nonzero();
            } else {
                content.stroke();
            }
            content.restore_state();
        }

        if !shape.text.is_empty() && shape.font_size > 0.0 {
            let fs = shape.font_size;
            let lines: Vec<&str> = shape.text.split('\n').collect();
            let line_h = fs * 1.2;
            let total_text_h = lines.len() as f32 * line_h;
            let text_top_y = diag_y - shape.y - (shape.height - total_text_h) / 2.0;
            content.save_state();
            if let Some([r, g, b]) = shape.text_color {
                content.set_fill_rgb(
                    r as f32 / 255.0,
                    g as f32 / 255.0,
                    b as f32 / 255.0,
                );
            } else {
                content.set_fill_gray(0.0);
            }
            for (i, line) in lines.iter().enumerate() {
                let tw = charts::text_width(line, fs, sa_font_entry);
                let tx = diag_x + shape.x + (shape.width - tw) / 2.0;
                let ty = text_top_y - fs - (i as f32) * line_h;
                charts::show_text_encoded(content, sa_font_pdf_name, fs, tx, ty, line, sa_font_entry);
            }
            content.restore_state();
        }
    }
}
