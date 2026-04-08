use std::collections::HashMap;

use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::geometry::{self, ResolvedCommand};
use crate::model::{ShapeGeometry, SmartArtDiagram, SmartArtTextAnchor, SmartArtTextAlign};

use super::charts;
use super::color;

pub(super) fn draw_shape_path(
    content: &mut Content,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    shape: &ShapeGeometry,
) {
    match evaluate_shape_geometry(shape, w as f64, h as f64) {
        Some(eval) => emit_evaluated_paths(content, x, y, &eval),
        None => {
            content.rect(x, y, w, h);
        }
    }
}

fn evaluate_shape_geometry(
    shape: &ShapeGeometry,
    w: f64,
    h: f64,
) -> Option<geometry::EvaluatedShape> {
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
            match *cmd {
                ResolvedCommand::MoveTo(px, py) => {
                    content.move_to(x + px as f32, y + py as f32);
                }
                ResolvedCommand::LineTo(px, py) => {
                    content.line_to(x + px as f32, y + py as f32);
                }
                ResolvedCommand::CubicTo {
                    x1,
                    y1,
                    x2,
                    y2,
                    x: px,
                    y: py,
                } => {
                    content.cubic_to(
                        x + x1 as f32,
                        y + y1 as f32,
                        x + x2 as f32,
                        y + y2 as f32,
                        x + px as f32,
                        y + py as f32,
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
    smartart_font_key: &str,
) {
    let sa_font_entry = seen_fonts
        .get(smartart_font_key)
        .or_else(|| seen_fonts.values().next());
    let sa_font_pdf_name = sa_font_entry.map(|e| e.pdf_name.as_str()).unwrap_or("F1");

    for shape in &diagram.shapes {
        let sx = diag_x + shape.x;
        let sy = diag_y - shape.y - shape.height;
        let has_fill = shape.fill.is_some();
        let has_stroke = shape.stroke_color.is_some() && shape.stroke_width > 0.0;

        // Apply rotation around the shape's center if non-zero
        let rotated = shape.rotation_deg.abs() > 0.01;
        if rotated {
            content.save_state();
            let cx = sx + shape.width / 2.0;
            let cy = sy + shape.height / 2.0;
            let rad = -shape.rotation_deg.to_radians();
            let cos = rad.cos();
            let sin = rad.sin();
            // Rotate around (cx, cy): translate to origin, rotate, translate back
            content.transform([
                cos, sin, -sin, cos,
                cx - cos * cx + sin * cy,
                cy - sin * cx - cos * cy,
            ]);
        }

        if has_fill || has_stroke {
            content.save_state();
            if let Some(fill) = shape.fill {
                color::fill_rgb(content, fill);
            }
            if let Some(stroke) = shape.stroke_color {
                content.set_line_width(shape.stroke_width);
                color::stroke_rgb(content, stroke);
            }
            draw_shape_path(
                content,
                sx,
                sy,
                shape.width,
                shape.height,
                &shape.shape_type,
            );
            match (has_fill, has_stroke) {
                (true, true) => { content.fill_nonzero_and_stroke(); }
                (true, false) => { content.fill_nonzero(); }
                (false, true) => { content.stroke(); }
                (false, false) => {}
            }
            content.restore_state();
        }

        // Close rotation before drawing text — txXfrm coordinates are in
        // the unrotated diagram space and already account for the shape rotation
        if rotated {
            content.restore_state();
        }

        if !shape.text.is_empty() && shape.font_size > 0.0 {
            let fs = shape.font_size;
            let para_lines: Vec<&str> = shape.text.split('\n').collect();
            let line_h = fs * 1.2;

            // Use txXfrm text rectangle if available, otherwise shape bounds
            let (txt_x, txt_y, txt_w, txt_h) = if let Some((tx, ty, tw, th)) = shape.text_rect {
                (tx, ty, tw, th)
            } else {
                (shape.x, shape.y, shape.width, shape.height)
            };

            let mut wrapped: Vec<String> = Vec::new();
            for para in &para_lines {
                wrap_text_into(para, fs, txt_w, sa_font_entry, &mut wrapped);
            }

            let total_text_h = wrapped.len() as f32 * line_h;
            let text_top_y = match shape.text_anchor {
                SmartArtTextAnchor::Top => diag_y - txt_y,
                SmartArtTextAnchor::Center => diag_y - txt_y - (txt_h - total_text_h) / 2.0,
                SmartArtTextAnchor::Bottom => diag_y - txt_y - (txt_h - total_text_h),
            };
            content.save_state();
            if let Some(color) = shape.text_color {
                color::fill_rgb(content, color);
            } else {
                content.set_fill_gray(0.0);
            }
            for (i, line) in wrapped.iter().enumerate() {
                let tw = charts::text_width(line, fs, sa_font_entry);
                let tx = diag_x + txt_x + match shape.text_align {
                    SmartArtTextAlign::Left => shape.text_inset_left,
                    SmartArtTextAlign::Center => (txt_w - tw) / 2.0,
                    SmartArtTextAlign::Right => txt_w - tw - shape.text_inset_left,
                };
                let ty = text_top_y - fs - (i as f32) * line_h;
                charts::show_text_encoded(
                    content,
                    sa_font_pdf_name,
                    fs,
                    tx,
                    ty,
                    line,
                    sa_font_entry,
                );
            }
            content.restore_state();
        }
    }
}

/// Word-wrap `text` into lines that fit within `max_width` at the given font size.
fn wrap_text_into(
    text: &str,
    font_size: f32,
    max_width: f32,
    font_entry: Option<&FontEntry>,
    out: &mut Vec<String>,
) {
    if text.is_empty() {
        out.push(String::new());
        return;
    }
    let full_w = charts::text_width(text, font_size, font_entry);
    if full_w <= max_width {
        out.push(text.to_string());
        return;
    }
    let words: Vec<&str> = text.split_whitespace().collect();
    if words.is_empty() {
        out.push(text.to_string());
        return;
    }
    let mut current_line = String::new();
    for word in &words {
        if current_line.is_empty() {
            current_line.push_str(word);
        } else {
            let candidate = format!("{} {}", current_line, word);
            let cw = charts::text_width(&candidate, font_size, font_entry);
            if cw <= max_width {
                current_line = candidate;
            } else {
                out.push(current_line);
                current_line = word.to_string();
            }
        }
    }
    if !current_line.is_empty() {
        out.push(current_line);
    }
}
