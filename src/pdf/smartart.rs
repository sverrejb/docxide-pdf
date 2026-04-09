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
    image_names: &HashMap<usize, String>,
) {
    let sa_font_entry = seen_fonts
        .get(smartart_font_key)
        .or_else(|| seen_fonts.values().next());
    let sa_font_pdf_name = sa_font_entry.map(|e| e.pdf_name.as_str()).unwrap_or("F1");

    let apply_rotation = |content: &mut Content, shape: &crate::model::SmartArtShape, sx: f32, sy: f32| -> bool {
        let rotated = shape.rotation_deg.abs() > 0.01;
        if rotated {
            content.save_state();
            let cx = sx + shape.width / 2.0;
            let cy = sy + shape.height / 2.0;
            let rad = -shape.rotation_deg.to_radians();
            let cos = rad.cos();
            let sin = rad.sin();
            content.transform([
                cos, sin, -sin, cos,
                cx - cos * cx + sin * cy,
                cy - sin * cx - cos * cy,
            ]);
        }
        rotated
    };

    // Pass 1: Fill all shapes
    for shape in &diagram.shapes {
        if shape.fill.is_none() && shape.image_fill.is_none() {
            continue;
        }
        let sx = diag_x + shape.x;
        let sy = diag_y - shape.y - shape.height;
        let rotated = apply_rotation(content, shape, sx, sy);

        if let Some(fill) = shape.fill {
            content.save_state();
            color::fill_rgb(content, fill);
            draw_shape_path(content, sx, sy, shape.width, shape.height, &shape.shape_type);
            content.fill_nonzero();
            content.restore_state();
        }

        if let Some(ref img) = shape.image_fill {
            let key = std::sync::Arc::as_ptr(&img.data) as usize;
            if let Some(pdf_name) = image_names.get(&key) {
                content.save_state();
                draw_shape_path(content, sx, sy, shape.width, shape.height, &shape.shape_type);
                content.clip_nonzero();
                content.end_path();
                let img_aspect = img.pixel_width as f32 / img.pixel_height.max(1) as f32;
                let shape_aspect = shape.width / shape.height.max(0.001);
                let (draw_w, draw_h) = if img_aspect > shape_aspect {
                    (shape.height * img_aspect, shape.height)
                } else {
                    (shape.width, shape.width / img_aspect)
                };
                let dx = sx + (shape.width - draw_w) / 2.0;
                let dy = sy + (shape.height - draw_h) / 2.0;
                content.transform([draw_w, 0.0, 0.0, draw_h, dx, dy]);
                content.x_object(pdf_writer::Name(pdf_name.as_bytes()));
                content.restore_state();
            }
        }

        if rotated { content.restore_state(); }
    }

    // Pass 2: Stroke all shapes (on top of all fills)
    for shape in &diagram.shapes {
        let has_stroke = shape.stroke_color.is_some() && shape.stroke_width > 0.0;
        if !has_stroke { continue; }
        let sx = diag_x + shape.x;
        let sy = diag_y - shape.y - shape.height;
        let rotated = apply_rotation(content, shape, sx, sy);

        content.save_state();
        if let Some(stroke) = shape.stroke_color {
            content.set_line_width(shape.stroke_width);
            color::stroke_rgb(content, stroke);
        }
        draw_shape_path(content, sx, sy, shape.width, shape.height, &shape.shape_type);
        content.stroke();
        content.restore_state();

        if rotated { content.restore_state(); }
    }

    // Pass 3: Render all text on top of geometry
    for shape in &diagram.shapes {
        if shape.paragraphs.is_empty() { continue; }

        let (txt_x, txt_y, txt_w, txt_h) = if let Some((tx, ty, tw, th)) = shape.text_rect {
            (tx, ty, tw, th)
        } else {
            (shape.x, shape.y, shape.width, shape.height)
        };

        let base_fs = shape.paragraphs.iter()
            .flat_map(|p| p.runs.iter())
            .map(|r| r.font_size)
            .find(|&fs| fs > 0.0)
            .unwrap_or(10.0);
        let line_h = base_fs * 1.2;

        let total_text_h = shape.paragraphs.len() as f32 * line_h;
        let text_top_y = match shape.text_anchor {
            SmartArtTextAnchor::Top => diag_y - txt_y,
            SmartArtTextAnchor::Center => diag_y - txt_y - (txt_h - total_text_h) / 2.0,
            SmartArtTextAnchor::Bottom => diag_y - txt_y - (txt_h - total_text_h),
        };

        for (para_idx, para) in shape.paragraphs.iter().enumerate() {
            let bullet_text = para.bullet.as_deref().map(|b| format!("{b} ")).unwrap_or_default();
            let mut line_parts: Vec<(&str, f32, Option<&FontEntry>, &crate::model::SmartArtRun)> = Vec::new();
            let mut total_w = 0.0_f32;

            // Resolve bullet font entry once (bullet inherits the first run's font)
            let bullet_fe = if !bullet_text.is_empty() {
                para.runs.first().map(|r| resolve_run_font(r, seen_fonts, sa_font_entry))
            } else {
                None
            };
            if let Some((fe, _)) = bullet_fe {
                let first_run = para.runs.first().unwrap();
                total_w += charts::text_width(&bullet_text, first_run.font_size, fe);
            }

            for run in &para.runs {
                let (fe, _) = resolve_run_font(run, seen_fonts, sa_font_entry);
                let efs = effective_font_size(run);
                let rw = charts::text_width(&run.text, efs, fe);
                line_parts.push((&run.text, rw, fe, run));
                total_w += rw;
            }

            let line_x = diag_x + txt_x + match para.align {
                SmartArtTextAlign::Left => shape.text_inset_left,
                SmartArtTextAlign::Center => (txt_w - total_w) / 2.0,
                SmartArtTextAlign::Right => txt_w - total_w - shape.text_inset_left,
            };
            let line_y = text_top_y - base_fs - (para_idx as f32) * line_h;

            content.save_state();
            let mut cx = line_x;

            if let Some((fe, fpn)) = bullet_fe {
                let first_run = para.runs.first().unwrap();
                let fpn = fpn.unwrap_or(sa_font_pdf_name);
                if let Some(c) = first_run.color { color::fill_rgb(content, c); }
                else { content.set_fill_gray(0.0); }
                let bw = charts::text_width(&bullet_text, first_run.font_size, fe);
                charts::show_text_encoded(content, fpn, first_run.font_size, cx, line_y, &bullet_text, fe);
                cx += bw;
            }

            for &(text, rw, fe, run) in &line_parts {
                let efs = effective_font_size(run);
                let fpn = fe.map(|e| e.pdf_name.as_str()).unwrap_or(sa_font_pdf_name);
                let y_offset = baseline_y_offset(run, base_fs);
                let ry = line_y + y_offset;

                if let Some(hl) = run.highlight {
                    content.save_state();
                    color::fill_rgb(content, hl);
                    content.rect(cx, ry - base_fs * 0.15, rw, base_fs * 1.15);
                    content.fill_nonzero();
                    content.restore_state();
                }

                if let Some(c) = run.color { color::fill_rgb(content, c); }
                else { content.set_fill_gray(0.0); }
                let needs_synthetic_bold = run.bold && fe.is_some_and(|e| e.synthetic_bold);
                if needs_synthetic_bold {
                    content.set_line_width(efs * 0.02);
                    if let Some(c) = run.color { color::stroke_rgb(content, c); }
                    else { content.set_stroke_gray(0.0); }
                    content.set_text_rendering_mode(pdf_writer::types::TextRenderingMode::FillStroke);
                }
                charts::show_text_encoded(content, fpn, efs, cx, ry, text, fe);
                if needs_synthetic_bold {
                    content.set_text_rendering_mode(pdf_writer::types::TextRenderingMode::Fill);
                }

                if run.underline {
                    let thick = (efs * 0.05).max(0.5);
                    let ul_y = ry - efs * 0.12;
                    content.rect(cx, ul_y - thick, rw, thick);
                    content.fill_nonzero();
                }

                if run.strikethrough {
                    let thick = (efs * 0.05).max(0.5);
                    let st_y = ry + efs * 0.3;
                    content.rect(cx, st_y, rw, thick);
                    content.fill_nonzero();
                }

                cx += rw;
            }
            content.restore_state();
        }
    }
}

/// Resolve the font entry and PDF name for a SmartArt run, falling back to the
/// diagram-level default font when the run has no per-shape font.
fn resolve_run_font<'a>(
    run: &crate::model::SmartArtRun,
    seen_fonts: &'a HashMap<String, FontEntry>,
    fallback: Option<&'a FontEntry>,
) -> (Option<&'a FontEntry>, Option<&'a str>) {
    let fe = run.font_name.as_ref()
        .and_then(|n| {
            let key = super::fonts::smartart_font_key_str(n, run.bold, run.italic);
            seen_fonts.get(&key)
        })
        .or(fallback);
    let pdf_name = fe.map(|e| e.pdf_name.as_str());
    (fe, pdf_name)
}

fn effective_font_size(run: &crate::model::SmartArtRun) -> f32 {
    if run.baseline != 0 { run.font_size * 0.58 } else { run.font_size }
}

fn baseline_y_offset(run: &crate::model::SmartArtRun, base_fs: f32) -> f32 {
    if run.baseline > 0 { base_fs * 0.35 }       // superscript
    else if run.baseline < 0 { -base_fs * 0.14 }  // subscript
    else { 0.0 }
}

