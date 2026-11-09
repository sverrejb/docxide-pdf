//! Preset and custom shape outlines as PDF paths, and pictures clipped to them.

use pdf_writer::Content;

use crate::geometry::{self, ResolvedCommand};
use crate::model::ShapeGeometry;

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
        Some(eval) => {
            for path in &eval.paths {
                emit_path_commands(content, x, y, path);
            }
        }
        None => {
            content.rect(x, y, w, h);
        }
    }
}

/// Evaluate a preset shape's text rectangle — the region Word lays text out
/// within, which for shapes like arrows, triangles, and pentagons is smaller
/// than the bounding box (e.g. an arrow's text sits in its rectangular body,
/// not over the arrowhead). Returns `(x_from_left, y_from_bottom, w, h)` in
/// points, or None for custom geometry / shapes without a defined text rect.
pub(super) fn shape_text_rect(
    shape: &ShapeGeometry,
    w: f32,
    h: f32,
) -> Option<(f32, f32, f32, f32)> {
    let eval = evaluate_shape_geometry(shape, w as f64, h as f64)?;
    let (l, b, rw, rh) = eval.text_rect?;
    Some((l as f32, b as f32, rw as f32, rh as f32))
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

fn emit_path_commands(content: &mut Content, x: f32, y: f32, path: &geometry::EvaluatedPath) {
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

/// Emit a shape's path for STROKING, honoring each subpath's OOXML `stroke`
/// flag. Shapes such as braces, brackets, arcs and callouts define a closed
/// fill-only subpath plus an open stroke-only subpath; stroking the fill
/// subpath too would draw a spurious closing segment (e.g. a vertical line
/// across the open side of a "{" brace). When no subpath is explicitly
/// strokable, all subpaths are emitted, preserving behavior for single-path
/// shapes whose definitions don't set the flag.
pub(super) fn draw_shape_stroke_path(
    content: &mut Content,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    shape: &ShapeGeometry,
) {
    match evaluate_shape_geometry(shape, w as f64, h as f64) {
        Some(eval) => {
            let any_strokable = eval.paths.iter().any(|p| p.stroke);
            for path in &eval.paths {
                if any_strokable && !path.stroke {
                    continue;
                }
                emit_path_commands(content, x, y, path);
            }
        }
        None => {
            content.rect(x, y, w, h);
        }
    }
}

pub(super) fn render_image_with_clip(
    content: &mut Content,
    pdf_name: &str,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    clip: Option<&ShapeGeometry>,
) {
    content.save_state();
    if let Some(geom) = clip {
        draw_shape_path(content, x, y, w, h, geom);
        content.clip_nonzero();
        content.end_path();
    }
    content.transform([w, 0.0, 0.0, h, x, y]);
    content.x_object(pdf_writer::Name(pdf_name.as_bytes()));
    content.restore_state();
}

pub(super) fn stroke_image_border(
    content: &mut Content,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    stroke_color: [u8; 3],
    stroke_width: f32,
    clip: Option<&ShapeGeometry>,
) {
    content.save_state();
    color::stroke_rgb(content, stroke_color);
    content.set_line_width(stroke_width);
    if let Some(geom) = clip {
        draw_shape_path(content, x, y, w, h, geom);
    } else {
        content.rect(x, y, w, h);
    }
    content.stroke();
    content.restore_state();
}
