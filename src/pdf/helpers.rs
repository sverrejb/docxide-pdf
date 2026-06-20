use pdf_writer::Content;

use crate::model::{LineSpacing, ParagraphBorder, ParagraphBorders};

/// Approximate a circle with 4 cubic Bézier curves (path only — caller fills/strokes).
pub(super) fn draw_circle(content: &mut Content, cx: f32, cy: f32, r: f32) {
    let k = r * 0.5522847498;
    content.move_to(cx + r, cy);
    content.cubic_to(cx + r, cy + k, cx + k, cy + r, cx, cy + r);
    content.cubic_to(cx - k, cy + r, cx - r, cy + k, cx - r, cy);
    content.cubic_to(cx - r, cy - k, cx - k, cy - r, cx, cy - r);
    content.cubic_to(cx + k, cy - r, cx + r, cy - k, cx + r, cy);
    content.close_path();
}

pub(crate) fn resolve_line_h(ls: LineSpacing, font_size: f32, tallest_lhr: Option<f32>) -> f32 {
    match ls {
        LineSpacing::Auto(mult) => tallest_lhr
            .map(|ratio| font_size * ratio * mult)
            .unwrap_or(font_size * 1.2 * mult),
        LineSpacing::Exact(pts) => pts,
        LineSpacing::AtLeast(min_pts) => {
            let natural = tallest_lhr
                .map(|ratio| font_size * ratio)
                .unwrap_or(font_size * 1.2);
            natural.max(min_pts)
        }
    }
}

fn border_eq(a: &Option<ParagraphBorder>, b: &Option<ParagraphBorder>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => a.width_pt == b.width_pt && a.color == b.color,
        _ => false,
    }
}

pub(super) fn borders_match(a: &ParagraphBorders, b: &ParagraphBorders) -> bool {
    border_eq(&a.top, &b.top)
        && border_eq(&a.bottom, &b.bottom)
        && border_eq(&a.left, &b.left)
        && border_eq(&a.right, &b.right)
        && border_eq(&a.between, &b.between)
}
