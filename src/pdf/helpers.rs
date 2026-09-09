use pdf_writer::Content;

use crate::model::{LineSpacing, Paragraph, ParagraphBorder, ParagraphBorders};

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

/// Word merges the borders of adjacent paragraphs into one box only when their
/// border *and* indentation settings are identical (Paragraph dialog's Indentation
/// group); a differing indent — even by a few twips — starts a new border group
/// that gets its own top/bottom rule.
pub(super) fn joins_border_group(a: &Paragraph, b: &Paragraph) -> bool {
    let same = |x: f32, y: f32| (x - y).abs() < 0.01;
    borders_match(&a.borders, &b.borders)
        && same(a.indent_left, b.indent_left)
        && same(a.indent_right, b.indent_right)
        && same(a.indent_hanging, b.indent_hanging)
        && same(a.indent_first_line, b.indent_first_line)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn bordered(indent_left: f32) -> Paragraph {
        Paragraph {
            indent_left,
            borders: ParagraphBorders {
                bottom: Some(ParagraphBorder {
                    width_pt: 2.25,
                    space_pt: 1.0,
                    color: [166, 166, 166],
                }),
                ..Default::default()
            },
            ..Default::default()
        }
    }

    #[test]
    fn identical_borders_join_only_with_identical_indents() {
        // samtale: a br-only spacer and the "Medarbeiderens navn" line share a
        // bottom rule and indent, so only the group's last rule is drawn.
        assert!(joins_border_group(&bordered(0.0), &bordered(0.0)));
        // samtale survey items 12/13: numbering ind 1131tw vs direct ind 1128tw
        // — Word keeps a rule under each.
        assert!(!joins_border_group(&bordered(56.55), &bordered(56.4)));
        let mut other = bordered(0.0);
        other.borders.bottom.as_mut().unwrap().color = [0, 0, 0];
        assert!(!joins_border_group(&bordered(0.0), &other));
    }
}
