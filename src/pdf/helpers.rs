use pdf_writer::Content;

use crate::model::{
    Alignment, HorizontalRule, LineSpacing, Paragraph, ParagraphBorder, ParagraphBorders, Run,
};

use super::color::fill_rgb;

/// How far an object `slack` narrower than its box moves to follow the
/// paragraph alignment: centred takes half of it, right-aligned all of it.
pub(super) fn align_offset(alignment: Alignment, slack: f32) -> f32 {
    match alignment {
        Alignment::Center => slack / 2.0,
        Alignment::Right => slack,
        _ => 0.0,
    }
}

/// Draw a VML horizontal rule (o:hr) on the line whose box ends at
/// `line_bottom`, in a column starting at `col_x`, `col_w` wide. The width
/// percentage and alignment apply inside the paragraph's indents (slovak's
/// 60.9% rule is 291.7pt of 478.95, centred there), and the bar's bottom sits
/// 2pt above the line bottom (croatian's 18pt and isla's 12pt lines both).
pub(super) fn draw_horizontal_rule(
    content: &mut Content,
    para: &Paragraph,
    hr: &HorizontalRule,
    col_x: f32,
    col_w: f32,
    line_bottom: f32,
) {
    let area_w = (col_w - para.indent_left - para.indent_right).max(0.0);
    let rule_w = area_w * hr.width_pct / 100.0;
    let rule_x = col_x + para.indent_left + align_offset(para.alignment, area_w - rule_w);
    // Standard HRs (o:hrstd) render as a thin 0.5pt line
    let draw_h = if hr.is_standard { 0.5 } else { hr.height_pt };
    content.save_state();
    fill_rgb(content, hr.fill_color);
    content.rect(rule_x, line_bottom + 2.0, rule_w, draw_h);
    content.fill_nonzero();
    content.restore_state();
}

/// Approximate a circle with 4 cubic Bézier curves (path only — caller fills/strokes).
pub(super) fn draw_circle(content: &mut Content, cx: f32, cy: f32, r: f32) {
    let k = r * 0.552_284_8;
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

/// What line spacing adds to a font's natural line (`line_h` from
/// `resolve_line_h`), never negative.
pub(crate) fn spacing_leading(line_h: f32, font_size: f32, tallest_lhr: Option<f32>) -> f32 {
    (line_h - font_size * tallest_lhr.unwrap_or(1.2)).max(0.0)
}

/// How much of its ascent a line's first baseline sits below the line top.
/// Below single spacing Word shrinks the whole line box, ascent and descent
/// alike; at or above it the extra leading all goes below the baseline (Word
/// probes: Times New Roman and Arial, 7-20pt lines at 0.8, 0.9, 1, 1.15, 1.5).
pub(crate) fn auto_ascent_scale(ls: LineSpacing) -> f32 {
    match ls {
        LineSpacing::Auto(m) if m < 1.0 => m,
        _ => 1.0,
    }
}

/// The band a paragraph border adds to its box: its `space` and its stroke.
pub(super) fn border_band(border: Option<&ParagraphBorder>) -> f32 {
    border.map_or(0.0, |b| b.space_pt + b.width_pt)
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
/// §17.3.1.9: contextualSpacing drops a paragraph's spacing next to a paragraph
/// of the same style (a Title line keeps it beside a Normal one).
pub(super) fn drops_contextual_spacing(para: &Paragraph, neighbour: Option<&Paragraph>) -> bool {
    para.contextual_spacing && neighbour.is_some_and(|n| n.style_id == para.style_id)
}

/// Word drops a list item's HTML auto spacing (`beforeAutospacing` /
/// `afterAutospacing`) beside another item of the same list, as HTML
/// collapses `<li>` margins: federal_design_services' ten 9pt items step
/// 10.25, not 24.3, while the first keeps 14pt above and the last 14pt below.
fn same_list(para: &Paragraph, neighbour: Option<&Paragraph>) -> bool {
    para.list_item
        .zip(neighbour.and_then(|n| n.list_item))
        .is_some_and(|(a, b)| a.list_id == b.list_id)
}

fn drops_space_before(para: &Paragraph, prev: Option<&Paragraph>) -> bool {
    drops_contextual_spacing(para, prev) || (para.space_before_auto && same_list(para, prev))
}

pub(super) fn drops_space_after(para: &Paragraph, next: Option<&Paragraph>) -> bool {
    drops_contextual_spacing(para, next) || (para.space_after_auto && same_list(para, next))
}

/// `space_before` unless contextual or list auto spacing drops it beside `prev`.
pub(super) fn effective_space_before(para: &Paragraph, prev: Option<&Paragraph>) -> f32 {
    if drops_space_before(para, prev) {
        0.0
    } else {
        para.space_before
    }
}

/// `space_after` unless contextual or list auto spacing drops it beside `next`.
pub(super) fn effective_space_after(para: &Paragraph, next: Option<&Paragraph>) -> f32 {
    if drops_space_after(para, next) {
        0.0
    } else {
        para.space_after
    }
}

pub(super) fn joins_border_group(a: &Paragraph, b: &Paragraph) -> bool {
    let same = |x: f32, y: f32| (x - y).abs() < 0.01;
    borders_match(&a.borders, &b.borders)
        && same(a.indent_left, b.indent_left)
        && same(a.indent_right, b.indent_right)
        && same(a.indent_hanging, b.indent_hanging)
        && same(a.indent_first_line, b.indent_first_line)
}

/// The paragraph's runs and those of the paragraphs in its textboxes, at any depth.
pub(super) fn para_runs_with_textboxes(para: &Paragraph) -> Vec<&Run> {
    let mut out: Vec<&Run> = para.runs.iter().collect();
    for tb in &para.textboxes {
        for tp in tb.content() {
            out.extend(para_runs_with_textboxes(tp));
        }
    }
    out
}

/// The paragraph and the paragraphs of its textboxes, at any depth.
pub(super) fn collect_paras(para: &Paragraph) -> Vec<&Paragraph> {
    let mut out = vec![para];
    for tb in &para.textboxes {
        for tp in tb.content() {
            out.extend(collect_paras(tp));
        }
    }
    out
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
