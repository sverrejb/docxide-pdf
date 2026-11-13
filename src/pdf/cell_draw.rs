//! Table cell drawing primitives: borders, shading, pictures, floats and list labels.

use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::FontEntry;
use crate::model::{Alignment, BorderStyle, CellBorder, CellMargins, CellVAlign};

use super::RenderContext;
use super::color::{fill_rgb, stroke_rgb};
use super::layout::LinkAnnotation;
use super::table_layout::{CellFloatingImageLayout, CellParagraphLayout, HfSubstitution};
use super::tagging::CellTagger;

// ponytail: `page_top` ignores w:vAlign's shift of the page body.
pub(super) fn draw_border(
    content: &mut Content,
    border: &CellBorder,
    page_top: f32,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
) {
    if !border.present {
        return;
    }
    // A horizontal border is a band; a vertical one's ends are points.
    let band = if (y1 - y2).abs() < 0.01 {
        border.width
    } else {
        0.0
    };
    let (y1, y2) = (
        super::layout::rule_on_grid(y1, page_top, band),
        super::layout::rule_on_grid(y2, page_top, band),
    );
    let w = border.width;
    content.save_state();
    content.set_line_width(w);
    if let Some(c) = border.color {
        stroke_rgb(content, c);
    }
    match border.style {
        BorderStyle::Dotted => {
            content.set_line_cap(pdf_writer::types::LineCapStyle::RoundCap);
            content.set_dash_pattern([0.0, w * 3.0], 0.0);
        }
        BorderStyle::Dashed | BorderStyle::DashSmallGap => {
            let dash = if border.style == BorderStyle::DashSmallGap {
                w * 3.0
            } else {
                w * 4.0
            };
            content.set_dash_pattern([dash, w * 2.0], 0.0);
        }
        BorderStyle::DashDot => {
            content.set_dash_pattern([w * 4.0, w * 2.0, 0.0, w * 2.0], 0.0);
        }
        BorderStyle::DashDotDot => {
            content.set_dash_pattern([w * 4.0, w * 2.0, 0.0, w * 2.0, 0.0, w * 2.0], 0.0);
        }
        BorderStyle::Double => {
            // Word renders each line of a double border at the full specified
            // width, separated by a clear gap of the same width. The gap here is
            // the center-to-center distance, so it must be 2×width: each line's
            // stroke spans ±width/2, and 2×width centers leave a width-wide gap
            // between their edges. (gap == width made the edges touch, so
            // antialiasing merged the pair into one thick line.)
            let thin = w.max(0.25);
            let gap = thin * 2.0;
            content.set_line_width(thin);
            content.move_to(x1, y1);
            content.line_to(x2, y2);
            content.stroke();
            let dx = if (x1 - x2).abs() < 0.01 { gap } else { 0.0 };
            let dy = if (y1 - y2).abs() < 0.01 { gap } else { 0.0 };
            content.move_to(x1 - dx, y1 - dy);
            content.line_to(x2 - dx, y2 - dy);
            content.stroke();
            content.restore_state();
            return;
        }
        // ponytail: cell borders draw wave and thin/thick pairs as one line
        BorderStyle::Single
        | BorderStyle::Wave
        | BorderStyle::ThinThick
        | BorderStyle::ThickThin
        | BorderStyle::ThinThickThin => {}
    }
    content.move_to(x1, y1);
    content.line_to(x2, y2);
    content.stroke();
    content.restore_state();
}

pub(super) fn draw_cell_borders(
    content: &mut Content,
    borders: &crate::model::CellBorders,
    page_top: f32,
    bx: f32,
    top: f32,
    bottom: f32,
    col_w: f32,
    draw_top: bool,
    draw_bottom: bool,
) {
    if draw_top {
        draw_border(content, &borders.top, page_top, bx, top, bx + col_w, top);
    }
    if draw_bottom {
        draw_border(
            content,
            &borders.bottom,
            page_top,
            bx,
            bottom,
            bx + col_w,
            bottom,
        );
    }
    draw_border(content, &borders.left, page_top, bx, top, bx, bottom);
    draw_border(
        content,
        &borders.right,
        page_top,
        bx + col_w,
        top,
        bx + col_w,
        bottom,
    );
}

/// Paint a cell background: hatching when a line/cross pattern is present,
/// otherwise a solid fill. No-op when both are absent.
fn paint_cell_background(
    content: &mut Content,
    shading: Option<[u8; 3]>,
    hatch: Option<crate::model::HatchPattern>,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) {
    if let Some(hp) = hatch {
        draw_hatch(content, &hp, x, y, w, h);
    } else if let Some(s) = shading {
        content.save_state();
        fill_rgb(content, s);
        content.rect(x, y, w, h);
        content.fill_nonzero();
        content.restore_state();
    }
}

/// Draw a `w:shd` line/cross pattern as real hatching: a background fill plus
/// thin stroked lines clipped to the cell rect.
fn draw_hatch(
    content: &mut Content,
    hp: &crate::model::HatchPattern,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) {
    use crate::model::HatchKind::*;
    if w <= 0.0 || h <= 0.0 {
        return;
    }
    content.save_state();
    fill_rgb(content, hp.bg);
    content.rect(x, y, w, h);
    content.fill_nonzero();
    // Clip subsequent strokes to the cell so the diagonal lines don't bleed out.
    content.rect(x, y, w, h);
    content.clip_nonzero();
    content.end_path();
    stroke_rgb(content, hp.fg);
    content.set_line_width(0.5);
    let s = 4.0_f32;
    if matches!(hp.kind, Horz | CrossHorzVert) {
        let mut yy = y + s;
        while yy < y + h {
            content.move_to(x, yy);
            content.line_to(x + w, yy);
            yy += s;
        }
    }
    if matches!(hp.kind, Vert | CrossHorzVert) {
        let mut xx = x + s;
        while xx < x + w {
            content.move_to(xx, y);
            content.line_to(xx, y + h);
            xx += s;
        }
    }
    if matches!(hp.kind, DiagFwd | CrossDiag) {
        let mut off = 0.0;
        while off <= w + h {
            content.move_to(x + off - h, y);
            content.line_to(x + off, y + h);
            off += s;
        }
    }
    if matches!(hp.kind, DiagBack | CrossDiag) {
        let mut off = 0.0;
        while off <= w + h {
            content.move_to(x + off - h, y + h);
            content.line_to(x + off, y);
            off += s;
        }
    }
    content.stroke();
    content.restore_state();
}

/// Word shades a cell up to the inner edge of each of its own borders (a thick
/// top border does not thin the shading at the other edges); the neighbours'
/// borders paint over it.
pub(super) fn draw_cell_shading(
    content: &mut Content,
    shading: Option<[u8; 3]>,
    hatch: Option<crate::model::HatchPattern>,
    borders: &crate::model::CellBorders,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
) {
    let half = |b: &crate::model::CellBorder| if b.present { b.width / 2.0 } else { 0.0 };
    let (top, bottom) = (half(&borders.top), half(&borders.bottom));
    let (left, right) = (half(&borders.left), half(&borders.right));
    paint_cell_background(
        content,
        shading,
        hatch,
        x + left,
        y + bottom,
        w - left - right,
        h - top - bottom,
    );
}

/// Render an inline image within a table cell paragraph, handling alignment,
/// shadow, glow, stroke, and the image transform. Returns the image height
/// consumed so the caller can advance its cursor.
pub(super) fn render_cell_inline_image(
    content: &mut Content,
    para: &CellParagraphLayout,
    img_name: &str,
    cell_x: f32,
    col_w: f32,
    cursor_y: f32,
    cm: &CellMargins,
) -> f32 {
    // The picture is the paragraph's first line, so it starts at the first-line
    // indent: nabl's logo paragraph (w:ind left=-198) sits 9.9pt into the cell
    // margin in Word.
    let left = cm.left + para.indent_left + para.indent_first_line - para.indent_hanging;
    let text_w = (col_w - left - cm.right - para.indent_right).max(0.0);
    let img_x = cell_x
        + left
        + match para.alignment {
            Alignment::Center => (text_w - para.image_width - para.image_trim) / 2.0,
            Alignment::Right => text_w - para.image_width - para.image_trim,
            _ => 0.0,
        };
    let img_y = cursor_y - para.image_height;

    // Word clips a picture wider than its cell to the cell's edges (nabl's
    // 90.75pt logo in an 81pt column).
    let clip = img_x < cell_x || img_x + para.image_width > cell_x + col_w;
    if clip {
        content.save_state();
        content
            .rect(cell_x, img_y - 1.0, col_w, para.image_height + 2.0)
            .clip_nonzero()
            .end_path();
    }

    super::color::draw_image_halo(
        content,
        para.image_shadow.as_ref(),
        para.image_glow.as_ref(),
        (img_x, img_y, para.image_width, para.image_height),
        para.image_fx.as_ref(),
    );

    super::shape_path::render_image_with_clip(
        content,
        img_name,
        img_x,
        img_y,
        para.image_width,
        para.image_height,
        para.image_clip.as_ref(),
    );

    if let Some(sc) = para.image_stroke_color {
        super::shape_path::stroke_image_border(
            content,
            img_x,
            img_y,
            para.image_width,
            para.image_height,
            sc,
            para.image_stroke_width,
            para.image_clip.as_ref(),
        );
    }
    if clip {
        content.restore_state();
    }

    para.image_height
}

pub(super) fn valign_offset(v_align: CellVAlign, available: f32, content_h: f32) -> f32 {
    match v_align {
        CellVAlign::Top => 0.0,
        CellVAlign::Center => ((available - content_h) / 2.0).max(0.0),
        CellVAlign::Bottom => (available - content_h).max(0.0),
    }
}

pub(super) fn clip_fixed_cell_pictures(
    content: &mut Content,
    para: &CellParagraphLayout,
    cell_x: f32,
    col_w: f32,
    top: f32,
    height: f32,
) -> bool {
    if !para.clip_inline_pictures
        || !para
            .lines
            .iter()
            .flat_map(|line| &line.chunks)
            .any(|chunk| chunk.inline_image_name.is_some())
        || para.image_line_ascent <= para.font_size * para.ascender_ratio
    {
        return false;
    }
    // Exact spacing leaves tall inline pictures on the fixed text baseline;
    // Word clips the part protruding above the paragraph's line box.
    content.save_state();
    content
        .rect(cell_x, top - height, col_w, height)
        .clip_nonzero()
        .end_path();
    true
}

/// A cell paragraph's `w:shd` fill behind its lines, from `top` down `h`.
pub(super) fn draw_cell_paragraph_shading(
    content: &mut Content,
    para: &crate::model::Paragraph,
    cell_x: f32,
    col_w: f32,
    cm: &CellMargins,
    top: f32,
    h: f32,
) {
    let Some(color) = para.shading else { return };
    let (box_left_in, box_right_in) = super::helpers::border_box_indents(para);
    let (left_out, right_out) = super::helpers::shading_outsets(&para.borders);
    let left = cell_x + cm.left + box_left_in - left_out;
    let right = cell_x + col_w - cm.right - box_right_in + right_out;
    content.save_state();
    super::color::fill_rgb(content, color);
    content.rect(left, top - h, right - left, h);
    content.fill_nonzero();
    content.restore_state();
}

pub(super) fn draw_cell_paragraph_bottom(
    content: &mut Content,
    para: &CellParagraphLayout,
    cell_x: f32,
    col_w: f32,
    cm: &CellMargins,
    bottom: f32,
) {
    if let Some(border) = &para.bottom_border {
        // Word extends a horizontal paragraph rule 1.5pt past the text area,
        // independently of the stroke width and vertical space.
        let left = cell_x + cm.left + para.indent_left - 1.5;
        let right = cell_x + col_w - cm.right - para.indent_right + 1.5;
        let y = bottom + para.bottom_band.1 - para.bottom_band.0 - border.extent() / 2.0;
        super::color::stroke_border(content, border, (left, y), (right, y), 1.0);
    }
}

/// Draw one floating picture anchored to a cell paragraph whose top is `para_y`.
pub(super) fn draw_cell_float(
    content: &mut Content,
    fi: &CellFloatingImageLayout,
    cell_x: f32,
    para_y: f32,
) {
    let fi_x = cell_x + fi.h_offset;
    let fi_y_bottom = para_y - fi.v_offset - fi.display_height;
    content.save_state();
    super::positioning::push_center_rotation(
        content,
        fi_x,
        fi_y_bottom,
        fi.display_width,
        fi.display_height,
        fi.rotation_deg,
    );
    content.transform([
        fi.display_width,
        0.0,
        0.0,
        fi.display_height,
        fi_x,
        fi_y_bottom,
    ]);
    content.x_object(Name(fi.pdf_name.as_bytes()));
    content.restore_state();
}

/// Minimal textbox rendering for cell contexts: stroke/fill the shape and
/// render text content. Doesn't support all features that body textboxes do.
pub(super) fn render_simple_textbox(
    content: &mut Content,
    tb: &crate::model::Textbox,
    tb_x: f32,
    tb_y_top: f32,
    ctx: &RenderContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
    links: &mut Vec<LinkAnnotation>,
    // Tags the box's part of its story in a Sect of the cell.
    tagger: Option<&mut CellTagger<'_>>,
    hf_sub: Option<&HfSubstitution>,
) {
    use super::shape_path::draw_shape_path;
    use super::textbox_render::render_textbox_paragraphs;
    use crate::model::ShapeFill;

    let tb_width = tb.width_pt;
    let tb_height = super::textbox_render::textbox_height(tb, ctx);

    // Fill
    if let Some(ShapeFill::Solid(color)) = &tb.fill {
        content.save_state();
        fill_rgb(content, *color);
        draw_shape_path(
            content,
            tb_x,
            tb_y_top - tb_height,
            tb_width,
            tb_height,
            &tb.shape_type,
        );
        content.fill_nonzero();
        content.restore_state();
    }
    super::textbox_render::render_picture_fill(
        content,
        tb,
        (tb_x, tb_y_top - tb_height, tb_width, tb_height),
        ctx,
    );

    super::textbox_render::stroke_textbox_outline(
        content,
        tb,
        tb_x,
        tb_y_top - tb_height,
        tb_width,
        tb_height,
    );

    // Text: route through the shared textbox renderer (same path as body
    // textboxes), anchored and clipped like a body box.
    let story = super::textbox_render::story_window(tb, tb_height, ctx);
    let tag = tagger
        .filter(|_| !story.is_empty())
        .map(|t| super::textbox_render::StoryTag {
            sect: t.sect(),
            tags: &mut *t.tags,
            page: t.page,
        });
    let content_x = tb_x + tb.margin_left;
    let content_w = (tb_width - tb.margin_left - tb.margin_right).max(0.0);
    let avail_h = tb_height - tb.margin_top - tb.margin_bottom;
    let anchor_offset =
        super::textbox_render::anchor_offset(tb, &story, content_w, content_w, avail_h, ctx);
    // Word clips a fixed-size box's text inside its top and bottom insets.
    let clip = super::layout::push_clip(
        content,
        (!matches!(tb.auto_fit, crate::model::AutoFit::Shape)).then_some((
            tb_x,
            tb_y_top - tb_height + tb.margin_bottom,
            tb_width,
            avail_h.max(0.0),
        )),
    );
    render_textbox_paragraphs(
        &story,
        content,
        content_x,
        content_w,
        content_w,
        tb_y_top - tb.margin_top - anchor_offset,
        0.0,
        0.0,
        None,
        true,
        links,
        ctx,
        clip.then_some(tb_y_top - tb_height),
        gradient_specs,
        tag,
        hf_sub,
        None,
    );
    if clip {
        content.restore_state();
    }
}

/// Per-line (x, width) for a cell paragraph whose first lines run beside a
/// left floating nested table (`CellParagraphLayout::beside_lines`).
pub(super) fn beside_float_geometry(
    para: &CellParagraphLayout,
    text_x: f32,
    text_w: f32,
) -> Option<Vec<(f32, f32)>> {
    (para.beside_lines > 0).then(|| {
        (0..para.lines.len())
            .map(|i| {
                if i < para.beside_lines {
                    (
                        text_x + para.beside_indent,
                        (text_w - para.beside_indent).max(1.0),
                    )
                } else {
                    (text_x, text_w)
                }
            })
            .collect()
    })
}

// A separately painted label follows the alignment of its first text line,
// including an empty numbered paragraph whose suffix still occupies a tab.
pub(super) fn cell_label_alignment_offset(para: &CellParagraphLayout, inner_w: f32) -> f32 {
    // Preserve signed remaining width when indents exceed a narrow cell.
    let text_w = inner_w
        - para.indent_left
        - para.indent_right
        - para.float_indent_left
        - if para.beside_lines > 0 {
            para.beside_indent
        } else {
            0.0
        };
    let first_width = para.lines.first().map_or(0.0, |line| line.total_width);
    let slack = (text_w + para.text_hanging - first_width).max(0.0);
    super::helpers::align_offset(para.alignment, slack)
}

pub(super) fn draw_cell_label(
    content: &mut Content,
    para: &CellParagraphLayout,
    label_x: f32,
    baseline_y: f32,
    fonts: &HashMap<String, FontEntry>,
) {
    // A one-piece label shows in its own font; a mixed one piece by piece.
    let single = if para.label_pieces.len() > 1 {
        None
    } else {
        let Some(entry) = para
            .label_pieces
            .first()
            .and_then(|(k, _)| k.as_deref())
            .and_then(|k| fonts.get(k))
        else {
            return;
        };
        let bytes = super::list_label::encode_label(entry, &para.list_label)
            .unwrap_or_else(|| entry.encode(&para.list_label));
        Some((entry, bytes))
    };

    if let Some(c) = para.label_color {
        fill_rgb(content, c);
    }
    content.begin_text();
    match &single {
        Some((entry, bytes)) => {
            content
                .set_font(Name(entry.pdf_name.as_bytes()), para.label_font_size)
                .next_line(label_x + para.label_shift, baseline_y)
                .show(Str(bytes));
        }
        None => {
            content.next_line(label_x + para.label_shift, baseline_y);
            super::list_label::show_label_pieces(
                content,
                &para.label_pieces,
                fonts,
                para.label_font_size,
            );
        }
    }
    content.end_text();
    if para.label_color.is_some() {
        content.set_fill_gray(0.0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn only_lines_level_with_a_nested_float_are_indented() {
        let para = CellParagraphLayout {
            lines: (0..3)
                .map(|_| super::super::layout::TextLine::default())
                .collect(),
            beside_indent: 90.0,
            beside_lines: 2,
            ..Default::default()
        };
        assert_eq!(
            beside_float_geometry(&para, 10.0, 400.0),
            Some(vec![(100.0, 310.0), (100.0, 310.0), (10.0, 400.0)])
        );
        let clear = CellParagraphLayout::default();
        assert_eq!(beside_float_geometry(&clear, 10.0, 400.0), None);
    }

    #[test]
    fn picture_clipping_does_not_clip_picture_free_text() {
        let para = CellParagraphLayout {
            font_size: 14.0,
            ascender_ratio: 0.8,
            image_line_ascent: 36.0,
            clip_inline_pictures: true,
            lines: vec![super::super::layout::TextLine::default()],
            ..Default::default()
        };
        let mut content = Content::new();
        assert!(!clip_fixed_cell_pictures(
            &mut content,
            &para,
            72.0,
            360.0,
            700.0,
            18.0
        ));
        assert!(content.finish().is_empty());
    }
}
