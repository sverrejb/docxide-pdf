use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{
    ArrowEnd, ConnectorShape, ConnectorType, FloatingImage, HRelativeFrom, HorizontalPosition,
    SectionProperties, VRelativeFrom, VerticalPosition, WrapType,
};

use super::color::{fill_rgb, stroke_rgb};
use super::helpers::draw_circle;
use super::tagging::Tags;
use super::*;

pub(crate) fn resolve_h_position(
    h_relative_from: HRelativeFrom,
    h_position: &HorizontalPosition,
    obj_width: f32,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> f32 {
    let (origin, area_width) = match h_relative_from {
        HRelativeFrom::Page => (0.0, sp.page_width),
        HRelativeFrom::Column => (col_x, col_w),
        HRelativeFrom::Margin => (sp.margin_left, text_width),
    };
    h_position.place(origin, area_width, obj_width)
}

pub(super) fn resolve_fi_x(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> f32 {
    let x = resolve_h_position(
        fi.h_relative_from,
        &fi.h_position,
        fi.image.display_width,
        sp,
        col_x,
        col_w,
        text_width,
    );
    // A wrapping picture wider than the page starts at its left edge (probe
    // p20d, 621pt at column -71.5pt on a 612pt page: x 0; an external corpus
    // cover's 615pt picture on A4: x 0). One that fits keeps a negative x
    // (cyprus_ucits's 592pt banner at -2.7 on A4).
    if fi.wrap_type != WrapType::None && fi.image.display_width > sp.page_width {
        0.0
    } else {
        x
    }
}

/// True when a float wraps text and its box (with its text distances) reaches
/// into the column: a QR code sitting entirely in the margin never narrows text.
pub(super) fn wraps_in_column(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> bool {
    fi.wrap_type.wraps_beside()
        && side_gaps(fi, sp, col_x, col_w, text_width).is_some()
        && !image_blocks_band(fi, sp, col_x, col_w, text_width)
}

/// The free room left and right of a float in the column, None when the
/// float (with its text distances) stays out of the column.
pub(super) fn side_gaps(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> Option<(f32, f32)> {
    let (left, right) = raw_side_gaps(fi, sp, col_x, col_w, text_width);
    (left < col_w && right < col_w).then_some((left, right))
}

fn raw_side_gaps(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> (f32, f32) {
    let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
    let left = fi_x - fi.dist_left - col_x;
    let right = col_x + col_w - (fi_x + fi.image.display_width + fi.dist_right);
    (left, right)
}

/// A picture leaves no usable strip beside it at any height:
/// beside its wrap polygon for tight and through, as Word runs lines right
/// of a column-wide picture whose wrap points enclose only its left half
/// (example: x 314 beside a 468pt picture, 330 beside a 500pt one).
pub(super) fn no_strip_beside(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> bool {
    let (left, right) = raw_side_gaps(fi, sp, col_x, col_w, text_width);
    let (l, r) = polygon_edges(fi);
    (left + l).max(right + fi.image.display_width - r) < super::MIN_EMPTY_STRIP
}

/// The innermost left and right edges of a picture's wrap polygon over its
/// height, from the picture's left edge; it may reach past the picture (Word
/// examples: a 388pt picture whose polygon spans 485pt keeps a 468pt column's
/// lines out of its band, one in column one whose polygon covers column two
/// steps column two's lines too). Without a polygon, the picture's edges.
pub(super) fn polygon_edges(fi: &FloatingImage) -> (f32, f32) {
    let w = fi.image.display_width;
    let Some(verts) = fi.wrap_polygon.as_ref().filter(|v| v.len() >= 3) else {
        return (0.0, w);
    };
    let pts: Vec<(f32, f32)> = verts
        .iter()
        .map(|&(x, y)| (x as f32 / 21600.0 * w, y as f32))
        .collect();
    let mut ys: Vec<f32> = pts.iter().map(|p| p.1).collect();
    ys.sort_by(f32::total_cmp);
    ys.dedup();
    let (l, r) = ys
        .windows(2)
        .flat_map(|y| [y[0] + 0.5, (y[0] + y[1]) / 2.0, y[1] - 0.5])
        .filter_map(|y| super::poly_scanline(&pts, y))
        .fold((f32::NEG_INFINITY, f32::INFINITY), |(l, r), s| {
            (l.max(s.0), r.min(s.1))
        });
    if l.is_finite() { (l, r) } else { (0.0, w) }
}

/// A page- or margin-anchored picture that keeps body lines out of its whole
/// band (`frame_bands`) like the matching textboxes: topAndBottom, or
/// square/tight/through with no usable strip beside it (beside the wrap
/// polygon for tight and through).
pub(super) fn image_blocks_band(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> bool {
    if fi.v_relative_from == VRelativeFrom::Paragraph
        || side_gaps(fi, sp, col_x, col_w, text_width).is_none()
    {
        return false;
    }
    match fi.wrap_type {
        WrapType::TopAndBottom => true,
        WrapType::Square | WrapType::Tight | WrapType::Through => {
            no_strip_beside(fi, sp, col_x, col_w, text_width)
        }
        _ => false,
    }
}

/// The band of a picture that `image_blocks_band`, down from the page top.
pub(super) fn image_band(fi: &FloatingImage, sp: &SectionProperties) -> (f32, f32) {
    // The slot top only places paragraph-relative floats.
    let top = sp.page_height - resolve_fi_y_top(fi, sp, 0.0);
    (
        top - fi.dist_top,
        top + fi.image.display_height + fi.dist_bottom,
    )
}

pub(crate) fn resolve_fi_y_top(fi: &FloatingImage, sp: &SectionProperties, slot_top: f32) -> f32 {
    let img = &fi.image;
    match fi.v_position {
        VerticalPosition::Offset(v_offset) => match fi.v_relative_from {
            VRelativeFrom::Page | VRelativeFrom::TopMargin => sp.page_height - v_offset,
            VRelativeFrom::Margin => sp.page_height - sp.margin_top - v_offset,
            VRelativeFrom::BottomMargin => sp.margin_bottom - v_offset,
            VRelativeFrom::Paragraph => slot_top - v_offset,
        },
        VerticalPosition::AlignTop => match fi.v_relative_from {
            VRelativeFrom::Page | VRelativeFrom::TopMargin => sp.page_height,
            VRelativeFrom::BottomMargin => sp.margin_bottom,
            // sample500kB's picture tops its anchor paragraph, 89pt below the margin.
            VRelativeFrom::Paragraph => slot_top,
            _ => sp.page_height - sp.margin_top,
        },
        VerticalPosition::AlignCenter => match fi.v_relative_from {
            VRelativeFrom::Page => (sp.page_height + img.display_height) / 2.0,
            VRelativeFrom::TopMargin => sp.page_height - (sp.margin_top - img.display_height) / 2.0,
            VRelativeFrom::BottomMargin => (sp.margin_bottom + img.display_height) / 2.0,
            _ => {
                let area = sp.page_height - sp.margin_top - sp.margin_bottom;
                sp.page_height - sp.margin_top - (area - img.display_height) / 2.0
            }
        },
        VerticalPosition::AlignBottom => match fi.v_relative_from {
            VRelativeFrom::Page | VRelativeFrom::BottomMargin => img.display_height,
            VRelativeFrom::TopMargin => sp.page_height - sp.margin_top + img.display_height,
            _ => sp.margin_bottom + img.display_height,
        },
    }
}

/// Render floating images into per-image content buffers pushed onto their
/// layer's page z-stack (`behind` or `front`), tagged with their
/// relativeHeight. Matches how textboxes/connectors are deferred so that a
/// floating image with a higher relativeHeight paints above a shape with a
/// lower one (annotation #191 — a cursor icon was hidden behind a bar shape
/// because images painted into the page content stream, always below deferred
/// shapes).
pub(super) fn render_floating_images(
    floating_images: &[FloatingImage],
    global_block_idx: usize,
    pdf_names: &HashMap<(usize, usize), String>,
    effect_pdf_names: &HashMap<(usize, usize), super::images::EffectXObjs>,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
    anchor: super::ParaAnchor,
    behind: &mut Vec<(u32, Content)>,
    front: &mut Vec<(u32, Content)>,
    tags: &mut Tags,
    page: usize,
) {
    for (fi_idx, fi) in floating_images.iter().enumerate() {
        let mut buf = super::tagging::artifact_content();
        if render_one_floating_image(
            fi,
            fi_idx,
            global_block_idx,
            pdf_names,
            effect_pdf_names,
            sp,
            col_x,
            col_w,
            text_width,
            anchor,
            &mut buf,
            tags,
            page,
        ) {
            let layer = if fi.behind_doc {
                &mut *behind
            } else {
                &mut *front
            };
            layer.push((fi.z_index, buf));
        }
    }
}

/// Render a single floating image (with shadow/glow/clip/border effects) into
/// `content`. Returns true if an image was actually emitted (i.e. it had a
/// registered XObject name).
fn render_one_floating_image(
    fi: &FloatingImage,
    fi_idx: usize,
    global_block_idx: usize,
    pdf_names: &HashMap<(usize, usize), String>,
    effect_pdf_names: &HashMap<(usize, usize), super::images::EffectXObjs>,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
    anchor: super::ParaAnchor,
    content: &mut Content,
    // The picture is tagged as a Figure hoisted after its anchor, like Word;
    // `content` must be inside the default artifact.
    tags: &mut Tags,
    page: usize,
) -> bool {
    let Some(pdf_name) = pdf_names.get(&(global_block_idx, fi_idx)) else {
        return false;
    };
    let img = &fi.image;
    let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
    let fi_y_top = anchor.fi_y_top(fi, sp);
    let fi_y_bottom = fi_y_top - img.display_height;

    // OOXML rotates a floating image about the center of its (unrotated) box, with
    // the anchor offset and effectExtent already accounting for the turn. Rotate
    // about that center so the image's center stays put — matching Word. The sign
    // is negated because OOXML rotation is clockwise while PDF's is counterclockwise.
    let rotated = img.rotation_deg.abs() > 0.01;
    if rotated {
        content.save_state();
        push_center_rotation(
            content,
            fi_x,
            fi_y_bottom,
            img.display_width,
            img.display_height,
            img.rotation_deg,
        );
    }

    let fi_fx = effect_pdf_names.get(&(global_block_idx, fi_idx));
    let rect = (fi_x, fi_y_bottom, img.display_width, img.display_height);
    super::color::draw_image_halo(content, img.shadow.as_ref(), img.glow.as_ref(), rect, fi_fx);

    // Only the picture itself: its shadow, glow, reflection and border are decoration.
    if !img.decorative {
        let figure = tags.hoist_figure(img.alt.as_deref(), fi.anchor_seq);
        tags.begin(content, page, figure);
    }
    super::shape_path::render_image_with_clip(
        content,
        pdf_name,
        fi_x,
        fi_y_bottom,
        img.display_width,
        img.display_height,
        img.clip_geometry.as_ref(),
    );
    if !img.decorative {
        Tags::end(content);
    }

    super::color::draw_image_afterglow(
        content,
        img.inner_shadow.as_ref(),
        img.reflection.as_ref(),
        rect,
        fi_fx,
    );

    if let Some(sc) = img.stroke_color {
        super::shape_path::stroke_image_border(
            content,
            fi_x,
            fi_y_bottom,
            img.display_width,
            img.display_height,
            sc,
            img.stroke_width,
            img.clip_geometry.as_ref(),
        );
    }
    if rotated {
        content.restore_state();
    }
    true
}

/// Turn subsequent drawing by `deg` (OOXML clockwise degrees) about the centre of
/// the box at (`x`, `y_bottom`) with size `w`×`h`. PDF rotates counterclockwise,
/// hence the negation. Push before the box's own transform; the caller brackets
/// it with save/restore. No-op for an unrotated box.
pub(super) fn push_center_rotation(
    content: &mut Content,
    x: f32,
    y_bottom: f32,
    w: f32,
    h: f32,
    deg: f32,
) {
    if deg.abs() <= 0.01 {
        return;
    }
    let cx = x + w / 2.0;
    let cy = y_bottom + h / 2.0;
    let (sin, cos) = (-deg.to_radians()).sin_cos();
    content.transform([
        cos,
        sin,
        -sin,
        cos,
        cx - cos * cx + sin * cy,
        cy - sin * cx - cos * cy,
    ]);
}

/// Top-left corner (PDF coordinates) of a body or header/footer connector,
/// resolved like a textbox against its anchor's frames.
pub(super) fn connector_top_left(
    conn: &ConnectorShape,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
    slot_top: f32,
) -> (f32, f32) {
    let x = resolve_h_position(
        conn.h_relative_from,
        &conn.h_position,
        conn.width,
        sp,
        col_x,
        col_w,
        text_width,
    );
    let y = super::header_footer::resolve_tb_y_top(
        conn.v_relative_from,
        &conn.v_position,
        conn.height,
        sp,
        slot_top,
    );
    (x, y)
}

pub(super) fn render_connector(conn: &ConnectorShape, content: &mut Content, cx: f32, cy: f32) {
    content.save_state();
    stroke_rgb(content, conn.stroke_color);
    content.set_line_width(conn.stroke_width);

    match &conn.connector_type {
        ConnectorType::Line { flip_h, flip_v } => {
            let (x0, y0, x1, y1) = match (*flip_h, *flip_v) {
                (false, false) => (cx, cy, cx + conn.width, cy - conn.height),
                (true, false) => (cx + conn.width, cy, cx, cy - conn.height),
                (false, true) => (cx, cy - conn.height, cx + conn.width, cy),
                (true, true) => (cx + conn.width, cy - conn.height, cx, cy),
            };
            content.move_to(x0, y0);
            content.line_to(x1, y1);
            content.stroke();
            if conn.head_end != ArrowEnd::None {
                draw_arrow_head(
                    content,
                    x0,
                    y0,
                    x1,
                    y1,
                    conn.head_end,
                    conn.stroke_color,
                    conn.stroke_width,
                );
            }
            if conn.tail_end != ArrowEnd::None {
                draw_arrow_head(
                    content,
                    x1,
                    y1,
                    x0,
                    y0,
                    conn.tail_end,
                    conn.stroke_color,
                    conn.stroke_width,
                );
            }
        }
        ConnectorType::Arc {
            start_angle,
            end_angle,
            rotation,
        } => {
            render_arc(
                content,
                cx,
                cy,
                conn.width,
                conn.height,
                *start_angle,
                *end_angle,
                *rotation,
            );
        }
    }

    content.restore_state();
}

fn draw_arrow_head(
    content: &mut Content,
    tip_x: f32,
    tip_y: f32,
    other_x: f32,
    other_y: f32,
    end_type: ArrowEnd,
    color: [u8; 3],
    stroke_width: f32,
) {
    let dx = tip_x - other_x;
    let dy = tip_y - other_y;
    let len = (dx * dx + dy * dy).sqrt();
    if len < 0.001 {
        return;
    }
    let ux = dx / len;
    let uy = dy / len;
    // Word's "med" arrow: width ~4.5x stroke, length ~4.5x stroke (heuristic)
    let arrow_len = (stroke_width * 4.5).max(3.0);
    let arrow_half = (stroke_width * 2.5).max(1.5);

    // Base point: along the line, behind the tip
    let bx = tip_x - ux * arrow_len;
    let by = tip_y - uy * arrow_len;
    // Perpendicular
    let px = -uy;
    let py = ux;

    let p1x = bx + px * arrow_half;
    let p1y = by + py * arrow_half;
    let p2x = bx - px * arrow_half;
    let p2y = by - py * arrow_half;

    content.save_state();
    match end_type {
        ArrowEnd::Arrow => {
            // Open arrow: two strokes from tip
            content.move_to(p1x, p1y);
            content.line_to(tip_x, tip_y);
            content.line_to(p2x, p2y);
            content.stroke();
        }
        ArrowEnd::Triangle | ArrowEnd::Stealth => {
            fill_rgb(content, color);
            content.move_to(tip_x, tip_y);
            content.line_to(p1x, p1y);
            if matches!(end_type, ArrowEnd::Stealth) {
                let mid_x = bx + ux * (arrow_len * 0.35);
                let mid_y = by + uy * (arrow_len * 0.35);
                content.line_to(mid_x, mid_y);
            }
            content.line_to(p2x, p2y);
            content.close_path();
            content.fill_nonzero();
        }
        ArrowEnd::Diamond => {
            fill_rgb(content, color);
            let mid_x = tip_x - ux * (arrow_len * 0.5);
            let mid_y = tip_y - uy * (arrow_len * 0.5);
            let back_x = tip_x - ux * arrow_len;
            let back_y = tip_y - uy * arrow_len;
            content.move_to(tip_x, tip_y);
            content.line_to(mid_x + px * arrow_half, mid_y + py * arrow_half);
            content.line_to(back_x, back_y);
            content.line_to(mid_x - px * arrow_half, mid_y - py * arrow_half);
            content.close_path();
            content.fill_nonzero();
        }
        ArrowEnd::Oval => {
            fill_rgb(content, color);
            let cx = tip_x - ux * (arrow_len * 0.5);
            let cy = tip_y - uy * (arrow_len * 0.5);
            draw_circle(content, cx, cy, arrow_half);
            content.fill_nonzero();
        }
        ArrowEnd::None => {}
    }
    content.restore_state();
}

fn render_arc(
    content: &mut Content,
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    start_deg: f32,
    end_deg: f32,
    rotation_deg: f32,
) {
    let rx = w / 2.0;
    let ry = h / 2.0;
    let cx = x + rx;
    let cy = y - ry;

    // OOXML sweep: from start_deg (adj1) to end_deg (adj2), always positive
    let mut sweep_deg = end_deg - start_deg;
    if sweep_deg <= 0.0 {
        sweep_deg += 360.0;
    }

    // OOXML uses standard trig angles (0°=right, CCW positive) displayed
    // in y-down coords (so visually clockwise). For PDF y-up, negate angles.
    let math_start = (-(start_deg + rotation_deg)).to_radians();
    let total = -(sweep_deg).to_radians();

    if total.abs() < 0.001 {
        return;
    }

    // Approximate arc with cubic bezier segments (max 90° each)
    let n_segs = ((total.abs() / std::f32::consts::FRAC_PI_2).ceil() as usize).max(1);
    let step = total / n_segs as f32;

    let pt = |a: f32| -> (f32, f32) { (cx + rx * a.cos(), cy + ry * a.sin()) };

    let mut angle = math_start;
    let (sx, sy) = pt(angle);
    content.move_to(sx, sy);

    for _ in 0..n_segs {
        let a0 = angle;
        let a1 = angle + step;
        let alpha = 4.0 * (1.0 - (step / 2.0).cos()) / (step / 2.0).sin() / 3.0;
        let (x0, y0) = pt(a0);
        let (x3, y3) = pt(a1);
        let cp1x = x0 - alpha * rx * a0.sin();
        let cp1y = y0 + alpha * ry * a0.cos();
        let cp2x = x3 + alpha * rx * a1.sin();
        let cp2y = y3 - alpha * ry * a1.cos();
        content.cubic_to(cp1x, cp1y, cp2x, cp2y, x3, y3);
        angle = a1;
    }
    content.stroke();
}

/// The room beside a float on the side(s) its wrapText lets text use.
pub(super) fn side_room(wrap_text: WrapText, left: f32, right: f32) -> f32 {
    match wrap_text {
        WrapText::Left => left,
        WrapText::Right => right,
        WrapText::BothSides | WrapText::Largest => left.max(right),
    }
}

/// Whether a page- or margin-anchored textbox keeps body lines out of its
/// whole band, like a lifted frame: a topAndBottom box (issue #148: text
/// flows above a page-placed one and resumes below it), or one leaving no
/// room on its wrap side (massachusetts' secretary box at the right edge wraps
/// text on its right only, so its anchor's line and the date start below it).
pub(super) fn textbox_blocks_band(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> bool {
    if tb.v_relative_from == VRelativeFrom::Paragraph {
        return false;
    }
    match tb.wrap_type {
        WrapType::TopAndBottom => true,
        WrapType::Square | WrapType::Tight => {
            textbox_room_beside(tb, para, sp, (col_x, col_w), text_width)
                .is_some_and(|room| room < MIN_EMPTY_STRIP)
        }
        _ => false,
    }
}

fn textbox_x(
    tb: &crate::model::Textbox,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> f32 {
    let (rel, pos) = (tb.h_relative_from, &tb.h_position);
    resolve_h_position(rel, pos, tb.width_pt, sp, col_x, col_w, text_width)
}

/// The rectangle (left, top, width, height; y from the page bottom) a text
/// box or shape at `x`, `y_top` keeps text off: its extent turned to the
/// nearest quarter turn (Word examples: a 128×60pt box rotated 90° wraps as
/// 60×128 about its centre, one rotated 30° as its unrotated extent).
fn textbox_wrap_rect(
    tb: &crate::model::Textbox,
    x: f32,
    y_top: f32,
    h: f32,
) -> (f32, f32, f32, f32) {
    let quarter = (tb.rotation_deg.rem_euclid(180.0) / 90.0).round() as i32 % 2 == 1;
    if !quarter {
        return (x, y_top, tb.width_pt, h);
    }
    let (cx, cy) = (x + tb.width_pt / 2.0, y_top - h / 2.0);
    (cx - h / 2.0, cy + tb.width_pt / 2.0, h, tb.width_pt)
}

/// A wrapping float's horizontal span with its wrap distances and its
/// vertical extent (top, bottom; y from the page bottom, its anchor at the
/// top margin).
type WrapSpan = ((f32, f32), (f32, f32));

fn textbox_span(
    tb: &crate::model::Textbox,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
) -> WrapSpan {
    let anchor = sp.page_height - sp.margin_top;
    let x = textbox_x(tb, sp, col, text_width);
    let (rel, pos) = (tb.v_relative_from, &tb.v_position);
    let y = header_footer::resolve_tb_y_top(rel, pos, tb.height_pt, sp, anchor);
    let (left, top, w, h) = textbox_wrap_rect(tb, x, y, tb.height_pt);
    (
        (left - tb.dist_left, left + w + tb.dist_right),
        (top + tb.dist_top, top - h - tb.dist_bottom),
    )
}

fn picture_span(
    fi: &crate::model::FloatingImage,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> WrapSpan {
    let x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
    let y = resolve_fi_y_top(fi, sp, sp.page_height - sp.margin_top);
    (
        (x - fi.dist_left, x + fi.image.display_width + fi.dist_right),
        (
            y + fi.dist_top,
            y - fi.image.display_height - fi.dist_bottom,
        ),
    )
}

/// Whether text wraps beside a floating text box or shape (square, tight or
/// through) like beside a picture: Word needs a strip of `MIN_EMPTY_STRIP`
/// on a side its wrapText allows (examples: 20pt beside a box keeps its
/// anchor's empty line there, 17pt moves it below the box).
pub(super) fn textbox_wraps_beside(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
) -> bool {
    !tb.on_line
        && tb.wrap_type.wraps_beside()
        && textbox_room_beside(tb, para, sp, col, text_width)
            .is_some_and(|room| room >= MIN_EMPTY_STRIP)
}

/// The room beside a text box on the side(s) its wrapText lets text use,
/// None for a box that stays out of the column.
pub(super) fn textbox_room_beside(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    (col_x, col_w): (f32, f32),
    text_width: f32,
) -> Option<f32> {
    let col = (col_x, col_w);
    let ((own_left, own_right), (top, bottom)) = textbox_span(tb, sp, col, text_width);
    if own_left >= col_x + col_w || own_right <= col_x {
        return None;
    }
    // The widest strips (left, right) the column leaves beside the box, the
    // paragraph's other wrapping floats beside it taken out too: two boxes
    // side by side across the column leave no room, though each alone would.
    let others = para
        .textboxes
        .iter()
        .filter(|o| !std::ptr::eq(*o, tb) && !o.on_line && o.wrap_type.wraps_beside())
        .map(|o| textbox_span(o, sp, col, text_width))
        .chain(
            para.floating_images
                .iter()
                .filter(|fi| fi.wrap_type.wraps_beside())
                .map(|fi| picture_span(fi, sp, col, text_width)),
        )
        .filter(|&(_, (t, b))| t > bottom && b < top);
    let mut free = vec![(col_x, col_x + col_w)];
    for (l, r) in std::iter::once((own_left, own_right)).chain(others.map(|(x, _)| x)) {
        free = free
            .into_iter()
            .flat_map(|(a, z)| [(a, z.min(l)), (a.max(r), z)])
            .filter(|(a, z)| z > a)
            .collect();
    }
    let widest = |keep: &dyn Fn(f32, f32) -> bool| {
        free.iter()
            .filter(|&&(a, z)| keep(a, z))
            .map(|(a, z)| z - a)
            .fold(0.0, f32::max)
    };
    Some(side_room(
        tb.wrap_text,
        widest(&|_, z| z <= own_left + 0.01),
        widest(&|a, _| a >= own_right - 0.01),
    ))
}

/// The wrap zones of a paragraph's text boxes that `textbox_wraps_beside`,
/// its drawings anchored at `anchor`.
pub(super) fn textbox_zones(
    para: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    anchor: ParaAnchor,
    ctx: &RenderContext,
) -> Vec<FloatZone> {
    para.textboxes
        .iter()
        .filter(|tb| textbox_wraps_beside(tb, para, sp, col, text_width))
        .map(|tb| {
            let top = anchor.drawing_top(false, tb.v_from_line);
            textbox_zone(tb, sp, col, text_width, top, ctx)
        })
        .collect()
}

/// The wrap zone of a text box that `textbox_wraps_beside`, its anchor's top
/// at `anchor_top`.
fn textbox_zone(
    tb: &crate::model::Textbox,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    anchor_top: f32,
    ctx: &RenderContext,
) -> FloatZone {
    let h = textbox_render::textbox_height(tb, ctx);
    let (rel, pos) = (tb.v_relative_from, &tb.v_position);
    let y_top = header_footer::resolve_tb_y_top(rel, pos, h, sp, anchor_top);
    FloatZone::for_textbox(tb, textbox_x(tb, sp, col, text_width), y_top, h)
}

/// The band (down from the page top) of a textbox that `textbox_blocks_band`.
pub(super) fn blocking_textbox_band(
    tb: &crate::model::Textbox,
    para: &Paragraph,
    sp: &SectionProperties,
    col: (f32, f32),
    text_width: f32,
    ctx: &RenderContext,
) -> Option<Band> {
    if !textbox_blocks_band(tb, para, sp, col, text_width) {
        return None;
    }
    let h = textbox_render::textbox_height(tb, ctx);
    let top = sp.page_height
        - header_footer::resolve_tb_y_top(tb.v_relative_from, &tb.v_position, h, sp, 0.0);
    Some(Band::new(top, top + h + tb.dist_bottom))
}

/// A paragraph-relative topAndBottom box hanging below its anchor's text
/// keeps the following paragraphs out of its band; they fill the gap above
/// it first (issue #148: a group 27mm below its anchor's top leaves the next
/// three paragraphs above it). One reaching into the anchor's own text keeps
/// the anchor's reserve.
pub(super) fn hangs_below_anchor(
    tb: &crate::model::Textbox,
    p: &Paragraph,
    col_w: f32,
    ctx: &RenderContext,
) -> bool {
    tb.v_relative_from == VRelativeFrom::Paragraph
        && !tb.on_line
        && tb.wrap_type == WrapType::TopAndBottom
        && matches!(tb.v_position, VerticalPosition::Offset(o) if o > 0.0 && {
            let lines = paragraph_line_count(p, col_w, ctx).max(1) as f32;
            o >= p.space_before + lines * para_line_h(p, ctx)
        })
}

/// A paragraph's line height from its runs, or from its mark when no run
/// sizes the line.
pub(super) fn plain_line_h(p: &Paragraph, ctx: &RenderContext) -> f32 {
    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
    let (fs, lhr) = if lhr.is_none() {
        unsized_line_metrics(p, fs, ctx.fonts)
    } else {
        (fs, lhr)
    };
    resolve_line_h(p.line_spacing.unwrap_or(ctx.doc_line_spacing), fs, lhr)
}

pub(super) fn para_line_h(p: &Paragraph, ctx: &RenderContext) -> f32 {
    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
    resolve_line_h(p.line_spacing.unwrap_or(ctx.doc_line_spacing), fs, lhr)
}

/// Minimum side-strip width for a line box to sit beside a wrapping float.
/// Empirical: brazilian_logistics_study absorbs empty paragraphs beside a
/// float with ~42pt strips; sample500kB (image width == column width, 0pt
/// strips) stacks them below.
pub(super) const MIN_EMPTY_STRIP: f32 = 18.0;

pub(super) const MIN_TABLE_SIDE_ROOM: f32 = 72.0;

pub(crate) struct FloatTableExtent {
    pub(super) left: f32,
    pub(super) right: f32,
    pub(super) left_from_text: f32,
    pub(super) right_from_text: f32,
}

/// The wider of the two strips a column leaves beside a float.
pub(super) fn side_strip(
    (obj_left, obj_right): (f32, f32),
    (left_from_text, right_from_text): (f32, f32),
    (col_x, col_w): (f32, f32),
) -> f32 {
    (obj_left - left_from_text - col_x).max(col_x + col_w - obj_right - right_from_text)
}

#[derive(Clone)]
pub(crate) struct FloatZone {
    pub top_y: f32,
    pub bottom_y: f32,
    pub obj_left: f32,
    pub obj_right: f32,
    pub left_from_text: f32,
    pub right_from_text: f32,
    /// Polygon vertices in absolute page coords (PDF: x from left, y from bottom)
    pub polygon_pts: Option<Vec<(f32, f32)>>,
    pub wrap_text: WrapText,
    /// True when the zone was created by a paragraph-relative floating image
    /// (positionV relativeFrom="paragraph").  Paragraphs whose cursor is
    /// slightly above the zone should still be pushed below wide images.
    pub para_relative: bool,
    /// Set for a floating table: like a paragraph-relative float it sits
    /// tblpY below its anchor paragraph, which comes next in the flow.
    pub from_table: bool,
}

/// The text region right of a float's wrap edge `float_right`: Word starts
/// the line at the wrap edge or at the paragraph's left indent, whichever is
/// further right, for tables, pictures and text boxes alike (examples: a box
/// whose wrap edge is 139pt in starts lines with a 0pt and a 129pt indent at
/// its edge, a 200pt one at the indent).
pub(super) fn right_text_region(
    col_x: f32,
    col_right: f32,
    float_right: f32,
    para: &Paragraph,
) -> (f32, f32) {
    let left = (col_x + para.indent_left).max(float_right);
    (left, (col_right - para.indent_right - left).max(0.0))
}

/// The text region left of a float's wrap edge `float_left`: the paragraph's
/// box cut off at the edge (example: a right indent inside the float's area
/// takes nothing more off the line).
pub(super) fn left_text_region(
    col_x: f32,
    col_right: f32,
    float_left: f32,
    para: &Paragraph,
) -> (f32, f32) {
    let left = col_x + para.indent_left;
    (
        left,
        (float_left.min(col_right - para.indent_right) - left).max(0.0),
    )
}

impl FloatZone {
    pub(super) fn wrap_left(&self) -> f32 {
        self.obj_left - self.left_from_text
    }

    pub(super) fn wrap_right(&self) -> f32 {
        self.obj_right + self.right_from_text
    }

    pub(super) fn side_strip(&self, col: (f32, f32)) -> f32 {
        side_strip(
            (self.obj_left, self.obj_right),
            (self.left_from_text, self.right_from_text),
            col,
        )
    }

    /// The wrap area lies wholly in the margin beside the column (`col_x`, `col_w`).
    pub(super) fn outside_column(&self, (col_x, col_w): (f32, f32)) -> bool {
        self.wrap_left() >= col_x + col_w || self.wrap_right() <= col_x
    }

    /// `other`'s wrap area lies inside this one's, so it excludes no more text.
    pub(super) fn covers(&self, other: &FloatZone) -> bool {
        self.top_y >= other.top_y
            && self.bottom_y <= other.bottom_y
            && self.wrap_left() <= other.wrap_left()
            && self.wrap_right() >= other.wrap_right()
    }

    /// The exclusion zone of a wrapping float whose frame's left edge is `fi_x`
    /// and top edge `fi_y_top` (page coordinates, y from the bottom).
    pub(super) fn for_float(fi: &crate::model::FloatingImage, fi_x: f32, fi_y_top: f32) -> Self {
        let (w, h) = (fi.image.display_width, fi.image.display_height);
        FloatZone {
            top_y: fi_y_top + fi.dist_top,
            bottom_y: fi_y_top - h - fi.dist_bottom,
            obj_left: fi_x,
            obj_right: fi_x + w,
            left_from_text: fi.dist_left,
            right_from_text: fi.dist_right,
            polygon_pts: fi
                .wrap_polygon
                .as_ref()
                .map(|verts| convert_polygon_to_page_coords(verts, fi_x, fi_y_top, w, h)),
            wrap_text: fi.wrap_text,
            para_relative: fi.v_relative_from == VRelativeFrom::Paragraph,
            from_table: false,
        }
    }

    /// The exclusion zone of a text box or shape `h` tall whose frame's left
    /// edge is `x` and top edge `y_top`: the same model as a picture's.
    fn for_textbox(tb: &crate::model::Textbox, x: f32, y_top: f32, h: f32) -> Self {
        let (left, top, w, h) = textbox_wrap_rect(tb, x, y_top, h);
        let unturned = left == x && top == y_top;
        FloatZone {
            top_y: top + tb.dist_top,
            bottom_y: top - h - tb.dist_bottom,
            obj_left: left,
            obj_right: left + w,
            left_from_text: tb.dist_left,
            right_from_text: tb.dist_right,
            polygon_pts: tb
                .wrap_polygon
                .as_ref()
                .filter(|_| unturned)
                .map(|verts| convert_polygon_to_page_coords(verts, left, top, w, h)),
            wrap_text: tb.wrap_text,
            para_relative: tb.v_relative_from == VRelativeFrom::Paragraph,
            from_table: false,
        }
    }

    /// Returns (left_edge, right_edge) of the exclusion zone at the given Y.
    /// Falls back to rectangular bounds if no polygon or scanline misses.
    pub(super) fn exclusion_at_y(&self, y: f32) -> (f32, f32) {
        if let Some(ref pts) = self.polygon_pts
            && let Some((left, right)) = poly_scanline(pts, y)
        {
            return (left, right);
        }
        (self.obj_left, self.obj_right)
    }

    /// Exclusion over a whole line box, `top` to `bottom`: Word keeps a line
    /// clear of the polygon anywhere in its height (case42's wrap edges match
    /// it to 0.3pt), where one scanline at the line top lags a line behind
    /// a shape that widens downwards.
    pub(super) fn exclusion_in_band(&self, top: f32, bottom: f32) -> (f32, f32) {
        let Some(pts) = self.polygon_pts.as_ref() else {
            return (self.obj_left, self.obj_right);
        };
        // A polygon's extremes over a band lie on the band's edges or its vertices.
        let inner = pts.iter().map(|p| p.1).filter(|&y| y < top && y > bottom);
        [top, bottom]
            .into_iter()
            .chain(inner)
            .filter_map(|y| poly_scanline(pts, y))
            .reduce(|(l0, r0), (l1, r1)| (l0.min(l1), r0.max(r1)))
            .unwrap_or((self.obj_left, self.obj_right))
    }

    /// Narrow a paragraph's text box (`text_x`, `text_w`, `label_x`) to fit
    /// beside this floating object when `y`, the paragraph's first line top,
    /// is inside the zone. Leaves the box as it is otherwise, or when no side
    /// is wide enough.
    #[allow(clippy::too_many_arguments)]
    pub(super) fn narrow_paragraph(
        &self,
        y: f32,
        col_x: f32,
        col_w: f32,
        para: &Paragraph,
        text_x: &mut f32,
        text_w: &mut f32,
        label_x: &mut f32,
        legacy_indents: bool,
    ) {
        if !(y <= self.top_y && y > self.bottom_y) {
            return;
        }
        let col_right = col_x + col_w;
        let (ex_left, ex_right) = self.exclusion_at_y(y);
        let space_right = col_right - (ex_right + self.right_from_text);
        let space_left = (ex_left - self.left_from_text) - col_x;

        let float_left = ex_left - self.left_from_text;
        let float_right = ex_right + self.right_from_text;
        let (_, lw) = left_text_region(col_x, col_right, float_left, para);
        let (rx, rw) = right_text_region(col_x, col_right, float_right, para);
        let mut use_right_region = || {
            *text_w = rw.max(1.0);
            *text_x = rx;
            let first =
                first_line_region_x(rx, float_right, para, para.indent_hanging, legacy_indents);
            *label_x = first - para.indent_hanging + para.indent_first_line;
        };
        if self.wrap_text == WrapText::BothSides {
            // For bothSides, use the wider region as primary text width (dual
            // geometry handles both regions per-line).
            if rw > lw {
                use_right_region();
            } else if lw > 0.0 {
                *text_w = lw.max(1.0);
            }
        } else {
            match wrap_side(self.wrap_text, space_left, space_right) {
                Some(true) => use_right_region(),
                Some(false) => *text_w = lw.max(1.0),
                None => {}
            }
        }
    }
}

/// The side of a float a line beside it takes (true: the right), None when
/// the side its wrapText allows (for largest, the wider) leaves less than
/// `MIN_EMPTY_STRIP` (Word examples: a 30pt strip beside a wrapText=largest
/// picture holds lines, a 17pt one sends them below it).
pub(super) fn wrap_side(wrap_text: WrapText, left: f32, right: f32) -> Option<bool> {
    let right_side = match wrap_text {
        WrapText::Left => false,
        WrapText::Right => true,
        WrapText::BothSides | WrapText::Largest => right >= left,
    };
    let room = if right_side { right } else { left };
    (room >= MIN_EMPTY_STRIP).then_some(right_side)
}

/// Where a paragraph's first line region beside a float starts, `x` being
/// the region's start for the other lines and `edge` the float's wrap edge:
/// the first-line indent moves the line right of it, a hanging indent never
/// left of the edge (Word examples: left 36pt + first line 18pt starts 18pt
/// past the edge; left 129pt or 150pt with a 40pt hanging indent at the
/// edge). Before compat 15 a hanging indent's first line counts its indent
/// from the edge instead (examples: left 150pt, hanging 40pt starts 110pt past
/// it; french_sexual_health_youth_strategy's arrows beside its logo). The
/// renderer shifts the first line by the hanging indent, so the region
/// starts that much further right.
pub(super) fn first_line_region_x(
    x: f32,
    edge: f32,
    para: &Paragraph,
    hanging: f32,
    legacy: bool,
) -> f32 {
    if legacy && para.indent_hanging > 0.0 {
        edge + (para.indent_left - para.indent_hanging).max(0.0) + para.indent_hanging
    } else {
        x.max(edge + hanging.max(0.0))
    }
}

/// The per-line entries of the lines `blocked` does not mark.
pub(super) fn keep_open<T: Copy>(per_line: &[T], blocked: &[bool]) -> Vec<T> {
    let blocked = blocked.iter().copied().chain(std::iter::repeat(false));
    per_line
        .iter()
        .zip(blocked)
        .filter(|(_, b)| !b)
        .map(|(&v, _)| v)
        .collect()
}

/// `built`, laid out over the lines `blocked` does not mark, with an empty
/// line (`float_skip`) at each marked one: a line beside a float with no room
/// on a side text may take stays empty and the text goes on below it.
pub(super) fn skip_blocked_lines(blocked: &[bool], built: Vec<TextLine>) -> Vec<TextLine> {
    let mut built = built.into_iter().peekable();
    let mut lines = Vec::new();
    for &b in blocked {
        if built.peek().is_none() {
            break;
        }
        lines.push(match b {
            false => built.next().unwrap(),
            true => TextLine {
                float_skip: true,
                ..TextLine::default()
            },
        });
    }
    lines.extend(built);
    lines
}

pub(super) fn wrap_room(wrap_text: WrapText, left: f32, right: f32) -> f32 {
    if wrap_text == WrapText::BothSides {
        // A float overhanging a margin has no strip on that side, not a
        // negative one (an external tracked corpus cover 71.5pt into the left
        // margin left 147pt on the right, and the title went below the
        // page-tall picture).
        left.max(0.0) + right.max(0.0)
    } else {
        side_room(wrap_text, left, right)
    }
}

/// A floating table that leaves no room for a line beside it in the column.
fn table_blocks_lines(zone: &FloatZone, (col_x, col_w): (f32, f32)) -> bool {
    if !zone.from_table {
        return false;
    }
    let space_left = zone.wrap_left() - col_x;
    let space_right = col_x + col_w - zone.wrap_right();
    wrap_room(zone.wrap_text, space_left, space_right) < MIN_TABLE_SIDE_ROOM
}

/// Word moves the first line of a paragraph that reaches into a floating
/// table leaving no room beside it below the table, with no gap; the lines
/// above stay (examples: a line 2pt or 8pt into the table's wrap area goes
/// below it, exact and auto spacing alike). The block start is cleared
/// before layout, so only a later line can cross the top here. Judged where
/// the lines are placed: a paragraph moved to the next page by keepLines or
/// widow control keeps its lines together there. The table is one more band
/// after the page's (`step_lines_below_bands`): a line moved below it can
/// land in a band. Returns the sum of the lines' gaps.
pub(super) fn step_lines_below_floats(
    lines: &mut [TextLine],
    pb: &PageBuilder,
    col: (f32, f32),
    page_height: f32,
    first_top: f32,
    line_h: f32,
    metrics: (f32, f32),
) -> f32 {
    let table = pb
        .float_zone
        .as_ref()
        .filter(|z| table_blocks_lines(z, col))
        .map(|z| Band::new(page_height - z.top_y - 0.1, page_height - z.bottom_y));
    let Some(table) = table else {
        return step_lines_below_bands(
            lines,
            &pb.frame_bands,
            col,
            page_height,
            first_top,
            line_h,
            metrics,
        );
    };
    let bands: Vec<Band> = pb.frame_bands.iter().copied().chain([table]).collect();
    step_lines_below_bands(lines, &bands, col, page_height, first_top, line_h, metrics)
}

/// Scanline intersection: find the leftmost and rightmost x where polygon edges cross y.
pub(super) fn poly_scanline(pts: &[(f32, f32)], y: f32) -> Option<(f32, f32)> {
    let n = pts.len();
    if n < 3 {
        return None;
    }
    let mut min_x = f32::MAX;
    let mut max_x = f32::MIN;
    for i in 0..n {
        let (x0, y0) = pts[i];
        let (x1, y1) = pts[(i + 1) % n];
        if (y0 <= y && y1 >= y) || (y1 <= y && y0 >= y) {
            if (y1 - y0).abs() < 0.001 {
                min_x = min_x.min(x0).min(x1);
                max_x = max_x.max(x0).max(x1);
            } else {
                let t = (y - y0) / (y1 - y0);
                let x = x0 + t * (x1 - x0);
                min_x = min_x.min(x);
                max_x = max_x.max(x);
            }
        }
    }
    if min_x <= max_x {
        Some((min_x, max_x))
    } else {
        None
    }
}

/// Convert polygon vertices from 1/21600-of-extent coords to absolute page coords.
fn convert_polygon_to_page_coords(
    vertices: &[(i32, i32)],
    img_x: f32,
    img_y_top: f32,
    display_w: f32,
    display_h: f32,
) -> Vec<(f32, f32)> {
    vertices
        .iter()
        .map(|&(px, py)| {
            let x_pt = img_x + (px as f32 / 21600.0) * display_w;
            let y_pt = img_y_top - (py as f32 / 21600.0) * display_h;
            (x_pt, y_pt)
        })
        .collect()
}

/// Draw debug overlays for the wrap polygon and effective exclusion zone.
/// Green = raw polygon, blue = wrap zone (polygon ± dist margins), red = top/bottom bounds.
pub(super) fn draw_debug_wrap_overlay(content: &mut Content, fz: &FloatZone) {
    let polyline = |content: &mut Content, pts: &[(f32, f32)], close: bool| {
        content.move_to(pts[0].0, pts[0].1);
        for &(x, y) in &pts[1..] {
            content.line_to(x, y);
        }
        if close {
            content.close_path();
        }
        content.stroke();
    };
    content.save_state();

    // Green: raw polygon outline
    if let Some(ref pts) = fz.polygon_pts
        && pts.len() >= 3
    {
        content.set_stroke_rgb(0.0, 0.7, 0.0);
        content.set_line_width(0.5);
        polyline(content, pts, true);
    }

    // Blue: effective wrap zone boundaries (polygon shifted by dist margins)
    if let Some(ref pts) = fz.polygon_pts {
        content.set_stroke_rgb(0.0, 0.0, 0.8);
        content.set_line_width(0.3);
        let steps = ((fz.top_y - fz.bottom_y) / 1.0).ceil() as usize;
        if steps > 0 {
            let mut left_pts = Vec::with_capacity(steps + 1);
            let mut right_pts = Vec::with_capacity(steps + 1);
            for i in 0..=steps {
                let y = fz.top_y - i as f32 * (fz.top_y - fz.bottom_y) / steps as f32;
                if let Some((l, r)) = poly_scanline(pts, y) {
                    left_pts.push((l - fz.left_from_text, y));
                    right_pts.push((r + fz.right_from_text, y));
                }
            }
            if left_pts.len() >= 2 {
                polyline(content, &left_pts, false);
            }
            if right_pts.len() >= 2 {
                polyline(content, &right_pts, false);
            }
        }
    }

    // Red: top and bottom zone boundaries
    content.set_stroke_rgb(0.8, 0.0, 0.0);
    content.set_line_width(0.3);
    let page_left = fz.obj_left - 20.0;
    let page_right = fz.obj_right + 20.0;
    polyline(
        content,
        &[(page_left, fz.top_y), (page_right, fz.top_y)],
        false,
    );
    polyline(
        content,
        &[(page_left, fz.bottom_y), (page_right, fz.bottom_y)],
        false,
    );

    content.restore_state();
}

pub(super) struct FloatingTablePos {
    pub x: f32,
    pub y: f32,
    pub top_from_text: f32,
    pub bottom_from_text: f32,
    pub left_from_text: f32,
    pub right_from_text: f32,
    /// Raw `tblpY` offset in points (positive = down from the anchor). Needed by
    /// the flushed-to-next-page path, which re-anchors a `vertAnchor="text"`
    /// table to the new page body top and must re-apply this offset so the table
    /// sits where Word puts it (often above the top margin for negative tblpY).
    pub v_offset_pt: f32,
    /// True when `vertAnchor="text"` — the offset is relative to the anchor
    /// paragraph (re-applicable on a fresh page), not the page/margin.
    pub v_anchor_text: bool,
}

impl FloatingTablePos {
    /// Preserve the text distance when a non-overlapping table moves below a float.
    pub(super) fn constrain_nonoverlap_left(
        &mut self,
        pos: &crate::model::TablePosition,
        sp: &SectionProperties,
        col_x: f32,
        compat_mode: u32,
    ) {
        if compat_mode >= 15
            && !pos.allow_overlap
            && matches!(pos.h_position, crate::model::HorizontalPosition::Offset(v) if v >= 0.0)
            && matches!(pos.h_anchor, "margin" | "column")
        {
            let text_left = if pos.h_anchor == "margin" {
                sp.margin_left
            } else {
                col_x
            };
            self.x = self.x.max(text_left + pos.left_from_text);
        }
    }

    /// Where a floating table goes: `tblpX`/`tblpXSpec` against its anchor
    /// column (`col_x`, `col_w`), `tblpY` below the page, the margin or the
    /// text (`text_y`: where its anchor paragraph's flow is).
    pub(super) fn resolve(
        table: &crate::model::Table,
        pos: &crate::model::TablePosition,
        sp: &SectionProperties,
        col_x: f32,
        col_w: f32,
        text_y: f32,
        ctx: &RenderContext,
    ) -> Self {
        let h_relative_from = match pos.h_anchor {
            "page" => HRelativeFrom::Page,
            "margin" => HRelativeFrom::Margin,
            _ => HRelativeFrom::Column,
        };
        let x = resolve_h_position(
            h_relative_from,
            &pos.h_position,
            table.col_widths.iter().sum(),
            sp,
            col_x,
            col_w,
            sp.text_width(),
        );
        // Before compat 15 an offset places the first cell's text, not
        // its border, as tblInd does inline: Word draws case46's R1 at
        // the margin and the border a cell margin left of it.
        // From compat 15 the offset places the left border's outer edge, as
        // tblInd does inline (Word probe: 0.5/1.5/6pt borders put floating and
        // inline tables' first-column text at the same x).
        // A right-aligned one mirrors it: its last cell's text ends at the
        // margin and its border a cell margin past (case40, case45).
        use crate::model::HorizontalPosition::{AlignRight, Offset};
        let x = match pos.h_position {
            Offset(_) if ctx.compat_mode < 15 => x - table.first_cell_left_margin(),
            Offset(_) => x + table::word2013_border_shift(table),
            AlignRight if ctx.compat_mode < 15 => x + table.last_cell_right_margin(),
            _ => x,
        };
        // Word puts a page or margin anchored table one twip above a
        // positive tblpY, and never above the page (Word for Mac probes:
        // tblpY 1, 3, 946 and 957 put the first line 0, 0.1, 47.25 and 47.8pt
        // down, tblpY −100 at the page top; text anchors are exact).
        let anchored = if pos.v_offset_pt > 0.0 {
            pos.v_offset_pt - 0.05
        } else {
            pos.v_offset_pt
        };
        let y = match pos.v_anchor {
            "page" => sp.page_height - anchored.max(0.0),
            "margin" => sp.page_height - sp.margin_top - anchored,
            _ => text_y - pos.v_offset_pt,
        };
        FloatingTablePos {
            x,
            y,
            top_from_text: pos.top_from_text,
            bottom_from_text: pos.bottom_from_text,
            left_from_text: pos.left_from_text,
            right_from_text: pos.right_from_text,
            v_offset_pt: pos.v_offset_pt,
            v_anchor_text: pos.v_anchor == "text",
        }
    }
}

// Full-width rectangular contour wrapping advances through candidate lines,
// rather than applying the square-wrap jump to image/effect bounds.
pub(super) struct RectangularClearance {
    pub(super) top: f32,
    pub(super) bottom: f32,
    pub(super) pitch: f32,
    pub(super) line_h: f32,
}

impl RectangularClearance {
    pub(super) fn cleared_y(&self, start: f32) -> f32 {
        let steps = ((start - self.bottom) / self.pitch).ceil().max(0.0);
        start - steps * self.pitch
    }
}

fn rectangular_polygon_bounds(pts: &[(f32, f32)]) -> Option<(f32, f32, f32, f32)> {
    if !(4..=5).contains(&pts.len()) {
        return None;
    }
    let left = pts.iter().map(|p| p.0).fold(f32::INFINITY, f32::min);
    let right = pts.iter().map(|p| p.0).fold(f32::NEG_INFINITY, f32::max);
    let bottom = pts.iter().map(|p| p.1).fold(f32::INFINITY, f32::min);
    let top = pts.iter().map(|p| p.1).fold(f32::NEG_INFINITY, f32::max);
    if left >= right || bottom >= top {
        return None;
    }
    for (i, &(x, y)) in pts.iter().enumerate() {
        if !(x == left || x == right) || !(y == bottom || y == top) {
            return None;
        }
        let next = pts[(i + 1) % pts.len()];
        if x != next.0 && y != next.1 {
            return None;
        }
    }
    [(left, bottom), (right, bottom), (left, top), (right, top)]
        .iter()
        .all(|c| pts.contains(c))
        .then_some((left, right, top, bottom))
}

pub(super) fn rectangular_polygon_clearance(
    block: &Block,
    zone: &FloatZone,
    prev_after: f32,
    ctx: &RenderContext,
    (col_x, col_w): (f32, f32),
) -> Option<RectangularClearance> {
    if zone.from_table || !zone.para_relative || prev_after != 0.0 {
        return None;
    }
    let (left, right, top, bottom) = rectangular_polygon_bounds(zone.polygon_pts.as_ref()?)?;
    if side_strip(
        (left, right),
        (zone.left_from_text, zone.right_from_text),
        (col_x, col_w),
    ) >= MIN_EMPTY_STRIP
    {
        return None;
    }
    let Block::Paragraph(p) = block else {
        return None;
    };
    if p.space_before != 0.0
        || p.space_before_auto
        || p.clears_floats
        || p.image.is_some()
        || p.inline_chart.is_some()
        || !p.smartart.is_empty()
        || !p.floating_images.is_empty()
        || !p.textboxes.is_empty()
        || p.frame_props.is_some()
        || p.runs
            .iter()
            .any(|r| r.inline_image.is_some() || r.field_code.is_some() || r.math.is_some())
    {
        return None;
    }
    let ls = p.line_spacing.unwrap_or(ctx.doc_line_spacing);
    if !matches!(ls, LineSpacing::Auto(_)) {
        return None;
    }
    let (fs, lhr, _) = tallest_run_metrics(&p.runs, ctx.fonts);
    let (fs, lhr) = if is_text_empty(&p.runs) {
        unsized_line_metrics(p, fs, ctx.fonts)
    } else {
        (fs, lhr)
    };
    let pitch = resolve_line_h(LineSpacing::Auto(1.0), fs, lhr);
    let line_h = resolve_line_h(ls, fs, lhr);
    (pitch > 0.0 && pitch.is_finite()).then_some(RectangularClearance {
        top,
        bottom,
        pitch,
        line_h,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Block;

    /// education_consultant_posting's header rule is page-anchored 60.6pt
    /// down; Word draws it there, not 60.6pt below its paragraph.
    #[test]
    fn page_anchored_connector_ignores_its_paragraph() {
        let bytes =
            include_bytes!("../../tests/fixtures/scraped/education_consultant_posting/input.docx");
        let doc = crate::docx::parse_bytes(bytes, crate::RevisionView::Final).unwrap();
        let sp = &doc.sections[0].properties;
        let conn = sp
            .header_default
            .as_ref()
            .unwrap()
            .blocks
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => p.connectors.first(),
                _ => None,
            })
            .next()
            .unwrap();
        assert_eq!(conn.v_relative_from, VRelativeFrom::Page);
        let tw = sp.text_width();
        let (_, y) = connector_top_left(conn, sp, sp.margin_left, tw, tw, 500.0);
        assert!((sp.page_height - y - 60.6).abs() < 0.1);
    }

    #[test]
    fn contour_clearance_keeps_the_natural_line_phase() {
        let rect = RectangularClearance {
            top: 80.0,
            bottom: 15.0,
            pitch: 10.0,
            line_h: 15.0,
        };
        assert_eq!(rect.cleared_y(100.0), 10.0);
        assert_eq!(rect.cleared_y(98.0), 8.0);
        assert_eq!(rect.cleared_y(15.0), 15.0);
        assert_eq!(rect.cleared_y(10.0), 10.0);
    }

    #[test]
    fn rectangle_bounds_reject_diagonal_and_nonrectangular_contours() {
        let rect = [
            (0.0, 0.0),
            (0.0, 20.0),
            (30.0, 20.0),
            (30.0, 0.0),
            (0.0, 0.0),
        ];
        assert_eq!(
            rectangular_polygon_bounds(&rect),
            Some((0.0, 30.0, 20.0, 0.0))
        );
        assert_eq!(
            rectangular_polygon_bounds(&rect[..4]),
            Some((0.0, 30.0, 20.0, 0.0))
        );
        assert!(
            rectangular_polygon_bounds(&[(0.0, 0.0), (30.0, 20.0), (0.0, 20.0), (30.0, 0.0)])
                .is_none()
        );
        assert!(
            rectangular_polygon_bounds(&[(0.0, 0.0), (0.0, 20.0), (15.0, 10.0), (30.0, 0.0)])
                .is_none()
        );
    }
}

#[cfg(test)]
mod table_wrap_indent_tests {
    use super::*;

    #[test]
    fn right_table_strip_intersects_margin_relative_paragraph_box() {
        let mut para = Paragraph {
            indent_left: 255.0,
            indent_right: 12.0,
            ..Default::default()
        };
        assert_eq!(right_text_region(78.0, 552.0, 299.0, &para), (333.0, 207.0));
        para.indent_left = 100.0;
        assert_eq!(right_text_region(78.0, 552.0, 299.0, &para), (299.0, 241.0));
        para.indent_left = -10.0;
        assert_eq!(right_text_region(78.0, 552.0, 299.0, &para), (299.0, 241.0));
        para.indent_left = 500.0;
        assert_eq!(right_text_region(78.0, 552.0, 299.0, &para), (578.0, 0.0));
    }
}
