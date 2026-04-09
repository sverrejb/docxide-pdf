use std::collections::HashMap;

use pdf_writer::{Content, Name};

use crate::model::{
    ConnectorShape, ConnectorType, FloatingImage, HRelativeFrom, HorizontalPosition,
    SectionProperties, VRelativeFrom, VerticalPosition,
};

use super::color::stroke_rgb;

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
    match h_position {
        HorizontalPosition::AlignCenter => origin + (area_width - obj_width) / 2.0,
        HorizontalPosition::AlignRight => origin + area_width - obj_width,
        HorizontalPosition::AlignLeft => origin,
        HorizontalPosition::Offset(o) => origin + o,
    }
}

pub(super) fn resolve_fi_x(
    fi: &FloatingImage,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
) -> f32 {
    resolve_h_position(
        fi.h_relative_from,
        &fi.h_position,
        fi.image.display_width,
        sp,
        col_x,
        col_w,
        text_width,
    )
}

pub(crate) fn resolve_fi_y_top(
    fi: &FloatingImage,
    sp: &SectionProperties,
    slot_top: f32,
) -> f32 {
    let img = &fi.image;
    match fi.v_position {
        VerticalPosition::Offset(v_offset) => match fi.v_relative_from {
            VRelativeFrom::Page => sp.page_height - v_offset,
            VRelativeFrom::Margin | VRelativeFrom::TopMargin => {
                sp.page_height - sp.margin_top - v_offset
            }
            VRelativeFrom::Paragraph => slot_top - v_offset,
        },
        VerticalPosition::AlignTop => match fi.v_relative_from {
            VRelativeFrom::Page => sp.page_height,
            _ => sp.page_height - sp.margin_top,
        },
        VerticalPosition::AlignCenter => match fi.v_relative_from {
            VRelativeFrom::Page => (sp.page_height + img.display_height) / 2.0,
            _ => {
                let area = sp.page_height - sp.margin_top - sp.margin_bottom;
                sp.page_height - sp.margin_top - (area - img.display_height) / 2.0
            }
        },
        VerticalPosition::AlignBottom => match fi.v_relative_from {
            VRelativeFrom::Page => img.display_height,
            _ => sp.margin_bottom + img.display_height,
        },
    }
}

pub(super) fn render_floating_images(
    floating_images: &[FloatingImage],
    behind_doc: bool,
    global_block_idx: usize,
    pdf_names: &HashMap<(usize, usize), String>,
    effect_pdf_names: &HashMap<(usize, usize), super::images::EffectXObjs>,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
    slot_top: f32,
    content: &mut Content,
) {
    for (fi_idx, fi) in floating_images.iter().enumerate() {
        if fi.behind_doc != behind_doc {
            continue;
        }
        if let Some(pdf_name) = pdf_names.get(&(global_block_idx, fi_idx)) {
            let img = &fi.image;
            let fi_x = resolve_fi_x(fi, sp, col_x, col_w, text_width);
            let fi_y_top = resolve_fi_y_top(fi, sp, slot_top);
            let fi_y_bottom = fi_y_top - img.display_height;

            let fi_fx = effect_pdf_names.get(&(global_block_idx, fi_idx));
            if let Some(ref shadow) = img.shadow {
                super::color::draw_image_shadow(
                    content, shadow, fi_x, fi_y_bottom,
                    img.display_width, img.display_height,
                    fi_fx.and_then(|fx| fx.shadow.as_deref()),
                );
            }
            if let Some(ref glow) = img.glow {
                super::color::draw_image_glow(
                    content, glow, fi_x, fi_y_bottom,
                    img.display_width, img.display_height,
                    fi_fx.and_then(|fx| fx.glow.as_deref()),
                );
            }

            content.save_state();
            content.transform([
                img.display_width,
                0.0,
                0.0,
                img.display_height,
                fi_x,
                fi_y_bottom,
            ]);
            content.x_object(Name(pdf_name.as_bytes()));
            content.restore_state();

            if let Some(ref inner) = img.inner_shadow {
                super::color::draw_inner_shadow(
                    content, inner, fi_x, fi_y_bottom,
                    img.display_width, img.display_height,
                    fi_fx.and_then(|fx| fx.inner_shadow.as_deref()),
                );
            }
            if let Some(ref refl) = img.reflection {
                super::color::draw_reflection(
                    content, refl, fi_x, fi_y_bottom,
                    img.display_width, img.display_height,
                    fi_fx.and_then(|fx| fx.reflection.as_deref()),
                );
            }

            if let Some(sc) = img.stroke_color {
                content.save_state();
                stroke_rgb(content, sc);
                content.set_line_width(img.stroke_width);
                content.rect(fi_x, fi_y_bottom, img.display_width, img.display_height);
                content.stroke();
                content.restore_state();
            }
        }
    }
}

pub(super) fn render_connector(
    conn: &ConnectorShape,
    content: &mut Content,
    col_x: f32,
    slot_top: f32,
) {
    let cx = col_x + conn.x;
    let cy = slot_top - conn.y;

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
