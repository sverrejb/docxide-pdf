//! SVG pictures as Form XObjects: the viewBox is stretched onto the unit square
//! (y flipped) so the draw sites' `w 0 0 h x y cm /Im Do` places it like any
//! raster or EMF picture.

use pdf_writer::types::{LineCapStyle, LineJoinStyle};
use pdf_writer::{Content, Pdf, Rect, Ref};

use super::color::{fill_rgb, stroke_rgb};
use crate::docx::svg::{IDENTITY, Seg, parse};

pub(super) fn svg_to_form_xobject(
    svg: &[u8],
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> Option<Ref> {
    let drawing = parse(svg)?;
    let [vx, vy, vw, vh] = drawing.view_box;
    let (sx, sy) = (1.0 / vw, 1.0 / vh);
    let mut c = Content::new();
    c.transform([sx, 0.0, 0.0, -sy, -vx * sx, 1.0 + vy * sy]);
    for item in &drawing.items {
        c.save_state();
        if item.transform != IDENTITY {
            c.transform(item.transform);
        }
        if let Some(rgb) = item.fill {
            fill_rgb(&mut c, rgb);
        }
        if let Some(rgb) = item.stroke {
            stroke_rgb(&mut c, rgb);
            c.set_line_width(item.stroke_width);
            c.set_line_cap(match item.cap {
                1 => LineCapStyle::RoundCap,
                2 => LineCapStyle::ProjectingSquareCap,
                _ => LineCapStyle::ButtCap,
            });
            c.set_line_join(match item.join {
                1 => LineJoinStyle::RoundJoin,
                2 => LineJoinStyle::BevelJoin,
                _ => LineJoinStyle::MiterJoin,
            });
            c.set_miter_limit(item.miter_limit);
            if !item.dash.is_empty() {
                c.set_dash_pattern(item.dash.iter().copied(), 0.0);
            }
        }
        for seg in &item.segs {
            match *seg {
                Seg::Move(x, y) => c.move_to(x, y),
                Seg::Line(x, y) => c.line_to(x, y),
                Seg::Cubic(x1, y1, x2, y2, x, y) => c.cubic_to(x1, y1, x2, y2, x, y),
                Seg::Close => c.close_path(),
            };
        }
        match (item.fill.is_some(), item.stroke.is_some(), item.even_odd) {
            (true, true, false) => c.fill_nonzero_and_stroke(),
            (true, true, true) => c.fill_even_odd_and_stroke(),
            (true, false, false) => c.fill_nonzero(),
            (true, false, true) => c.fill_even_odd(),
            (false, true, _) => c.stroke(),
            (false, false, _) => c.end_path(),
        };
        c.restore_state();
    }
    let form_ref = alloc();
    pdf.form_xobject(form_ref, &c.finish())
        .bbox(Rect::new(0.0, 0.0, 1.0, 1.0));
    Some(form_ref)
}
