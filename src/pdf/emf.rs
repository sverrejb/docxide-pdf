//! Minimal EMF → PDF Form XObject translator.
//!
//! Walks the EMF record stream from `crate::docx::emf` and emits PDF drawing
//! operators into a Form XObject whose bbox is `[0 0 1 1]`. The Form XObject
//! is then placed by the existing image-rendering pipeline, which scales it
//! by the inline image's `display_width × display_height`.

use std::collections::HashMap;

use pdf_writer::{Content, Pdf, Rect, Ref};

use super::color::{fill_rgb, stroke_rgb};
use crate::docx::emf::{EmfRecord, FillRule, for_each_record, parse_header};

#[derive(Clone, Copy)]
enum EmfObject {
    Brush { color: [u8; 3] },
    Pen { color: [u8; 3], width: i32 },
}

#[derive(Clone)]
struct EmfState {
    window_org: (i32, i32),
    window_ext: (i32, i32),
    viewport_org: (i32, i32),
    viewport_ext: (i32, i32),
    current_pt: (i32, i32),
    fill_rule: FillRule,
    selected_brush: Option<[u8; 3]>,
    selected_pen: Option<([u8; 3], i32)>,
    in_path: bool,
}

impl EmfState {
    fn new(default_ext: (i32, i32)) -> Self {
        Self {
            window_org: (0, 0),
            window_ext: default_ext,
            viewport_org: (0, 0),
            viewport_ext: default_ext,
            current_pt: (0, 0),
            fill_rule: FillRule::Alternate,
            selected_brush: None,
            selected_pen: None,
            in_path: false,
        }
    }
}

/// Compose the EMF → form-XObject (1×1 unit box) mapping for a logical point.
struct Mapper {
    bounds: (i32, i32, i32, i32),
}

impl Mapper {
    fn map(&self, state: &EmfState, x: i32, y: i32) -> (f32, f32) {
        // logical → device
        let dx = ((x - state.window_org.0) as i64 * state.viewport_ext.0 as i64)
            .checked_div(state.window_ext.0 as i64)
            .unwrap_or(0) as i32
            + state.viewport_org.0;
        let dy = ((y - state.window_org.1) as i64 * state.viewport_ext.1 as i64)
            .checked_div(state.window_ext.1 as i64)
            .unwrap_or(0) as i32
            + state.viewport_org.1;
        // device → form [0,1] with Y flip (PDF is Y-up, EMF is Y-down).
        let (bl, bt, br, bb) = self.bounds;
        let w = (br - bl).max(1) as f32;
        let h = (bb - bt).max(1) as f32;
        let u = (dx - bl) as f32 / w;
        let v = 1.0 - (dy - bt) as f32 / h;
        (u, v)
    }
}

/// Translate an EMF byte stream into a PDF Form XObject. Returns the
/// allocated Ref, or `None` if the input isn't a parseable EMF.
pub(super) fn emf_to_form_xobject(
    emf: &[u8],
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> Option<Ref> {
    let header = parse_header(emf)?;
    let bounds_size = header.bounds_size();
    if bounds_size.0 == 0 || bounds_size.1 == 0 {
        return None;
    }
    let mapper = Mapper { bounds: header.bounds };

    let mut state = EmfState::new(bounds_size);
    let mut state_stack: Vec<EmfState> = Vec::new();
    let mut objects: HashMap<u32, EmfObject> = HashMap::new();
    let mut content = Content::new();

    for_each_record(emf, |rec| {
        translate_record(rec, &mut state, &mut state_stack, &mut objects, &mapper, &mut content);
        true
    });

    let form_ref = alloc();
    let bytes = content.finish();
    let mut form = pdf.form_xobject(form_ref, &bytes);
    form.bbox(Rect::new(0.0, 0.0, 1.0, 1.0));
    drop(form);
    Some(form_ref)
}

fn translate_record(
    rec: &EmfRecord,
    state: &mut EmfState,
    stack: &mut Vec<EmfState>,
    objects: &mut HashMap<u32, EmfObject>,
    mapper: &Mapper,
    content: &mut Content,
) {
    use EmfRecord::*;
    match rec {
        Header | Eof | Skip => {}
        SetMapMode(_) | SetBkMode(_) => {} // We honour window/viewport explicitly.
        SetPolyFillMode(rule) => state.fill_rule = *rule,
        SetWindowOrgEx(x, y) => state.window_org = (*x, *y),
        SetWindowExtEx(x, y) => state.window_ext = ((*x).max(1), (*y).max(1)),
        SetViewportOrgEx(x, y) => state.viewport_org = (*x, *y),
        SetViewportExtEx(x, y) => state.viewport_ext = ((*x).max(1), (*y).max(1)),
        SaveDc => {
            stack.push(state.clone());
            content.save_state();
        }
        RestoreDc => {
            if let Some(prev) = stack.pop() {
                *state = prev;
            }
            content.restore_state();
        }
        CreateBrushIndirect { handle, color } => {
            objects.insert(*handle, EmfObject::Brush { color: *color });
        }
        ExtCreatePen { handle, color, width } => {
            objects.insert(*handle, EmfObject::Pen { color: *color, width: *width });
        }
        DeleteObject(handle) => {
            objects.remove(handle);
        }
        SelectObject(handle) => match objects.get(handle) {
            Some(EmfObject::Brush { color }) => state.selected_brush = Some(*color),
            Some(EmfObject::Pen { color, width }) => state.selected_pen = Some((*color, *width)),
            None => {}
        },
        BeginPath => state.in_path = true,
        EndPath => state.in_path = false,
        MoveToEx(x, y) => {
            let (u, v) = mapper.map(state, *x, *y);
            content.move_to(u, v);
            state.current_pt = (*x, *y);
        }
        LineTo(x, y) => {
            let (u, v) = mapper.map(state, *x, *y);
            content.line_to(u, v);
            state.current_pt = (*x, *y);
        }
        PolyBezierTo16(pts) => {
            for triple in pts.chunks_exact(3) {
                let (c1x, c1y) = mapper.map(state, triple[0].0 as i32, triple[0].1 as i32);
                let (c2x, c2y) = mapper.map(state, triple[1].0 as i32, triple[1].1 as i32);
                let (ex, ey) = mapper.map(state, triple[2].0 as i32, triple[2].1 as i32);
                content.cubic_to(c1x, c1y, c2x, c2y, ex, ey);
                state.current_pt = (triple[2].0 as i32, triple[2].1 as i32);
            }
        }
        PolyLineTo16(pts) => {
            for p in pts {
                let (u, v) = mapper.map(state, p.0 as i32, p.1 as i32);
                content.line_to(u, v);
                state.current_pt = (p.0 as i32, p.1 as i32);
            }
        }
        CloseFigure => {
            content.close_path();
        }
        FillPath => {
            if let Some(c) = state.selected_brush {
                fill_rgb(content, c);
            }
            match state.fill_rule {
                FillRule::Alternate => {
                    content.fill_even_odd();
                }
                FillRule::Winding => {
                    content.fill_nonzero();
                }
            }
        }
        StrokePath => {
            if let Some((c, w)) = state.selected_pen {
                stroke_rgb(content, c);
                if w > 0 {
                    content.set_line_width(stroke_width_in_form(state, mapper, w));
                }
            }
            content.stroke();
        }
        StrokeAndFillPath => {
            if let Some(c) = state.selected_brush {
                fill_rgb(content, c);
            }
            if let Some((c, w)) = state.selected_pen {
                stroke_rgb(content, c);
                if w > 0 {
                    content.set_line_width(stroke_width_in_form(state, mapper, w));
                }
            }
            match state.fill_rule {
                FillRule::Alternate => {
                    content.fill_even_odd_and_stroke();
                }
                FillRule::Winding => {
                    content.fill_nonzero_and_stroke();
                }
            }
        }
    }
}

/// EMF pen widths are in logical units. The form XObject is 1×1 unit; the
/// emitted line-width operator runs in that space. Convert via the
/// logical→device→form mapping at the origin and at (w, 0), then take the
/// horizontal delta as a rough scalar.
fn stroke_width_in_form(state: &EmfState, mapper: &Mapper, w: i32) -> f32 {
    let (u0, _) = mapper.map(state, 0, 0);
    let (u1, _) = mapper.map(state, w, 0);
    (u1 - u0).abs().max(0.0005)
}
