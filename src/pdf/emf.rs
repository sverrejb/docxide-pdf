//! Minimal EMF → PDF Form XObject translator.
//!
//! Walks the EMF record stream from `crate::docx::emf` and emits PDF drawing
//! operators into a Form XObject whose bbox is `[0 0 1 1]`. The Form XObject
//! is then placed by the existing image-rendering pipeline, which scales it
//! by the inline image's `display_width × display_height`.

use std::collections::{BTreeMap, HashMap};

use pdf_writer::{Content, Name, Pdf, Rect, Ref, Str};

use super::color::{fill_rgb, stroke_rgb};
use crate::docx::emf::{EmfRecord, FillRule, for_each_record, parse_header};
use crate::fonts::{FontEntry, font_key_for};

#[derive(Clone)]
enum EmfObject {
    Brush(Option<[u8; 3]>),
    Pen(Option<([u8; 3], i32)>),
    Font { key: String, height: i32 },
}

#[derive(Clone)]
struct EmfState {
    window_org: (i32, i32),
    window_ext: (i32, i32),
    viewport_org: (i32, i32),
    viewport_ext: (i32, i32),
    current_pt: (i32, i32),
    fill_rule: FillRule,
    /// `None` is the null brush / pen.
    selected_brush: Option<[u8; 3]>,
    selected_pen: Option<([u8; 3], i32)>,
    font: Option<(String, i32)>,
    text_color: [u8; 3],
    text_align: u32,
    /// Between BEGINPATH and the record that paints or discards the path.
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
            // GDI's defaults: WHITE_BRUSH, BLACK_PEN, black text.
            selected_brush: Some([255, 255, 255]),
            selected_pen: Some(([0, 0, 0], 0)),
            font: None,
            text_color: [0, 0, 0],
            text_align: 0,
            in_path: false,
        }
    }
}

/// Compose the EMF → form-XObject (1×1 unit box) mapping for a logical point.
struct Mapper {
    /// The header's frame in device pixels, mapped onto the unit box.
    frame: (f64, f64, f64, f64),
}

impl Mapper {
    fn map(&self, state: &EmfState, x: i32, y: i32) -> (f32, f32) {
        self.map_f(state, x as f64, y as f64)
    }

    fn map_f(&self, state: &EmfState, x: f64, y: f64) -> (f32, f32) {
        // logical → device
        let dx = (x - state.window_org.0 as f64) * state.viewport_ext.0 as f64
            / state.window_ext.0 as f64
            + state.viewport_org.0 as f64;
        let dy = (y - state.window_org.1 as f64) * state.viewport_ext.1 as f64
            / state.window_ext.1 as f64
            + state.viewport_org.1 as f64;
        // device → form [0,1] with Y flip (PDF is Y-up, EMF is Y-down).
        let (fl, ft, fr, fb) = self.frame;
        let w = (fr - fl).max(1.0);
        let h = (fb - ft).max(1.0);
        (((dx - fl) / w) as f32, (1.0 - (dy - ft) / h) as f32)
    }

    /// One device pixel in form units: GDI's width for a zero-width pen.
    fn pixel(&self) -> f32 {
        1.0 / (self.frame.2 - self.frame.0).max(1.0) as f32
    }
}

/// Stock objects (`SelectObject` with the high bit set), [MS-EMF] §2.1.31.
fn stock_object(index: u32) -> Option<EmfObject> {
    let gray = |v: u8| Some([v, v, v]);
    Some(match index {
        0 => EmfObject::Brush(gray(255)),
        1 => EmfObject::Brush(gray(192)),
        2 => EmfObject::Brush(gray(128)),
        3 => EmfObject::Brush(gray(64)),
        4 => EmfObject::Brush(gray(0)),
        5 => EmfObject::Brush(None),
        6 => EmfObject::Pen(Some(([255, 255, 255], 0))),
        7 => EmfObject::Pen(Some(([0, 0, 0], 0))),
        8 => EmfObject::Pen(None),
        // ponytail: stock fonts keep the selected font; no fixture draws with one.
        _ => return None,
    })
}

struct Translator<'a> {
    mapper: Mapper,
    state: EmfState,
    stack: Vec<EmfState>,
    objects: HashMap<u32, EmfObject>,
    content: Content,
    fonts: &'a HashMap<String, FontEntry>,
    /// PDF font name → font object, for the form's resources.
    used_fonts: BTreeMap<String, Ref>,
    /// `a:lum` brightness/contrast, applied to every colour the metafile
    /// sets (Word recolours the rasterised picture; swiss_materials' stamp
    /// is drawn light grey).
    lum: Option<(f32, f32)>,
}


/// Translate an EMF byte stream into a PDF Form XObject. Returns the
/// allocated Ref, or `None` if the input isn't a parseable EMF. Text draws
/// with the fonts `collect_and_register_fonts` registered for it.
pub(super) fn emf_to_form_xobject(
    emf: &[u8],
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
    fonts: &HashMap<String, FontEntry>,
    lum: Option<(f32, f32)>,
) -> Option<Ref> {
    let header = parse_header(emf)?;
    let bounds_size = header.bounds_size();
    if bounds_size.0 == 0 || bounds_size.1 == 0 {
        return None;
    }
    let mut t = Translator {
        mapper: Mapper {
            frame: header.frame,
        },
        state: EmfState::new(bounds_size),
        stack: Vec::new(),
        objects: HashMap::new(),
        content: Content::new(),
        fonts,
        used_fonts: BTreeMap::new(),
        lum,
    };
    for_each_record(emf, |rec| {
        t.record(rec);
        true
    });

    let form_ref = alloc();
    let bytes = t.content.finish();
    let mut form = pdf.form_xobject(form_ref, &bytes);
    form.bbox(Rect::new(0.0, 0.0, 1.0, 1.0));
    if !t.used_fonts.is_empty() {
        let mut res = form.resources();
        let mut dict = res.fonts();
        for (name, font_ref) in &t.used_fonts {
            dict.pair(Name(name.as_bytes()), *font_ref);
        }
    }
    drop(form);
    Some(form_ref)
}

impl Translator<'_> {
    fn record(&mut self, rec: &EmfRecord) {
        use EmfRecord::*;
        let lum = self.lum;
        let lum_color = |c: [u8; 3]| -> [u8; 3] {
            let Some((bright, contrast)) = lum else {
                return c;
            };
            let mut c = c;
            super::images::apply_lum(&mut c, bright, contrast);
            c
        };
        let state = &mut self.state;
        match rec {
            Header | Eof | Skip => {}
            SetMapMode | SetBkMode => {} // We honour window/viewport explicitly.
            SetPolyFillMode(rule) => state.fill_rule = *rule,
            SetWindowOrgEx(x, y) => state.window_org = (*x, *y),
            SetWindowExtEx(x, y) => state.window_ext = ((*x).max(1), (*y).max(1)),
            SetViewportOrgEx(x, y) => state.viewport_org = (*x, *y),
            SetViewportExtEx(x, y) => state.viewport_ext = ((*x).max(1), (*y).max(1)),
            SetTextColor(c) => state.text_color = lum_color(*c),
            SetTextAlign(a) => state.text_align = *a,
            SaveDc => {
                self.stack.push(state.clone());
                self.content.save_state();
            }
            RestoreDc => {
                if let Some(prev) = self.stack.pop() {
                    *state = prev;
                }
                self.content.restore_state();
            }
            CreateBrushIndirect { handle, color } => {
                let color = color.map(lum_color);
                self.objects.insert(*handle, EmfObject::Brush(color));
            }
            CreatePen {
                handle,
                color,
                width,
            } => {
                self.objects
                    .insert(*handle, EmfObject::Pen(color.map(|c| (lum_color(c), *width))));
            }
            CreateFont {
                handle,
                height,
                bold,
                italic,
                face,
            } => {
                let key = font_key_for(face, *bold, *italic);
                self.objects.insert(
                    *handle,
                    EmfObject::Font {
                        key,
                        height: *height,
                    },
                );
            }
            DeleteObject(handle) => {
                self.objects.remove(handle);
            }
            SelectObject(handle) => {
                let obj = if handle & 0x8000_0000 != 0 {
                    stock_object(handle & 0x7FFF_FFFF)
                } else {
                    self.objects.get(handle).cloned()
                };
                match obj {
                    Some(EmfObject::Brush(c)) => state.selected_brush = c,
                    Some(EmfObject::Pen(p)) => state.selected_pen = p,
                    Some(EmfObject::Font { key, height }) => state.font = Some((key, height)),
                    None => {}
                }
            }
            // PDF path construction has no explicit bracket; segments accumulate until
            // a painting or clipping operator consumes them.
            BeginPath => state.in_path = true,
            EndPath => {}
            MoveToEx(x, y) => {
                if state.in_path {
                    let (u, v) = self.mapper.map(state, *x, *y);
                    self.content.move_to(u, v);
                }
                state.current_pt = (*x, *y);
            }
            LineTo(x, y) => self.line_to(&[(*x, *y)]),
            PolyLineTo16(pts) => {
                let pts: Vec<_> = pts.iter().map(|p| (p.0 as i32, p.1 as i32)).collect();
                self.line_to(&pts);
            }
            PolyBezierTo16(pts) => {
                let open = !state.in_path;
                if open {
                    self.start_at_current();
                }
                let state = &mut self.state;
                for triple in pts.chunks_exact(3) {
                    let m = |p: (i16, i16)| self.mapper.map(state, p.0 as i32, p.1 as i32);
                    let (c1x, c1y) = m(triple[0]);
                    let (c2x, c2y) = m(triple[1]);
                    let (ex, ey) = m(triple[2]);
                    self.content.cubic_to(c1x, c1y, c2x, c2y, ex, ey);
                    state.current_pt = (triple[2].0 as i32, triple[2].1 as i32);
                }
                if open {
                    self.stroke();
                }
            }
            CloseFigure => {
                self.content.close_path();
            }
            FillPath => {
                state.in_path = false;
                match state.selected_brush {
                    Some(c) => {
                        fill_rgb(&mut self.content, c);
                        self.fill();
                    }
                    None => {
                        self.content.end_path();
                    }
                }
            }
            StrokePath => {
                state.in_path = false;
                self.stroke();
            }
            // A PDF path stays open until a painting operator consumes it, so an ignored
            // clip path would merge into the next fill (a logo's clip rectangle came out
            // as a solid black block). Clip with the current fill rule and end the path.
            SelectClipPath => {
                state.in_path = false;
                match state.fill_rule {
                    FillRule::Alternate => self.content.clip_even_odd(),
                    FillRule::Winding => self.content.clip_nonzero(),
                };
                self.content.end_path();
            }
            AbortPath => {
                state.in_path = false;
                self.content.end_path();
            }
            StrokeAndFillPath => {
                state.in_path = false;
                match (state.selected_brush, state.selected_pen) {
                    (Some(c), Some((pc, w))) => {
                        fill_rgb(&mut self.content, c);
                        let rule = state.fill_rule;
                        self.set_pen(pc, w);
                        match rule {
                            FillRule::Alternate => self.content.fill_even_odd_and_stroke(),
                            FillRule::Winding => self.content.fill_nonzero_and_stroke(),
                        };
                    }
                    (Some(c), None) => {
                        fill_rgb(&mut self.content, c);
                        self.fill();
                    }
                    (None, _) => self.stroke(),
                }
            }
            Rectangle(l, t, r, b) => {
                let paint = (state.selected_brush, state.selected_pen);
                self.rect(*l, *t, *r - *l, *b - *t);
                match paint {
                    (Some(c), Some((pc, w))) => {
                        fill_rgb(&mut self.content, c);
                        self.set_pen(pc, w);
                        self.content.fill_nonzero_and_stroke();
                    }
                    (Some(c), None) => {
                        fill_rgb(&mut self.content, c);
                        self.content.fill_nonzero();
                    }
                    (None, _) => self.stroke(),
                }
            }
            PatBlt { x, y, w, h, rop } => {
                let color = match rop {
                    0x00F0_0021 => state.selected_brush,  // PATCOPY
                    0x0000_0042 => Some([0, 0, 0]),       // BLACKNESS
                    0x00FF_0062 => Some([255, 255, 255]), // WHITENESS
                    _ => None,
                };
                if let Some(c) = color {
                    self.rect(*x, *y, *w, *h);
                    fill_rgb(&mut self.content, c);
                    self.content.fill_nonzero();
                }
            }
            ExtTextOut { x, y, text, dx } => self.text(*x, *y, text, dx),
        }
    }

    /// GDI draws LINETO and POLYLINETO outside a path bracket at once with the
    /// current pen (spreadsheet gridlines); inside one they only extend the path.
    fn line_to(&mut self, pts: &[(i32, i32)]) {
        let open = !self.state.in_path;
        if open {
            self.start_at_current();
        }
        for &(x, y) in pts {
            let (u, v) = self.mapper.map(&self.state, x, y);
            self.content.line_to(u, v);
            self.state.current_pt = (x, y);
        }
        if open {
            self.stroke();
        }
    }

    fn start_at_current(&mut self) {
        let (cx, cy) = self.state.current_pt;
        let (u, v) = self.mapper.map(&self.state, cx, cy);
        self.content.move_to(u, v);
    }

    fn rect(&mut self, x: i32, y: i32, w: i32, h: i32) {
        let (u0, v0) = self.mapper.map(&self.state, x, y);
        let (u1, v1) = self.mapper.map(&self.state, x + w, y + h);
        self.content
            .rect(u0.min(u1), v0.min(v1), (u1 - u0).abs(), (v1 - v0).abs());
    }

    fn fill(&mut self) {
        match self.state.fill_rule {
            FillRule::Alternate => self.content.fill_even_odd(),
            FillRule::Winding => self.content.fill_nonzero(),
        };
    }

    fn set_pen(&mut self, color: [u8; 3], width: i32) {
        stroke_rgb(&mut self.content, color);
        let w = if width > 0 {
            stroke_width_in_form(&self.state, &self.mapper, width)
        } else {
            self.mapper.pixel()
        };
        self.content.set_line_width(w);
    }

    fn stroke(&mut self) {
        match self.state.selected_pen {
            Some((c, w)) => {
                self.set_pen(c, w);
                self.content.stroke();
            }
            None => {
                self.content.end_path();
            }
        }
    }

    /// EMR_EXTTEXTOUTW: the reference point sits per SETTEXTALIGN (top, baseline or
    /// bottom; left, centre or right) and each glyph advances by its Dx entry.
    // ponytail: no clipping (INTERSECTCLIPRECT/EXTSELECTCLIPRGN), opaque
    // backgrounds, escapement or TA_UPDATECP; add when a fixture needs them.
    fn text(&mut self, x: i32, y: i32, text: &str, dx: &[i32]) {
        let Some((key, height)) = &self.state.font else {
            return;
        };
        let Some(entry) = self.fonts.get(key) else {
            return;
        };
        let line_ratio = entry
            .plain_line_h_ratio
            .or(entry.line_h_ratio)
            .unwrap_or(1.15);
        let ascent_ratio = entry
            .plain_ascender_ratio
            .or(entry.ascender_ratio)
            .unwrap_or(0.9);
        let em = match *height {
            h if h < 0 => -h as f64,
            h if h > 0 => h as f64 / line_ratio as f64,
            _ => return,
        };
        let chars: Vec<char> = text.chars().collect();
        let advances: Vec<f64> = if dx.len() >= chars.len() {
            dx.iter().map(|&d| d as f64).collect()
        } else {
            chars
                .iter()
                .map(|&c| entry.char_width_1000(c) as f64 * em / 1000.0)
                .collect()
        };
        let width: f64 = advances.iter().take(chars.len()).sum();
        let align = self.state.text_align;
        let x0 = match align & 6 {
            2 => x as f64 - width,       // TA_RIGHT
            6 => x as f64 - width / 2.0, // TA_CENTER
            _ => x as f64,
        };
        let baseline = match align & 24 {
            24 => y as f64,                                          // TA_BASELINE
            8 => y as f64 - (line_ratio - ascent_ratio) as f64 * em, // TA_BOTTOM
            _ => y as f64 + ascent_ratio as f64 * em,                // TA_TOP
        };
        let (u, v) = self.mapper.map_f(&self.state, x0, baseline);
        let (u1, _) = self.mapper.map_f(&self.state, x0 + 1.0, baseline);
        let (_, v1) = self.mapper.map_f(&self.state, x0, baseline + 1.0);
        let (sx, sy) = (u1 - u, v - v1);

        self.used_fonts
            .insert(entry.pdf_name.clone(), entry.font_ref);
        let c = &mut self.content;
        c.begin_text();
        c.set_font(Name(entry.pdf_name.as_bytes()), 1.0);
        fill_rgb(c, self.state.text_color);
        c.set_text_matrix([sx * em as f32, 0.0, 0.0, sy * em as f32, u, v]);
        let mut tj = c.show_positioned();
        let mut items = tj.items();
        let mut buf = [0u8; 4];
        for (i, &ch) in chars.iter().enumerate() {
            items.show(Str(&entry.encode(ch.encode_utf8(&mut buf))));
            if i + 1 < chars.len() {
                // TJ subtracts thousandths of the em from the glyph's own advance.
                let adjust = entry.char_width_1000(ch) as f64 - advances[i] * 1000.0 / em;
                if adjust.abs() > 0.01 {
                    items.adjust(adjust as f32);
                }
            }
        }
        drop(items);
        drop(tj);
        c.end_text();
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
