use std::collections::HashMap;

use pdf_writer::types::TextRenderingMode;
use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::geometry::{self, ResolvedCommand};
use crate::model::{TextFill, TextGlow, TextOutline, TextShadow, Textbox};

/// Apply text outline rendering state to the PDF content stream.
/// Returns `true` if an outline mode was set (caller must call `reset_text_outline` afterward).
pub(super) fn apply_text_outline(
    content: &mut Content,
    outline: &TextOutline,
    text_fill: Option<&TextFill>,
) -> bool {
    content.set_line_width(outline.width_pt);
    stroke_rgb(content, outline.color);

    let mode = match text_fill {
        Some(TextFill::NoFill) => TextRenderingMode::Stroke,
        _ => TextRenderingMode::FillStroke,
    };
    content.set_text_rendering_mode(mode);
    true
}

pub(super) fn reset_text_outline(content: &mut Content) {
    content.set_text_rendering_mode(TextRenderingMode::Fill);
}

pub(super) fn find_text_shadow(tb: &Textbox) -> Option<&TextShadow> {
    tb.paragraphs
        .iter()
        .flat_map(|p| &p.runs)
        .find_map(|r| r.text_shadow.as_ref())
}

pub(super) fn find_text_glow(tb: &Textbox) -> Option<&TextGlow> {
    tb.paragraphs
        .iter()
        .flat_map(|p| &p.runs)
        .find_map(|r| r.text_glow.as_ref())
}

// ---------------------------------------------------------------------------
// Glyph outline extraction
// ---------------------------------------------------------------------------

pub(super) struct GlyphPath {
    pub commands: Vec<GlyphCommand>,
    pub advance_width: f32,
}

pub(super) enum GlyphCommand {
    MoveTo(f32, f32),
    LineTo(f32, f32),
    QuadTo(f32, f32, f32, f32),
    CubicTo(f32, f32, f32, f32, f32, f32),
    Close,
}

struct GlyphOutlineCollector {
    commands: Vec<GlyphCommand>,
}

impl ttf_parser::OutlineBuilder for GlyphOutlineCollector {
    fn move_to(&mut self, x: f32, y: f32) {
        self.commands.push(GlyphCommand::MoveTo(x, y));
    }
    fn line_to(&mut self, x: f32, y: f32) {
        self.commands.push(GlyphCommand::LineTo(x, y));
    }
    fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
        self.commands.push(GlyphCommand::QuadTo(x1, y1, x, y));
    }
    fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
        self.commands.push(GlyphCommand::CubicTo(x1, y1, x2, y2, x, y));
    }
    fn close(&mut self) {
        self.commands.push(GlyphCommand::Close);
    }
}

pub(super) fn extract_glyph_path(face: &ttf_parser::Face, ch: char) -> Option<GlyphPath> {
    let gid = face.glyph_index(ch)?;
    let advance = face.glyph_hor_advance(gid)? as f32;
    let mut collector = GlyphOutlineCollector {
        commands: Vec::new(),
    };
    face.outline_glyph(gid, &mut collector)?;
    Some(GlyphPath {
        commands: collector.commands,
        advance_width: advance,
    })
}

/// Search seen_fonts for a matching entry with the right bold/italic variant.
pub(super) fn load_font_data(
    font_name: &str,
    bold: bool,
    italic: bool,
    seen_fonts: &HashMap<String, FontEntry>,
) -> Option<(Vec<u8>, u32)> {
    let suffix = match (bold, italic) {
        (true, true) => "/BI",
        (true, false) => "/B",
        (false, true) => "/I",
        (false, false) => "",
    };
    let preferred_key = format!("{}{}", font_name, suffix);
    let entry = seen_fonts
        .get(&preferred_key)
        .filter(|e| e.font_path.is_some())
        .or_else(|| {
            seen_fonts
                .iter()
                .find(|(k, e)| k.starts_with(font_name) && e.font_path.is_some())
                .map(|(_, e)| e)
        })?;
    let path = entry.font_path.as_ref()?;
    let data = std::fs::read(path).ok()?;
    Some((data, entry.face_index))
}

// ---------------------------------------------------------------------------
// Shared helpers for glyph rendering
// ---------------------------------------------------------------------------

/// Text and formatting collected from a textbox's paragraphs.
struct WordArtTextInfo {
    total_text: String,
    font_name: String,
    font_size: f32,
    text_color: Option<[u8; 3]>,
    outline: Option<TextOutline>,
    fill: Option<TextFill>,
    bold: bool,
    italic: bool,
}

/// Collect all text and the first run's formatting from a textbox.
fn collect_text_info(tb: &Textbox) -> Option<WordArtTextInfo> {
    let mut info = WordArtTextInfo {
        total_text: String::new(),
        font_name: String::new(),
        font_size: 12.0,
        text_color: None,
        outline: None,
        fill: None,
        bold: false,
        italic: false,
    };

    for para in &tb.paragraphs {
        for run in &para.runs {
            if !run.text.is_empty() {
                info.total_text.push_str(&run.text);
                if info.font_name.is_empty() {
                    info.font_name = run.font_name.clone();
                    info.font_size = run.font_size;
                    info.text_color = run.color;
                    info.outline = run.text_outline.clone();
                    info.fill = run.text_fill.clone();
                    info.bold = run.bold;
                    info.italic = run.italic;
                }
            }
        }
    }

    if info.total_text.is_empty() || info.font_name.is_empty() {
        return None;
    }
    Some(info)
}

use super::color::{fill_color_or_black, stroke_rgb};

/// Resolve the effective fill color from text_fill + run color.
fn resolve_fill_color(
    text_fill: &Option<TextFill>,
    text_color: Option<[u8; 3]>,
) -> Option<[u8; 3]> {
    match text_fill {
        Some(TextFill::Solid(c)) => Some(*c),
        Some(TextFill::NoFill) => None,
        _ => text_color,
    }
}

/// Emit glyph path commands into a PDF content stream, applying a coordinate
/// transform to each point. Handles quad-to-cubic conversion and endpoint tracking.
fn emit_glyph_commands(
    glyph: &GlyphPath,
    content: &mut Content,
    transform: impl Fn(f64, f64) -> (f32, f32),
) {
    let mut prev_x = 0.0_f64;
    let mut prev_y = 0.0_f64;

    for cmd in &glyph.commands {
        match cmd {
            GlyphCommand::MoveTo(gx, gy) => {
                let (px, py) = transform(*gx as f64, *gy as f64);
                content.move_to(px, py);
                prev_x = *gx as f64;
                prev_y = *gy as f64;
            }
            GlyphCommand::LineTo(gx, gy) => {
                let (px, py) = transform(*gx as f64, *gy as f64);
                content.line_to(px, py);
                prev_x = *gx as f64;
                prev_y = *gy as f64;
            }
            GlyphCommand::QuadTo(qcx, qcy, qx, qy) => {
                let qcx = *qcx as f64;
                let qcy = *qcy as f64;
                let p1x = *qx as f64;
                let p1y = *qy as f64;

                let cp1x = prev_x + (2.0 / 3.0) * (qcx - prev_x);
                let cp1y = prev_y + (2.0 / 3.0) * (qcy - prev_y);
                let cp2x = p1x + (2.0 / 3.0) * (qcx - p1x);
                let cp2y = p1y + (2.0 / 3.0) * (qcy - p1y);

                let (wx1, wy1) = transform(cp1x, cp1y);
                let (wx2, wy2) = transform(cp2x, cp2y);
                let (wx3, wy3) = transform(p1x, p1y);
                content.cubic_to(wx1, wy1, wx2, wy2, wx3, wy3);

                prev_x = p1x;
                prev_y = p1y;
            }
            GlyphCommand::CubicTo(x1, y1, x2, y2, x3, y3) => {
                let (wx1, wy1) = transform(*x1 as f64, *y1 as f64);
                let (wx2, wy2) = transform(*x2 as f64, *y2 as f64);
                let (wx3, wy3) = transform(*x3 as f64, *y3 as f64);
                content.cubic_to(wx1, wy1, wx2, wy2, wx3, wy3);
                prev_x = *x3 as f64;
                prev_y = *y3 as f64;
            }
            GlyphCommand::Close => {
                content.close_path();
            }
        }
    }
}

/// Emit glyph commands using a transform that also receives the cursor_x position.
fn emit_glyph_commands_with_cursor(
    glyph: &GlyphPath,
    content: &mut Content,
    cursor_x: f64,
    transform: &impl Fn(f64, f64, f64) -> (f32, f32),
) {
    emit_glyph_commands(glyph, content, |gx, gy| transform(cursor_x, gx, gy));
}

/// Fill the current path, then optionally re-emit glyphs and stroke.
fn fill_and_stroke_glyphs(
    content: &mut Content,
    outline: &Option<TextOutline>,
    fill: &Option<TextFill>,
    emit_stroke_paths: impl FnOnce(&mut Content),
) {
    let no_fill = matches!(fill, Some(TextFill::NoFill));

    if !no_fill {
        content.fill_nonzero();
    } else {
        content.end_path();
    }

    if let Some(outline) = outline {
        content.set_line_width(outline.width_pt);
        stroke_rgb(content, outline.color);
        emit_stroke_paths(content);
        content.stroke();
    }
}

/// Evaluate a cubic bezier at parameter `t` given start point `(cx, cy)` and
/// control/end points `(x1, y1, x2, y2, x, y)`.
fn eval_cubic(cx: f64, cy: f64, x1: f64, y1: f64, x2: f64, y2: f64, x: f64, y: f64, t: f64) -> (f64, f64) {
    let mt = 1.0 - t;
    let mt2 = mt * mt;
    let mt3 = mt2 * mt;
    let t2 = t * t;
    let t3 = t2 * t;
    let px = mt3 * cx + 3.0 * mt2 * t * x1 + 3.0 * mt * t2 * x2 + t3 * x;
    let py = mt3 * cy + 3.0 * mt2 * t * y1 + 3.0 * mt * t2 * y2 + t3 * y;
    (px, py)
}

// ---------------------------------------------------------------------------
// Text warp algorithm
// ---------------------------------------------------------------------------

/// A boundary curve sampled into (x, y) points for fast interpolation.
/// Points are sorted by x-coordinate.
struct SampledBoundary {
    points: Vec<(f64, f64)>,
}

impl SampledBoundary {
    fn from_commands(commands: &[ResolvedCommand]) -> Self {
        let mut points = Vec::with_capacity(256);
        let mut cx = 0.0;
        let mut cy = 0.0;

        for cmd in commands {
            match cmd {
                ResolvedCommand::MoveTo(x, y) => {
                    cx = *x;
                    cy = *y;
                    points.push((cx, cy));
                }
                ResolvedCommand::LineTo(x, y) => {
                    cx = *x;
                    cy = *y;
                    points.push((cx, cy));
                }
                ResolvedCommand::CubicTo {
                    x1, y1, x2, y2, x, y,
                } => {
                    let steps = 20;
                    for i in 1..=steps {
                        let t = i as f64 / steps as f64;
                        let (px, py) = eval_cubic(cx, cy, *x1, *y1, *x2, *y2, *x, *y, t);
                        points.push((px, py));
                    }
                    cx = *x;
                    cy = *y;
                }
                ResolvedCommand::Close => {}
            }
        }

        points.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
        points.dedup_by(|a, b| (a.0 - b.0).abs() < 0.001);

        SampledBoundary { points }
    }

    fn y_at(&self, x: f64) -> f64 {
        if self.points.is_empty() {
            return 0.0;
        }
        if self.points.len() == 1 {
            return self.points[0].1;
        }

        let first = self.points.first().unwrap();
        let last = self.points.last().unwrap();
        if x <= first.0 {
            return first.1;
        }
        if x >= last.0 {
            return last.1;
        }

        let idx = self
            .points
            .partition_point(|p| p.0 < x)
            .min(self.points.len() - 1)
            .max(1);
        let (x0, y0) = self.points[idx - 1];
        let (x1, y1) = self.points[idx];
        let dx = x1 - x0;
        if dx.abs() < 1e-10 {
            return y0;
        }
        let t = (x - x0) / dx;
        y0 + t * (y1 - y0)
    }

    fn min_y(&self) -> f64 {
        self.points
            .iter()
            .map(|p| p.1)
            .fold(f64::INFINITY, f64::min)
    }

    fn max_y(&self) -> f64 {
        self.points
            .iter()
            .map(|p| p.1)
            .fold(f64::NEG_INFINITY, f64::max)
    }

    fn x_range(&self) -> f64 {
        if self.points.len() < 2 {
            return 0.0;
        }
        self.points.last().unwrap().0 - self.points.first().unwrap().0
    }
}

pub(super) fn evaluate_warp_boundaries(
    preset: &str,
    adjustments: &[(String, i64)],
    w: f64,
    h: f64,
) -> Option<(SampledBoundary, SampledBoundary)> {
    let def = geometry::text_warp_definitions::lookup_text_warp(preset)?;
    let shape = geometry::evaluate_def(def, w, h, adjustments);

    if shape.paths.len() < 2 {
        return None;
    }

    let top = SampledBoundary::from_commands(&shape.paths[0].commands);
    let bottom = SampledBoundary::from_commands(&shape.paths[1].commands);
    Some((top, bottom))
}

/// Warp a single point through the envelope boundaries.
/// `gx`, `gy` are in flat text coordinate space (0..text_w, 0..text_h).
fn warp_point(
    gx: f64,
    gy: f64,
    text_w: f64,
    text_h: f64,
    boundary_w: f64,
    top: &SampledBoundary,
    bottom: &SampledBoundary,
) -> (f64, f64) {
    let v = if text_h > 0.0 { gy / text_h } else { 0.5 };

    let boundary_x = if text_w > 0.0 {
        gx / text_w * boundary_w
    } else {
        boundary_w * 0.5
    };

    let top_y = top.y_at(boundary_x);
    let bot_y = bottom.y_at(boundary_x);
    let warped_y = bot_y + v * (top_y - bot_y);

    (boundary_x, warped_y)
}

/// Render a warped textbox: extract glyph outlines, warp them through the envelope,
/// and emit as filled PDF paths. Returns true if rendering succeeded, false to fall back to flat.
pub(super) fn render_warped_textbox(
    tb: &Textbox,
    content: &mut Content,
    seen_fonts: &HashMap<String, FontEntry>,
    tb_x: f32,
    tb_y_top: f32,
    content_w: f32,
) -> bool {
    let warp = match &tb.text_warp {
        Some(w) => w,
        None => return false,
    };

    let Some(info) = collect_text_info(tb) else {
        return false;
    };

    let Some((font_data, face_index)) =
        load_font_data(&info.font_name, info.bold, info.italic, seen_fonts)
    else {
        return false;
    };
    let Some(face) = ttf_parser::Face::parse(&font_data, face_index).ok() else {
        return false;
    };

    let units_per_em = face.units_per_em() as f64;
    let scale = info.font_size as f64 / units_per_em;

    let descender = face.descender() as f64 / units_per_em * info.font_size as f64;
    let ascender = face.ascender() as f64 / units_per_em * info.font_size as f64;
    let glyph_extent = ascender - descender;

    let char_advances = compute_char_advances(&face, &info.total_text, scale);
    let total_advance: f64 = char_advances.iter().map(|(_, a)| a).sum();

    let text_w = total_advance.max(1.0);
    // Word normalizes WordArt text height using font_size * (ascender / glyph_extent),
    // which accounts for the ascender:descender ratio so that capital letters
    // fill the shape height more completely (matching Word's rasterized output).
    let text_h = info.font_size as f64 * ascender / glyph_extent;

    let boundary_w = text_w;
    let boundary_h = tb.height_pt as f64;
    let Some((top_boundary, bottom_boundary)) =
        evaluate_warp_boundaries(&warp.preset, &warp.adjustments, boundary_w, boundary_h)
    else {
        return false;
    };

    let fill_color = resolve_fill_color(&info.fill, info.text_color);

    content.save_state();
    fill_color_or_black(content, fill_color);

    // Geometry engine already converts boundaries to PDF y-up coords (shape_h - scaled),
    // so boundary y=0 is bottom of envelope, y=boundary_h is top.
    let envelope_top = top_boundary.max_y();
    let x_offset = (content_w as f64 - boundary_w) / 2.0;

    let transform = |cursor_x: f64, gx_font: f64, gy_font: f64| -> (f32, f32) {
        let gx = cursor_x + gx_font * scale;
        let gy = gy_font * scale - descender;

        let (wx, wy) =
            warp_point(gx, gy, text_w, text_h, boundary_w, &top_boundary, &bottom_boundary);

        let pdf_x = (tb_x as f64 + x_offset + wx) as f32;
        let pdf_y = (tb_y_top as f64 - envelope_top + wy) as f32;
        (pdf_x, pdf_y)
    };

    let emit = |content: &mut Content| {
        emit_all_glyphs(&face, &char_advances, content, |cursor_x, gx, gy| {
            transform(cursor_x, gx, gy)
        });
    };

    emit(content);

    fill_and_stroke_glyphs(content, &info.outline, &info.fill, |content| {
        emit(content);
    });

    content.restore_state();
    true
}

// ---------------------------------------------------------------------------
// Text-on-a-path rendering (for single-path presets: arch, circle)
// ---------------------------------------------------------------------------

/// A path sampled at uniform parameter steps with cumulative arc lengths.
struct ArcLengthPath {
    samples: Vec<(f64, f64, f64)>,
}

impl ArcLengthPath {
    fn from_commands(commands: &[ResolvedCommand]) -> Self {
        let mut samples = Vec::with_capacity(256);
        let mut cx = 0.0;
        let mut cy = 0.0;
        let mut cum_len = 0.0;

        for cmd in commands {
            match cmd {
                ResolvedCommand::MoveTo(x, y) => {
                    cx = *x;
                    cy = *y;
                    if samples.is_empty() {
                        samples.push((cx, cy, 0.0));
                    }
                }
                ResolvedCommand::LineTo(x, y) => {
                    let dx = *x - cx;
                    let dy = *y - cy;
                    cum_len += (dx * dx + dy * dy).sqrt();
                    cx = *x;
                    cy = *y;
                    samples.push((cx, cy, cum_len));
                }
                ResolvedCommand::CubicTo {
                    x1, y1, x2, y2, x, y,
                } => {
                    let steps = 32;
                    for i in 1..=steps {
                        let t = i as f64 / steps as f64;
                        let (px, py) = eval_cubic(cx, cy, *x1, *y1, *x2, *y2, *x, *y, t);
                        let default = (cx, cy, 0.0);
                        let prev = samples.last().unwrap_or(&default);
                        let dx = px - prev.0;
                        let dy = py - prev.1;
                        cum_len += (dx * dx + dy * dy).sqrt();
                        samples.push((px, py, cum_len));
                    }
                    cx = *x;
                    cy = *y;
                }
                ResolvedCommand::Close => {}
            }
        }

        ArcLengthPath { samples }
    }

    fn total_length(&self) -> f64 {
        self.samples.last().map(|s| s.2).unwrap_or(0.0)
    }

    fn position_at(&self, s: f64) -> (f64, f64) {
        if self.samples.len() < 2 {
            return self
                .samples
                .first()
                .map(|p| (p.0, p.1))
                .unwrap_or((0.0, 0.0));
        }
        let s = s.clamp(0.0, self.total_length());
        let idx = self
            .samples
            .partition_point(|p| p.2 < s)
            .min(self.samples.len() - 1)
            .max(1);
        let (x0, y0, s0) = self.samples[idx - 1];
        let (x1, y1, s1) = self.samples[idx];
        let ds = s1 - s0;
        if ds.abs() < 1e-10 {
            return (x0, y0);
        }
        let t = (s - s0) / ds;
        (x0 + t * (x1 - x0), y0 + t * (y1 - y0))
    }

    fn tangent_at(&self, s: f64) -> f64 {
        if self.samples.len() < 2 {
            return 0.0;
        }
        let s = s.clamp(0.0, self.total_length());
        let idx = self
            .samples
            .partition_point(|p| p.2 < s)
            .min(self.samples.len() - 1)
            .max(1);
        let (x0, y0, _) = self.samples[idx - 1];
        let (x1, y1, _) = self.samples[idx];
        (y1 - y0).atan2(x1 - x0)
    }
}

/// Render text along a single path (for arch/circle presets).
/// Returns true if rendering succeeded.
pub(super) fn render_text_on_path(
    tb: &Textbox,
    content: &mut Content,
    seen_fonts: &HashMap<String, FontEntry>,
    tb_x: f32,
    tb_y_top: f32,
    content_w: f32,
) -> bool {
    let warp = match &tb.text_warp {
        Some(w) => w,
        None => return false,
    };

    let Some(info) = collect_text_info(tb) else {
        return false;
    };

    let Some((font_data, face_index)) =
        load_font_data(&info.font_name, info.bold, info.italic, seen_fonts)
    else {
        return false;
    };
    let Some(face) = ttf_parser::Face::parse(&font_data, face_index).ok() else {
        return false;
    };

    let units_per_em = face.units_per_em() as f64;
    let scale = info.font_size as f64 / units_per_em;

    let char_advances = compute_char_advances(&face, &info.total_text, scale);
    let total_advance: f64 = char_advances.iter().map(|(_, a)| a).sum();

    let Some(def) = geometry::text_warp_definitions::lookup_text_warp(&warp.preset) else {
        return false;
    };
    let boundary_w = total_advance;
    let boundary_h = tb.height_pt as f64;
    let shape = geometry::evaluate_def(def, boundary_w, boundary_h, &warp.adjustments);
    if shape.paths.is_empty() {
        return false;
    }
    let arc_path = ArcLengthPath::from_commands(&shape.paths[0].commands);
    let total_arc = arc_path.total_length();
    if total_arc < 1.0 {
        return false;
    }

    let fill_color = resolve_fill_color(&info.fill, info.text_color);

    content.save_state();
    fill_color_or_black(content, fill_color);

    let start_s = (total_arc - total_advance) / 2.0;
    let x_offset = (content_w as f64 - boundary_w) / 2.0;
    let path_anchor = info.font_size as f64 / 2.0;

    let emit_on_path = |content: &mut Content| {
        let mut cursor_s = start_s;
        for &(ch, advance) in &char_advances {
            let Some(glyph) = extract_glyph_path(&face, ch) else {
                cursor_s += advance;
                continue;
            };

            let char_center_s = cursor_s + advance / 2.0;
            let (cx, cy) = arc_path.position_at(char_center_s);
            let angle = arc_path.tangent_at(char_center_s);
            let cos_a = angle.cos();
            let sin_a = angle.sin();

            let transform_pt = |gx_font: f64, gy_font: f64| -> (f32, f32) {
                let lx = gx_font * scale - advance / 2.0;
                let ly = gy_font * scale - path_anchor;
                let rx = lx * cos_a - ly * sin_a;
                let ry = lx * sin_a + ly * cos_a;
                let pdf_x = (tb_x as f64 + x_offset + cx + rx) as f32;
                let pdf_y = (tb_y_top as f64 - boundary_h + cy + ry) as f32;
                (pdf_x, pdf_y)
            };

            emit_glyph_commands(&glyph, content, &transform_pt);
            cursor_s += advance;
        }
    };

    emit_on_path(content);

    fill_and_stroke_glyphs(content, &info.outline, &info.fill, emit_on_path);

    content.restore_state();
    true
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

fn compute_char_advances(face: &ttf_parser::Face, text: &str, scale: f64) -> Vec<(char, f64)> {
    text.chars()
        .map(|ch| {
            let adv = face
                .glyph_index(ch)
                .and_then(|gid| face.glyph_hor_advance(gid))
                .unwrap_or(0) as f64
                * scale;
            (ch, adv)
        })
        .collect()
}

/// Emit all glyph outlines for each character, using a transform that
/// receives `(cursor_x, glyph_x, glyph_y)` and returns PDF coordinates.
fn emit_all_glyphs(
    face: &ttf_parser::Face,
    char_advances: &[(char, f64)],
    content: &mut Content,
    transform: impl Fn(f64, f64, f64) -> (f32, f32),
) {
    let mut cursor_x = 0.0_f64;
    for &(ch, advance) in char_advances {
        if let Some(glyph) = extract_glyph_path(face, ch) {
            emit_glyph_commands_with_cursor(&glyph, content, cursor_x, &transform);
        }
        cursor_x += advance;
    }
}
