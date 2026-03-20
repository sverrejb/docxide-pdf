use std::collections::HashMap;

use pdf_writer::types::TextRenderingMode;
use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::model::{TextFill, TextGlow, TextOutline, TextShadow, Textbox};

/// Apply text outline rendering state to the PDF content stream.
/// Returns `true` if an outline mode was set (caller must call `reset_text_outline` afterward).
pub(super) fn apply_text_outline(
    content: &mut Content,
    outline: &TextOutline,
    text_fill: Option<&TextFill>,
) -> bool {
    content.set_line_width(outline.width_pt);
    let [r, g, b] = outline.color;
    content.set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);

    let mode = match text_fill {
        Some(TextFill::NoFill) => TextRenderingMode::Stroke,
        _ => TextRenderingMode::FillStroke,
    };
    content.set_text_rendering_mode(mode);
    true
}

/// Reset text rendering mode back to normal fill after outline rendering.
pub(super) fn reset_text_outline(content: &mut Content) {
    content.set_text_rendering_mode(TextRenderingMode::Fill);
}

/// Find the first text shadow defined on any run in the textbox.
pub(super) fn find_text_shadow(tb: &Textbox) -> Option<&TextShadow> {
    tb.paragraphs
        .iter()
        .flat_map(|p| &p.runs)
        .find_map(|r| r.text_shadow.as_ref())
}

/// Find the first text glow defined on any run in the textbox.
pub(super) fn find_text_glow(tb: &Textbox) -> Option<&TextGlow> {
    tb.paragraphs
        .iter()
        .flat_map(|p| &p.runs)
        .find_map(|r| r.text_glow.as_ref())
}

// ---------------------------------------------------------------------------
// Glyph outline extraction (Phase 3: text warping)
// ---------------------------------------------------------------------------

/// A single glyph's outline path in font units.
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

/// Extract glyph outline from a font face. Returns path commands in font units.
pub(super) fn extract_glyph_path(
    face: &ttf_parser::Face,
    ch: char,
) -> Option<GlyphPath> {
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

/// Load font data for glyph outline extraction.
/// Searches seen_fonts for an entry whose key starts with the given font name and has a stored path.
pub(super) fn load_font_data(
    font_name: &str,
    seen_fonts: &HashMap<String, FontEntry>,
) -> Option<(Vec<u8>, u32)> {
    // Font keys are "FontName", "FontName/B", "FontName/I", "FontName/BI"
    let entry = seen_fonts
        .iter()
        .find(|(k, e)| k.starts_with(font_name) && e.font_path.is_some())
        .map(|(_, e)| e)?;
    let path = entry.font_path.as_ref()?;
    let data = std::fs::read(path).ok()?;
    Some((data, entry.face_index))
}

// ---------------------------------------------------------------------------
// Text warp algorithm (Phase 3)
// ---------------------------------------------------------------------------

use crate::geometry::{self, ResolvedCommand};

/// A boundary curve sampled into (x, y) points for fast interpolation.
/// Points are sorted by x-coordinate.
struct SampledBoundary {
    points: Vec<(f64, f64)>,
}

impl SampledBoundary {
    /// Build from resolved geometry path commands by flattening curves into line segments.
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
                    x1,
                    y1,
                    x2,
                    y2,
                    x,
                    y,
                } => {
                    // Flatten cubic bezier into ~20 segments
                    let steps = 20;
                    for i in 1..=steps {
                        let t = i as f64 / steps as f64;
                        let t2 = t * t;
                        let t3 = t2 * t;
                        let mt = 1.0 - t;
                        let mt2 = mt * mt;
                        let mt3 = mt2 * mt;
                        let px = mt3 * cx + 3.0 * mt2 * t * x1 + 3.0 * mt * t2 * x2 + t3 * x;
                        let py = mt3 * cy + 3.0 * mt2 * t * y1 + 3.0 * mt * t2 * y2 + t3 * y;
                        points.push((px, py));
                    }
                    cx = *x;
                    cy = *y;
                }
                ResolvedCommand::Close => {}
            }
        }

        // Sort by x for binary search
        points.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
        // Deduplicate x values (keep last occurrence)
        points.dedup_by(|a, b| (a.0 - b.0).abs() < 0.001);

        SampledBoundary { points }
    }

    /// Interpolate y at a given x. Clamps to boundary endpoints.
    fn y_at(&self, x: f64) -> f64 {
        if self.points.is_empty() {
            return 0.0;
        }
        if self.points.len() == 1 {
            return self.points[0].1;
        }

        // Clamp to range
        let first = self.points.first().unwrap();
        let last = self.points.last().unwrap();
        if x <= first.0 {
            return first.1;
        }
        if x >= last.0 {
            return last.1;
        }

        // Binary search for the segment containing x
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

    /// Minimum y value across all sampled points.
    fn min_y(&self) -> f64 {
        self.points
            .iter()
            .map(|p| p.1)
            .fold(f64::INFINITY, f64::min)
    }

    /// Maximum y value across all sampled points.
    fn max_y(&self) -> f64 {
        self.points
            .iter()
            .map(|p| p.1)
            .fold(f64::NEG_INFINITY, f64::max)
    }

    /// X-extent of the boundary.
    fn x_range(&self) -> f64 {
        if self.points.len() < 2 {
            return 0.0;
        }
        self.points.last().unwrap().0 - self.points.first().unwrap().0
    }
}

/// Evaluate the warp preset and return sampled top and bottom boundaries.
/// Returns None if the preset is not found or has fewer than 2 paths.
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
/// `gx`, `gy` are in the flat text coordinate system (0..text_w, 0..text_h).
/// `boundary_w` is the x-extent of the boundary curves (typically content_w).
/// Returns warped (x, y) in boundary coordinates.
fn warp_point(
    gx: f64,
    gy: f64,
    text_w: f64,
    text_h: f64,
    boundary_w: f64,
    top: &SampledBoundary,
    bottom: &SampledBoundary,
) -> (f64, f64) {
    // v = vertical position (0 = bottom, 1 = top of text)
    let v = if text_h > 0.0 { gy / text_h } else { 0.5 };

    // Map glyph x to boundary x-coordinate space
    let boundary_x = if text_w > 0.0 {
        gx / text_w * boundary_w
    } else {
        boundary_w * 0.5
    };

    let top_y = top.y_at(boundary_x);
    let bot_y = bottom.y_at(boundary_x);

    // Interpolate between bottom and top boundary
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

    // Collect all text and first run's formatting
    let mut total_text = String::new();
    let mut font_name = String::new();
    let mut font_size = 12.0_f32;
    let mut text_color: Option<[u8; 3]> = None;
    let mut first_outline: Option<TextOutline> = None;
    let mut first_fill: Option<TextFill> = None;
    for para in &tb.paragraphs {
        for run in &para.runs {
            if !run.text.is_empty() {
                total_text.push_str(&run.text);
                if font_name.is_empty() {
                    font_name = run.font_name.clone();
                    font_size = run.font_size;
                    text_color = run.color;
                    first_outline = run.text_outline.clone();
                    first_fill = run.text_fill.clone();
                }
            }
        }
    }

    if total_text.is_empty() || font_name.is_empty() {
        return false;
    }

    // Load font for glyph outlines
    let Some((font_data, face_index)) = load_font_data(&font_name, seen_fonts) else {
        return false;
    };
    let Some(face) = ttf_parser::Face::parse(&font_data, face_index).ok() else {
        return false;
    };

    let units_per_em = face.units_per_em() as f64;
    let scale = font_size as f64 / units_per_em;

    // Font metrics for coordinate transform (Fix 1)
    let descender = face.descender() as f64 / units_per_em * font_size as f64; // negative
    let ascender = face.ascender() as f64 / units_per_em * font_size as f64;
    let glyph_extent = ascender - descender; // total glyph height in points

    // Compute character advances to get total text width
    let mut char_advances: Vec<(char, f64)> = Vec::new();
    let mut total_advance = 0.0_f64;
    for ch in total_text.chars() {
        let adv = face
            .glyph_index(ch)
            .and_then(|gid| face.glyph_hor_advance(gid))
            .unwrap_or(0) as f64
            * scale;
        char_advances.push((ch, adv));
        total_advance += adv;
    }

    let text_w = total_advance.max(1.0);
    let text_h = glyph_extent; // glyph vertical extent for v normalization

    // Evaluate warp boundaries using textbox dimensions (not glyph extent)
    // WordArt stretches text to fill the textbox envelope
    let boundary_w = content_w as f64;
    let boundary_h = tb.height_pt as f64;
    let Some((top_boundary, bottom_boundary)) =
        evaluate_warp_boundaries(&warp.preset, &warp.adjustments, boundary_w, boundary_h)
    else {
        return false;
    };

    // Determine fill color from text_fill, falling back to run color
    let fill_color = match &first_fill {
        Some(TextFill::Solid([r, g, b])) => Some((*r, *g, *b)),
        Some(TextFill::NoFill) => None,
        _ => text_color.map(|[r, g, b]| (r, g, b)),
    };

    // Set up rendering
    content.save_state();
    if let Some((r, g, b)) = fill_color {
        content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
    } else {
        content.set_fill_gray(0.0);
    }

    // Geometry engine already converts boundaries to PDF y-up coords (shape_h - scaled),
    // so boundary y=0 is bottom of envelope, y=boundary_h is top.
    // Anchor envelope top to textbox top.
    let envelope_top = top_boundary.max_y();

    // Helper closure: transform a glyph point to PDF page coordinates
    let transform = |gx_font: f64, gy_font: f64, cursor_x: f64| -> (f32, f32) {
        // gx in text advance space
        let gx = cursor_x + gx_font * scale;
        // gy: shift from font-space (baseline=0, up=positive) to flat text space (0..text_h)
        let gy = gy_font * scale - descender;

        let (wx, wy) = warp_point(gx, gy, text_w, text_h, boundary_w, &top_boundary, &bottom_boundary);

        // wx is in boundary x-space [0, boundary_w]
        // wy is in PDF y-up space (already converted by geometry engine)
        let pdf_x = (tb_x as f64 + wx) as f32;
        let pdf_y = (tb_y_top as f64 - envelope_top + wy) as f32;
        (pdf_x, pdf_y)
    };

    // For each character, extract glyph outline and warp it
    let mut cursor_x = 0.0_f64;
    // Track previous endpoint for quad-to-cubic conversion
    let mut prev_x = 0.0_f64;
    let mut prev_y = 0.0_f64;

    for &(ch, advance) in &char_advances {
        if let Some(glyph) = extract_glyph_path(&face, ch) {
            for cmd in &glyph.commands {
                match cmd {
                    GlyphCommand::MoveTo(gx, gy) => {
                        let (px, py) = transform(*gx as f64, *gy as f64, cursor_x);
                        content.move_to(px, py);
                        prev_x = *gx as f64;
                        prev_y = *gy as f64;
                    }
                    GlyphCommand::LineTo(gx, gy) => {
                        let (px, py) = transform(*gx as f64, *gy as f64, cursor_x);
                        content.line_to(px, py);
                        prev_x = *gx as f64;
                        prev_y = *gy as f64;
                    }
                    GlyphCommand::QuadTo(qcx, qcy, qx, qy) => {
                        // Quad-to-cubic: cp1 = p0 + 2/3*(qcp - p0), cp2 = p1 + 2/3*(qcp - p1)
                        let p0x = prev_x;
                        let p0y = prev_y;
                        let qcx = *qcx as f64;
                        let qcy = *qcy as f64;
                        let p1x = *qx as f64;
                        let p1y = *qy as f64;

                        let cp1x = p0x + (2.0 / 3.0) * (qcx - p0x);
                        let cp1y = p0y + (2.0 / 3.0) * (qcy - p0y);
                        let cp2x = p1x + (2.0 / 3.0) * (qcx - p1x);
                        let cp2y = p1y + (2.0 / 3.0) * (qcy - p1y);

                        let (wx1, wy1) = transform(cp1x, cp1y, cursor_x);
                        let (wx2, wy2) = transform(cp2x, cp2y, cursor_x);
                        let (wx3, wy3) = transform(p1x, p1y, cursor_x);
                        content.cubic_to(wx1, wy1, wx2, wy2, wx3, wy3);

                        prev_x = p1x;
                        prev_y = p1y;
                    }
                    GlyphCommand::CubicTo(x1, y1, x2, y2, x3, y3) => {
                        let (wx1, wy1) = transform(*x1 as f64, *y1 as f64, cursor_x);
                        let (wx2, wy2) = transform(*x2 as f64, *y2 as f64, cursor_x);
                        let (wx3, wy3) = transform(*x3 as f64, *y3 as f64, cursor_x);
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
        cursor_x += advance;
    }

    // Fill all glyph paths
    let has_outline = first_outline.is_some();
    let no_fill = matches!(&first_fill, Some(TextFill::NoFill));

    if has_outline && !no_fill {
        // Need to fill AND stroke — use fill_nonzero first, then re-emit paths for stroke
        content.fill_nonzero();
    } else if no_fill {
        // NoFill: don't fill, just end the path (stroke happens below)
        content.end_path();
    } else {
        content.fill_nonzero();
    }

    // Apply text outline by re-emitting glyph paths and stroking
    if let Some(ref outline) = first_outline {
        content.set_line_width(outline.width_pt);
        let [r, g, b] = outline.color;
        content.set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);

        cursor_x = 0.0;
        for &(ch, advance) in &char_advances {
            if let Some(glyph) = extract_glyph_path(&face, ch) {
                prev_x = 0.0;
                prev_y = 0.0;
                for cmd in &glyph.commands {
                    match cmd {
                        GlyphCommand::MoveTo(gx, gy) => {
                            let (px, py) = transform(*gx as f64, *gy as f64, cursor_x);
                            content.move_to(px, py);
                            prev_x = *gx as f64;
                            prev_y = *gy as f64;
                        }
                        GlyphCommand::LineTo(gx, gy) => {
                            let (px, py) = transform(*gx as f64, *gy as f64, cursor_x);
                            content.line_to(px, py);
                            prev_x = *gx as f64;
                            prev_y = *gy as f64;
                        }
                        GlyphCommand::QuadTo(qcx, qcy, qx, qy) => {
                            let p0x = prev_x;
                            let p0y = prev_y;
                            let qcx = *qcx as f64;
                            let qcy = *qcy as f64;
                            let p1x = *qx as f64;
                            let p1y = *qy as f64;
                            let cp1x = p0x + (2.0 / 3.0) * (qcx - p0x);
                            let cp1y = p0y + (2.0 / 3.0) * (qcy - p0y);
                            let cp2x = p1x + (2.0 / 3.0) * (qcx - p1x);
                            let cp2y = p1y + (2.0 / 3.0) * (qcy - p1y);
                            let (wx1, wy1) = transform(cp1x, cp1y, cursor_x);
                            let (wx2, wy2) = transform(cp2x, cp2y, cursor_x);
                            let (wx3, wy3) = transform(p1x, p1y, cursor_x);
                            content.cubic_to(wx1, wy1, wx2, wy2, wx3, wy3);
                            prev_x = p1x;
                            prev_y = p1y;
                        }
                        GlyphCommand::CubicTo(x1, y1, x2, y2, x3, y3) => {
                            let (wx1, wy1) = transform(*x1 as f64, *y1 as f64, cursor_x);
                            let (wx2, wy2) = transform(*x2 as f64, *y2 as f64, cursor_x);
                            let (wx3, wy3) = transform(*x3 as f64, *y3 as f64, cursor_x);
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
            cursor_x += advance;
        }
        content.stroke();
    }

    content.restore_state();
    true
}
