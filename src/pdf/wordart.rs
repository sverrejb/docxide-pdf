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
/// Returns warped (x, y) in PDF coordinates.
fn warp_point(
    gx: f64,
    gy: f64,
    text_w: f64,
    text_h: f64,
    top: &SampledBoundary,
    bottom: &SampledBoundary,
) -> (f64, f64) {
    // u = horizontal position (0..text_w maps to boundary x range)
    let u = if text_w > 0.0 { gx / text_w } else { 0.5 };

    // v = vertical position (0 = bottom, 1 = top of text)
    let v = if text_h > 0.0 { gy / text_h } else { 0.5 };

    // Sample boundaries at this x position
    let x_pos = gx; // boundaries are in the same coordinate space
    let top_y = top.y_at(x_pos);
    let bot_y = bottom.y_at(x_pos);

    // Interpolate between bottom and top
    let warped_x = gx; // x stays the same (boundaries define y distortion)
    let warped_y = bot_y + v * (top_y - bot_y);

    // For warps that also move x (like textWave), we need to interpolate x too.
    // The boundary x-coordinates may not be uniformly distributed.
    // For now, use a simple approach: x comes from the boundary at parameter u.
    let _ = u; // future: use u for x-interpolation on circular warps

    (warped_x, warped_y)
}

/// Render a warped textbox: extract glyph outlines, warp them through the envelope,
/// and emit as filled PDF paths.
pub(super) fn render_warped_textbox(
    tb: &Textbox,
    content: &mut Content,
    seen_fonts: &HashMap<String, FontEntry>,
    tb_x: f32,
    tb_y_top: f32,
    content_w: f32,
) {
    let warp = match &tb.text_warp {
        Some(w) => w,
        None => return,
    };

    // Collect all text and compute total width
    let mut total_text = String::new();
    let mut font_name = String::new();
    let mut font_size = 12.0_f32;
    let mut text_color: Option<[u8; 3]> = None;
    for para in &tb.paragraphs {
        for run in &para.runs {
            if !run.text.is_empty() {
                total_text.push_str(&run.text);
                if font_name.is_empty() {
                    font_name = run.font_name.clone();
                    font_size = run.font_size;
                    text_color = run.color;
                }
            }
        }
    }

    if total_text.is_empty() || font_name.is_empty() {
        return;
    }

    // Load font for glyph outlines
    let Some((font_data, face_index)) = load_font_data(&font_name, seen_fonts) else {
        return;
    };
    let Some(face) = ttf_parser::Face::parse(&font_data, face_index).ok() else {
        return;
    };

    let units_per_em = face.units_per_em() as f64;
    let scale = font_size as f64 / units_per_em;

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
    let text_h = font_size as f64;

    // Evaluate warp boundaries
    let Some((top_boundary, bottom_boundary)) =
        evaluate_warp_boundaries(&warp.preset, &warp.adjustments, content_w as f64, text_h)
    else {
        return;
    };

    // Set up rendering
    content.save_state();
    if let Some([r, g, b]) = text_color {
        content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
    } else {
        content.set_fill_gray(0.0);
    }

    // Horizontal centering offset
    let x_offset = (content_w as f64 - text_w) / 2.0;

    // For each character, extract glyph outline and warp it
    let mut cursor_x = 0.0_f64;
    let ascender = face.ascender() as f64 / units_per_em * font_size as f64;

    for &(ch, advance) in &char_advances {
        if let Some(glyph) = extract_glyph_path(&face, ch) {
            for cmd in &glyph.commands {
                match cmd {
                    GlyphCommand::MoveTo(gx, gy) => {
                        let px = cursor_x + *gx as f64 * scale;
                        let py = *gy as f64 * scale;
                        let (wx, wy) =
                            warp_point(px, py, text_w, text_h, &top_boundary, &bottom_boundary);
                        content.move_to(
                            (tb_x as f64 + x_offset + wx) as f32,
                            (tb_y_top as f64 - text_h + wy) as f32,
                        );
                    }
                    GlyphCommand::LineTo(gx, gy) => {
                        let px = cursor_x + *gx as f64 * scale;
                        let py = *gy as f64 * scale;
                        let (wx, wy) =
                            warp_point(px, py, text_w, text_h, &top_boundary, &bottom_boundary);
                        content.line_to(
                            (tb_x as f64 + x_offset + wx) as f32,
                            (tb_y_top as f64 - text_h + wy) as f32,
                        );
                    }
                    GlyphCommand::QuadTo(x1, y1, x2, y2) => {
                        // Convert quad to cubic for PDF
                        let qx1 = cursor_x + *x1 as f64 * scale;
                        let qy1 = *y1 as f64 * scale;
                        let qx2 = cursor_x + *x2 as f64 * scale;
                        let qy2 = *y2 as f64 * scale;

                        // Get previous point for quad-to-cubic conversion
                        let (wx1, wy1) = warp_point(
                            qx1, qy1, text_w, text_h, &top_boundary, &bottom_boundary,
                        );
                        let (wx2, wy2) = warp_point(
                            qx2, qy2, text_w, text_h, &top_boundary, &bottom_boundary,
                        );
                        // Approximate as line segments for warped quads
                        // (warping distorts control points, so subdivision is more correct)
                        content.line_to(
                            (tb_x as f64 + x_offset + wx1) as f32,
                            (tb_y_top as f64 - text_h + wy1) as f32,
                        );
                        content.line_to(
                            (tb_x as f64 + x_offset + wx2) as f32,
                            (tb_y_top as f64 - text_h + wy2) as f32,
                        );
                    }
                    GlyphCommand::CubicTo(x1, y1, x2, y2, x3, y3) => {
                        let cx1 = cursor_x + *x1 as f64 * scale;
                        let cy1 = *y1 as f64 * scale;
                        let cx2 = cursor_x + *x2 as f64 * scale;
                        let cy2 = *y2 as f64 * scale;
                        let cx3 = cursor_x + *x3 as f64 * scale;
                        let cy3 = *y3 as f64 * scale;

                        let (wx1, wy1) = warp_point(
                            cx1, cy1, text_w, text_h, &top_boundary, &bottom_boundary,
                        );
                        let (wx2, wy2) = warp_point(
                            cx2, cy2, text_w, text_h, &top_boundary, &bottom_boundary,
                        );
                        let (wx3, wy3) = warp_point(
                            cx3, cy3, text_w, text_h, &top_boundary, &bottom_boundary,
                        );
                        content.cubic_to(
                            (tb_x as f64 + x_offset + wx1) as f32,
                            (tb_y_top as f64 - text_h + wy1) as f32,
                            (tb_x as f64 + x_offset + wx2) as f32,
                            (tb_y_top as f64 - text_h + wy2) as f32,
                            (tb_x as f64 + x_offset + wx3) as f32,
                            (tb_y_top as f64 - text_h + wy3) as f32,
                        );
                    }
                    GlyphCommand::Close => {
                        content.close_path();
                    }
                }
            }
        }
        cursor_x += advance;
    }

    // Fill all glyph paths at once
    content.fill_nonzero();
    content.restore_state();
}
