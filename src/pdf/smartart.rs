use std::collections::HashMap;

use pdf_writer::Content;

use crate::fonts::FontEntry;
use crate::model::{SmartArtDiagram, SmartArtTextAlign, SmartArtTextAnchor};

use super::charts;
use super::color;
// The shape helpers grew up here and the rest of the renderer still reaches
// them through this module.
pub(super) use super::shape_path::{
    draw_shape_path, draw_shape_stroke_path, render_image_with_clip, shape_text_rect,
    stroke_image_border,
};

pub(super) fn render_smartart(
    content: &mut Content,
    diagram: &SmartArtDiagram,
    diag_x: f32,
    diag_y: f32,
    seen_fonts: &HashMap<String, FontEntry>,
    smartart_font_key: &str,
    image_names: &HashMap<usize, String>,
) {
    let sa_font_entry = seen_fonts.get(smartart_font_key);
    let sa_font_pdf_name = sa_font_entry.map(|e| e.pdf_name.as_str()).unwrap_or("F1");

    let apply_rotation =
        |content: &mut Content, shape: &crate::model::SmartArtShape, sx: f32, sy: f32| -> bool {
            let rotated = shape.rotation_deg.abs() > 0.01;
            if rotated {
                content.save_state();
                super::positioning::push_center_rotation(
                    content,
                    sx,
                    sy,
                    shape.width,
                    shape.height,
                    shape.rotation_deg,
                );
            }
            rotated
        };

    // Pass 1: Fill all shapes
    for shape in &diagram.shapes {
        if shape.fill.is_none() && shape.image_fill.is_none() {
            continue;
        }
        let sx = diag_x + shape.x;
        let sy = diag_y - shape.y - shape.height;
        let rotated = apply_rotation(content, shape, sx, sy);

        if let Some(fill) = shape.fill {
            content.save_state();
            color::fill_rgb(content, fill);
            draw_shape_path(
                content,
                sx,
                sy,
                shape.width,
                shape.height,
                &shape.shape_type,
            );
            content.fill_nonzero();
            content.restore_state();
        }

        if let Some(ref img) = shape.image_fill {
            if let Some(pdf_name) = image_names.get(&img.id) {
                content.save_state();
                draw_shape_path(
                    content,
                    sx,
                    sy,
                    shape.width,
                    shape.height,
                    &shape.shape_type,
                );
                content.clip_nonzero();
                content.end_path();
                let img_aspect = img.pixel_width as f32 / img.pixel_height.max(1) as f32;
                let shape_aspect = shape.width / shape.height.max(0.001);
                let (draw_w, draw_h) = if img_aspect > shape_aspect {
                    (shape.height * img_aspect, shape.height)
                } else {
                    (shape.width, shape.width / img_aspect)
                };
                let dx = sx + (shape.width - draw_w) / 2.0;
                let dy = sy + (shape.height - draw_h) / 2.0;
                content.transform([draw_w, 0.0, 0.0, draw_h, dx, dy]);
                content.x_object(pdf_writer::Name(pdf_name.as_bytes()));
                content.restore_state();
            }
        }

        if rotated {
            content.restore_state();
        }
    }

    // Pass 2: Stroke all shapes (on top of all fills)
    for shape in &diagram.shapes {
        let has_stroke = shape.stroke_color.is_some() && shape.stroke_width > 0.0;
        if !has_stroke {
            continue;
        }
        let sx = diag_x + shape.x;
        let sy = diag_y - shape.y - shape.height;
        let rotated = apply_rotation(content, shape, sx, sy);

        content.save_state();
        if let Some(stroke) = shape.stroke_color {
            content.set_line_width(shape.stroke_width);
            color::stroke_rgb(content, stroke);
        }
        draw_shape_stroke_path(
            content,
            sx,
            sy,
            shape.width,
            shape.height,
            &shape.shape_type,
        );
        content.stroke();
        content.restore_state();

        if rotated {
            content.restore_state();
        }
    }

    // Pass 3: Render all text on top of geometry (with word-wrapping)
    for shape in &diagram.shapes {
        if shape.paragraphs.is_empty() {
            continue;
        }

        let (txt_x, txt_y, txt_w, txt_h) =
            shape
                .text_rect
                .unwrap_or((shape.x, shape.y, shape.width, shape.height));

        let (ins_top, ins_right, ins_bottom, ins_left) = shape.text_insets;
        let avail_w = (txt_w - ins_left - ins_right).max(1.0);

        let base_fs = shape
            .paragraphs
            .iter()
            .flat_map(|p| p.runs.iter())
            .map(|r| r.font_size)
            .find(|&fs| fs > 0.0)
            .unwrap_or(10.0);

        // Build all wrapped lines across all paragraphs first (for vertical centering)
        /// A word, a space or a bullet, measured in its run's font.
        #[derive(Clone, Copy)]
        struct Piece<'a> {
            text: &'a str,
            width: f32,
            space_w: f32, // width of space in this run's font/size
            fe: Option<&'a FontEntry>,
            run: &'a crate::model::SmartArtRun,
        }
        struct WrappedLine<'a> {
            pieces: Vec<Piece<'a>>,
            total_w: f32,
            align: SmartArtTextAlign,
            line_h: f32,
        }

        // A bullet is the paragraph's first word; the strings live here so
        // the pieces can borrow them like run text.
        let bullets: Vec<Option<String>> = shape
            .paragraphs
            .iter()
            .map(|p| p.bullet.as_deref().map(|b| format!("{b} ")))
            .collect();
        let mut all_lines: Vec<WrappedLine<'_>> = Vec::new();

        for (para, bullet) in shape.paragraphs.iter().zip(&bullets) {
            let para_fs = para
                .runs
                .iter()
                .map(|r| r.font_size)
                .find(|&fs| fs > 0.0)
                .unwrap_or(base_fs);
            // spcPct is percentage of single spacing; single ≈ 1.2× font size
            let line_h = if para.line_spacing_pct > 0.0 {
                para_fs * 1.2 * para.line_spacing_pct
            } else {
                para_fs * 1.2
            };

            // Collect word segments: split each run's text by spaces
            let mut words: Vec<Piece<'_>> = Vec::new();

            if let Some(bullet_text) = bullet.as_deref()
                && let Some(first_run) = para.runs.first()
            {
                let fe = resolve_run_font(first_run, seen_fonts, sa_font_entry);
                let bw = charts::text_width(bullet_text, first_run.font_size, fe);
                words.push(Piece {
                    text: bullet_text,
                    width: bw,
                    space_w: 0.0,
                    fe,
                    run: first_run,
                });
            }

            for run in &para.runs {
                let fe = resolve_run_font(run, seen_fonts, sa_font_entry);
                let efs = effective_font_size(run);
                let space_w = charts::text_width(" ", efs, fe);

                for (i, word) in run.text.split(' ').enumerate() {
                    if word.is_empty() && i > 0 {
                        continue;
                    }
                    let ww = if word.is_empty() {
                        0.0
                    } else {
                        charts::text_width(word, efs, fe)
                    };
                    words.push(Piece {
                        text: word,
                        width: ww,
                        space_w,
                        fe,
                        run,
                    });
                }
            }

            // Greedy line-fill
            let mut cur_line: Vec<Piece<'_>> = Vec::new();
            let mut cur_w = 0.0_f32;

            for &seg in &words {
                let needed = if cur_line.is_empty() {
                    seg.width
                } else {
                    seg.space_w + seg.width
                };
                if !cur_line.is_empty() && cur_w + needed > avail_w + 0.05 {
                    let total_w = cur_w;
                    all_lines.push(WrappedLine {
                        pieces: std::mem::take(&mut cur_line),
                        total_w,
                        align: para.align,
                        line_h,
                    });
                    cur_w = 0.0;
                }
                if !cur_line.is_empty() {
                    // Add space before this word
                    cur_w += seg.space_w;
                    cur_line.push(Piece {
                        text: " ",
                        width: seg.space_w,
                        ..seg
                    });
                }
                cur_line.push(seg);
                cur_w += seg.width;
            }
            if !cur_line.is_empty() {
                let total_w = cur_w;
                all_lines.push(WrappedLine {
                    pieces: cur_line,
                    total_w,
                    align: para.align,
                    line_h,
                });
            }
        }

        // Compute total text height and vertical anchor position
        let total_text_h: f32 = all_lines.iter().map(|l| l.line_h).sum();
        let content_h = txt_h - ins_top - ins_bottom;
        let text_top_y = match shape.text_anchor {
            SmartArtTextAnchor::Top => diag_y - txt_y - ins_top,
            SmartArtTextAnchor::Center => {
                diag_y - txt_y - ins_top - (content_h - total_text_h) / 2.0
            }
            SmartArtTextAnchor::Bottom => diag_y - txt_y - ins_top - (content_h - total_text_h),
        };

        // Distance from the text block's top to the first baseline. The baseline
        // sits at the font's ascent within the line box, which is a fraction
        // (ascender/line-height) of the line height — not the full em. Using the
        // full em (base_fs) drops every line ~0.15em too low, leaving centered
        // text noticeably below center. Derive the fraction from the first line's
        // resolved font metrics, defaulting to a typical ascent ratio.
        let first_ascent = {
            let fl = all_lines.first();
            let frac = fl
                .and_then(|l| l.pieces.first())
                .and_then(|p| p.fe)
                .and_then(|e| match (e.ascender_ratio, e.line_h_ratio) {
                    (Some(a), Some(lh)) if lh > 0.0 => Some(a / lh),
                    _ => None,
                })
                .unwrap_or(0.85);
            fl.map(|l| l.line_h * frac).unwrap_or(base_fs)
        };

        // Render each wrapped line
        let mut y_cursor = 0.0_f32;
        for line in &all_lines {
            let line_x = diag_x
                + txt_x
                + ins_left
                + match line.align {
                    SmartArtTextAlign::Left => 0.0,
                    SmartArtTextAlign::Center => (avail_w - line.total_w) / 2.0,
                    SmartArtTextAlign::Right => avail_w - line.total_w,
                };
            let line_y = text_top_y - first_ascent - y_cursor;

            content.save_state();
            let mut cx = line_x;

            for piece in &line.pieces {
                let efs = effective_font_size(piece.run);
                let fpn = piece
                    .fe
                    .map(|e| e.pdf_name.as_str())
                    .unwrap_or(sa_font_pdf_name);
                let y_offset = baseline_y_offset(piece.run, base_fs);
                let ry = line_y + y_offset;

                if let Some(hl) = piece.run.highlight {
                    content.save_state();
                    color::fill_rgb(content, hl);
                    content.rect(cx, ry - base_fs * 0.15, piece.width, base_fs * 1.15);
                    content.fill_nonzero();
                    content.restore_state();
                }

                if let Some(c) = piece.run.color {
                    color::fill_rgb(content, c);
                } else {
                    content.set_fill_gray(0.0);
                }
                let needs_synthetic_bold =
                    piece.run.bold && piece.fe.is_some_and(|e| e.synthetic_bold);
                if needs_synthetic_bold {
                    content.set_line_width(efs * 0.02);
                    if let Some(c) = piece.run.color {
                        color::stroke_rgb(content, c);
                    } else {
                        content.set_stroke_gray(0.0);
                    }
                    content
                        .set_text_rendering_mode(pdf_writer::types::TextRenderingMode::FillStroke);
                }
                charts::show_text(content, fpn, efs, cx, ry, piece.text, piece.fe);
                if needs_synthetic_bold {
                    content.set_text_rendering_mode(pdf_writer::types::TextRenderingMode::Fill);
                }

                if piece.run.underline {
                    let thick = (efs * 0.05).max(0.5);
                    let ul_y = ry - efs * 0.12;
                    content.rect(cx, ul_y - thick, piece.width, thick);
                    content.fill_nonzero();
                }

                if piece.run.strikethrough {
                    let thick = (efs * 0.05).max(0.5);
                    let st_y = ry + efs * 0.3;
                    content.rect(cx, st_y, piece.width, thick);
                    content.fill_nonzero();
                }

                cx += piece.width;
            }
            content.restore_state();
            y_cursor += line.line_h;
        }
    }
}

/// Resolve the font entry for a SmartArt run, falling back to the
/// diagram-level default font when the run has no per-shape font.
fn resolve_run_font<'a>(
    run: &crate::model::SmartArtRun,
    seen_fonts: &'a HashMap<String, FontEntry>,
    fallback: Option<&'a FontEntry>,
) -> Option<&'a FontEntry> {
    run.font_name
        .as_ref()
        .and_then(|n| {
            let key = crate::fonts::font_key_for(n, run.bold, run.italic);
            seen_fonts.get(&key)
        })
        .or(fallback)
}

fn effective_font_size(run: &crate::model::SmartArtRun) -> f32 {
    if run.baseline != 0 {
        run.font_size * 0.58
    } else {
        run.font_size
    }
}

fn baseline_y_offset(run: &crate::model::SmartArtRun, base_fs: f32) -> f32 {
    if run.baseline > 0 {
        base_fs * 0.35
    }
    // superscript
    else if run.baseline < 0 {
        -base_fs * 0.14
    }
    // subscript
    else {
        0.0
    }
}
