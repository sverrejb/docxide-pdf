use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{Paragraph, SectionProperties, TextAnchor, Textbox};

use super::color::{fill_rgb, stroke_rgb};
use super::header_footer::resolve_tb_y_top;
use super::layout::{
    LinkAnnotation, build_paragraph_lines, build_tabbed_line, render_paragraph_lines,
    tallest_run_metrics,
};
use super::list_label::render_list_label;
use super::positioning::resolve_h_position;
use super::smartart::draw_shape_path;
use super::wordart;
use super::{GradientSpec, RenderContext, render_shape_fill, resolve_line_h};

pub(super) fn render_single_textbox(
    tb: &Textbox,
    sp: &SectionProperties,
    col_x: f32,
    col_w: f32,
    text_width: f32,
    slot_top: f32,
    content: &mut Content,
    gradient_specs: &mut Vec<GradientSpec>,
    ctx: &RenderContext,
    page_links: &mut Vec<LinkAnnotation>,
) {
    let tb_x = resolve_h_position(
        tb.h_relative_from,
        &tb.h_position,
        tb.width_pt,
        sp,
        col_x,
        col_w,
        text_width,
    );
    let tb_y_top = resolve_tb_y_top(
        tb.v_relative_from,
        &tb.v_position,
        tb.height_pt,
        sp,
        slot_top,
    );


    // For AutoFit::Shape, compute height from content instead of using tb.height_pt
    let tb_height = if matches!(tb.auto_fit, crate::model::AutoFit::Shape) {
        let tmp_w = if tb.no_text_wrap {
            10000.0
        } else {
            (tb.width_pt - tb.margin_left - tb.margin_right).max(0.0)
        };
        let empty_imgs: HashMap<usize, String> = HashMap::new();
        let empty_fx: HashMap<usize, super::images::EffectXObjs> = HashMap::new();
        let mut h = 0.0f32;
        for tp in &tb.paragraphs {
            let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
            let tw = (tmp_w - tp.indent_left - tp.indent_right).max(1.0);
            let hang = if !tp.list_label.is_empty() {
                if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 {
                    -tp.indent_first_line
                } else {
                    0.0
                }
            } else if tp.indent_hanging > 0.0 {
                tp.indent_hanging
            } else {
                -tp.indent_first_line
            };
            let lines = if tp.runs.iter().any(|r| r.is_tab) {
                build_tabbed_line(
                    &tp.runs, ctx.fonts, &tp.tab_stops, tp.indent_left,
                    tw, tp.indent_right, hang, &empty_imgs, &empty_fx, ctx.default_tab_stop,
                )
            } else {
                build_paragraph_lines(
                    &tp.runs, ctx.fonts, tw, hang, &empty_imgs, &empty_fx, None, None, None, true,
                )
            };
            let (fs, lhr, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
            let lh = resolve_line_h(tp_ls, fs, lhr);
            h += tp.space_before + lines.len().max(1) as f32 * lh + tp.space_after;
        }
        h + tb.margin_top + tb.margin_bottom
    } else {
        tb.height_pt
    };

    if let Some(ref fill) = tb.fill {
        render_shape_fill(
            content,
            fill,
            tb_x,
            tb_y_top - tb_height,
            tb.width_pt,
            tb_height,
            &tb.shape_type,
            gradient_specs,
        );
    }

    if let Some(stroke) = tb.stroke_color {
        if tb.stroke_width > 0.0 {
            content.save_state();
            content.set_line_width(tb.stroke_width);
            stroke_rgb(content, stroke);
            draw_shape_path(
                content,
                tb_x,
                tb_y_top - tb_height,
                tb.width_pt,
                tb_height,
                &tb.shape_type,
            );
            content.stroke();
            content.restore_state();
        }
    }

    let content_x = tb_x + tb.margin_left;
    let natural_w = (tb.width_pt - tb.margin_left - tb.margin_right).max(0.0);
    // no_text_wrap uses a huge width for line-breaking to prevent wrapping,
    // but alignment still uses the natural textbox width
    let content_w = if tb.no_text_wrap { 10000.0 } else { natural_w };
    let align_w = natural_w;

    // Clip text content to textbox bounds for fixed-size textboxes (Word clips overflow)
    let needs_clip = !matches!(tb.auto_fit, crate::model::AutoFit::Shape);
    if needs_clip {
        content.save_state();
        content.rect(tb_x, tb_y_top - tb_height, tb.width_pt, tb_height);
        content.clip_nonzero();
        content.end_path();
    }

    // Text warp: render glyphs as warped paths instead of normal text
    if tb
        .text_warp
        .as_ref()
        .is_some_and(|w| w.preset != "textNoShape")
    {
        if wordart::render_warped_textbox(tb, content, ctx.fonts, tb_x, tb_y_top, align_w) {
            if needs_clip { content.restore_state(); }
            return;
        }
        if wordart::render_text_on_path(tb, content, ctx.fonts, tb_x, tb_y_top, align_w) {
            if needs_clip { content.restore_state(); }
            return;
        }
    }

    let anchor_offset = match tb.text_anchor {
        TextAnchor::Top => 0.0,
        TextAnchor::Middle | TextAnchor::Bottom => {
            let empty_inline_imgs_pre: HashMap<usize, String> = HashMap::new();
            let empty_fx_pre: HashMap<usize, super::images::EffectXObjs> = HashMap::new();
            let mut total_h = 0.0f32;
            for tp in &tb.paragraphs {
                let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let tp_text_w = (content_w - tp.indent_left - tp.indent_right).max(1.0);
                let text_hanging = if !tp.list_label.is_empty() {
                    if let Some(nts) = tp.num_level_tab_stop {
                        if nts < tp.indent_left && (tp.indent_left - tp.indent_hanging).abs() < 0.5 {
                            (tp.indent_left - nts).max(0.0)
                        } else if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 {
                            -tp.indent_first_line
                        } else {
                            0.0
                        }
                    } else if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 {
                        -tp.indent_first_line
                    } else {
                        0.0
                    }
                } else if tp.indent_hanging > 0.0 {
                    tp.indent_hanging
                } else {
                    -tp.indent_first_line
                };
                let has_tabs = tp.runs.iter().any(|r| r.is_tab);
                let lines = if has_tabs {
                    build_tabbed_line(
                        &tp.runs, ctx.fonts, &tp.tab_stops, tp.indent_left,
                        tp_text_w, tp.indent_right, text_hanging, &empty_inline_imgs_pre,
                        &empty_fx_pre, ctx.default_tab_stop,
                    )
                } else {
                    build_paragraph_lines(
                        &tp.runs, ctx.fonts, tp_text_w, text_hanging, &empty_inline_imgs_pre, &empty_fx_pre, None, None, None, true,
                    )
                };
                let (fs, lhr, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
                let lh = resolve_line_h(tp_ls, fs, lhr);
                let n = lines.len().max(1) as f32;
                total_h += tp.space_before + n * lh + tp.space_after;
            }
            let available = tb_height - tb.margin_top - tb.margin_bottom;
            let gap = (available - total_h).max(0.0);
            match tb.text_anchor {
                TextAnchor::Middle => gap / 2.0,
                TextAnchor::Bottom => gap,
                TextAnchor::Top => 0.0,
            }
        }
    };

    // For fixed-size textboxes, stop rendering content below the textbox bounds
    let clip_bottom = if needs_clip {
        Some(tb_y_top - tb_height)
    } else {
        None
    };

    // Glow pass: render text as thick stroke in glow color behind everything
    if let Some(glow) = wordart::find_text_glow(tb) {
        content.save_state();
        stroke_rgb(content, glow.color);
        content.set_line_width(glow.radius_pt * 2.0);
        content.set_line_join(pdf_writer::types::LineJoinStyle::RoundJoin);
        content.set_text_rendering_mode(pdf_writer::types::TextRenderingMode::Stroke);
        let mut discard_links: Vec<LinkAnnotation> = Vec::new();
        render_textbox_paragraphs(
            &tb.paragraphs, content, content_x, content_w, align_w,
            tb_y_top - tb.margin_top - anchor_offset,
            0.0, 0.0, None, false, &mut discard_links, ctx, clip_bottom,
        );
        content.restore_state();
    }

    // Shadow pass: render all text offset by shadow distance in shadow color
    if let Some(shadow) = wordart::find_text_shadow(tb) {
        content.save_state();
        let [sr, sg, sb] = shadow.color;
        let shadow_color = [
            (sr as f32 * shadow.alpha) as u8,
            (sg as f32 * shadow.alpha) as u8,
            (sb as f32 * shadow.alpha) as u8,
        ];
        let mut discard_links: Vec<LinkAnnotation> = Vec::new();
        render_textbox_paragraphs(
            &tb.paragraphs, content, content_x, content_w, align_w,
            tb_y_top - tb.margin_top - anchor_offset,
            shadow.offset_x, shadow.offset_y, Some(shadow_color),
            false, &mut discard_links, ctx, clip_bottom,
        );
        content.restore_state();
    }

    render_textbox_paragraphs(
        &tb.paragraphs, content, content_x, content_w, align_w,
        tb_y_top - tb.margin_top - anchor_offset,
        0.0, 0.0, None, true, page_links, ctx, clip_bottom,
    );

    if needs_clip {
        content.restore_state();
    }
}

/// Shared paragraph iteration loop used by the glow, shadow, and normal text passes
/// within a textbox. Each pass differs only in positional offsets, forced color override,
/// and whether list labels are rendered.
pub(super) fn render_textbox_paragraphs(
    paragraphs: &[Paragraph],
    content: &mut Content,
    content_x: f32,
    content_w: f32,
    align_w: f32,
    start_y: f32,
    x_offset: f32,
    y_offset: f32,
    force_color: Option<[u8; 3]>,
    render_labels: bool,
    links: &mut Vec<LinkAnnotation>,
    ctx: &RenderContext,
    clip_bottom: Option<f32>,
) {
    let mut cursor_y = start_y;
    let mut prev_space_after = 0.0f32;
    let empty_imgs: HashMap<usize, String> = HashMap::new();
    let empty_fx: HashMap<usize, super::images::EffectXObjs> = HashMap::new();
    for (tp_idx, tp) in paragraphs.iter().enumerate() {
        // Collapse adjacent spacing: use max(prev_after, current_before) like body text
        let inter_gap = if tp_idx == 0 {
            tp.space_before
        } else {
            prev_space_after.max(tp.space_before)
        };
        // Stop rendering when content overflows the textbox bounds
        if let Some(bottom) = clip_bottom {
            if cursor_y - inter_gap < bottom {
                break;
            }
        }
        let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
        let tp_text_w = (content_w - tp.indent_left - tp.indent_right).max(1.0);
        let tp_align_w = (align_w - tp.indent_left - tp.indent_right).max(1.0);
        let text_hanging = if !tp.list_label.is_empty() {
            if let Some(nts) = tp.num_level_tab_stop {
                (tp.indent_left - nts).max(0.0)
            } else if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 {
                -tp.indent_first_line
            } else {
                0.0
            }
        } else if tp.indent_hanging > 0.0 {
            tp.indent_hanging
        } else {
            -tp.indent_first_line
        };
        let tb_lines = if tp.runs.iter().any(|r| r.is_tab) {
            build_tabbed_line(
                &tp.runs, ctx.fonts, &tp.tab_stops, tp.indent_left,
                tp_text_w, tp.indent_right, text_hanging, &empty_imgs, &empty_fx, ctx.default_tab_stop,
            )
        } else {
            build_paragraph_lines(
                &tp.runs, ctx.fonts, tp_text_w, text_hanging, &empty_imgs, &empty_fx, None, None, None, true,
            )
        };
        if tb_lines.is_empty() {
            let (fs, lhr, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
            let lh = resolve_line_h(tp_ls, fs, lhr);
            cursor_y -= inter_gap + lh;
            prev_space_after = tp.space_after;
            continue;
        }
        let (tb_fs, tb_lhr, tb_ar) = tallest_run_metrics(&tp.runs, ctx.fonts);
        let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_lhr);
        let tb_baseline = cursor_y - inter_gap - tb_fs * tb_ar.unwrap_or(0.75) - y_offset;
        let tp_text_x = content_x + tp.indent_left + x_offset;
        if let Some(c) = force_color {
            fill_rgb(content, c);
        }
        if render_labels {
            render_list_label(
                content,
                tp,
                ctx.fonts,
                content_x + tp.indent_left - tp.indent_hanging,
                tb_baseline,
                tb_fs,
            );
        }
        render_paragraph_lines(
            content, &tb_lines, &tp.alignment, tp_text_x, tp_align_w,
            tb_baseline, tb_line_h, tb_lines.len(), 0,
            links, 0.0, ctx.fonts, None,
        );
        cursor_y -= inter_gap + (tb_lines.len() as f32) * tb_line_h;
        prev_space_after = tp.space_after;
    }
}
