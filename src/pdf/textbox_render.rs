use pdf_writer::Content;

use crate::model::{EmbeddedImage, Paragraph, SectionProperties, TextAnchor, Textbox};

use super::color::{fill_rgb, stroke_rgb};
use super::header_footer::resolve_tb_y_top;
use super::helpers::align_offset;
use super::layout::{
    LineOpts, LinkAnnotation, LinkTagger, build_lines, lines_height, picture_line_bottom,
    render_paragraph_lines, tallest_run_metrics,
};
use super::list_label::render_list_label;
use super::positioning::resolve_h_position;
use super::tagging::{Lists, Tags};
use super::wordart;
use super::{GradientSpec, RenderContext, render_shape_fill, resolve_line_h};

/// Returns the block image for an image-only textbox paragraph (no text).
/// Textbox paragraphs are parsed with `resolve_drawings: false`, so the
/// `inline_image_count == 1 && !has_text` promotion in `docx::paragraph`
/// doesn't fire — `tp.image` stays `None` and the image lives on a run.
/// This helper re-derives the block image at render time so callers can
/// size the paragraph as image-height (not text-line-height) and emit the
/// image directly instead of going through line layout.
pub(super) fn textbox_para_block_image(tp: &Paragraph) -> Option<&EmbeddedImage> {
    if let Some(img) = &tp.image {
        return Some(img);
    }
    let has_text = tp.runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
    if has_text {
        return None;
    }
    tp.runs.iter().find_map(|r| r.inline_image.as_ref())
}

fn image_block_height(img: &EmbeddedImage) -> f32 {
    img.display_height + img.layout_extra_height
}

/// Rendered height of a textbox: the stored height, or for `a:spAutoFit` shapes the
/// height of their content (Word shrinks or grows the shape to fit). Shared by body
/// and table-cell textboxes so a cell-anchored autofit box does not keep Word's
/// 110.6pt default and paint over neighbouring shapes (japanese_land_development p1).
pub(super) fn textbox_height(tb: &Textbox, ctx: &RenderContext) -> f32 {
    if matches!(tb.auto_fit, crate::model::AutoFit::Shape) {
        let tmp_w = if tb.no_text_wrap {
            10000.0
        } else {
            tb.text_area(0.0).0.max(0.0)
        };
        let h: f32 = tb
            .paragraphs
            .iter()
            .map(|tp| tp.space_before + paragraph_body_height(tp, tmp_w, ctx) + tp.space_after)
            .sum();
        h + tb.margin_top + tb.margin_bottom
    } else {
        tb.height_pt
    }
}

/// A textbox paragraph's height without its spacing, at text width `text_w`.
fn paragraph_body_height(tp: &Paragraph, text_w: f32, ctx: &RenderContext) -> f32 {
    if let Some(img) = textbox_para_block_image(tp) {
        return image_block_height(img);
    }
    let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let tw = (text_w - tp.indent_left - tp.indent_right).max(1.0);
    let hang = if !tp.list_label.is_empty() {
        match tp.num_level_tab_stop {
            Some(nts)
                if nts < tp.indent_left && (tp.indent_left - tp.indent_hanging).abs() < 0.5 =>
            {
                (tp.indent_left - nts).max(0.0)
            }
            _ if tp.indent_first_line > 0.0 && tp.indent_hanging == 0.0 => -tp.indent_first_line,
            _ => 0.0,
        }
    } else if tp.indent_hanging > 0.0 {
        tp.indent_hanging
    } else {
        -tp.indent_first_line
    };
    let lines = build_lines(
        &tp.runs,
        ctx,
        tw,
        ctx.cjk(true, tp),
        &LineOpts {
            tab_stops: &tp.tab_stops,
            indent_left: tp.indent_left,
            indent_right: tp.indent_right,
            hanging: hang,
            ..Default::default()
        },
    );
    let (fs, lhr, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
    lines.len().max(1) as f32 * resolve_line_h(tp_ls, fs, lhr)
}

/// The part of its story a linked box shows: after what the boxes before it
/// hold, and only what fits when a later box takes the rest (torture_test's
/// second paragraph opens the right-hand box). ponytail: whole paragraphs;
/// Word splits a story between lines.
fn story_window<'t>(tb: &'t Textbox, height: f32, ctx: &RenderContext) -> &'t [Paragraph] {
    let Some(link) = &tb.link else {
        return &tb.paragraphs;
    };
    // A chain outside the body was never linked: its first box keeps the text.
    let mut rest = if link.story.is_empty() {
        &tb.paragraphs[..]
    } else {
        &link.story[..]
    };
    for &area in &link.before {
        rest = &rest[paragraphs_fitting(rest, area, ctx)..];
    }
    if !link.continues {
        return rest;
    }
    &rest[..paragraphs_fitting(rest, tb.text_area(height), ctx)]
}

/// How many of `paras` a text area holds, at least one when there are any.
fn paragraphs_fitting(paras: &[Paragraph], (w, h): (f32, f32), ctx: &RenderContext) -> usize {
    let mut used = 0.0;
    let mut n = 0;
    for tp in paras {
        let bottom = used + tp.space_before + paragraph_body_height(tp, w, ctx);
        if n > 0 && bottom > h + 0.01 {
            break;
        }
        used = bottom + tp.space_after;
        n += 1;
    }
    n
}

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
    // Tags its text, otherwise an artifact, in a Sect hoisted after the anchor
    // paragraph; `content` must be inside the default artifact.
    tags: &mut Tags,
    page: usize,
) {
    let sect = tags.hoist("Sect", tb.anchor_seq);
    let tb_x = resolve_h_position(
        tb.h_relative_from,
        &tb.h_position,
        tb.width_pt,
        sp,
        col_x,
        col_w,
        text_width,
    );
    let tb_height = textbox_height(tb, ctx);
    let paragraphs = story_window(tb, tb_height, ctx);
    let tb_y_top = resolve_tb_y_top(tb.v_relative_from, &tb.v_position, tb_height, sp, slot_top);
    log::debug!(
        "TEXTBOX x={tb_x:.1} top={tb_y_top:.1} w={:.1} h={tb_height:.1} rot={} paras={}",
        tb.width_pt,
        tb.rotation_deg,
        paragraphs.len()
    );
    let rotated =
        rotate_about_center(content, tb.rotation_deg, tb_x, tb_y_top, tb.width_pt, tb_height);

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

    if let Some(stroke) = tb.stroke_color
        && tb.stroke_width > 0.0
    {
        content.save_state();
        content.set_line_width(tb.stroke_width);
        stroke_rgb(content, stroke);
        super::smartart::draw_shape_stroke_path(
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

    // Lay text out within the shape's text rectangle (e.g. an arrow's body,
    // not its head). For plain rectangles this is the full box, so non-shape
    // textboxes are unaffected. Body-margin insets apply within the text rect.
    let (txt_l, txt_b, txt_w, txt_h) = super::smartart::shape_text_rect(
        &tb.shape_type,
        tb.width_pt,
        tb_height,
    )
    .unwrap_or((0.0, 0.0, tb.width_pt, tb_height));
    let text_top = tb_y_top - tb_height + txt_b + txt_h;

    let content_x = tb_x + txt_l + tb.margin_left;
    let natural_w = (txt_w - tb.margin_left - tb.margin_right).max(0.0);
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
        if wordart::render_warped_textbox(tb, content, ctx.fonts, tb_x, tb_y_top, align_w)
            || wordart::render_text_on_path(tb, content, ctx.fonts, tb_x, tb_y_top, align_w)
        {
            // The outlines stay artifacts; the text reads as one P (Word:
            // Sect > P with the text).
            let p = tags.add(sect, "P");
            tags.begin(content, page, p);
            let (x, top) = (tb_x + tb.margin_left, tb_y_top - tb.margin_top);
            wordart::invisible_text(tb, content, ctx.fonts, x, top);
            Tags::end(content);
            if needs_clip {
                content.restore_state();
            }
            if rotated {
                content.restore_state();
            }
            return;
        }
    }

    let anchor_offset = match tb.text_anchor {
        TextAnchor::Top => 0.0,
        TextAnchor::Middle | TextAnchor::Bottom => {
            let total_h: f32 = paragraphs
                .iter()
                .map(|tp| {
                    tp.space_before + paragraph_body_height(tp, content_w, ctx) + tp.space_after
                })
                .sum();
            let available = txt_h - tb.margin_top - tb.margin_bottom;
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
            paragraphs,
            content,
            content_x,
            content_w,
            align_w,
            text_top - tb.margin_top - anchor_offset,
            0.0,
            0.0,
            None,
            false,
            &mut discard_links,
            ctx,
            clip_bottom,
            gradient_specs,
            None,
            None,
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
            paragraphs,
            content,
            content_x,
            content_w,
            align_w,
            text_top - tb.margin_top - anchor_offset,
            shadow.offset_x,
            shadow.offset_y,
            Some(shadow_color),
            false,
            &mut discard_links,
            ctx,
            clip_bottom,
            gradient_specs,
            None,
            None,
        );
        content.restore_state();
    }

    render_textbox_paragraphs(
        paragraphs,
        content,
        content_x,
        content_w,
        align_w,
        text_top - tb.margin_top - anchor_offset,
        0.0,
        0.0,
        None,
        true,
        page_links,
        ctx,
        clip_bottom,
        gradient_specs,
        Some((tags, page, sect)),
        None,
    );

    if needs_clip {
        content.restore_state();
    }
    if rotated {
        content.restore_state();
    }
}

/// Turn the graphics state by a shape's rotation (OOXML: degrees clockwise)
/// about its box centre. Returns whether a state was pushed to restore.
pub(super) fn rotate_about_center(
    content: &mut Content,
    rotation_deg: f32,
    x: f32,
    y_top: f32,
    w: f32,
    h: f32,
) -> bool {
    if rotation_deg.rem_euclid(360.0).abs() < 0.01 {
        return false;
    }
    let (s, c) = (-rotation_deg).to_radians().sin_cos();
    let (cx, cy) = (x + w / 2.0, y_top - h / 2.0);
    content.save_state();
    content.transform([c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]);
    true
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
    gradient_specs: &mut Vec<GradientSpec>,
    // As in `render_single_textbox`; each paragraph becomes a P or H1–H6
    // (a list paragraph an LI) in the Sect.
    mut tag: Option<(&mut Tags, usize, usize)>,
    hf_sub: Option<&super::table_layout::HfSubstitution>,
) {
    let mut cursor_y = start_y;
    let mut prev_space_after = 0.0f32;
    let mut lists = Lists::default();
    for (tp_idx, tp) in paragraphs.iter().enumerate() {
        // Collapse adjacent spacing: use max(prev_after, current_before) like body text
        let inter_gap = if tp_idx == 0 {
            tp.space_before
        } else {
            prev_space_after.max(tp.space_before)
        };
        // Stop rendering when content overflows the textbox bounds
        if let Some(bottom) = clip_bottom
            && cursor_y - inter_gap < bottom
        {
            break;
        }
        // Word keeps empty and picture-only paragraphs as empty P elements.
        // (Lbl, text element) as in the body: see `para_tags`.
        let tag_nodes = tag.as_mut().map(|(tags, _, sect)| {
            let item = tp.list_item.filter(|_| tp.outline_level.is_none());
            // Word tags headings in textboxes too (never in table cells).
            let style_name = tp
                .style_id
                .as_ref()
                .and_then(|id| ctx.style_id_to_name.get(id));
            let kind = super::para_tag_kind(tp, style_name);
            tags.para_nodes(&mut lists, *sect, item, !tp.list_label.is_empty(), kind)
        });
        let para_tag = tag_nodes.map(|(_, text)| text);
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

        if let Some(img) = textbox_para_block_image(tp) {
            if let Some(pdf_name) = ctx.textbox_image_names.get(&img.key()) {
                let img_x = content_x
                    + tp.indent_left
                    + x_offset
                    + align_offset(tp.alignment, (tp_align_w - img.display_width).max(0.0));
                let img_y = cursor_y - inter_gap - img.display_height - y_offset;
                // The paragraph's P stays empty and the picture follows it as
                // a Figure in the Sect, as in the body.
                let figure = tag.as_mut().filter(|_| !img.decorative);
                if let Some((tags, page, sect)) = figure {
                    let figure = tags.add_figure(*sect, img.alt.as_deref());
                    tags.begin(content, *page, figure);
                }
                super::smartart::render_image_with_clip(
                    content,
                    pdf_name,
                    img_x,
                    img_y,
                    img.display_width,
                    img.display_height,
                    img.clip_geometry.as_ref(),
                );
                if tag.is_some() && !img.decorative {
                    Tags::end(content);
                }
            }
            cursor_y -= inter_gap + img.display_height + img.layout_extra_height;
            prev_space_after = tp.space_after;
            continue;
        }

        // Header/footer page numbers may sit in a text box: evaluate the
        // field runs for this page before laying them out. The substituted
        // vector lives for the rest of the iteration.
        let hf_runs: Vec<crate::model::Run>;
        let runs: &[crate::model::Run] = match hf_sub {
            Some(sub) => {
                hf_runs = super::header_footer::substitute_hf_runs(
                    &tp.runs,
                    sub.page_num,
                    sub.total_pages,
                    sub.styleref_values,
                    sub.page_num_format,
                );
                &hf_runs
            }
            None => &tp.runs,
        };
        let inline_imgs = super::layout::run_image_map(runs, ctx.textbox_image_names);
        let mut tb_lines = build_lines(
            runs,
            ctx,
            tp_text_w,
            ctx.cjk(true, tp),
            &LineOpts {
                inline_images: Some(&inline_imgs),
                tab_stops: &tp.tab_stops,
                indent_left: tp.indent_left,
                indent_right: tp.indent_right,
                hanging: text_hanging,
                ..Default::default()
            },
        );
        if tb_lines.is_empty() {
            let (fs, lhr, _) = tallest_run_metrics(runs, ctx.fonts);
            let lh = resolve_line_h(tp_ls, fs, lhr);
            cursor_y -= inter_gap + lh;
            prev_space_after = tp.space_after;
            continue;
        }
        let (tb_fs, tb_lhr, tb_ar) = tallest_run_metrics(runs, ctx.fonts);
        let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_lhr);
        let tb_text_ascent = tb_fs * tb_ar.unwrap_or(0.75);
        let tb_ascent = super::layout::boxed_line_ascent(
            tp_ls, tb_line_h, tb_fs, tb_lhr, tb_ar, runs, ctx.fonts,
        )
        .unwrap_or(tb_text_ascent);
        let tb_baseline = cursor_y - inter_gap - tb_ascent - y_offset;
        // Each line is as tall as its own runs, as in the body: air_pollution's
        // 10.5pt dot-leader line before a break steps 18.1 at 1.5 lines, not
        // the 20.7 of the paragraph's 12pt text.
        if !matches!(tp_ls, crate::model::LineSpacing::Exact(_)) {
            super::layout::size_lines_by_own_runs(
                &mut tb_lines,
                ctx.fonts,
                tp_ls,
                tb_line_h,
                tb_text_ascent,
                tp.text_alignment == crate::model::TextAlignment::Top,
            );
            super::size_break_lines(&mut tb_lines, runs, tp_ls, tb_line_h, ctx.fonts);
            super::size_mark_line(&mut tb_lines, tp, runs, tp_ls, tb_lhr, ctx.fonts);
        }
        let tb_metrics = (
            tb_text_ascent,
            if inline_imgs.is_empty() {
                0.0
            } else {
                picture_line_bottom(runs, tp, ctx.fonts, tp_ls)
            },
        );
        let tp_text_x = content_x + tp.indent_left + x_offset;
        if let Some(c) = force_color {
            fill_rgb(content, c);
        }
        let draw_label = |content: &mut Content| {
            if render_labels {
                let label_x = content_x + tp.indent_left - tp.indent_hanging;
                render_list_label(content, tp, ctx.fonts, label_x, tb_baseline, tb_fs);
            }
        };
        match tag.as_mut().zip(tag_nodes) {
            Some(((tags, page, _), nodes)) => tags.begin_para(content, *page, nodes, draw_label),
            None => draw_label(content),
        }
        render_paragraph_lines(
            content,
            &tb_lines,
            &tp.alignment,
            tp_text_x,
            tp.indent_left,
            tp_align_w,
            tb_baseline,
            tb_line_h,
            tb_metrics,
            tb_lines.len(),
            0,
            links,
            // The first line's indent also moves centred and right-aligned text
            // (air_pollution's caption: firstLine 709 shifts it 17.6pt).
            text_hanging,
            ctx.fonts,
            None,
            gradient_specs,
            None,
            None,
            tag.as_mut()
                .zip(para_tag)
                .map(|((tags, page, _), p)| LinkTagger::new(tags, *page, p)),
        );
        if para_tag.is_some() {
            Tags::end(content);
        }
        cursor_y -= inter_gap + lines_height(&tb_lines, tb_line_h, tb_metrics);
        prev_space_after = tp.space_after;
    }
}
