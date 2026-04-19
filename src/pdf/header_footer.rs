use std::collections::HashMap;

use pdf_writer::Content;

use crate::model::{
    Alignment, Block, Document, FieldCode, HeaderFooter, Paragraph, Run, SectionProperties,
    VRelativeFrom, VerticalPosition, WrapType,
};

use super::layout::{
    TextLine, build_paragraph_lines, build_tabbed_line, is_text_empty, render_paragraph_lines,
    tallest_run_metrics,
};
use super::positioning::resolve_h_position;
use super::table;
use super::{RenderContext, resolve_line_h};

pub(super) fn substitute_hf_runs(
    runs: &[Run],
    page_num: usize,
    total_pages: usize,
    styleref_values: &HashMap<String, String>,
    page_num_format: Option<&str>,
) -> Vec<Run> {
    runs.iter()
        .map(|run| {
            let mut r = run.clone();
            if let Some(ref fc) = run.field_code {
                r.field_code = None;
                r.text = match fc {
                    FieldCode::Page => {
                        if let Some(fmt) = page_num_format {
                            crate::docx::numbering::format_number(page_num as u32, fmt)
                        } else {
                            page_num.to_string()
                        }
                    }
                    FieldCode::NumPages => total_pages.to_string(),
                    FieldCode::StyleRef(name) => {
                        styleref_values.get(name).cloned().unwrap_or_default()
                    }
                    FieldCode::PageRef(_) => run.text.clone(),
                };
            }
            r
        })
        .collect()
}

pub(super) fn compute_header_height(
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
) -> f32 {
    let text_width = sp.page_width - sp.margin_left - sp.margin_right;
    let mut height = 0.0f32;
    let mut prev_space_after = 0.0f32;
    // Track bottom of wrapping float zones (Square/Tight/Through): subsequent
    // paragraphs flow beside the image so their height is absorbed, not additive.
    let mut float_bottom_h = 0.0f32;
    for block in &hf.blocks {
        match block {
            Block::Paragraph(para) if para.frame_props.is_some() => {
                // Frame paragraphs are out-of-flow; skip height contribution
            }
            Block::Paragraph(para) => {
                height += prev_space_after.max(para.space_before);
                let (font_size, tallest_lhr, _) = tallest_run_metrics(&para.runs, ctx.fonts);
                let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                let max_img_h = para
                    .runs
                    .iter()
                    .filter_map(|r| r.inline_image.as_ref())
                    .map(|img| img.display_height + img.layout_extra_height)
                    .fold(0.0f32, f32::max);
                let mut content_h = max_img_h.max(line_h);

                for fi in &para.floating_images {
                    if matches!(fi.wrap_type, WrapType::None) {
                        continue;
                    }
                    let fi_h = match fi.v_position {
                        VerticalPosition::Offset(o) => o.max(0.0) + fi.image.display_height,
                        _ => fi.image.display_height,
                    };
                    if matches!(
                        fi.wrap_type,
                        WrapType::Square | WrapType::Tight | WrapType::Through
                    ) && fi.image.display_width < text_width * 0.5
                    {
                        // Narrow wrapping image: text flows beside it. Track
                        // its bottom separately instead of inflating content_h.
                        float_bottom_h = float_bottom_h.max(height + fi_h);
                    } else {
                        // TopAndBottom or wide wrapping image
                        content_h = content_h.max(fi_h);
                    }
                }

                for tb in &para.textboxes {
                    if matches!(tb.wrap_type, WrapType::TopAndBottom) {
                        let tb_bottom = tb.v_offset_pt + tb.height_pt + tb.dist_bottom;
                        match tb.v_relative_from {
                            VRelativeFrom::Paragraph => {
                                content_h = content_h.max(tb_bottom);
                            }
                            VRelativeFrom::Page if is_header => {
                                // tb_bottom is absolute from page top; convert to
                                // distance relative to the current position in the header
                                let current_pos = sp.header_margin + height;
                                let contribution = (tb_bottom - current_pos).max(0.0);
                                content_h = content_h.max(contribution);
                            }
                            _ => {
                                content_h += tb_bottom;
                            }
                        }
                    }
                }

                // Each w:br (line break) in the paragraph creates an additional line.
                let br_count = para.runs.iter().filter(|r| r.is_line_break).count();
                content_h += br_count as f32 * line_h;

                height += content_h;
                prev_space_after = para.space_after;
            }
            Block::Table(table) => {
                height += table::compute_hf_table_height(table, ctx);
                prev_space_after = 0.0;
            }
        }
    }
    (height + prev_space_after).max(float_bottom_h)
}

pub(super) fn effective_slot_top(
    sp: &SectionProperties,
    is_first: bool,
    ctx: &RenderContext,
) -> f32 {
    let header = select_hf(
        is_first,
        sp.different_first_page,
        &sp.header_first,
        &sp.header_default,
    );
    let base = sp.page_height - sp.margin_top;
    match header {
        Some(hf) => {
            base.min(sp.page_height - sp.header_margin - compute_header_height(hf, ctx, sp, true))
        }
        None => base,
    }
}

pub(super) fn compute_effective_margin_bottom(
    sp: &SectionProperties,
    is_first: bool,
    ctx: &RenderContext,
) -> f32 {
    let footer = select_hf(
        is_first,
        sp.different_first_page,
        &sp.footer_first,
        &sp.footer_default,
    );
    let base = sp.margin_bottom;
    match footer {
        Some(hf) => {
            base.max(sp.footer_margin + compute_header_height(hf, ctx, sp, false))
        }
        None => base,
    }
}

fn select_hf<'a>(
    is_first: bool,
    different_first_page: bool,
    first: &'a Option<HeaderFooter>,
    default: &'a Option<HeaderFooter>,
) -> Option<&'a HeaderFooter> {
    if is_first && different_first_page {
        first.as_ref()
    } else {
        default.as_ref()
    }
}

pub(super) fn hf_paragraphs(hf: &HeaderFooter) -> Vec<&Paragraph> {
    hf.blocks
        .iter()
        .flat_map(|block| match block {
            Block::Paragraph(p) => vec![p],
            Block::Table(t) => t
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .flat_map(|cell| cell.all_paragraphs())
                .collect(),
        })
        .collect()
}

fn resolve_tb_y_top(
    v_relative_from: VRelativeFrom,
    v_offset_pt: f32,
    sp: &SectionProperties,
    slot_top: f32,
) -> f32 {
    match v_relative_from {
        VRelativeFrom::Page => sp.page_height - v_offset_pt,
        VRelativeFrom::Margin | VRelativeFrom::TopMargin => {
            sp.page_height - sp.margin_top - v_offset_pt
        }
        VRelativeFrom::Paragraph => slot_top - v_offset_pt,
    }
}

fn build_lines(
    runs: &[Run],
    fonts: &HashMap<String, crate::fonts::FontEntry>,
    tab_stops: &[crate::model::TabStop],
    text_width: f32,
    inline_images: &HashMap<usize, String>,
    default_tab_stop: f32,
    indent_left: f32,
    text_hanging: f32,
) -> Vec<TextLine> {
    build_lines_with_float(runs, fonts, tab_stops, text_width, inline_images, default_tab_stop, indent_left, text_hanging, None)
}

fn build_lines_with_float(
    runs: &[Run],
    fonts: &HashMap<String, crate::fonts::FontEntry>,
    tab_stops: &[crate::model::TabStop],
    text_width: f32,
    inline_images: &HashMap<usize, String>,
    default_tab_stop: f32,
    indent_left: f32,
    text_hanging: f32,
    per_line_widths: Option<&[f32]>,
) -> Vec<TextLine> {
    let empty_fx: HashMap<usize, super::images::EffectXObjs> = HashMap::new();
    let has_tabs = runs.iter().any(|r| r.is_tab);
    if has_tabs {
        build_tabbed_line(runs, fonts, tab_stops, indent_left, text_width, text_hanging, inline_images, &empty_fx, default_tab_stop)
    } else {
        build_paragraph_lines(runs, fonts, text_width, text_hanging, inline_images, &empty_fx, None, per_line_widths, None, true)
    }
}

/// Per-page context for header/footer rendering: field substitution values
/// and image name mappings resolved by the caller.
pub(super) struct HfPageContext<'a> {
    pub(super) page_num: usize,
    pub(super) total_pages: usize,
    pub(super) para_image_names: &'a HashMap<usize, String>,
    pub(super) inline_image_names: &'a HashMap<(usize, usize), String>,
    pub(super) floating_image_names: &'a HashMap<(usize, usize), String>,
    pub(super) effect_para_names: &'a HashMap<usize, super::images::EffectXObjs>,
    #[allow(dead_code)]
    pub(super) effect_floating_names: &'a HashMap<(usize, usize), super::images::EffectXObjs>,
    pub(super) styleref_values: &'a HashMap<String, String>,
    pub(super) page_num_format: Option<&'a str>,
}

pub(super) fn render_header_footer(
    content: &mut Content,
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
    pc: &HfPageContext,
    gradient_specs: &mut Vec<super::GradientSpec>,
) {
    let page_num = pc.page_num;
    let total_pages = pc.total_pages;
    let para_image_names = pc.para_image_names;
    let inline_image_names = pc.inline_image_names;
    let floating_image_names = pc.floating_image_names;
    let styleref_values = pc.styleref_values;
    let page_num_format = pc.page_num_format;
    let text_width = sp.page_width - sp.margin_left - sp.margin_right;
    let mut cursor_y = if is_header {
        sp.page_height - sp.header_margin
    } else {
        sp.footer_margin + compute_header_height(hf, ctx, sp, false)
    };

    let mut pi = 0usize;
    let mut prev_space_after = 0.0f32;
    // Float zone: (fi_x, fi_y_top, fi_y_bottom, obj_right, dist_left, dist_right)
    let mut hdr_fz: Option<(f32, f32, f32, f32, f32, f32)> = None;
    for block in &hf.blocks {
        match block {
            Block::Table(table) => {
                table::render_header_footer_table(
                    table,
                    sp,
                    ctx,
                    content,
                    &mut cursor_y,
                    page_num,
                    total_pages,
                    styleref_values,
                    page_num_format,
                );
                prev_space_after = 0.0;
            }
            Block::Paragraph(para) if para.frame_props.is_some() => {
                let fp = para.frame_props.as_ref().unwrap();
                let substituted_runs = substitute_hf_runs(
                    &para.runs, page_num, total_pages, styleref_values,
                    page_num_format,
                );
                let (font_size, _, tallest_ar) =
                    tallest_run_metrics(&substituted_runs, ctx.fonts);
                let ascender_ratio = tallest_ar.unwrap_or(0.75);

                let empty_inline_imgs: HashMap<usize, String> = HashMap::new();
                let lines = build_lines(
                    &substituted_runs, ctx.fonts, &para.tab_stops,
                    text_width, &empty_inline_imgs, ctx.default_tab_stop,
                    0.0, 0.0,
                );
                let content_width = lines.iter()
                    .map(|l| l.total_width)
                    .fold(0.0f32, f32::max);

                let frame_x = resolve_h_position(
                    fp.h_relative_from, &fp.h_position, content_width,
                    sp, sp.margin_left, text_width, text_width,
                );
                let frame_baseline = cursor_y - font_size * ascender_ratio;

                render_paragraph_lines(
                    content, &lines, &Alignment::Left,
                    frame_x, content_width, frame_baseline,
                    font_size, lines.len(), 0,
                    &mut Vec::new(), 0.0, ctx.fonts, None,
                );
                // Frame paragraphs are out-of-flow: do not advance cursor_y
                pi += 1;
            }
            Block::Paragraph(para) => {
                let has_para_image = para.image.is_some();
                let has_field_code = para.runs.iter().any(|r| r.field_code.is_some());
                let text_empty = !has_field_code && is_text_empty(&para.runs);

                cursor_y -= prev_space_after.max(para.space_before);

                let substituted_runs =
                    substitute_hf_runs(&para.runs, page_num, total_pages, styleref_values, page_num_format);

                let (font_size, tallest_lhr, tallest_ar) =
                    tallest_run_metrics(&substituted_runs, ctx.fonts);
                let ascender_ratio = tallest_ar.unwrap_or(0.75);
                let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);

                let baseline_y = cursor_y - font_size * ascender_ratio;
                let slot_top = cursor_y;

                // Render textboxes
                for tb in &para.textboxes {
                    let tb_x = super::resolve_h_position(
                        tb.h_relative_from,
                        &tb.h_position,
                        tb.width_pt,
                        sp,
                        sp.margin_left,
                        text_width,
                        text_width,
                    );
                    let tb_y_top =
                        resolve_tb_y_top(tb.v_relative_from, tb.v_offset_pt, sp, slot_top);

                    if let Some(ref fill) = tb.fill {
                        super::render_shape_fill(
                            content,
                            fill,
                            tb_x,
                            tb_y_top - tb.height_pt,
                            tb.width_pt,
                            tb.height_pt,
                            &tb.shape_type,
                            gradient_specs,
                        );
                    }

                    let content_x = tb_x + tb.margin_left;
                    let content_w = (tb.width_pt - tb.margin_left - tb.margin_right).max(0.0);
                    let mut tb_cursor = tb_y_top - tb.margin_top;
                    let empty_inline_imgs: HashMap<usize, String> = HashMap::new();
                    for tp in &tb.paragraphs {
                        let tp_ls = tp.line_spacing.unwrap_or(ctx.doc_line_spacing);
                        let tp_text_w = (content_w - tp.indent_left - tp.indent_right).max(1.0);
                        let tp_hanging = if !tp.list_label.is_empty() {
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
                        let tb_lines = build_lines(
                            &tp.runs,
                            ctx.fonts,
                            &tp.tab_stops,
                            tp_text_w,
                            &empty_inline_imgs,
                            ctx.default_tab_stop,
                            tp.indent_left,
                            tp_hanging,
                        );
                        if tb_lines.is_empty() {
                            let (fs, _, _) = tallest_run_metrics(&tp.runs, ctx.fonts);
                            let lh = resolve_line_h(tp_ls, fs, None);
                            tb_cursor -= tp.space_before + lh + tp.space_after;
                            continue;
                        }
                        let (tb_fs, _, tb_ar) = tallest_run_metrics(&tp.runs, ctx.fonts);
                        let tb_ascender = tb_ar.unwrap_or(0.75);
                        let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_ar);
                        let tb_baseline = tb_cursor - tp.space_before - tb_fs * tb_ascender;
                        super::render_list_label(
                            content,
                            tp,
                            ctx.fonts,
                            content_x + tp.indent_left - tp.indent_hanging,
                            tb_baseline,
                            tb_fs,
                        );
                        render_paragraph_lines(
                            content,
                            &tb_lines,
                            &tp.alignment,
                            content_x + tp.indent_left,
                            tp_text_w,
                            tb_baseline,
                            tb_line_h,
                            tb_lines.len(),
                            0,
                            &mut Vec::new(),
                            0.0,
                            ctx.fonts,
                            None,
                        );
                        tb_cursor -=
                            tp.space_before + (tb_lines.len() as f32) * tb_line_h + tp.space_after;
                    }
                }

                // Render floating images and register float zones
                for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                    if let Some(pdf_name) = floating_image_names.get(&(pi, fi_idx)) {
                        let img = &fi.image;
                        let fi_x = super::resolve_h_position(
                            fi.h_relative_from,
                            &fi.h_position,
                            img.display_width,
                            sp,
                            sp.margin_left,
                            text_width,
                            text_width,
                        );
                        let fi_y_top = super::resolve_fi_y_top(fi, sp, slot_top);
                        super::smartart::render_image_with_clip(
                            content,
                            pdf_name,
                            fi_x,
                            fi_y_top - img.display_height,
                            img.display_width,
                            img.display_height,
                            img.clip_geometry.as_ref(),
                        );
                        // Register float zone for wrapping images
                        if matches!(
                            fi.wrap_type,
                            WrapType::Square | WrapType::Tight | WrapType::Through
                        ) {
                            hdr_fz = Some((
                                fi_x,
                                fi_y_top,
                                fi_y_top - img.display_height,
                                fi_x + img.display_width,
                                fi.dist_left,
                                fi.dist_right,
                            ));
                        }
                    }
                }

                if (has_para_image || text_empty) && para.content_height > 0.0 {
                    if let Some(pdf_name) = para_image_names.get(&pi) {
                        let img = para.image.as_ref().unwrap();
                        let y_bottom = baseline_y + font_size * ascender_ratio - img.display_height;
                        let x = sp.margin_left
                            + match para.alignment {
                                Alignment::Center => {
                                    (text_width - img.display_width).max(0.0) / 2.0
                                }
                                Alignment::Right => (text_width - img.display_width).max(0.0),
                                _ => 0.0,
                            };
                        let hf_fx = pc.effect_para_names.get(&pi);
                        if let Some(ref shadow) = img.shadow {
                            super::color::draw_image_shadow(
                                content, shadow, x, y_bottom,
                                img.display_width, img.display_height,
                                hf_fx.and_then(|fx| fx.shadow.as_deref()),
                            );
                        }
                        if let Some(ref glow) = img.glow {
                            super::color::draw_image_glow(
                                content, glow, x, y_bottom,
                                img.display_width, img.display_height,
                                hf_fx.and_then(|fx| fx.glow.as_deref()),
                            );
                        }
                        super::smartart::render_image_with_clip(
                            content,
                            pdf_name,
                            x,
                            y_bottom,
                            img.display_width,
                            img.display_height,
                            img.clip_geometry.as_ref(),
                        );
                        if let Some(sc) = img.stroke_color {
                            super::smartart::stroke_image_border(
                                content, x, y_bottom,
                                img.display_width, img.display_height,
                                sc, img.stroke_width,
                                img.clip_geometry.as_ref(),
                            );
                        }
                    }
                    cursor_y -= line_h;
                    prev_space_after = para.space_after;
                    pi += 1;
                    continue;
                }

                // Before text_empty skip so empty paragraphs with borders still render
                {
                    let bdr = &para.borders;
                    let box_left = sp.margin_left;
                    let box_right = sp.margin_left + text_width;
                    let box_top = cursor_y;
                    let box_bottom = cursor_y - line_h;
                    let draw_h_border =
                        |content: &mut Content, b: &crate::model::ParagraphBorder, y: f32| {
                            content.save_state();
                            content.set_line_width(b.width_pt);
                            super::color::stroke_rgb(content, b.color);
                            content.move_to(box_left, y);
                            content.line_to(box_right, y);
                            content.stroke();
                            content.restore_state();
                        };
                    if let Some(b) = &bdr.top {
                        draw_h_border(content, b, box_top);
                    }
                    if let Some(b) = &bdr.bottom {
                        draw_h_border(content, b, box_bottom);
                    }
                }

                if text_empty {
                    let mut advance = line_h;
                    // TopAndBottom textboxes push content below them
                    for tb in &para.textboxes {
                        if matches!(tb.wrap_type, WrapType::TopAndBottom) {
                            let tb_bottom_y = match tb.v_relative_from {
                                VRelativeFrom::Page => {
                                    sp.page_height
                                        - (tb.v_offset_pt + tb.height_pt + tb.dist_bottom)
                                }
                                _ => cursor_y - tb.v_offset_pt - tb.height_pt - tb.dist_bottom,
                            };
                            let needed = (cursor_y - tb_bottom_y).max(0.0);
                            advance = advance.max(needed);
                        }
                    }
                    cursor_y -= advance;
                    prev_space_after = para.space_after;
                    pi += 1;
                    continue;
                }

                let block_inline_images: HashMap<usize, String> = inline_image_names
                    .iter()
                    .filter(|((pi2, _), _)| *pi2 == pi)
                    .map(|((_, ri), name)| (*ri, name.clone()))
                    .collect();

                let mut para_text_x = sp.margin_left + para.indent_left;
                let mut para_text_width =
                    (text_width - para.indent_left - para.indent_right).max(1.0);
                let text_hanging = if !para.list_label.is_empty() {
                    if let Some(nts) = para.num_level_tab_stop {
                        if nts < para.indent_left && (para.indent_left - para.indent_hanging).abs() < 0.5 {
                            (para.indent_left - nts).max(0.0)
                        } else if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
                            -para.indent_first_line
                        } else {
                            0.0
                        }
                    } else if para.indent_first_line > 0.0 && para.indent_hanging == 0.0 {
                        -para.indent_first_line
                    } else {
                        0.0
                    }
                } else if para.indent_hanging > 0.0 {
                    para.indent_hanging
                } else {
                    -para.indent_first_line
                };

                // Narrow text width when a wrapping floating image overlaps this paragraph
                // Check both same-paragraph floats and cross-paragraph float zone
                let mut hdr_line_geom: Option<Vec<(f32, f32)>> = None;
                let fz_params: Option<(f32, f32, f32, f32, f32, f32)> =
                    para.floating_images
                        .iter()
                        .find(|fi| {
                            matches!(
                                fi.wrap_type,
                                WrapType::Square | WrapType::Tight | WrapType::Through
                            )
                        })
                        .map(|fi| {
                            let img = &fi.image;
                            let fi_x = super::resolve_h_position(
                                fi.h_relative_from,
                                &fi.h_position,
                                img.display_width,
                                sp,
                                sp.margin_left,
                                text_width,
                                text_width,
                            );
                            let fi_y_top = super::resolve_fi_y_top(fi, sp, slot_top);
                            (
                                fi_x,
                                fi_y_top,
                                fi_y_top - img.display_height,
                                fi_x + img.display_width,
                                fi.dist_left,
                                fi.dist_right,
                            )
                        })
                        .or(hdr_fz);

                if let Some((fi_x, fi_y_top, fi_y_bottom, obj_right, dist_left, dist_right)) =
                    fz_params
                {
                    let col_x = sp.margin_left;
                    let col_right = sp.margin_left + text_width;

                    // Check if paragraph overlaps the float vertically
                    if slot_top > fi_y_bottom && baseline_y < fi_y_top {
                        let space_right =
                            col_right - (obj_right + dist_right);
                        let space_left = (fi_x - dist_left) - col_x;

                        if space_right >= space_left && space_right >= 36.0 {
                            let new_left = obj_right + dist_right;
                            para_text_width =
                                (col_right - new_left - para.indent_right).max(1.0);
                            para_text_x = new_left + para.indent_left;
                        } else if space_left >= 36.0 {
                            let avail_right = fi_x - dist_left;
                            para_text_width = (avail_right - col_x
                                - para.indent_left
                                - para.indent_right)
                                .max(1.0);
                        }

                        // Build per-line geometry for multi-line paragraphs that may
                        // span above and through the float zone
                        let ascender_ratio_e = tallest_ar.unwrap_or(0.75);
                        let full_w =
                            (text_width - para.indent_left - para.indent_right).max(1.0);
                        let max_lines = ((slot_top - fi_y_bottom) / line_h)
                            .ceil() as usize
                            + 10;
                        let max_lines = max_lines.max(20);
                        let mut geom = Vec::with_capacity(max_lines);
                        for i in 0..max_lines {
                            let y = slot_top
                                - font_size * ascender_ratio_e
                                - i as f32 * line_h;
                            if y <= fi_y_top && y > fi_y_bottom {
                                let sr = col_right - (obj_right + dist_right);
                                let sl = (fi_x - dist_left) - col_x;
                                if sr >= sl && sr >= 36.0 {
                                    let nl = obj_right + dist_right;
                                    let w = (col_right - nl - para.indent_right)
                                        .max(1.0);
                                    geom.push((nl + para.indent_left, w));
                                } else if sl >= 36.0 {
                                    let ar = fi_x - dist_left;
                                    let w = (ar - col_x
                                        - para.indent_left
                                        - para.indent_right)
                                        .max(1.0);
                                    geom.push((col_x + para.indent_left, w));
                                } else {
                                    geom.push((col_x + para.indent_left, full_w));
                                }
                            } else {
                                geom.push((col_x + para.indent_left, full_w));
                            }
                        }
                        hdr_line_geom = Some(geom);
                    }
                }

                let per_line_widths: Option<Vec<f32>> =
                    hdr_line_geom.as_ref().map(|g| g.iter().map(|&(_, w)| w).collect());

                let lines = build_lines_with_float(
                    &substituted_runs,
                    ctx.fonts,
                    &para.tab_stops,
                    para_text_width,
                    &block_inline_images,
                    ctx.default_tab_stop,
                    para.indent_left,
                    text_hanging,
                    per_line_widths.as_deref(),
                );

                render_paragraph_lines(
                    content,
                    &lines,
                    &para.alignment,
                    para_text_x,
                    para_text_width,
                    baseline_y,
                    line_h,
                    lines.len(),
                    0,
                    &mut Vec::new(),
                    text_hanging,
                    ctx.fonts,
                    hdr_line_geom.as_deref(),
                );

                let max_img_h = lines
                    .iter()
                    .flat_map(|l| l.chunks.iter())
                    .map(|c| c.inline_image_height)
                    .fold(0.0f32, f32::max);
                cursor_y -= lines.len().max(1) as f32 * max_img_h.max(line_h);
                prev_space_after = para.space_after;
                pi += 1;
            }
        }
    }
}

/// Resolve which header to use for a given page, walking sections backward
/// for inheritance. Returns `(header_data, hf_type_id, section_index)`.
pub(super) fn resolve_header_for_page<'a>(
    doc: &'a Document,
    section_idx: usize,
    is_first_page: bool,
    page_num: usize,
) -> (Option<&'a HeaderFooter>, u8, usize) {
    for idx in (0..=section_idx).rev() {
        let s = &doc.sections[idx].properties;
        let (h, t) = if idx == section_idx {
            if is_first_page && s.different_first_page {
                (s.header_first.as_ref(), 1u8)
            } else if doc.even_and_odd_headers && page_num % 2 == 0 && s.header_even.is_some() {
                (s.header_even.as_ref(), 4u8)
            } else {
                (s.header_default.as_ref(), 0u8)
            }
        } else {
            (s.header_default.as_ref(), 0u8)
        };
        if h.is_some() {
            return (h, t, idx);
        }
    }
    (None, 0, section_idx)
}

/// Resolve which footer to use for a given page, walking sections backward
/// for inheritance. Returns `(footer_data, hf_type_id, section_index)`.
pub(super) fn resolve_footer_for_page<'a>(
    doc: &'a Document,
    section_idx: usize,
    is_first_page: bool,
    page_num: usize,
) -> (Option<&'a HeaderFooter>, u8, usize) {
    for idx in (0..=section_idx).rev() {
        let s = &doc.sections[idx].properties;
        let (f, t) = if idx == section_idx {
            if is_first_page && s.different_first_page {
                (s.footer_first.as_ref(), 3u8)
            } else if doc.even_and_odd_headers && page_num % 2 == 0 && s.footer_even.is_some() {
                (s.footer_even.as_ref(), 5u8)
            } else {
                (s.footer_default.as_ref(), 2u8)
            }
        } else {
            (s.footer_default.as_ref(), 2u8)
        };
        if f.is_some() {
            return (f, t, idx);
        }
    }
    (None, 2, section_idx)
}
