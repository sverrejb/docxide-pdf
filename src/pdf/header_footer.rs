use std::collections::HashMap;

use pdf_writer::{Content, Name};

use crate::model::{
    Alignment, Block, FieldCode, HeaderFooter, Paragraph, Run, SectionProperties, VRelativeFrom,
};

use super::layout::{
    TextLine, build_paragraph_lines, build_tabbed_line, is_text_empty, render_paragraph_lines,
    tallest_run_metrics,
};
use super::table;
use super::{RenderContext, resolve_line_h};

pub(super) fn substitute_hf_runs(
    runs: &[Run],
    page_num: usize,
    total_pages: usize,
    styleref_values: &HashMap<String, String>,
) -> Vec<Run> {
    runs.iter()
        .map(|run| {
            let mut r = run.clone();
            if let Some(ref fc) = run.field_code {
                r.field_code = None;
                r.text = match fc {
                    FieldCode::Page => page_num.to_string(),
                    FieldCode::NumPages => total_pages.to_string(),
                    FieldCode::StyleRef(name) => {
                        styleref_values.get(name).cloned().unwrap_or_default()
                    }
                };
            }
            r
        })
        .collect()
}

pub(super) fn compute_header_height(hf: &HeaderFooter, ctx: &RenderContext) -> f32 {
    let mut height = 0.0f32;
    let mut prev_space_after = 0.0f32;
    for block in &hf.blocks {
        match block {
            Block::Paragraph(para) => {
                height += prev_space_after.max(para.space_before);
                let (font_size, tallest_lhr, _) = tallest_run_metrics(&para.runs, ctx.fonts);
                let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                let max_img_h = para
                    .runs
                    .iter()
                    .filter_map(|r| r.inline_image.as_ref())
                    .map(|img| img.display_height)
                    .fold(0.0f32, f32::max);
                height += max_img_h.max(line_h);
                prev_space_after = para.space_after;
            }
            Block::Table(table) => {
                height += table::compute_hf_table_height(table, ctx);
                prev_space_after = 0.0;
            }
        }
    }
    // Include trailing space_after: the header's last paragraph's space_after
    // represents the extent of the header region and affects where body text starts.
    height + prev_space_after
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
        Some(hf) => base.min(sp.page_height - sp.header_margin - compute_header_height(hf, ctx)),
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
        Some(hf) => base.max(sp.footer_margin + compute_header_height(hf, ctx)),
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
                .flat_map(|cell| cell.paragraphs.iter())
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

fn emit_image_xobject(
    content: &mut Content,
    pdf_name: &str,
    x: f32,
    y_bottom: f32,
    w: f32,
    h: f32,
) {
    content.save_state();
    content.transform([w, 0.0, 0.0, h, x, y_bottom]);
    content.x_object(Name(pdf_name.as_bytes()));
    content.restore_state();
}

fn build_lines(
    runs: &[Run],
    fonts: &HashMap<String, crate::fonts::FontEntry>,
    tab_stops: &[crate::model::TabStop],
    text_width: f32,
    inline_images: &HashMap<usize, String>,
) -> Vec<TextLine> {
    let has_tabs = runs.iter().any(|r| r.is_tab);
    if has_tabs {
        build_tabbed_line(runs, fonts, tab_stops, 0.0, text_width, 0.0, inline_images)
    } else {
        build_paragraph_lines(runs, fonts, text_width, 0.0, inline_images)
    }
}

pub(super) fn render_header_footer(
    content: &mut Content,
    hf: &HeaderFooter,
    ctx: &RenderContext,
    sp: &SectionProperties,
    is_header: bool,
    page_num: usize,
    total_pages: usize,
    para_image_names: &HashMap<usize, String>,
    inline_image_names: &HashMap<(usize, usize), String>,
    floating_image_names: &HashMap<(usize, usize), String>,
    styleref_values: &HashMap<String, String>,
    gradient_specs: &mut Vec<super::GradientSpec>,
) {
    let text_width = sp.page_width - sp.margin_left - sp.margin_right;
    let mut cursor_y = if is_header {
        sp.page_height - sp.header_margin
    } else {
        sp.footer_margin + compute_header_height(hf, ctx)
    };

    let mut pi = 0usize;
    let mut prev_space_after = 0.0f32;
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
                );
                prev_space_after = 0.0;
            }
            Block::Paragraph(para) => {
                let has_para_image = para.image.is_some();
                let has_field_code = para.runs.iter().any(|r| r.field_code.is_some());
                let text_empty = !has_field_code && is_text_empty(&para.runs);

                cursor_y -= prev_space_after.max(para.space_before);

                let substituted_runs =
                    substitute_hf_runs(&para.runs, page_num, total_pages, styleref_values);

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
                        let tb_lines = build_lines(
                            &tp.runs,
                            ctx.fonts,
                            &tp.tab_stops,
                            content_w,
                            &empty_inline_imgs,
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
                            content_x,
                            content_w,
                            tb_baseline,
                            tb_line_h,
                            tb_lines.len(),
                            0,
                            &mut Vec::new(),
                            0.0,
                            ctx.fonts,
                        );
                        tb_cursor -=
                            tp.space_before + (tb_lines.len() as f32) * tb_line_h + tp.space_after;
                    }
                }

                // Render floating images
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
                        emit_image_xobject(
                            content,
                            pdf_name,
                            fi_x,
                            fi_y_top - img.display_height,
                            img.display_width,
                            img.display_height,
                        );
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
                        emit_image_xobject(
                            content,
                            pdf_name,
                            x,
                            y_bottom,
                            img.display_width,
                            img.display_height,
                        );
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
                            let [r, g, b_c] = b.color;
                            content.save_state();
                            content.set_line_width(b.width_pt);
                            content.set_stroke_rgb(
                                r as f32 / 255.0,
                                g as f32 / 255.0,
                                b_c as f32 / 255.0,
                            );
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
                    cursor_y -= line_h;
                    prev_space_after = para.space_after;
                    pi += 1;
                    continue;
                }

                let block_inline_images: HashMap<usize, String> = inline_image_names
                    .iter()
                    .filter(|((pi2, _), _)| *pi2 == pi)
                    .map(|((_, ri), name)| (*ri, name.clone()))
                    .collect();

                let lines = build_lines(
                    &substituted_runs,
                    ctx.fonts,
                    &para.tab_stops,
                    text_width,
                    &block_inline_images,
                );

                render_paragraph_lines(
                    content,
                    &lines,
                    &para.alignment,
                    sp.margin_left,
                    text_width,
                    baseline_y,
                    line_h,
                    lines.len(),
                    0,
                    &mut Vec::new(),
                    0.0,
                    ctx.fonts,
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
