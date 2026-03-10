use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::FontEntry;
use crate::model::{
    Alignment, Block, FieldCode, HeaderFooter, HorizontalPosition, LineSpacing, Paragraph, Run,
    SectionProperties,
};

use super::layout::{
    build_paragraph_lines, build_tabbed_line, is_text_empty, render_paragraph_lines,
    tallest_run_metrics,
};
use super::table;
use super::{label_for_paragraph, resolve_line_h};

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

pub(super) fn compute_header_height(
    hf: &HeaderFooter,
    seen_fonts: &HashMap<String, FontEntry>,
    doc_line_spacing: LineSpacing,
) -> f32 {
    let mut height = 0.0f32;
    let mut prev_space_after = 0.0f32;
    for block in &hf.blocks {
        match block {
            Block::Paragraph(para) => {
                let inter_gap = f32::max(prev_space_after, para.space_before);
                height += inter_gap;
                let (font_size, tallest_lhr, _) = tallest_run_metrics(&para.runs, seen_fonts);
                let effective_ls = para.line_spacing.unwrap_or(doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);
                let max_img_h = para
                    .runs
                    .iter()
                    .filter_map(|r| r.inline_image.as_ref())
                    .map(|img| img.display_height)
                    .fold(0.0f32, f32::max);
                height += if max_img_h > line_h {
                    max_img_h
                } else {
                    line_h
                };
                prev_space_after = para.space_after;
            }
            Block::Table(table) => {
                let row_layouts =
                    table::compute_hf_table_height(table, doc_line_spacing, seen_fonts);
                height += row_layouts;
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
    seen_fonts: &HashMap<String, FontEntry>,
    doc_line_spacing: LineSpacing,
) -> f32 {
    let header = if is_first && sp.different_first_page {
        sp.header_first.as_ref()
    } else {
        sp.header_default.as_ref()
    };
    let base = sp.page_height - sp.margin_top;
    if let Some(hf) = header {
        let hdr_h = compute_header_height(hf, seen_fonts, doc_line_spacing);
        let hdr_bottom = sp.page_height - sp.header_margin - hdr_h;
        base.min(hdr_bottom)
    } else {
        base
    }
}

pub(super) fn compute_effective_margin_bottom(
    sp: &SectionProperties,
    is_first: bool,
    seen_fonts: &HashMap<String, FontEntry>,
    doc_line_spacing: LineSpacing,
) -> f32 {
    let footer = if is_first && sp.different_first_page {
        sp.footer_first.as_ref()
    } else {
        sp.footer_default.as_ref()
    };
    let base = sp.margin_bottom;
    if let Some(hf) = footer {
        let ftr_h = compute_header_height(hf, seen_fonts, doc_line_spacing);
        let ftr_top = sp.footer_margin + ftr_h;
        base.max(ftr_top)
    } else {
        base
    }
}

pub(super) fn hf_paragraphs(hf: &HeaderFooter) -> Vec<&Paragraph> {
    let mut out = Vec::new();
    for block in &hf.blocks {
        match block {
            Block::Paragraph(p) => out.push(p),
            Block::Table(t) => {
                for row in &t.rows {
                    for cell in &row.cells {
                        for p in &cell.paragraphs {
                            out.push(p);
                        }
                    }
                }
            }
        }
    }
    out
}

pub(super) fn render_header_footer(
    content: &mut Content,
    hf: &HeaderFooter,
    seen_fonts: &HashMap<String, FontEntry>,
    sp: &SectionProperties,
    doc_line_spacing: LineSpacing,
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
        sp.footer_margin + compute_header_height(hf, seen_fonts, doc_line_spacing)
    };

    let mut pi = 0usize;
    let mut prev_space_after = 0.0f32;
    for block in &hf.blocks {
        match block {
            Block::Table(table) => {
                table::render_header_footer_table(
                    table,
                    sp,
                    doc_line_spacing,
                    seen_fonts,
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

                let inter_gap = f32::max(prev_space_after, para.space_before);
                cursor_y -= inter_gap;

                let substituted_runs =
                    substitute_hf_runs(&para.runs, page_num, total_pages, styleref_values);

                let (font_size, tallest_lhr, tallest_ar) =
                    tallest_run_metrics(&substituted_runs, seen_fonts);
                let ascender_ratio = tallest_ar.unwrap_or(0.75);
                let effective_ls = para.line_spacing.unwrap_or(doc_line_spacing);
                let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);

                let baseline_y = cursor_y - font_size * ascender_ratio;
                let slot_top = cursor_y;

                // Render textboxes
                for tb in &para.textboxes {
                    let tb_x = match tb.h_relative_from {
                        "page" => match tb.h_position {
                            HorizontalPosition::AlignCenter => (sp.page_width - tb.width_pt) / 2.0,
                            HorizontalPosition::AlignRight => sp.page_width - tb.width_pt,
                            HorizontalPosition::AlignLeft => 0.0,
                            HorizontalPosition::Offset(o) => o,
                        },
                        "column" => match tb.h_position {
                            HorizontalPosition::AlignCenter => {
                                sp.margin_left + (text_width - tb.width_pt) / 2.0
                            }
                            HorizontalPosition::AlignRight => {
                                sp.margin_left + text_width - tb.width_pt
                            }
                            HorizontalPosition::AlignLeft => sp.margin_left,
                            HorizontalPosition::Offset(o) => sp.margin_left + o,
                        },
                        "margin" | _ => match tb.h_position {
                            HorizontalPosition::AlignCenter => {
                                sp.margin_left + (text_width - tb.width_pt) / 2.0
                            }
                            HorizontalPosition::AlignRight => {
                                sp.margin_left + text_width - tb.width_pt
                            }
                            HorizontalPosition::AlignLeft => sp.margin_left,
                            HorizontalPosition::Offset(o) => sp.margin_left + o,
                        },
                    };
                    let tb_y_top = match tb.v_relative_from {
                        "page" => sp.page_height - tb.v_offset_pt,
                        "margin" | "topMargin" => sp.page_height - sp.margin_top - tb.v_offset_pt,
                        _ => slot_top - tb.v_offset_pt,
                    };

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
                        let tp_ls = tp.line_spacing.unwrap_or(doc_line_spacing);
                        let has_tabs = tp.runs.iter().any(|r| r.is_tab);
                        let tb_lines = if has_tabs {
                            build_tabbed_line(
                                &tp.runs,
                                seen_fonts,
                                &tp.tab_stops,
                                0.0,
                                content_w,
                                0.0,
                                &empty_inline_imgs,
                            )
                        } else {
                            build_paragraph_lines(
                                &tp.runs,
                                seen_fonts,
                                content_w,
                                0.0,
                                &empty_inline_imgs,
                            )
                        };
                        if tb_lines.is_empty() {
                            let (fs, _, _) = tallest_run_metrics(&tp.runs, seen_fonts);
                            let lh = resolve_line_h(tp_ls, fs, None);
                            tb_cursor -= tp.space_before + lh + tp.space_after;
                            continue;
                        }
                        let (tb_fs, _, tb_ar) = tallest_run_metrics(&tp.runs, seen_fonts);
                        let tb_ascender = tb_ar.unwrap_or(0.75);
                        let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_ar);
                        let tb_baseline = tb_cursor - tp.space_before - tb_fs * tb_ascender;
                        if !tp.list_label.is_empty() {
                            let label_x = content_x + tp.indent_left - tp.indent_hanging;
                            let (label_font_name, label_bytes) =
                                label_for_paragraph(tp, seen_fonts);
                            if let Some([r, g, b]) = tp.runs.first().and_then(|r| r.color) {
                                content.set_fill_rgb(
                                    r as f32 / 255.0,
                                    g as f32 / 255.0,
                                    b as f32 / 255.0,
                                );
                            }
                            content
                                .begin_text()
                                .set_font(Name(label_font_name.as_bytes()), tb_fs)
                                .next_line(label_x, tb_baseline)
                                .show(Str(&label_bytes))
                                .end_text();
                            if tp.runs.first().and_then(|r| r.color).is_some() {
                                content.set_fill_gray(0.0);
                            }
                        }
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
                            seen_fonts,
                        );
                        tb_cursor -=
                            tp.space_before + (tb_lines.len() as f32) * tb_line_h + tp.space_after;
                    }
                }

                // Render floating images
                for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                    if let Some(pdf_name) = floating_image_names.get(&(pi, fi_idx)) {
                        let img = &fi.image;
                        let fi_x = match fi.h_relative_from {
                            "page" => match fi.h_position {
                                HorizontalPosition::AlignCenter => {
                                    (sp.page_width - img.display_width) / 2.0
                                }
                                HorizontalPosition::AlignRight => sp.page_width - img.display_width,
                                HorizontalPosition::AlignLeft => 0.0,
                                HorizontalPosition::Offset(o) => o,
                            },
                            "margin" | _ => match fi.h_position {
                                HorizontalPosition::AlignCenter => {
                                    sp.margin_left + (text_width - img.display_width) / 2.0
                                }
                                HorizontalPosition::AlignRight => {
                                    sp.margin_left + text_width - img.display_width
                                }
                                HorizontalPosition::AlignLeft => sp.margin_left,
                                HorizontalPosition::Offset(o) => sp.margin_left + o,
                            },
                        };
                        let fi_y_top = super::resolve_fi_y_top(fi, sp, slot_top);
                        let fi_y_bottom = fi_y_top - img.display_height;
                        content.save_state();
                        content.transform([
                            img.display_width,
                            0.0,
                            0.0,
                            img.display_height,
                            fi_x,
                            fi_y_bottom,
                        ]);
                        content.x_object(Name(pdf_name.as_bytes()));
                        content.restore_state();
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
                        content.save_state();
                        content.transform([
                            img.display_width,
                            0.0,
                            0.0,
                            img.display_height,
                            x,
                            y_bottom,
                        ]);
                        content.x_object(Name(pdf_name.as_bytes()));
                        content.restore_state();
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

                let has_tabs = substituted_runs.iter().any(|r| r.is_tab);
                let lines = if has_tabs {
                    build_tabbed_line(
                        &substituted_runs,
                        seen_fonts,
                        &para.tab_stops,
                        0.0,
                        text_width,
                        0.0,
                        &block_inline_images,
                    )
                } else {
                    build_paragraph_lines(
                        &substituted_runs,
                        seen_fonts,
                        text_width,
                        0.0,
                        &block_inline_images,
                    )
                };

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
                    seen_fonts,
                );

                let max_img_h = lines
                    .iter()
                    .flat_map(|l| l.chunks.iter())
                    .map(|c| c.inline_image_height)
                    .fold(0.0f32, f32::max);
                let effective_line_h = if max_img_h > line_h {
                    max_img_h
                } else {
                    line_h
                };
                cursor_y -= lines.len().max(1) as f32 * effective_line_h;
                prev_space_after = para.space_after;
                pi += 1;
            }
        }
    }
}
