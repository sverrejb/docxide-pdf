use std::collections::HashMap;

use pdf_writer::Content;

use crate::fonts::{FontEntry, font_key_buf};
use crate::model::{Alignment, CellVAlign, SectionProperties, Table, VMerge};

use super::header_footer::substitute_hf_runs;
use super::resolve_line_h;
use super::RenderContext;

use super::layout::{
    TextLine, build_paragraph_lines, font_metric, is_text_empty, render_paragraph_lines,
};

/// Auto-fit column widths so that the longest non-breakable word in each column
/// fits within the cell (including padding). Columns that need more space grow;
/// other columns shrink proportionally. Total width is preserved.
fn auto_fit_columns(table: &Table, fonts: &HashMap<String, FontEntry>) -> Vec<f32> {
    let ncols = table.col_widths.len();
    if ncols == 0 {
        return table.col_widths.clone();
    }

    let mut min_widths = vec![0.0f32; ncols];

    for row in &table.rows {
        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            if grid_col >= ncols || span > 1 {
                grid_col += span;
                continue;
            }
            let mut key_buf = String::new();
            for para in &cell.paragraphs {
                for run in &para.runs {
                    let key = font_key_buf(run, &mut key_buf);
                    let Some(entry) = fonts.get(key) else {
                        continue;
                    };
                    let text = if run.caps || run.small_caps {
                        std::borrow::Cow::Owned(run.text.to_uppercase())
                    } else {
                        std::borrow::Cow::Borrowed(&run.text)
                    };
                    let fs = if run.small_caps {
                        (run.font_size - 2.0).max(1.0)
                    } else {
                        run.font_size
                    };
                    for word in text.split_whitespace() {
                        let kern = run.kern_threshold.is_some_and(|t| fs >= t);
                        let ww = entry.word_width(word, fs, kern);
                        min_widths[grid_col] = min_widths[grid_col].max(ww);
                    }
                }
            }
            grid_col += span;
        }
    }

    let total: f32 = table.col_widths.iter().sum();
    let mut widths = table.col_widths.clone();

    // Expand columns that need it, track how much extra space is needed
    let mut extra_needed: f32 = 0.0;
    let mut shrinkable: f32 = 0.0;
    for i in 0..ncols {
        if min_widths[i] > widths[i] {
            extra_needed += min_widths[i] - widths[i];
            widths[i] = min_widths[i];
        } else {
            shrinkable += widths[i] - min_widths[i];
        }
    }

    if extra_needed > 0.0 && shrinkable > 0.0 {
        let factor = extra_needed.min(shrinkable) / shrinkable;
        for i in 0..ncols {
            if widths[i] > min_widths[i] {
                let available = widths[i] - min_widths[i];
                widths[i] -= available * factor;
            }
        }
        // Normalize to preserve total
        let new_total: f32 = widths.iter().sum();
        if (new_total - total).abs() > 0.01 {
            let scale = total / new_total;
            for w in &mut widths {
                *w *= scale;
            }
        }
    }

    widths
}

struct RowLayout {
    height: f32,
    cell_lines: Vec<(Vec<TextLine>, f32, f32)>, // (lines, line_h, font_size) per cell
}

/// When provided, field codes in header/footer table runs are substituted with
/// their resolved values before layout.
struct HfSubstitution<'a> {
    page_num: usize,
    total_pages: usize,
    styleref_values: &'a HashMap<String, String>,
}

fn compute_row_layouts(
    table: &Table,
    col_widths: &[f32],
    ctx: &RenderContext,
    hf_sub: Option<&HfSubstitution>,
) -> Vec<RowLayout> {
    let cm = &table.cell_margins;
    table
        .rows
        .iter()
        .map(|row| {
            let mut max_h: f32 = 0.0;
            let mut grid_col = 0usize;
            let cell_lines: Vec<(Vec<TextLine>, f32, f32)> = row
                .cells
                .iter()
                .map(|cell| {
                    let span = cell.grid_span.max(1) as usize;
                    let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
                        .iter()
                        .sum::<f32>()
                        .max(cell.width);
                    grid_col += span;

                    if cell.v_merge == VMerge::Continue {
                        return (vec![], 14.4, 12.0);
                    }

                    let cell_text_w = (col_w - cm.left - cm.right).max(0.0);
                    let mut total_h: f32 = cm.top + cm.bottom;
                    let mut all_lines = Vec::new();
                    let mut first_font_size = 12.0f32;
                    let mut first_line_h = 14.4f32;
                    let mut prev_space_after = 0.0f32;

                    for (pi, para) in cell.paragraphs.iter().enumerate() {
                        let substituted;
                        let runs = if let Some(sub) = hf_sub {
                            substituted = substitute_hf_runs(
                                &para.runs,
                                sub.page_num,
                                sub.total_pages,
                                sub.styleref_values,
                            );
                            &substituted
                        } else {
                            &para.runs
                        };
                        let font_size = runs.first().map_or(12.0, |r| r.font_size);
                        let effective_ls = para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                        let tallest_lhr = font_metric(runs, ctx.fonts, |e| e.line_h_ratio);
                        let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);

                        if all_lines.is_empty() {
                            first_font_size = font_size;
                            first_line_h = line_h;
                        }

                        if pi > 0 {
                            total_h += f32::max(prev_space_after, para.space_before);
                        }

                        if !is_text_empty(runs) {
                            let lines = build_paragraph_lines(
                                runs,
                                ctx.fonts,
                                cell_text_w,
                                0.0,
                                &std::collections::HashMap::new(),
                            );
                            total_h += lines.len() as f32 * line_h;
                            all_lines.extend(lines);
                        }

                        prev_space_after = para.space_after;
                    }

                    total_h += prev_space_after;
                    max_h = max_h.max(total_h);
                    (all_lines, first_line_h, first_font_size)
                })
                .collect();

            // Word's row height includes the end-of-cell paragraph mark glyph,
            // adding roughly 0.5pt beyond the content metrics.
            let content_h = max_h + 0.5;
            let height = match (row.height, row.height_exact) {
                (Some(h), true) => h,
                (Some(h), false) => content_h.max(h),
                _ => content_h,
            };

            RowLayout { height, cell_lines }
        })
        .collect()
}

fn render_table_row(
    row: &crate::model::TableRow,
    layout: &RowLayout,
    col_widths: &[f32],
    cm: &crate::model::CellMargins,
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
) {
    let row_h = layout.height;
    let row_top = pb.slot_top;
    let row_bottom = row_top - row_h;

    let mut grid_col = 0usize;
    for (cell, (lines, line_h, font_size)) in row.cells.iter().zip(layout.cell_lines.iter()) {
        let span = cell.grid_span.max(1) as usize;
        let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
            .iter()
            .sum();
        let cell_x = table_left
            + col_widths[..grid_col.min(col_widths.len())]
                .iter()
                .sum::<f32>();
        grid_col += span;

        if cell.v_merge == VMerge::Continue {
            continue;
        }

        if let Some([r, g, b]) = cell.shading {
            let b_borders = &cell.borders;
            let inset = (b_borders.top.width
                + b_borders.bottom.width
                + b_borders.left.width
                + b_borders.right.width)
                / 8.0;
            pb.content.save_state();
            pb.content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
            pb.content.rect(
                cell_x + inset,
                row_bottom + inset,
                col_w - 2.0 * inset,
                row_h - 2.0 * inset,
            );
            pb.content.fill_nonzero();
            pb.content.restore_state();
        }

        if !lines.is_empty() && !lines.iter().all(|l| l.chunks.is_empty()) {
            let text_x = cell_x + cm.left;
            let text_w = (col_w - cm.left - cm.right).max(0.0);
            let first_run = cell.paragraphs.first().and_then(|p| p.runs.first());
            let mut kb = String::new();
            let ascender_ratio = first_run
                .map(|r| font_key_buf(r, &mut kb))
                .and_then(|k| ctx.fonts.get(k))
                .and_then(|e| e.ascender_ratio)
                .unwrap_or(0.75);

            let content_h = lines.len() as f32 * line_h;
            let baseline_y = match cell.v_align {
                CellVAlign::Top => row_top - cm.top - font_size * ascender_ratio,
                CellVAlign::Center => {
                    let avail = row_h - cm.top - cm.bottom;
                    let offset = (avail - content_h) / 2.0;
                    row_top - cm.top - offset.max(0.0) - font_size * ascender_ratio
                }
                CellVAlign::Bottom => {
                    let avail = row_h - cm.top - cm.bottom;
                    let offset = avail - content_h;
                    row_top - cm.top - offset.max(0.0) - font_size * ascender_ratio
                }
            };

            let alignment = cell
                .paragraphs
                .first()
                .map(|p| p.alignment)
                .unwrap_or(Alignment::Left);

            render_paragraph_lines(
                &mut pb.content,
                lines,
                &alignment,
                text_x,
                text_w,
                baseline_y,
                *line_h,
                lines.len(),
                0,
                &mut Vec::new(),
                0.0,
                ctx.fonts,
            );
        }
    }

    // Draw per-cell borders
    let mut grid_col = 0usize;
    for cell in &row.cells {
        let span = cell.grid_span.max(1) as usize;
        let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
            .iter()
            .sum();
        let bx = table_left
            + col_widths[..grid_col.min(col_widths.len())]
                .iter()
                .sum::<f32>();
        grid_col += span;

        let b = &cell.borders;
        let draw_border =
            |c: &mut Content, border: &crate::model::CellBorder, x1, y1, x2, y2| {
                if !border.present {
                    return;
                }
                c.save_state();
                c.set_line_width(border.width);
                if let Some([r, g, b]) = border.color {
                    c.set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                }
                c.move_to(x1, y1);
                c.line_to(x2, y2);
                c.stroke();
                c.restore_state();
            };

        if cell.v_merge != VMerge::Continue {
            draw_border(&mut pb.content, &b.top, bx, row_top, bx + col_w, row_top);
        }
        draw_border(
            &mut pb.content,
            &b.bottom,
            bx,
            row_bottom,
            bx + col_w,
            row_bottom,
        );
        draw_border(&mut pb.content, &b.left, bx, row_top, bx, row_bottom);
        draw_border(
            &mut pb.content,
            &b.right,
            bx + col_w,
            row_top,
            bx + col_w,
            row_bottom,
        );
    }

    pb.slot_top = row_bottom;
}

pub(super) fn render_table(
    table: &Table,
    sp: &SectionProperties,
    ctx: &RenderContext,
    pb: &mut super::PageBuilder,
    sect_idx: usize,
    prev_space_after: f32,
    override_pos: Option<(f32, f32, bool)>,
) {
    let col_widths = auto_fit_columns(table, ctx.fonts);
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, None);
    let cm = &table.cell_margins;

    let is_truly_floating = override_pos.is_some_and(|(.., restore)| restore);
    let (table_left, saved_slot_top) = if let Some((x, y, restore)) = override_pos {
        let saved = if restore { Some(pb.slot_top) } else { None };
        pb.slot_top = y;
        (x, saved)
    } else {
        (sp.margin_left + table.table_indent - cm.left, None)
    };

    if !is_truly_floating {
        pb.slot_top -= prev_space_after;
    }

    // Count contiguous header rows from the start of the table (per OOXML spec,
    // only contiguous header rows starting from row 0 are repeated).
    let header_count = table
        .rows
        .iter()
        .take_while(|r| r.is_header)
        .count();

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        log::debug!(
            "TABLE row={} row_h={:.2} cells={} slot_top={:.2}",
            ri,
            row_h,
            layout.cell_lines.len(),
            pb.slot_top
        );
        let at_page_top = (pb.slot_top - (sp.page_height - sp.margin_top)).abs() < 1.0;

        if !is_truly_floating && !at_page_top && pb.slot_top - row_h < sp.margin_bottom {
            pb.flush_page(sect_idx);
            pb.is_first_page_of_section = false;
            pb.slot_top = sp.page_height - sp.margin_top;

            // Repeat header rows on the new page
            if header_count > 0 && ri >= header_count {
                for hi in 0..header_count {
                    render_table_row(
                        &table.rows[hi],
                        &row_layouts[hi],
                        &col_widths,
                        cm,
                        table_left,
                        pb,
                        ctx,
                    );
                }
            }
        }

        render_table_row(row, layout, &col_widths, cm, table_left, pb, ctx);
    }

    if let Some(saved) = saved_slot_top {
        pb.slot_top = saved;
    }
}

pub(super) fn compute_hf_table_height(
    table: &Table,
    ctx: &RenderContext,
) -> f32 {
    let col_widths = auto_fit_columns(table, ctx.fonts);
    let row_layouts = compute_row_layouts(table, &col_widths, ctx, None);
    row_layouts.iter().map(|r| r.height).sum()
}

pub(super) fn render_header_footer_table(
    table: &Table,
    sp: &SectionProperties,
    ctx: &RenderContext,
    content: &mut Content,
    cursor_y: &mut f32,
    page_num: usize,
    total_pages: usize,
    styleref_values: &HashMap<String, String>,
) {
    let col_widths = auto_fit_columns(table, ctx.fonts);
    let hf_sub = HfSubstitution {
        page_num,
        total_pages,
        styleref_values,
    };
    let row_layouts = compute_row_layouts(
        table,
        &col_widths,
        ctx,
        Some(&hf_sub),
    );
    let cm = &table.cell_margins;
    let table_left = sp.margin_left + table.table_indent - cm.left;

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        let row_top = *cursor_y;
        let row_bottom = row_top - row_h;

        let mut grid_col = 0usize;
        for (cell, (lines, line_h, font_size)) in row.cells.iter().zip(layout.cell_lines.iter()) {
            let span = cell.grid_span.max(1) as usize;
            let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
                .iter()
                .sum();
            let cell_x = table_left
                + col_widths[..grid_col.min(col_widths.len())]
                    .iter()
                    .sum::<f32>();
            grid_col += span;

            if cell.v_merge == VMerge::Continue {
                continue;
            }

            if let Some([r, g, b]) = cell.shading {
                content.save_state();
                content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                content.rect(cell_x, row_bottom, col_w, row_h);
                content.fill_nonzero();
                content.restore_state();
            }

            if !lines.is_empty() && !lines.iter().all(|l| l.chunks.is_empty()) {
                let text_x = cell_x + cm.left;
                let text_w = (col_w - cm.left - cm.right).max(0.0);
                let first_run = cell.paragraphs.first().and_then(|p| p.runs.first());
                let mut kb = String::new();
                let ascender_ratio = first_run
                    .map(|r| font_key_buf(r, &mut kb))
                    .and_then(|k| ctx.fonts.get(k))
                    .and_then(|e| e.ascender_ratio)
                    .unwrap_or(0.75);

                let content_h = lines.len() as f32 * line_h;
                let baseline_y = match cell.v_align {
                    CellVAlign::Top => row_top - cm.top - font_size * ascender_ratio,
                    CellVAlign::Center => {
                        let avail = row_h - cm.top - cm.bottom;
                        let offset = (avail - content_h) / 2.0;
                        row_top - cm.top - offset.max(0.0) - font_size * ascender_ratio
                    }
                    CellVAlign::Bottom => {
                        let avail = row_h - cm.top - cm.bottom;
                        let offset = avail - content_h;
                        row_top - cm.top - offset.max(0.0) - font_size * ascender_ratio
                    }
                };

                let alignment = cell
                    .paragraphs
                    .first()
                    .map(|p| p.alignment)
                    .unwrap_or(Alignment::Left);

                render_paragraph_lines(
                    content,
                    lines,
                    &alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    *line_h,
                    lines.len(),
                    0,
                    &mut Vec::new(),
                    0.0,
                    ctx.fonts,
                );
            }
        }

        // Draw borders
        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
                .iter()
                .sum();
            let bx = table_left
                + col_widths[..grid_col.min(col_widths.len())]
                    .iter()
                    .sum::<f32>();
            grid_col += span;

            let b = &cell.borders;
            let draw_border = |content: &mut Content,
                               border: &crate::model::CellBorder,
                               x1,
                               y1,
                               x2,
                               y2| {
                if !border.present {
                    return;
                }
                content.save_state();
                content.set_line_width(border.width);
                if let Some([r, g, b]) = border.color {
                    content.set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                }
                content.move_to(x1, y1);
                content.line_to(x2, y2);
                content.stroke();
                content.restore_state();
            };

            if cell.v_merge != VMerge::Continue && ri == 0 {
                draw_border(content, &b.top, bx, row_top, bx + col_w, row_top);
            }
            draw_border(content, &b.bottom, bx, row_bottom, bx + col_w, row_bottom);
            draw_border(content, &b.left, bx, row_top, bx, row_bottom);
            draw_border(
                content,
                &b.right,
                bx + col_w,
                row_top,
                bx + col_w,
                row_bottom,
            );
        }

        *cursor_y = row_bottom;
    }
}
