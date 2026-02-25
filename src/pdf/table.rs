use std::collections::HashMap;

use pdf_writer::Content;

use crate::fonts::{FontEntry, font_key, to_winansi_bytes};
use crate::model::{Alignment, CellVAlign, Document, Table, VMerge};

use super::layout::{
    LinkAnnotation, TextLine, build_paragraph_lines, font_metric, is_text_empty,
    render_paragraph_lines,
};

/// Auto-fit column widths so that the longest non-breakable word in each column
/// fits within the cell (including padding). Columns that need more space grow;
/// other columns shrink proportionally. Total width is preserved.
fn auto_fit_columns(table: &Table, seen_fonts: &HashMap<String, FontEntry>) -> Vec<f32> {
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
            for para in &cell.paragraphs {
                for run in &para.runs {
                    let key = font_key(run);
                    let Some(entry) = seen_fonts.get(&key) else {
                        continue;
                    };
                    for word in run.text.split_whitespace() {
                        let ww: f32 = to_winansi_bytes(word)
                            .iter()
                            .filter(|&&b| b >= 32)
                            .map(|&b| entry.widths_1000[(b - 32) as usize] * run.font_size / 1000.0)
                            .sum();
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

fn compute_row_layouts(
    table: &Table,
    col_widths: &[f32],
    doc: &Document,
    seen_fonts: &HashMap<String, FontEntry>,
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

                    for para in &cell.paragraphs {
                        let font_size = para.runs.first().map_or(12.0, |r| r.font_size);
                        let effective_ls = para.line_spacing.unwrap_or(doc.line_spacing);
                        let line_h = font_metric(&para.runs, seen_fonts, |e| e.line_h_ratio)
                            .map(|ratio| font_size * ratio * effective_ls)
                            .unwrap_or(font_size * 1.2 * effective_ls);

                        if all_lines.is_empty() {
                            first_font_size = font_size;
                            first_line_h = line_h;
                        }

                        if !is_text_empty(&para.runs) {
                            let lines = build_paragraph_lines(&para.runs, seen_fonts, cell_text_w, 0.0, &std::collections::HashMap::new());
                            total_h += lines.len() as f32 * line_h;
                            all_lines.extend(lines);
                        }
                    }

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

            RowLayout {
                height,
                cell_lines,
            }
        })
        .collect()
}

pub(super) fn render_table(
    table: &Table,
    doc: &Document,
    seen_fonts: &HashMap<String, FontEntry>,
    content: &mut Content,
    all_contents: &mut Vec<Content>,
    all_page_links: &mut Vec<Vec<LinkAnnotation>>,
    current_page_links: &mut Vec<LinkAnnotation>,
    slot_top: &mut f32,
    prev_space_after: f32,
) {
    let col_widths = auto_fit_columns(table, seen_fonts);
    let row_layouts = compute_row_layouts(table, &col_widths, doc, seen_fonts);
    let cm = &table.cell_margins;
    // Word positions tables so cell text aligns with the paragraph margin
    let table_left = doc.margin_left + table.table_indent - cm.left;

    *slot_top -= prev_space_after;

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        log::debug!(
            "TABLE row={} row_h={:.2} cells={} slot_top={:.2}",
            ri,
            row_h,
            layout.cell_lines.len(),
            *slot_top
        );
        let at_page_top = (*slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;

        if !at_page_top && *slot_top - row_h < doc.margin_bottom {
            all_contents.push(std::mem::replace(content, Content::new()));
            all_page_links.push(std::mem::take(current_page_links));
            *slot_top = doc.page_height - doc.margin_top;
        }

        let row_top = *slot_top;
        let row_bottom = row_top - row_h;

        let mut grid_col = 0usize;
        for (cell, (lines, line_h, font_size)) in
            row.cells.iter().zip(layout.cell_lines.iter())
        {
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
                let inset = (b_borders.top.width + b_borders.bottom.width
                    + b_borders.left.width + b_borders.right.width)
                    / 8.0;
                content.save_state();
                content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                content.rect(
                    cell_x + inset,
                    row_bottom + inset,
                    col_w - 2.0 * inset,
                    row_h - 2.0 * inset,
                );
                content.fill_nonzero();
                content.restore_state();
            }

            if !lines.is_empty() && !lines.iter().all(|l| l.chunks.is_empty()) {
                let text_x = cell_x + cm.left;
                let text_w = (col_w - cm.left - cm.right).max(0.0);
                let first_run = cell.paragraphs.first().and_then(|p| p.runs.first());
                let ascender_ratio = first_run
                    .map(font_key)
                    .and_then(|k| seen_fonts.get(&k))
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
                    seen_fonts,
                );
            }
        }

        // Draw per-cell borders with color/width
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
                |content: &mut Content, border: &crate::model::CellBorder, x1, y1, x2, y2| {
                    if !border.present {
                        return;
                    }
                    content.save_state();
                    content.set_line_width(border.width);
                    if let Some([r, g, b]) = border.color {
                        content
                            .set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                    }
                    content.move_to(x1, y1);
                    content.line_to(x2, y2);
                    content.stroke();
                    content.restore_state();
                };

            if cell.v_merge != VMerge::Continue {
                draw_border(content, &b.top, bx, row_top, bx + col_w, row_top);
            }
            draw_border(content, &b.bottom, bx, row_bottom, bx + col_w, row_bottom);
            draw_border(content, &b.left, bx, row_top, bx, row_bottom);
            draw_border(content, &b.right, bx + col_w, row_top, bx + col_w, row_bottom);
        }

        *slot_top = row_bottom;
    }
}
