use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, encode_as_gids, font_key_buf, to_winansi_bytes};
use crate::model::{Alignment, CellVAlign, SectionProperties, Table, TextDirection, VMerge};

use super::header_footer::substitute_hf_runs;
use super::resolve_line_h;
use super::RenderContext;

use super::layout::{
    TextLine, build_paragraph_lines, encode_text_for_pdf, font_metric, is_text_empty,
    render_paragraph_lines,
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

struct CellParagraphLayout {
    lines: Vec<TextLine>,
    line_h: f32,
    font_size: f32,
    ascender_ratio: f32,
    alignment: Alignment,
    space_before: f32,
    indent_left: f32,
    indent_right: f32,
    indent_hanging: f32,
    list_label: String,
    list_label_font: Option<String>,
    label_color: Option<[u8; 3]>,
    first_run_font_key: String,
}

struct CellLayout {
    paragraphs: Vec<CellParagraphLayout>,
    total_height: f32,
    text_direction: TextDirection,
}

struct RowLayout {
    height: f32,
    cells: Vec<CellLayout>,
}

fn draw_cell_label(
    content: &mut Content,
    para: &CellParagraphLayout,
    label_x: f32,
    baseline_y: f32,
    fonts: &HashMap<String, FontEntry>,
) {
    let (pdf_name, bytes) = if let Some(ref lf) = para.list_label_font {
        if let Some(entry) = fonts.get(lf.as_str()) {
            let b = match &entry.char_to_gid {
                Some(map) => encode_as_gids(&para.list_label, map),
                None => to_winansi_bytes(&para.list_label),
            };
            (entry.pdf_name.as_str(), b)
        } else {
            return;
        }
    } else if let Some(entry) = fonts.get(para.first_run_font_key.as_str()) {
        let b = match &entry.char_to_gid {
            Some(map) => encode_as_gids(&para.list_label, map),
            None => to_winansi_bytes(&para.list_label),
        };
        (entry.pdf_name.as_str(), b)
    } else {
        return;
    };

    if let Some([r, g, b]) = para.label_color {
        content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
    }
    content
        .begin_text()
        .set_font(Name(pdf_name.as_bytes()), para.font_size)
        .next_line(label_x, baseline_y)
        .show(Str(&bytes))
        .end_text();
    if para.label_color.is_some() {
        content.set_fill_gray(0.0);
    }
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
            let cells: Vec<CellLayout> = row
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
                        return CellLayout {
                            paragraphs: vec![],
                            total_height: 14.4,
                            text_direction: TextDirection::LrTb,
                        };
                    }

                    let is_rotated = cell.text_direction != TextDirection::LrTb;
                    let cell_text_w = if is_rotated {
                        10000.0
                    } else {
                        (col_w - cm.left - cm.right).max(0.0)
                    };
                    let mut total_h: f32 = cm.top + cm.bottom;
                    let mut max_rotated_line_w: f32 = 0.0;
                    let mut paragraphs = Vec::new();
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

                        let space_before = if pi > 0 {
                            f32::max(prev_space_after, para.space_before)
                        } else {
                            para.space_before
                        };
                        total_h += space_before;

                        let mut kb = String::new();
                        let ascender_ratio = runs
                            .first()
                            .map(|r| font_key_buf(r, &mut kb))
                            .and_then(|k| ctx.fonts.get(k))
                            .and_then(|e| e.ascender_ratio)
                            .unwrap_or(0.75);

                        let lines = if !is_text_empty(runs) {
                            let para_text_w = (cell_text_w - para.indent_left - para.indent_right).max(0.0);
                            let lines = build_paragraph_lines(
                                runs,
                                ctx.fonts,
                                para_text_w,
                                para.indent_hanging,
                                &std::collections::HashMap::new(),
                            );
                            if is_rotated {
                                for line in &lines {
                                    max_rotated_line_w = max_rotated_line_w.max(line.total_width);
                                }
                            }
                            total_h += lines.len() as f32 * line_h;
                            lines
                        } else {
                            vec![]
                        };

                        let first_run_font_key = runs
                            .first()
                            .map(|r| {
                                let mut kb2 = String::new();
                                font_key_buf(r, &mut kb2).to_owned()
                            })
                            .unwrap_or_default();

                        paragraphs.push(CellParagraphLayout {
                            lines,
                            line_h,
                            font_size,
                            ascender_ratio,
                            alignment: para.alignment,
                            space_before,
                            indent_left: para.indent_left,
                            indent_right: para.indent_right,
                            indent_hanging: para.indent_hanging,
                            list_label: para.list_label.clone(),
                            list_label_font: para.list_label_font.clone(),
                            label_color: para.runs.first().and_then(|r| r.color),
                            first_run_font_key,
                        });

                        prev_space_after = para.space_after;
                    }

                    total_h += prev_space_after;
                    if is_rotated {
                        total_h = cm.top + cm.bottom + max_rotated_line_w;
                    }
                    max_h = max_h.max(total_h);
                    CellLayout {
                        paragraphs,
                        total_height: total_h,
                        text_direction: cell.text_direction,
                    }
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

            RowLayout { height, cells }
        })
        .collect()
}

/// Look up the vMerge value for the cell at `target_col` in `row`.
fn vmerge_at_col(row: &crate::model::TableRow, target_col: usize) -> VMerge {
    let mut col = 0usize;
    for cell in &row.cells {
        if col == target_col {
            return cell.v_merge;
        }
        col += cell.grid_span.max(1) as usize;
        if col > target_col {
            break;
        }
    }
    VMerge::None
}

/// Pre-compute how much extra height each vMerge Restart cell spans beyond its own row.
/// Returns a map from (row_idx, grid_col) to the sum of Continue row heights below.
fn compute_merge_spans(
    table: &Table,
    row_layouts: &[RowLayout],
) -> HashMap<(usize, usize), f32> {
    let mut spans = HashMap::new();
    for (ri, row) in table.rows.iter().enumerate() {
        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            if cell.v_merge == VMerge::Restart {
                let mut extra = 0.0f32;
                for next_ri in (ri + 1)..table.rows.len() {
                    if vmerge_at_col(&table.rows[next_ri], grid_col) != VMerge::Continue {
                        break;
                    }
                    extra += row_layouts[next_ri].height;
                }
                if extra > 0.0 {
                    spans.insert((ri, grid_col), extra);
                }
            }
            grid_col += span;
        }
    }
    spans
}

fn render_table_row(
    row: &crate::model::TableRow,
    layout: &RowLayout,
    col_widths: &[f32],
    cm: &crate::model::CellMargins,
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
    row_idx: usize,
    merge_spans: &HashMap<(usize, usize), f32>,
) {
    let row_h = layout.height;
    let row_top = pb.slot_top;
    let row_bottom = row_top - row_h;

    let mut grid_col = 0usize;
    for (cell, cell_layout) in row.cells.iter().zip(layout.cells.iter()) {
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

        let has_content = cell_layout
            .paragraphs
            .iter()
            .any(|p| !p.lines.is_empty() && !p.lines.iter().all(|l| l.chunks.is_empty()));

        if has_content && cell_layout.text_direction == TextDirection::TbRl {
            // Vertical CJK text: characters upright, stacked top-to-bottom
            let pdf_name_to_entry: HashMap<&str, &FontEntry> = ctx
                .fonts
                .values()
                .map(|e| (e.pdf_name.as_str(), e))
                .collect();

            // Compute total content height (characters × font_size)
            let mut total_char_h = 0.0f32;
            for para in &cell_layout.paragraphs {
                for line in &para.lines {
                    for chunk in &line.chunks {
                        if !chunk.text.is_empty() {
                            total_char_h += chunk.text.chars().count() as f32 * chunk.font_size;
                        }
                    }
                }
            }

            let avail_h = row_h - cm.top - cm.bottom;
            let v_offset = match cell.v_align {
                CellVAlign::Top => 0.0,
                CellVAlign::Center => ((avail_h - total_char_h) / 2.0).max(0.0),
                CellVAlign::Bottom => (avail_h - total_char_h).max(0.0),
            };

            let avail_w = col_w - cm.left - cm.right;
            let mut char_y = row_top - cm.top - v_offset;

            for para in &cell_layout.paragraphs {
                for line in &para.lines {
                    for chunk in &line.chunks {
                        if chunk.text.is_empty() {
                            continue;
                        }
                        let fs = chunk.font_size;
                        let ascender_ratio = pdf_name_to_entry
                            .get(chunk.pdf_font.as_str())
                            .and_then(|e| e.ascender_ratio)
                            .unwrap_or(0.75);

                        if let Some([r, g, b]) = chunk.color {
                            pb.content.set_fill_rgb(
                                r as f32 / 255.0,
                                g as f32 / 255.0,
                                b as f32 / 255.0,
                            );
                        }

                        pb.content.begin_text();
                        pb.content
                            .set_font(Name(chunk.pdf_font.as_bytes()), fs);

                        let mut td_x = 0.0f32;
                        let mut td_y = 0.0f32;
                        for ch in chunk.text.chars() {
                            let baseline_y = char_y - fs * ascender_ratio;
                            // Center each character horizontally using its actual width
                            let char_w = pdf_name_to_entry
                                .get(chunk.pdf_font.as_str())
                                .and_then(|e| e.char_widths_1000.as_ref())
                                .and_then(|m| m.get(&ch))
                                .map(|w| w * fs / 1000.0)
                                .unwrap_or(fs);
                            let cx = cell_x + cm.left + (avail_w - char_w) / 2.0;

                            let ch_str = ch.to_string();
                            let bytes = encode_text_for_pdf(
                                &ch_str,
                                &chunk.pdf_font,
                                &pdf_name_to_entry,
                            );
                            pb.content.next_line(cx - td_x, baseline_y - td_y);
                            td_x = cx;
                            td_y = baseline_y;
                            pb.content.show(Str(&bytes));

                            char_y -= fs;
                        }
                        pb.content.end_text();

                        if chunk.color.is_some() {
                            pb.content.set_fill_gray(0.0);
                        }
                    }
                }
            }
        } else if has_content {
            let content_h: f32 = cell_layout
                .paragraphs
                .iter()
                .map(|p| p.space_before + p.lines.len() as f32 * p.line_h)
                .sum();

            let v_offset = match cell.v_align {
                CellVAlign::Top => 0.0,
                CellVAlign::Center => {
                    let avail = row_h - cm.top - cm.bottom;
                    ((avail - content_h) / 2.0).max(0.0)
                }
                CellVAlign::Bottom => {
                    let avail = row_h - cm.top - cm.bottom;
                    (avail - content_h).max(0.0)
                }
            };

            let mut cursor_y = row_top - cm.top - v_offset;

            for para in &cell_layout.paragraphs {
                if para.lines.is_empty() || para.lines.iter().all(|l| l.chunks.is_empty()) {
                    cursor_y -= para.space_before + para.lines.len() as f32 * para.line_h;
                    continue;
                }

                cursor_y -= para.space_before;
                let text_x = cell_x + cm.left + para.indent_left;
                let text_w = (col_w - cm.left - cm.right - para.indent_left).max(0.0);
                let baseline_y = cursor_y - para.font_size * para.ascender_ratio;

                if !para.list_label.is_empty() {
                    let label_x = cell_x + cm.left + para.indent_left - para.indent_hanging;
                    draw_cell_label(&mut pb.content, para, label_x, baseline_y, ctx.fonts);
                }

                render_paragraph_lines(
                    &mut pb.content,
                    &para.lines,
                    &para.alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    para.line_h,
                    para.lines.len(),
                    0,
                    &mut Vec::new(),
                    0.0,
                    ctx.fonts,
                );

                cursor_y -= para.lines.len() as f32 * para.line_h;
            }
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

        if cell.v_merge == VMerge::Continue {
            grid_col += span;
            continue;
        }

        let merge_extra = merge_spans
            .get(&(row_idx, grid_col))
            .copied()
            .unwrap_or(0.0);
        let effective_bottom = row_bottom - merge_extra;

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

        draw_border(&mut pb.content, &b.top, bx, row_top, bx + col_w, row_top);
        draw_border(
            &mut pb.content,
            &b.bottom,
            bx,
            effective_bottom,
            bx + col_w,
            effective_bottom,
        );
        draw_border(&mut pb.content, &b.left, bx, row_top, bx, effective_bottom);
        draw_border(
            &mut pb.content,
            &b.right,
            bx + col_w,
            row_top,
            bx + col_w,
            effective_bottom,
        );

        grid_col += span;
    }

    pb.slot_top = row_bottom;
}

/// Find how many paragraphs (from `start`) fit within `available_h`.
/// Always includes at least one paragraph to guarantee progress.
fn find_cell_split(
    cell: &CellLayout,
    start: usize,
    available_h: f32,
    cm: &crate::model::CellMargins,
) -> usize {
    if start >= cell.paragraphs.len() {
        return cell.paragraphs.len();
    }
    let mut h = cm.top + cm.bottom;
    for pi in start..cell.paragraphs.len() {
        let para = &cell.paragraphs[pi];
        let sb = if pi == start { 0.0 } else { para.space_before };
        let para_h = sb + para.lines.len() as f32 * para.line_h;
        if h + para_h > available_h && pi > start {
            return pi;
        }
        h += para_h;
    }
    cell.paragraphs.len()
}

/// Render a subset of each cell's paragraphs for a split row.
/// `starts[ci]..ends[ci]` gives the paragraph range for cell `ci`.
/// `is_first`/`is_last` control top/bottom border drawing.
fn render_partial_row(
    row: &crate::model::TableRow,
    layout: &RowLayout,
    col_widths: &[f32],
    cm: &crate::model::CellMargins,
    table_left: f32,
    pb: &mut super::PageBuilder,
    ctx: &RenderContext,
    starts: &[usize],
    ends: &[usize],
    is_first: bool,
    is_last: bool,
) {
    let mut max_h: f32 = cm.top + cm.bottom;
    for (ci, cell_layout) in layout.cells.iter().enumerate() {
        let start = starts[ci];
        let end = ends[ci];
        let mut h = cm.top + cm.bottom;
        for pi in start..end {
            let para = &cell_layout.paragraphs[pi];
            let sb = if pi == start { 0.0 } else { para.space_before };
            h += sb + para.lines.len() as f32 * para.line_h;
        }
        max_h = max_h.max(h);
    }

    let row_h = max_h;
    let row_top = pb.slot_top;
    let row_bottom = row_top - row_h;

    let mut grid_col = 0usize;
    for (ci, (cell, cell_layout)) in row.cells.iter().zip(layout.cells.iter()).enumerate() {
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

        let start = starts[ci];
        let end = ends[ci];

        if let Some([r, g, b]) = cell.shading {
            let b_borders = &cell.borders;
            let inset = (b_borders.top.width
                + b_borders.bottom.width
                + b_borders.left.width
                + b_borders.right.width)
                / 8.0;
            pb.content.save_state();
            pb.content
                .set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
            pb.content.rect(
                cell_x + inset,
                row_bottom + inset,
                col_w - 2.0 * inset,
                row_h - 2.0 * inset,
            );
            pb.content.fill_nonzero();
            pb.content.restore_state();
        }

        let has_content = (start..end).any(|pi| {
            let p = &cell_layout.paragraphs[pi];
            !p.lines.is_empty() && !p.lines.iter().all(|l| l.chunks.is_empty())
        });

        if has_content {
            let mut cursor_y = row_top - cm.top;

            for pi in start..end {
                let para = &cell_layout.paragraphs[pi];
                let sb = if pi == start { 0.0 } else { para.space_before };

                if para.lines.is_empty() || para.lines.iter().all(|l| l.chunks.is_empty()) {
                    cursor_y -= sb + para.lines.len() as f32 * para.line_h;
                    continue;
                }

                cursor_y -= sb;
                let text_x = cell_x + cm.left + para.indent_left;
                let text_w = (col_w - cm.left - cm.right - para.indent_left).max(0.0);
                let baseline_y = cursor_y - para.font_size * para.ascender_ratio;

                if !para.list_label.is_empty() {
                    let label_x = cell_x + cm.left + para.indent_left - para.indent_hanging;
                    draw_cell_label(&mut pb.content, para, label_x, baseline_y, ctx.fonts);
                }

                render_paragraph_lines(
                    &mut pb.content,
                    &para.lines,
                    &para.alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    para.line_h,
                    para.lines.len(),
                    0,
                    &mut Vec::new(),
                    0.0,
                    ctx.fonts,
                );

                cursor_y -= para.lines.len() as f32 * para.line_h;
            }
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

        if is_first && cell.v_merge != VMerge::Continue {
            draw_border(&mut pb.content, &b.top, bx, row_top, bx + col_w, row_top);
        }
        if is_last {
            draw_border(
                &mut pb.content,
                &b.bottom,
                bx,
                row_bottom,
                bx + col_w,
                row_bottom,
            );
        }
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
    let merge_spans = compute_merge_spans(table, &row_layouts);
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
            layout.cells.len(),
            pb.slot_top
        );
        let at_page_top = (pb.slot_top - (sp.page_height - sp.margin_top)).abs() < 1.0;
        let available_h = pb.slot_top - sp.margin_bottom;
        let page_content_h = sp.page_height - sp.margin_top - sp.margin_bottom;

        if !is_truly_floating && row_h > available_h && row_h > page_content_h {
            // Row is too tall for any single page — split across pages,
            // starting on the current page to avoid wasting space
            let ncells = layout.cells.len();
            let mut starts = vec![0usize; ncells];
            let mut is_first_chunk = true;

            loop {
                let avail = pb.slot_top - sp.margin_bottom;
                let mut ends = Vec::with_capacity(ncells);
                let mut all_done = true;

                for ci in 0..ncells {
                    let end = find_cell_split(&layout.cells[ci], starts[ci], avail, cm);
                    if end < layout.cells[ci].paragraphs.len() {
                        all_done = false;
                    }
                    ends.push(end);
                }

                render_partial_row(
                    row, layout, &col_widths, cm, table_left, pb, ctx,
                    &starts, &ends, is_first_chunk, all_done,
                );

                if all_done {
                    break;
                }

                starts = ends;
                is_first_chunk = false;
                pb.flush_page(sect_idx);
                pb.is_first_page_of_section = false;
                pb.slot_top = sp.page_height - sp.margin_top;

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
                            hi,
                            &merge_spans,
                        );
                    }
                }
            }
        } else if !is_truly_floating && !at_page_top && row_h > available_h {
            // Row fits on a fresh page but not here — flush first
            pb.flush_page(sect_idx);
            pb.is_first_page_of_section = false;
            pb.slot_top = sp.page_height - sp.margin_top;

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
                        hi,
                        &merge_spans,
                    );
                }
            }
            render_table_row(row, layout, &col_widths, cm, table_left, pb, ctx, ri, &merge_spans);
        } else {
            render_table_row(row, layout, &col_widths, cm, table_left, pb, ctx, ri, &merge_spans);
        }
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
        for (cell, cell_layout) in row.cells.iter().zip(layout.cells.iter()) {
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

            let has_content = cell_layout
                .paragraphs
                .iter()
                .any(|p| !p.lines.is_empty() && !p.lines.iter().all(|l| l.chunks.is_empty()));

            if has_content {
                let content_h: f32 = cell_layout
                    .paragraphs
                    .iter()
                    .map(|p| p.space_before + p.lines.len() as f32 * p.line_h)
                    .sum();

                let v_offset = match cell.v_align {
                    CellVAlign::Top => 0.0,
                    CellVAlign::Center => {
                        let avail = row_h - cm.top - cm.bottom;
                        ((avail - content_h) / 2.0).max(0.0)
                    }
                    CellVAlign::Bottom => {
                        let avail = row_h - cm.top - cm.bottom;
                        (avail - content_h).max(0.0)
                    }
                };

                let mut cursor_y = row_top - cm.top - v_offset;

                for para in &cell_layout.paragraphs {
                    if para.lines.is_empty() || para.lines.iter().all(|l| l.chunks.is_empty()) {
                        cursor_y -= para.space_before + para.lines.len() as f32 * para.line_h;
                        continue;
                    }

                    cursor_y -= para.space_before;
                    let text_x = cell_x + cm.left + para.indent_left;
                    let text_w = (col_w - cm.left - cm.right - para.indent_left).max(0.0);
                    let baseline_y = cursor_y - para.font_size * para.ascender_ratio;

                    if !para.list_label.is_empty() {
                        let label_x = cell_x + cm.left + para.indent_left - para.indent_hanging;
                        draw_cell_label(content, para, label_x, baseline_y, ctx.fonts);
                    }

                    render_paragraph_lines(
                        content,
                        &para.lines,
                        &para.alignment,
                        text_x,
                        text_w,
                        baseline_y,
                        para.line_h,
                        para.lines.len(),
                        0,
                        &mut Vec::new(),
                        0.0,
                        ctx.fonts,
                    );

                    cursor_y -= para.lines.len() as f32 * para.line_h;
                }
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
