use std::collections::HashMap;

use crate::fonts::{FontEntry, font_key_buf};
use crate::model::{
    Alignment, Block, CellMargins, HorizontalPosition, Table, TableRow, TextDirection, VMerge,
    VerticalPosition,
};

use super::RenderContext;
use super::header_footer::substitute_hf_runs;
use super::layout::{TextLine, build_paragraph_lines, font_metric, is_text_empty};
use super::resolve_line_h;

pub(super) fn cell_span_width(col_widths: &[f32], grid_col: usize, span: usize) -> f32 {
    col_widths[grid_col..col_widths.len().min(grid_col + span)]
        .iter()
        .sum()
}

pub(super) fn cell_x_offset(col_widths: &[f32], table_left: f32, grid_col: usize) -> f32 {
    table_left
        + col_widths[..grid_col.min(col_widths.len())]
            .iter()
            .sum::<f32>()
}

/// Height of a paragraph's text content, matching the layout computation in
/// `compute_row_layouts`. Empty paragraphs (no lines) still occupy one line
/// height unless they carry an explicit `content_height` (e.g. from an image).
pub(super) fn para_block_height(p: &CellParagraphLayout) -> f32 {
    if p.lines.is_empty() {
        if p.paragraph_mark_vanish {
            0.0
        } else if p.content_height > 0.0 {
            p.content_height
        } else {
            p.line_h
        }
    } else {
        p.lines.len() as f32 * p.line_h
    }
}

/// Auto-fit column widths so that the longest non-breakable word in each column
/// fits within the cell (including padding). Columns that need more space grow;
/// other columns shrink proportionally. Total width is preserved.
pub(super) fn auto_fit_columns(table: &Table, fonts: &HashMap<String, FontEntry>) -> Vec<f32> {
    let ncols = table.col_widths.len();
    if ncols == 0 {
        return table.col_widths.clone();
    }

    let cm = &table.cell_margins;
    let mut min_widths = vec![0.0f32; ncols];

    for row in &table.rows {
        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            if grid_col >= ncols || span > 1 {
                grid_col += span;
                continue;
            }
            let ecm = cell.cell_margins.as_ref().unwrap_or(cm);
            let h_pad = ecm.left + ecm.right;
            let mut key_buf = String::new();
            for para in cell.all_paragraphs() {
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
                        let ww = entry.word_width(word, fs, kern) + h_pad;
                        min_widths[grid_col] = min_widths[grid_col].max(ww);
                    }
                }
            }
            grid_col += span;
        }
    }

    let total: f32 = table.col_widths.iter().sum();
    let mut widths = table.col_widths.clone();

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

pub(super) struct CellFloatingImageLayout {
    pub(super) pdf_name: String,
    pub(super) display_width: f32,
    pub(super) display_height: f32,
    pub(super) h_offset: f32,
    pub(super) v_offset: f32,
    pub(super) behind_doc: bool,
}

pub(super) struct CellParagraphLayout {
    pub(super) lines: Vec<TextLine>,
    pub(super) line_h: f32,
    pub(super) font_size: f32,
    pub(super) ascender_ratio: f32,
    pub(super) alignment: Alignment,
    pub(super) space_before: f32,
    pub(super) indent_left: f32,
    pub(super) indent_right: f32,
    pub(super) indent_hanging: f32,
    pub(super) list_label: String,
    pub(super) list_label_font: Option<String>,
    pub(super) label_color: Option<[u8; 3]>,
    pub(super) first_run_font_key: String,
    pub(super) image_name: Option<String>,
    pub(super) image_width: f32,
    pub(super) image_height: f32,
    pub(super) image_stroke_color: Option<[u8; 3]>,
    pub(super) image_stroke_width: f32,
    pub(super) content_height: f32,
    pub(super) paragraph_mark_vanish: bool,
    pub(super) floating_images: Vec<CellFloatingImageLayout>,
    pub(super) space_after: f32,
}

pub(super) enum CellContentItem {
    Paragraph(CellParagraphLayout),
    NestedTable { height: f32 },
}

pub(super) struct CellLayout {
    pub(super) items: Vec<CellContentItem>,
    pub(super) total_height: f32,
    pub(super) text_direction: TextDirection,
}

pub(super) struct RowLayout {
    pub(super) height: f32,
    pub(super) cells: Vec<CellLayout>,
}

/// When provided, field codes in header/footer table runs are substituted with
/// their resolved values before layout.
pub(super) struct HfSubstitution<'a> {
    pub(super) page_num: usize,
    pub(super) total_pages: usize,
    pub(super) styleref_values: &'a HashMap<String, String>,
    pub(super) page_num_format: Option<&'a str>,
}

pub(super) fn compute_row_layouts(
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
                    let col_w = cell_span_width(col_widths, grid_col, span).max(cell.width);
                    grid_col += span;

                    if cell.v_merge == VMerge::Continue {
                        return CellLayout {
                            items: vec![],
                            total_height: 14.4,
                            text_direction: TextDirection::LrTb,
                        };
                    }

                    let ecm = cell.cell_margins.as_ref().unwrap_or(cm);
                    let is_rotated = cell.text_direction != TextDirection::LrTb;
                    let cell_text_w = if is_rotated {
                        10000.0
                    } else {
                        (col_w - ecm.left - ecm.right).max(0.0)
                    };
                    let mut total_h: f32 = ecm.top + ecm.bottom;
                    let mut max_rotated_line_w: f32 = 0.0;
                    let mut items: Vec<CellContentItem> = Vec::new();
                    let mut prev_space_after = 0.0f32;
                    let mut para_idx = 0usize;

                    for block in &cell.content {
                        match block {
                            Block::Paragraph(para) => {
                                let substituted;
                                let runs = if let Some(sub) = hf_sub {
                                    substituted = substitute_hf_runs(
                                        &para.runs,
                                        sub.page_num,
                                        sub.total_pages,
                                        sub.styleref_values,
                                        sub.page_num_format,
                                    );
                                    &substituted
                                } else {
                                    &para.runs
                                };
                                let font_size = runs.first().map_or(12.0, |r| r.font_size);
                                let effective_ls =
                                    para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                                let tallest_lhr =
                                    font_metric(runs, ctx.fonts, |e| e.line_h_ratio);
                                let line_h =
                                    resolve_line_h(effective_ls, font_size, tallest_lhr);

                                let space_before = if para_idx > 0 {
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
                                    let para_text_w = (cell_text_w
                                        - para.indent_left
                                        - para.indent_right)
                                        .max(0.0);
                                    let lines = build_paragraph_lines(
                                        runs,
                                        ctx.fonts,
                                        para_text_w,
                                        para.indent_hanging,
                                        &std::collections::HashMap::new(),
                                        None,
                                        None,
                                        None,
                                    );
                                    if is_rotated {
                                        for line in &lines {
                                            max_rotated_line_w =
                                                max_rotated_line_w.max(line.total_width);
                                        }
                                    }
                                    total_h += lines.len() as f32 * line_h;
                                    lines
                                } else {
                                    if para.paragraph_mark_vanish {
                                        // vanished paragraph mark: zero height
                                    } else if para.content_height > 0.0 {
                                        total_h += para.content_height;
                                    } else {
                                        total_h += line_h;
                                    }
                                    vec![]
                                };

                                let first_run_font_key = runs
                                    .first()
                                    .map(|r| {
                                        let mut kb2 = String::new();
                                        font_key_buf(r, &mut kb2).to_owned()
                                    })
                                    .unwrap_or_default();

                                let image_name = para.image.as_ref().and_then(|img| {
                                    let key = std::sync::Arc::as_ptr(&img.data) as usize;
                                    ctx.table_cell_image_names.get(&key).cloned()
                                });
                                let (image_width, image_height, img_stroke_color, img_stroke_width) = para
                                    .image
                                    .as_ref()
                                    .map(|img| (img.display_width, img.display_height, img.stroke_color, img.stroke_width))
                                    .unwrap_or((0.0, 0.0, None, 0.0));

                                let cell_floats: Vec<CellFloatingImageLayout> = para
                                    .floating_images
                                    .iter()
                                    .filter_map(|fi| {
                                        let key =
                                            std::sync::Arc::as_ptr(&fi.image.data) as usize;
                                        let pdf_name =
                                            ctx.table_cell_image_names.get(&key)?.clone();
                                        let h_offset = match fi.h_position {
                                            HorizontalPosition::Offset(o) => o,
                                            HorizontalPosition::AlignCenter => {
                                                (col_w - fi.image.display_width) / 2.0
                                            }
                                            HorizontalPosition::AlignRight => {
                                                col_w - fi.image.display_width
                                            }
                                            HorizontalPosition::AlignLeft => 0.0,
                                        };
                                        let v_offset = match fi.v_position {
                                            VerticalPosition::Offset(o) => o,
                                            _ => 0.0,
                                        };
                                        Some(CellFloatingImageLayout {
                                            pdf_name,
                                            display_width: fi.image.display_width,
                                            display_height: fi.image.display_height,
                                            h_offset,
                                            v_offset,
                                            behind_doc: fi.behind_doc,
                                        })
                                    })
                                    .collect();

                                items.push(CellContentItem::Paragraph(CellParagraphLayout {
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
                                    image_name,
                                    image_width,
                                    image_height,
                                    image_stroke_color: img_stroke_color,
                                    image_stroke_width: img_stroke_width,
                                    content_height: para.content_height,
                                    paragraph_mark_vanish: para.paragraph_mark_vanish,
                                    floating_images: cell_floats,
                                    space_after: para.space_after,
                                }));

                                prev_space_after = para.space_after;
                                para_idx += 1;
                            }
                            Block::Table(nested_table) => {
                                let nested_cw = auto_fit_columns(nested_table, ctx.fonts);
                                let nested_layouts =
                                    compute_row_layouts(nested_table, &nested_cw, ctx, hf_sub);
                                let nested_h: f32 =
                                    nested_layouts.iter().map(|rl| rl.height).sum();
                                total_h += nested_h;
                                items.push(CellContentItem::NestedTable { height: nested_h });
                                prev_space_after = 0.0;
                                para_idx += 1;
                            }
                        }
                    }

                    total_h += prev_space_after;
                    if is_rotated {
                        total_h = ecm.top + ecm.bottom + max_rotated_line_w;
                    }
                    if cell.v_merge != VMerge::Restart {
                        max_h = max_h.max(total_h);
                    }
                    CellLayout {
                        items,
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
pub(super) fn vmerge_at_col(row: &TableRow, target_col: usize) -> VMerge {
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
pub(super) fn compute_merge_spans(table: &Table, row_layouts: &[RowLayout]) -> HashMap<(usize, usize), f32> {
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

/// Find how many items (from `start`) fit within `available_h`.
/// Always includes at least one item to guarantee progress.
pub(super) fn find_cell_split(cell: &CellLayout, start: usize, available_h: f32, cm: &CellMargins) -> usize {
    if start >= cell.items.len() {
        return cell.items.len();
    }
    let mut h = cm.top + cm.bottom;
    for pi in start..cell.items.len() {
        let item_h = match &cell.items[pi] {
            CellContentItem::Paragraph(para) => {
                let sb = if pi == start { 0.0 } else { para.space_before };
                sb + para_block_height(para)
            }
            CellContentItem::NestedTable { height } => *height,
        };
        if h + item_h > available_h && pi > start {
            return pi;
        }
        h += item_h;
    }
    cell.items.len()
}
