use std::collections::HashMap;
use std::sync::LazyLock;

use crate::fonts::{FontEntry, font_key_buf};

static EMPTY_INLINE_IMAGE_MAP: LazyLock<HashMap<usize, String>> =
    LazyLock::new(HashMap::new);
use crate::model::{
    Alignment, Block, CellMargins, HorizontalPosition, Table, TextDirection, VMerge,
    VerticalPosition, WrapType,
};

use super::RenderContext;
use super::header_footer::substitute_hf_runs;
use super::layout::{TextLine, build_paragraph_lines, build_tabbed_line, font_metric, is_text_empty};
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
/// When `available_width` is provided and the table exceeds it, all columns
/// are scaled down proportionally to fit (matching Word's behavior).
/// For nested auto-fit tables (`available_width` is Some and not fixed layout),
/// Word shrinks columns to content-based minimum widths rather than using the
/// gridCol preferred widths.
pub(super) fn auto_fit_columns(table: &Table, fonts: &HashMap<String, FontEntry>, available_width: Option<f32>) -> Vec<f32> {
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
            if cell.text_direction != TextDirection::LrTb {
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

    // For nested auto-fit tables, Word shrinks columns to content-based widths
    // rather than preserving the gridCol total. Each column is sized based on
    // a blend of the minimum width (longest word) and the natural width
    // (longest single-line paragraph), capped by the gridCol preferred width.
    if available_width.is_some() && !table.fixed_layout {
        let mut natural_widths = vec![0.0f32; ncols];
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
                    let mut para_w = 0.0f32;
                    for run in &para.runs {
                        let key = font_key_buf(run, &mut key_buf);
                        let Some(entry) = fonts.get(key) else { continue };
                        let fs = if run.small_caps {
                            (run.font_size - 2.0).max(1.0)
                        } else {
                            run.font_size
                        };
                        let text = if run.caps || run.small_caps {
                            std::borrow::Cow::Owned(run.text.to_uppercase())
                        } else {
                            std::borrow::Cow::Borrowed(&run.text)
                        };
                        let kern = run.kern_threshold.is_some_and(|t| fs >= t);
                        para_w += entry.word_width(&text, fs, kern);
                    }
                    natural_widths[grid_col] =
                        natural_widths[grid_col].max(para_w + h_pad);
                }
                grid_col += span;
            }
        }
        // Word's auto-fit for nested tables produces column widths slightly
        // below the full natural paragraph width. Scale down by 0.9 to
        // approximate Word's sizing, ensuring text wraps where Word wraps it.
        let min_cell = cm.left + cm.right;
        let mut widths: Vec<f32> = (0..ncols)
            .map(|i| {
                let mw = min_widths[i].max(min_cell);
                let nw = natural_widths[i].max(mw);
                let fitted = (nw * 0.9).max(mw);
                fitted.min(table.col_widths.get(i).copied().unwrap_or(f32::MAX))
            })
            .collect();
        if let Some(avail) = available_width {
            let total: f32 = widths.iter().sum();
            if total > avail && avail > 0.0 {
                let scale = avail / total;
                for w in &mut widths {
                    *w *= scale;
                }
            }
        }
        return widths;
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

    if let Some(avail) = available_width {
        let final_total: f32 = widths.iter().sum();
        if final_total > avail && avail > 0.0 {
            let scale = avail / final_total;
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
    pub(super) indent_first_line: f32,
    /// Extra left indent from wrapSquare/Tight floating images in this paragraph.
    /// Text lines are laid out narrower and rendered further right to avoid the image.
    pub(super) float_indent_left: f32,
    pub(super) list_label: String,
    pub(super) list_label_font: Option<String>,
    pub(super) label_color: Option<[u8; 3]>,
    pub(super) first_run_font_key: String,
    pub(super) image_name: Option<String>,
    pub(super) image_width: f32,
    pub(super) image_height: f32,
    pub(super) image_stroke_color: Option<[u8; 3]>,
    pub(super) image_stroke_width: f32,
    pub(super) image_shadow: Option<crate::model::ImageShadow>,
    pub(super) image_shadow_xobj: Option<String>,
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
                    let span_w = cell_span_width(col_widths, grid_col, span);
                    // When content-based column widths are narrower than the
                    // cell's preferred width (nested auto-fit tables), respect
                    // the computed widths instead of overriding with cell.width.
                    let grid_w = cell_span_width(&table.col_widths, grid_col, span);
                    let col_w = if span_w < grid_w * 0.9 {
                        span_w
                    } else {
                        span_w.max(cell.width)
                    };
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
                    let mut prev_was_nested_table = false;

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

                                // Compute extra left indent from left-aligned
                                // wrapSquare/Tight floating images so text wraps
                                // to the right of the image within the cell.
                                let float_indent_left: f32 = para
                                    .floating_images
                                    .iter()
                                    .filter(|fi| {
                                        matches!(
                                            fi.wrap_type,
                                            WrapType::Square | WrapType::Tight | WrapType::Through
                                        ) && matches!(
                                            fi.h_position,
                                            HorizontalPosition::AlignLeft
                                                | HorizontalPosition::Offset(_)
                                        )
                                    })
                                    .map(|fi| {
                                        let left_edge = match fi.h_position {
                                            HorizontalPosition::Offset(o) => o,
                                            _ => 0.0,
                                        };
                                        left_edge + fi.image.display_width + fi.dist_right
                                    })
                                    .fold(0.0f32, f32::max);

                                let lines = if !is_text_empty(runs) {
                                    let para_text_w = (cell_text_w
                                        - para.indent_left
                                        - para.indent_right
                                        - float_indent_left)
                                        .max(0.0);
                                    // Match the rendering's first_line_hanging: when a
                                    // list label is present, the label is drawn separately
                                    // and the text starts at indent_left, so the first
                                    // line has no extra hanging width.
                                    let hanging = if !para.list_label.is_empty() {
                                        if para.indent_first_line > 0.0
                                            && para.indent_hanging == 0.0
                                        {
                                            -para.indent_first_line
                                        } else {
                                            0.0
                                        }
                                    } else {
                                        para.indent_hanging
                                    };
                                    let has_tabs = runs.iter().any(|r| r.is_tab);
                                    let lines = if has_tabs {
                                        build_tabbed_line(
                                            runs,
                                            ctx.fonts,
                                            &para.tab_stops,
                                            para.indent_left,
                                            para_text_w,
                                            hanging,
                                            &EMPTY_INLINE_IMAGE_MAP,
                                            &EMPTY_INLINE_IMAGE_MAP,
                                            ctx.default_tab_stop,
                                        )
                                    } else {
                                        build_paragraph_lines(
                                            runs,
                                            ctx.fonts,
                                            para_text_w,
                                            hanging,
                                            &EMPTY_INLINE_IMAGE_MAP,
                                            &EMPTY_INLINE_IMAGE_MAP,
                                            None,
                                            None,
                                            None,
                                        )
                                    };
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
                                    } else if prev_was_nested_table && para_idx == 1 {
                                        // Cell = [nested table, empty ¶]: the mark
                                        // glyph height is covered by the +0.5pt
                                        // row addition; space_after is suppressed
                                        // in the trailing-space block below.
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
                                let (image_width, image_height, img_stroke_color, img_stroke_width, img_shadow) = para
                                    .image
                                    .as_ref()
                                    .map(|img| (img.display_width, img.display_height, img.stroke_color, img.stroke_width, img.shadow.clone()))
                                    .unwrap_or((0.0, 0.0, None, 0.0, None));
                                let img_shadow_xobj = para.image.as_ref().and_then(|img| {
                                    let key = std::sync::Arc::as_ptr(&img.data) as usize;
                                    ctx.shadow_table_names.get(&key).cloned()
                                });

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
                                    indent_first_line: para.indent_first_line,
                                    float_indent_left,
                                    list_label: para.list_label.clone(),
                                    list_label_font: para.list_label_font.clone(),
                                    label_color: para.runs.first().and_then(|r| r.color),
                                    first_run_font_key,
                                    image_name,
                                    image_width,
                                    image_height,
                                    image_stroke_color: img_stroke_color,
                                    image_stroke_width: img_stroke_width,
                                    image_shadow: img_shadow,
                                    image_shadow_xobj: img_shadow_xobj,
                                    content_height: para.content_height,
                                    paragraph_mark_vanish: para.paragraph_mark_vanish,
                                    floating_images: cell_floats,
                                    space_after: para.space_after,
                                }));

                                prev_space_after = para.space_after;
                                prev_was_nested_table = false;
                                para_idx += 1;
                            }
                            Block::Table(nested_table) => {
                                let nested_cw = auto_fit_columns(nested_table, ctx.fonts, Some(cell_text_w));
                                let nested_layouts =
                                    compute_row_layouts(nested_table, &nested_cw, ctx, hf_sub);
                                let nested_h: f32 =
                                    nested_layouts.iter().map(|rl| rl.height).sum();
                                total_h += nested_h;
                                items.push(CellContentItem::NestedTable { height: nested_h });
                                prev_space_after = 0.0;
                                prev_was_nested_table = true;
                                para_idx += 1;
                            }
                        }
                    }

                    // When a cell contains only a nested table plus the
                    // mandatory end-of-cell paragraph mark (empty, no text),
                    // Word does not count the trailing paragraph's space_after
                    // toward the row height — the mark glyph height and
                    // line_h are already suppressed above via prev_was_nested_table.
                    let sole_table_plus_mark = items.len() == 2
                        && matches!(items.first(), Some(CellContentItem::NestedTable { .. }))
                        && matches!(items.last(), Some(CellContentItem::Paragraph(p)) if p.lines.is_empty() && p.image_name.is_none() && p.floating_images.is_empty());
                    if !sole_table_plus_mark {
                        total_h += prev_space_after;
                    }
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

/// Pre-compute how much extra height each vMerge Restart cell spans beyond its own row.
/// Returns a map from (row_idx, grid_col) to the sum of Continue row heights below.
pub(super) fn compute_merge_spans(table: &Table, row_layouts: &[RowLayout]) -> HashMap<(usize, usize), f32> {
    // Build a grid index: vmerge_grid[row][grid_col] = VMerge value
    let max_cols = table.rows.iter().map(|r| {
        r.cells.iter().map(|c| c.grid_span.max(1) as usize).sum::<usize>()
    }).max().unwrap_or(0);
    let mut vmerge_grid: Vec<Vec<VMerge>> = Vec::with_capacity(table.rows.len());
    for row in &table.rows {
        let mut row_vmerge = vec![VMerge::None; max_cols];
        let mut col = 0usize;
        for cell in &row.cells {
            if col < max_cols {
                row_vmerge[col] = cell.v_merge;
            }
            col += cell.grid_span.max(1) as usize;
        }
        vmerge_grid.push(row_vmerge);
    }

    let mut spans = HashMap::new();
    for (ri, row) in table.rows.iter().enumerate() {
        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            if cell.v_merge == VMerge::Restart {
                let mut extra = 0.0f32;
                for next_ri in (ri + 1)..table.rows.len() {
                    if grid_col >= max_cols || vmerge_grid[next_ri][grid_col] != VMerge::Continue {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_span_width_single() {
        let widths = vec![100.0, 200.0, 300.0];
        assert_eq!(cell_span_width(&widths, 0, 1), 100.0);
        assert_eq!(cell_span_width(&widths, 1, 1), 200.0);
        assert_eq!(cell_span_width(&widths, 2, 1), 300.0);
    }

    #[test]
    fn test_cell_span_width_multi() {
        let widths = vec![100.0, 200.0, 300.0];
        assert_eq!(cell_span_width(&widths, 0, 2), 300.0);
        assert_eq!(cell_span_width(&widths, 0, 3), 600.0);
        assert_eq!(cell_span_width(&widths, 1, 2), 500.0);
    }

    #[test]
    fn test_cell_span_width_clamps_to_len() {
        let widths = vec![100.0, 200.0];
        // span=5 but only 2 columns from index 0
        assert_eq!(cell_span_width(&widths, 0, 5), 300.0);
    }

    #[test]
    fn test_cell_x_offset() {
        let widths = vec![100.0, 200.0, 300.0];
        assert_eq!(cell_x_offset(&widths, 50.0, 0), 50.0);
        assert_eq!(cell_x_offset(&widths, 50.0, 1), 150.0);
        assert_eq!(cell_x_offset(&widths, 50.0, 2), 350.0);
        assert_eq!(cell_x_offset(&widths, 50.0, 3), 650.0);
    }
}
