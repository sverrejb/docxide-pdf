//! Table column widths: fixed and percentage layouts and Word's AutoFit.

use std::collections::HashMap;

use crate::fonts::{FontEntry, font_key_buf};

use crate::model::{Block, CellMargins, Table, TableCell, TextDirection};

use super::layout::word_width_for_run;

pub(super) fn cell_span_width(col_widths: &[f32], grid_col: usize, span: usize) -> f32 {
    // Clamp the start too: malformed tables (missing tblGrid, gridSpan
    // overrun) can push grid_col past the grid — yield 0 instead of panicking.
    let start = grid_col.min(col_widths.len());
    col_widths[start..col_widths.len().min(grid_col + span)]
        .iter()
        .sum()
}

/// The width a `w:tblW type="pct"` table is scaled to (`apply_pct_width`'s
/// target); `None` for a table without a percentage width.
fn pct_target(table: &Table, available_w: f32, compat_mode: u32) -> Option<f32> {
    let pct = table.width_pct?;
    let basis = if compat_mode >= 15 {
        available_w
    } else {
        available_w + table.first_cell_left_margin() + table.last_cell_right_margin()
    };
    // Word caps a table's width at the text column width.
    Some(basis * pct.min(1.0))
}

/// A `tblLayout fixed` table's columns: from the cells' preferred widths in a
/// percentage or auto-width table whose cells all name one, else the saved grid.
pub(super) fn fixed_columns(table: &Table, available_w: f32, compat_mode: u32) -> Vec<f32> {
    if (table.width_pct.is_some() || table.auto_width) && !table.cells_unsized {
        fixed_layout_columns(table, pct_target(table, available_w, compat_mode))
    } else {
        table.col_widths.clone()
    }
}

/// Columns of a `tblLayout fixed` table come from the cells' preferred widths,
/// not the saved grid (Word for Mac probes p28a–g: 720-twip cells under a
/// 2746…-twip grid make 36pt columns in a `tblW auto` table, a doubled grid
/// changes nothing, and a 100% table shares its width over them). In a
/// percentage-width table each column's preference is the widest over its
/// rows; pct cells take their share of the table width, dxa cells share the
/// rest in proportion (p28c: a 25% cell and five 360-twip ones → 25% and five
/// equal columns). When the shares overrun the width, columns are starved
/// from the right down to 12pt each and the others share what is left
/// (p28e, p28f and a corpus document: five 25% cells and a 360-twip one →
/// four columns of 25% of the width less 24pt, then two 12pt slivers where
/// Word sets one letter per line).
fn fixed_layout_columns(table: &Table, table_w: Option<f32>) -> Vec<f32> {
    const MIN_COL: f32 = 12.0;
    let n = table.col_widths.len();
    let mut dxa = vec![0.0f32; n];
    let mut pct = vec![None::<f32>; n];
    let mut used = 0usize;
    for row in &table.rows {
        for (col, span, cell) in row.grid_cells() {
            if col >= n {
                break;
            }
            used = used.max((col + span).min(n));
            if span != 1 {
                continue;
            }
            match cell.width_pct {
                Some(p) => pct[col] = Some(pct[col].map_or(p, |q| q.max(p))),
                None => dxa[col] = dxa[col].max(cell.width),
            }
        }
    }
    let Some(w) = table_w else {
        // `tblW auto`: the cells' twips as they are, the saved grid where a
        // column names none.
        return (0..n)
            .map(|i| {
                if pct[i].is_some() || dxa[i] <= 0.0 {
                    table.col_widths[i]
                } else {
                    dxa[i]
                }
            })
            .collect();
    };
    // ponytail: a column only spanned cells cover keeps the saved grid; how
    // Word splits a spanned preference over its columns is unprobed. A table
    // over 100% keeps it too, as in `apply_pct_width` (a 105% table in the
    // fixtures fills its 478pt grid, wider than the 454pt column).
    if table.width_pct.is_some_and(|p| p > 1.0)
        || (0..used).any(|i| pct[i].is_none() && dxa[i] <= 0.0)
    {
        return table.col_widths.clone();
    }
    let mut widths = vec![0.0f32; n];
    for starved in 0..=used {
        let live = used - starved;
        let base = w - MIN_COL * starved as f32;
        let pct_sum: f32 = (0..live).filter_map(|i| pct[i]).sum();
        let dxa_cols: Vec<usize> = (0..live).filter(|&i| pct[i].is_none()).collect();
        if live == 0 || pct_sum * base + dxa_cols.len() as f32 * MIN_COL <= base + 0.01 {
            for i in 0..live {
                if let Some(p) = pct[i] {
                    widths[i] = p * base;
                }
            }
            let remaining = base - pct_sum * base;
            let dxa_pref: f32 = dxa_cols.iter().map(|&i| dxa[i].max(MIN_COL)).sum();
            for &i in &dxa_cols {
                widths[i] = remaining * dxa[i].max(MIN_COL) / dxa_pref;
            }
            for width in &mut widths[live..used] {
                *width = MIN_COL;
            }
            break;
        }
    }
    widths
}

/// Scale column widths so the table occupies its `w:tblW type="pct"` share
/// of the available content width. tcW pct values are mis-read as twips at
/// parse time, but their proportions survive — only the total needs fixing.
/// A real tblGrid only shrinks, to a share of at most 100% (Word probes: 100%
/// and 80% tables over 575pt and 540pt grids; maine's 544.5pt grids wrap at
/// its 540pt column). Word keeps the saved grid when the share would widen it
/// (polish_ministry), is over 100%, or has no room for a nested table or
/// picture.
/// Before compat 15 the share is of the column plus the edge cells' left and
/// right margins, the margins the table hangs out by (Word probe: 100% tables
/// 550.7pt wide in a 540pt column with 5.4pt margins; cell spacing included).
// ponytail: keeps the grid's proportions; Word's autofit re-splits columns
// by content (a 150:425 grid came out 200:350).
pub(super) fn apply_pct_width(
    table: &Table,
    widths: &mut [f32],
    available_w: f32,
    compat_mode: u32,
) {
    let (Some(pct), Some(target)) = (table.width_pct, pct_target(table, available_w, compat_mode))
    else {
        return;
    };
    let total: f32 = widths.iter().sum();
    if total <= 0.0 || target <= 0.0 {
        return;
    }
    let scale = target / total;
    let keeps_grid = || {
        pct > 1.0
            || scale > 1.0
            || widths
                .iter()
                .zip(hard_content_widths(table, true))
                .any(|(w, m)| w * scale < m - 0.5)
    };
    if !table.grid_inferred && keeps_grid() {
        // Keeping the grid means the saved one, not columns content-fitted
        // below it: a 100% CV table over a 450pt grid fills its 451pt column
        // in Word (white cell shading over grey table shading), where ours
        // shrank to its 355pt of text.
        let grid: f32 = table.col_widths.iter().sum();
        if pct <= 1.0
            && grid > total
            && grid <= target + 0.5
            && widths.len() == table.col_widths.len()
        {
            widths.copy_from_slice(&table.col_widths);
        }
        return;
    }
    for w in widths {
        *w *= scale;
    }
}

/// Nested percentage preferences refer to the parent cell's content width,
/// even when a saved grid exists. Content fitting must not shrink a 100% table.
pub(super) fn apply_nested_pct_width(table: &Table, widths: &mut [f32], available_w: f32) {
    let Some(pct) = table.width_pct else { return };
    // Word caps a table's width at the text column width.
    let target = available_w * pct.min(1.0);
    let total: f32 = widths.iter().sum();
    if total > 0.0 && target > 0.0 {
        let scale = target / total;
        for w in widths.iter_mut() {
            *w *= scale;
        }
    }
    // The share never clips a picture: Word widens the column to it even
    // in a fixed-layout nested table (issue #142, control-02-fixed-outer).
    for (w, hard) in widths.iter_mut().zip(hard_content_widths(table, false)) {
        *w = w.max(hard);
    }
}

/// A spacer column (every cell empty) keeps its written `tcW`: Word's
/// AutoFit has no content to size it by (advocacy's 225-twip gutter beside
/// a picture; content-fitting left it 4.5pt wide and the story 7.7pt left).
fn spacer_widths(table: &Table) -> Vec<f32> {
    // The widest tcW per column, None once a cell has content.
    let mut pref: Vec<Option<f32>> = vec![Some(0.0); table.col_widths.len()];
    for row in &table.rows {
        for (grid_col, span, cell) in row.grid_cells() {
            let Some(Some(w)) = (span == 1).then(|| pref.get_mut(grid_col)).flatten() else {
                continue;
            };
            let has_content = cell.content.iter().any(|b| match b {
                Block::Table(_) => true,
                Block::Paragraph(p) => {
                    p.image.is_some()
                        || p.runs
                            .iter()
                            .any(|r| !r.text.trim().is_empty() || r.inline_image.is_some())
                }
            });
            if has_content {
                pref[grid_col] = None;
            } else {
                *w = w.max(cell.width);
            }
        }
    }
    pref.into_iter().map(|w| w.unwrap_or(0.0)).collect()
}

/// Each grid column's widest picture (and nested table, when asked) plus its
/// cell padding: content that never wraps narrower. Pictures alone size the
/// auto-fit minimum (nested tables fit their column, case51).
fn hard_content_widths(table: &Table, nested_tables: bool) -> Vec<f32> {
    column_max(table, |cell| {
        cell_hard_content_width(table, cell, nested_tables)
    })
}

/// Each grid column's largest `f` over the cells that sit in it alone.
fn column_max(table: &Table, f: impl Fn(&TableCell) -> f32) -> Vec<f32> {
    let mut max = vec![0.0f32; table.col_widths.len()];
    for row in &table.rows {
        for (grid_col, span, cell) in row.grid_cells() {
            if grid_col < max.len() && span == 1 {
                max[grid_col] = max[grid_col].max(f(cell));
            }
        }
    }
    max
}

fn cell_hard_content_width(table: &Table, cell: &TableCell, nested_tables: bool) -> f32 {
    let pad = cell.cell_margins.as_ref().unwrap_or(&table.cell_margins);
    let pad = pad.left + pad.right;
    let mut min = 0.0f32;
    for block in &cell.content {
        let w = match block {
            Block::Table(t) if nested_tables => t.col_widths.iter().sum(),
            Block::Table(_) => 0.0,
            Block::Paragraph(p) => p
                .image
                .iter()
                .chain(p.runs.iter().filter_map(|r| r.inline_image.as_ref()))
                .map(|i| i.layout_size().0)
                .fold(0.0, f32::max),
        };
        if w > 0.0 {
            min = min.max(w + pad);
        }
    }
    min
}

pub(super) fn cell_x_offset(col_widths: &[f32], table_left: f32, grid_col: usize) -> f32 {
    table_left
        + col_widths[..grid_col.min(col_widths.len())]
            .iter()
            .sum::<f32>()
}

/// Auto-fit column widths so that the longest non-breakable word in each column
/// fits within the cell (including padding). Columns that need more space grow;
/// other columns shrink proportionally. Total width is preserved.
/// Auto-fit tables wider than `available_width` are scaled down to fit.
/// Fixed-layout tables preserve their grid even when it exceeds the host cell.
/// For nested auto-fit tables (`available_width` is Some and not fixed layout),
/// Word shrinks columns to content-based minimum widths rather than using the
/// gridCol preferred widths.
/// Per-column natural (unwrapped, single-line) content width including cell
/// horizontal padding. This is the "max" width input to Word's AutoFit.
/// Columns holding nested tables are raised to fit them.
fn natural_widths(
    table: &Table,
    fonts: &HashMap<String, FontEntry>,
    cm: &crate::model::CellMargins,
) -> Vec<f32> {
    let ncols = table.col_widths.len();
    let mut natural = vec![0.0f32; ncols];
    for spanning in [false, true] {
        for row in &table.rows {
            for (grid_col, span, cell) in row.grid_cells() {
                if grid_col >= ncols || (span > 1) != spanning {
                    continue;
                }
                let ecm = cell.cell_margins.as_ref().unwrap_or(cm);
                let h_pad = ecm.left + ecm.right;
                let mut key_buf = String::new();
                for para in cell.all_paragraphs() {
                    let mut para_w = 0.0f32;
                    for run in &para.runs {
                        let key = font_key_buf(run, &mut key_buf);
                        let Some(entry) = fonts.get(key) else {
                            continue;
                        };
                        let fs = run.font_size;
                        let text = super::layout::caps_word(run, &run.text);
                        para_w += super::layout::soft_split_width(&text, |t| {
                            word_width_for_run(entry, run, t, fs)
                        });
                    }
                    let end = (grid_col + span).min(ncols);
                    let current: f32 = natural[grid_col..end].iter().sum();
                    let required = para_w + h_pad;
                    if required > current {
                        let preferred: f32 = table.col_widths[grid_col..end].iter().sum();
                        for col in grid_col..end {
                            let share = if preferred > 0.0 {
                                table.col_widths[col] / preferred
                            } else {
                                1.0 / (end - grid_col) as f32
                            };
                            natural[col] += (required - current) * share;
                        }
                    }
                }
            }
        }
    }
    raise_natural_for_nested_tables(table, fonts, cm, &mut natural);
    natural
}

/// Raise each column's natural width to fit any directly nested table at its
/// own content-fitted width (plus the host cell's h-padding). The flattened
/// paragraph widths from `natural_widths` can't see the nested grid: three
/// side-by-side nested columns need their *sum*, not the widest paragraph.
fn raise_natural_for_nested_tables(
    table: &Table,
    fonts: &HashMap<String, FontEntry>,
    cm: &CellMargins,
    natural: &mut [f32],
) {
    let ncols = natural.len();
    for row in &table.rows {
        for (grid_col, span, cell) in row.grid_cells() {
            if grid_col >= ncols || span > 1 {
                continue;
            }
            let ecm = cell.cell_margins.as_ref().unwrap_or(cm);
            let h_pad = ecm.left + ecm.right;
            for block in &cell.content {
                if let Block::Table(nt) = block {
                    let ncm = &nt.cell_margins;
                    let nat = natural_widths(nt, fonts, ncm);
                    let min_cell = ncm.left + ncm.right;
                    // A nested table's pictures never wrap either (issue #142:
                    // a 166pt plot in a fixed 40% nested table inside an
                    // auto-fit host collapsed the host column to its padding).
                    let hard = hard_content_widths(nt, false);
                    let w: f32 = nat
                        .iter()
                        .zip(&hard)
                        .map(|(&x, &h)| x.max(h).max(min_cell))
                        .sum();
                    natural[grid_col] = natural[grid_col].max(w + h_pad);
                }
            }
        }
    }
}

/// Distribute `avail` across columns proportionally to their natural ("max")
/// width, floored at each column's minimum width. Columns that would fall
/// below their minimum are pinned and the remaining width is re-shared among
/// the rest (Word's AutoFit-to-Window behavior). If the minimums alone exceed
/// `avail`, everything is scaled down to fit.
fn distribute_autofit(minw: &[f32], maxw: &[f32], avail: f32) -> Vec<f32> {
    let n = minw.len();
    let mut widths = vec![0.0f32; n];
    let mut pinned = vec![false; n];
    loop {
        let rem_avail: f32 = avail
            - (0..n)
                .filter(|&i| pinned[i])
                .map(|i| widths[i])
                .sum::<f32>();
        let active_max: f32 = (0..n).filter(|&i| !pinned[i]).map(|i| maxw[i]).sum();
        if active_max <= 0.0 {
            break;
        }
        let mut newly_pinned = false;
        for i in 0..n {
            if pinned[i] {
                continue;
            }
            let w = rem_avail * maxw[i] / active_max;
            if w < minw[i] {
                widths[i] = minw[i];
                pinned[i] = true;
                newly_pinned = true;
            }
        }
        if !newly_pinned {
            for i in 0..n {
                if !pinned[i] {
                    widths[i] = rem_avail * maxw[i] / active_max;
                }
            }
            break;
        }
    }
    let total: f32 = widths.iter().sum();
    if total > avail && total > 0.0 {
        let scale = avail / total;
        for w in &mut widths {
            *w *= scale;
        }
    }
    widths
}

/// Word squeezes an autofit table wider than its newspaper column by taking
/// from each grid column in proportion to its room above its longest word
/// (case80: two 234pt cells in a 252pt column come out 130 and 121.5).
pub(super) fn squeeze_to_width(widths: &mut [f32], min: &[f32], avail: f32) {
    let total: f32 = widths.iter().sum();
    if total <= avail {
        return;
    }
    let excess = total - avail;
    let slack: f32 = widths.iter().zip(min).map(|(w, m)| (w - m).max(0.0)).sum();
    if slack >= excess {
        for (w, m) in widths.iter_mut().zip(min) {
            *w -= (*w - m).max(0.0) * excess / slack;
        }
    } else {
        for w in widths {
            *w *= avail / total;
        }
    }
}

/// Each grid column's longest unbreakable word plus its cell padding.
pub(super) fn min_content_widths(table: &Table, fonts: &HashMap<String, FontEntry>) -> Vec<f32> {
    column_max(table, |cell| cell_min_content_width(table, cell, fonts))
}

fn cell_min_content_width(
    table: &Table,
    cell: &TableCell,
    fonts: &HashMap<String, FontEntry>,
) -> f32 {
    if cell.text_direction != TextDirection::LrTb {
        return 0.0;
    }
    let mut min = 0.0f32;
    let ecm = cell.cell_margins.as_ref().unwrap_or(&table.cell_margins);
    let h_pad = ecm.left + ecm.right;
    let mut key_buf = String::new();
    for para in cell.all_paragraphs() {
        // AutoFit reserves paragraph indents inside the cell width.
        // Fixed layout keeps its authored grid; negative indents
        // still expose extra text area there.
        let inset = if table.fixed_layout {
            h_pad + para.indent_left.min(0.0) + para.indent_right.min(0.0)
        } else {
            h_pad + para.indent_left + para.indent_right + para.indent_first_line.max(0.0)
        };
        // A word continued from the previous run adds to its width.
        let mut joined_width = 0.0f32;
        let mut prev_char = None;
        for run in &para.runs {
            if run.is_tab || run.is_line_break {
                prev_char = None;
            }
            let key = font_key_buf(run, &mut key_buf);
            let Some(entry) = fonts.get(key) else {
                prev_char = None;
                continue;
            };
            let text = super::layout::caps_word(run, &run.text);
            let fs = run.font_size;
            // Same break opportunities as line layout: a CJK sentence holds
            // no spaces but may wrap after any ideograph.
            for (spaces, word) in super::layout::split_preserving_spaces(&text) {
                let ww = word_width_for_run(entry, run, word, fs);
                let continues = !table.fixed_layout
                    && spaces == 0
                    && prev_char
                        .zip(word.chars().next())
                        .is_some_and(|(a, b)| !super::layout::breaks_between(a, b));
                joined_width = if continues { joined_width + ww } else { ww };
                prev_char = word.chars().last();
                let required = (joined_width + inset).max(h_pad);
                min = min.max(required);
            }
            if run.text.ends_with(char::is_whitespace) {
                prev_char = None;
            }
        }
    }
    min
}

/// An auto-width AutoFit table whose cells name their widths in twips is laid
/// out from those widths: each column as wide as its widest cell, the table as
/// wide as the columns, a saved grid that disagrees ignored. A column some row
/// leaves empty (gridBefore, gridAfter) is at least its grid width; one only
/// merged cells cover starts at zero. A merged cell wider than its columns
/// widens the last of them (in reading order, so the leftmost in a
/// `bidiVisual` table) by the difference, in whatever row it is. Only a
/// longer word widens a column further; a merged cell's longest word or
/// picture widens its columns to it together, each to its share of it in
/// proportion to its width plus its own longest word. A nested
/// table does the same inside its cell, its grid agreeing or not
/// (Word examples: `autofit-cell-widths`, `autofit-cell-widths-spans`,
/// `autofit-cell-widths-nested`). A top-level grid that agrees with the cells,
/// and cells too wide for `avail` (the page, or the cell holding a nested
/// table), keep the other paths: a Word-saved grid holds Word's squeeze to
/// the page.
fn cell_preferred_widths(
    table: &Table,
    fonts: &HashMap<String, FontEntry>,
    min_widths: &[f32],
    avail: Option<f32>,
    nested: bool,
) -> Option<Vec<f32>> {
    let ncols = table.col_widths.len();
    if !table.auto_width || !table.cells_twip_sized || table.fixed_layout {
        return None;
    }
    let merged = || {
        table
            .rows
            .iter()
            .flat_map(|row| row.grid_cells())
            .filter(|&(col, span, _)| span > 1 && col < ncols)
            .map(|(col, span, cell)| (col..(col + span).min(ncols), cell))
    };
    let mut widths: Vec<Option<f32>> = vec![None; ncols];
    for row in &table.rows {
        let mut end = row.grid_before;
        for (col, span, cell) in row.grid_cells() {
            end = col + span;
            if span == 1 && col < ncols {
                widths[col] = Some(widths[col].unwrap_or(0.0).max(cell.width));
            }
        }
        for col in (0..ncols).filter(|&c| c < row.grid_before || c >= end) {
            widths[col] = Some(widths[col].unwrap_or(0.0).max(table.col_widths[col]));
        }
    }
    if widths.iter().any(|w| w.is_some_and(|w| w <= 0.0)) {
        return None;
    }
    // A merged cell whose other columns are known settles its last unknown
    // one first: federal_procurement's 3192 tw span over two unsized columns
    // waits for the 5256 tw span beside a 3084 tw cell to fix the first.
    while let Some((col, w)) = merged().find_map(|(cols, cell)| {
        let mut unknown = cols.clone().filter(|&c| widths[c].is_none());
        let col = unknown.next().filter(|_| unknown.next().is_none())?;
        let known: f32 = cols.filter_map(|c| widths[c]).sum();
        Some((col, (cell.width - known).max(0.0)))
    }) {
        widths[col] = Some(w);
    }
    let mut widths: Vec<f32> = widths.into_iter().map(|w| w.unwrap_or(0.0)).collect();
    for (cols, cell) in merged() {
        let spanned: f32 = widths[cols.clone()].iter().sum();
        if cell.width > spanned {
            let last = if table.bidi_visual {
                cols.start
            } else {
                cols.end - 1
            };
            widths[last] += cell.width - spanned;
        }
    }
    if widths.iter().sum::<f32>() <= 0.0 {
        return None;
    }
    let stale = widths
        .iter()
        .zip(&table.col_widths)
        .any(|(cell, grid)| (cell - grid).abs() > 0.5);
    let avail = avail?;
    let overflows = |widths: &[f32]| widths.iter().sum::<f32>() > avail + 0.5;
    if !stale && !nested && !overflows(&widths) {
        return None;
    }
    for (w, min) in widths.iter_mut().zip(min_widths) {
        *w = w.max(*min);
    }
    for (cols, cell) in merged() {
        let need = cell_min_content_width(table, cell, fonts)
            .max(cell_hard_content_width(table, cell, false));
        if need <= widths[cols.clone()].iter().sum::<f32>() {
            continue;
        }
        let weight: f32 = cols.clone().map(|c| widths[c] + min_widths[c]).sum();
        for c in cols.clone() {
            let share = if weight > 0.0 {
                need * (widths[c] + min_widths[c]) / weight
            } else {
                need / cols.len() as f32
            };
            widths[c] = widths[c].max(share);
        }
    }
    // Cells too wide for the page give up width in proportion to their room
    // above their longest word, nested table or cell margins (Word probes on
    // case100: 9889 tw of cells on a 9357 tw page, every column within
    // 0.6pt; autofit-saved-grid-auto-nested-address's empty column keeps
    // its 10.8pt of margins).
    if overflows(&widths) {
        if nested {
            return None;
        }
        let floor: Vec<f32> = min_widths
            .iter()
            .zip(hard_content_widths(table, true))
            .map(|(m, h)| {
                m.max(h)
                    .max(table.cell_margins.left + table.cell_margins.right)
            })
            .collect();
        squeeze_to_width(&mut widths, &floor, avail);
    }
    Some(widths)
}

/// Each grid column's narrowest auto-fit width: its longest word, and at
/// least its widest picture, since a picture never wraps (advocacy's nested
/// picture table collapsed to its padding, the picture clipped away).
pub(super) fn autofit_min_widths(table: &Table, fonts: &HashMap<String, FontEntry>) -> Vec<f32> {
    min_content_widths(table, fonts)
        .iter()
        .zip(hard_content_widths(table, false))
        .map(|(text, hard)| text.max(hard))
        .collect()
}

/// `fill_width`, when `Some`, is the content width a top-level AutoFit-to-Window
/// table should fill. It is kept separate from `available_width` (which drives
/// the nested-table shrink path) so passing a fill target does not accidentally
/// push a top-level table onto the shrink path.
pub(super) fn auto_fit_columns(
    table: &Table,
    fonts: &HashMap<String, FontEntry>,
    available_width: Option<f32>,
    fill_width: Option<f32>,
    overhang: f32,
) -> Vec<f32> {
    let ncols = table.col_widths.len();
    if ncols == 0 {
        return table.col_widths.clone();
    }

    // A nested table sized in twips (tblW dxa) keeps its grid like a top-level
    // one: a survey's 465pt nested table inside an auto-width host fills it
    // in Word, where content-fitting made it half as wide.
    if available_width.is_some() && !table.auto_width && !table.fixed_layout && !table.grid_inferred
    {
        return auto_fit_columns(table, fonts, None, None, 0.0);
    }

    let cm = &table.cell_margins;
    let min_widths = autofit_min_widths(table, fonts);

    // Word's AutoFit (tblLayout=autofit, the default) ignores the stored
    // gridCol widths for a `tblW type="auto"` table and re-derives column
    // widths from cell *content*, filling the available width (AutoFit to
    // Window). OOXML §17.18.87: "uses the contents of each cell to determine
    // final column widths."
    //
    // We only apply this to a uniform-grid `type="auto"` table whose cells
    // directly hold a nested table. A Word-saved autofit table stores
    // content-derived (unequal) gridCol widths, so honoring those reproduces
    // Word — recomputing from our own font metrics would only drift. The narrow
    // case where the stored grid is provably meaningless is an equal-column grid
    // wrapping a nested table (case51's 4680/4680 outer tables): the nested
    // table establishes a hard content width that the equal split ignores, so
    // Word sizes purely to content. Gating on a directly-nested table keeps
    // ordinary text tables (which legitimately keep ~equal columns) on the
    // gridCol path and avoids the corpus-wide redistribution regressions.
    let grid_uniform = ncols >= 2 && table.col_widths.iter().all(|&w| w > 0.0) && {
        let first = table.col_widths[0];
        table
            .col_widths
            .iter()
            .all(|&w| (w - first).abs() <= first * 0.02 + 0.5)
    };
    let has_nested_table = table.rows.iter().any(|r| {
        r.cells
            .iter()
            .any(|c| c.content.iter().any(|b| matches!(b, Block::Table(_))))
    });
    // Cells without a preferred width (tcW absent, auto or 0) leave Word only
    // the content to go by: a 2262/2828/2828 grid of "Item"/"Value"/"Notes"
    // shrinks to the words (Word probes, issue #140). A dxa tcW keeps the grid.
    // ponytail: content too wide for the page keeps the grid, as a Word-saved
    // grid already holds Word's full-width result (nabl's 9-column amendment
    // sheet) and our redistribution drifts from it; a made-up grid stays wrong.
    if table.auto_width
        && !table.fixed_layout
        && ((grid_uniform && has_nested_table) || table.cells_unsized)
        && let Some(avail) = fill_width.filter(|a| *a > 0.0)
    {
        let natural = natural_widths(table, fonts, cm);
        let min_cell = cm.left + cm.right;
        let maxw: Vec<f32> = (0..ncols)
            .map(|i| natural[i].max(min_widths[i]).max(min_cell))
            .collect();
        // AutoFit to Contents: when every column fits at max-content
        // width, Word leaves the table narrower than the window rather
        // than stretching it to fill.
        if maxw.iter().sum::<f32>() <= avail {
            return maxw;
        }
        if grid_uniform && has_nested_table {
            let minw: Vec<f32> = (0..ncols).map(|i| min_widths[i].max(min_cell)).collect();
            return distribute_autofit(&minw, &maxw, avail);
        }
    }

    // Uniform cells against an asymmetric grid take the rule too, as do nested
    // tables inside their cell; only a content-derived grid around a nested
    // table keeps its host column (`keep_saved_grid` below). A top-level
    // table's cells must fit beside its indent: uk_commercial_lease's 8720 tw
    // of cells at tblInd 851 overflow the 9073 tw text width, and Word drew
    // its saved 8222 tw squeeze.
    // A top-level table holding a nested one takes it as well: Word sized
    // case100's host from its 5036/340/340/4173 tw cells, not its grid.
    if let Some(widths) = cell_preferred_widths(
        table,
        fonts,
        &min_widths,
        fill_width
            .map(|w| w - table.table_indent + overhang)
            .or(available_width),
        fill_width.is_none(),
    ) {
        return widths;
    }

    // For nested auto-fit tables, Word shrinks columns to content-based widths
    // rather than preserving the gridCol total. Each column is sized based on
    // a blend of the minimum width (longest word) and the natural width
    // (longest single-line paragraph), capped by the gridCol preferred width.
    // A single-row nested table with explicit cell preferences is already
    // sized by the author. Word keeps that grid when it fits the parent,
    // including empty cells reserved for later number/date filling.
    // Do not collapse those slots to the current text's natural width.
    if let Some(avail) = available_width
        && table.auto_width
        && !table.fixed_layout
        && !table.grid_inferred
        && table.rows.len() == 1
        && table.col_widths.iter().sum::<f32>() <= avail
        && table.rows[0]
            .grid_cells()
            .all(|(_, span, cell)| span == 1 && cell.width > 0.0)
        && min_widths
            .iter()
            .zip(&table.col_widths)
            .all(|(minimum, preferred)| minimum <= preferred)
    {
        return table.col_widths.clone();
    }

    if available_width.is_some() && !table.fixed_layout {
        let natural_widths = natural_widths(table, fonts, cm);
        // Word's auto-fit for nested tables produces column widths slightly
        // below the full natural paragraph width. Scale down by 0.9 to
        // approximate Word's sizing, ensuring text wraps where Word wraps it.
        let min_cell = cm.left + cm.right;
        let avail = available_width.unwrap_or(0.0);
        let preferred_total: f32 = table.col_widths.iter().sum();
        // When the nested table has an explicit tblInd and the gridCol
        // preferred widths fit inside the parent cell, Word uses those
        // preferred widths rather than shrinking to content. An explicit
        // tblInd signals the author deliberately sized and positioned the
        // nested table, so its column hints should be honored.
        let mut widths: Vec<f32> =
            if table.table_indent_explicit && preferred_total > 0.0 && preferred_total <= avail {
                (0..ncols)
                    .map(|i| {
                        let pref = table.col_widths.get(i).copied().unwrap_or(0.0);
                        let mw = min_widths[i].max(min_cell);
                        pref.max(mw)
                    })
                    .collect()
            } else {
                // At full natural width the table fits the parent cell → Word
                // keeps the content-fitted widths (AutoFit to Contents), ignoring
                // the stored gridCol hints (§17.18.87 derives purely from cell
                // content). Only when it overflows does Word squeeze below
                // natural width; the 0.9 factor approximates that squeeze.
                let spacer = spacer_widths(table);
                let full: Vec<f32> = (0..ncols)
                    .map(|i| {
                        natural_widths[i]
                            .max(min_widths[i].max(min_cell))
                            .max(spacer[i])
                    })
                    .collect();
                if full.iter().sum::<f32>() <= avail {
                    full
                } else {
                    (0..ncols)
                        .map(|i| {
                            let mw = min_widths[i].max(min_cell);
                            let nw = natural_widths[i].max(mw);
                            let fitted = (nw * 0.9).max(mw);
                            fitted.min(table.col_widths.get(i).copied().unwrap_or(f32::MAX))
                        })
                        .collect()
                }
            };
        let total: f32 = widths.iter().sum();
        if total > avail && avail > 0.0 {
            let scale = avail / total;
            for w in &mut widths {
                *w *= scale;
            }
        }
        return widths;
    }

    // Per OOXML §17.18.87: the fixed-width base (used by auto-fit too)
    // sets each grid column to the maximum preferred width (tcW) from
    // all cells at that column, then scales proportionally if the total
    // exceeds the table width.
    let total: f32 = table.col_widths.iter().sum();
    let mut preferred = table.col_widths.clone();
    // A non-uniform saved AutoFit grid containing nested tables is already
    // content-derived. Enlarging it from stale tcW preferences and rescaling
    // the total can narrow the host column below its nested table's width.
    // Uniform tcW values can be stale after Word has fitted the saved grid
    // to content. Reapplying and scaling them erases its asymmetric widths.
    let uniform_cell_preferences = table
        .rows
        .iter()
        .flat_map(|row| row.grid_cells())
        .filter(|(_, span, _)| *span == 1)
        .map(|(_, _, cell)| cell.width)
        .collect::<Vec<_>>();
    let saved_grid_asymmetric = table
        .col_widths
        .iter()
        .any(|&w| (w - table.col_widths[0]).abs() > 0.01);
    let stale_uniform_preferences = uniform_cell_preferences.len() >= ncols
        && uniform_cell_preferences
            .iter()
            .all(|&w| w > 0.0 && (w - uniform_cell_preferences[0]).abs() < 0.01)
        && uniform_cell_preferences[0] * ncols as f32 > total;
    // An explicit table width can also carry a content-derived saved grid.
    // A tcW that no longer fits its cell's text is a stale preference, not
    // evidence that the saved nested-table host column should be rescaled.
    let stale_content_preference = || {
        table.rows.iter().any(|row| {
            row.grid_cells().any(|(col, span, cell)| {
                span == 1
                    && cell.width > 0.0
                    && min_widths
                        .get(col)
                        .is_some_and(|minimum| *minimum > cell.width + 0.01)
            })
        })
    };
    let keep_saved_grid = !table.fixed_layout
        && ((!grid_uniform
            && has_nested_table
            && (table.auto_width || stale_content_preference()))
            || (table.auto_width && saved_grid_asymmetric && stale_uniform_preferences));
    if !keep_saved_grid {
        // A row whose cell widths add up to the table but disagree with the
        // saved grid replaces that stale grid: Word draws case91's
        // 1000/1517/479/1754 cells over an 851/1666/479/1754 grid as stated,
        // where the larger of each pair, rescaled, moved columns by up to 5pt.
        // An AutoFit-width table's saved grid is Word's own fit and stays
        // (primary_care_safeguarding re-wrapped under its stated cells).
        let stated_row = table.rows.iter().filter(|_| !table.auto_width).find(|row| {
            let cells: Vec<_> = row.grid_cells().collect();
            cells.len() == ncols
                && cells
                    .iter()
                    .all(|(_, span, cell)| *span == 1 && cell.width > 0.0)
                && (cells.iter().map(|(_, _, c)| c.width).sum::<f32>() - total).abs() < 0.5
                && cells
                    .iter()
                    .any(|(col, _, cell)| (cell.width - table.col_widths[*col]).abs() > 0.5)
        });
        if let Some(row) = stated_row {
            for (grid_col, _, cell) in row.grid_cells() {
                preferred[grid_col] = cell.width;
            }
        }
        for row in &table.rows {
            for (grid_col, span, cell) in row.grid_cells() {
                if grid_col >= ncols {
                    break;
                }
                if span == 1 {
                    preferred[grid_col] = preferred[grid_col].max(cell.width);
                } else {
                    // Distribute multi-span cell width proportionally across
                    // the spanned grid columns.
                    let grid_sum: f32 = table.col_widths[grid_col..ncols.min(grid_col + span)]
                        .iter()
                        .sum();
                    if grid_sum > 0.0 && cell.width > grid_sum {
                        for g in grid_col..ncols.min(grid_col + span) {
                            let share = cell.width * (table.col_widths[g] / grid_sum);
                            preferred[g] = preferred[g].max(share);
                        }
                    }
                }
            }
        }
    }
    let pref_total: f32 = preferred.iter().sum();
    let mut widths = if pref_total > total && total > 0.0 {
        let scale = total / pref_total;
        preferred.iter().map(|&w| w * scale).collect::<Vec<_>>()
    } else {
        preferred
    };

    // Apply content minimums: ensure each column fits its longest word.
    // If any column needed boosting, shrink others proportionally.
    // Not for a percentage-width table whose cells all state percentages:
    // Word keeps those shares and breaks a word that does not fit its
    // column (a 23-column landscape table: "Total" prints "Tot|al" in a
    // 23pt column where we widened it and starved the first column to
    // its longest word, wrapping "Indicator" and running a second page).
    let pct_shares = table.width_pct.is_some()
        && table
            .rows
            .iter()
            .flat_map(|r| r.grid_cells())
            .filter(|(_, span, _)| *span == 1)
            .all(|(_, _, c)| c.width_pct.is_some());
    let mut extra_needed: f32 = 0.0;
    let mut shrinkable: f32 = 0.0;
    for i in (0..ncols).filter(|_| !pct_shares) {
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

    if let Some(avail) = available_width
        && !table.fixed_layout
    {
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

#[cfg(test)]
mod tests {
    use super::*;

    /// A fixed-layout table parsed from its XML: a 2746…-twip grid of six
    /// columns and the given rows of `w:tcW` (width, type) pairs.
    fn fixed_table(tbl_w: &str, rows: &[&[(u32, &str)]]) -> Table {
        let grid = [2746, 2770, 2869, 1409, 1410, 2744]
            .iter()
            .map(|w| format!(r#"<w:gridCol w:w="{w}"/>"#))
            .collect::<String>();
        let rows = rows
            .iter()
            .map(|cells| {
                let cells = cells
                    .iter()
                    .map(|(w, t)| {
                        format!(r#"<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="{t}"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>"#)
                    })
                    .collect::<String>();
                format!("<w:tr>{cells}</w:tr>")
            })
            .collect::<String>();
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr>{tbl_w}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl><w:sectPr/></w:body></w:document>"#
        );
        let mut doc = crate::docx::parse_document_xml(&xml);
        let Block::Table(table) = doc.sections[0].blocks.remove(0) else {
            panic!("table");
        };
        table
    }

    #[test]
    fn fixed_layout_columns_come_from_the_cells() {
        let pct = r#"<w:tblW w:w="5000" w:type="pct"/>"#;
        let auto = r#"<w:tblW w:w="0" w:type="auto"/>"#;
        let six_720: &[(u32, &str)] = &[(720, "dxa"); 6];
        let quarter_then_360: &[(u32, &str)] = &[
            (1250, "pct"),
            (360, "dxa"),
            (360, "dxa"),
            (360, "dxa"),
            (360, "dxa"),
            (360, "dxa"),
        ];
        let five_quarters: &[(u32, &str)] = &[
            (1250, "pct"),
            (1250, "pct"),
            (1250, "pct"),
            (1250, "pct"),
            (1250, "pct"),
            (360, "dxa"),
        ];
        let close = |got: &[f32], want: &[f32]| {
            assert_eq!(got.len(), want.len(), "{got:?}");
            for (g, w) in got.iter().zip(want) {
                assert!((g - w).abs() < 0.01, "{got:?} vs {want:?}");
            }
        };
        // p28d: tblW auto keeps the cells' twips, the grid is ignored.
        close(
            &fixed_layout_columns(&fixed_table(auto, &[six_720, six_720]), None),
            &[36.0; 6],
        );
        // p28a: equal cells share a 100% table equally.
        close(
            &fixed_layout_columns(&fixed_table(pct, &[six_720]), Some(468.0)),
            &[78.0; 6],
        );
        // p28c: a 25% cell takes its share, the dxa cells share the rest.
        close(
            &fixed_layout_columns(&fixed_table(pct, &[six_720, quarter_then_360]), Some(512.0)),
            &[128.0, 76.8, 76.8, 76.8, 76.8, 76.8],
        );
        // p28e and the corpus document: five 25% cells overrun the width; the
        // last two columns starve to 12pt and the rest share 512 − 24.
        close(
            &fixed_layout_columns(
                &fixed_table(pct, &[six_720, quarter_then_360, five_quarters]),
                Some(512.0),
            ),
            &[122.0, 122.0, 122.0, 122.0, 12.0, 12.0],
        );
    }

    #[test]
    fn nested_autofit_preserves_each_level_of_cell_padding() {
        let mut content = String::from(
            r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="TestFont" w:hAnsi="TestFont"/><w:sz w:val="20"/></w:rPr><w:t>abc def</w:t></w:r></w:p>"#,
        );
        let fonts = HashMap::from([("TestFont".into(), FontEntry::stub())]);
        for levels in 1..=3 {
            content = format!(
                r#"<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>{content}</w:tc></w:tr></w:tbl>"#
            );
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{content}<w:sectPr/></w:body></w:document>"#
            );
            let doc = crate::docx::parse_document_xml(&xml);
            let Block::Table(table) = &doc.sections[0].blocks[0] else {
                panic!("table");
            };
            let width = auto_fit_columns(table, &fonts, Some(150.0), None, 0.0)[0];
            assert!((width - (35.0 + 10.0 * levels as f32)).abs() < 0.001);
        }
    }

    #[test]
    fn natural_width_includes_merged_cell_text() {
        const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="3"/></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="TestFont" w:hAnsi="TestFont"/><w:sz w:val="20"/></w:rPr><w:t>{}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#,
            "x".repeat(100)
        );
        let mut doc = crate::docx::parse_document_xml(&xml);
        let Block::Table(table) = doc.sections[0].blocks.remove(0) else {
            panic!("table")
        };
        let fonts = HashMap::from([("TestFont".into(), FontEntry::stub())]);
        let widths = natural_widths(&table, &fonts, &table.cell_margins);
        let required = 500.0 + table.cell_margins.left + table.cell_margins.right;
        assert!((widths.iter().sum::<f32>() - required).abs() < 0.01);
        assert!((widths[1] - widths[0] * 2.0).abs() < 0.01);
        assert!((widths[2] - widths[0] * 3.0).abs() < 0.01);
    }

    #[test]
    fn autofit_word_width_is_invariant_to_run_formatting() {
        let measure = |parts: &[&str], indent_right| {
            let runs: Vec<crate::model::Run> = parts
                .iter()
                .enumerate()
                .map(|(i, text)| crate::model::Run {
                    text: (*text).into(),
                    font_name: "TestFont".into(),
                    font_size: 10.0,
                    color: Some([0, 0, i as u8]),
                    ..Default::default()
                })
                .collect();
            let mut key_buf = String::new();
            let key = crate::fonts::font_key_buf(&runs[0], &mut key_buf).to_string();
            let fonts = HashMap::from([(key, FontEntry::stub())]);
            let mut doc = crate::docx::parse_bytes(
                include_bytes!("../../tests/fixtures/cases/min-width-run-whole/input.docx"),
                crate::RevisionView::Final,
            )
            .unwrap();
            let Block::Table(mut table) = doc.sections[0].blocks.remove(0) else {
                panic!("table")
            };
            table.rows[0].cells[1].content = vec![Block::Paragraph(crate::model::Paragraph {
                runs,
                indent_right,
                ..Default::default()
            })];
            table.fixed_layout = false;
            min_content_widths(&table, &fonts)[1]
        };
        for indent in [0.0, 11.0] {
            let whole = measure(&["abcdef"], indent);
            assert_eq!(whole, measure(&["abc", "def"], indent));
            assert!(measure(&["abc ", "def"], indent) < whole);
            assert!(measure(&["abc", " def"], indent) < whole);
        }
    }

    #[test]
    fn nested_empty_reserved_cells_keep_explicit_grid_when_it_fits() {
        let bytes = include_bytes!("../../tests/fixtures/cases/case88/input.docx");
        let doc = crate::docx::parse_bytes(bytes, crate::RevisionView::Final).unwrap();
        let Block::Table(outer) = &doc.sections[0].blocks[0] else {
            panic!("outer table")
        };
        let Block::Table(nested) = &outer.rows[0].cells[0].content[0] else {
            panic!("nested table")
        };
        let fonts = HashMap::new();
        let kept = auto_fit_columns(nested, &fonts, Some(260.0), None, 0.0);
        assert_eq!(kept, nested.col_widths);
        assert!(kept.iter().sum::<f32>() > 230.0);
        let squeezed = auto_fit_columns(nested, &fonts, Some(80.0), None, 0.0);
        assert!(squeezed.iter().sum::<f32>() <= 80.001);
    }

    #[test]
    fn nested_explicit_width_uses_fixed_base_even_when_overflowing() {
        let bytes = include_bytes!("../../tests/fixtures/cases/case90/input.docx");
        let doc = crate::docx::parse_bytes(bytes, crate::RevisionView::Final).unwrap();
        let Block::Table(outer) = &doc.sections[0].blocks[0] else {
            panic!("outer table")
        };
        let Block::Table(nested) = &outer.rows[0].cells[0].content[0] else {
            panic!("nested table")
        };
        let fonts = HashMap::new();
        let kept = auto_fit_columns(nested, &fonts, Some(260.0), None, 0.0);
        assert_eq!(kept, auto_fit_columns(nested, &fonts, None, None, 0.0));
        assert!(kept.iter().sum::<f32>() > 230.0);
        let squeezed = auto_fit_columns(nested, &fonts, Some(80.0), None, 0.0);
        assert_eq!(squeezed, kept);
    }

    /// A one-row auto-width table over `grid` whose cells prefer `preferred`.
    fn one_row_table(grid: Vec<f32>, preferred: Vec<f32>) -> Table {
        use crate::model::{CellVAlign, TableCell, TableRow, VMerge};
        Table {
            col_widths: grid,
            rows: vec![TableRow {
                cells: preferred
                    .into_iter()
                    .map(|width| TableCell {
                        width,
                        width_pct: None,
                        content: vec![],
                        borders: Default::default(),
                        shading: None,
                        hatch: None,
                        grid_span: 1,
                        v_merge: VMerge::None,
                        v_align: CellVAlign::Top,
                        text_direction: Default::default(),
                        cell_margins: None,
                        hide_mark: false,
                    })
                    .collect(),
                grid_before: 0,
                height: None,
                height_exact: false,
                is_header: false,
                cant_split: false,
            }],
            table_indent: 0.0,
            table_indent_explicit: false,
            has_table_style: true,
            cell_margins: Default::default(),
            position: None,
            alignment: Default::default(),
            fixed_layout: false,
            auto_width: true,
            width_pct: None,
            grid_inferred: false,
            cells_unsized: false,
            cells_twip_sized: false,
            bidi_visual: false,
            header_first_row: false,
            header_first_col: false,
        }
    }

    #[test]
    fn asymmetric_autofit_parent_keeps_saved_grid_with_nested_table() {
        let mut parent = one_row_table(vec![250.0, 150.0], vec![250.0, 180.0]);
        parent.rows[0].cells[0]
            .content
            .push(Block::Table(one_row_table(vec![230.0], vec![230.0])));
        let fonts = HashMap::new();
        assert_eq!(
            auto_fit_columns(&parent, &fonts, None, None, 0.0),
            vec![250.0, 150.0]
        );
        parent.fixed_layout = true;
        assert!(auto_fit_columns(&parent, &fonts, None, None, 0.0)[0] < 240.0);
        parent.fixed_layout = false;
        parent.auto_width = false;
        assert!(auto_fit_columns(&parent, &fonts, None, None, 0.0)[0] < 240.0);
    }

    #[test]
    fn asymmetric_autofit_keeps_grid_with_uniform_stale_cell_widths() {
        let mut table = one_row_table(vec![250.0, 150.0], vec![300.0, 300.0]);
        let fonts = HashMap::new();
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None, 0.0),
            vec![250.0, 150.0]
        );
        table.col_widths = vec![235.6, 232.25];
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None, 0.0),
            vec![235.6, 232.25]
        );
        table.col_widths = vec![250.0, 150.0];
        table.fixed_layout = true;
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None, 0.0),
            vec![200.0, 200.0]
        );
        table.fixed_layout = false;
        table.auto_width = false;
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None, 0.0),
            vec![200.0, 200.0]
        );
        table.auto_width = true;
        table.rows[0].cells[0].width = 250.0;
        table.rows[0].cells[1].width = 150.0;
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None, 0.0),
            vec![250.0, 150.0]
        );
    }

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
    fn test_cell_span_width_start_past_end() {
        // Missing tblGrid (empty widths) or gridSpan overrun must not panic.
        assert_eq!(cell_span_width(&[], 1, 1), 0.0);
        let widths = vec![100.0, 200.0];
        assert_eq!(cell_span_width(&widths, 5, 2), 0.0);
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
