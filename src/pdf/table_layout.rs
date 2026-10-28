use std::collections::HashMap;

use crate::fonts::{FontEntry, font_key_buf};

use crate::model::{
    Alignment, Block, CellMargins, HorizontalPosition, Table, TextDirection, VMerge,
    VerticalPosition, WrapType,
};

use super::RenderContext;
use super::header_footer::substitute_hf_runs;
use super::helpers::{effective_space_after, effective_space_before};
use super::layout::{
    TextLine, build_paragraph_lines, build_tabbed_line, east_asian_leading, is_text_empty,
    run_line_metrics, word_width_for_run,
};
use super::resolve_line_h;

pub(super) fn cell_span_width(col_widths: &[f32], grid_col: usize, span: usize) -> f32 {
    // Clamp the start too: malformed tables (missing tblGrid, gridSpan
    // overrun) can push grid_col past the grid — yield 0 instead of panicking.
    let start = grid_col.min(col_widths.len());
    col_widths[start..col_widths.len().min(grid_col + span)]
        .iter()
        .sum()
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
    let Some(pct) = table.width_pct else { return };
    let basis = if compat_mode >= 15 {
        available_w
    } else {
        available_w + table.first_cell_left_margin() + table.last_cell_right_margin()
    };
    // Word caps a table's width at the text column width.
    let target = basis * pct.min(1.0);
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
                        || p.runs.iter().any(|r| !r.text.trim().is_empty() || r.inline_image.is_some())
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
    let mut min = vec![0.0f32; table.col_widths.len()];
    for row in &table.rows {
        for (grid_col, span, cell) in row.grid_cells() {
            if grid_col >= min.len() || span > 1 {
                continue;
            }
            let pad = cell.cell_margins.as_ref().unwrap_or(&table.cell_margins);
            let pad = pad.left + pad.right;
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
                    min[grid_col] = min[grid_col].max(w + pad);
                }
            }
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

/// Height of a paragraph's text content, matching the layout computation in
/// `compute_row_layouts`. Empty paragraphs (no lines) still occupy one line
/// height unless they carry an explicit `content_height` (e.g. from an image).
pub(super) fn para_block_height(p: &CellParagraphLayout) -> f32 {
    if p.lines.is_empty() {
        if p.paragraph_mark_vanish {
            0.0
        } else {
            let h = if p.content_height > 0.0 { p.content_height } else { p.line_h };
            h + super::helpers::border_band(p.bottom_border.as_ref())
        }
    } else {
        cell_lines_h(p, 0..p.lines.len())
    }
}

/// Height of a cell paragraph's lines `range`, each at its own pitch
/// (`size_lines_by_own_runs`).
pub(super) fn cell_lines_h(p: &CellParagraphLayout, range: std::ops::Range<usize>) -> f32 {
    let bottom_band = if range.end == p.lines.len() {
        super::helpers::border_band(p.bottom_border.as_ref())
    } else { 0.0 };
    p.lines[range]
        .iter()
        .map(|l| super::layout::inline_line_advance(l, p.line_h, p.picture_metrics()))
        .sum::<f32>() + bottom_band
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
                        let text = if run.caps {
                            std::borrow::Cow::Owned(run.text.to_uppercase())
                        } else {
                            std::borrow::Cow::Borrowed(&run.text)
                        };
                        para_w +=
                            word_width_for_run(entry, run, &text, fs, run.kerns_at(fs), 0.0, 1.0);
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
    let ncols = table.col_widths.len();
    let cm = &table.cell_margins;
    let mut min_widths = vec![0.0f32; ncols];
    for row in &table.rows {
        for (grid_col, span, cell) in row.grid_cells() {
            if grid_col >= ncols || span > 1 {
                continue;
            }
            if cell.text_direction != TextDirection::LrTb {
                continue;
            }
            let ecm = cell.cell_margins.as_ref().unwrap_or(cm);
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
                    let text = if run.caps {
                        std::borrow::Cow::Owned(run.text.to_uppercase())
                    } else {
                        std::borrow::Cow::Borrowed(&run.text)
                    };
                    let fs = run.font_size;
                    // Same break opportunities as line layout: a CJK sentence holds
                    // no spaces but may wrap after any ideograph.
                    for (spaces, word) in super::layout::split_preserving_spaces(&text) {
                        let ww =
                            word_width_for_run(entry, run, word, fs, run.kerns_at(fs), 0.0, 1.0);
                        let continues = !table.fixed_layout
                            && spaces == 0
                            && prev_char
                                .zip(word.chars().next())
                                .is_some_and(|(a, b)| !super::layout::breaks_between(a, b));
                        joined_width = if continues { joined_width + ww } else { ww };
                        prev_char = word.chars().last();
                        let required = (joined_width + inset).max(h_pad);
                        min_widths[grid_col] = min_widths[grid_col].max(required);
                    }
                    if run.text.ends_with(char::is_whitespace) {
                        prev_char = None;
                    }
                }
            }
        }
    }
    min_widths
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
) -> Vec<f32> {
    let ncols = table.col_widths.len();
    if ncols == 0 {
        return table.col_widths.clone();
    }

    // A nested table sized in twips (tblW dxa) keeps its grid like a top-level
    // one: a survey's 465pt nested table inside an auto-width host fills it
    // in Word, where content-fitting made it half as wide.
    if available_width.is_some()
        && !table.auto_width
        && !table.fixed_layout
        && !table.grid_inferred
    {
        return auto_fit_columns(table, fonts, None, None);
    }

    let cm = &table.cell_margins;
    // A picture never wraps: its column is at least that wide (advocacy's
    // nested picture table collapsed to its padding, the picture clipped away).
    let min_widths: Vec<f32> = min_content_widths(table, fonts)
        .iter()
        .zip(hard_content_widths(table, false))
        .map(|(text, hard)| text.max(hard))
        .collect();

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
                    .map(|i| natural_widths[i].max(min_widths[i].max(min_cell)).max(spacer[i]))
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

pub(super) struct CellFloatingImageLayout {
    pub(super) pdf_name: String,
    pub(super) display_width: f32,
    pub(super) display_height: f32,
    pub(super) h_offset: f32,
    pub(super) v_offset: f32,
    /// In-plane rotation in degrees (OOXML clockwise). Cell-anchored floats must
    /// carry this just like body floats, else e.g. a 90°-rotated vertical label
    /// renders horizontally.
    pub(super) rotation_deg: f32,
    /// wp:anchor relativeHeight, so a cell picture can paint over a
    /// connector/textbox anchored in the same paragraph (annotation #241).
    pub(super) z_index: u32,
    /// The picture's `docPr@descr` and decorative flag, for tagging.
    pub(super) alt: Option<String>,
    pub(super) decorative: bool,
}

impl CellFloatingImageLayout {
    /// (alt, decorative), for `cell_figure`.
    pub(super) fn tagging(&self) -> (Option<&str>, bool) {
        (self.alt.as_deref(), self.decorative)
    }
}

#[derive(Default)]
pub(super) struct CellParagraphLayout {
    pub(super) lines: Vec<TextLine>,
    pub(super) line_h: f32,
    pub(super) font_size: f32,
    pub(super) ascender_ratio: f32,
    /// Picture-line ascent; exact spacing keeps the fixed baseline and pitch.
    pub(super) image_line_ascent: f32,
    pub(super) clip_inline_pictures: bool,
    pub(super) descender_ratio: f32,
    pub(super) font_substituted: bool,
    pub(super) alignment: Alignment,
    pub(super) space_before: f32,
    /// Collapsed spacing less the previous page's already consumed after gap.
    pub(super) continuation_space_before: f32,
    pub(super) indent_left: f32,
    pub(super) indent_right: f32,
    pub(super) indent_hanging: f32,
    pub(super) indent_first_line: f32,
    pub(super) text_hanging: f32,
    /// `list_label::label_shift`: where a right/centre-aligned label starts.
    pub(super) label_shift: f32,
    /// Extra left indent from wrapSquare/Tight floating images in this paragraph.
    /// Text lines are laid out narrower and rendered further right to avoid the image.
    pub(super) float_indent_left: f32,
    /// The first `beside_lines` lines start `beside_indent` further right,
    /// level with a left floating nested table.
    pub(super) beside_indent: f32,
    pub(super) beside_lines: usize,
    /// `w:keepLines`: the paragraph never splits over a row break.
    pub(super) keep_lines: bool,
    /// Word 2013+ layout (compatibilityMode 15): a row break keeps two lines
    /// of this paragraph on each side; earlier layout lets a row start with one.
    pub(super) two_line_widow: bool,
    pub(super) list_label: String,
    /// The label by font (`list_label::label_pieces`), empty without one.
    pub(super) label_pieces: Vec<(Option<String>, String)>,
    pub(super) label_font_size: f32,
    /// For L/LI tagging inside the cell.
    pub(super) list_item: Option<crate::model::ListItem>,
    /// Tagged as the Formula itself (`sole_formula`).
    pub(super) sole_formula: bool,
    pub(super) label_color: Option<[u8; 3]>,
    pub(super) image_name: Option<String>,
    /// The picture's `docPr@descr` and decorative flag, for tagging.
    pub(super) image_alt: Option<String>,
    pub(super) image_decorative: bool,
    pub(super) image_width: f32,
    pub(super) image_height: f32,
    pub(super) image_stroke_color: Option<[u8; 3]>,
    pub(super) image_stroke_width: f32,
    pub(super) image_shadow: Option<crate::model::ImageShadow>,
    pub(super) image_shadow_xobj: Option<String>,
    pub(super) image_glow: Option<crate::model::ImageGlow>,
    pub(super) image_glow_xobj: Option<String>,
    pub(super) image_clip: Option<crate::model::ShapeGeometry>,
    pub(super) content_height: f32,
    pub(super) bottom_border: Option<crate::model::ParagraphBorder>,
    pub(super) paragraph_mark_vanish: bool,
    pub(super) floating_images: Vec<CellFloatingImageLayout>,
    pub(super) space_after: f32,
    pub(super) has_textboxes: bool,
    pub(super) has_connectors: bool,
}

impl CellParagraphLayout {
    pub(super) fn tag_kind(&self) -> &'static str {
        if self.sole_formula { "Formula" } else { "P" }
    }

    /// (ascent, descent) that size this paragraph's picture lines.
    pub(super) fn picture_metrics(&self) -> (f32, f32) {
        (
            self.image_line_ascent,
            self.font_size * self.descender_ratio,
        )
    }
}

// Synthetic Word controls switch between 18.75 pt and 18.70 pt of free
// right strip; the transition is unchanged for 8–12 pt marks and no borders.
const MIN_NESTED_FLOAT_STRIP: f32 = 18.75;

pub(super) enum CellContentItem {
    Paragraph(CellParagraphLayout),
    /// A table inside the cell, laid out (and drawn) at `col_widths`. A row
    /// split breaks between its rows (the cursor's `line` counts nested rows)
    /// or inside one that can split (`CellCursor::nested`).
    NestedTable {
        col_widths: Vec<f32>,
        rows: Vec<RowLayout>,
        space_before: f32,
        floating_offset: Option<f32>,
        floating_x_offset: Option<f32>,
    },
}

pub(super) struct CellLayout {
    pub(super) items: Vec<CellContentItem>,
    /// The cell's margins (its own or the table's), border bands included.
    pub(super) cm: CellMargins,
    pub(super) total_height: f32,
    /// The space after the last paragraph that `total_height` includes; the
    /// chunk of a split row that finishes the cell charges it too.
    pub(super) trailing_space_after: f32,
    pub(super) text_direction: TextDirection,
}

pub(super) struct RowLayout {
    pub(super) height: f32,
    pub(super) cells: Vec<CellLayout>,
    /// The room a row needs on a page to break across it: None for cantSplit
    /// and exact-height rows, which Word never breaks; an at-least row's
    /// trHeight with its cell margins (croatian_grant's 15pt rows of five
    /// bullets break at the foot of a page, stem_partnerships' 76.55pt row
    /// and traditional_skills' form rows move whole).
    pub(super) split_min: Option<f32>,
}

/// When provided, field codes in header/footer table runs are substituted with
/// their resolved values before layout.
pub(super) struct HfSubstitution<'a> {
    pub(super) page_num: usize,
    pub(super) total_pages: usize,
    pub(super) styleref_values: &'a HashMap<String, String>,
    pub(super) page_num_format: Option<&'a str>,
}

/// A one-row, text-anchored table floated at the cell's left margin with room
/// for text beside it: its vertical extent below its anchor.
fn registration_table_extent(table: &Table, cell_text_w: f32, ctx: &RenderContext) -> Option<f32> {
    let pos = table.position.as_ref()?;
    if table.rows.len() != 1
        || pos.v_anchor != "text"
        || pos.h_anchor != "margin"
        || !matches!(pos.h_position, HorizontalPosition::Offset(x) if x == 0.0)
        || table.table_indent != 0.0
        || table.alignment != crate::model::TableAlignment::Left
    {
        return None;
    }
    let mut widths = auto_fit_columns(table, ctx.fonts, Some(cell_text_w), None);
    apply_nested_pct_width(table, &mut widths, cell_text_w);
    let strip = cell_text_w - widths.iter().sum::<f32>() - pos.left_from_text - pos.right_from_text;
    if strip < MIN_NESTED_FLOAT_STRIP {
        return None;
    }
    let height = compute_row_layouts(table, &widths, ctx, None)
        .iter()
        .map(|r| r.height)
        .sum::<f32>();
    Some(height + pos.v_offset_pt.max(0.0) + pos.top_from_text + pos.bottom_from_text)
}

/// A paragraph whose height is only its lines and spacing.
fn plain_spacer(p: &crate::model::Paragraph) -> bool {
    !(p.frame_props.is_some()
        || p.paragraph_mark_vanish
        || p.image.is_some()
        || p.inline_chart.is_some()
        || !p.smartart.is_empty()
        || !p.floating_images.is_empty()
        || !p.textboxes.is_empty()
        || !p.connectors.is_empty()
        || p.contextual_spacing
        || p.space_before_auto
        || p.space_after_auto
        || p.runs.iter().any(|r| {
            r.is_tab
                || r.is_line_break
                || r.inline_image.is_some()
                || r.field_code.is_some()
                || r.math.is_some()
                || r.vanish
                || r.footnote_id.is_some()
                || r.endnote_id.is_some()
        }))
}

fn spacer_line_h(p: &crate::model::Paragraph, ctx: &RenderContext) -> f32 {
    let (fs, lhr, _) = super::layout::tallest_run_metrics(&p.runs, ctx.fonts);
    resolve_line_h(p.line_spacing.unwrap_or(ctx.doc_line_spacing), fs, lhr)
}

// Interleaved empty anchors contribute flow height; their floating tables do
// not, provided that the accumulated spacer lines clear every table's extent.
fn interleaved_spacers_clear_nested_registration(
    cell: &crate::model::TableCell,
    block_idx: usize,
    cell_text_w: f32,
    ctx: &RenderContext,
    hf_sub: Option<&HfSubstitution>,
) -> bool {
    if hf_sub.is_some() || cell.text_direction != TextDirection::LrTb {
        return false;
    }
    let mut flow = 0.0f32;
    let mut occupied = 0.0f32;
    let mut prev_after = 0.0f32;
    let mut tables = 0;
    for (index, block) in cell.content.iter().enumerate() {
        match block {
            Block::Table(table) => {
                tables += 1;
                if tables > 2 || (tables == 2 && flow <= 0.0) {
                    return false;
                }
                let Some(extent) = registration_table_extent(table, cell_text_w, ctx) else {
                    return false;
                };
                occupied = occupied.max(flow + prev_after + extent);
            }
            Block::Paragraph(p) => {
                if !plain_spacer(p) {
                    return false;
                }
                flow += prev_after.max(p.space_before);
                if p.runs.iter().any(|r| !r.text.trim().is_empty()) {
                    return tables == 2 && block_idx < index && flow >= occupied;
                }
                flow += spacer_line_h(p, ctx);
                prev_after = p.space_after;
            }
        }
    }
    false
}

// Leading margin-anchored registration tables may precede visible cell text
// without contributing to flow height when measured empty paragraphs already
// clear their entire extent. Keep the old narrow path for all other layouts.
fn spacers_clear_nested_registration(
    cell: &crate::model::TableCell,
    block_idx: usize,
    cell_text_w: f32,
    ctx: &RenderContext,
    hf_sub: Option<&HfSubstitution>,
) -> bool {
    if hf_sub.is_some() || cell.text_direction != TextDirection::LrTb {
        return false;
    }
    let tables: Vec<&Table> = cell
        .content
        .iter()
        .map_while(|b| match b {
            Block::Table(t) => Some(t),
            Block::Paragraph(_) => None,
        })
        .collect();
    if tables.is_empty() || tables.len() > 2 || block_idx >= tables.len() {
        return false;
    }
    let mut occupied = 0.0f32;
    for table in &tables {
        let Some(extent) = registration_table_extent(table, cell_text_w, ctx) else {
            return false;
        };
        occupied += extent;
    }
    let mut empty_space = 0.0f32;
    let mut prev_after = 0.0f32;
    let mut visible_seen = false;
    for block in &cell.content[tables.len()..] {
        let Block::Paragraph(p) = block else {
            return false;
        };
        if !plain_spacer(p) {
            return false;
        }
        if visible_seen {
            continue;
        }
        empty_space += prev_after.max(p.space_before);
        if p.runs.iter().any(|r| !r.text.trim().is_empty()) {
            if empty_space < occupied {
                return false;
            }
            visible_seen = true;
        } else {
            empty_space += spacer_line_h(p, ctx);
            prev_after = p.space_after;
        }
    }
    visible_seen
}

/// The single-line width of a paragraph of plain visible text with its
/// indents; None for anything else or a missing font.
fn plain_line_width(p: &crate::model::Paragraph, fonts: &HashMap<String, FontEntry>) -> Option<f32> {
    if p.frame_props.is_some()
        || p.list_item.is_some()
        || p.image.is_some()
        || !p.floating_images.is_empty()
        || !p.textboxes.is_empty()
        || !p.runs.iter().any(|r| !r.text.trim().is_empty())
        || p.runs.iter().any(|r| {
            r.is_tab
                || r.is_line_break
                || r.inline_image.is_some()
                || r.field_code.is_some()
                || r.math.is_some()
                || r.vanish
        })
    {
        return None;
    }
    let mut width = p.indent_left + p.indent_right + p.indent_first_line.max(0.0);
    let mut key = String::new();
    for run in &p.runs {
        let font = fonts.get(font_key_buf(run, &mut key))?;
        let text = if run.caps {
            std::borrow::Cow::Owned(run.text.to_uppercase())
        } else {
            std::borrow::Cow::Borrowed(run.text.as_str())
        };
        width += word_width_for_run(font, run, &text, run.font_size, run.kerns_at(run.font_size), 0.0, 1.0);
    }
    Some(width)
}

fn plain_text_fits_nested_strip(
    cell: &crate::model::TableCell,
    idx: usize,
    table: &Table,
    widths: &[f32],
    cell_text_w: f32,
    ctx: &RenderContext,
) -> bool {
    let Some(pos) = &table.position else {
        return false;
    };
    if !matches!(pos.h_anchor, "margin" | "column")
        || !matches!(pos.h_position, HorizontalPosition::Offset(x) if x == 0.0)
        || pos.v_offset_pt < 0.0
        || table.table_indent != 0.0
        || table.alignment != crate::model::TableAlignment::Left
    {
        return false;
    }
    let Some(Block::Paragraph(p)) = cell.content.get(idx + 1) else {
        return false;
    };
    if p.runs.iter().any(|r| r.caps) {
        return false;
    }
    let strip = cell_text_w - widths.iter().sum::<f32>() - pos.right_from_text;
    let Some(line_w) = plain_line_width(p, ctx.fonts) else {
        return false;
    };
    // A cell's sole table lets text of any length, over several paragraphs,
    // wrap beside it (advocacy's picture stories); beside another table only
    // a lone paragraph fitting one line is known to.
    let sole_table = cell
        .content
        .iter()
        .enumerate()
        .all(|(i, b)| i == idx || matches!(b, Block::Paragraph(_)));
    strip >= MIN_NESTED_FLOAT_STRIP
        && (sole_table || (cell.content.len() == idx + 2 && line_w <= strip))
}

// A single cell heading can precede positive-offset floating tables even
// when the tables are serialized before its anchor paragraph.
fn text_fits_above_nested_floats(
    cell: &crate::model::TableCell,
    cell_text_w: f32,
    fonts: &HashMap<String, FontEntry>,
    doc_line_spacing: crate::model::LineSpacing,
    hf_sub: Option<&HfSubstitution>,
) -> bool {
    if hf_sub.is_some()
        || cell.text_direction != TextDirection::LrTb
        || !matches!(cell.v_align, crate::model::CellVAlign::Top)
    {
        return false;
    }
    let Some((Block::Paragraph(p), prefix)) = cell.content.split_last() else {
        return false;
    };
    if prefix.is_empty() {
        return false;
    }
    let Some(width) = plain_line_width(p, fonts) else {
        return false;
    };
    let mut first_top = f32::INFINITY;
    for block in prefix {
        let Block::Table(t) = block else {
            return false;
        };
        let Some(pos) = &t.position else {
            return false;
        };
        if t.rows.len() != 1
            || pos.v_anchor != "text"
            || pos.h_anchor != "margin"
            || pos.v_offset_pt <= 0.0
            || t.table_indent != 0.0
            || !matches!(pos.h_position, HorizontalPosition::Offset(x) if x == 0.0)
            || t.alignment != crate::model::TableAlignment::Left
        {
            return false;
        }
        first_top = first_top.min(pos.v_offset_pt - pos.top_from_text);
    }
    let (fs, lhr, _) = super::layout::tallest_run_metrics(&p.runs, fonts);
    let height = p.space_before.max(0.0)
        + resolve_line_h(p.line_spacing.unwrap_or(doc_line_spacing), fs, lhr)
        + p.space_after.max(0.0);
    width <= cell_text_w && height <= first_top
}

pub(super) fn compute_row_layouts(
    table: &Table,
    col_widths: &[f32],
    ctx: &RenderContext,
    hf_sub: Option<&HfSubstitution>,
) -> Vec<RowLayout> {
    let cm = &table.cell_margins;
    let mut layouts: Vec<RowLayout> = table
        .rows
        .iter()
        .map(|row| {
            let mut max_h: f32 = 0.0;
            let row_keeps_next = row.cells.iter().flat_map(|cell| &cell.content)
                .any(|block| matches!(block, Block::Paragraph(p) if p.keep_next));
            // Word uses the largest top/bottom cell margins across a row,
            // even when the cell carrying the margin has only a hidden mark.
            // Left/right margins remain specific to each cell.
            let (row_top, row_bottom) = row.cells.iter().fold((0.0f32, 0.0f32), |(top, bottom), cell| {
                let margins = cell.cell_margins.as_ref().unwrap_or(cm);
                (top.max(margins.top - cell.borders.top.band() / 2.0),
                 bottom.max(margins.bottom - cell.borders.bottom.band() / 2.0))
            });
            let cells: Vec<CellLayout> = row
                .grid_cells()
                .map(|(grid_col, span, cell)| {
                    let span_w = cell_span_width(col_widths, grid_col, span);
                    // For auto-fit tables the resolved grid width is what the
                    // renderer draws borders and content at, so the layout must
                    // use the same width — a larger tcW preference otherwise wraps
                    // text past the drawn cell border (#115). Fixed-layout tables
                    // keep honoring the cell's preferred width (their gridCol can
                    // be narrower than Word's effective column).
                    let col_w = if table.fixed_layout {
                        span_w.max(cell.width)
                    } else {
                        span_w
                    };
                    if cell.v_merge == VMerge::Continue {
                        return CellLayout {
                            items: vec![],
                            cm: *cm,
                            total_height: 14.4,
                            trailing_space_after: 0.0,
                            text_direction: TextDirection::LrTb,
                        };
                    }

                    let effective_margins = CellMargins {
                        top: row_top + cell.borders.top.band() / 2.0,
                        bottom: row_bottom + cell.borders.bottom.band() / 2.0,
                        ..*cell.cell_margins.as_ref().unwrap_or(cm)
                    };
                    let ecm = &effective_margins;
                    let is_rotated = cell.text_direction != TextDirection::LrTb;
                    let cell_text_w = if is_rotated {
                        10000.0
                    } else {
                        (col_w - ecm.left - ecm.right).max(0.0)
                    };
                    let mut total_h: f32 = ecm.top + ecm.bottom;
                    let mut floating_extent: f32 = 0.0;
                    let mut left_float_bottom: f32 = 0.0;
                    let mut left_float_top: f32 = 0.0;
                    let mut left_float_indent: f32 = 0.0;
                    let mut first_positive_float = false;
                    let mut left_float_right_gap: f32 = f32::INFINITY;
                    let mut max_rotated_line_w: f32 = 0.0;
                    let mut items: Vec<CellContentItem> = Vec::new();
                    let mut prev_space_after = 0.0f32;
                    let mut para_idx = 0usize;
                    let mut prev_was_nested_table = false;
                    let mut negative_anchor_shift = 0.0f32;
                    let mut final_float_mark_hidden = false;
                    let top_aligned = matches!(cell.v_align, crate::model::CellVAlign::Top);
                    let text_above_floats =
                        text_fits_above_nested_floats(cell, cell_text_w, ctx.fonts, ctx.doc_line_spacing, hf_sub);
                    let centered = matches!(cell.v_align, crate::model::CellVAlign::Center);

                    let block_count = cell.content.len();
                    // Modern Word drops the relocated anchor line only when
                    // its empty paragraph has expanded automatic line spacing.
                    // Single/exact/at-least spacing and legacy modes retain it.
                    let suppress_negative_mark = |idx: usize, offset: Option<f32>| {
                        ctx.compat_mode >= 15 && offset.is_some_and(|y| y < 0.0)
                            && idx.checked_sub(1).and_then(|i| cell.content.get(i))
                                .is_some_and(|b| matches!(b, Block::Paragraph(p)
                                    if is_text_empty(&p.runs)
                                        && matches!(p.line_spacing.unwrap_or(ctx.doc_line_spacing),
                                            crate::model::LineSpacing::Auto(factor) if factor > 1.0)))
                    };
                    for (block_idx, block) in cell.content.iter().enumerate() {
                        match block {
                            Block::Paragraph(para) => {
                                // An HTML rule (o:hr) alone in a hideMark cell keeps a
                                // 5.25pt line in Word (advocacy's story separators:
                                // 0.75pt margins, the 0.5pt rule 2.5pt below the top),
                                // where the hidden mark would otherwise collapse it.
                                // ponytail: one measurement; Word's rule for the
                                // shape's line box is unknown.
                                // Word 2013+ layout (compatibilityMode 15) keeps the
                                // mark's full text line instead: maine's separator
                                // rows step one line, 5.25pt once saved as mode 14
                                // (Word probes).
                                let rule_only = para.horizontal_rule.is_some() && is_text_empty(&para.runs);
                                let rule_keeps_line = rule_only && ctx.compat_mode >= 15;
                                let para_content_h = if cell.hide_mark && rule_only && !rule_keeps_line {
                                    5.25
                                } else {
                                    para.content_height
                                };
                                let nested_float_clearance = if left_float_right_gap < MIN_NESTED_FLOAT_STRIP
                                    && para.runs.iter().all(|r| r.text.trim().is_empty() && !r.is_tab && !r.is_line_break)
                                {
                                    let (fs,lhr,_) = super::layout::tallest_run_metrics(&para.runs,ctx.fonts);
                                    let pitch=resolve_line_h(para.line_spacing.unwrap_or(ctx.doc_line_spacing),fs,lhr);
                                    if !first_positive_float || total_h + pitch > left_float_top { (left_float_bottom-total_h).max(0.0) } else { 0.0 }
                                } else { 0.0 };
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
                                } else if let Some(marked) = ctx.with_note_marks(&para.runs) {
                                    // A note reference mark's run is empty until it
                                    // shows its note's number, as in body paragraphs.
                                    substituted = marked;
                                    &substituted
                                } else {
                                    &para.runs
                                };
                                // A cell paragraph containing only ordinary
                                // padding spaces has the paragraph mark's line
                                // metrics in Word, not the spaces' run format.
                                let padding_mark;
                                let runs = if para_content_h == 0.0
                                    && runs.iter().any(|run| !run.text.is_empty())
                                    && runs.iter().all(|run| !run.is_tab
                                        && !run.is_line_break && run.math.is_none()
                                        && run.inline_image.is_none()
                                        && run.text.bytes().all(|b| b == b' '))
                                {
                                    padding_mark = crate::model::Run {
                                        font_size: para.paragraph_mark_font_size.unwrap_or(10.0),
                                        font_name: para.paragraph_mark_font_name.clone().unwrap_or_default(),
                                        ..Default::default()
                                    };
                                    std::slice::from_ref(&padding_mark)
                                } else { runs };
                                // Size the cell line from the tallest run, not the
                                // first — a small leading run (e.g. padding spaces)
                                // must not pull the baseline up. Math runs are
                                // clamped inside tallest_run_metrics so a header
                                // cell leading with math doesn't balloon the row.
                                // Cell metrics come from the first run carrying
                                // real text — a leading padding-spaces run must
                                // not set the baseline (Word sizes the line from
                                // the content run; observed in header cells like
                                // "                g" where g is larger).
                                let metric_run = runs
                                    .iter()
                                    .find(|r| {
                                        r.is_tab
                                            || r.text.is_empty()
                                            || !r.text.trim().is_empty()
                                    })
                                    .or(runs.first());
                                // A tab-only empty cell paragraph is sized by its
                                // paragraph mark, as in body text. Do not let an
                                // inherited tab font enlarge a smaller explicit mark.
                                let mark_run;
                                let metric_run = if !runs.is_empty()
                                    && runs.iter().all(|r| r.is_tab)
                                    && para.tab_stops.iter().all(|tab| tab.leader.is_none())
                                    && let Some(size) = para.paragraph_mark_font_size
                                    && let Some(name) = &para.paragraph_mark_font_name
                                {
                                    mark_run = crate::model::Run {
                                        font_size: size,
                                        font_name: name.clone(),
                                        ..Default::default()
                                    };
                                    Some(&mark_run)
                                } else {
                                    metric_run
                                };
                                let font_size = metric_run.map_or(12.0, |r| r.font_size);
                                let mut kb0 = String::new();
                                let metric_font = metric_run
                                    .map(|r| font_key_buf(r, &mut kb0).to_owned())
                                    .and_then(|k| ctx.fonts.get(&k));
                                let (tallest_lhr, tallest_ar) = metric_run
                                    .zip(metric_font)
                                    .map_or((None, None), |(r, e)| run_line_metrics(e, &r.text));
                                let effective_ls =
                                    para.line_spacing.unwrap_or(ctx.doc_line_spacing);
                                let line_h =
                                    resolve_line_h(effective_ls, font_size, tallest_lhr);
                                // A negative floating table moves its preceding empty
                                // anchor mark below itself. The mandatory end-of-cell
                                // mark does not replace that anchor's line metrics.
                                let relocated_anchor_line = (block_idx + 1 == block_count
                                    && nested_float_clearance > 0.0
                                    && is_text_empty(runs)
                                    && matches!(items.last(), Some(CellContentItem::NestedTable {
                                        floating_offset: Some(offset), ..
                                    }) if *offset < 0.0)
                                    && block_idx.checked_sub(2).and_then(|i| cell.content.get(i))
                                        .is_some_and(|b| matches!(b, Block::Paragraph(p)
                                            if is_text_empty(&p.runs) && p.content_height == 0.0
                                                && p.image.is_none() && !p.paragraph_mark_vanish)))
                                    .then(|| items.get(items.len().saturating_sub(2)))
                                    .flatten()
                                    .and_then(|item| match item {
                                        CellContentItem::Paragraph(p) => Some(p.line_h),
                                        _ => None,
                                    });
                                let line_h = relocated_anchor_line.unwrap_or(line_h);

                                // Auto-spaced cell lines snap to the section's line grid
                                // like body lines: japanese_medical's ten empty cell
                                // paragraphs step 18pt (the grid), not their 15pt
                                // natural height; chinese_student's at-least-0 cell
                                // lines keep their own height.
                                let grid_pitch = ctx.cell_grid_pitch.get();
                                let grid_snapped = para.snap_to_grid
                                    && grid_pitch > 0.0
                                    && matches!(effective_ls, crate::model::LineSpacing::Auto(_));
                                let line_h = if grid_snapped {
                                    super::layout::grid_snapped_line_h(
                                        runs,
                                        ctx.fonts,
                                        effective_ls,
                                        line_h,
                                        grid_pitch,
                                    )
                                } else if is_text_empty(runs) && para_content_h == 0.0 {
                                    line_h + super::mark_position_stretch(para, effective_ls, false, ctx.fonts)
                                } else {
                                    line_h
                                };

                                // A numbering label taller than the text raises the
                                // first line (see `label_boosted_line_h`); CV's 9pt
                                // Symbol bullets on 9pt Georgia add 0.8pt per item.
                                // ponytail: folded into space_before, which equals the
                                // baseline drop at single spacing; split them if a
                                // multiple-spaced list cell ever drifts.
                                let label_extra = super::label_boosted_line_h(
                                    para,
                                    ctx.fonts,
                                    line_h,
                                    effective_ls,
                                    font_size,
                                    tallest_lhr,
                                    tallest_ar,
                                ) - line_h;
                                // Contextual spacing applies between a cell's
                                // paragraphs as in the body (erasmus_plus: Normal's
                                // 12pt after "Faculty/Department" goes).
                                let space_after = effective_space_after(
                                    para,
                                    super::block_para(&cell.content, block_idx + 1),
                                );
                                // hideMark drops the cell's trailing empty mark with
                                // its spacing; the paragraph above keeps its space
                                // after (Word probes).
                                let mark_hidden = cell.hide_mark
                                    && block_idx == block_count - 1
                                    && para_content_h == 0.0
                                    && is_text_empty(runs)
                                    && !para.paragraph_mark_vanish
                                    && !rule_keeps_line;
                                // A cell drops HTML auto spacing at its edges.
                                let space_before = if mark_hidden {
                                    prev_space_after
                                } else if para_idx > 0 {
                                    let prev = super::block_para(&cell.content, block_idx - 1);
                                    f32::max(prev_space_after, effective_space_before(para, prev))
                                } else if para.space_before_auto {
                                    0.0
                                } else {
                                    para.space_before
                                } + label_extra;
                                let space_before = space_before.max(nested_float_clearance);
                                total_h += space_before;

                                // The first baseline sits this far (per em) below the
                                // cell top: the ascent, line gap included, as in body
                                // text; one em for an East Asian font, whose 1.3×
                                // leading Word keeps out of the cell's top
                                // (japanese_interlibrary: 11pt MS Mincho 11.0).
                                let east_asian = metric_run
                                    .zip(metric_font)
                                    .is_some_and(|(r, e)| east_asian_leading(e, &r.text));
                                let ascender_ratio =
                                    if east_asian { 1.0 } else { tallest_ar.unwrap_or(0.75) };
                                // Win-path metrics identity: line_h_ratio −
                                // ascender_ratio = usWinDescent/units. Fallback
                                // 0.2 pairs with the 1.2 default line ratio so
                                // standard fonts get zero trailing leading.
                                let descender_ratio = if runs.iter().any(|r| r.inline_image.is_some())
                                    && runs.iter().all(|r| !r.is_tab && r.text.trim().is_empty()) {
                                    0.0
                                } else {
                                    tallest_lhr.zip(tallest_ar)
                                        .map(|(lh, ar)| (lh - ar).max(0.0)).unwrap_or(0.2)
                                };
                                let exact_spacing =
                                    matches!(effective_ls, crate::model::LineSpacing::Exact(_));
                                let image_line_ascent = if exact_spacing {
                                    super::layout::runs_max_image_h(runs).max(font_size * ascender_ratio)
                                } else {
                                    font_size * ascender_ratio
                                };
                                let font_substituted =
                                    metric_font.is_some_and(|e| e.is_substituted);

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
                                // Lines level with a left floating nested table
                                // keep clear of it, the rest run full width, as
                                // beside a wrapping picture.
                                let beside_float = total_h < left_float_bottom
                                    && total_h + line_h > left_float_top
                                    && left_float_right_gap >= MIN_NESTED_FLOAT_STRIP;
                                let mut beside_indent = if beside_float {
                                    (left_float_indent - float_indent_left).max(0.0)
                                } else {
                                    0.0
                                };
                                let mut beside_lines = if beside_indent > 0.0 {
                                    ((left_float_bottom - total_h) / line_h.max(1.0)).ceil().max(1.0) as usize
                                } else {
                                    0
                                };
                                // ponytail: tabbed lines take no per-line widths, so
                                // a tabbed paragraph stays indented throughout.
                                let float_indent_left = if beside_lines > 0 && runs.iter().any(|r| r.is_tab) {
                                    beside_lines = 0;
                                    float_indent_left + std::mem::take(&mut beside_indent)
                                } else {
                                    float_indent_left
                                };

                                // An inline-image paragraph whose only other run
                                // is a trailing line break (Word's logo-in-cell
                                // idiom) is NOT text-empty, but its row height must
                                // come from the image's content_height, not a stray
                                // text line — otherwise the cell collapses and a
                                // tall header logo fails to push the body down.
                                let image_only = para_content_h > 0.0
                                    && runs.iter().all(|r| r.text.is_empty() && !r.is_tab);
                                let mut break_mark_hidden = false;
                                let lines = if !is_text_empty(runs) && !image_only {
                                    let para_text_w = (cell_text_w
                                        - para.indent_left
                                        - para.indent_right
                                        - float_indent_left)
                                        .max(0.0);
                                    // Match the rendering's first_line_hanging: when a
                                    // list label is present, the label is drawn separately
                                    // and the text starts at indent_left, so the first
                                    // line has no extra hanging width.
                                    let hanging = super::list_label::text_hanging(para, ctx.default_tab_stop, ctx.fonts);
                                    let has_tabs = runs.iter().any(|r| r.is_tab);
                                    let inline_images = super::layout::run_image_map(runs, ctx.table_cell_image_names);
                                    let inline_effects = super::layout::run_image_map(runs, ctx.effect_table_names);
                                    let mut lines = if has_tabs {
                                        build_tabbed_line(
                                            runs,
                                            ctx.fonts,
                                            &para.tab_stops,
                                            para.indent_left,
                                            para_text_w,
                                            para.indent_right,
                                            hanging,
                                            &inline_images,
                                            &inline_effects,
                                            ctx.default_tab_stop,
                                            &[],
                                            ctx.compat_mode,
                                            ctx.cjk(false, para),
                                        )
                                    } else {
                                        let beside_widths = vec![(para_text_w - beside_indent).max(1.0); beside_lines];
                                        build_paragraph_lines(
                                            runs,
                                            ctx.fonts,
                                            para_text_w,
                                            hanging,
                                            &inline_images,
                                            &inline_effects,
                                            None,
                                            (beside_lines > 0).then_some(beside_widths.as_slice()),
                                            None,
                                            ctx.cjk(para.auto_space_de || para.auto_space_dn, para),
                                        )
                                    };
                                    if is_rotated {
                                        for line in &lines {
                                            max_rotated_line_w =
                                                max_rotated_line_w.max(line.total_width);
                                        }
                                    }
                                    // Each line is as tall as its own runs, as in body
                                    // text: nabl's "(Mark √ in the" header line, √ a
                                    // w:sym Symbol run, steps 12.24 where Arial gives 11.50.
                                    if !east_asian && !grid_snapped && !matches!(effective_ls, crate::model::LineSpacing::Exact(_)) {
                                        super::layout::size_lines_by_own_runs(
                                            &mut lines,
                                            ctx.fonts,
                                            effective_ls,
                                            line_h,
                                            font_size * ascender_ratio,
                                            para.text_alignment == crate::model::TextAlignment::Top,
                                        );
                                        super::size_mark_line(&mut lines, para, runs, effective_ls, tallest_lhr, ctx.fonts);
                                    }
                                    // hideMark also hides the mark left alone on the
                                    // line after the cell's closing break, with its
                                    // spacing (maine's "…this Bureau.<w:br/>"; Word
                                    // probes: one or two breaks, 0 or 12pt after).
                                    if cell.hide_mark
                                        && block_idx == block_count - 1
                                        && lines.len() > 1
                                        && lines.last().is_some_and(|l| l.chunks.is_empty())
                                        && lines[lines.len() - 2].ends_with_break
                                    {
                                        lines.pop();
                                        break_mark_hidden = true;
                                    }
                                    total_h += lines
                                        .iter()
                                        .map(|l| super::layout::inline_line_advance(l, line_h, (image_line_ascent, font_size * descender_ratio)))
                                        .sum::<f32>();
                                    lines
                                } else {
                                    if para.paragraph_mark_vanish {
                                        // vanished paragraph mark: zero height
                                    } else if mark_hidden {
                                        // hideMark: last empty paragraph in cell
                                        // contributes no height
                                    } else if prev_was_nested_table
                                        && block_idx == block_count - 1
                                        && para_content_h == 0.0
                                    {
                                        // End-of-cell mark directly after a nested
                                        // table: Word hides it; space_after is
                                        // suppressed in the trailing-space block
                                        // below.
                                    } else if para_content_h > 0.0 {
                                        // Image paragraph: the image is line 1; each
                                        // trailing w:br adds a further blank line
                                        // (matches the non-table header height path).
                                        let br_count = runs
                                            .iter()
                                            .filter(|r| r.is_line_break)
                                            .count();
                                        total_h += para_content_h
                                            + br_count as f32 * line_h;
                                    } else {
                                        total_h += line_h;
                                    }
                                    vec![]
                                };

                                let image_name = para.image.as_ref().and_then(|img| {
                                    let key = std::sync::Arc::as_ptr(&img.data) as usize;
                                    ctx.table_cell_image_names.get(&key).cloned()
                                });
                                let (image_width, image_height, img_stroke_color, img_stroke_width, img_shadow) = para
                                    .image
                                    .as_ref()
                                    .map(|img| (img.display_width, img.display_height, img.stroke_color, img.stroke_width, img.shadow.clone()))
                                    .unwrap_or((0.0, 0.0, None, 0.0, None));
                                let table_fx = para.image.as_ref().and_then(|img| {
                                    let key = std::sync::Arc::as_ptr(&img.data) as usize;
                                    ctx.effect_table_names.get(&key)
                                });
                                let img_shadow_xobj = table_fx.and_then(|fx| fx.shadow.clone());
                                let img_glow = para.image.as_ref().and_then(|img| img.glow.clone());
                                let img_glow_xobj = table_fx.and_then(|fx| fx.glow.clone());

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
                                                // Modern Word aligns to the cell's right text edge.
                                                // Legacy mode also shifts the origin by the left inset.
                                                let legacy_origin = if ctx.compat_mode < 15 {
                                                    ecm.left
                                                } else {
                                                    0.0
                                                };
                                                col_w + legacy_origin - ecm.right - fi.image.display_width
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
                                            rotation_deg: fi.image.rotation_deg,
                                            z_index: fi.z_index,
                                            alt: fi.image.alt.clone(),
                                            decorative: fi.image.decorative,
                                        })
                                    })
                                    .collect();

                                let next = super::block_para(&cell.content, block_idx + 1);
                                let bottom_border = if next.is_some_and(|n| super::helpers::joins_border_group(para, n)) {
                                    para.borders.between.clone()
                                } else { para.borders.bottom.clone() };
                                total_h += super::helpers::border_band(bottom_border.as_ref());

                                items.push(CellContentItem::Paragraph(CellParagraphLayout {
                                    lines,
                                    line_h,
                                    font_size,
                                    // Places the first baseline only, which rises
                                    // with a box shrunk below single spacing.
                                    // Exact and winning at-least lines place it as
                                    // in the body: carbon_farming's 8pt footer cells
                                    // sit on their descent under Normal's 13pt
                                    // minimum, arizona's exact 13.2pt lines 80% down.
                                    // Grid-snapped lines centre in their cells as
                                    // in the body (japanese_hr's 9pt lines sit
                                    // 2.9pt into a 17.5pt cell).
                                    ascender_ratio: grid_snapped
                                        .then(|| super::layout::grid_baseline_offset(runs, ctx.fonts, line_h))
                                        .flatten()
                                        .or_else(|| super::layout::boxed_line_ascent(
                                            effective_ls,
                                            line_h,
                                            font_size,
                                            tallest_lhr,
                                            tallest_ar,
                                            runs,
                                            ctx.fonts,
                                        ))
                                        .map_or(
                                            ascender_ratio
                                                * super::helpers::auto_ascent_scale(effective_ls),
                                            |a| a / font_size,
                                        ),
                                    descender_ratio,
                                    image_line_ascent,
                                    clip_inline_pictures: exact_spacing,
                                    font_substituted,
                                    alignment: para.alignment,
                                    space_before,
                                    continuation_space_before: super::own_space_before(space_before, prev_space_after),
                                    indent_left: para.indent_left,
                                    indent_right: para.indent_right,
                                    indent_hanging: para.indent_hanging,
                                    indent_first_line: para.indent_first_line,
                                    text_hanging: super::list_label::text_hanging(para, ctx.default_tab_stop, ctx.fonts),
                                    label_shift: super::list_label::label_shift(para, ctx.fonts),
                                    float_indent_left,
                                    beside_indent,
                                    beside_lines,
                                    // Word 2010 layout splits a keepLines cell paragraph too.
                                    keep_lines: para.keep_lines && ctx.compat_mode >= 15,
                                    // Without widow control (and nothing keeping the row
                                    // with the next) Word 2013+ also lets one line stand.
                                    two_line_widow: ctx.compat_mode >= 15 && (para.widow_control || row_keeps_next),
                                    list_label: para.list_label.clone(),
                                    sole_formula: super::sole_formula(para),
                                    label_pieces: if para.list_label.is_empty() {
                                        Vec::new()
                                    } else {
                                        super::list_label::label_pieces(para)
                                            .into_iter()
                                            .map(|(k, t)| (k, t.to_string()))
                                            .collect()
                                    },
                                    label_font_size: para.list_label_font_size.unwrap_or(font_size),
                                    list_item: para.list_item,
                                    label_color: para.list_label_color,
                                    image_name,
                                    image_alt: para.image.as_ref().and_then(|i| i.alt.clone()),
                                    image_decorative: para
                                        .image
                                        .as_ref()
                                        .is_some_and(|i| i.decorative),
                                    image_width,
                                    image_height,
                                    image_stroke_color: img_stroke_color,
                                    image_stroke_width: img_stroke_width,
                                    image_shadow: img_shadow,
                                    image_shadow_xobj: img_shadow_xobj,
                                    image_glow: img_glow,
                                    image_glow_xobj: img_glow_xobj,
                                    image_clip: para.image.as_ref().and_then(|img| img.clip_geometry.clone()),
                                    content_height: para_content_h,
                                    bottom_border,
                                    paragraph_mark_vanish: para.paragraph_mark_vanish,
                                    floating_images: cell_floats,
                                    space_after,
                                    has_textboxes: !para.textboxes.is_empty(),
                                    has_connectors: !para.connectors.is_empty(),
                                }));

                                prev_space_after = if mark_hidden
                                    || break_mark_hidden
                                    || (para.space_after_auto && block_idx == block_count - 1)
                                {
                                    0.0
                                } else {
                                    space_after
                                };
                                prev_was_nested_table = false;
                                para_idx += 1;
                            }
                            Block::Table(nested_table) => {
                                // Adjacent text-relative floats share the flow anchor.
                                // A preceding negative float displaces its empty
                                // anchor below the table before the next float's
                                // own offset is applied (not just to the table edge).
                                if nested_table.position.as_ref().is_some_and(|p| p.v_anchor == "text")
                                    && left_float_right_gap < MIN_NESTED_FLOAT_STRIP
                                    // Modern expanded anchors are not relocated (Word 15).
                                    && !(ctx.compat_mode >= 15
                                        && cell.content.get(block_idx.saturating_sub(2))
                                            .is_some_and(|b| matches!(b, Block::Paragraph(p)
                                                if matches!(p.line_spacing.unwrap_or(ctx.doc_line_spacing),
                                                    crate::model::LineSpacing::Auto(factor) if factor > 1.0))))
                                    && matches!(items.last(), Some(CellContentItem::NestedTable {
                                        floating_offset: Some(offset), ..
                                    }) if *offset < 0.0)
                                    && let Some(CellContentItem::Paragraph(anchor)) = items.get(items.len().saturating_sub(2))
                                    && block_idx.checked_sub(2).and_then(|i| cell.content.get(i))
                                        .is_some_and(|b| matches!(b, Block::Paragraph(p) if is_text_empty(&p.runs)))
                                    && anchor.image_name.is_none() && anchor.floating_images.is_empty()
                                {
                                    negative_anchor_shift = (left_float_bottom + anchor.line_h - total_h).max(0.0);
                                }
                                // The widths the nested table is drawn at, so a split's
                                // line cursors index the lines it draws.
                                let mut nested_cw = auto_fit_columns(nested_table, ctx.fonts, Some(cell_text_w), None);
                                apply_nested_pct_width(nested_table, &mut nested_cw, cell_text_w);
                                let nested_layouts =
                                    compute_row_layouts(nested_table, &nested_cw, ctx, hf_sub);
                                let nested_h = nested_layouts.iter().map(|rl| rl.height).sum::<f32>();
                                let plain_side_text = plain_text_fits_nested_strip(cell, block_idx, nested_table, &nested_cw, cell_text_w, ctx);
                                let blank = |p: &crate::model::Paragraph| {
                                    p.runs.iter().all(|r| r.text.trim().is_empty()
                                        && !r.is_tab && !r.is_line_break && r.inline_image.is_none() && r.field_code.is_none())
                                        && p.image.is_none() && p.floating_images.is_empty() && p.textboxes.is_empty()
                                };
                                // A left float leaving a text strip of at least
                                // MIN_NESTED_FLOAT_STRIP beside it may be followed by
                                // text: the text flows beside it (advocacy's picture
                                // tables with their stories to the right).
                                let mut floating_offset = nested_table.position.as_ref()
                                    .filter(|pos| pos.v_anchor == "text" && (nested_table.rows.len() == 1 || plain_side_text) && (top_aligned || (centered && pos.v_offset_pt < 0.0) || (centered && block_idx == 0 && pos.v_offset_pt > 0.0)))
                                    .filter(|pos| {
                                        let at_left_edge = matches!(pos.h_position, HorizontalPosition::Offset(v) if v == 0.0);
                                        let next_is_blank = matches!(
                                            cell.content.get(block_idx + 1),
                                            Some(Block::Paragraph(p)) if blank(p) && p.frame_props.is_none() && !p.paragraph_mark_vanish
                                        );
                                        cell.content[block_idx + 1..].iter().all(|b| match b {
                                            Block::Paragraph(p) => blank(p),
                                            Block::Table(t) => t.rows.len() == 1 && (
                                                t.position.as_ref().is_some_and(|p| p.v_anchor == "text")
                                                || (at_left_edge && t.position.is_none() && next_is_blank)
                                            ),
                                        }) || plain_side_text
                                            || text_above_floats
                                            || (top_aligned && (spacers_clear_nested_registration(cell, block_idx, cell_text_w, ctx, hf_sub)
                                                || interleaved_spacers_clear_nested_registration(cell, block_idx, cell_text_w, ctx, hf_sub)))
                                    })
                                    // Word keeps a table at the cell's top when its first
                                    // text anchor has a negative vertical offset.
                                    .map(|pos| if block_idx == 0 { pos.v_offset_pt.max(0.0) } else { pos.v_offset_pt })
                                    .map(|offset| offset + negative_anchor_shift);
                                let space_before = prev_space_after;
                                let mut floating_x_offset = None;
                                let auto_left_float = nested_table.position.as_ref().is_some_and(|p|
                                    matches!(p.h_position, crate::model::HorizontalPosition::Offset(v) if (p.h_anchor == "margin" && v <= 0.0 && v + nested_cw.iter().sum::<f32>() > 0.0)
                                        || (p.h_anchor == "column" && plain_side_text && v == 0.0)))
                                    && nested_table.alignment == crate::model::TableAlignment::Left
                                    && nested_table.table_indent == 0.0;
                                if auto_left_float && let Some(offset) = floating_offset {
                                    let top = total_h + space_before + offset;
                                    if left_float_bottom > top {
                                        floating_offset = Some(offset + left_float_bottom - top);
                                        floating_x_offset = Some(nested_table.position.as_ref().unwrap().left_from_text);
                                    }
                                }
                                if let Some(offset) = floating_offset {
                                    floating_extent = floating_extent.max(total_h + space_before + offset + nested_h);
                                    if auto_left_float {
                                        left_float_bottom = total_h + space_before + offset + nested_h;
                                        left_float_top = total_h + space_before + offset;
                                        left_float_indent = if plain_side_text { nested_cw.iter().sum::<f32>() + nested_table.position.as_ref().unwrap().right_from_text } else { 0.0 };
                                        first_positive_float = block_idx == 0 && nested_table.position.as_ref().is_some_and(|p| p.v_offset_pt >= 0.0);
                                        let float_x = match nested_table.position.as_ref().unwrap().h_position { crate::model::HorizontalPosition::Offset(x) => x, _ => 0.0 };
                                        left_float_right_gap = cell_text_w - float_x - nested_cw.iter().sum::<f32>()
                                            - nested_table.position.as_ref().unwrap().right_from_text;
                                    }
                                } else {
                                    total_h += space_before + nested_h;
                                }
                                items.push(CellContentItem::NestedTable {
                                    col_widths: nested_cw,
                                    rows: nested_layouts,
                                    space_before,
                                    floating_offset,
                                    floating_x_offset,
                                });
                                if floating_offset.is_none() { prev_space_after = 0.0; }
                                // A lone end-of-cell mark directly anchoring a
                                // wide float does not add a line below that float.
                                final_float_mark_hidden = block_idx + 2 == block_count
                                    && auto_left_float && left_float_right_gap < MIN_NESTED_FLOAT_STRIP
                                    && floating_offset.is_some_and(|offset| offset >= 0.0)
                                    && matches!(cell.content.get(block_idx+1),Some(Block::Paragraph(p)) if is_text_empty(&p.runs));
                                // Positive offsets suppress the centered cell's trailing mark; negative offsets displace its anchor line.
                                prev_was_nested_table = final_float_mark_hidden || floating_offset.is_none() || (centered && floating_offset.is_some_and(|offset| offset > 0.0))
                                    || suppress_negative_mark(block_idx, floating_offset)
                                    || (negative_anchor_shift > 0.0 && floating_offset.is_some());
                                para_idx += 1;
                            }
                        }
                    }

                    // When a cell ends with a nested table plus the mandatory
                    // end-of-cell paragraph mark (empty, no text), Word does
                    // not count the trailing paragraph's space_after toward
                    // the row height — the mark glyph height and line_h are
                    // already suppressed above via prev_was_nested_table.
                    let trailing_mark_after_table = items.len() >= 2
                        && matches!(items.get(items.len() - 2), Some(CellContentItem::NestedTable { floating_offset, .. })
                            if final_float_mark_hidden || floating_offset.is_none() || (centered && floating_offset.is_some_and(|offset| offset > 0.0))
                                || suppress_negative_mark(block_count.saturating_sub(2), *floating_offset)
                                || (negative_anchor_shift > 0.0 && floating_offset.is_some()))
                        && matches!(items.last(), Some(CellContentItem::Paragraph(p)) if p.lines.is_empty() && p.image_name.is_none() && p.floating_images.is_empty());
                    let trailing_space_after =
                        if trailing_mark_after_table { 0.0 } else { prev_space_after };
                    total_h += trailing_space_after;
                    total_h = total_h.max(floating_extent);
                    if is_rotated {
                        total_h = ecm.top + ecm.bottom + max_rotated_line_w;
                    }
                    if cell.v_merge != VMerge::Restart {
                        max_h = max_h.max(total_h);
                    }
                    CellLayout {
                        items,
                        cm: *ecm,
                        total_height: total_h,
                        trailing_space_after,
                        text_direction: cell.text_direction,
                    }
                })
                .collect();

            // Border bands are already in the cell insets (docx::tables); the
            // 0.5pt once added here was Table Grid's border width.
            let content_h = max_h;
            // An at-least trHeight bounds the row between its border bands,
            // which sit on top of it: 0.5pt Table Grid rows step trHeight + 0.5
            // (belgian_youth 19.85 → 20.5, italian_evaluation 12.0 → 12.48,
            // japanese_interlibrary 21.25 → 21.77). An exact height is the
            // border-to-border pitch (case15: 36.0).
            // The cells' top and bottom margins sit on top of it too, like the
            // bands in their insets: a 30pt row with 5pt margins is 40.5 in
            // Word, and polish_ministry's 33.15pt rows with 1.4pt ones 36.25.
            let insets = cells
                .iter()
                .map(|c| c.cm.top + c.cm.bottom)
                .fold(0.0f32, f32::max);
            let height = match (row.height, row.height_exact) {
                // An exact one gains only the bottom margin: 30pt with 5pt
                // margins is 35.0, with only a 5pt top margin 30.0 (Word probes).
                // The inset's half band is already in the exact pitch.
                (Some(h), true) => {
                    h + cells
                        .iter()
                        .zip(row.grid_cells())
                        .map(|(l, (_, _, c))| l.cm.bottom - c.borders.bottom.band() / 2.0)
                        .fold(0.0f32, f32::max)
                }
                (Some(h), false) => content_h.max(h + insets),
                _ => content_h,
            };


            RowLayout {
                height,
                cells,
                split_min: (!row.cant_split && !row.height_exact)
                    .then(|| row.height.map_or(0.0, |h| h + insets)),
            }
        })
        .collect();

    // A merged cell taller than the rows it spans grows the last of them:
    // indonesian_school_admission_checklist's "NO." header (49.7pt of content)
    // makes the second row 29.28pt where its at-least trHeight asks 24.15.
    for (ri, row) in table.rows.iter().enumerate() {
        for (ci, (grid_col, _, cell)) in row.grid_cells().enumerate() {
            if cell.v_merge != VMerge::Restart {
                continue;
            }
            let mut last = ri;
            while table.rows.get(last + 1).is_some_and(|next| {
                next.grid_cells()
                    .any(|(c, _, n)| c == grid_col && n.v_merge == VMerge::Continue)
            }) {
                last += 1;
            }
            let spanned: f32 = layouts[ri..=last].iter().map(|l| l.height).sum();
            let overflow = layouts[ri].cells[ci].total_height - spanned;
            if overflow > 0.0
                && let Some(r) = (ri..=last).rev().find(|&r| !table.rows[r].height_exact)
            {
                layouts[r].height += overflow;
            }
        }
    }
    layouts
}

/// Pre-compute how much extra height each vMerge Restart cell spans beyond its own row.
/// Returns a map from (row_idx, grid_col) to the sum of Continue row heights below;
/// a Continue cell maps to the height of the span rows still below it, absent
/// on the last, so a row can tell whether the merged cell ends there.
pub(super) fn compute_merge_spans(
    table: &Table,
    row_layouts: &[RowLayout],
) -> HashMap<(usize, usize), f32> {
    // Build a grid index: vmerge_grid[row][grid_col] = VMerge value
    let max_cols = table
        .rows
        .iter()
        .map(|r| {
            r.grid_before
                + r.cells
                    .iter()
                    .map(|c| c.grid_span.max(1) as usize)
                    .sum::<usize>()
        })
        .max()
        .unwrap_or(0);
    let mut vmerge_grid: Vec<Vec<VMerge>> = Vec::with_capacity(table.rows.len());
    for row in &table.rows {
        let mut row_vmerge = vec![VMerge::None; max_cols];
        for (col, _, cell) in row.grid_cells() {
            if col < max_cols {
                row_vmerge[col] = cell.v_merge;
            }
        }
        vmerge_grid.push(row_vmerge);
    }

    let mut spans = HashMap::new();
    for (ri, row) in table.rows.iter().enumerate() {
        for (grid_col, _, cell) in row.grid_cells() {
            if cell.v_merge == VMerge::Restart {
                let end = (ri + 1..table.rows.len())
                    .find(|&next_ri| {
                        grid_col >= max_cols || vmerge_grid[next_ri][grid_col] != VMerge::Continue
                    })
                    .unwrap_or(table.rows.len());
                let mut below = 0.0f32;
                for next_ri in (ri + 1..end).rev() {
                    if below > 0.0 {
                        spans.insert((next_ri, grid_col), below);
                    }
                    below += row_layouts[next_ri].height;
                }
                if below > 0.0 {
                    spans.insert((ri, grid_col), below);
                }
            }
        }
    }
    spans
}

/// Position in a cell's content for row splitting: items before `item` are
/// emitted; when `line > 0`, paragraph `item` is emitted up to that line.
#[derive(Clone, Default, PartialEq, Debug)]
pub(super) struct CellCursor {
    pub(super) item: usize,
    pub(super) line: usize,
    /// Nested table `item` broken inside its row `line`: where each cell of
    /// that row stands (empty when the break is between rows).
    pub(super) nested: Vec<CellCursor>,
}

/// A cell's beginning, for borrowing where a cursor is missing.
pub(super) static CELL_START: CellCursor = CellCursor {
    item: 0,
    line: 0,
    nested: Vec::new(),
};

impl CellCursor {
    pub(super) fn at(item: usize, line: usize) -> Self {
        CellCursor {
            item,
            line,
            nested: Vec::new(),
        }
    }
}

impl CellContentItem {
    /// The item's full height (a paragraph's block, a nested table's rows).
    pub(super) fn height(&self) -> f32 {
        match self {
            CellContentItem::Paragraph(p) => para_block_height(p),
            CellContentItem::NestedTable {
                rows,
                space_before,
                floating_offset,
                ..
            } => {
                if floating_offset.is_some() {
                    0.0
                } else {
                    space_before + rows.iter().map(|r| r.height).sum::<f32>()
                }
            }
        }
    }
}

/// One piece of a cell between two cursors: item `item`, its lines (or nested
/// rows) `l0..l1` (`l1` None = to the end), and for a nested table the row
/// cursors `from` continuing row `l0` partway and `to` ending row `l1` partway.
#[derive(Clone, Copy, PartialEq, Debug)]
pub(super) struct Chunk<'a> {
    pub(super) item: usize,
    pub(super) l0: usize,
    pub(super) l1: Option<usize>,
    pub(super) from: &'a [CellCursor],
    pub(super) to: &'a [CellCursor],
}

/// The pieces of a cell between two cursors.
pub(super) fn cursor_chunks<'a>(
    items: &'a [CellContentItem],
    start: &'a CellCursor,
    end: &'a CellCursor,
) -> impl Iterator<Item = Chunk<'a>> + 'a {
    let partial_end = end.line > 0 || !end.nested.is_empty();
    let last = if partial_end { end.item + 1 } else { end.item };
    (start.item..last.min(items.len())).map(move |pi| {
        let first = pi == start.item;
        let at_end = pi == end.item && partial_end;
        Chunk {
            item: pi,
            l0: if first { start.line } else { 0 },
            l1: at_end.then_some(end.line),
            from: if first { &start.nested } else { &[] },
            to: if at_end { &end.nested } else { &[] },
        }
    })
}

/// How a nested table's chunk draws: whole rows, or one row from `starts` to
/// `ends` (missing cursors: the cell's start, or its end).
pub(super) enum RowPiece<'a> {
    Rows(std::ops::Range<usize>),
    Partial {
        row: usize,
        starts: &'a [CellCursor],
        ends: &'a [CellCursor],
    },
}

/// A nested table chunk as row pieces: the rest of row `l0` when it continues
/// partway, the whole rows, then row `l1` up to `to` when it ends partway.
pub(super) fn row_pieces<'a>(n_rows: usize, c: &Chunk<'a>) -> Vec<RowPiece<'a>> {
    let end = c.l1.unwrap_or(n_rows).min(n_rows);
    let mut pieces = Vec::new();
    let mut r = c.l0;
    if !c.from.is_empty() && r < n_rows {
        let ends = if c.l1 == Some(r) { c.to } else { &[] };
        pieces.push(RowPiece::Partial {
            row: r,
            starts: c.from,
            ends,
        });
        r += 1;
    }
    if r < end {
        pieces.push(RowPiece::Rows(r..end));
    }
    if !c.to.is_empty() && end < n_rows && end >= r {
        pieces.push(RowPiece::Partial {
            row: end,
            starts: &[],
            ends: c.to,
        });
    }
    pieces
}

/// Height of a chunk; items without lines use their block height.
pub(super) fn item_chunk_height(item: &CellContentItem, c: &Chunk) -> f32 {
    match item {
        CellContentItem::Paragraph(p) if !p.lines.is_empty() => {
            cell_lines_h(p, c.l0..c.l1.unwrap_or(p.lines.len()))
        }
        CellContentItem::Paragraph(p) => para_block_height(p),
        CellContentItem::NestedTable {
            floating_offset: Some(_),
            ..
        } => 0.0,
        CellContentItem::NestedTable {
            rows, space_before, ..
        } => {
            (if c.l0 == 0 && c.from.is_empty() {
                *space_before
            } else {
                0.0
            }) + row_pieces(rows.len(), c)
                .into_iter()
                .map(|piece| match piece {
                    RowPiece::Rows(range) => rows[range].iter().map(|rl| rl.height).sum::<f32>(),
                    RowPiece::Partial { row, starts, ends } => {
                        partial_row_height(&rows[row], starts, ends)
                    }
                })
                .sum::<f32>()
        }
    }
}

/// Reanchor dependent floats after dropping the first anchor on a new page.
pub(super) fn floating_anchor_gap(
    space_before: f32,
    offset: f32,
    collided: bool,
    drop_anchor: bool,
    reanchor_shift: &mut f32,
) -> f32 {
    if drop_anchor {
        *reanchor_shift = space_before + offset;
        0.0
    } else {
        space_before + offset - if collided { *reanchor_shift } else { 0.0 }
    }
}

/// Height of a row's piece from `starts` to `ends` (one cursor per cell; a
/// missing start is the cell's beginning, a missing end its finish): its
/// tallest cell's content.
pub(super) fn partial_row_height(
    layout: &RowLayout,
    starts: &[CellCursor],
    ends: &[CellCursor],
) -> f32 {
    let mut max_h: f32 = 0.0;
    for (ci, cell_layout) in layout.cells.iter().enumerate() {
        let cm = &cell_layout.cm;
        let start = starts.get(ci).unwrap_or(&CELL_START);
        let done = CellCursor::at(cell_layout.items.len(), 0);
        let end = ends.get(ci).unwrap_or(&done);
        let mut h = cm.top + cm.bottom;
        let mut floating_extent: f32 = 0.0;
        let mut reanchor_shift = 0.0;
        for c in cursor_chunks(&cell_layout.items, start, end) {
            let item = &cell_layout.items[c.item];
            if let CellContentItem::NestedTable {
                rows,
                space_before,
                floating_offset: Some(offset),
                floating_x_offset,
                ..
            } = item
            {
                let anchor_gap = floating_anchor_gap(
                    *space_before,
                    *offset,
                    floating_x_offset.is_some(),
                    c.item == start.item && start.item > 0 && start.line == 0 && c.from.is_empty(),
                    &mut reanchor_shift,
                );
                floating_extent = floating_extent
                    .max(h + anchor_gap + rows.iter().map(|r| r.height).sum::<f32>());
            }
            let mut piece_h = item_chunk_height(item, &c);
            if c.item == start.item
                && start.item > 0
                && start.line == 0
                && c.from.is_empty()
                && let CellContentItem::NestedTable {
                    space_before,
                    floating_offset: None,
                    ..
                } = item
            {
                piece_h -= space_before;
            }
            h += chunk_space_before(item, c.item, start) + piece_h;
        }
        // The chunk that finishes the cell keeps its last paragraph's space
        // after, as an unsplit row does.
        if end.item >= cell_layout.items.len() {
            h += cell_layout.trailing_space_after;
        }
        max_h = max_h.max(h.max(floating_extent));
    }
    max_h
}

/// The paragraph's space_before as charged inside a chunk starting at
/// `start`: a paragraph continuing mid-way never repeats it, but one starting
/// a chunk whole keeps it like an unsplit row (croatian_grant's floating
/// "Važno!" box starts 6pt below its top border in Word; nabl's carried-over
/// "Remarks" paragraph keeps its 4pt).
pub(super) fn chunk_space_before(item: &CellContentItem, pi: usize, start: &CellCursor) -> f32 {
    match item {
        CellContentItem::Paragraph(p)
            if pi == start.item && start.line == 0 && *start != CELL_START =>
        {
            p.continuation_space_before
        }
        CellContentItem::Paragraph(p) if pi != start.item || start.line == 0 => p.space_before,
        _ => 0.0,
    }
}

/// Where the cell content from `start` must break to fit `available_h`.
/// Word breaks a row's paragraph between lines. In Word 2013+ layout
/// (compatibilityMode 15) two lines stay on each side (strategi, italian,
/// carbon_farming: two-line cell paragraphs move whole; croatian's three-
/// line one leaves no single line above the footer). In Word 2010 layout
/// (mode 14) a row starts with as many lines as fit, one included, and one
/// may follow alone (Word online probes 2026-10-22 with 1-3 lines of room,
/// widowControl set or absent; advocacy's two-line heading splits 1/1).
/// A keepLines paragraph moves whole. Always progresses by at least one item.
pub(super) fn find_cell_split(
    cell: &CellLayout,
    start: &CellCursor,
    available_h: f32,
) -> CellCursor {
    let cm = &cell.cm;
    let done = CellCursor::at(cell.items.len(), 0);
    if start.item >= cell.items.len() {
        return done;
    }
    let mut h = cm.top + cm.bottom;
    let mut reanchor_shift = 0.0;
    for pi in start.item..cell.items.len() {
        let first = pi == start.item;
        let l0 = if first { start.line } else { 0 };
        let item = &cell.items[pi];
        let sb = chunk_space_before(item, pi, start);
        let from: &[CellCursor] = if first { &start.nested } else { &[] };
        let rest = Chunk {
            item: pi,
            l0,
            l1: None,
            from,
            to: &[],
        };
        let mut item_h = sb + item_chunk_height(item, &rest);
        if first
            && start.item > 0
            && l0 == 0
            && from.is_empty()
            && let CellContentItem::NestedTable {
                space_before,
                floating_offset: None,
                ..
            } = item
        {
            item_h -= space_before;
        }
        if let CellContentItem::NestedTable {
            rows,
            space_before,
            floating_offset: Some(offset),
            floating_x_offset,
            ..
        } = item
        {
            let anchor_gap = floating_anchor_gap(
                *space_before,
                *offset,
                floating_x_offset.is_some(),
                first && start.item > 0 && l0 == 0 && from.is_empty(),
                &mut reanchor_shift,
            );
            let required = (anchor_gap + rows.iter().map(|r| r.height).sum::<f32>()).max(0.0);
            if h + required <= available_h {
                continue;
            }
            return CellCursor::at(if pi == start.item { pi + 1 } else { pi }, 0);
        }
        // A paragraph fits only with its space after: nabl's "Remarks" row
        // moves its last (4pt after) paragraph to the next page in Word.
        let sa = match item {
            CellContentItem::Paragraph(p) => p.space_after,
            CellContentItem::NestedTable { .. } => 0.0,
        };
        if h + item_h + sa <= available_h {
            h += item_h;
            continue;
        }
        if let CellContentItem::Paragraph(p) = item {
            let remaining = p.lines.len().saturating_sub(l0);
            let mut used = h + sb;
            let room = p.lines[l0..]
                .iter()
                .take_while(|l| {
                    used += l.pitch.unwrap_or(p.line_h);
                    used <= available_h
                })
                .count();
            // ponytail: the one-line start is only taken where the row itself
            // starts (the probes' case); later paragraphs keep two lines.
            let keep = if first && *start == CELL_START && !p.two_line_widow { 1 } else { 2 };
            let fit = room.min(remaining.saturating_sub(keep));
            if fit >= keep && !p.keep_lines {
                return CellCursor::at(pi, l0 + fit);
            }
        }
        // Word breaks a nested table between its rows (radiographer's
        // "Internal / External to the Trust" table starts on page 1), and
        // inside the first row that does not fit when that row may split, each
        // of its cells by these same rules (its "Administrative teams within
        // Radiology" closes page 1).
        if let CellContentItem::NestedTable { rows, .. } = item {
            let mut used = h + sb;
            let mut r = l0;
            // Only the first row can continue partway; the rest are whole.
            let mut row_start = from;
            while r < rows.len() {
                let rest = if row_start.is_empty() {
                    rows[r].height
                } else {
                    partial_row_height(&rows[r], row_start, &[])
                };
                if used + rest > available_h {
                    break;
                }
                used += rest;
                r += 1;
                row_start = &[];
            }
            if r < rows.len() && rows[r].split_min.is_some_and(|m| available_h - used >= m) {
                let starts: Vec<CellCursor> = (0..rows[r].cells.len())
                    .map(|ci| row_start.get(ci).unwrap_or(&CELL_START).clone())
                    .collect();
                let ends: Vec<CellCursor> = rows[r]
                    .cells
                    .iter()
                    .zip(&starts)
                    .map(|(c, s)| find_cell_split(c, s, available_h - used))
                    .collect();
                if ends != starts
                    && used + partial_row_height(&rows[r], &starts, &ends) <= available_h
                {
                    return CellCursor {
                        item: pi,
                        line: r,
                        nested: ends,
                    };
                }
            }
            if r > l0 && r < rows.len() {
                return CellCursor::at(pi, r);
            }
        }
        if !first {
            return CellCursor::at(pi, 0);
        }
        // The first item is force-included so the split makes progress.
        h += item_h;
    }
    done
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn nested_autofit_preserves_each_level_of_cell_padding() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;
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
            let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
            zip.start_file("word/document.xml", SimpleFileOptions::default())
                .unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
            let doc = crate::docx::parse_bytes(&zip.finish().unwrap().into_inner()).unwrap();
            let Block::Table(table) = &doc.sections[0].blocks[0] else {
                panic!("table");
            };
            let width = auto_fit_columns(table, &fonts, Some(150.0), None)[0];
            assert!((width - (35.0 + 10.0 * levels as f32)).abs() < 0.001);
        }
    }

    #[test]
    fn text_above_float_requires_both_horizontal_and_vertical_room() {
        let mut doc = crate::docx::parse_bytes(include_bytes!(
            "../../tests/fixtures/cases/cell-text-above-floats-y1200-n1-long0/input.docx"
        ))
        .unwrap();
        let Block::Table(table) = &mut doc.sections[0].blocks[0] else {
            panic!("outer table");
        };
        let cell = &mut table.rows[0].cells[0];
        let fonts = HashMap::from([("Times New Roman".into(), FontEntry::stub())]);
        let fits = |cell: &crate::model::TableCell, width| {
            text_fits_above_nested_floats(
                cell,
                width,
                &fonts,
                crate::model::LineSpacing::Auto(1.0),
                None,
            )
        };
        assert!(fits(cell, 230.0));
        assert!(!fits(cell, 10.0));
        let Block::Table(float) = &mut cell.content[0] else {
            panic!("float");
        };
        float.position.as_mut().unwrap().v_offset_pt = 1.0;
        assert!(!fits(cell, 230.0));
        let Block::Table(float) = &mut cell.content[0] else {
            unreachable!();
        };
        float.position.as_mut().unwrap().v_offset_pt = 60.0;
        let Some(Block::Paragraph(heading)) = cell.content.last_mut() else {
            panic!("heading");
        };
        heading.runs[0].is_line_break = true;
        assert!(!fits(cell, 230.0));
    }

    #[test]
    fn natural_width_includes_merged_cell_text() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;
        const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="3"/></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="TestFont" w:hAnsi="TestFont"/><w:sz w:val="20"/></w:rPr><w:t>{}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#,
            "x".repeat(100)
        );
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file("word/document.xml", SimpleFileOptions::default())
            .unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        let mut doc = crate::docx::parse_bytes(&zip.finish().unwrap().into_inner()).unwrap();
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
            let mut doc = crate::docx::parse_bytes(include_bytes!(
                "../../tests/fixtures/cases/min-width-run-whole/input.docx"
            ))
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
        let doc = crate::docx::parse_bytes(bytes).unwrap();
        let Block::Table(outer) = &doc.sections[0].blocks[0] else {
            panic!("outer table")
        };
        let Block::Table(nested) = &outer.rows[0].cells[0].content[0] else {
            panic!("nested table")
        };
        let fonts = HashMap::new();
        let kept = auto_fit_columns(nested, &fonts, Some(260.0), None);
        assert_eq!(kept, nested.col_widths);
        assert!(kept.iter().sum::<f32>() > 230.0);
        let squeezed = auto_fit_columns(nested, &fonts, Some(80.0), None);
        assert!(squeezed.iter().sum::<f32>() <= 80.001);
    }

    #[test]
    fn nested_explicit_width_uses_fixed_base_even_when_overflowing() {
        let bytes = include_bytes!("../../tests/fixtures/cases/case90/input.docx");
        let doc = crate::docx::parse_bytes(bytes).unwrap();
        let Block::Table(outer) = &doc.sections[0].blocks[0] else {
            panic!("outer table")
        };
        let Block::Table(nested) = &outer.rows[0].cells[0].content[0] else {
            panic!("nested table")
        };
        let fonts = HashMap::new();
        let kept = auto_fit_columns(nested, &fonts, Some(260.0), None);
        assert_eq!(kept, auto_fit_columns(nested, &fonts, None, None));
        assert!(kept.iter().sum::<f32>() > 230.0);
        let squeezed = auto_fit_columns(nested, &fonts, Some(80.0), None);
        assert_eq!(squeezed, kept);
    }

    #[test]
    fn asymmetric_autofit_parent_keeps_saved_grid_with_nested_table() {
        use crate::model::{CellVAlign, TableCell, TableRow, VMerge};
        let make = |grid: Vec<f32>, preferred: Vec<f32>| Table {
            col_widths: grid,
            rows: vec![TableRow {
                cells: preferred
                    .into_iter()
                    .map(|width| TableCell {
                        width,
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
            has_style_or_cell_mar: true,
            cell_margins: Default::default(),
            position: None,
            alignment: Default::default(),
            fixed_layout: false,
            auto_width: true,
            width_pct: None,
            grid_inferred: false,
            cells_unsized: false,
            header_first_row: false,
            header_first_col: false,
        };
        let mut parent = make(vec![250.0, 150.0], vec![250.0, 180.0]);
        parent.rows[0].cells[0]
            .content
            .push(Block::Table(make(vec![230.0], vec![230.0])));
        let fonts = HashMap::new();
        assert_eq!(
            auto_fit_columns(&parent, &fonts, None, None),
            vec![250.0, 150.0]
        );
        parent.fixed_layout = true;
        assert!(auto_fit_columns(&parent, &fonts, None, None)[0] < 240.0);
        parent.fixed_layout = false;
        parent.auto_width = false;
        assert!(auto_fit_columns(&parent, &fonts, None, None)[0] < 240.0);
    }

    #[test]
    fn asymmetric_autofit_keeps_grid_with_uniform_stale_cell_widths() {
        use crate::model::{CellVAlign, TableCell, TableRow, VMerge};
        let make = |grid: Vec<f32>, preferred: Vec<f32>| Table {
            col_widths: grid,
            rows: vec![TableRow {
                cells: preferred
                    .into_iter()
                    .map(|width| TableCell {
                        width,
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
            has_style_or_cell_mar: true,
            cell_margins: Default::default(),
            position: None,
            alignment: Default::default(),
            fixed_layout: false,
            auto_width: true,
            width_pct: None,
            grid_inferred: false,
            cells_unsized: false,
            header_first_row: false,
            header_first_col: false,
        };
        let mut table = make(vec![250.0, 150.0], vec![300.0, 300.0]);
        let fonts = HashMap::new();
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None),
            vec![250.0, 150.0]
        );
        table.col_widths = vec![235.6, 232.25];
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None),
            vec![235.6, 232.25]
        );
        table.col_widths = vec![250.0, 150.0];
        table.fixed_layout = true;
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None),
            vec![200.0, 200.0]
        );
        table.fixed_layout = false;
        table.auto_width = false;
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None),
            vec![200.0, 200.0]
        );
        table.auto_width = true;
        table.rows[0].cells[0].width = 250.0;
        table.rows[0].cells[1].width = 150.0;
        assert_eq!(
            auto_fit_columns(&table, &fonts, None, None),
            vec![250.0, 150.0]
        );
    }

    #[test]
    fn continued_cell_drops_previous_after_but_keeps_own_before() {
        for own in [0.0, 4.0] {
            let item = CellContentItem::Paragraph(CellParagraphLayout {
                space_before: 8.0,
                continuation_space_before: own,
                ..Default::default()
            });
            assert_eq!(chunk_space_before(&item, 1, &CELL_START), 8.0);
            assert_eq!(chunk_space_before(&item, 1, &CellCursor::at(1, 0)), own);
            assert_eq!(chunk_space_before(&item, 1, &CellCursor::at(1, 1)), 0.0);
        }
    }

    /// stem_partnerships p4 (annotation #237): a 10-line cell paragraph after a
    /// heading must break between lines, two lines minimum on either side.
    #[test]
    fn row_split_breaks_inside_a_long_paragraph() {
        let para = |n: usize| {
            CellContentItem::Paragraph(CellParagraphLayout {
                lines: (0..n).map(|_| TextLine::default()).collect(),
                line_h: 10.0,
                space_before: 5.0,
                ..Default::default()
            })
        };
        let cell = CellLayout {
            items: vec![para(1), para(10)],
            cm: CellMargins {
                top: 0.0,
                left: 0.0,
                bottom: 0.0,
                right: 0.0,
            },
            total_height: 0.0,
            trailing_space_after: 0.0,
            text_direction: TextDirection::default(),
        };
        let split = |start, avail| find_cell_split(&cell, &start, avail);
        let at = CellCursor::at;

        // the heading's own space_before (5, the cell opens with it) + heading
        // (10) + space_before (5) + four of the ten lines
        assert_eq!(split(at(0, 0), 60.0), at(1, 4));
        // the remaining six lines fit, no space_before on a continuation
        assert_eq!(split(at(1, 4), 60.0), at(2, 0));
        // room for one line only: the paragraph moves whole
        assert_eq!(split(at(0, 0), 26.0), at(1, 0));
        // nine lines would fit but two must stay for the next page
        assert_eq!(split(at(0, 0), 110.0), at(1, 8));
        let chunks: Vec<_> = cursor_chunks(&cell.items, &at(0, 0), &at(1, 4))
            .map(|c| (c.item, c.l0, c.l1))
            .collect();
        assert_eq!(chunks, vec![(0, 0, None), (1, 0, Some(4))]);
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

#[cfg(test)]
mod floating_nested_tests {
    use super::*;

    fn item(height: f32, gap: f32, floating_offset: Option<f32>) -> CellContentItem {
        CellContentItem::NestedTable {
            col_widths: vec![80.0],
            rows: vec![RowLayout {
                height,
                cells: vec![],
                split_min: None,
            }],
            space_before: gap,
            floating_offset,
            floating_x_offset: None,
        }
    }

    #[test]
    fn floating_extent_moves_whole_and_drops_previous_page_anchor_gap() {
        let floating = item(48.0, 8.0, Some(-1.65));
        assert_eq!(floating.height(), 0.0);
        let row = RowLayout {
            height: 66.35,
            cells: vec![CellLayout {
                items: vec![item(12.0, 0.0, None), floating],
                cm: CellMargins::default(),
                total_height: 66.35,
                trailing_space_after: 0.0,
                text_direction: TextDirection::LrTb,
            }],
            split_min: None,
        };
        let split = find_cell_split(&row.cells[0], &CELL_START, 30.0);
        assert_eq!(split.item, 1);
        assert!((partial_row_height(&row, &[], &[split.clone()]) - 12.0).abs() < 0.001);
        assert!((partial_row_height(&row, &[split.clone()], &[]) - 48.0).abs() < 0.001);
        assert!((partial_row_height(&row, &[], &[]) - 66.35).abs() < 0.001);
        assert_eq!(find_cell_split(&row.cells[0], &split, 48.0).item, 2);
    }

    #[test]
    fn continued_pair_reanchors_collision_extent_before_splitting() {
        let mut second = item(40.0, 0.0, Some(16.85));
        if let CellContentItem::NestedTable {
            floating_x_offset, ..
        } = &mut second
        {
            *floating_x_offset = Some(9.0);
        }
        let row = RowLayout {
            height: 88.35,
            cells: vec![CellLayout {
                items: vec![
                    item(12.0, 0.0, None),
                    item(30.0, 8.0, Some(-1.65)),
                    item(19.5, 0.0, None),
                    second,
                ],
                cm: CellMargins::default(),
                total_height: 88.35,
                trailing_space_after: 0.0,
                text_direction: TextDirection::LrTb,
            }],
            split_min: None,
        };
        let start = CellCursor::at(1, 0);
        assert!((partial_row_height(&row, &[start.clone()], &[]) - 70.0).abs() < 0.001);
        assert_eq!(find_cell_split(&row.cells[0], &start, 70.01).item, 4);
        assert_eq!(find_cell_split(&row.cells[0], &start, 69.9).item, 3);
    }
    #[test]
    fn bottom_border_band_belongs_only_to_the_last_line_chunk() {
        let para = CellParagraphLayout {
            lines: vec![TextLine::default(), TextLine::default()],
            line_h: 10.0,
            bottom_border: Some(crate::model::ParagraphBorder {
                width_pt: 1.5, space_pt: 1.0, color: [0, 0, 0],
            }),
            ..Default::default()
        };
        assert_eq!(cell_lines_h(&para, 0..1), 10.0);
        assert_eq!(cell_lines_h(&para, 1..2), 12.5);
        assert_eq!(para_block_height(&para), 22.5);
        let empty = CellParagraphLayout { lines: vec![], ..para };
        assert_eq!(para_block_height(&empty), 12.5);
    }

}
