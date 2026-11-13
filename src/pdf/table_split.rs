//! Row splitting across pages: cell cursors, chunks and where a cell breaks.

use std::collections::HashMap;

use crate::model::{Table, VMerge};

use super::table_layout::{
    CellContentItem, CellLayout, RowLayout, cell_lines_h, para_block_height,
};

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
                bottom_band,
                ..
            } => {
                if floating_offset.is_some() {
                    0.0
                } else {
                    space_before + rows.iter().map(|r| r.height).sum::<f32>() + bottom_band
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

/// Whether chunk `c` opens the continuation `start` (past the cell's first
/// item) with its item whole: a nested table there starts flush at the
/// piece's top, without its space before, and its anchor gap moves with it
/// (`floating_anchor_gap`).
pub(super) fn continues_whole(c: &Chunk, start: &CellCursor) -> bool {
    c.item == start.item && start.item > 0 && start.line == 0 && c.from.is_empty()
}

/// Height of a chunk with `room` left for it on its page, in a cell drawn
/// from `start`; items without lines use their block height.
pub(super) fn item_chunk_height(
    item: &CellContentItem,
    c: &Chunk,
    start: &CellCursor,
    room: f32,
) -> f32 {
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
            rows,
            space_before,
            bottom_band,
            ..
        } => {
            let mut h = if c.l0 == 0 && c.from.is_empty() && !continues_whole(c, start) {
                *space_before
            } else {
                0.0
            };
            for piece in row_pieces(rows.len(), c) {
                h += match piece {
                    RowPiece::Rows(range) => rows[range].iter().map(|rl| rl.height).sum::<f32>(),
                    RowPiece::Partial { row, starts, ends } => {
                        partial_row_height(&rows[row], starts, ends, room - h)
                    }
                };
            }
            if c.l1.is_none() && c.to.is_empty() {
                h += *bottom_band;
            }
            h
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
/// tallest cell's content, but no less than an at-least row's trHeight with
/// its cell margins on every page the row spans, up to the `room` left on
/// that page (Word examples: a 45pt row split 2/2 lines draws two
/// 55.5pt pieces, a 19.5pt row's one-line continuation 30pt; a 410pt row's
/// 190pt continuation fills a 350pt page body and ends there).
pub(super) fn partial_row_height(
    layout: &RowLayout,
    starts: &[CellCursor],
    ends: &[CellCursor],
    room: f32,
) -> f32 {
    let mut max_h: f32 = layout.split_min.map_or(0.0, |m| m.min(room));
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
                    continues_whole(&c, start),
                    &mut reanchor_shift,
                );
                floating_extent = floating_extent
                    .max(h + anchor_gap + rows.iter().map(|r| r.height).sum::<f32>());
            }
            let sb = chunk_space_before(item, c.item, start);
            h += sb + item_chunk_height(item, &c, start, room - h - sb);
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

/// `cursor` moved past every nested row it has finished in every cell, at
/// any depth, so the next page does not open with an empty piece of one.
pub(super) fn skip_finished_row(cell: &CellLayout, mut cursor: CellCursor) -> CellCursor {
    let Some(CellContentItem::NestedTable { rows, .. }) = cell.items.get(cursor.item) else {
        return cursor;
    };
    let Some(row) = rows.get(cursor.line) else {
        return cursor;
    };
    if cursor.nested.is_empty() || row.cells.len() != cursor.nested.len() {
        return cursor;
    }
    cursor.nested = std::mem::take(&mut cursor.nested)
        .into_iter()
        .zip(&row.cells)
        .map(|(n, c)| skip_finished_row(c, n))
        .collect();
    let finished = row
        .cells
        .iter()
        .zip(&cursor.nested)
        .all(|(c, n)| n.item >= c.items.len());
    if !finished {
        cursor
    } else if cursor.line + 1 < rows.len() {
        CellCursor::at(cursor.item, cursor.line + 1)
    } else {
        CellCursor::at(cursor.item + 1, 0)
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
/// A keepLines paragraph moves whole. Progresses by at least one item, except
/// that a nested table whose leading keep-with-next rows do not fit stays
/// whole for the next page (`start` comes back) unless `force` says this page
/// is that next page.
pub(super) fn find_cell_split(
    cell: &CellLayout,
    start: &CellCursor,
    available_h: f32,
    force: bool,
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
        let item_h = sb + item_chunk_height(item, &rest, start, available_h - h - sb);
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
                continues_whole(&rest, start),
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
            let keep = if first && *start == CELL_START && !p.two_line_widow {
                1
            } else {
                2
            };
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
                    partial_row_height(&rows[r], row_start, &[], available_h - used)
                };
                if used + rest > available_h {
                    break;
                }
                used += rest;
                r += 1;
                row_start = &[];
            }
            // Rows keeping with their successor move with the row that did
            // not fit (issue #145's panels); a chain too tall for any page
            // breaks where it must once it has a page of its own.
            let fits = r;
            while r > l0 && rows[r - 1].keeps_next {
                r -= 1;
            }
            if r < fits && r == l0 && first && *start == CELL_START {
                if !force {
                    return start.clone();
                }
                r = fits;
            }
            if r == fits
                && r < rows.len()
                && rows[r]
                    .split_min
                    .is_some_and(|m| available_h - used >= m || (first && r == l0))
            {
                let starts: Vec<CellCursor> = (0..rows[r].cells.len())
                    .map(|ci| row_start.get(ci).unwrap_or(&CELL_START).clone())
                    .collect();
                let ends: Vec<CellCursor> = rows[r]
                    .cells
                    .iter()
                    .zip(&starts)
                    .map(|(c, s)| find_cell_split(c, s, available_h - used, force))
                    .collect();
                if ends != starts
                    && used + partial_row_height(&rows[r], &starts, &ends, available_h - used)
                        <= available_h
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
    use crate::model::{CellMargins, TextDirection};
    use crate::pdf::layout::TextLine;
    use crate::pdf::table_layout::CellParagraphLayout;

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
        let split = |start, avail| find_cell_split(&cell, &start, avail, false);
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
}

#[cfg(test)]
mod floating_nested_tests {
    use super::*;
    use crate::model::{CellMargins, TextDirection};
    use crate::pdf::layout::TextLine;
    use crate::pdf::table_layout::CellParagraphLayout;

    fn item(height: f32, gap: f32, floating_offset: Option<f32>) -> CellContentItem {
        CellContentItem::NestedTable {
            col_widths: vec![80.0],
            rows: vec![RowLayout {
                height,
                cells: vec![],
                split_min: None,
                keeps_next: false,
            }],
            space_before: gap,
            floating_offset,
            floating_x_offset: None,
            bottom_band: 0.0,
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
            keeps_next: false,
        };
        let split = find_cell_split(&row.cells[0], &CELL_START, 30.0, false);
        assert_eq!(split.item, 1);
        assert!(
            (partial_row_height(&row, &[], &[split.clone()], f32::INFINITY) - 12.0).abs() < 0.001
        );
        assert!(
            (partial_row_height(&row, &[split.clone()], &[], f32::INFINITY) - 48.0).abs() < 0.001
        );
        assert!((partial_row_height(&row, &[], &[], f32::INFINITY) - 66.35).abs() < 0.001);
        assert_eq!(find_cell_split(&row.cells[0], &split, 48.0, false).item, 2);
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
            keeps_next: false,
        };
        let start = CellCursor::at(1, 0);
        assert!(
            (partial_row_height(&row, &[start.clone()], &[], f32::INFINITY) - 70.0).abs() < 0.001
        );
        assert_eq!(find_cell_split(&row.cells[0], &start, 70.01, false).item, 4);
        assert_eq!(find_cell_split(&row.cells[0], &start, 69.9, false).item, 3);
    }
    #[test]
    fn bottom_border_band_belongs_only_to_the_last_line_chunk() {
        let para = CellParagraphLayout {
            lines: vec![TextLine::default(), TextLine::default()],
            line_h: 10.0,
            bottom_band: (1.0, 2.5),
            ..Default::default()
        };
        assert_eq!(cell_lines_h(&para, 0..1), 10.0);
        assert_eq!(cell_lines_h(&para, 1..2), 12.5);
        assert_eq!(para_block_height(&para), 22.5);
        let empty = CellParagraphLayout {
            lines: vec![],
            ..para
        };
        assert_eq!(para_block_height(&empty), 12.5);
    }
}
