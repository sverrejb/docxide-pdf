//! Table border conflicts and cell insets: shared edges, vertical merges, spacing, RTL.

use crate::model::{BorderStyle, CellBorder, CellMargins, TableCell, TableRow, VMerge};

/// Border style precedence for the §17.4.66 conflict tiebreaker: when two
/// borders have equal weight (same width, same source level), the more
/// prominent style wins. Lower number = wins. Double is ranked above Single
/// because our `width` field carries only the nominal stroke width and does
/// not encode the double rule's extra visual weight; everything beats the
/// dotted/dashed family, which is what makes Word collapse a `single`+`dotted`
/// edge to a solid line.
fn border_style_precedence(style: BorderStyle) -> u8 {
    match style {
        BorderStyle::Double
        | BorderStyle::ThinThick
        | BorderStyle::ThickThin
        | BorderStyle::ThinThickThin => 0,
        BorderStyle::Single | BorderStyle::Wave => 1,
        BorderStyle::Dotted => 2,
        BorderStyle::Dashed | BorderStyle::DashSmallGap => 3,
        BorderStyle::DashDot => 4,
        BorderStyle::DashDotDot => 5,
    }
}

fn resolve_h_border(upper_bottom: CellBorder, lower_top: CellBorder) -> CellBorder {
    if !upper_bottom.present {
        return lower_top;
    }
    if !lower_top.present {
        return upper_bottom;
    }
    // §17.4.66 weight ranking comes first: wider wins.
    if upper_bottom.width > lower_top.width + 0.01 {
        return upper_bottom;
    }
    if lower_top.width > upper_bottom.width + 0.01 {
        return lower_top;
    }
    // Equal width: the more prominent style wins (single beats dotted, etc.).
    // Word applies this even when one side is a cell-level override and the other
    // is the inherited table border (e.g. a cell with an explicit `dotted` bottom
    // meeting a neighbor that inherits `insideH single` collapses to solid), so
    // style precedence must be checked BEFORE the cell-vs-table tiebreaker below —
    // otherwise a cell's weaker explicit style would win and render too faint.
    if border_style_precedence(upper_bottom.style) < border_style_precedence(lower_top.style) {
        return upper_bottom;
    }
    if border_style_precedence(lower_top.style) < border_style_precedence(upper_bottom.style) {
        return lower_top;
    }
    // Equal width and style: a cell-level override beats a table-level default.
    if upper_bottom.is_override && !lower_top.is_override {
        return upper_bottom;
    }
    if lower_top.is_override && !upper_bottom.is_override {
        return lower_top;
    }
    // Equal in every respect: prefer upper (first in reading order).
    upper_bottom
}

/// A right-to-left table laid out left to right: each row's cells and the grid
/// reversed, with left and right borders and margins swapped.
/// ponytail: gridBefore keeps its side, and tagging reads the mirrored order
/// (the first-column TH lands on the logically last column).
pub(super) fn mirror_columns(
    rows: &mut [TableRow],
    col_widths: &mut [f32],
    margins: &mut CellMargins,
) {
    col_widths.reverse();
    std::mem::swap(&mut margins.left, &mut margins.right);
    for row in rows {
        row.cells.reverse();
        for cell in &mut row.cells {
            std::mem::swap(&mut cell.borders.left, &mut cell.borders.right);
            if let Some(m) = &mut cell.cell_margins {
                std::mem::swap(&mut m.left, &mut m.right);
            }
        }
    }
}

/// A row's cells all start their text at the row's largest top margin. Word
/// probe: two cells with tcMar top 4pt beside one with none line up; raising
/// the two to 20pt moves the third's text down the same 16pt.
// ponytail: bottom margins untested, left per cell.
pub(super) fn share_row_top_margins(rows: &mut [TableRow], table_margins: CellMargins) {
    for row in rows {
        let top = row
            .cells
            .iter()
            .map(|c| c.cell_margins.unwrap_or(table_margins).top)
            .fold(0.0f32, f32::max);
        for cell in &mut row.cells {
            let mut m = cell.cell_margins.unwrap_or(table_margins);
            if m.top < top {
                m.top = top;
                cell.cell_margins = Some(m);
            }
        }
    }
}

/// Shared edges between rows and vertical merges get one border each, and the
/// cells' insets make room for their border bands.
pub(super) fn settle_row_borders(rows: &mut [TableRow], cell_margins: CellMargins) {
    resolve_h_border_conflicts(rows);
    propagate_vmerge_borders(rows);
    // Word lays cell content between the border bands, not under them: the
    // first line box starts at the top border's lower edge and the next border
    // starts where the content ends. Borders are drawn centred on the row edge,
    // so each side's inset grows by half its band (measured in reference PDFs:
    // rehab_centre 2.25pt rows pitch 3×12.07 + 2.25, case6 0.5pt rows 14.65 + 0.5).
    // Sideways the band doesn't add: the margin only grows to half the band
    // when it is narrower (table probe 6, compat 14 and 15, margins 0/0.5/1/5.4pt
    // under 0.5–6pt borders: text at max(margin, band/2) from the border's
    // centre line, each cell by its own left/right border).
    for cell in rows.iter_mut().flat_map(|r| r.cells.iter_mut()) {
        let b = &cell.borders;
        let (top, bottom, left, right) =
            (b.top.band(), b.bottom.band(), b.left.band(), b.right.band());
        if top > 0.0 || bottom > 0.0 || left > 0.0 || right > 0.0 {
            let mut m = cell.cell_margins.unwrap_or(cell_margins);
            m.top += top / 2.0;
            m.bottom += bottom / 2.0;
            m.left = m.left.max(left / 2.0);
            m.right = m.right.max(right / 2.0);
            cell.cell_margins = Some(m);
        }
    }
}

/// `w:tblCellSpacing` s leaves 2s between neighbouring cells and between the
/// cells and the table's edge (Word probes, compat 14 and 15: a 3-row table
/// grows by 4 × 2s; the first column's text moves 2s right, the second's s).
/// Cell content sees that as wider insets: s on a side facing another cell,
/// 2s on a side facing the table edge. Compat 14's outdent follows the first
/// cell's left inset, so its text stays at the indent as in Word.
// ponytail: borders stay on the grid lines; Word draws each cell's own box
// 2s apart (and adds one border width per gap), so bordered spaced tables are
// still short by a border per row.
pub(super) fn apply_cell_spacing(
    rows: &mut [TableRow],
    spacing: &[f32],
    num_cols: usize,
    base: CellMargins,
) {
    let last_row = rows.len().saturating_sub(1);
    for (ri, row) in rows.iter_mut().enumerate() {
        let s = spacing[ri];
        if s <= 0.0 {
            continue;
        }
        let gap = |at_edge: bool| if at_edge { 2.0 * s } else { s };
        let spans: Vec<_> = row.grid_cells().map(|(col, span, _)| (col, span)).collect();
        for (cell, (col, span)) in row.cells.iter_mut().zip(spans) {
            let mut m = cell.cell_margins.unwrap_or(base);
            m.left += gap(col == 0);
            m.right += gap(col + span >= num_cols);
            m.top += gap(ri == 0);
            m.bottom += gap(ri == last_row);
            cell.cell_margins = Some(m);
        }
    }
}

/// Resolve border conflicts between adjacent rows. When a cell-level override
/// border meets a table-level default at the same horizontal edge, the
/// cell-level border wins per OOXML §17.4.38.
fn resolve_h_border_conflicts(rows: &mut [TableRow]) {
    for ri in 0..rows.len().saturating_sub(1) {
        let (upper, lower) = rows.split_at_mut(ri + 1);
        let upper_row = &mut upper[ri];
        let lower_row = &mut lower[0];
        let mut ug = upper_row.grid_before;
        let mut lg = lower_row.grid_before;
        let mut ui = 0usize;
        let mut li = 0usize;
        while ui < upper_row.cells.len() && li < lower_row.cells.len() {
            let u_span = upper_row.cells[ui].grid_span.max(1) as usize;
            let l_span = lower_row.cells[li].grid_span.max(1) as usize;
            if ug == lg {
                let ub = &upper_row.cells[ui].borders.bottom;
                let lb = &lower_row.cells[li].borders.top;
                let winner = resolve_h_border(*ub, *lb);
                upper_row.cells[ui].borders.bottom = winner;
                let lower = &mut lower_row.cells[li].borders;
                lower.own_top = Some(lower.top);
                lower.top = winner;
            }
            let u_end = ug + u_span;
            let l_end = lg + l_span;
            if u_end <= l_end {
                ug = u_end;
                ui += 1;
            }
            if l_end <= u_end {
                lg = l_end;
                li += 1;
            }
        }
    }
}

/// For vertically merged cells, copy the last continuation cell's bottom
/// border to the restart cell so it draws the correct edge style.
fn propagate_vmerge_borders(rows: &mut [TableRow]) {
    for ri in 0..rows.len() {
        let restarts: Vec<(usize, usize)> = rows[ri]
            .grid_cells()
            .enumerate()
            .filter(|(_, (_, _, c))| c.v_merge == VMerge::Restart)
            .map(|(ci, (g, _, _))| (ci, g))
            .collect();
        for (ci, grid_col) in restarts {
            let last_ri = (ri + 1..rows.len())
                .take_while(|&n| {
                    cell_at(&rows[n], grid_col).is_some_and(|c| c.v_merge == VMerge::Continue)
                })
                .last();
            // The table's bottom border sizes only its last row (`own_bottom`).
            let at_table_end = last_ri == Some(rows.len() - 1);
            if let Some(bottom) = last_ri
                .and_then(|n| cell_at(&rows[n], grid_col))
                .map(|c| c.borders.bottom)
            {
                let borders = &mut rows[ri].cells[ci].borders;
                if at_table_end {
                    borders.own_bottom = Some(borders.bottom);
                }
                borders.bottom = bottom;
            }
        }
    }
}

/// The cell of `row` that starts at `grid_col`, if any.
fn cell_at(row: &TableRow, grid_col: usize) -> Option<&TableCell> {
    row.grid_cells()
        .find(|(g, _, _)| *g == grid_col)
        .map(|(_, _, c)| c)
}

#[cfg(test)]
mod border_conflict_tests {
    use super::*;
    use crate::docx::styles::TableBordersDef;
    use crate::docx::tables::merge_table_borders;
    use crate::model::Block;

    #[test]
    fn row_margin_exceptions_inherit_per_side_and_do_not_leak() {
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
          <w:tblPr><w:tblCellMar><w:top w:w="20"/><w:left w:w="100"/><w:bottom w:w="40"/><w:right w:w="120"/></w:tblCellMar></w:tblPr>
          <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
          <w:tr><w:tblPrEx><w:tblCellMar><w:start w:w="28"/><w:end w:w="32"/></w:tblCellMar></w:tblPrEx>
            <w:tc><w:p/></w:tc><w:tc><w:tcPr><w:tcMar><w:left w:w="0"/><w:bottom w:w="60"/></w:tcMar></w:tcPr><w:p/></w:tc>
          </w:tr><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
        </w:tbl><w:sectPr/></w:body></w:document>"#;
        let doc = crate::docx::parse_document_xml(xml);
        let Block::Table(table) = &doc.sections[0].blocks[0] else {
            panic!("expected table")
        };
        let margins = |r: usize, c: usize| {
            let m = table.rows[r].cells[c]
                .cell_margins
                .unwrap_or(table.cell_margins);
            (m.top, m.left, m.bottom, m.right)
        };
        assert_eq!(margins(0, 0), (1.0, 1.4, 2.0, 1.6));
        assert_eq!(margins(0, 1), (1.0, 0.0, 3.0, 1.6));
        assert_eq!(margins(1, 0), (1.0, 5.0, 2.0, 6.0));
        assert!(table.rows[1].cells[0].cell_margins.is_none());
    }

    #[test]
    fn cell_spacing_widens_insets_double_at_the_table_edge() {
        // 20 twips = 1pt spacing on the table, 0 on the second row.
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
          <w:tblPr><w:tblCellSpacing w:w="20" w:type="dxa"/><w:tblCellMar><w:left w:w="0"/><w:right w:w="0"/></w:tblCellMar></w:tblPr>
          <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
          <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
          <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
          <w:tr><w:trPr><w:tblCellSpacing w:w="0" w:type="dxa"/></w:trPr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
        </w:tbl><w:sectPr/></w:body></w:document>"#;
        let doc = crate::docx::parse_document_xml(xml);
        let Block::Table(table) = &doc.sections[0].blocks[0] else {
            panic!("expected table")
        };
        let margins = |r: usize, c: usize| {
            let m = table.rows[r].cells[c]
                .cell_margins
                .unwrap_or(table.cell_margins);
            (m.top, m.left, m.bottom, m.right)
        };
        assert_eq!(margins(0, 0), (2.0, 2.0, 1.0, 1.0));
        assert_eq!(margins(0, 1), (2.0, 1.0, 1.0, 2.0));
        assert_eq!(margins(1, 1), (1.0, 1.0, 1.0, 2.0));
        assert_eq!(margins(2, 0), (0.0, 0.0, 0.0, 0.0));
    }

    #[test]
    fn default_side_margins_need_a_default_style_or_tbl_cell_mar() {
        let left = |styles: Option<&str>, tbl_pr: &str| {
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr>{tbl_pr}</w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcMar><w:right w:w="108"/></w:tcMar></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#
            );
            let styles = styles.map(|s| {
                format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{s}</w:styles>"#)
            });
            let mut parts = vec![("word/document.xml", xml.as_str())];
            parts.extend(styles.as_deref().map(|s| ("word/styles.xml", s)));
            let doc = crate::docx::parse_parts(&parts);
            let Block::Table(table) = &doc.sections[0].blocks[0] else {
                panic!("expected table")
            };
            (table.cell_margins.left, table.cell_margins.right)
        };
        let default_style = r#"<w:style w:type="table" w:default="1" w:styleId="TableNormal"/>"#;
        let named_style = r#"<w:style w:type="table" w:styleId="Plain"/>"#;
        assert_eq!(left(None, ""), (0.5, 0.5));
        assert_eq!(
            left(None, r#"<w:tblInd w:w="0" w:type="dxa"/>"#),
            (0.5, 0.5)
        );
        assert_eq!(left(Some(named_style), ""), (0.5, 0.5));
        assert_eq!(
            left(Some(named_style), r#"<w:tblStyle w:val="Plain"/>"#),
            (0.5, 0.5)
        );
        assert_eq!(left(Some(default_style), ""), (5.4, 5.4));
        assert_eq!(
            left(None, r#"<w:tblCellMar><w:left w:w="108"/></w:tblCellMar>"#),
            (5.4, 5.4)
        );
    }

    #[test]
    fn side_insets_grow_to_half_the_border_band() {
        let lefts = |tbl_cell_mar: &str| {
            let b = |tag: &str, sz: u32| format!(r#"<w:{tag} w:val="single" w:sz="{sz}"/>"#);
            let borders = [b("left", 48), b("right", 4), b("insideV", 4)].concat();
            let cell = r#"<w:tc><w:p/></w:tc>"#;
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblBorders>{borders}</w:tblBorders>{tbl_cell_mar}</w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>{cell}{cell}</w:tr></w:tbl><w:sectPr/></w:body></w:document>"#
            );
            let doc = crate::docx::parse_document_xml(&xml);
            let Block::Table(table) = &doc.sections[0].blocks[0] else {
                panic!("expected table")
            };
            let cells = &table.rows[0].cells;
            let m = |i: usize| cells[i].cell_margins.unwrap_or(table.cell_margins);
            (m(0).left, m(1).left, m(1).right)
        };
        // 6pt left border: 3pt beats the 0.5pt default; the 0.5pt insideV doesn't.
        assert_eq!(lefts(""), (3.0, 0.5, 0.5));
        assert_eq!(
            lefts(r#"<w:tblCellMar><w:left w:w="0"/><w:right w:w="0"/></w:tblCellMar>"#),
            (3.0, 0.25, 0.25)
        );
        assert_eq!(
            lefts(r#"<w:tblCellMar><w:left w:w="108"/><w:right w:w="108"/></w:tblCellMar>"#),
            (5.4, 5.4, 5.4)
        );
    }

    fn border(style: BorderStyle, is_override: bool) -> CellBorder {
        CellBorder {
            present: true,
            color: None,
            width: 0.5,
            style,
            is_override,
        }
    }

    #[test]
    fn inline_tbl_borders_keep_unspecified_sides_from_style() {
        // croatian_grant_guidelines: inline tblBorders name only the outer sides;
        // insideH/insideV must still come from the Table Grid style.
        let none = CellBorder::default();
        let inline = TableBordersDef {
            top: border(BorderStyle::Single, false),
            bottom: border(BorderStyle::Single, false),
            left: none,
            right: none,
            inside_h: none,
            inside_v: none,
        };
        let style = TableBordersDef {
            top: border(BorderStyle::Dotted, false),
            bottom: none,
            left: none,
            right: none,
            inside_h: border(BorderStyle::Single, false),
            inside_v: border(BorderStyle::Single, false),
        };
        let merged = merge_table_borders(inline, style);
        assert_eq!(merged.top.style, BorderStyle::Single);
        assert!(merged.inside_h.present && merged.inside_v.present);
        assert!(!merged.left.present);
    }

    #[test]
    fn single_beats_dotted_at_equal_width() {
        // Word collapses a dotted+single shared edge to solid (§17.4.66 precedence).
        let upper = border(BorderStyle::Dotted, true);
        let lower = border(BorderStyle::Single, true);
        assert_eq!(resolve_h_border(upper, lower).style, BorderStyle::Single);
        // Order-independent.
        assert_eq!(resolve_h_border(lower, upper).style, BorderStyle::Single);
    }

    #[test]
    fn double_beats_single_at_equal_width() {
        let upper = border(BorderStyle::Single, true);
        let lower = border(BorderStyle::Double, true);
        assert_eq!(resolve_h_border(upper, lower).style, BorderStyle::Double);
    }

    #[test]
    fn wider_border_still_wins_over_precedence() {
        let mut upper = border(BorderStyle::Dotted, true);
        upper.width = 2.0;
        let lower = border(BorderStyle::Single, true);
        assert_eq!(resolve_h_border(upper, lower).style, BorderStyle::Dotted);
    }

    #[test]
    fn override_dotted_loses_to_inherited_single() {
        // The real corpus case (turkish_chemistry): a cell with an explicit dotted
        // bottom (override) meets a neighbor that inherits `insideH single` (not an
        // override). Word renders this solid — style precedence must beat the
        // cell-vs-table override rule at equal width.
        let upper = border(BorderStyle::Dotted, true);
        let lower = border(BorderStyle::Single, false);
        assert_eq!(resolve_h_border(upper, lower).style, BorderStyle::Single);
        assert_eq!(resolve_h_border(lower, upper).style, BorderStyle::Single);
    }

    #[test]
    fn equal_style_prefers_upper() {
        let upper = border(BorderStyle::Single, true);
        let lower = border(BorderStyle::Single, true);
        // Same everything: upper (first in reading order) is kept.
        assert!(resolve_h_border(upper, lower).present);
    }
}
