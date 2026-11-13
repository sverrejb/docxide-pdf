use std::io::{Read, Seek};

use crate::model::{
    Alignment, Block, CellBorder, CellBorders, CellMargins, CellVAlign, FitText, HatchPattern,
    HorizontalPosition, LineSpacing, Paragraph, Run, Table, TableAlignment, TableCell,
    TablePosition, TableRow, TextDirection, VMerge,
};

use super::numbering::{ListCounters, parse_list_info};
use super::runs::parse_runs;
use super::styles::{TableBordersDef, TableStyleDef, parse_table_borders_def};
use super::table_borders::{
    apply_cell_spacing, mirror_columns, settle_row_borders, share_row_top_margins,
};
use super::{
    ParseContext, WML_NS, collect_block_nodes, parse_cell_border, parse_cell_border_left,
    parse_cell_border_right, parse_hex_color, parse_on_off, parse_paragraph_spacing, twips_attr,
    wml, wml_attr, wml_bool,
};

/// Approximate a `w:shd` stripe/cross pattern as a solid color for render
/// paths that don't hatch. The pattern ink (`w:color`) defaults to black;
/// named patterns blend by an estimated coverage (`solid`/`pctNN` go through
/// `shd_color`).
fn approx_pattern_shade(val: &str, color: Option<[u8; 3]>) -> Option<[u8; 3]> {
    let coverage = match val {
        "thinHorzStripe" | "thinVertStripe" | "thinDiagStripe" | "thinReverseDiagStripe" => 0.30,
        "horzStripe" | "vertStripe" | "diagStripe" | "reverseDiagStripe" => 0.45,
        "thinHorzCross" | "thinDiagCross" => 0.40,
        "horzCross" | "diagCross" => 0.55,
        _ => return None,
    };
    let fg = color.unwrap_or([0, 0, 0]);
    let blend =
        |bg: u8, ink: u8| (bg as f32 * (1.0 - coverage) + ink as f32 * coverage).round() as u8;
    Some([blend(255, fg[0]), blend(255, fg[1]), blend(255, fg[2])])
}

/// Map a `w:shd` line/cross pattern value to a `HatchKind`. Returns `None` for
/// `pctNN`, `clear`, `solid`, `nil`, etc. — only the geometric stripe/cross
/// patterns are drawn as hatching.
fn hatch_kind(val: &str) -> Option<crate::model::HatchKind> {
    use crate::model::HatchKind::*;
    Some(match val {
        "thinHorzStripe" | "horzStripe" => Horz,
        "thinVertStripe" | "vertStripe" => Vert,
        "thinDiagStripe" | "diagStripe" => DiagFwd,
        "thinReverseDiagStripe" | "reverseDiagStripe" => DiagBack,
        "thinHorzCross" | "horzCross" => CrossHorzVert,
        "thinDiagCross" | "diagCross" => CrossDiag,
        _ => return None,
    })
}

pub(super) fn margin_twips(mar: roxmltree::Node, primary: &str, fallback: &str) -> Option<f32> {
    wml(mar, primary)
        .or_else(|| wml(mar, fallback))
        .and_then(|n| twips_attr(n, "w"))
}

fn merge_cell_margins(mar: roxmltree::Node, base: CellMargins) -> CellMargins {
    CellMargins {
        top: wml(mar, "top")
            .and_then(|n| twips_attr(n, "w"))
            .unwrap_or(base.top),
        left: margin_twips(mar, "left", "start").unwrap_or(base.left),
        bottom: wml(mar, "bottom")
            .and_then(|n| twips_attr(n, "w"))
            .unwrap_or(base.bottom),
        right: margin_twips(mar, "right", "end").unwrap_or(base.right),
    }
}

/// `w:vAlign` under a `w:tcPr`.
pub(super) fn cell_v_align(tc_pr: roxmltree::Node) -> Option<CellVAlign> {
    match wml_attr(tc_pr, "vAlign")? {
        "center" => Some(CellVAlign::Center),
        "bottom" => Some(CellVAlign::Bottom),
        "top" => Some(CellVAlign::Top),
        _ => None,
    }
}

/// `w:jc` under a `w:tblPr` or `w:trPr`.
pub(super) fn table_jc(pr: roxmltree::Node) -> Option<TableAlignment> {
    wml_attr(pr, "jc").map(|val| match val {
        "center" => TableAlignment::Center,
        "right" | "end" => TableAlignment::Right,
        _ => TableAlignment::Left,
    })
}

/// Per-side merge: sides `over` specifies (or explicitly clears) win, the rest
/// fall through to `base`.
pub(super) fn merge_table_borders(over: TableBordersDef, base: TableBordersDef) -> TableBordersDef {
    TableBordersDef {
        top: border_or_fallback(over.top, base.top),
        bottom: border_or_fallback(over.bottom, base.bottom),
        left: border_or_fallback(over.left, base.left),
        right: border_or_fallback(over.right, base.right),
        inside_h: border_or_fallback(over.inside_h, base.inside_h),
        inside_v: border_or_fallback(over.inside_v, base.inside_v),
    }
}

fn border_or_fallback(inline: CellBorder, fallback: CellBorder) -> CellBorder {
    if inline.present || inline.is_override {
        CellBorder {
            is_override: true,
            ..inline
        }
    } else {
        fallback
    }
}

/// Conditional table-style formatting (`w:tblStylePr`) accumulated for one
/// cell; the caller overlays regions in spec order (bands, first/last
/// row/column, corners), later ones winning.
#[derive(Default)]
struct CondFormat {
    borders: CellBorders,
    shading: Option<[u8; 3]>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<[u8; 3]>,
    font_size: Option<f32>,
    font_name: Option<String>,
}

impl CondFormat {
    /// Overlay the `key` region. Per OOXML §17.4.23 its top/bottom/left/right
    /// borders apply only where the cell sits on that edge of the region
    /// (`edges` = (top, bottom, left, right)); elsewhere insideH/insideV do.
    fn apply(&mut self, style: &TableStyleDef, key: &str, edges: (bool, bool, bool, bool)) {
        let Some(cond) = style.conditionals.get(key) else {
            return;
        };
        if let Some(cb) = &cond.borders {
            let (top, bottom, left, right) = edges;
            let ct = if top { cb.top } else { cb.inside_h };
            let cb_b = if bottom { cb.bottom } else { cb.inside_h };
            let cl = if left { cb.left } else { cb.inside_v };
            let cr = if right { cb.right } else { cb.inside_v };
            self.borders.top = border_or_fallback(ct, self.borders.top);
            self.borders.bottom = border_or_fallback(cb_b, self.borders.bottom);
            self.borders.left = border_or_fallback(cl, self.borders.left);
            self.borders.right = border_or_fallback(cr, self.borders.right);
        }
        self.shading = cond.shading.or(self.shading);
        self.bold = cond.bold.or(self.bold);
        self.italic = cond.italic.or(self.italic);
        self.color = cond.color.or(self.color);
        self.font_size = cond.font_size.or(self.font_size);
        self.font_name = cond.font_name.clone().or(self.font_name.take());
    }
}

pub(in crate::docx) fn parse_table_node<R: Read + Seek>(
    node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    lists: &mut ListCounters,
) -> Table {
    let mut col_widths: Vec<f32> = wml(node, "tblGrid")
        .into_iter()
        .flat_map(|grid| grid.children())
        .filter(|n| n.has_tag_name((WML_NS, "gridCol")))
        .filter_map(|n| twips_attr(n, "w"))
        .collect();

    let tbl_pr = wml(node, "tblPr");
    let table_indent_opt = tbl_pr
        .and_then(|pr| wml(pr, "tblInd"))
        .and_then(|ind| twips_attr(ind, "w"));
    let table_indent = table_indent_opt.unwrap_or(0.0);
    let table_indent_explicit = table_indent_opt.is_some();

    let fixed_layout = tbl_pr
        .and_then(|pr| wml(pr, "tblLayout"))
        .and_then(|n| n.attribute((WML_NS, "type")))
        .is_some_and(|v| v == "fixed");

    let auto_width = tbl_pr.and_then(|pr| wml(pr, "tblW")).is_none_or(|n| {
        n.attribute((WML_NS, "type")).is_none_or(|t| t == "auto")
            || twips_attr(n, "w").is_none_or(|w| w <= 0.0)
    });

    // ST_MeasurementOrPercent: pct values are in 5000ths ("5000" = 100%),
    // or the literal "NN%" text form.
    let width_pct = tbl_pr
        .and_then(|pr| wml(pr, "tblW"))
        .filter(|n| n.attribute((WML_NS, "type")) == Some("pct"))
        .and_then(|n| n.attribute((WML_NS, "w")))
        .and_then(|raw| {
            if let Some(stripped) = raw.strip_suffix('%') {
                stripped.trim().parse::<f32>().ok().map(|p| p / 100.0)
            } else {
                raw.parse::<f32>().ok().map(|v| v / 5000.0)
            }
        })
        .filter(|p| *p > 0.0);

    let tbl_style = tbl_pr
        .and_then(|pr| wml_attr(pr, "tblStyle"))
        .and_then(|id| ctx.styles.table_styles.get(id));
    // The table style and its basedOn ancestors, nearest first.
    let style_chain = || {
        std::iter::successors(tbl_style, |s| {
            s.based_on
                .as_deref()
                .and_then(|id| ctx.styles.table_styles.get(id))
        })
        .take(8)
    };

    // The table's own jc, else its style's, else the first row's trPr/jc.
    let alignment = tbl_pr
        .and_then(table_jc)
        .or_else(|| style_chain().find_map(|s| s.jc))
        .unwrap_or_else(|| {
            collect_block_nodes(node, ctx.revisions.markup)
                .into_iter()
                .find(|n| n.has_tag_name((WML_NS, "tr")))
                .and_then(|tr| wml(tr, "trPr"))
                .and_then(table_jc)
                .unwrap_or_default()
        });

    // Each side: the table's own, then its style's (estonian_community's
    // TableGrid zeroes them), then Word's default.
    let own_mar = tbl_pr.and_then(|pr| wml(pr, "tblCellMar"));
    let word_default_margins = ctx.styles.has_default_table_style || own_mar.is_some();
    // Word's 5.4pt default comes with a table style: a table with neither a
    // style nor its own tblCellMar gets 0.5pt side margins and so no visible
    // outdent (Word probes on case200/case213 with cell borders: the border
    // at the margin, the text 0.5pt inside it and a right-aligned line
    // ending ~0.5pt short of the right border; case131's nested table text
    // 1.0pt inside the parent's, its second column 0.58pt inside; an empty
    // default table style brings the 5.4pt back).
    // It is the document's default table style that brings it, not the
    // table's own: table-style-centered's "Centered" style, without
    // tblCellMar in a document without a default table style, gets 0.5pt.
    let default_side = if word_default_margins { 5.4 } else { 0.5 };
    let side = |i: usize, a: &str, b: &str, default: f32| {
        own_mar
            .and_then(|mar| margin_twips(mar, a, b))
            .or_else(|| style_chain().find_map(|s| s.cell_margins[i]))
            .unwrap_or(default)
    };
    let mut cell_margins = CellMargins {
        top: side(0, "top", "top", 0.0),
        left: side(1, "left", "start", default_side),
        bottom: side(2, "bottom", "bottom", 0.0),
        right: side(3, "right", "end", default_side),
    };
    let style_space_before = style_chain().find_map(|s| s.space_before);
    let style_space_after = style_chain().find_map(|s| s.space_after);
    let style_line_spacing = style_chain().find_map(|s| s.line_spacing);
    let style_alignment = style_chain().find_map(|s| s.alignment);
    let style_v_align = style_chain().find_map(|s| s.v_align);

    let table_position = tbl_pr.and_then(|pr| wml(pr, "tblpPr")).map(|tblp| {
        let v_anchor = match tblp.attribute((WML_NS, "vertAnchor")) {
            Some("page") => "page",
            Some("text") => "text",
            _ => "margin",
        };
        let h_anchor = match tblp.attribute((WML_NS, "horzAnchor")) {
            Some("page") => "page",
            Some("margin") => "margin",
            _ => "column",
        };
        let v_offset_pt = twips_attr(tblp, "tblpY").unwrap_or(0.0);
        let h_position = match tblp.attribute((WML_NS, "tblpXSpec")) {
            Some("center") => HorizontalPosition::AlignCenter,
            Some("right") => HorizontalPosition::AlignRight,
            Some(_) => HorizontalPosition::AlignLeft,
            None => {
                let offset = twips_attr(tblp, "tblpX").unwrap_or(0.0);
                HorizontalPosition::Offset(offset)
            }
        };
        let top_from_text = twips_attr(tblp, "topFromText").unwrap_or(0.0);
        let bottom_from_text = twips_attr(tblp, "bottomFromText").unwrap_or(0.0);
        let left_from_text = twips_attr(tblp, "leftFromText").unwrap_or(0.0);
        let right_from_text = twips_attr(tblp, "rightFromText").unwrap_or(0.0);
        TablePosition {
            allow_overlap: tbl_pr
                .and_then(|pr| wml(pr, "tblOverlap"))
                .and_then(|n| n.attribute((WML_NS, "val")))
                != Some("never"),
            h_position,
            h_anchor,
            v_offset_pt,
            v_anchor,
            top_from_text,
            bottom_from_text,
            left_from_text,
            right_from_text,
        }
    });

    let tbl_style_borders = tbl_style.and_then(|s| s.base_borders.as_ref());
    // A bordered table style without paragraph spacing sets its cells single
    // with no space only over Word's built-in defaults; a document's own
    // pPrDefault reaches the cells (a resume's 1.2-line docDefaults step its
    // cell lines 14.64pt in Word).
    let has_tbl_style = tbl_style_borders.is_some() && ctx.styles.defaults.builtin_para_defaults;

    let inline_tbl_borders = tbl_pr
        .and_then(|pr| wml(pr, "tblBorders"))
        .map(parse_table_borders_def);
    // A table's own tblBorders overrides the style side by side (§17.4.39); sides
    // it leaves out still come from the style. croatian_grant_guidelines sets only
    // the outer borders inline and takes insideH/insideV from Table Grid.
    let merged_tbl_borders = match (inline_tbl_borders, tbl_style_borders) {
        (Some(inline), Some(style)) => Some(merge_table_borders(inline, *style)),
        (inline, style) => inline.or(style.copied()),
    };

    // Parse tblLook — controls which conditional formats from the style apply.
    // Supports both named attributes (w:firstRow="1") and legacy hex bitmask (w:val="04A0").
    let tbl_look_node = tbl_pr.and_then(|pr| wml(pr, "tblLook"));
    let look_flag = |attr: &str, bit: u32| -> bool {
        tbl_look_node
            .and_then(|n| n.attribute((WML_NS, attr)))
            .map(parse_on_off)
            .unwrap_or_else(|| {
                tbl_look_node
                    .and_then(|n| n.attribute((WML_NS, "val")))
                    .and_then(|v| u32::from_str_radix(v, 16).ok())
                    .is_some_and(|mask| mask & bit != 0)
            })
    };
    let look_first_row = look_flag("firstRow", 0x0020);
    let look_last_row = look_flag("lastRow", 0x0040);
    let look_first_col = look_flag("firstColumn", 0x0080);
    let look_last_col = look_flag("lastColumn", 0x0100);
    let look_no_h_band = look_flag("noHBand", 0x0200);
    let look_no_v_band = look_flag("noVBand", 0x0400);

    let tbl_rows: Vec<_> = collect_block_nodes(node, ctx.revisions.markup)
        .into_iter()
        .filter(|n| n.has_tag_name((WML_NS, "tr")))
        // The final view drops a deleted row (Word online export probe
        // 2026-11-05); the markup view keeps it, struck through.
        .filter(|n| ctx.revisions.markup || wml(*n, "trPr").and_then(|pr| wml(pr, "del")).is_none())
        .collect();

    // OOXML §17.4.48 requires tblGrid, but some generators (e.g. SpecLink)
    // omit it. Infer the grid from the widest row's cell widths so layout
    // has real columns to work with, splitting spanned cells evenly.
    let grid_inferred = col_widths.is_empty() && !tbl_rows.is_empty();
    let row_cells: Vec<Vec<_>> = tbl_rows
        .iter()
        .map(|tr| {
            collect_block_nodes(*tr, ctx.revisions.markup)
                .into_iter()
                .filter(|n| n.has_tag_name((WML_NS, "tc")))
                .collect()
        })
        .collect();
    let cell_widths: Vec<_> = row_cells
        .iter()
        .flatten()
        .map(|tc| wml(*tc, "tcPr").and_then(|pr| wml(pr, "tcW")))
        .collect();
    let cells_unsized = cell_widths.iter().all(|w| {
        w.is_none_or(|w| {
            matches!(w.attribute((WML_NS, "type")), Some("auto" | "nil"))
                || twips_attr(w, "w").is_none_or(|v| v == 0.0)
        })
    });
    let cells_twip_sized = cell_widths.iter().all(|w| {
        w.is_some_and(|w| {
            matches!(
                w.attribute((WML_NS, "type")),
                None | Some("dxa" | "auto" | "nil")
            )
        })
    });
    if col_widths.is_empty() {
        for cells in &row_cells {
            let mut row_widths: Vec<f32> = Vec::new();
            for tc in cells {
                let tc_pr = wml(*tc, "tcPr");
                let w = tc_pr
                    .and_then(|pr| wml(pr, "tcW"))
                    .and_then(|n| twips_attr(n, "w"))
                    .unwrap_or(72.0);
                let span = tc_pr
                    .and_then(|pr| wml_attr(pr, "gridSpan"))
                    .and_then(|v| v.parse::<u16>().ok())
                    .unwrap_or(1)
                    .max(1) as usize;
                for _ in 0..span {
                    row_widths.push(w / span as f32);
                }
            }
            if row_widths.len() > col_widths.len() {
                col_widths = row_widths;
            }
        }
    }

    let num_rows = tbl_rows.len();
    let num_cols = col_widths.len();

    let cell_spacing = |pr: Option<roxmltree::Node>| {
        pr.and_then(|pr| wml(pr, "tblCellSpacing"))
            .and_then(|n| twips_attr(n, "w"))
    };
    let table_spacing = cell_spacing(tbl_pr).unwrap_or(0.0);
    let mut row_spacing = Vec::with_capacity(num_rows);
    let mut rows = Vec::new();
    for (ri, tr) in tbl_rows.iter().enumerate() {
        let tr_pr = wml(*tr, "trPr");
        row_spacing.push(cell_spacing(tr_pr).unwrap_or(table_spacing));
        let (row_height, height_exact) = tr_pr
            .and_then(|pr| wml(pr, "trHeight"))
            .map(|h| {
                let val = twips_attr(h, "val");
                let exact = h.attribute((WML_NS, "hRule")) == Some("exact");
                (val, exact)
            })
            .unwrap_or((None, false));
        let is_header = tr_pr.and_then(|pr| wml(pr, "tblHeader")).is_some();
        let cant_split = tr_pr
            .and_then(|pr| wml_bool(pr, "cantSplit"))
            .unwrap_or(false);
        let grid_before = tr_pr
            .and_then(|pr| wml_attr(pr, "gridBefore"))
            .and_then(|v| v.parse::<u16>().ok())
            .map_or(0, usize::from);

        // Per-row table property exceptions (§17.4.60): merge with base table
        // borders — specified exception borders override, unspecified inherit.
        let row_effective_tbl_borders =
            match wml(*tr, "tblPrEx").and_then(|prex| wml(prex, "tblBorders")) {
                Some(bdr_node) => {
                    let exc = parse_table_borders_def(bdr_node);
                    Some(merged_tbl_borders.map_or(exc, |base| merge_table_borders(exc, base)))
                }
                None => merged_tbl_borders,
            };

        // Row exceptions override table defaults per side; explicit tcMar wins.
        let row_cell_margins = wml(*tr, "tblPrEx")
            .and_then(|pr| wml(pr, "tblCellMar"))
            .map(|mar| merge_cell_margins(mar, cell_margins));
        let mut cells = Vec::new();
        let mut grid_col = grid_before;
        for &tc in &row_cells[ri] {
            let ci = grid_col;
            let tc_pr = wml(tc, "tcPr");
            let tc_w = tc_pr.and_then(|pr| wml(pr, "tcW"));
            let cell_width = tc_w
                .and_then(|w| twips_attr(w, "w"))
                .unwrap_or_else(|| col_widths.get(ci).copied().unwrap_or(72.0));
            let cell_width_pct = tc_w
                .filter(|w| w.attribute((WML_NS, "type")) == Some("pct"))
                .and_then(|w| w.attribute((WML_NS, "w")))
                .and_then(|v| v.trim_end_matches('%').parse::<f32>().ok())
                .map(|v| v / 5000.0);

            // Per OOXML §17.4.17, absent gridSpan defaults to 1.
            // Word strictly honours this — it never infers larger spans
            // from tcW. Rows with fewer cells than grid columns simply
            // leave the trailing columns empty.
            let grid_span = tc_pr
                .and_then(|pr| wml_attr(pr, "gridSpan"))
                .and_then(|v| v.parse::<u16>().ok())
                .unwrap_or(1);

            let v_merge = tc_pr
                .and_then(|pr| wml(pr, "vMerge"))
                .map(|n| match n.attribute((WML_NS, "val")) {
                    Some("restart") => VMerge::Restart,
                    _ => VMerge::Continue,
                })
                .unwrap_or(VMerge::None);

            let v_align = tc_pr
                .and_then(cell_v_align)
                .or(style_v_align)
                .unwrap_or(CellVAlign::Top);

            let text_direction = match tc_pr.and_then(|pr| wml_attr(pr, "textDirection")) {
                Some("tbRlV" | "tbRl" | "rlV" | "rl" | "tbV" | "tb") => TextDirection::TbRl,
                Some("btLr" | "lr" | "lrV" | "lrTbV") => TextDirection::BtLr,
                _ => TextDirection::LrTb,
            };
            let hide_mark = tc_pr
                .and_then(|pr| wml_bool(pr, "hideMark"))
                .unwrap_or(false);

            // Word shows only a vertically merged cell's first part; the
            // numbered paragraphs of its continuations don't count either:
            // nabl's checklist numbers its sections 13, 14, … though every
            // row's merged first cell carries a numbered paragraph.
            let mut hidden_lists = ListCounters::default();
            let cell_lists: &mut ListCounters = if v_merge == VMerge::Continue {
                &mut hidden_lists
            } else {
                &mut *lists
            };

            let span_end = ci + grid_span as usize;

            // Base style borders (position-aware: outer vs inner)
            let style_borders = row_effective_tbl_borders.map(|tb| CellBorders {
                top: if ri == 0 { tb.top } else { tb.inside_h },
                bottom: if ri == num_rows - 1 {
                    tb.bottom
                } else {
                    tb.inside_h
                },
                left: if ci == 0 { tb.left } else { tb.inside_v },
                right: if span_end >= num_cols {
                    tb.right
                } else {
                    tb.inside_v
                },
                own_top: None,
                own_bottom: None,
            });

            // Apply conditional formatting overrides from tblStylePr.
            // Order per spec: wholeTable → bands → first/last row/col → corners.
            //
            // Per OOXML §17.4.23, tblStylePr borders use inside/outside semantics:
            // top/bottom/left/right are the outer edges of the conditional region,
            // insideH/insideV are borders between cells within the region.
            let mut cond = CondFormat {
                borders: style_borders.unwrap_or_default(),
                ..CondFormat::default()
            };
            if let Some(style_def) = tbl_style {
                let is_first_col = ci == 0;
                let is_last_col = span_end >= num_cols;
                let is_first_row = ri == 0;
                let is_last_row = ri == num_rows - 1;
                // The first/last regions the cell is in and the tblLook turns on.
                let first_row = look_first_row && is_first_row;
                let last_row = look_last_row && is_last_row;
                let first_col = look_first_col && is_first_col;
                let last_col = look_last_col && is_last_col;
                // Row banding — skip rows consumed by firstRow/lastRow
                if !look_no_h_band && !first_row && !last_row {
                    let band_row = if look_first_row { ri - 1 } else { ri };
                    let key = if band_row % 2 == 0 {
                        "band1Horz"
                    } else {
                        "band2Horz"
                    };
                    // Row bands: single row, so top/bottom are always edges
                    cond.apply(style_def, key, (true, true, is_first_col, is_last_col));
                }
                // Column banding — skip cols consumed by firstCol/lastCol
                if !look_no_v_band && !first_col && !last_col {
                    let band_col = if look_first_col { ci - 1 } else { ci };
                    let key = if band_col % 2 == 0 {
                        "band1Vert"
                    } else {
                        "band2Vert"
                    };
                    // Column bands: single column, so left/right are always edges
                    cond.apply(style_def, key, (is_first_row, is_last_row, true, true));
                }
                // Row regions: top/bottom are edges, left/right depend on the
                // column; column regions the other way round; a corner cell
                // is a single cell, all edges. Spec order, later ones winning.
                let row = (true, true, is_first_col, is_last_col);
                let col = (is_first_row, is_last_row, true, true);
                let all = (true, true, true, true);
                for (on, key, edges) in [
                    (first_row, "firstRow", row),
                    (last_row, "lastRow", row),
                    (first_col, "firstCol", col),
                    (last_col, "lastCol", col),
                    (first_row && first_col, "nwCell", all),
                    (first_row && last_col, "neCell", all),
                    (last_row && first_col, "swCell", all),
                    (last_row && last_col, "seCell", all),
                ] {
                    if on {
                        cond.apply(style_def, key, edges);
                    }
                }
            }

            // Inline cell borders override conditional/style borders
            let borders = tc_pr
                .and_then(|pr| wml(pr, "tcBorders"))
                .map(|bdr| CellBorders {
                    top: border_or_fallback(parse_cell_border(bdr, "top"), cond.borders.top),
                    bottom: border_or_fallback(
                        parse_cell_border(bdr, "bottom"),
                        cond.borders.bottom,
                    ),
                    left: border_or_fallback(parse_cell_border_left(bdr), cond.borders.left),
                    right: border_or_fallback(parse_cell_border_right(bdr), cond.borders.right),
                    own_top: None,
                    own_bottom: None,
                })
                .unwrap_or(cond.borders);

            // Inline shading overrides conditional shading. An explicit fill
            // color wins; otherwise a pattern fill (stripes/cross/pctNN) with an
            // auto fill is approximated as a solid tint, since the renderer only
            // supports solid cell fills.
            let has_explicit_shd = tc_pr.and_then(|pr| wml(pr, "shd")).is_some();
            let mut cell_hatch: Option<HatchPattern> = None;
            let mut pattern_shading: Option<[u8; 3]> = None;
            if let Some(shd) = tc_pr.and_then(|pr| wml(pr, "shd")) {
                let val = shd.attribute((WML_NS, "val")).unwrap_or("clear");
                let fill = shd.attribute((WML_NS, "fill")).unwrap_or("auto");
                let bg = if fill != "none" && fill != "auto" {
                    parse_hex_color(fill)
                } else {
                    None
                };
                let ink = shd.attribute((WML_NS, "color")).and_then(parse_hex_color);
                if let Some(kind) = hatch_kind(val) {
                    cell_hatch = Some(HatchPattern {
                        kind,
                        fg: ink.unwrap_or([0, 0, 0]),
                        bg: bg.unwrap_or([255, 255, 255]),
                    });
                    // Solid fallback for any render path that doesn't hatch.
                    pattern_shading = approx_pattern_shade(val, ink).or(bg);
                } else {
                    pattern_shading = super::shd_color(shd);
                }
            }
            // A direct cell <w:shd> overrides table-style conditional banding —
            // even fill="auto" (explicit no-fill) suppresses a band1Horz/Vert
            // tint. Only fall back to the conditional shading when the cell has
            // no explicit shd at all. (Pendulum data table: every cell is
            // shd fill="auto", so Word renders it white despite PlainTable1's
            // grey row banding.)
            let shading = if has_explicit_shd {
                pattern_shading
            } else {
                cond.shading
            };
            let hatch = cell_hatch;

            let per_cell_margins = tc_pr
                .and_then(|pr| wml(pr, "tcMar"))
                .map(|mar| merge_cell_margins(mar, row_cell_margins.unwrap_or(cell_margins)))
                .or(row_cell_margins);

            let mut cell_blocks: Vec<Block> = Vec::new();
            let block_nodes = collect_block_nodes(tc, ctx.revisions.markup);
            ctx.fields = Default::default();
            for n in &block_nodes {
                if n.has_tag_name((WML_NS, "p")) {
                    let p = *n;
                    let parsed = parse_runs(p, ctx);
                    let mut runs = parsed.runs;
                    // Apply table style rPr: conditional > base, only when
                    // the run inherited from doc defaults (not set explicitly).
                    let eff_tbl_font_size = cond
                        .font_size
                        .or_else(|| tbl_style.and_then(|s| s.base_font_size));
                    let eff_tbl_font_name = cond
                        .font_name
                        .as_deref()
                        .or_else(|| tbl_style.and_then(|s| s.base_font_name.as_deref()));
                    let eff_tbl_bold = cond.bold.or_else(|| tbl_style.and_then(|s| s.base_bold));
                    let eff_tbl_italic = cond
                        .italic
                        .or_else(|| tbl_style.and_then(|s| s.base_italic));

                    let mut has_text = false;
                    let mut has_inline_images = false;
                    for run in &mut runs {
                        if let Some(tfs) = eff_tbl_font_size
                            && run.font_size_from_default
                        {
                            run.font_size = tfs;
                        }
                        if let Some(tfn) = eff_tbl_font_name
                            && run.font_name_from_default
                        {
                            run.font_name = tfn.to_string();
                        }
                        if eff_tbl_bold == Some(true) && !run.bold_is_direct {
                            run.bold = true;
                        }
                        if eff_tbl_italic == Some(true) && !run.italic_is_direct {
                            run.italic = true;
                        }
                        if let Some(cc) = cond.color
                            && run.color.is_none()
                        {
                            run.color = Some(cc);
                        }
                        if !run.text.is_empty() || run.is_tab {
                            has_text = true;
                        }
                        if run.inline_image.is_some() {
                            has_inline_images = true;
                        }
                    }
                    // Underlined pictures stay picture runs and hidden ones
                    // are not shown, as in a body paragraph.
                    let (para_image, content_height) = if !has_text
                        && runs.iter().any(|r| r.shown_image().is_some())
                        && !runs.iter().any(Run::underlined_picture)
                    {
                        crate::model::take_shown_image(&mut runs)
                    } else if !has_text && !has_inline_images {
                        (None, super::images::inline_drawing_height(p))
                    } else {
                        (None, 0.0)
                    };
                    let ppr = wml(p, "pPr");
                    let (para_style_id, para_style) = ctx.styles.paragraph_style_of(ppr);
                    let alignment = super::runs::paragraph_alignment(p, ppr, para_style)
                        .or(style_alignment)
                        .unwrap_or(Alignment::Left);
                    let (sp_before, sp_after, ls) =
                        parse_paragraph_spacing(ppr, para_style, &ctx.styles.defaults);
                    let (space_before_auto, space_after_auto) =
                        super::autospacing(ppr, para_style, &ctx.styles.defaults);
                    let line_spacing = ls
                        .or(style_line_spacing)
                        .or_else(|| has_tbl_style.then_some(LineSpacing::Auto(1.0)));
                    let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                    let style_num = para_style.and_then(|s| s.num_id.as_deref());
                    let style_ilvl = para_style.and_then(|s| s.num_ilvl);
                    let mark = super::runs::paragraph_mark_run(ppr, para_style, ctx);
                    let numbering = parse_list_info(
                        num_pr,
                        style_num,
                        style_ilvl,
                        Some(para_style_id),
                        &ctx.styles.paragraph_styles,
                        ctx.numbering,
                        cell_lists,
                        Some(&mark),
                        super::numbering::item_revision(p, ctx.revisions),
                    );
                    let (mut indents, indent_chars) = super::paragraph::resolve_indents(
                        ppr,
                        para_style,
                        &numbering,
                        &ctx.styles.defaults,
                    );
                    super::apply_indent_chars(
                        &mut indents,
                        indent_chars,
                        super::paragraph::indent_char_widths(
                            &runs,
                            para_style,
                            &ctx.styles,
                            ctx.grid_char_delta,
                        ),
                    );
                    let (indent_left, indent_right, indent_hanging, indent_first_line) = indents;
                    let num_tab_stop =
                        super::paragraph::num_tab_override(ppr).or(numbering.tab_stop);
                    let (auto_space_de, auto_space_dn) =
                        super::paragraph::auto_space(ppr, para_style);
                    let space_before =
                        sp_before
                            .or(style_space_before)
                            .unwrap_or(if has_tbl_style {
                                0.0
                            } else {
                                ctx.styles.defaults.space_before
                            });
                    let space_after = sp_after.or(style_space_after).unwrap_or(if has_tbl_style {
                        0.0
                    } else {
                        ctx.styles.defaults.space_after
                    });
                    // The end-of-cell mark uses the document's default
                    // paragraph style. Contextual after-spacing collapses only
                    // when the last paragraph has that same style.
                    let contextual_spacing = super::paragraph::contextual_spacing(ppr, para_style);
                    let space_after = if contextual_spacing
                        && para_style_id == ctx.styles.default_paragraph_style_id
                        && p.next_siblings().skip(1).all(|node| !node.is_element())
                    {
                        0.0
                    } else {
                        space_after
                    };
                    let tab_stops = super::resolve_tab_stops(ppr, para_style);
                    cell_blocks.push(Block::Paragraph(Paragraph {
                        widow_control: super::paragraph::widow_control(
                            ppr,
                            para_style,
                            &ctx.styles.defaults,
                        ),
                        keep_lines: super::paragraph::keep_lines(ppr, para_style),
                        runs,
                        alignment,
                        borders: super::paragraph::borders(ppr, para_style),
                        space_before_auto,
                        space_after_auto,
                        indent_left,
                        indent_right,
                        indent_hanging,
                        indent_first_line,
                        num_level_tab_stop: num_tab_stop,
                        no_indent_num_tab: ctx.styles.defaults.no_indent_as_num_tab,
                        horizontal_rule: parsed.horizontal_rule,
                        field_code_mark: parsed.mark_in_field_code,
                        line_spacing,
                        space_before,
                        space_after,
                        image: para_image,
                        content_height,
                        snap_to_grid: super::paragraph::snap_to_grid(ppr, para_style),
                        text_alignment: super::paragraph::text_alignment(
                            ppr,
                            para_style,
                            &ctx.styles.defaults,
                        ),
                        // Cells space Latin text from ideographs as the body
                        // does (chinese_costume's "5-10 分钟").
                        auto_space_de,
                        auto_space_dn,
                        floating_images: parsed.floating_images,
                        textboxes: parsed.textboxes,
                        connectors: parsed.connectors,
                        style_id: Some(para_style_id.to_string()),
                        contextual_spacing,
                        keep_next: super::paragraph::keep_next(ppr, para_style),
                        paragraph_mark_font_size: Some(mark.font_size),
                        paragraph_mark_font_name: Some(mark.font_name),
                        paragraph_mark_position: super::paragraph::paragraph_mark_position(
                            ppr.and_then(|ppr| wml(ppr, "rPr")),
                            para_style,
                        ),
                        tab_stops,
                        frame_props: super::parse_frame_props(
                            ppr,
                            para_style.and_then(|s| s.frame_attrs.as_ref()),
                        ),
                        ..numbering.into_paragraph()
                    }));
                    super::paragraph::merge_field_code_paragraphs(&mut cell_blocks);
                } else if n.has_tag_name((WML_NS, "tbl")) {
                    ctx.fields = Default::default();
                    let nested = parse_table_node(*n, ctx, cell_lists);
                    cell_blocks.push(Block::Table(nested));
                }
            }
            if tc_pr.and_then(|pr| wml_bool(pr, "tcFitText")) == Some(true) {
                fit_cell_text(&mut cell_blocks);
            }
            cells.push(TableCell {
                width: cell_width,
                width_pct: cell_width_pct,
                content: cell_blocks,
                borders,
                shading,
                hatch,
                grid_span,
                v_merge,
                v_align,
                cell_margins: per_cell_margins,
                text_direction,
                hide_mark,
            });
            grid_col += grid_span as usize;
        }
        rows.push(TableRow {
            cells,
            grid_before,
            height: row_height,
            height_exact,
            is_header,
            cant_split,
        });
    }
    share_row_top_margins(&mut rows, cell_margins);
    settle_row_borders(&mut rows, cell_margins);
    apply_cell_spacing(&mut rows, &row_spacing, num_cols, cell_margins);
    let bidi_visual = tbl_pr
        .and_then(|p| wml_bool(p, "bidiVisual"))
        .unwrap_or(false);
    if bidi_visual {
        mirror_columns(&mut rows, &mut col_widths, &mut cell_margins);
    }

    Table {
        col_widths,
        rows,
        table_indent,
        table_indent_explicit,
        has_table_style: tbl_style.is_some() || ctx.styles.has_default_table_style,
        cell_margins,
        position: table_position,
        alignment,
        fixed_layout,
        auto_width,
        width_pct,
        grid_inferred,
        cells_unsized,
        cells_twip_sized,
        bidi_visual,
        header_first_row: tbl_look_node.is_none() || look_first_row,
        header_first_col: tbl_look_node.is_none() || look_first_col,
    }
}

/// `w:tcFitText`: each paragraph of the cell fills the cell's text width,
/// as one region apart from any fitText of its own.
fn fit_cell_text(blocks: &mut [Block]) {
    const CELL_FIT_ID: i64 = i64::MIN;
    for block in blocks {
        if let Block::Paragraph(p) = block {
            for run in p.runs.iter_mut().filter(|r| r.fit_text.is_none()) {
                run.fit_text = Some(FitText {
                    width: None,
                    id: CELL_FIT_ID,
                });
            }
        }
    }
}
