use std::io::{Read, Seek};

use crate::model::{
    Alignment, Block, BorderStyle, CellBorder, CellBorders, CellMargins, CellVAlign, FitText,
    HatchPattern, HorizontalPosition, LineSpacing, Paragraph, Table, TableAlignment, TableCell,
    TablePosition, TableRow, TextDirection, VMerge,
};

use super::numbering::{ListCounters, ListLabelInfo, parse_list_info};
use super::runs::parse_runs;
use super::styles::{TableBordersDef, TableStyleDef, parse_table_borders_def};
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

/// `w:jc` under a `w:tblPr` or `w:trPr`.
fn table_jc(pr: roxmltree::Node) -> Option<TableAlignment> {
    wml_attr(pr, "jc").map(|val| match val {
        "center" => TableAlignment::Center,
        "right" | "end" => TableAlignment::Right,
        _ => TableAlignment::Left,
    })
}

/// Per-side merge: sides `over` specifies (or explicitly clears) win, the rest
/// fall through to `base`.
fn merge_table_borders(over: TableBordersDef, base: TableBordersDef) -> TableBordersDef {
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
        if let Some(s) = cond.shading {
            self.shading = Some(s);
        }
        if let Some(b) = cond.bold {
            self.bold = Some(b);
        }
        if let Some(i) = cond.italic {
            self.italic = Some(i);
        }
        if let Some(c) = cond.color {
            self.color = Some(c);
        }
        if let Some(fs) = cond.font_size {
            self.font_size = Some(fs);
        }
        if let Some(ref fn_name) = cond.font_name {
            self.font_name = Some(fn_name.clone());
        }
    }
}

/// Border style precedence for the §17.4.66 conflict tiebreaker: when two
/// borders have equal weight (same width, same source level), the more
/// prominent style wins. Lower number = wins. Double is ranked above Single
/// because our `width` field carries only the nominal stroke width and does
/// not encode the double rule's extra visual weight; everything beats the
/// dotted/dashed family, which is what makes Word collapse a `single`+`dotted`
/// edge to a solid line.
fn border_style_precedence(style: BorderStyle) -> u8 {
    match style {
        BorderStyle::Double => 0,
        BorderStyle::Single => 1,
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

    let alignment = tbl_pr.and_then(table_jc).unwrap_or_else(|| {
        // Fall back to first row's w:trPr/w:jc if table-level jc absent
        collect_block_nodes(node)
            .into_iter()
            .find(|n| n.has_tag_name((WML_NS, "tr")))
            .and_then(|tr| wml(tr, "trPr"))
            .and_then(table_jc)
            .unwrap_or_default()
    });

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
    // Each side: the table's own, then its style's (estonian_community's
    // TableGrid zeroes them), then Word's default.
    let own_mar = tbl_pr.and_then(|pr| wml(pr, "tblCellMar"));
    let has_style_or_cell_mar =
        tbl_style.is_some() || ctx.styles.has_default_table_style || own_mar.is_some();
    // Word's 5.4pt default comes with a table style: a table with neither a
    // style nor its own tblCellMar gets 0.5pt side margins and so no visible
    // outdent (Word probes on case200/case213 with cell borders: the border
    // at the margin, the text 0.5pt inside it and a right-aligned line
    // ending ~0.5pt short of the right border; case131's nested table text
    // 1.0pt inside the parent's, its second column 0.58pt inside; an empty
    // default table style brings the 5.4pt back).
    let default_side = if has_style_or_cell_mar { 5.4 } else { 0.5 };
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

    let tbl_rows: Vec<_> = collect_block_nodes(node)
        .into_iter()
        .filter(|n| n.has_tag_name((WML_NS, "tr")))
        .collect();

    // OOXML §17.4.48 requires tblGrid, but some generators (e.g. SpecLink)
    // omit it. Infer the grid from the widest row's cell widths so layout
    // has real columns to work with, splitting spanned cells evenly.
    let grid_inferred = col_widths.is_empty() && !tbl_rows.is_empty();
    let cells_unsized = tbl_rows.iter().all(|tr| {
        collect_block_nodes(*tr)
            .into_iter()
            .filter(|n| n.has_tag_name((WML_NS, "tc")))
            .all(|tc| {
                wml(tc, "tcPr")
                    .and_then(|pr| wml(pr, "tcW"))
                    .is_none_or(|w| {
                        matches!(w.attribute((WML_NS, "type")), Some("auto" | "nil"))
                            || twips_attr(w, "w").is_none_or(|v| v == 0.0)
                    })
            })
    });
    if col_widths.is_empty() {
        for tr in &tbl_rows {
            let mut row_widths: Vec<f32> = Vec::new();
            for tc in collect_block_nodes(*tr)
                .into_iter()
                .filter(|n| n.has_tag_name((WML_NS, "tc")))
            {
                let tc_pr = wml(tc, "tcPr");
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
        for tc in collect_block_nodes(*tr)
            .into_iter()
            .filter(|n| n.has_tag_name((WML_NS, "tc")))
        {
            let ci = grid_col;
            let tc_pr = wml(tc, "tcPr");
            let cell_width = tc_pr
                .and_then(|pr| wml(pr, "tcW"))
                .and_then(|w| twips_attr(w, "w"))
                .unwrap_or_else(|| col_widths.get(ci).copied().unwrap_or(72.0));

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

            let v_align = match tc_pr.and_then(|pr| wml_attr(pr, "vAlign")) {
                Some("center") => CellVAlign::Center,
                Some("bottom") => CellVAlign::Bottom,
                _ => CellVAlign::Top,
            };

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
                // Row banding — skip rows consumed by firstRow/lastRow
                if !look_no_h_band {
                    let skip_first = look_first_row && is_first_row;
                    let skip_last = look_last_row && is_last_row;
                    if !skip_first && !skip_last {
                        let band_row = if look_first_row { ri - 1 } else { ri };
                        let key = if band_row % 2 == 0 {
                            "band1Horz"
                        } else {
                            "band2Horz"
                        };
                        // Row bands: single row, so top/bottom are always edges
                        cond.apply(style_def, key, (true, true, is_first_col, is_last_col));
                    }
                }
                // Column banding — skip cols consumed by firstCol/lastCol
                if !look_no_v_band {
                    let skip_first = look_first_col && is_first_col;
                    let skip_last = look_last_col && is_last_col;
                    if !skip_first && !skip_last {
                        let band_col = if look_first_col { ci - 1 } else { ci };
                        let key = if band_col % 2 == 0 {
                            "band1Vert"
                        } else {
                            "band2Vert"
                        };
                        // Column bands: single column, so left/right are always edges
                        cond.apply(style_def, key, (is_first_row, is_last_row, true, true));
                    }
                }
                // First/last row — row region: top/bottom are edges, left/right depend on col
                if look_first_row && is_first_row {
                    cond.apply(
                        style_def,
                        "firstRow",
                        (true, true, is_first_col, is_last_col),
                    );
                }
                if look_last_row && is_last_row {
                    cond.apply(
                        style_def,
                        "lastRow",
                        (true, true, is_first_col, is_last_col),
                    );
                }
                // First/last column — column region: left/right are edges, top/bottom depend on row
                if look_first_col && is_first_col {
                    cond.apply(
                        style_def,
                        "firstCol",
                        (is_first_row, is_last_row, true, true),
                    );
                }
                if look_last_col && is_last_col {
                    cond.apply(
                        style_def,
                        "lastCol",
                        (is_first_row, is_last_row, true, true),
                    );
                }
                // Corner cells — single cell, all edges
                if look_first_row && is_first_row && look_first_col && is_first_col {
                    cond.apply(style_def, "nwCell", (true, true, true, true));
                }
                if look_first_row && is_first_row && look_last_col && is_last_col {
                    cond.apply(style_def, "neCell", (true, true, true, true));
                }
                if look_last_row && is_last_row && look_first_col && is_first_col {
                    cond.apply(style_def, "swCell", (true, true, true, true));
                }
                if look_last_row && is_last_row && look_last_col && is_last_col {
                    cond.apply(style_def, "seCell", (true, true, true, true));
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
            let block_nodes = collect_block_nodes(tc);
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
                    let (para_image, content_height) = if has_inline_images && !has_text {
                        let idx = runs.iter().position(|r| r.inline_image.is_some());
                        let img = idx.and_then(|i| runs[i].inline_image.take());
                        let h = img
                            .as_ref()
                            .map(|i| i.display_height + i.layout_extra_height)
                            .unwrap_or(0.0);
                        (img, h)
                    } else if !has_text {
                        (None, super::images::inline_drawing_height(p))
                    } else {
                        (None, 0.0)
                    };
                    let ppr = wml(p, "pPr");
                    let para_style_id = ppr
                        .and_then(|ppr| wml_attr(ppr, "pStyle"))
                        .unwrap_or(&ctx.styles.default_paragraph_style_id);
                    let para_style = ctx.styles.paragraph_styles.get(para_style_id);
                    let alignment = super::runs::paragraph_alignment(p, ppr, para_style)
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
                    let ListLabelInfo {
                        indent_left: _,
                        indent_hanging: _,
                        tab_stop: mut num_tab_stop,
                        label: list_label,
                        font: list_label_font,
                        east_asia_font: list_label_east_asia_font,
                        font_size: list_label_font_size,
                        rise: list_label_rise,
                        bold: list_label_bold,
                        italic: list_label_italic,
                        color: list_label_color,
                        suff: list_label_suff,
                        jc: list_label_jc,
                        item: list_item,
                    } = numbering;
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
                    if let Some(tabs) = ppr.and_then(|ppr| wml(ppr, "tabs")) {
                        for t in tabs.children().filter(|n| n.has_tag_name((WML_NS, "tab"))) {
                            if t.attribute((WML_NS, "val")) == Some("num")
                                && let Some(pos) = super::twips_attr(t, "pos")
                                && pos > 0.0
                            {
                                num_tab_stop = Some(pos);
                            }
                        }
                    }
                    let tab_stops = super::resolve_tab_stops(ppr, para_style);
                    cell_blocks.push(Block::Paragraph(Paragraph {
                        widow_control: super::paragraph::widow_control(ppr, para_style, &ctx.styles.defaults),
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
                        list_label,
                        num_level_tab_stop: num_tab_stop,
                        no_indent_num_tab: ctx.styles.defaults.no_indent_as_num_tab,
                        horizontal_rule: parsed.horizontal_rule,
                        field_code_mark: parsed.mark_in_field_code,
                        list_label_font,
                        list_label_east_asia_font,
                        list_label_font_size,
                        list_label_rise,
                        list_label_bold,
                        list_label_italic,
                        list_label_color,
                        list_label_jc,
                        list_label_suff: super::numbering::label_suffix(&list_label_suff),
                        list_item,
                        line_spacing,
                        space_before,
                        space_after,
                        image: para_image,
                        content_height,
                        snap_to_grid: ppr
                            .and_then(|ppr| wml_bool(ppr, "snapToGrid"))
                            .or_else(|| para_style.and_then(|s| s.snap_to_grid))
                            .unwrap_or(true),
                        text_alignment: super::paragraph::text_alignment(ppr, para_style, &ctx.styles.defaults),
                        // Cells space Latin text from ideographs as the body
                        // does (chinese_costume's "5-10 分钟").
                        auto_space_de,
                        auto_space_dn,
                        floating_images: parsed.floating_images,
                        textboxes: parsed.textboxes,
                        connectors: parsed.connectors,
                        style_id: Some(para_style_id.to_string()),
                        contextual_spacing,
                        keep_next: ppr
                            .and_then(|ppr| wml_bool(ppr, "keepNext"))
                            .or_else(|| para_style.and_then(|s| s.keep_next))
                            .unwrap_or(false),
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
                        ..Paragraph::default()
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
        has_style_or_cell_mar,
        cell_margins,
        position: table_position,
        alignment,
        fixed_layout,
        auto_width,
        width_pct,
        grid_inferred,
        cells_unsized,
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

/// A right-to-left table laid out left to right: each row's cells and the grid
/// reversed, with left and right borders and margins swapped.
/// ponytail: gridBefore keeps its side, and tagging reads the mirrored order
/// (the first-column TH lands on the logically last column).
fn mirror_columns(rows: &mut [TableRow], col_widths: &mut [f32], margins: &mut CellMargins) {
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
fn share_row_top_margins(rows: &mut [TableRow], table_margins: CellMargins) {
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
    for cell in rows.iter_mut().flat_map(|r| r.cells.iter_mut()) {
        let (top, bottom) = (cell.borders.top.band(), cell.borders.bottom.band());
        if top > 0.0 || bottom > 0.0 {
            let mut m = cell.cell_margins.unwrap_or(cell_margins);
            m.top += top / 2.0;
            m.bottom += bottom / 2.0;
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
fn apply_cell_spacing(rows: &mut [TableRow], spacing: &[f32], num_cols: usize, base: CellMargins) {
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
            if let Some(bottom) = last_ri
                .and_then(|n| cell_at(&rows[n], grid_col))
                .map(|c| c.borders.bottom)
            {
                rows[ri].cells[ci].borders.bottom = bottom;
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

    #[test]
    fn row_margin_exceptions_inherit_per_side_and_do_not_leak() {
        use std::io::{Cursor, Write};
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
          <w:tblPr><w:tblCellMar><w:top w:w="20"/><w:left w:w="100"/><w:bottom w:w="40"/><w:right w:w="120"/></w:tblCellMar></w:tblPr>
          <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
          <w:tr><w:tblPrEx><w:tblCellMar><w:start w:w="28"/><w:end w:w="32"/></w:tblCellMar></w:tblPrEx>
            <w:tc><w:p/></w:tc><w:tc><w:tcPr><w:tcMar><w:left w:w="0"/><w:bottom w:w="60"/></w:tcMar></w:tcPr><w:p/></w:tc>
          </w:tr><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
        </w:tbl><w:sectPr/></w:body></w:document>"#;
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file(
            "word/document.xml",
            zip::write::SimpleFileOptions::default(),
        )
        .unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        let bytes = zip.finish().unwrap().into_inner();
        let doc = crate::docx::parse_bytes(&bytes).unwrap();
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
        use std::io::{Cursor, Write};
        // 20 twips = 1pt spacing on the table, 0 on the second row.
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
          <w:tblPr><w:tblCellSpacing w:w="20" w:type="dxa"/><w:tblCellMar><w:left w:w="0"/><w:right w:w="0"/></w:tblCellMar></w:tblPr>
          <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
          <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
          <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
          <w:tr><w:trPr><w:tblCellSpacing w:w="0" w:type="dxa"/></w:trPr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
        </w:tbl><w:sectPr/></w:body></w:document>"#;
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file(
            "word/document.xml",
            zip::write::SimpleFileOptions::default(),
        )
        .unwrap();
        zip.write_all(xml.as_bytes()).unwrap();
        let bytes = zip.finish().unwrap().into_inner();
        let doc = crate::docx::parse_bytes(&bytes).unwrap();
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
    fn default_side_margins_need_a_style_or_tbl_cell_mar() {
        use std::io::{Cursor, Write};
        let left = |styles: Option<&str>, tbl_pr: &str| {
            let xml = format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr>{tbl_pr}</w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcMar><w:right w:w="108"/></w:tcMar></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#
            );
            let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
            let opts = zip::write::SimpleFileOptions::default();
            zip.start_file("word/document.xml", opts).unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
            if let Some(s) = styles {
                zip.start_file("word/styles.xml", opts).unwrap();
                zip.write_all(format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{s}</w:styles>"#).as_bytes()).unwrap();
            }
            let bytes = zip.finish().unwrap().into_inner();
            let doc = crate::docx::parse_bytes(&bytes).unwrap();
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
            (5.4, 5.4)
        );
        assert_eq!(left(Some(default_style), ""), (5.4, 5.4));
        assert_eq!(
            left(None, r#"<w:tblCellMar><w:left w:w="108"/></w:tblCellMar>"#),
            (5.4, 5.4)
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
