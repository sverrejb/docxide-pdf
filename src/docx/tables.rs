use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use crate::model::{
    Alignment, Block, BorderStyle, CellBorder, CellBorders, CellMargins, CellVAlign, HatchPattern,
    HorizontalPosition, LineSpacing, Paragraph, Table, TableAlignment, TableCell, TablePosition,
    TableRow, TextDirection, VMerge,
};

use super::numbering::{ListLabelInfo, parse_list_info};
use super::runs::parse_runs;
use super::styles::{TableBordersDef, parse_alignment};
use super::{
    ParseContext, WML_NS, collect_block_nodes, extract_indents, parse_cell_border,
    parse_cell_border_left, parse_cell_border_right, parse_hex_color, parse_on_off,
    parse_paragraph_spacing, twips_attr, twips_to_pts, wml, wml_attr, wml_bool,
};

/// Approximate a `w:shd` pattern fill (stripes/cross/pctNN) as a solid color,
/// since the renderer only does solid cell fills. The pattern ink (`w:color`)
/// defaults to black and the background to white when `auto`; `pctNN` blends by
/// its percentage, named patterns by an estimated coverage. Returns `None` for
/// `clear`/`solid`/`nil` (no pattern tint to approximate).
fn approx_pattern_shade(val: &str, color: Option<[u8; 3]>) -> Option<[u8; 3]> {
    let coverage = if let Some(rest) = val.strip_prefix("pct") {
        (rest.parse::<f32>().ok()? / 100.0).clamp(0.0, 1.0)
    } else {
        match val {
            "thinHorzStripe" | "thinVertStripe" | "thinDiagStripe"
            | "thinReverseDiagStripe" => 0.30,
            "horzStripe" | "vertStripe" | "diagStripe" | "reverseDiagStripe" => 0.45,
            "thinHorzCross" | "thinDiagCross" => 0.40,
            "horzCross" | "diagCross" => 0.55,
            _ => return None,
        }
    };
    let fg = color.unwrap_or([0, 0, 0]);
    let blend = |bg: u8, ink: u8| {
        (bg as f32 * (1.0 - coverage) + ink as f32 * coverage).round() as u8
    };
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

fn is_wml(node: &roxmltree::Node, name: &str) -> bool {
    node.tag_name().name() == name && node.tag_name().namespace() == Some(WML_NS)
}

fn margin_twips(mar: roxmltree::Node, primary: &str, fallback: &str) -> Option<f32> {
    wml(mar, primary)
        .or_else(|| wml(mar, fallback))
        .and_then(|n| twips_attr(n, "w"))
}

fn parse_table_borders_def(bdr_node: roxmltree::Node) -> TableBordersDef {
    TableBordersDef {
        top: parse_cell_border(bdr_node, "top"),
        bottom: parse_cell_border(bdr_node, "bottom"),
        left: parse_cell_border_left(bdr_node),
        right: parse_cell_border_right(bdr_node),
        inside_h: parse_cell_border(bdr_node, "insideH"),
        inside_v: parse_cell_border(bdr_node, "insideV"),
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
    counters: &mut HashMap<(u32, u8), u32>,
    last_seen_level: &mut HashMap<u32, u8>,
    applied_overrides: &mut HashSet<(u32, u8)>,
) -> Table {
    let mut col_widths: Vec<f32> = wml(node, "tblGrid")
        .into_iter()
        .flat_map(|grid| grid.children())
        .filter(|n| is_wml(n, "gridCol"))
        .filter_map(|n| twips_attr(n, "w"))
        .collect();

    let tbl_pr = wml(node, "tblPr");
    let table_indent_opt = tbl_pr
        .and_then(|pr| wml(pr, "tblInd"))
        .and_then(|ind| twips_attr(ind, "w"));
    let table_indent = table_indent_opt.unwrap_or(0.0);
    let table_indent_explicit = table_indent_opt.is_some();

    let alignment = tbl_pr
        .and_then(|pr| wml(pr, "jc"))
        .and_then(|jc| jc.attribute((WML_NS, "val")))
        .map(|val| match val {
            "center" => TableAlignment::Center,
            "right" | "end" => TableAlignment::Right,
            _ => TableAlignment::Left,
        })
        .unwrap_or_else(|| {
            // Fall back to first row's w:trPr/w:jc if table-level jc absent
            collect_block_nodes(node)
                .into_iter()
                .find(|n| is_wml(n, "tr"))
                .and_then(|tr| wml(tr, "trPr"))
                .and_then(|pr| wml(pr, "jc"))
                .and_then(|jc| jc.attribute((WML_NS, "val")))
                .map(|val| match val {
                    "center" => TableAlignment::Center,
                    "right" | "end" => TableAlignment::Right,
                    _ => TableAlignment::Left,
                })
                .unwrap_or_default()
        });

    let fixed_layout = tbl_pr
        .and_then(|pr| wml(pr, "tblLayout"))
        .and_then(|n| n.attribute((WML_NS, "type")))
        .is_some_and(|v| v == "fixed");

    let auto_width = tbl_pr
        .and_then(|pr| wml(pr, "tblW"))
        .is_none_or(|n| {
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

    let cell_margins = tbl_pr
        .and_then(|pr| wml(pr, "tblCellMar"))
        .map(|mar| CellMargins {
            top: wml(mar, "top")
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(0.0),
            left: margin_twips(mar, "left", "start").unwrap_or(5.4),
            bottom: wml(mar, "bottom")
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(0.0),
            right: margin_twips(mar, "right", "end").unwrap_or(5.4),
        })
        .unwrap_or_default();

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
        let v_offset_pt = tblp
            .attribute((WML_NS, "tblpY"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .unwrap_or(0.0);
        let h_position = match tblp.attribute((WML_NS, "tblpXSpec")) {
            Some("center") => HorizontalPosition::AlignCenter,
            Some("right") => HorizontalPosition::AlignRight,
            Some(_) => HorizontalPosition::AlignLeft,
            None => {
                let offset = tblp
                    .attribute((WML_NS, "tblpX"))
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(twips_to_pts)
                    .unwrap_or(0.0);
                HorizontalPosition::Offset(offset)
            }
        };
        let top_from_text = tblp
            .attribute((WML_NS, "topFromText"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .unwrap_or(0.0);
        let bottom_from_text = tblp
            .attribute((WML_NS, "bottomFromText"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .unwrap_or(0.0);
        let left_from_text = tblp
            .attribute((WML_NS, "leftFromText"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .unwrap_or(0.0);
        let right_from_text = tblp
            .attribute((WML_NS, "rightFromText"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .unwrap_or(0.0);
        TablePosition {
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

    let tbl_style = tbl_pr
        .and_then(|pr| wml_attr(pr, "tblStyle"))
        .and_then(|id| ctx.styles.table_styles.get(id));
    let tbl_style_borders = tbl_style.and_then(|s| s.base_borders.as_ref());
    let has_tbl_style = tbl_style_borders.is_some();

    let inline_tbl_borders = tbl_pr
        .and_then(|pr| wml(pr, "tblBorders"))
        .map(parse_table_borders_def);

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
        .filter(|n| is_wml(n, "tr"))
        .collect();

    // OOXML §17.4.48 requires tblGrid, but some generators (e.g. SpecLink)
    // omit it. Infer the grid from the widest row's cell widths so layout
    // has real columns to work with, splitting spanned cells evenly.
    let grid_inferred = col_widths.is_empty() && !tbl_rows.is_empty();
    if col_widths.is_empty() {
        for tr in &tbl_rows {
            let mut row_widths: Vec<f32> = Vec::new();
            for tc in collect_block_nodes(*tr)
                .into_iter()
                .filter(|n| is_wml(n, "tc"))
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

    let mut rows = Vec::new();
    for (ri, tr) in tbl_rows.iter().enumerate() {
        let tr_pr = wml(*tr, "trPr");
        let (row_height, height_exact) = tr_pr
            .and_then(|pr| wml(pr, "trHeight"))
            .map(|h| {
                let val = twips_attr(h, "val");
                let exact = h.attribute((WML_NS, "hRule")) == Some("exact");
                (val, exact)
            })
            .unwrap_or((None, false));
        let is_header = tr_pr.and_then(|pr| wml(pr, "tblHeader")).is_some();
        let cant_split = tr_pr.and_then(|pr| wml_bool(pr, "cantSplit")).unwrap_or(false);

        // Per-row table property exceptions (§17.4.60): merge with base table
        // borders — specified exception borders override, unspecified inherit.
        let base_tbl_borders: Option<&TableBordersDef> =
            inline_tbl_borders.as_ref().or(tbl_style_borders);
        let merged_row_borders;
        let row_effective_tbl_borders = match wml(*tr, "tblPrEx")
            .and_then(|prex| wml(prex, "tblBorders"))
        {
            Some(bdr_node) => {
                let exc = parse_table_borders_def(bdr_node);
                merged_row_borders = if let Some(base) = base_tbl_borders {
                    TableBordersDef {
                        top: border_or_fallback(exc.top, base.top),
                        bottom: border_or_fallback(exc.bottom, base.bottom),
                        left: border_or_fallback(exc.left, base.left),
                        right: border_or_fallback(exc.right, base.right),
                        inside_h: border_or_fallback(exc.inside_h, base.inside_h),
                        inside_v: border_or_fallback(exc.inside_v, base.inside_v),
                    }
                } else {
                    exc
                };
                Some(&merged_row_borders)
            }
            None => base_tbl_borders,
        };

        let mut cells = Vec::new();
        let mut grid_col = 0usize;
        for tc in collect_block_nodes(*tr)
            .into_iter()
            .filter(|n| is_wml(n, "tc"))
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
            let hide_mark =
                tc_pr.and_then(|pr| wml_bool(pr, "hideMark")).unwrap_or(false);

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
            });

            // Apply conditional formatting overrides from tblStylePr.
            // Order per spec: wholeTable → bands → first/last row/col → corners.
            //
            // Per OOXML §17.4.23, tblStylePr borders use inside/outside semantics:
            // top/bottom/left/right are the outer edges of the conditional region,
            // insideH/insideV are borders between cells within the region.
            let mut cond_borders = style_borders.unwrap_or_default();
            let mut cond_shading: Option<[u8; 3]> = None;
            let mut cond_bold: Option<bool> = None;
            let mut cond_italic: Option<bool> = None;
            let mut cond_color: Option<[u8; 3]> = None;
            let mut cond_font_size: Option<f32> = None;
            let mut cond_font_name: Option<String> = None;
            if let Some(style_def) = tbl_style {
                let apply_cond =
                    |key: &str,
                     borders: &mut CellBorders,
                     shading: &mut Option<[u8; 3]>,
                     bold: &mut Option<bool>,
                     italic: &mut Option<bool>,
                     color: &mut Option<[u8; 3]>,
                     font_size: &mut Option<f32>,
                     font_name: &mut Option<String>,
                     top_edge: bool,
                     bottom_edge: bool,
                     left_edge: bool,
                     right_edge: bool| {
                        if let Some(cond) = style_def.conditionals.get(key) {
                            if let Some(cb) = &cond.borders {
                                let ct = if top_edge { cb.top } else { cb.inside_h };
                                let cb_b = if bottom_edge { cb.bottom } else { cb.inside_h };
                                let cl = if left_edge { cb.left } else { cb.inside_v };
                                let cr = if right_edge { cb.right } else { cb.inside_v };
                                borders.top = border_or_fallback(ct, borders.top);
                                borders.bottom = border_or_fallback(cb_b, borders.bottom);
                                borders.left = border_or_fallback(cl, borders.left);
                                borders.right = border_or_fallback(cr, borders.right);
                            }
                            if let Some(s) = cond.shading {
                                *shading = Some(s);
                            }
                            if let Some(b) = cond.bold {
                                *bold = Some(b);
                            }
                            if let Some(i) = cond.italic {
                                *italic = Some(i);
                            }
                            if let Some(c) = cond.color {
                                *color = Some(c);
                            }
                            if let Some(fs) = cond.font_size {
                                *font_size = Some(fs);
                            }
                            if let Some(ref fn_name) = cond.font_name {
                                *font_name = Some(fn_name.clone());
                            }
                        }
                    };
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
                        let key = if band_row % 2 == 0 { "band1Horz" } else { "band2Horz" };
                        // Row bands: single row, so top/bottom are always edges
                        apply_cond(key, &mut cond_borders, &mut cond_shading,
                            &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                            true, true, is_first_col, is_last_col);
                    }
                }
                // Column banding — skip cols consumed by firstCol/lastCol
                if !look_no_v_band {
                    let skip_first = look_first_col && is_first_col;
                    let skip_last = look_last_col && is_last_col;
                    if !skip_first && !skip_last {
                        let band_col = if look_first_col { ci - 1 } else { ci };
                        let key = if band_col % 2 == 0 { "band1Vert" } else { "band2Vert" };
                        // Column bands: single column, so left/right are always edges
                        apply_cond(key, &mut cond_borders, &mut cond_shading,
                            &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                            is_first_row, is_last_row, true, true);
                    }
                }
                // First/last row — row region: top/bottom are edges, left/right depend on col
                if look_first_row && is_first_row {
                    apply_cond("firstRow", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        true, true, is_first_col, is_last_col);
                }
                if look_last_row && is_last_row {
                    apply_cond("lastRow", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        true, true, is_first_col, is_last_col);
                }
                // First/last column — column region: left/right are edges, top/bottom depend on row
                if look_first_col && is_first_col {
                    apply_cond("firstCol", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        is_first_row, is_last_row, true, true);
                }
                if look_last_col && is_last_col {
                    apply_cond("lastCol", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        is_first_row, is_last_row, true, true);
                }
                // Corner cells — single cell, all edges
                if look_first_row && is_first_row && look_first_col && is_first_col {
                    apply_cond("nwCell", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        true, true, true, true);
                }
                if look_first_row && is_first_row && look_last_col && is_last_col {
                    apply_cond("neCell", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        true, true, true, true);
                }
                if look_last_row && is_last_row && look_first_col && is_first_col {
                    apply_cond("swCell", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        true, true, true, true);
                }
                if look_last_row && is_last_row && look_last_col && is_last_col {
                    apply_cond("seCell", &mut cond_borders, &mut cond_shading,
                        &mut cond_bold, &mut cond_italic, &mut cond_color, &mut cond_font_size, &mut cond_font_name,
                        true, true, true, true);
                }
            }

            // Inline cell borders override conditional/style borders
            let borders = tc_pr
                .and_then(|pr| wml(pr, "tcBorders"))
                .map(|bdr| CellBorders {
                    top: border_or_fallback(parse_cell_border(bdr, "top"), cond_borders.top),
                    bottom: border_or_fallback(parse_cell_border(bdr, "bottom"), cond_borders.bottom),
                    left: border_or_fallback(parse_cell_border_left(bdr), cond_borders.left),
                    right: border_or_fallback(parse_cell_border_right(bdr), cond_borders.right),
                })
                .unwrap_or(cond_borders);

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
                } else if let Some(c) = bg {
                    pattern_shading = Some(c);
                } else {
                    pattern_shading = approx_pattern_shade(val, ink);
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
                cond_shading
            };
            let hatch = cell_hatch;

            let per_cell_margins = tc_pr
                .and_then(|pr| wml(pr, "tcMar"))
                .map(|mar| CellMargins {
                    top: wml(mar, "top")
                        .and_then(|n| twips_attr(n, "w"))
                        .unwrap_or(cell_margins.top),
                    left: margin_twips(mar, "left", "start")
                        .unwrap_or(cell_margins.left),
                    bottom: wml(mar, "bottom")
                        .and_then(|n| twips_attr(n, "w"))
                        .unwrap_or(cell_margins.bottom),
                    right: margin_twips(mar, "right", "end")
                        .unwrap_or(cell_margins.right),
                });

            let mut cell_blocks: Vec<Block> = Vec::new();
            let block_nodes = collect_block_nodes(tc);
            for n in &block_nodes {
                if is_wml(n, "p") {
                    let p = *n;
                    let parsed = parse_runs(p, ctx);
                    let mut runs = parsed.runs;
                    // Apply table style rPr: conditional > base, only when
                    // the run inherited from doc defaults (not set explicitly).
                    let eff_tbl_font_size = cond_font_size
                        .or_else(|| tbl_style.and_then(|s| s.base_font_size));
                    let eff_tbl_font_name = cond_font_name.as_deref()
                        .or_else(|| tbl_style.and_then(|s| s.base_font_name.as_deref()));
                    let eff_tbl_bold = cond_bold
                        .or_else(|| tbl_style.and_then(|s| s.base_bold));
                    let eff_tbl_italic = cond_italic
                        .or_else(|| tbl_style.and_then(|s| s.base_italic));

                    let mut has_text = false;
                    let mut has_inline_images = false;
                    for run in &mut runs {
                        if let Some(tfs) = eff_tbl_font_size {
                            if run.font_size_from_default {
                                run.font_size = tfs;
                            }
                        }
                        if let Some(tfn) = eff_tbl_font_name {
                            if run.font_name_from_default {
                                run.font_name = tfn.to_string();
                            }
                        }
                        if eff_tbl_bold == Some(true) {
                            run.bold = true;
                        }
                        if eff_tbl_italic == Some(true) {
                            run.italic = true;
                        }
                        if let Some(cc) = cond_color {
                            if run.color.is_none() {
                                run.color = Some(cc);
                            }
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
                    } else {
                        (None, 0.0)
                    };
                    let ppr = wml(p, "pPr");
                    let para_style_id = ppr
                        .and_then(|ppr| wml_attr(ppr, "pStyle"))
                        .unwrap_or(&ctx.styles.default_paragraph_style_id);
                    let para_style = ctx.styles.paragraph_styles.get(para_style_id);
                    let alignment = ppr
                        .and_then(|ppr| wml_attr(ppr, "jc"))
                        .map(parse_alignment)
                        .or_else(|| super::runs::display_math_alignment(p))
                        .or_else(|| para_style.and_then(|s| s.alignment))
                        .unwrap_or(Alignment::Left);
                    let (sp_before, sp_after, ls) = parse_paragraph_spacing(ppr, para_style, None);
                    let line_spacing =
                        ls.or_else(|| has_tbl_style.then_some(LineSpacing::Auto(1.0)));
                    let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                    let style_num = para_style.and_then(|s| s.num_id.as_deref());
                    let style_ilvl = para_style.and_then(|s| s.num_ilvl);
                    let ListLabelInfo {
                        mut indent_left,
                        mut indent_hanging,
                        tab_stop: _,
                        label: list_label,
                        font: list_label_font,
                        font_size: list_label_font_size,
                        bold: list_label_bold,
                        color: list_label_color,
                        suff: _,
                    } = parse_list_info(
                        num_pr,
                        style_num,
                        style_ilvl,
                        Some(para_style_id),
                        ctx.numbering,
                        counters,
                        last_seen_level,
                        applied_overrides,
                    );
                    let mut indent_first_line = 0.0;
                    let mut indent_right = 0.0;
                    if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                        let cw_fs = para_style
                            .and_then(|s| s.font_size)
                            .unwrap_or(ctx.styles.defaults.font_size);
                        let (left, right, hanging, first) = extract_indents(ind, Some(cw_fs / 2.0));
                        if let Some(v) = left {
                            indent_left = v;
                        }
                        if let Some(v) = right {
                            indent_right = v;
                        }
                        if let Some(v) = hanging {
                            indent_hanging = v;
                        }
                        if let Some(v) = first {
                            indent_first_line = v;
                        }
                    }
                    let space_before = sp_before.unwrap_or(0.0);
                    let space_after = sp_after.unwrap_or(if has_tbl_style {
                        0.0
                    } else {
                        ctx.styles.defaults.space_after
                    });
                    cell_blocks.push(Block::Paragraph(Paragraph {
                        runs,
                        alignment,
                        indent_left,
                        indent_right,
                        indent_hanging,
                        indent_first_line,
                        list_label,
                        list_label_font,
                        list_label_font_size,
                        list_label_bold,
                        list_label_color,
                        line_spacing,
                        space_before,
                        space_after,
                        image: para_image,
                        content_height,
                        snap_to_grid: true,
                        floating_images: parsed.floating_images,
                        textboxes: parsed.textboxes,
                        connectors: parsed.connectors,
                        ..Paragraph::default()
                    }));
                } else if is_wml(n, "tbl") {
                    let nested = parse_table_node(
                        *n, ctx, counters, last_seen_level, applied_overrides,
                    );
                    cell_blocks.push(Block::Table(nested));
                }
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
            height: row_height,
            height_exact,
            is_header,
            cant_split,
        });
    }
    resolve_h_border_conflicts(&mut rows);
    propagate_vmerge_borders(&mut rows);

    Table {
        col_widths,
        rows,
        table_indent,
        table_indent_explicit,
        cell_margins,
        position: table_position,
        alignment,
        fixed_layout,
        auto_width,
        width_pct,
        grid_inferred,
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
        let mut ug = 0usize;
        let mut lg = 0usize;
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
                lower_row.cells[li].borders.top = winner;
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
        let mut grid_col = 0usize;
        for ci in 0..rows[ri].cells.len() {
            let span = rows[ri].cells[ci].grid_span.max(1) as usize;
            if rows[ri].cells[ci].v_merge == VMerge::Restart {
                let mut last_ri = ri;
                for next_ri in (ri + 1)..rows.len() {
                    let mut g = 0usize;
                    let mut is_continue = false;
                    for c in &rows[next_ri].cells {
                        if g == grid_col {
                            is_continue = c.v_merge == VMerge::Continue;
                            break;
                        }
                        g += c.grid_span.max(1) as usize;
                        if g > grid_col {
                            break;
                        }
                    }
                    if is_continue {
                        last_ri = next_ri;
                    } else {
                        break;
                    }
                }
                if last_ri > ri {
                    let mut g = 0usize;
                    for c in &rows[last_ri].cells {
                        if g == grid_col {
                            rows[ri].cells[ci].borders.bottom = c.borders.bottom;
                            break;
                        }
                        g += c.grid_span.max(1) as usize;
                        if g > grid_col {
                            break;
                        }
                    }
                }
            }
            grid_col += span;
        }
    }
}

#[cfg(test)]
mod border_conflict_tests {
    use super::*;

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
