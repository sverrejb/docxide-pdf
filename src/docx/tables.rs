use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    Alignment, CellBorder, CellBorders, CellMargins, CellVAlign, HorizontalPosition, LineSpacing,
    Paragraph, Table, TableCell, TablePosition, TableRow, VMerge,
};

use super::numbering::{self, parse_list_info};
use super::runs::parse_runs;
use super::styles::{self, TableBordersDef, parse_alignment, parse_line_spacing};
use super::{
    WML_NS, collect_block_nodes, parse_hex_color, twips_attr, twips_to_pts, wml, wml_attr,
};

pub(in crate::docx) fn parse_table_node<R: Read + std::io::Seek>(
    node: roxmltree::Node,
    styles: &styles::StylesInfo,
    theme: &styles::ThemeFonts,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    numbering: &numbering::NumberingInfo,
    counters: &mut HashMap<(String, u8), u32>,
    last_seen_level: &mut HashMap<String, u8>,
) -> Table {
    let col_widths: Vec<f32> = wml(node, "tblGrid")
        .into_iter()
        .flat_map(|grid| grid.children())
        .filter(|n| n.tag_name().name() == "gridCol" && n.tag_name().namespace() == Some(WML_NS))
        .filter_map(|n| twips_attr(n, "w"))
        .collect();

    let tbl_pr = wml(node, "tblPr");
    let table_indent = tbl_pr
        .and_then(|pr| wml(pr, "tblInd"))
        .and_then(|ind| twips_attr(ind, "w"))
        .unwrap_or(0.0);

    let cell_margins = tbl_pr
        .and_then(|pr| wml(pr, "tblCellMar"))
        .map(|mar| CellMargins {
            top: wml(mar, "top")
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(0.0),
            left: wml(mar, "left")
                .or_else(|| wml(mar, "start"))
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(5.4),
            bottom: wml(mar, "bottom")
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(0.0),
            right: wml(mar, "right")
                .or_else(|| wml(mar, "end"))
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(5.4),
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
        let h_position = if let Some(spec) = tblp.attribute((WML_NS, "tblpXSpec")) {
            match spec {
                "center" => HorizontalPosition::AlignCenter,
                "right" => HorizontalPosition::AlignRight,
                _ => HorizontalPosition::AlignLeft,
            }
        } else {
            let offset = tblp
                .attribute((WML_NS, "tblpX"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(0.0);
            HorizontalPosition::Offset(offset)
        };
        TablePosition {
            h_position,
            h_anchor,
            v_offset_pt,
            v_anchor,
        }
    });

    let has_tbl_style = tbl_pr.and_then(|pr| wml_attr(pr, "tblStyle")).is_some();
    let tbl_style_borders = tbl_pr
        .and_then(|pr| wml_attr(pr, "tblStyle"))
        .and_then(|id| styles.table_border_styles.get(id));

    let inline_tbl_borders = tbl_pr.and_then(|pr| wml(pr, "tblBorders")).map(|bdr_node| {
        let parse_bdr = |name: &str| -> CellBorder {
            let Some(n) = wml(bdr_node, name) else {
                return CellBorder::default();
            };
            let val = n.attribute((WML_NS, "val")).unwrap_or("none");
            if val == "nil" || val == "none" {
                return CellBorder::default();
            }
            let width = n
                .attribute((WML_NS, "sz"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|v| v / 8.0)
                .unwrap_or(0.5);
            let color = n.attribute((WML_NS, "color")).and_then(parse_hex_color);
            CellBorder::visible(color, width)
        };
        let left = parse_bdr("left");
        let left = if left.present {
            left
        } else {
            parse_bdr("start")
        };
        let right = parse_bdr("right");
        let right = if right.present {
            right
        } else {
            parse_bdr("end")
        };
        TableBordersDef {
            top: parse_bdr("top"),
            bottom: parse_bdr("bottom"),
            left,
            right,
            inside_h: parse_bdr("insideH"),
            inside_v: parse_bdr("insideV"),
        }
    });

    let effective_tbl_borders: Option<&TableBordersDef> =
        inline_tbl_borders.as_ref().or(tbl_style_borders);

    let tbl_rows: Vec<_> = collect_block_nodes(node)
        .into_iter()
        .filter(|n| n.tag_name().name() == "tr" && n.tag_name().namespace() == Some(WML_NS))
        .collect();
    let num_rows = tbl_rows.len();
    let num_cols = col_widths.len();

    let parse_cell_border = |bdr_node: roxmltree::Node, name: &str| -> CellBorder {
        let Some(n) = wml(bdr_node, name) else {
            return CellBorder::default();
        };
        let val = n.attribute((WML_NS, "val")).unwrap_or("none");
        if val == "nil" || val == "none" {
            return CellBorder::default();
        }
        let width = n
            .attribute((WML_NS, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|v| v / 8.0)
            .unwrap_or(0.5);
        let color = n.attribute((WML_NS, "color")).and_then(parse_hex_color);
        CellBorder::visible(color, width)
    };

    let mut rows = Vec::new();
    for (ri, tr) in tbl_rows.iter().enumerate() {
        let tr_pr = wml(*tr, "trPr");
        let (row_height, height_exact) = tr_pr
            .and_then(|pr| wml(pr, "trHeight"))
            .map(|h| {
                let val = h
                    .attribute((WML_NS, "val"))
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(twips_to_pts);
                let exact = h.attribute((WML_NS, "hRule")) == Some("exact");
                (val, exact)
            })
            .unwrap_or((None, false));

        let mut cells = Vec::new();
        let mut grid_col = 0usize;
        for tc in collect_block_nodes(*tr)
            .into_iter()
            .filter(|n| n.tag_name().name() == "tc" && n.tag_name().namespace() == Some(WML_NS))
        {
            let ci = grid_col;
            let tc_pr = wml(tc, "tcPr");
            let cell_width = tc_pr
                .and_then(|pr| wml(pr, "tcW"))
                .and_then(|w| twips_attr(w, "w"))
                .unwrap_or_else(|| col_widths.get(ci).copied().unwrap_or(72.0));

            let grid_span = tc_pr
                .and_then(|pr| wml(pr, "gridSpan"))
                .and_then(|n| n.attribute((WML_NS, "val")))
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
                .and_then(|pr| wml(pr, "vAlign"))
                .and_then(|n| n.attribute((WML_NS, "val")))
                .map(|v| match v {
                    "center" => CellVAlign::Center,
                    "bottom" => CellVAlign::Bottom,
                    _ => CellVAlign::Top,
                })
                .unwrap_or(CellVAlign::Top);

            let span_end = ci + grid_span as usize;

            let style_borders = effective_tbl_borders.map(|tb| CellBorders {
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

            let borders = tc_pr
                .and_then(|pr| wml(pr, "tcBorders"))
                .map(|bdr| {
                    let fallback = style_borders.unwrap_or_default();
                    let top = parse_cell_border(bdr, "top");
                    let bottom = parse_cell_border(bdr, "bottom");
                    let left = parse_cell_border(bdr, "left");
                    let left = if left.present {
                        left
                    } else {
                        parse_cell_border(bdr, "start")
                    };
                    let right = parse_cell_border(bdr, "right");
                    let right = if right.present {
                        right
                    } else {
                        parse_cell_border(bdr, "end")
                    };
                    CellBorders {
                        top: if top.present { top } else { fallback.top },
                        bottom: if bottom.present {
                            bottom
                        } else {
                            fallback.bottom
                        },
                        left: if left.present { left } else { fallback.left },
                        right: if right.present { right } else { fallback.right },
                    }
                })
                .unwrap_or_else(|| style_borders.unwrap_or_default());

            let shading = tc_pr
                .and_then(|pr| wml(pr, "shd"))
                .and_then(|shd| shd.attribute((WML_NS, "fill")))
                .filter(|f| *f != "none")
                .and_then(parse_hex_color);

            let mut cell_paras = Vec::new();
            for p in tc
                .children()
                .filter(|n| n.tag_name().name() == "p" && n.tag_name().namespace() == Some(WML_NS))
            {
                let parsed = parse_runs(p, styles, theme, rels, zip, numbering);
                let ppr = wml(p, "pPr");
                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"))
                    .unwrap_or("Normal");
                let para_style = styles.paragraph_styles.get(para_style_id);
                let alignment = ppr
                    .and_then(|ppr| wml_attr(ppr, "jc"))
                    .map(parse_alignment)
                    .or_else(|| para_style.and_then(|s| s.alignment))
                    .unwrap_or(Alignment::Left);
                let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
                let line_spacing = inline_spacing
                    .and_then(|n| {
                        n.attribute((WML_NS, "line"))
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|line_val| parse_line_spacing(n, line_val))
                    })
                    .or_else(|| para_style.and_then(|s| s.line_spacing))
                    .or_else(|| {
                        if has_tbl_style {
                            Some(LineSpacing::Auto(1.0))
                        } else {
                            None
                        }
                    });
                let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                let (mut indent_left, mut indent_hanging, list_label, list_label_font) =
                    parse_list_info(num_pr, numbering, counters, last_seen_level);
                let mut indent_first_line = 0.0f32;
                let mut indent_right = 0.0f32;
                if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                    if let Some(v) = twips_attr(ind, "start").or_else(|| twips_attr(ind, "left")) {
                        indent_left = v;
                    }
                    if let Some(v) = twips_attr(ind, "end").or_else(|| twips_attr(ind, "right")) {
                        indent_right = v;
                    }
                    if let Some(v) = twips_attr(ind, "hanging") {
                        indent_hanging = v;
                    }
                    if let Some(v) = twips_attr(ind, "firstLine") {
                        indent_first_line = v;
                    }
                }
                let space_before = inline_spacing
                    .and_then(|n| twips_attr(n, "before"))
                    .or_else(|| para_style.and_then(|s| s.space_before))
                    .unwrap_or(0.0);
                let space_after = inline_spacing
                    .and_then(|n| twips_attr(n, "after"))
                    .or_else(|| para_style.and_then(|s| s.space_after))
                    .unwrap_or_else(|| {
                        if has_tbl_style {
                            0.0
                        } else {
                            styles.defaults.space_after
                        }
                    });
                cell_paras.push(Paragraph {
                    runs: parsed.runs,
                    alignment,
                    indent_left,
                    indent_right,
                    indent_hanging,
                    indent_first_line,
                    list_label,
                    list_label_font,
                    line_spacing,
                    space_before,
                    space_after,
                    extra_line_breaks: parsed.line_break_count,
                    ..Paragraph::default()
                });
            }
            cells.push(TableCell {
                width: cell_width,
                paragraphs: cell_paras,
                borders,
                shading,
                grid_span,
                v_merge,
                v_align,
            });
            grid_col += grid_span as usize;
        }
        rows.push(TableRow {
            cells,
            height: row_height,
            height_exact,
        });
    }
    Table {
        col_widths,
        rows,
        table_indent,
        cell_margins,
        position: table_position,
    }
}
