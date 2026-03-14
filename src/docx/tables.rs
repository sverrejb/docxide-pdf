use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    Alignment, CellBorder, CellBorders, CellMargins, CellVAlign, HorizontalPosition, LineSpacing,
    Paragraph, Table, TableCell, TablePosition, TableRow, TextDirection, VMerge,
};

use super::numbering::{self, ListLabelInfo, parse_list_info};
use super::runs::parse_runs;
use super::styles::{self, TableBordersDef, parse_alignment};
use super::{
    WML_NS, collect_block_nodes, extract_indents, parse_cell_border, parse_cell_border_left,
    parse_cell_border_right, parse_hex_color, parse_paragraph_spacing, twips_attr, twips_to_pts,
    wml, wml_attr,
};

fn is_wml(node: &roxmltree::Node, name: &str) -> bool {
    node.tag_name().name() == name && node.tag_name().namespace() == Some(WML_NS)
}

fn margin_twips(mar: roxmltree::Node, primary: &str, fallback: &str) -> Option<f32> {
    wml(mar, primary)
        .or_else(|| wml(mar, fallback))
        .and_then(|n| twips_attr(n, "w"))
}

fn border_or_fallback(inline: CellBorder, fallback: CellBorder) -> CellBorder {
    if inline.present { inline } else { fallback }
}

struct AnnotatedNode<'a> {
    node: roxmltree::Node<'a, 'a>,
    extra_space_before: f32,
    extra_space_after: f32,
}

/// Flatten nested table content: extract all w:p nodes from a w:tbl's cells,
/// skipping vMerge=continue cells to avoid duplicating merged content.
/// Preserves nested table cell margins as extra spacing on first/last paragraphs.
fn collect_nested_table_paragraphs<'a>(
    tbl: roxmltree::Node<'a, 'a>,
    out: &mut Vec<AnnotatedNode<'a>>,
) {
    let tbl_pr = wml(tbl, "tblPr");
    let (margin_top, margin_bottom) = tbl_pr
        .and_then(|pr| wml(pr, "tblCellMar"))
        .map(|mar| {
            let top = wml(mar, "top")
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(0.0);
            let bottom = wml(mar, "bottom")
                .and_then(|n| twips_attr(n, "w"))
                .unwrap_or(0.0);
            (top, bottom)
        })
        .unwrap_or((0.0, 0.0));

    for tr in collect_block_nodes(tbl)
        .into_iter()
        .filter(|n| is_wml(n, "tr"))
    {
        for tc in collect_block_nodes(tr)
            .into_iter()
            .filter(|n| is_wml(n, "tc"))
        {
            let tc_pr = wml(tc, "tcPr");
            let is_continue = tc_pr
                .and_then(|pr| wml(pr, "vMerge"))
                .is_some_and(|n| n.attribute((WML_NS, "val")) != Some("restart"));
            if is_continue {
                continue;
            }
            let (cell_margin_top, cell_margin_bottom) = tc_pr
                .and_then(|pr| wml(pr, "tcMar"))
                .map(|mar| {
                    let top = wml(mar, "top")
                        .and_then(|n| twips_attr(n, "w"))
                        .unwrap_or(margin_top);
                    let bottom = wml(mar, "bottom")
                        .and_then(|n| twips_attr(n, "w"))
                        .unwrap_or(margin_bottom);
                    (top, bottom)
                })
                .unwrap_or((margin_top, margin_bottom));

            let start_idx = out.len();
            for n in collect_block_nodes(tc) {
                if is_wml(&n, "p") {
                    out.push(AnnotatedNode {
                        node: n,
                        extra_space_before: 0.0,
                        extra_space_after: 0.0,
                    });
                } else if is_wml(&n, "tbl") {
                    collect_nested_table_paragraphs(n, out);
                }
            }
            let end_idx = out.len();
            if start_idx < end_idx {
                out[start_idx].extra_space_before += cell_margin_top;
                out[end_idx - 1].extra_space_after += cell_margin_bottom;
            }
        }
    }
}

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
        .filter(|n| is_wml(n, "gridCol"))
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
        TablePosition {
            h_position,
            h_anchor,
            v_offset_pt,
            v_anchor,
        }
    });

    let tbl_style_borders = tbl_pr
        .and_then(|pr| wml_attr(pr, "tblStyle"))
        .and_then(|id| styles.table_border_styles.get(id));
    let has_tbl_style = tbl_style_borders.is_some();

    let inline_tbl_borders =
        tbl_pr
            .and_then(|pr| wml(pr, "tblBorders"))
            .map(|bdr_node| TableBordersDef {
                top: parse_cell_border(bdr_node, "top"),
                bottom: parse_cell_border(bdr_node, "bottom"),
                left: parse_cell_border_left(bdr_node),
                right: parse_cell_border_right(bdr_node),
                inside_h: parse_cell_border(bdr_node, "insideH"),
                inside_v: parse_cell_border(bdr_node, "insideV"),
            });

    let effective_tbl_borders: Option<&TableBordersDef> =
        inline_tbl_borders.as_ref().or(tbl_style_borders);

    let tbl_rows: Vec<_> = collect_block_nodes(node)
        .into_iter()
        .filter(|n| is_wml(n, "tr"))
        .collect();
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
                    CellBorders {
                        top: border_or_fallback(parse_cell_border(bdr, "top"), fallback.top),
                        bottom: border_or_fallback(
                            parse_cell_border(bdr, "bottom"),
                            fallback.bottom,
                        ),
                        left: border_or_fallback(parse_cell_border_left(bdr), fallback.left),
                        right: border_or_fallback(parse_cell_border_right(bdr), fallback.right),
                    }
                })
                .unwrap_or_else(|| style_borders.unwrap_or_default());

            let shading = tc_pr
                .and_then(|pr| wml(pr, "shd"))
                .and_then(|shd| shd.attribute((WML_NS, "fill")))
                .filter(|f| *f != "none")
                .and_then(parse_hex_color);

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

            let mut cell_paras = Vec::new();
            let block_nodes = collect_block_nodes(tc);
            let mut p_nodes: Vec<AnnotatedNode> = Vec::new();
            for n in &block_nodes {
                if is_wml(n, "p") {
                    p_nodes.push(AnnotatedNode {
                        node: *n,
                        extra_space_before: 0.0,
                        extra_space_after: 0.0,
                    });
                } else if is_wml(n, "tbl") {
                    collect_nested_table_paragraphs(*n, &mut p_nodes);
                }
            }
            for ap in p_nodes {
                let p = ap.node;
                let parsed = parse_runs(p, styles, theme, rels, zip, numbering);
                let mut runs = parsed.runs;
                let has_text = runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
                let has_inline_images = runs.iter().any(|r| r.inline_image.is_some());
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
                    .unwrap_or(&styles.default_paragraph_style_id);
                let para_style = styles.paragraph_styles.get(para_style_id);
                let alignment = ppr
                    .and_then(|ppr| wml_attr(ppr, "jc"))
                    .map(parse_alignment)
                    .or_else(|| para_style.and_then(|s| s.alignment))
                    .unwrap_or(Alignment::Left);
                let (sp_before, sp_after, ls) = parse_paragraph_spacing(ppr, para_style);
                let line_spacing = ls.or_else(|| has_tbl_style.then_some(LineSpacing::Auto(1.0)));
                let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                let style_num = para_style.and_then(|s| s.num_id.as_deref());
                let style_ilvl = para_style.and_then(|s| s.num_ilvl);
                let ListLabelInfo {
                    mut indent_left,
                    mut indent_hanging,
                    label: list_label,
                    font: list_label_font,
                    font_size: list_label_font_size,
                    bold: list_label_bold,
                    color: list_label_color,
                } = parse_list_info(
                    num_pr,
                    style_num,
                    style_ilvl,
                    numbering,
                    counters,
                    last_seen_level,
                );
                let mut indent_first_line = 0.0f32;
                let mut indent_right = 0.0f32;
                if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                    let (left, right, hanging, first) = extract_indents(ind);
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
                let space_before = sp_before.unwrap_or(0.0) + ap.extra_space_before;
                let space_after = sp_after.unwrap_or(if has_tbl_style {
                    0.0
                } else {
                    styles.defaults.space_after
                }) + ap.extra_space_after;
                cell_paras.push(Paragraph {
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
                cell_margins: per_cell_margins,
                text_direction,
            });
            grid_col += grid_span as usize;
        }
        rows.push(TableRow {
            cells,
            height: row_height,
            height_exact,
            is_header,
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
