mod embedded_fonts;
mod headers_footers;
mod images;
mod numbering;
mod runs;
mod sections;
mod styles;
mod textbox;

use std::collections::HashMap;
use std::io::Read;

use crate::error::Error;
use crate::model::{
    Alignment, Block, CellBorder, CellBorders, CellMargins, CellVAlign, Document,
    HorizontalPosition, LineSpacing, Paragraph, ParagraphBorders, Section, SectionBreakType,
    SectionProperties, TabAlignment, TabStop, Table, TableCell, TablePosition, TableRow, VMerge,
};

use styles::{parse_alignment, parse_line_spacing, parse_styles, parse_theme};

use embedded_fonts::parse_font_table;
use headers_footers::parse_footnotes;
use images::compute_drawing_info;
use numbering::{parse_list_info, parse_numbering};
use runs::parse_runs;
use sections::parse_section_properties;
use textbox::collect_textboxes_from_paragraph;

pub(super) const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPD_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS_TOP: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(super) fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

pub(super) fn parse_hex_color(val: &str) -> Option<[u8; 3]> {
    if val == "auto" || val.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&val[0..2], 16).ok()?;
    let g = u8::from_str_radix(&val[2..4], 16).ok()?;
    let b = u8::from_str_radix(&val[4..6], 16).ok()?;
    Some([r, g, b])
}

pub(super) fn parse_text_color(val: &str) -> Option<[u8; 3]> {
    if val == "auto" {
        return Some([0, 0, 0]);
    }
    parse_hex_color(val)
}

fn highlight_color(name: &str) -> Option<[u8; 3]> {
    match name {
        "yellow" => Some([255, 255, 0]),
        "green" => Some([0, 255, 0]),
        "cyan" => Some([0, 255, 255]),
        "magenta" => Some([255, 0, 255]),
        "red" => Some([255, 0, 0]),
        "blue" => Some([0, 0, 255]),
        "darkYellow" => Some([128, 128, 0]),
        "darkGreen" => Some([0, 128, 0]),
        "darkCyan" => Some([0, 128, 128]),
        "darkMagenta" => Some([128, 0, 128]),
        "darkRed" => Some([128, 0, 0]),
        "darkBlue" => Some([0, 0, 128]),
        "lightGray" => Some([192, 192, 192]),
        "darkGray" => Some([128, 128, 128]),
        "black" => Some([0, 0, 0]),
        "white" => Some([255, 255, 255]),
        _ => None,
    }
}

pub(super) fn wml_bool(parent: roxmltree::Node, name: &str) -> Option<bool> {
    wml(parent, name).map(|n| {
        n.attribute((WML_NS, "val"))
            .is_none_or(|v| v != "0" && v != "false")
    })
}

pub(super) fn wml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(WML_NS))
}

pub(super) fn wml_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    wml(node, child).and_then(|n| n.attribute((WML_NS, "val")))
}

pub(super) fn twips_attr(node: roxmltree::Node, attr: &str) -> Option<f32> {
    node.attribute((WML_NS, attr))
        .and_then(|v| v.parse::<f32>().ok())
        .map(twips_to_pts)
}

fn parse_one_border(node: roxmltree::Node) -> Option<crate::model::ParagraphBorder> {
    let val = node.attribute((WML_NS, "val")).unwrap_or("none");
    if val == "none" || val == "nil" {
        return None;
    }
    let width_pt = node
        .attribute((WML_NS, "sz"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 8.0)
        .unwrap_or(0.5);
    let space_pt = node
        .attribute((WML_NS, "space"))
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0);
    let color = node
        .attribute((WML_NS, "color"))
        .and_then(parse_hex_color)
        .unwrap_or([0, 0, 0]);
    Some(crate::model::ParagraphBorder {
        width_pt,
        space_pt,
        color,
    })
}

pub(super) fn parse_paragraph_borders(ppr: roxmltree::Node) -> ParagraphBorders {
    let Some(pbdr) = wml(ppr, "pBdr") else {
        return ParagraphBorders::default();
    };
    ParagraphBorders {
        top: wml(pbdr, "top").and_then(parse_one_border),
        bottom: wml(pbdr, "bottom").and_then(parse_one_border),
        left: wml(pbdr, "left").and_then(parse_one_border),
        right: wml(pbdr, "right").and_then(parse_one_border),
        between: wml(pbdr, "between").and_then(parse_one_border),
    }
}

fn parse_tab_stops(ppr: roxmltree::Node) -> Vec<TabStop> {
    let Some(tabs) = wml(ppr, "tabs") else {
        return vec![];
    };
    let mut stops: Vec<TabStop> = tabs
        .children()
        .filter(|n| n.tag_name().name() == "tab" && n.tag_name().namespace() == Some(WML_NS))
        .filter_map(|n| {
            let pos = twips_attr(n, "pos")?;
            let val = n.attribute((WML_NS, "val")).unwrap_or("left");
            if val == "clear" {
                return None;
            }
            let alignment = match val {
                "center" => TabAlignment::Center,
                "right" => TabAlignment::Right,
                "decimal" => TabAlignment::Decimal,
                _ => TabAlignment::Left,
            };
            let leader = n.attribute((WML_NS, "leader")).and_then(|l| match l {
                "dot" => Some('.'),
                "hyphen" => Some('-'),
                "underscore" => Some('_'),
                _ => None,
            });
            Some(TabStop {
                position: pos,
                alignment,
                leader,
            })
        })
        .collect();
    stops.sort_by(|a, b| a.position.partial_cmp(&b.position).unwrap());
    stops
}

fn collect_block_nodes<'a>(parent: roxmltree::Node<'a, 'a>) -> Vec<roxmltree::Node<'a, 'a>> {
    let mut nodes = Vec::new();
    for child in parent.children() {
        if child.tag_name().name() == "sdt" && child.tag_name().namespace() == Some(WML_NS) {
            if let Some(content) = wml(child, "sdtContent") {
                nodes.extend(collect_block_nodes(content));
            }
        } else {
            nodes.push(child);
        }
    }
    nodes
}

pub(super) fn read_zip_text<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    name: &str,
) -> Option<String> {
    let mut content = String::new();
    zip.by_name(name).ok()?.read_to_string(&mut content).ok()?;
    Some(content)
}

// --- Relationship parsing (too small for own module) ---

mod relationships {
    use std::collections::HashMap;
    use std::io::Read;

    use super::read_zip_text;

    fn parse_rels_xml(xml_content: &str) -> HashMap<String, String> {
        let mut rels = HashMap::new();
        let Ok(xml) = roxmltree::Document::parse(xml_content) else {
            return rels;
        };
        for node in xml.root_element().children() {
            if node.tag_name().name() == "Relationship"
                && let (Some(id), Some(target)) =
                    (node.attribute("Id"), node.attribute("Target"))
            {
                rels.insert(id.to_string(), target.to_string());
            }
        }
        rels
    }

    pub(in crate::docx) fn parse_relationships<R: Read + std::io::Seek>(
        zip: &mut zip::ZipArchive<R>,
    ) -> HashMap<String, String> {
        let Some(xml_content) = read_zip_text(zip, "word/_rels/document.xml.rels") else {
            return HashMap::new();
        };
        parse_rels_xml(&xml_content)
    }

    pub(in crate::docx) fn parse_part_relationships<R: Read + std::io::Seek>(
        zip: &mut zip::ZipArchive<R>,
        part_path: &str,
    ) -> HashMap<String, String> {
        let (dir, file) = match part_path.rsplit_once('/') {
            Some((d, f)) => (d, f),
            None => ("", part_path),
        };
        let rels_path = if dir.is_empty() {
            format!("_rels/{}.rels", file)
        } else {
            format!("{}/_rels/{}.rels", dir, file)
        };
        let Some(xml_content) = read_zip_text(zip, &rels_path) else {
            return HashMap::new();
        };
        parse_rels_xml(&xml_content)
    }
}

use relationships::parse_relationships;

pub fn parse(path: &std::path::Path) -> Result<Document, Error> {
    let file = std::fs::File::open(path).map_err(|e| match e.kind() {
        std::io::ErrorKind::NotFound | std::io::ErrorKind::PermissionDenied => Error::Io(
            std::io::Error::new(e.kind(), format!("{}: {}", e, path.display())),
        ),
        _ => Error::Io(e),
    })?;

    let mut zip = zip::ZipArchive::new(file)
        .map_err(|_| Error::InvalidDocx("file is not a ZIP archive".into()))?;

    parse_zip(&mut zip)
}

pub fn parse_bytes(bytes: &[u8]) -> Result<Document, Error> {
    let cursor = std::io::Cursor::new(bytes);
    let mut zip = zip::ZipArchive::new(cursor)
        .map_err(|_| Error::InvalidDocx("data is not a valid ZIP/DOCX archive".into()))?;

    parse_zip(&mut zip)
}

fn parse_zip<R: Read + std::io::Seek>(zip: &mut zip::ZipArchive<R>) -> Result<Document, Error> {
    let theme = parse_theme(zip);
    let styles = parse_styles(zip, &theme);
    let numbering = parse_numbering(zip);
    let rels = parse_relationships(zip);
    let embedded_fonts = parse_font_table(zip);
    let footnotes = parse_footnotes(zip, &styles, &theme);

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")
        .map_err(|_| {
            Error::InvalidDocx("missing word/document.xml (is this a DOCX file?)".into())
        })?
        .read_to_string(&mut xml_content)?;

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = wml(root, "body").ok_or_else(|| Error::Pdf("Missing w:body".into()))?;

    let default_line_pitch = styles.defaults.font_size * 1.2;

    let mut sections: Vec<Section> = Vec::new();
    let mut blocks = Vec::new();
    let mut counters: HashMap<(String, u8), u32> = HashMap::new();
    let mut last_seen_level: HashMap<String, u8> = HashMap::new();

    for node in collect_block_nodes(body) {
        if node.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "tbl" => {
                let col_widths: Vec<f32> = wml(node, "tblGrid")
                    .into_iter()
                    .flat_map(|grid| grid.children())
                    .filter(|n| {
                        n.tag_name().name() == "gridCol"
                            && n.tag_name().namespace() == Some(WML_NS)
                    })
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

                let table_position = tbl_pr
                    .and_then(|pr| wml(pr, "tblpPr"))
                    .map(|tblp| {
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
                        let h_position =
                            if let Some(spec) = tblp.attribute((WML_NS, "tblpXSpec")) {
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

                let tbl_style_borders = tbl_pr
                    .and_then(|pr| wml_attr(pr, "tblStyle"))
                    .and_then(|id| styles.table_border_styles.get(id));

                let tbl_rows: Vec<_> = collect_block_nodes(node)
                    .into_iter()
                    .filter(|n| {
                        n.tag_name().name() == "tr"
                            && n.tag_name().namespace() == Some(WML_NS)
                    })
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
                    let color = n
                        .attribute((WML_NS, "color"))
                        .and_then(parse_hex_color);
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
                    for tc in collect_block_nodes(*tr).into_iter().filter(|n| {
                        n.tag_name().name() == "tc"
                            && n.tag_name().namespace() == Some(WML_NS)
                    }) {
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

                        let style_borders = tbl_style_borders.map(|tb| CellBorders {
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
                                    right: if right.present {
                                        right
                                    } else {
                                        fallback.right
                                    },
                                }
                            })
                            .unwrap_or_else(|| style_borders.unwrap_or_default());

                        let shading = tc_pr
                            .and_then(|pr| wml(pr, "shd"))
                            .and_then(|shd| shd.attribute((WML_NS, "fill")))
                            .filter(|f| *f != "auto" && *f != "none")
                            .and_then(|hex| {
                                if hex.len() == 6 {
                                    Some([
                                        u8::from_str_radix(&hex[0..2], 16).ok()?,
                                        u8::from_str_radix(&hex[2..4], 16).ok()?,
                                        u8::from_str_radix(&hex[4..6], 16).ok()?,
                                    ])
                                } else {
                                    None
                                }
                            });

                        let mut cell_paras = Vec::new();
                        for p in tc.children().filter(|n| {
                            n.tag_name().name() == "p"
                                && n.tag_name().namespace() == Some(WML_NS)
                        }) {
                            let parsed =
                                parse_runs(p, &styles, &theme, &rels, zip, &numbering);
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
                            let line_spacing = Some(
                                inline_spacing
                                    .and_then(|n| {
                                        n.attribute((WML_NS, "line"))
                                            .and_then(|v| v.parse::<f32>().ok())
                                            .map(|line_val| parse_line_spacing(n, line_val))
                                    })
                                    .or_else(|| para_style.and_then(|s| s.line_spacing))
                                    .unwrap_or(LineSpacing::Auto(1.0)),
                            );
                            let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                            let (
                                mut indent_left,
                                mut indent_hanging,
                                list_label,
                                list_label_font,
                            ) = parse_list_info(
                                num_pr,
                                &numbering,
                                &mut counters,
                                &mut last_seen_level,
                            );
                            let mut indent_first_line = 0.0f32;
                            let mut indent_right = 0.0f32;
                            if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                                if let Some(v) = twips_attr(ind, "left") {
                                    indent_left = v;
                                }
                                if let Some(v) = twips_attr(ind, "right") {
                                    indent_right = v;
                                }
                                if let Some(v) = twips_attr(ind, "hanging") {
                                    indent_hanging = v;
                                }
                                if let Some(v) = twips_attr(ind, "firstLine") {
                                    indent_first_line = v;
                                }
                            }
                            cell_paras.push(Paragraph {
                                runs: parsed.runs,
                                space_before: 0.0,
                                space_after: 0.0,
                                content_height: 0.0,
                                alignment,
                                indent_left,
                                indent_right,
                                indent_hanging,
                                indent_first_line,
                                list_label,
                                list_label_font,
                                contextual_spacing: false,
                                keep_next: false,
                                keep_lines: false,
                                line_spacing,
                                image: None,
                                borders: ParagraphBorders::default(),
                                shading: None,
                                page_break_before: false,
                                column_break_before: false,
                                tab_stops: vec![],
                                extra_line_breaks: parsed.line_break_count,
                                floating_images: vec![],
                                textboxes: vec![],
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
                blocks.push(Block::Table(Table {
                    col_widths,
                    rows,
                    table_indent,
                    cell_margins,
                    position: table_position,
                }));
            }
            "p" => {
                let ppr = wml(node, "pPr");

                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"))
                    .unwrap_or("Normal");

                let para_style = styles.paragraph_styles.get(para_style_id);

                let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));

                let inline_borders =
                    ppr.map(parse_paragraph_borders).unwrap_or_default();
                let has_inline_borders = inline_borders.top.is_some()
                    || inline_borders.bottom.is_some()
                    || inline_borders.left.is_some()
                    || inline_borders.right.is_some();
                let borders = if has_inline_borders {
                    inline_borders
                } else {
                    para_style
                        .map(|s| s.borders.clone())
                        .unwrap_or_default()
                };
                let bdr_bottom_extra = borders
                    .bottom
                    .as_ref()
                    .map(|b| b.space_pt + b.width_pt)
                    .unwrap_or(0.0);
                let space_before = inline_spacing
                    .and_then(|n| twips_attr(n, "before"))
                    .or_else(|| para_style.and_then(|s| s.space_before))
                    .unwrap_or(0.0);

                let para_shading = ppr
                    .and_then(|ppr| wml(ppr, "shd"))
                    .and_then(|shd| shd.attribute((WML_NS, "fill")))
                    .and_then(parse_hex_color);
                let space_after = inline_spacing
                    .and_then(|n| twips_attr(n, "after"))
                    .or_else(|| para_style.and_then(|s| s.space_after))
                    .unwrap_or(styles.defaults.space_after)
                    + bdr_bottom_extra;

                let style_color: Option<[u8; 3]> = para_style.and_then(|s| s.color);

                let alignment = ppr
                    .and_then(|ppr| wml_attr(ppr, "jc"))
                    .map(parse_alignment)
                    .or_else(|| para_style.and_then(|s| s.alignment))
                    .unwrap_or(Alignment::Left);

                let contextual_spacing =
                    ppr.and_then(|ppr| wml(ppr, "contextualSpacing")).is_some()
                        || para_style.is_some_and(|s| s.contextual_spacing);

                let keep_next = ppr.and_then(|ppr| wml(ppr, "keepNext")).is_some()
                    || para_style.is_some_and(|s| s.keep_next);

                let keep_lines = ppr.and_then(|ppr| wml(ppr, "keepLines")).is_some()
                    || para_style.is_some_and(|s| s.keep_lines);

                let line_spacing = inline_spacing
                    .and_then(|n| {
                        n.attribute((WML_NS, "line"))
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|line_val| parse_line_spacing(n, line_val))
                    })
                    .or_else(|| para_style.and_then(|s| s.line_spacing));

                let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                let (mut indent_left, mut indent_hanging, list_label, list_label_font) =
                    parse_list_info(
                        num_pr,
                        &numbering,
                        &mut counters,
                        &mut last_seen_level,
                    );

                let mut indent_first_line = 0.0f32;
                let mut indent_right = 0.0f32;
                if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                    if let Some(v) = twips_attr(ind, "left") {
                        indent_left = v;
                    }
                    if let Some(v) = twips_attr(ind, "right") {
                        indent_right = v;
                    }
                    if let Some(v) = twips_attr(ind, "hanging") {
                        indent_hanging = v;
                    }
                    if let Some(v) = twips_attr(ind, "firstLine") {
                        indent_first_line = v;
                    }
                } else if list_label.is_empty()
                    && let Some(s) = para_style
                {
                    if let Some(v) = s.indent_left {
                        indent_left = v;
                    }
                    if let Some(v) = s.indent_right {
                        indent_right = v;
                    }
                    if let Some(v) = s.indent_hanging {
                        indent_hanging = v;
                    }
                    if let Some(v) = s.indent_first_line {
                        indent_first_line = v;
                    }
                }

                let parsed = parse_runs(node, &styles, &theme, &rels, zip, &numbering);
                let mut runs = parsed.runs;

                for run in &mut runs {
                    if run.color.is_none() && style_color.is_some() {
                        run.color = style_color;
                    }
                }

                let tab_stops = ppr.map(parse_tab_stops).unwrap_or_default();

                let has_text = runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
                let has_inline_images = runs.iter().any(|r| r.inline_image.is_some());

                let mut floating_images = parsed.floating_images;

                let (para_image, content_height) = if has_inline_images && !has_text {
                    let img_run_idx = runs.iter().position(|r| r.inline_image.is_some());
                    let img = img_run_idx.and_then(|i| runs[i].inline_image.take());
                    let h = img.as_ref().map(|i| i.display_height).unwrap_or(0.0);
                    (img, h)
                } else if has_inline_images {
                    (None, 0.0)
                } else {
                    let drawing = compute_drawing_info(node, &rels, zip);
                    floating_images.extend(drawing.floating_images);
                    (drawing.image, drawing.height)
                };

                blocks.push(Block::Paragraph(Paragraph {
                    runs,
                    space_before,
                    space_after,
                    content_height,
                    alignment,
                    indent_left,
                    indent_right,
                    indent_hanging,
                    indent_first_line,
                    list_label,
                    list_label_font,
                    contextual_spacing,
                    keep_next,
                    keep_lines,
                    line_spacing,
                    image: para_image,
                    borders,
                    shading: para_shading,
                    page_break_before: parsed.has_page_break,
                    column_break_before: parsed.has_column_break,
                    tab_stops,
                    extra_line_breaks: parsed.line_break_count,
                    floating_images,
                    textboxes: {
                        let mut tbs = parsed.textboxes;
                        tbs.extend(collect_textboxes_from_paragraph(
                            node, &rels, zip, &styles, &theme, &numbering,
                        ));
                        tbs
                    },
                }));

                // Mid-document section break: sectPr inside pPr ends the current section
                if let Some(sect_node) = ppr.and_then(|ppr| wml(ppr, "sectPr")) {
                    let props = parse_section_properties(
                        sect_node,
                        &rels,
                        &styles,
                        &theme,
                        zip,
                        default_line_pitch,
                    );
                    sections.push(Section {
                        properties: props,
                        blocks: std::mem::take(&mut blocks),
                    });
                }
            }
            _ => {}
        }
    }

    // Final section: body-level sectPr
    let final_props = if let Some(sect_node) = wml(body, "sectPr") {
        parse_section_properties(sect_node, &rels, &styles, &theme, zip, default_line_pitch)
    } else {
        SectionProperties {
            page_width: 612.0,
            page_height: 792.0,
            margin_top: 72.0,
            margin_bottom: 72.0,
            margin_left: 72.0,
            margin_right: 72.0,
            header_margin: 36.0,
            footer_margin: 36.0,
            header_default: None,
            header_first: None,
            footer_default: None,
            footer_first: None,
            different_first_page: false,
            line_pitch: default_line_pitch,
            break_type: SectionBreakType::NextPage,
            columns: None,
        }
    };
    sections.push(Section {
        properties: final_props,
        blocks,
    });

    Ok(Document {
        sections,
        line_spacing: styles.defaults.line_spacing,
        embedded_fonts,
        footnotes,
    })
}
