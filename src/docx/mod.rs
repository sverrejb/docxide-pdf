mod alt_chunk;
mod charts;
mod embedded_fonts;
mod headers_footers;
mod images;
mod numbering;
mod runs;
mod sections;
mod settings;
pub(crate) mod smartart;
mod styles;
mod tables;
mod textbox;

use std::collections::HashMap;
use std::io::Read;

use crate::error::Error;
use crate::model::{
    Alignment, Block, Document, LineSpacing, Paragraph, ParagraphBorders, Section,
    SectionBreakType, SectionProperties, TabAlignment, TabStop,
};

use styles::{ParagraphStyle, parse_alignment, parse_line_spacing, parse_styles, parse_theme};

use embedded_fonts::parse_font_table;
use headers_footers::parse_footnotes;
use images::compute_drawing_info;
use numbering::{ListLabelInfo, parse_list_info, parse_numbering};
use relationships::parse_relationships;
use runs::parse_runs;
use sections::parse_section_properties;
use settings::parse_settings;
use tables::parse_table_node;
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

pub(crate) fn is_east_asian_char(ch: char) -> bool {
    matches!(ch as u32,
        0x2E80..=0x2EFF   // CJK Radicals Supplement
        | 0x2F00..=0x2FDF // Kangxi Radicals
        | 0x2FF0..=0x2FFF // Ideographic Description Characters
        | 0x3000..=0x303F // CJK Symbols and Punctuation
        | 0x3040..=0x309F // Hiragana
        | 0x30A0..=0x30FF // Katakana
        | 0x3100..=0x312F // Bopomofo
        | 0x3130..=0x318F // Hangul Compatibility Jamo
        | 0x31A0..=0x31BF // Bopomofo Extended
        | 0x31F0..=0x31FF // Katakana Phonetic Extensions
        | 0x3200..=0x32FF // Enclosed CJK Letters and Months
        | 0x3300..=0x33FF // CJK Compatibility
        | 0x3400..=0x4DBF // CJK Unified Ideographs Extension A
        | 0x4E00..=0x9FFF // CJK Unified Ideographs
        | 0xAC00..=0xD7AF // Hangul Syllables
        | 0xF900..=0xFAFF // CJK Compatibility Ideographs
        | 0xFE30..=0xFE4F // CJK Compatibility Forms
        | 0xFF00..=0xFFEF // Halfwidth and Fullwidth Forms
        | 0x1100..=0x11FF // Hangul Jamo
        | 0x20000..=0x2A6DF // CJK Unified Ideographs Extension B
        | 0x2A700..=0x2B73F // CJK Unified Ideographs Extension C
        | 0x2B740..=0x2B81F // CJK Unified Ideographs Extension D
    )
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

pub(super) fn wml<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
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

pub(super) fn parse_cell_border(parent: roxmltree::Node, name: &str) -> crate::model::CellBorder {
    let Some(n) = wml(parent, name) else {
        return crate::model::CellBorder::default();
    };
    let val = n.attribute((WML_NS, "val")).unwrap_or("none");
    if val == "nil" || val == "none" {
        return crate::model::CellBorder::default();
    }
    let width = n
        .attribute((WML_NS, "sz"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 8.0)
        .unwrap_or(0.5);
    let color = n.attribute((WML_NS, "color")).and_then(parse_hex_color);
    crate::model::CellBorder::visible(color, width)
}

fn parse_cell_border_with_fallback(
    parent: roxmltree::Node,
    primary: &str,
    fallback: &str,
) -> crate::model::CellBorder {
    let border = parse_cell_border(parent, primary);
    if border.present {
        border
    } else {
        parse_cell_border(parent, fallback)
    }
}

/// Parse left border with "start" fallback per OOXML bidi naming.
pub(super) fn parse_cell_border_left(parent: roxmltree::Node) -> crate::model::CellBorder {
    parse_cell_border_with_fallback(parent, "left", "start")
}

/// Parse right border with "end" fallback per OOXML bidi naming.
pub(super) fn parse_cell_border_right(parent: roxmltree::Node) -> crate::model::CellBorder {
    parse_cell_border_with_fallback(parent, "right", "end")
}

pub(super) fn parse_tab_stops(ppr: roxmltree::Node) -> Vec<TabStop> {
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
    stops.sort_by(|a, b| a.position.total_cmp(&b.position));
    stops
}

pub(super) fn resolve_theme_color_key(scheme_name: &str) -> &str {
    match scheme_name {
        "dk1" | "lt1" | "dk2" | "lt2" => scheme_name,
        "tx1" => "dk1",
        "tx2" => "dk2",
        "bg1" => "lt1",
        "bg2" => "lt2",
        other => other,
    }
}

pub(in crate::docx) fn parse_paragraph_spacing(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
) -> (Option<f32>, Option<f32>, Option<LineSpacing>) {
    let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
    let space_before = inline_spacing
        .and_then(|n| twips_attr(n, "before"))
        .or_else(|| para_style.and_then(|s| s.space_before));
    let space_after = inline_spacing
        .and_then(|n| twips_attr(n, "after"))
        .or_else(|| para_style.and_then(|s| s.space_after));
    let line_spacing = inline_spacing
        .and_then(|n| {
            n.attribute((WML_NS, "line"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|line_val| parse_line_spacing(n, line_val))
        })
        .or_else(|| para_style.and_then(|s| s.line_spacing));
    (space_before, space_after, line_spacing)
}

pub(super) fn extract_indents(
    ind: roxmltree::Node,
) -> (Option<f32>, Option<f32>, Option<f32>, Option<f32>) {
    (
        twips_attr(ind, "start").or_else(|| twips_attr(ind, "left")),
        twips_attr(ind, "end").or_else(|| twips_attr(ind, "right")),
        twips_attr(ind, "hanging"),
        twips_attr(ind, "firstLine"),
    )
}

pub(super) fn collect_block_nodes<'a>(
    parent: roxmltree::Node<'a, 'a>,
) -> Vec<roxmltree::Node<'a, 'a>> {
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

mod relationships {
    use std::collections::HashMap;
    use std::io::Read;

    use super::read_zip_text;

    fn parse_rels_xml(xml_content: &str) -> HashMap<String, String> {
        let Ok(xml) = roxmltree::Document::parse(xml_content) else {
            return HashMap::new();
        };
        xml.root_element()
            .children()
            .filter(|n| n.tag_name().name() == "Relationship")
            .filter_map(|n| {
                Some((
                    n.attribute("Id")?.to_string(),
                    n.attribute("Target")?.to_string(),
                ))
            })
            .collect()
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
    let settings = parse_settings(zip);
    let theme = parse_theme(zip, settings.east_asia_lang.as_deref());
    let styles = parse_styles(zip, &theme);
    let numbering = parse_numbering(zip);
    let rels = parse_relationships(zip);
    let ft = parse_font_table(zip);
    let (embedded_fonts, font_table) = (ft.embedded_fonts, ft.font_table);
    let footnotes = parse_footnotes(zip, &styles, &theme);

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")
        .map_err(|_| Error::InvalidDocx("missing word/document.xml (is this a DOCX file?)".into()))?
        .read_to_string(&mut xml_content)?;

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = wml(root, "body").ok_or_else(|| Error::InvalidDocx("Missing w:body".into()))?;

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
                let table = parse_table_node(
                    node,
                    &styles,
                    &theme,
                    &rels,
                    zip,
                    &numbering,
                    &mut counters,
                    &mut last_seen_level,
                );
                blocks.push(Block::Table(table));
            }
            "p" => {
                let ppr = wml(node, "pPr");

                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"))
                    .unwrap_or(&styles.default_paragraph_style_id);

                let para_style = styles.paragraph_styles.get(para_style_id);

                let inline_borders = ppr.map(parse_paragraph_borders).unwrap_or_default();
                let has_inline_borders = inline_borders.top.is_some()
                    || inline_borders.bottom.is_some()
                    || inline_borders.left.is_some()
                    || inline_borders.right.is_some();
                let borders = if has_inline_borders {
                    inline_borders
                } else {
                    para_style.map(|s| s.borders.clone()).unwrap_or_default()
                };
                let bdr_bottom_extra = borders
                    .bottom
                    .as_ref()
                    .map(|b| b.space_pt + b.width_pt)
                    .unwrap_or(0.0);

                let (sp_before, sp_after, line_spacing) = parse_paragraph_spacing(ppr, para_style);
                let space_before = sp_before.unwrap_or(0.0);
                let space_after =
                    sp_after.unwrap_or(styles.defaults.space_after) + bdr_bottom_extra;

                let para_shading = ppr
                    .and_then(|ppr| wml(ppr, "shd"))
                    .and_then(|shd| shd.attribute((WML_NS, "fill")))
                    .and_then(parse_hex_color);

                let style_color = para_style.and_then(|s| s.color);

                let alignment = ppr
                    .and_then(|ppr| wml_attr(ppr, "jc"))
                    .map(parse_alignment)
                    .or_else(|| para_style.and_then(|s| s.alignment))
                    .unwrap_or(Alignment::Left);

                let contextual_spacing = ppr
                    .and_then(|ppr| wml_bool(ppr, "contextualSpacing"))
                    .unwrap_or_else(|| para_style.is_some_and(|s| s.contextual_spacing));

                let keep_next = ppr
                    .and_then(|ppr| wml_bool(ppr, "keepNext"))
                    .unwrap_or_else(|| para_style.is_some_and(|s| s.keep_next));

                let keep_lines = ppr
                    .and_then(|ppr| wml_bool(ppr, "keepLines"))
                    .unwrap_or_else(|| para_style.is_some_and(|s| s.keep_lines));

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
                    &numbering,
                    &mut counters,
                    &mut last_seen_level,
                );

                let mut indent_first_line = 0.0f32;
                let mut indent_right = 0.0f32;
                let (left, right, hanging, first) =
                    if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                        extract_indents(ind)
                    } else if list_label.is_empty()
                        && let Some(s) = para_style
                    {
                        (
                            s.indent_left,
                            s.indent_right,
                            s.indent_hanging,
                            s.indent_first_line,
                        )
                    } else {
                        (None, None, None, None)
                    };
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

                let parsed = parse_runs(node, &styles, &theme, &rels, zip, &numbering);
                let mut runs = parsed.runs;

                if let Some(color) = style_color {
                    for run in &mut runs {
                        run.color.get_or_insert(color);
                    }
                }

                let mut tab_stops = ppr.map(parse_tab_stops).unwrap_or_default();
                if tab_stops.is_empty()
                    && let Some(s) = para_style
                {
                    tab_stops = s.tab_stops.clone();
                }
                // OOXML §17.3.1.38: hanging indent implicitly creates a tab stop
                if indent_hanging > 0.0 {
                    let hang_pos = indent_left;
                    if !tab_stops
                        .iter()
                        .any(|t| (t.position - hang_pos).abs() < 0.5)
                    {
                        tab_stops.push(TabStop {
                            position: hang_pos,
                            alignment: TabAlignment::Left,
                            leader: None,
                        });
                        tab_stops.sort_by(|a, b| a.position.total_cmp(&b.position));
                    }
                }

                let has_text = runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
                let has_inline_images = runs.iter().any(|r| r.inline_image.is_some());

                let mut floating_images = parsed.floating_images;

                let (para_image, mut content_height) = if has_inline_images && !has_text {
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

                if let Some(ref ic) = parsed.inline_chart {
                    content_height = content_height.max(ic.display_height);
                }
                if let Some(ref sa) = parsed.smartart {
                    content_height = content_height.max(sa.display_height);
                }

                blocks.push(Block::Paragraph(Paragraph {
                    runs,
                    style_id: Some(para_style_id.to_string()),
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
                    list_label_font_size,
                    list_label_bold,
                    list_label_color,
                    contextual_spacing,
                    keep_next,
                    keep_lines,
                    line_spacing,
                    image: para_image,
                    borders,
                    shading: para_shading,
                    page_break_before: parsed.has_page_break_before
                        || para_style.is_some_and(|s| s.page_break_before),
                    page_break_after: parsed.has_page_break_after,
                    column_break_before: parsed.has_column_break,
                    tab_stops,
                    floating_images,
                    textboxes: {
                        let mut tbs = parsed.textboxes;
                        tbs.extend(collect_textboxes_from_paragraph(
                            node, &rels, zip, &styles, &theme, &numbering,
                        ));
                        tbs
                    },
                    connectors: parsed.connectors,
                    inline_chart: parsed.inline_chart,
                    smartart: parsed.smartart,
                    is_section_break: false,
                }));

                // Mid-document section break: sectPr inside pPr ends the current section
                if let Some(sect_node) = ppr.and_then(|ppr| wml(ppr, "sectPr")) {
                    if let Some(Block::Paragraph(last_para)) = blocks.last_mut() {
                        last_para.is_section_break = true;
                    }
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
            "altChunk" => {
                if let Some(id) = node.attribute((REL_NS, "id")) {
                    blocks.extend(alt_chunk::parse_alt_chunk(id, &rels, zip));
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
            header_even: None,
            footer_default: None,
            footer_first: None,
            footer_even: None,
            different_first_page: false,
            line_pitch: default_line_pitch,
            break_type: SectionBreakType::NextPage,
            columns: None,
            page_num_start: None,
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
        font_table,
        even_and_odd_headers: settings.even_and_odd_headers,
        style_id_to_name: styles.style_id_to_name,
    })
}
