mod styles;

use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

use crate::error::Error;
use crate::model::{
    Alignment, Block, CellBorder, CellBorders, CellMargins, CellVAlign, ColumnDef, ColumnsConfig,
    Document, EmbeddedImage, FieldCode, FloatingImage, Footnote, HeaderFooter, HorizontalPosition,
    ImageFormat, LineSpacing, Paragraph, ParagraphBorder, ParagraphBorders, Run, Section,
    SectionBreakType, SectionProperties, TabAlignment, TabStop, Table, TableCell, TableRow, VMerge,
    VertAlign,
};

use styles::{
    StylesInfo, ThemeFonts, parse_alignment, parse_line_spacing, parse_styles, parse_theme,
    resolve_font_from_node,
};

struct LevelDef {
    num_fmt: String,
    lvl_text: String,
    indent_left: f32,
    indent_hanging: f32,
    start: u32,
}

struct NumberingInfo {
    abstract_nums: HashMap<String, HashMap<u8, LevelDef>>,
    num_to_abstract: HashMap<String, String>,
}

pub(super) const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPD_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

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

/// Parse a WML boolean toggle element (e.g., w:b, w:i, w:strike).
/// Present with no val or val != "0"/"false" means true.
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

fn parse_one_border(node: roxmltree::Node) -> Option<ParagraphBorder> {
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
    Some(ParagraphBorder {
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
    }
}

pub(super) fn paragraph_borders_extra(ppr: roxmltree::Node) -> f32 {
    let borders = parse_paragraph_borders(ppr);
    borders
        .bottom
        .as_ref()
        .map(|b| b.space_pt + b.width_pt)
        .unwrap_or(0.0)
}

/// Parse GUID string like "{302EE813-EB4A-4642-A93A-89EF99B2457E}" into 16 bytes.
/// Returns bytes in standard GUID mixed-endian layout, then reversed to big-endian.
fn parse_guid_to_bytes(guid: &str) -> Option<[u8; 16]> {
    let hex: String = guid.chars().filter(|c| c.is_ascii_hexdigit()).collect();
    if hex.len() != 32 {
        return None;
    }
    let mut bytes = [0u8; 16];
    for i in 0..16 {
        bytes[i] = u8::from_str_radix(&hex[i * 2..i * 2 + 2], 16).ok()?;
    }
    // Standard GUID byte order: first 4 bytes LE, next 2 LE, next 2 LE, rest big-endian
    // Convert from the string representation to actual GUID byte layout
    let guid_bytes: [u8; 16] = [
        bytes[3], bytes[2], bytes[1], bytes[0], // Data1 (LE)
        bytes[5], bytes[4], // Data2 (LE)
        bytes[7], bytes[6], // Data3 (LE)
        bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15],
    ];
    // Reverse for XOR key per spec §17.8.1
    let mut reversed = guid_bytes;
    reversed.reverse();
    Some(reversed)
}

/// Deobfuscate an embedded DOCX font by XORing the first 32 bytes with the reversed GUID key.
fn deobfuscate_font(data: &mut [u8], key: &[u8; 16]) {
    for i in 0..16.min(data.len()) {
        data[i] ^= key[i];
    }
    for i in 16..32.min(data.len()) {
        data[i] ^= key[i - 16];
    }
}

/// Parse word/_rels/fontTable.xml.rels to get relationship ID → target path mapping.
fn parse_font_table_rels(zip: &mut zip::ZipArchive<std::fs::File>) -> HashMap<String, String> {
    let mut rels = HashMap::new();
    let Some(xml_content) = read_zip_text(zip, "word/_rels/fontTable.xml.rels") else {
        return rels;
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return rels;
    };
    for node in xml.root_element().children() {
        if node.tag_name().name() == "Relationship"
            && let (Some(id), Some(target)) = (node.attribute("Id"), node.attribute("Target"))
        {
            rels.insert(id.to_string(), target.to_string());
        }
    }
    rels
}

struct EmbedInfo {
    font_name: String,
    bold: bool,
    italic: bool,
    rel_id: String,
    font_key: Option<String>,
}

/// Parse word/fontTable.xml for embedded fonts, extract and deobfuscate them.
fn parse_font_table(
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> HashMap<(String, bool, bool), Vec<u8>> {
    let mut result = HashMap::new();

    let embeds = {
        let Some(xml_content) = read_zip_text(zip, "word/fontTable.xml") else {
            return result;
        };
        let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
            return result;
        };

        let embed_variants: &[(&str, bool, bool)] = &[
            ("embedRegular", false, false),
            ("embedBold", true, false),
            ("embedItalic", false, true),
            ("embedBoldItalic", true, true),
        ];

        let mut embeds = Vec::new();
        for font_node in xml.root_element().children() {
            if font_node.tag_name().name() != "font"
                || font_node.tag_name().namespace() != Some(WML_NS)
            {
                continue;
            }
            let Some(font_name) = font_node.attribute((WML_NS, "name")) else {
                continue;
            };

            for &(embed_tag, bold, italic) in embed_variants {
                let Some(embed_node) = wml(font_node, embed_tag) else {
                    continue;
                };
                let Some(r_id) = embed_node.attribute((REL_NS, "id")) else {
                    continue;
                };
                let font_key = embed_node
                    .attribute((WML_NS, "fontKey"))
                    .map(|s| s.to_string());

                embeds.push(EmbedInfo {
                    font_name: font_name.to_string(),
                    bold,
                    italic,
                    rel_id: r_id.to_string(),
                    font_key,
                });
            }
        }
        embeds
    };

    if embeds.is_empty() {
        return result;
    }

    // Phase 2: resolve relationships and extract font data
    let font_rels = parse_font_table_rels(zip);

    for info in embeds {
        let Some(target) = font_rels.get(&info.rel_id) else {
            continue;
        };

        let zip_path = target
            .strip_prefix('/')
            .map(String::from)
            .unwrap_or_else(|| format!("word/{}", target));

        let Ok(mut entry) = zip.by_name(&zip_path) else {
            continue;
        };
        let mut data = Vec::new();
        if entry.read_to_end(&mut data).is_err() {
            continue;
        }
        drop(entry);

        if let Some(ref guid_str) = info.font_key
            && let Some(key) = parse_guid_to_bytes(guid_str)
        {
            deobfuscate_font(&mut data, &key);
        }

        log::info!(
            "Extracted embedded font: {} bold={} italic={} ({} bytes)",
            info.font_name,
            info.bold,
            info.italic,
            data.len()
        );
        result.insert(
            (info.font_name.to_lowercase(), info.bold, info.italic),
            data,
        );
    }

    result
}

fn parse_numbering(zip: &mut zip::ZipArchive<std::fs::File>) -> NumberingInfo {
    let mut abstract_nums: HashMap<String, HashMap<u8, LevelDef>> = HashMap::new();
    let mut num_to_abstract: HashMap<String, String> = HashMap::new();

    let Some(xml_content) = read_zip_text(zip, "word/numbering.xml") else {
        return NumberingInfo {
            abstract_nums,
            num_to_abstract,
        };
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return NumberingInfo {
            abstract_nums,
            num_to_abstract,
        };
    };

    let root = xml.root_element();

    for node in root.children() {
        if node.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "abstractNum" => {
                let Some(abs_id) = node.attribute((WML_NS, "abstractNumId")) else {
                    continue;
                };
                let mut levels: HashMap<u8, LevelDef> = HashMap::new();
                for lvl in node.children() {
                    if lvl.tag_name().name() != "lvl" || lvl.tag_name().namespace() != Some(WML_NS)
                    {
                        continue;
                    }
                    let Some(ilvl) = lvl
                        .attribute((WML_NS, "ilvl"))
                        .and_then(|v| v.parse::<u8>().ok())
                    else {
                        continue;
                    };
                    let num_fmt = wml_attr(lvl, "numFmt").unwrap_or("bullet").to_string();
                    let lvl_text = wml_attr(lvl, "lvlText").unwrap_or("").to_string();
                    let start = wml_attr(lvl, "start")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1);
                    let ind = wml(lvl, "pPr").and_then(|ppr| wml(ppr, "ind"));
                    let indent_left = ind.and_then(|n| twips_attr(n, "left")).unwrap_or(0.0);
                    let indent_hanging = ind.and_then(|n| twips_attr(n, "hanging")).unwrap_or(0.0);
                    levels.insert(
                        ilvl,
                        LevelDef {
                            num_fmt,
                            lvl_text,
                            indent_left,
                            indent_hanging,
                            start,
                        },
                    );
                }
                abstract_nums.insert(abs_id.to_string(), levels);
            }
            "num" => {
                let Some(num_id) = node.attribute((WML_NS, "numId")) else {
                    continue;
                };
                let Some(abs_id) = wml_attr(node, "abstractNumId") else {
                    continue;
                };
                num_to_abstract.insert(num_id.to_string(), abs_id.to_string());
            }
            _ => {}
        }
    }

    NumberingInfo {
        abstract_nums,
        num_to_abstract,
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

/// Flatten SDT wrappers: descend into w:sdtContent and collect effective children.
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

struct ParsedRuns {
    runs: Vec<Run>,
    has_page_break: bool,
    has_column_break: bool,
    line_break_count: u32,
    floating_images: Vec<FloatingImage>,
}

fn parse_runs(
    para_node: roxmltree::Node,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> ParsedRuns {
    let ppr = wml(para_node, "pPr");
    let para_style_id = ppr
        .and_then(|ppr| wml_attr(ppr, "pStyle"))
        .unwrap_or("Normal");
    let para_style = styles.paragraph_styles.get(para_style_id);

    let style_font_size = para_style
        .and_then(|s| s.font_size)
        .unwrap_or(styles.defaults.font_size);
    let style_font_name = para_style
        .and_then(|s| s.font_name.as_deref())
        .unwrap_or(&styles.defaults.font_name)
        .to_string();
    let style_bold = para_style.and_then(|s| s.bold).unwrap_or(false);
    let style_italic = para_style.and_then(|s| s.italic).unwrap_or(false);
    let style_caps = para_style.and_then(|s| s.caps).unwrap_or(false);
    let style_small_caps = para_style.and_then(|s| s.small_caps).unwrap_or(false);
    let style_vanish = para_style.and_then(|s| s.vanish).unwrap_or(false);
    let style_color: Option<[u8; 3]> = para_style.and_then(|s| s.color);

    fn collect_run_nodes<'a>(
        parent: roxmltree::Node<'a, 'a>,
        rels: &HashMap<String, String>,
        out: &mut Vec<(roxmltree::Node<'a, 'a>, Option<String>)>,
    ) {
        for child in parent.children() {
            let name = child.tag_name().name();
            let is_wml = child.tag_name().namespace() == Some(WML_NS);
            if is_wml && name == "r" {
                out.push((child, None));
            } else if is_wml && name == "hyperlink" {
                let url = child
                    .attribute((REL_NS, "id"))
                    .and_then(|rid| rels.get(rid))
                    .cloned();
                for n in child.children().filter(|n| {
                    n.tag_name().name() == "r" && n.tag_name().namespace() == Some(WML_NS)
                }) {
                    out.push((n, url.clone()));
                }
            } else if is_wml && name == "sdt" {
                if let Some(content) = wml(child, "sdtContent") {
                    collect_run_nodes(content, rels, out);
                }
            }
        }
    }
    let mut run_nodes: Vec<(roxmltree::Node, Option<String>)> = Vec::new();
    collect_run_nodes(para_node, rels, &mut run_nodes);

    let mut runs = Vec::new();
    let mut floating_images: Vec<FloatingImage> = Vec::new();
    let mut has_page_break = false;
    let mut has_column_break = false;
    let mut line_break_count: u32 = 0;
    let mut in_field = false;
    let mut field_instr = String::new();

    for (run_node, hyperlink_url) in run_nodes {
        let rpr = wml(run_node, "rPr");

        let char_style = rpr
            .and_then(|n| wml_attr(n, "rStyle"))
            .and_then(|id| styles.character_styles.get(id));

        let font_size = rpr
            .and_then(|n| wml_attr(n, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0)
            .or_else(|| char_style.and_then(|cs| cs.font_size))
            .unwrap_or(style_font_size);

        let font_name = rpr
            .and_then(|n| wml(n, "rFonts"))
            .map(|rfonts| resolve_font_from_node(rfonts, theme, &style_font_name))
            .or_else(|| char_style.and_then(|cs| cs.font_name.clone()))
            .unwrap_or_else(|| style_font_name.clone());

        let bold = rpr
            .and_then(|n| wml_bool(n, "b"))
            .or_else(|| char_style.and_then(|cs| cs.bold))
            .unwrap_or(style_bold);
        let italic = rpr
            .and_then(|n| wml_bool(n, "i"))
            .or_else(|| char_style.and_then(|cs| cs.italic))
            .unwrap_or(style_italic);
        let underline = rpr
            .and_then(|n| {
                wml(n, "u")
                    .and_then(|u| u.attribute((WML_NS, "val")))
                    .map(|v| v != "none")
            })
            .or_else(|| char_style.and_then(|cs| cs.underline))
            .unwrap_or(false);
        let strikethrough = rpr
            .and_then(|n| wml_bool(n, "strike"))
            .or_else(|| char_style.and_then(|cs| cs.strikethrough))
            .unwrap_or(false);
        let caps = rpr
            .and_then(|n| wml_bool(n, "caps"))
            .or_else(|| char_style.and_then(|cs| cs.caps))
            .unwrap_or(style_caps);
        let small_caps = rpr
            .and_then(|n| wml_bool(n, "smallCaps"))
            .or_else(|| char_style.and_then(|cs| cs.small_caps))
            .unwrap_or(style_small_caps);
        let vanish = rpr
            .and_then(|n| wml_bool(n, "vanish"))
            .or_else(|| char_style.and_then(|cs| cs.vanish))
            .unwrap_or(style_vanish);

        let color = rpr
            .and_then(|n| wml_attr(n, "color"))
            .and_then(parse_text_color)
            .or_else(|| char_style.and_then(|cs| cs.color))
            .or(style_color);

        let vertical_align = rpr
            .and_then(|n| wml_attr(n, "vertAlign"))
            .map(|v| match v {
                "superscript" => VertAlign::Superscript,
                "subscript" => VertAlign::Subscript,
                _ => VertAlign::Baseline,
            })
            .unwrap_or(VertAlign::Baseline);

        let highlight = rpr
            .and_then(|n| wml_attr(n, "highlight"))
            .and_then(highlight_color);

        // Iterate children in document order to handle w:t, w:tab, w:br, w:fldChar, w:instrText
        let mut pending_text = String::new();
        for child in run_node.children() {
            if child.tag_name().namespace() != Some(WML_NS) {
                continue;
            }
            match child.tag_name().name() {
                "fldChar" => {
                    match child.attribute((WML_NS, "fldCharType")) {
                        Some("begin") => {
                            // Flush pending text before entering field
                            if !pending_text.is_empty() {
                                runs.push(Run {
                                    text: std::mem::take(&mut pending_text),
                                    font_size,
                                    font_name: font_name.clone(),
                                    bold,
                                    italic,
                                    underline,
                                    strikethrough,
                                    caps,
                                    small_caps,
                                    vanish,
                                    color,
                                    is_tab: false,
                                    vertical_align,
                                    field_code: None,
                                    hyperlink_url: hyperlink_url.clone(),
                                    highlight,
                                    inline_image: None,
                                    footnote_id: None,
                                    is_footnote_ref_mark: false,
                                });
                            }
                            in_field = true;
                            field_instr.clear();
                        }
                        Some("end") => {
                            if in_field {
                                let trimmed = field_instr.trim();
                                let fc = if trimmed.eq_ignore_ascii_case("PAGE") {
                                    Some(FieldCode::Page)
                                } else if trimmed.eq_ignore_ascii_case("NUMPAGES") {
                                    Some(FieldCode::NumPages)
                                } else {
                                    None
                                };
                                if let Some(code) = fc {
                                    runs.push(Run {
                                        text: String::new(),
                                        font_size,
                                        font_name: font_name.clone(),
                                        bold,
                                        italic,
                                        underline: false,
                                        strikethrough: false,
                                        caps: false,
                                        small_caps: false,
                                        vanish: false,
                                        color,
                                        is_tab: false,
                                        vertical_align: VertAlign::Baseline,
                                        field_code: Some(code),
                                        hyperlink_url: hyperlink_url.clone(),
                                        highlight: None,
                                        inline_image: None,
                                        footnote_id: None,
                                        is_footnote_ref_mark: false,
                                    });
                                }
                                in_field = false;
                                field_instr.clear();
                            }
                        }
                        _ => {}
                    }
                }
                "instrText" if in_field => {
                    if let Some(t) = child.text() {
                        field_instr.push_str(t);
                    }
                }
                "t" if !in_field => {
                    if let Some(t) = child.text() {
                        // Word treats newlines in w:t as whitespace; only w:br creates line breaks
                        let normalized = t.replace('\n', " ");
                        pending_text.push_str(&normalized);
                    }
                }
                "tab" if !in_field => {
                    // Flush any pending text before the tab
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    // Insert tab marker run
                    runs.push(Run {
                        text: String::new(),
                        font_size,
                        font_name: font_name.clone(),
                        bold: false,
                        italic: false,
                        underline: false,
                        strikethrough: false,
                        caps: false,
                        small_caps: false,
                        vanish: false,
                        color: None,
                        is_tab: true,
                        vertical_align: VertAlign::Baseline,
                        field_code: None,
                        hyperlink_url: None,
                        highlight: None,
                        inline_image: None,
                        footnote_id: None,
                        is_footnote_ref_mark: false,
                    });
                }
                "br" if !in_field => {
                    match child.attribute((WML_NS, "type")) {
                        Some("page") => has_page_break = true,
                        Some("column") => has_column_break = true,
                        _ => line_break_count += 1,
                    }
                }
                "drawing" if !in_field => {
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    match parse_run_drawing(child, rels, zip) {
                        Some(RunDrawingResult::Inline(img)) => {
                            runs.push(Run {
                                text: String::new(),
                                font_size,
                                font_name: font_name.clone(),
                                bold: false,
                                italic: false,
                                underline: false,
                                strikethrough: false,
                                caps: false,
                                small_caps: false,
                                vanish: false,
                                color: None,
                                is_tab: false,
                                vertical_align: VertAlign::Baseline,
                                field_code: None,
                                hyperlink_url: None,
                                highlight: None,
                                inline_image: Some(img),
                                footnote_id: None,
                                is_footnote_ref_mark: false,
                            });
                        }
                        Some(RunDrawingResult::Floating(fi)) => {
                            floating_images.push(fi);
                        }
                        None => {}
                    }
                }
                "footnoteReference" if !in_field => {
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    if let Some(id) = child
                        .attribute((WML_NS, "id"))
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        runs.push(Run {
                            text: String::new(),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline: false,
                            strikethrough: false,
                            caps: false,
                            small_caps: false,
                            vanish: false,
                            color,
                            is_tab: false,
                            vertical_align: VertAlign::Superscript,
                            field_code: None,
                            hyperlink_url: None,
                            highlight: None,
                            inline_image: None,
                            footnote_id: Some(id),
                            is_footnote_ref_mark: false,
                        });
                    }
                }
                "footnoteRef" if !in_field => {
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    runs.push(Run {
                        text: String::new(),
                        font_size,
                        font_name: font_name.clone(),
                        bold,
                        italic,
                        underline: false,
                        strikethrough: false,
                        caps: false,
                        small_caps: false,
                        vanish: false,
                        color,
                        is_tab: false,
                        vertical_align: VertAlign::Superscript,
                        field_code: None,
                        hyperlink_url: None,
                        highlight: None,
                        inline_image: None,
                        footnote_id: None,
                        is_footnote_ref_mark: true,
                    });
                }
                _ => {}
            }
        }
        // Flush remaining text
        if !pending_text.is_empty() {
            runs.push(Run {
                text: pending_text,
                font_size,
                font_name,
                bold,
                italic,
                underline,
                strikethrough,
                caps,
                small_caps,
                vanish,
                color,
                is_tab: false,
                vertical_align,
                field_code: None,
                hyperlink_url: hyperlink_url.clone(),
                highlight,
                inline_image: None,
                footnote_id: None,
                is_footnote_ref_mark: false,
            });
        }
    }

    if ppr
        .and_then(|ppr| wml_bool(ppr, "pageBreakBefore"))
        .unwrap_or(false)
    {
        has_page_break = true;
    }

    // Empty paragraphs with explicit font sizing in their paragraph mark (pPr/rPr)
    // need a synthetic run so the renderer computes the correct line height.
    if runs.is_empty() && !has_page_break {
        let mark_rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
        let has_explicit_sz = mark_rpr.and_then(|n| wml_attr(n, "sz")).is_some();
        if has_explicit_sz {
            let mark_font_size = mark_rpr
                .and_then(|n| wml_attr(n, "sz"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|hp| hp / 2.0)
                .unwrap_or(style_font_size);
            let mark_font_name = mark_rpr
                .and_then(|n| wml(n, "rFonts"))
                .map(|rfonts| resolve_font_from_node(rfonts, theme, &style_font_name))
                .unwrap_or_else(|| style_font_name.clone());
            runs.push(Run {
                text: String::new(),
                font_size: mark_font_size,
                font_name: mark_font_name,
                bold: style_bold,
                italic: style_italic,
                underline: false,
                strikethrough: false,
                caps: false,
                small_caps: false,
                vanish: false,
                color: None,
                highlight: None,
                is_tab: false,
                vertical_align: VertAlign::Baseline,
                field_code: None,
                hyperlink_url: None,
                inline_image: None,
                footnote_id: None,
                is_footnote_ref_mark: false,
            });
        }
    }

    // Word's paragraph mark (¶) uses the paragraph style's font even in empty
    // paragraphs; ensure we carry that font info so line height is correct.
    if runs.is_empty() {
        runs.push(Run {
            text: String::new(),
            font_size: style_font_size,
            font_name: style_font_name.clone(),
            bold: style_bold,
            italic: style_italic,
            underline: false,
            strikethrough: false,
            caps: false,
            small_caps: false,
            vanish: false,
            color: None,
            highlight: None,
            is_tab: false,
            vertical_align: VertAlign::Baseline,
            field_code: None,
            hyperlink_url: None,
            inline_image: None,
            footnote_id: None,
            is_footnote_ref_mark: false,
        });
    }

    ParsedRuns {
        runs,
        has_page_break,
        has_column_break,
        line_break_count,
        floating_images,
    }
}

fn parse_header_footer_xml(
    xml_content: &str,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> Option<HeaderFooter> {
    let xml = roxmltree::Document::parse(xml_content).ok()?;
    let root = xml.root_element();
    let mut paragraphs = Vec::new();

    for node in root.children() {
        if node.tag_name().namespace() != Some(WML_NS) || node.tag_name().name() != "p" {
            continue;
        }
        let ppr = wml(node, "pPr");
        let para_style_id = ppr
            .and_then(|ppr| wml_attr(ppr, "pStyle"))
            .unwrap_or("Normal");
        let para_style = styles.paragraph_styles.get(para_style_id);

        let alignment = ppr
            .and_then(|ppr| wml_attr(ppr, "jc"))
            .map(parse_alignment)
            .or_else(|| para_style.and_then(|s| s.alignment))
            .unwrap_or(Alignment::Left);

        let parsed = parse_runs(node, styles, theme, rels, zip);
        let mut runs = parsed.runs;
        let mut floating_images = parsed.floating_images;

        let has_inline_images = runs.iter().any(|r| r.inline_image.is_some());
        let has_text = runs.iter().any(|r| !r.text.is_empty());
        let (para_image, content_height) = if has_inline_images && !has_text {
            let img_run_idx = runs.iter().position(|r| r.inline_image.is_some());
            let img = img_run_idx.and_then(|i| runs[i].inline_image.take());
            let h = img.as_ref().map(|i| i.display_height).unwrap_or(0.0);
            (img, h)
        } else if has_inline_images {
            (None, 0.0)
        } else {
            let drawing = compute_drawing_info(node, rels, zip);
            floating_images.extend(drawing.floating_images);
            (drawing.image, drawing.height)
        };

        paragraphs.push(Paragraph {
            runs,
            space_before: 0.0,
            space_after: 0.0,
            content_height,
            alignment,
            indent_left: 0.0,
            indent_right: 0.0,
            indent_hanging: 0.0,
            indent_first_line: 0.0,
            list_label: String::new(),
            contextual_spacing: false,
            keep_next: false,
            line_spacing: None,
            image: para_image,
            borders: ParagraphBorders::default(),
            shading: None,
            page_break_before: false,
            column_break_before: false,
            tab_stops: vec![],
            extra_line_breaks: parsed.line_break_count,
            floating_images,
        });
    }

    if paragraphs.is_empty() {
        None
    } else {
        Some(HeaderFooter { paragraphs })
    }
}

fn parse_footnotes(
    zip: &mut zip::ZipArchive<std::fs::File>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
) -> HashMap<u32, Footnote> {
    let mut footnotes = HashMap::new();
    let Some(xml_text) = read_zip_text(zip, "word/footnotes.xml") else {
        return footnotes;
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_text) else {
        return footnotes;
    };
    let root = xml.root_element();

    for node in root.children() {
        if node.tag_name().namespace() != Some(WML_NS) || node.tag_name().name() != "footnote" {
            continue;
        }
        // Skip separator/continuationSeparator footnotes (type attribute, IDs 0 and 1)
        if node.attribute((WML_NS, "type")).is_some() {
            continue;
        }
        let Some(id) = node
            .attribute((WML_NS, "id"))
            .and_then(|v| v.parse::<u32>().ok())
        else {
            continue;
        };

        let mut paragraphs = Vec::new();
        let empty_rels = HashMap::new();
        for p in node.children() {
            if p.tag_name().namespace() != Some(WML_NS) || p.tag_name().name() != "p" {
                continue;
            }
            let ppr = wml(p, "pPr");
            let para_style_id = ppr
                .and_then(|ppr| wml_attr(ppr, "pStyle"))
                .unwrap_or("FootnoteText");
            let para_style = styles.paragraph_styles.get(para_style_id);

            let alignment = ppr
                .and_then(|ppr| wml_attr(ppr, "jc"))
                .map(parse_alignment)
                .or_else(|| para_style.and_then(|s| s.alignment))
                .unwrap_or(Alignment::Left);

            let parsed = parse_runs(p, styles, theme, &empty_rels, zip);

            let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
            let space_before = inline_spacing
                .and_then(|n| twips_attr(n, "before"))
                .or_else(|| para_style.and_then(|s| s.space_before))
                .unwrap_or(0.0);
            let space_after = inline_spacing
                .and_then(|n| twips_attr(n, "after"))
                .or_else(|| para_style.and_then(|s| s.space_after))
                .unwrap_or(0.0);
            let line_spacing = inline_spacing
                .and_then(|n| {
                    n.attribute((WML_NS, "line"))
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|line_val| parse_line_spacing(n, line_val))
                })
                .or_else(|| para_style.and_then(|s| s.line_spacing))
                .or(Some(LineSpacing::Auto(1.0)));

            paragraphs.push(Paragraph {
                runs: parsed.runs,
                space_before,
                space_after,
                content_height: 0.0,
                alignment,
                indent_left: 0.0,
                indent_right: 0.0,
                indent_hanging: 0.0,
                indent_first_line: 0.0,
                list_label: String::new(),
                contextual_spacing: false,
                keep_next: false,
                line_spacing,
                image: None,
                borders: ParagraphBorders::default(),
                shading: None,
                page_break_before: false,
                column_break_before: false,
                tab_stops: vec![],
                extra_line_breaks: parsed.line_break_count,
                floating_images: vec![],
            });
        }

        if !paragraphs.is_empty() {
            footnotes.insert(id, Footnote { paragraphs });
        }
    }

    footnotes
}

pub(super) fn read_zip_text(zip: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Option<String> {
    let mut content = String::new();
    zip.by_name(name).ok()?.read_to_string(&mut content).ok()?;
    Some(content)
}

fn parse_section_properties(
    sect_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    zip: &mut zip::ZipArchive<std::fs::File>,
    default_line_pitch: f32,
) -> SectionProperties {
    let pg_sz = wml(sect_node, "pgSz");
    let pg_mar = wml(sect_node, "pgMar");
    let doc_grid = wml(sect_node, "docGrid");

    let page_width = pg_sz.and_then(|n| twips_attr(n, "w")).unwrap_or(612.0);
    let page_height = pg_sz.and_then(|n| twips_attr(n, "h")).unwrap_or(792.0);
    let margin_top = pg_mar.and_then(|n| twips_attr(n, "top")).unwrap_or(72.0);
    let margin_bottom = pg_mar.and_then(|n| twips_attr(n, "bottom")).unwrap_or(72.0);
    let margin_left = pg_mar.and_then(|n| twips_attr(n, "left")).unwrap_or(72.0);
    let margin_right = pg_mar.and_then(|n| twips_attr(n, "right")).unwrap_or(72.0);
    let header_margin = pg_mar.and_then(|n| twips_attr(n, "header")).unwrap_or(36.0);
    let footer_margin = pg_mar.and_then(|n| twips_attr(n, "footer")).unwrap_or(36.0);
    let line_pitch = doc_grid
        .and_then(|n| twips_attr(n, "linePitch"))
        .unwrap_or(default_line_pitch);

    let different_first_page = wml(sect_node, "titlePg").is_some();

    let break_type = wml(sect_node, "type")
        .and_then(|n| n.attribute((WML_NS, "val")))
        .map(|v| match v {
            "continuous" => SectionBreakType::Continuous,
            "oddPage" => SectionBreakType::OddPage,
            "evenPage" => SectionBreakType::EvenPage,
            _ => SectionBreakType::NextPage,
        })
        .unwrap_or(SectionBreakType::NextPage);

    let available = page_width - margin_left - margin_right;
    let columns = wml(sect_node, "cols").and_then(|cols_node| {
        let num: u32 = cols_node
            .attribute((WML_NS, "num"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(1);
        let equal_width = cols_node
            .attribute((WML_NS, "equalWidth"))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(true);
        let sep = cols_node
            .attribute((WML_NS, "sep"))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false);

        let child_cols: Vec<_> = cols_node
            .children()
            .filter(|c| c.tag_name().name() == "col" && c.tag_name().namespace() == Some(WML_NS))
            .collect();

        let col_defs: Vec<ColumnDef> = if !equal_width && !child_cols.is_empty() {
            child_cols
                .iter()
                .map(|c| {
                    let w = twips_attr(*c, "w").unwrap_or(0.0);
                    let sp = twips_attr(*c, "space").unwrap_or(0.0);
                    ColumnDef { width: w, space: sp }
                })
                .collect()
        } else if num > 1 {
            let default_space = cols_node
                .attribute((WML_NS, "space"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(36.0);
            let col_width = (available - (num - 1) as f32 * default_space) / num as f32;
            (0..num)
                .map(|i| ColumnDef {
                    width: col_width.max(1.0),
                    space: if i < num - 1 { default_space } else { 0.0 },
                })
                .collect()
        } else {
            return None;
        };

        Some(ColumnsConfig {
            columns: col_defs,
            sep,
        })
    });

    let mut header_default_rid = None;
    let mut header_first_rid = None;
    let mut footer_default_rid = None;
    let mut footer_first_rid = None;
    for child in sect_node.children() {
        if child.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        let hf_type = child.attribute((WML_NS, "type")).unwrap_or("");
        let rid = child.attribute((REL_NS, "id"));
        match child.tag_name().name() {
            "headerReference" => match hf_type {
                "default" => header_default_rid = rid,
                "first" => header_first_rid = rid,
                _ => {}
            },
            "footerReference" => match hf_type {
                "default" => footer_default_rid = rid,
                "first" => footer_first_rid = rid,
                _ => {}
            },
            _ => {}
        }
    }

    let resolve_hf =
        |rid: Option<&str>, zip: &mut zip::ZipArchive<std::fs::File>| -> Option<HeaderFooter> {
            let target = rels.get(rid?)?;
            let zip_path = target
                .strip_prefix('/')
                .map(String::from)
                .unwrap_or_else(|| format!("word/{}", target));
            let part_rels = parse_part_relationships(zip, &zip_path);
            let xml_text = read_zip_text(zip, &zip_path)?;
            parse_header_footer_xml(&xml_text, styles, theme, &part_rels, zip)
        };

    let header_default = resolve_hf(header_default_rid, zip);
    let header_first = resolve_hf(header_first_rid, zip);
    let footer_default = resolve_hf(footer_default_rid, zip);
    let footer_first = resolve_hf(footer_first_rid, zip);

    SectionProperties {
        page_width,
        page_height,
        margin_top,
        margin_bottom,
        margin_left,
        margin_right,
        header_margin,
        footer_margin,
        header_default,
        header_first,
        footer_default,
        footer_first,
        different_first_page,
        line_pitch,
        break_type,
        columns,
    }
}

pub fn parse(path: &Path) -> Result<Document, Error> {
    let file = std::fs::File::open(path).map_err(|e| match e.kind() {
        std::io::ErrorKind::NotFound | std::io::ErrorKind::PermissionDenied => Error::Io(
            std::io::Error::new(e.kind(), format!("{}: {}", e, path.display())),
        ),
        _ => Error::Io(e),
    })?;

    let mut zip = zip::ZipArchive::new(file)
        .map_err(|_| Error::InvalidDocx("file is not a ZIP archive".into()))?;

    let theme = parse_theme(&mut zip);
    let styles = parse_styles(&mut zip, &theme);
    let numbering = parse_numbering(&mut zip);
    let rels = parse_relationships(&mut zip);
    let embedded_fonts = parse_font_table(&mut zip);
    let footnotes = parse_footnotes(&mut zip, &styles, &theme);

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")
        .map_err(|_| Error::InvalidDocx("missing word/document.xml (is this a DOCX file?)".into()))?
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
                        n.tag_name().name() == "gridCol" && n.tag_name().namespace() == Some(WML_NS)
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
                        n.tag_name().name() == "tc" && n.tag_name().namespace() == Some(WML_NS)
                    }) {
                        let ci = grid_col;
                        let tc_pr = wml(tc, "tcPr");
                        let cell_width = tc_pr
                            .and_then(|pr| wml(pr, "tcW"))
                            .and_then(|w| twips_attr(w, "w"))
                            .unwrap_or_else(|| {
                                col_widths.get(ci).copied().unwrap_or(72.0)
                            });

                        let grid_span = tc_pr
                            .and_then(|pr| wml(pr, "gridSpan"))
                            .and_then(|n| n.attribute((WML_NS, "val")))
                            .and_then(|v| v.parse::<u16>().ok())
                            .unwrap_or(1);

                        let v_merge = tc_pr
                            .and_then(|pr| wml(pr, "vMerge"))
                            .map(|n| {
                                match n.attribute((WML_NS, "val")) {
                                    Some("restart") => VMerge::Restart,
                                    _ => VMerge::Continue,
                                }
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
                            bottom: if ri == num_rows - 1 { tb.bottom } else { tb.inside_h },
                            left: if ci == 0 { tb.left } else { tb.inside_v },
                            right: if span_end >= num_cols { tb.right } else { tb.inside_v },
                        });

                        let borders = tc_pr
                            .and_then(|pr| wml(pr, "tcBorders"))
                            .map(|bdr| {
                                let fallback = style_borders.unwrap_or_default();
                                let top = parse_cell_border(bdr, "top");
                                let bottom = parse_cell_border(bdr, "bottom");
                                let left = parse_cell_border(bdr, "left");
                                let left = if left.present { left } else { parse_cell_border(bdr, "start") };
                                let right = parse_cell_border(bdr, "right");
                                let right = if right.present { right } else { parse_cell_border(bdr, "end") };
                                CellBorders {
                                    top: if top.present { top } else { fallback.top },
                                    bottom: if bottom.present { bottom } else { fallback.bottom },
                                    left: if left.present { left } else { fallback.left },
                                    right: if right.present { right } else { fallback.right },
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
                            n.tag_name().name() == "p" && n.tag_name().namespace() == Some(WML_NS)
                        }) {
                            let parsed = parse_runs(p, &styles, &theme, &rels, &mut zip);
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
                            cell_paras.push(Paragraph {
                                runs: parsed.runs,
                                space_before: 0.0,
                                space_after: 0.0,
                                content_height: 0.0,
                                alignment,
                                indent_left: 0.0,
                                indent_right: 0.0,
                                indent_hanging: 0.0,
                                indent_first_line: 0.0,
                                list_label: String::new(),
                                contextual_spacing: false,
                                keep_next: false,
                                line_spacing,
                                image: None,
                                borders: ParagraphBorders::default(),
                                shading: None,
                                page_break_before: false,
                                column_break_before: false,
                                tab_stops: vec![],
                                extra_line_breaks: parsed.line_break_count,
                                floating_images: vec![],
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
                }));
            }
            "p" => {
                let ppr = wml(node, "pPr");

                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"))
                    .unwrap_or("Normal");

                let para_style = styles.paragraph_styles.get(para_style_id);

                let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));

                let inline_borders = ppr.map(parse_paragraph_borders).unwrap_or_default();
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

                let line_spacing = inline_spacing
                    .and_then(|n| {
                        n.attribute((WML_NS, "line"))
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|line_val| parse_line_spacing(n, line_val))
                    })
                    .or_else(|| para_style.and_then(|s| s.line_spacing));

                let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                let (mut indent_left, mut indent_hanging, list_label) =
                    parse_list_info(num_pr, &numbering, &mut counters, &mut last_seen_level);

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

                let parsed = parse_runs(node, &styles, &theme, &rels, &mut zip);
                let mut runs = parsed.runs;

                // Override font defaults from style for runs that used doc defaults
                for run in &mut runs {
                    if run.color.is_none() && style_color.is_some() {
                        run.color = style_color;
                    }
                }

                let tab_stops = ppr.map(parse_tab_stops).unwrap_or_default();

                // Determine if this is an image-only paragraph or mixed text+image
                let has_text = runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
                let has_inline_images = runs.iter().any(|r| r.inline_image.is_some());

                let mut floating_images = parsed.floating_images;

                let (para_image, content_height) = if has_inline_images && !has_text {
                    // Image-only paragraph: extract image for block-level rendering
                    let img_run_idx = runs.iter().position(|r| r.inline_image.is_some());
                    let img = img_run_idx.and_then(|i| runs[i].inline_image.take());
                    let h = img.as_ref().map(|i| i.display_height).unwrap_or(0.0);
                    (img, h)
                } else if has_inline_images {
                    // Mixed text+image: images stay in runs, no paragraph-level image
                    (None, 0.0)
                } else {
                    let drawing = compute_drawing_info(node, &rels, &mut zip);
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
                    contextual_spacing,
                    keep_next,
                    line_spacing,
                    image: para_image,
                    borders,
                    shading: para_shading,
                    page_break_before: parsed.has_page_break,
                    column_break_before: parsed.has_column_break,
                    tab_stops,
                    extra_line_breaks: parsed.line_break_count,
                    floating_images,
                }));

                // Mid-document section break: sectPr inside pPr ends the current section
                if let Some(sect_node) = ppr.and_then(|ppr| wml(ppr, "sectPr")) {
                    let props = parse_section_properties(
                        sect_node, &rels, &styles, &theme, &mut zip, default_line_pitch,
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
        parse_section_properties(
            sect_node, &rels, &styles, &theme, &mut zip, default_line_pitch,
        )
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

fn to_roman(mut n: u32) -> String {
    const TABLE: &[(u32, &str)] = &[
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut result = String::new();
    for &(value, numeral) in TABLE {
        while n >= value {
            result.push_str(numeral);
            n -= value;
        }
    }
    result
}

fn format_number(value: u32, num_fmt: &str) -> String {
    match num_fmt {
        "decimal" => value.to_string(),
        "decimalZero" => format!("{value:02}"),
        "lowerLetter" => {
            if value == 0 {
                return String::new();
            }
            let mut n = value - 1;
            let mut result = String::new();
            loop {
                result.insert(0, (b'a' + (n % 26) as u8) as char);
                if n < 26 {
                    break;
                }
                n = n / 26 - 1;
            }
            result
        }
        "upperLetter" => {
            if value == 0 {
                return String::new();
            }
            let mut n = value - 1;
            let mut result = String::new();
            loop {
                result.insert(0, (b'A' + (n % 26) as u8) as char);
                if n < 26 {
                    break;
                }
                n = n / 26 - 1;
            }
            result
        }
        "lowerRoman" => to_roman(value),
        "upperRoman" => to_roman(value).to_uppercase(),
        "none" => String::new(),
        _ => value.to_string(),
    }
}

fn normalize_bullet_text(text: &str) -> String {
    text.chars()
        .map(|c| {
            let cp = c as u32;
            if (0xF000..=0xF0FF).contains(&cp) {
                symbol_pua_to_unicode(cp).unwrap_or(c)
            } else {
                c
            }
        })
        .collect()
}

fn symbol_pua_to_unicode(cp: u32) -> Option<char> {
    let sym = cp - 0xF000;
    let mapped = match sym {
        0xB7 => '\u{2022}', // bullet •
        0xA7 => '\u{25A0}', // black square ■ (Wingdings §)
        0xA8 => '\u{25CB}', // white circle ○
        0xD8 => '\u{2666}', // diamond ◆
        0x76 => '\u{221A}', // check mark √
        _ => return char::from_u32(sym),
    };
    Some(mapped)
}

fn parse_list_info(
    num_pr: Option<roxmltree::Node>,
    numbering: &NumberingInfo,
    counters: &mut HashMap<(String, u8), u32>,
    last_seen_level: &mut HashMap<String, u8>,
) -> (f32, f32, String) {
    let Some(num_pr) = num_pr else {
        return (0.0, 0.0, String::new());
    };
    let Some(num_id) = wml_attr(num_pr, "numId") else {
        return (0.0, 0.0, String::new());
    };
    if num_id == "0" {
        return (0.0, 0.0, String::new());
    }
    let ilvl = wml_attr(num_pr, "ilvl")
        .and_then(|v| v.parse::<u8>().ok())
        .unwrap_or(0);

    let Some(abs_id) = numbering.num_to_abstract.get(num_id) else {
        return (0.0, 0.0, String::new());
    };
    let Some(levels) = numbering.abstract_nums.get(abs_id.as_str()) else {
        return (0.0, 0.0, String::new());
    };
    let Some(def) = levels.get(&ilvl) else {
        return (0.0, 0.0, String::new());
    };

    // Reset deeper-level counters when returning to a higher level
    let prev_level = last_seen_level.get(num_id).copied();
    if let Some(prev) = prev_level {
        if ilvl <= prev {
            for deeper in (ilvl + 1)..=prev {
                counters.remove(&(num_id.to_string(), deeper));
            }
        }
    }
    last_seen_level.insert(num_id.to_string(), ilvl);

    // Increment or initialize counter using the level's start value
    let start = def.start;
    let current_counter = *counters
        .entry((num_id.to_string(), ilvl))
        .and_modify(|c| *c += 1)
        .or_insert(start);

    let label = if def.num_fmt == "bullet" {
        let text = normalize_bullet_text(&def.lvl_text);
        if text.is_empty() { "\u{2022}".to_string() } else { text }
    } else {
        let mut label = def.lvl_text.clone();
        for lvl_idx in 0..9u8 {
            let placeholder = format!("%{}", lvl_idx + 1);
            if label.contains(&placeholder) {
                let lvl_counter = if lvl_idx == ilvl {
                    current_counter
                } else {
                    counters
                        .get(&(num_id.to_string(), lvl_idx))
                        .copied()
                        .unwrap_or(
                            levels
                                .get(&lvl_idx)
                                .map(|d| d.start)
                                .unwrap_or(1),
                        )
                };
                let lvl_fmt = levels
                    .get(&lvl_idx)
                    .map(|d| d.num_fmt.as_str())
                    .unwrap_or("decimal");
                label = label.replace(&placeholder, &format_number(lvl_counter, lvl_fmt));
            }
        }
        label
    };
    (def.indent_left, def.indent_hanging, label)
}

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn parse_rels_xml(xml_content: &str) -> HashMap<String, String> {
    let mut rels = HashMap::new();
    let Ok(xml) = roxmltree::Document::parse(xml_content) else {
        return rels;
    };
    for node in xml.root_element().children() {
        if node.tag_name().name() == "Relationship"
            && let (Some(id), Some(target)) = (node.attribute("Id"), node.attribute("Target"))
        {
            rels.insert(id.to_string(), target.to_string());
        }
    }
    rels
}

fn parse_relationships(zip: &mut zip::ZipArchive<std::fs::File>) -> HashMap<String, String> {
    let Some(xml_content) = read_zip_text(zip, "word/_rels/document.xml.rels") else {
        return HashMap::new();
    };
    parse_rels_xml(&xml_content)
}

/// Load relationships for a part like "word/header1.xml" → "word/_rels/header1.xml.rels"
fn parse_part_relationships(
    zip: &mut zip::ZipArchive<std::fs::File>,
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

fn image_dimensions(data: &[u8]) -> Option<(u32, u32, ImageFormat)> {
    // JPEG: starts with FF D8
    if data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8 {
        let mut i = 2;
        while i + 4 < data.len() {
            if data[i] != 0xFF {
                return None;
            }
            let marker = data[i + 1];
            if marker == 0xD9 {
                break;
            }
            let len = u16::from_be_bytes([data[i + 2], data[i + 3]]) as usize;
            if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2) && i + 9 < data.len() {
                let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
                let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
                return Some((width, height, ImageFormat::Jpeg));
            }
            i += 2 + len;
        }
        return None;
    }

    // PNG: starts with 89 50 4E 47, dimensions in IHDR chunk at bytes 16-23
    if data.len() >= 24 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
    {
        let width = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
        let height = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
        return Some((width, height, ImageFormat::Png));
    }

    None
}

enum RunDrawingResult {
    Inline(EmbeddedImage),
    Floating(FloatingImage),
}

fn parse_run_drawing(
    drawing_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> Option<RunDrawingResult> {
    for container in drawing_node.children() {
        let name = container.tag_name().name();
        if name != "inline" && name != "anchor" {
            continue;
        }
        if container.tag_name().namespace() != Some(WPD_NS) {
            continue;
        }

        let extent = container.children().find(|n| {
            n.tag_name().name() == "extent" && n.tag_name().namespace() == Some(WPD_NS)
        });
        let cx = extent
            .and_then(|n| n.attribute("cx"))
            .and_then(|v| v.parse::<f32>().ok())
            .unwrap_or(0.0);
        let cy = extent
            .and_then(|n| n.attribute("cy"))
            .and_then(|v| v.parse::<f32>().ok())
            .unwrap_or(0.0);
        let display_w = cx / 12700.0;
        let display_h = cy / 12700.0;

        if name == "anchor" {
            let has_wrap_none = container.children().any(|n| {
                n.tag_name().name() == "wrapNone"
                    && n.tag_name().namespace() == Some(WPD_NS)
            });
            if has_wrap_none {
                if let Some(embed_id) = find_blip_embed(container) {
                    if let Some(img) = read_image_from_zip(embed_id, rels, zip, display_w, display_h) {
                        let (h_position, h_relative, v_offset, v_relative, behind_doc) =
                            parse_anchor_position(container);
                        return Some(RunDrawingResult::Floating(FloatingImage {
                            image: img,
                            h_position,
                            h_relative_from: h_relative,
                            v_offset_pt: v_offset,
                            v_relative_from: v_relative,
                            behind_doc,
                        }));
                    }
                }
                continue;
            }
        }

        if let Some(embed_id) = find_blip_embed(container) {
            if let Some(img) = read_image_from_zip(embed_id, rels, zip, display_w, display_h) {
                return Some(RunDrawingResult::Inline(img));
            }
        }
    }
    None
}

fn find_blip_embed<'a>(container: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    container
        .descendants()
        .find(|n| n.tag_name().name() == "blip" && n.tag_name().namespace() == Some(DML_NS))
        .and_then(|n| n.attribute((REL_NS, "embed")))
}

struct DrawingInfo {
    height: f32,
    image: Option<EmbeddedImage>,
    floating_images: Vec<FloatingImage>,
}

fn parse_anchor_position(container: roxmltree::Node) -> (HorizontalPosition, &'static str, f32, &'static str, bool) {
    let behind_doc = container.attribute("behindDoc") == Some("1");

    let pos_h = container.children().find(|n| {
        n.tag_name().name() == "positionH" && n.tag_name().namespace() == Some(WPD_NS)
    });
    let h_relative = match pos_h.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => "page",
        Some("margin") => "margin",
        _ => "column",
    };
    let h_position = if let Some(align_node) = pos_h.and_then(|n| n.children().find(|c| c.tag_name().name() == "align")) {
        match align_node.text().unwrap_or("") {
            "center" => HorizontalPosition::AlignCenter,
            "right" => HorizontalPosition::AlignRight,
            _ => HorizontalPosition::AlignLeft,
        }
    } else if let Some(offset_node) = pos_h.and_then(|n| n.children().find(|c| c.tag_name().name() == "posOffset")) {
        let emu = offset_node.text().unwrap_or("0").parse::<f32>().unwrap_or(0.0);
        HorizontalPosition::Offset(emu / 12700.0)
    } else {
        HorizontalPosition::AlignLeft
    };

    let pos_v = container.children().find(|n| {
        n.tag_name().name() == "positionV" && n.tag_name().namespace() == Some(WPD_NS)
    });
    let v_relative = match pos_v.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => "page",
        Some("margin") => "margin",
        Some("topMargin") => "topMargin",
        _ => "paragraph",
    };
    let v_offset = if let Some(offset_node) = pos_v.and_then(|n| n.children().find(|c| c.tag_name().name() == "posOffset")) {
        offset_node.text().unwrap_or("0").parse::<f32>().unwrap_or(0.0) / 12700.0
    } else {
        0.0
    };

    (h_position, h_relative, v_offset, v_relative, behind_doc)
}

fn read_image_from_zip(
    embed_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<std::fs::File>,
    display_w: f32,
    display_h: f32,
) -> Option<EmbeddedImage> {
    let target = rels.get(embed_id)?;
    let zip_path = target
        .strip_prefix('/')
        .map(String::from)
        .unwrap_or_else(|| format!("word/{}", target));
    let mut entry = zip.by_name(&zip_path).ok()?;
    let mut data = Vec::new();
    entry.read_to_end(&mut data).ok()?;
    let (pw, ph, fmt) = image_dimensions(&data)?;
    Some(EmbeddedImage {
        data,
        format: fmt,
        pixel_width: pw,
        pixel_height: ph,
        display_width: display_w,
        display_height: display_h,
    })
}

fn compute_drawing_info(
    para_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> DrawingInfo {
    let mut max_height: f32 = 0.0;
    let mut image: Option<EmbeddedImage> = None;
    let mut floating_images: Vec<FloatingImage> = Vec::new();

    for child in para_node.children() {
        let is_wml = child.tag_name().namespace() == Some(WML_NS);
        let drawing_node = match child.tag_name().name() {
            "drawing" if is_wml => Some(child),
            "r" if is_wml => wml(child, "drawing"),
            _ => None,
        };

        let Some(drawing) = drawing_node else {
            continue;
        };
        for container in drawing.children() {
            let name = container.tag_name().name();
            if name != "inline" && name != "anchor" {
                continue;
            }
            if container.tag_name().namespace() != Some(WPD_NS) {
                continue;
            }

            let extent = container.children().find(|n| {
                n.tag_name().name() == "extent" && n.tag_name().namespace() == Some(WPD_NS)
            });
            let cx = extent
                .and_then(|n| n.attribute("cx"))
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0);
            let cy = extent
                .and_then(|n| n.attribute("cy"))
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0);
            let display_w = cx / 12700.0;
            let display_h = cy / 12700.0;

            // Anchored images with wrapNone float independently — they don't
            // affect paragraph layout height (text flows as if they're absent).
            if name == "anchor" {
                let has_wrap_none = container.children().any(|n| {
                    n.tag_name().name() == "wrapNone"
                        && n.tag_name().namespace() == Some(WPD_NS)
                });
                if has_wrap_none {
                    if let Some(embed_id) = find_blip_embed(container) {
                        if let Some(img) = read_image_from_zip(embed_id, rels, zip, display_w, display_h) {
                            let (h_position, h_relative, v_offset, v_relative, behind_doc) =
                                parse_anchor_position(container);
                            floating_images.push(FloatingImage {
                                image: img,
                                h_position,
                                h_relative_from: h_relative,
                                v_offset_pt: v_offset,
                                v_relative_from: v_relative,
                                behind_doc,
                            });
                        }
                    }
                    continue;
                }
            }

            max_height = max_height.max(display_h);

            if image.is_none() {
                if let Some(embed_id) = find_blip_embed(container) {
                    image = read_image_from_zip(embed_id, rels, zip, display_w, display_h);
                }
            }
        }
    }
    DrawingInfo {
        height: max_height,
        image,
        floating_images,
    }
}
