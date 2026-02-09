use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

use crate::error::Error;
use crate::model::{
    Alignment, Block, Document, EmbeddedImage, Paragraph, Run, Table, TableCell, TableRow,
};

struct LevelDef {
    num_fmt: String,
    lvl_text: String,
    indent_left: f32,
    indent_hanging: f32,
}

struct NumberingInfo {
    abstract_nums: HashMap<String, HashMap<u8, LevelDef>>,
    num_to_abstract: HashMap<String, String>,
}

const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPD_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

fn parse_hex_color(val: &str) -> Option<[u8; 3]> {
    if val == "auto" || val.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&val[0..2], 16).ok()?;
    let g = u8::from_str_radix(&val[2..4], 16).ok()?;
    let b = u8::from_str_radix(&val[4..6], 16).ok()?;
    Some([r, g, b])
}

fn wml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(WML_NS))
}

fn wml_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    wml(node, child).and_then(|n| n.attribute((WML_NS, "val")))
}

fn twips_attr(node: roxmltree::Node, attr: &str) -> Option<f32> {
    node.attribute((WML_NS, attr))
        .and_then(|v| v.parse::<f32>().ok())
        .map(twips_to_pts)
}

fn parse_border_bottom(ppr: roxmltree::Node) -> Option<crate::model::BorderBottom> {
    let bottom = wml(ppr, "pBdr").and_then(|pbdr| wml(pbdr, "bottom"))?;
    let val = bottom.attribute((WML_NS, "val")).unwrap_or("none");
    if val == "none" || val == "nil" {
        return None;
    }
    // sz is in 1/8 of a point
    let width_pt = bottom
        .attribute((WML_NS, "sz"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 8.0)
        .unwrap_or(0.5);
    let space_pt = bottom
        .attribute((WML_NS, "space"))
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0);
    let color = bottom
        .attribute((WML_NS, "color"))
        .and_then(parse_hex_color)
        .unwrap_or([0, 0, 0]);
    Some(crate::model::BorderBottom {
        width_pt,
        space_pt,
        color,
    })
}

fn border_bottom_extra(ppr: roxmltree::Node) -> f32 {
    parse_border_bottom(ppr)
        .map(|b| b.space_pt + b.width_pt)
        .unwrap_or(0.0)
}

fn dml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

fn latin_typeface<'a>(node: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    dml(node, "latin")
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

struct ThemeFonts {
    major: String,
    minor: String,
}

struct StyleDefaults {
    font_size: f32,
    font_name: String,
    space_after: f32,
    line_spacing: f32, // multiplier from w:spacing @line / 240
}

struct ParagraphStyle {
    font_size: Option<f32>,
    font_name: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<[u8; 3]>,
    space_before: f32,
    space_after: Option<f32>,
    alignment: Option<Alignment>,
    contextual_spacing: bool,
    keep_next: bool,
    line_spacing: Option<f32>, // auto line spacing factor override
    border_bottom_extra: f32,
    border_bottom: Option<crate::model::BorderBottom>,
    based_on: Option<String>,
}

struct StylesInfo {
    defaults: StyleDefaults,
    paragraph_styles: HashMap<String, ParagraphStyle>,
}

fn parse_alignment(val: &str) -> Alignment {
    match val {
        "center" => Alignment::Center,
        "right" | "end" => Alignment::Right,
        "both" => Alignment::Justify,
        _ => Alignment::Left,
    }
}

fn parse_theme(zip: &mut zip::ZipArchive<std::fs::File>) -> ThemeFonts {
    let mut major = String::from("Aptos Display");
    let mut minor = String::from("Aptos");

    let mut xml_content = String::new();
    let names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
    let Some(theme_name) = names
        .iter()
        .find(|n| n.starts_with("word/theme/") && n.ends_with(".xml"))
    else {
        return ThemeFonts { major, minor };
    };
    let theme_name = theme_name.clone();
    let Ok(mut file) = zip.by_name(&theme_name) else {
        return ThemeFonts { major, minor };
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return ThemeFonts { major, minor };
    }
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return ThemeFonts { major, minor };
    };

    for node in xml.descendants() {
        if node.tag_name().namespace() != Some(DML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "majorFont" => {
                if let Some(tf) = latin_typeface(node) {
                    major = tf.to_string();
                }
            }
            "minorFont" => {
                if let Some(tf) = latin_typeface(node) {
                    minor = tf.to_string();
                }
            }
            _ => {}
        }
    }

    ThemeFonts { major, minor }
}

fn resolve_font(
    ascii: Option<&str>,
    ascii_theme: Option<&str>,
    theme: &ThemeFonts,
    default_font: &str,
) -> String {
    if let Some(f) = ascii {
        return f.to_string();
    }
    match ascii_theme {
        Some("majorHAnsi") => theme.major.clone(),
        Some("minorHAnsi") => theme.minor.clone(),
        _ => default_font.to_string(),
    }
}

fn resolve_font_from_node(
    rfonts: roxmltree::Node,
    theme: &ThemeFonts,
    default_font: &str,
) -> String {
    resolve_font(
        rfonts.attribute((WML_NS, "ascii")),
        rfonts.attribute((WML_NS, "asciiTheme")),
        theme,
        default_font,
    )
}

fn parse_styles(zip: &mut zip::ZipArchive<std::fs::File>, theme: &ThemeFonts) -> StylesInfo {
    let mut defaults = StyleDefaults {
        font_size: 12.0,
        font_name: theme.minor.clone(),
        space_after: 8.0,
        line_spacing: 1.2,
    };
    let mut paragraph_styles = HashMap::new();

    let mut xml_content = String::new();
    let Ok(mut file) = zip.by_name("word/styles.xml") else {
        return StylesInfo {
            defaults,
            paragraph_styles,
        };
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return StylesInfo {
            defaults,
            paragraph_styles,
        };
    }
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return StylesInfo {
            defaults,
            paragraph_styles,
        };
    };

    let root = xml.root_element();

    if let Some(doc_defaults) = wml(root, "docDefaults") {
        if let Some(rpr) = wml(doc_defaults, "rPrDefault").and_then(|n| wml(n, "rPr")) {
            if let Some(sz_val) = wml_attr(rpr, "sz").and_then(|v| v.parse::<f32>().ok()) {
                defaults.font_size = sz_val / 2.0;
            }
            if let Some(rfonts) = wml(rpr, "rFonts") {
                defaults.font_name = resolve_font_from_node(rfonts, theme, &theme.minor);
            }
        }
        let default_spacing = wml(doc_defaults, "pPrDefault")
            .and_then(|n| wml(n, "pPr"))
            .and_then(|n| wml(n, "spacing"));
        if let Some(spacing) = default_spacing {
            if let Some(after_val) = twips_attr(spacing, "after") {
                defaults.space_after = after_val;
            }
            if let Some(line_val) = spacing
                .attribute((WML_NS, "line"))
                .and_then(|v| v.parse::<f32>().ok())
            {
                defaults.line_spacing = line_val / 240.0;
            }
        }
    }

    for style_node in root.children() {
        if style_node.tag_name().name() != "style"
            || style_node.tag_name().namespace() != Some(WML_NS)
        {
            continue;
        }
        if style_node.attribute((WML_NS, "type")) != Some("paragraph") {
            continue;
        }
        let Some(style_id) = style_node.attribute((WML_NS, "styleId")) else {
            continue;
        };

        let ppr = wml(style_node, "pPr");
        let spacing = ppr.and_then(|n| wml(n, "spacing"));
        let space_before = spacing.and_then(|n| twips_attr(n, "before")).unwrap_or(0.0);
        let space_after = spacing.and_then(|n| twips_attr(n, "after"));
        let bdr_extra = ppr.map(border_bottom_extra).unwrap_or(0.0);
        let border_bottom = ppr.and_then(parse_border_bottom);

        let rpr = wml(style_node, "rPr");

        let font_size = rpr
            .and_then(|n| wml_attr(n, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);

        let font_name = rpr
            .and_then(|n| wml(n, "rFonts"))
            .map(|rfonts| resolve_font_from_node(rfonts, theme, &defaults.font_name));

        let bold = rpr.and_then(|n| wml(n, "b")).map(|n| {
            n.attribute((WML_NS, "val"))
                .is_none_or(|v| v != "0" && v != "false")
        });
        let italic = rpr.and_then(|n| wml(n, "i")).map(|n| {
            n.attribute((WML_NS, "val"))
                .is_none_or(|v| v != "0" && v != "false")
        });

        let color = rpr
            .and_then(|n| wml_attr(n, "color"))
            .and_then(parse_hex_color);

        let alignment = ppr.and_then(|ppr| wml_attr(ppr, "jc")).map(parse_alignment);

        let contextual_spacing = ppr.and_then(|ppr| wml(ppr, "contextualSpacing")).is_some();

        let keep_next = ppr.and_then(|ppr| wml(ppr, "keepNext")).is_some();

        let line_spacing = spacing
            .and_then(|n| n.attribute((WML_NS, "line")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|val| val / 240.0);

        let based_on = wml(style_node, "basedOn")
            .and_then(|n| n.attribute((WML_NS, "val")))
            .map(|s| s.to_string());

        paragraph_styles.insert(
            style_id.to_string(),
            ParagraphStyle {
                font_size,
                font_name,
                bold,
                italic,
                color,
                space_before,
                space_after,
                alignment,
                contextual_spacing,
                keep_next,
                line_spacing,
                border_bottom_extra: bdr_extra,
                border_bottom,
                based_on,
            },
        );
    }

    resolve_based_on(&mut paragraph_styles);

    StylesInfo {
        defaults,
        paragraph_styles,
    }
}

fn resolve_based_on(styles: &mut HashMap<String, ParagraphStyle>) {
    let ids: Vec<String> = styles.keys().cloned().collect();
    for id in ids {
        let mut chain: Vec<String> = Vec::new();
        let mut current = id.clone();
        // Walk the basedOn chain collecting parent values
        loop {
            if chain.contains(&current) {
                break; // cycle
            }
            chain.push(current.clone());
            let next = styles.get(&current).and_then(|s| s.based_on.clone());
            match next {
                Some(parent) => current = parent,
                None => break,
            }
        }
        // Apply inheritance from furthest ancestor down to the style itself
        let mut inherited_font_name: Option<String> = None;
        let mut inherited_font_size: Option<f32> = None;
        let mut inherited_bold: Option<bool> = None;
        let mut inherited_italic: Option<bool> = None;
        let mut inherited_color: Option<[u8; 3]> = None;
        let mut inherited_alignment: Option<Alignment> = None;
        let mut inherited_space_after: Option<f32> = None;
        let mut inherited_line_spacing: Option<f32> = None;

        for ancestor_id in chain.iter().rev() {
            if let Some(s) = styles.get(ancestor_id) {
                if s.font_name.is_some() {
                    inherited_font_name = s.font_name.clone();
                }
                if s.font_size.is_some() {
                    inherited_font_size = s.font_size;
                }
                if s.bold.is_some() {
                    inherited_bold = s.bold;
                }
                if s.italic.is_some() {
                    inherited_italic = s.italic;
                }
                if s.color.is_some() {
                    inherited_color = s.color;
                }
                if s.alignment.is_some() {
                    inherited_alignment = s.alignment;
                }
                if s.space_after.is_some() {
                    inherited_space_after = s.space_after;
                }
                if s.line_spacing.is_some() {
                    inherited_line_spacing = s.line_spacing;
                }
            }
        }

        if let Some(s) = styles.get_mut(&id) {
            if s.font_name.is_none() {
                s.font_name = inherited_font_name;
            }
            if s.font_size.is_none() {
                s.font_size = inherited_font_size;
            }
            if s.bold.is_none() {
                s.bold = inherited_bold;
            }
            if s.italic.is_none() {
                s.italic = inherited_italic;
            }
            if s.color.is_none() {
                s.color = inherited_color;
            }
            if s.alignment.is_none() {
                s.alignment = inherited_alignment;
            }
            if s.space_after.is_none() {
                s.space_after = inherited_space_after;
            }
            if s.line_spacing.is_none() {
                s.line_spacing = inherited_line_spacing;
            }
        }
    }
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
    let mut xml_content = String::new();
    let Ok(mut file) = zip.by_name("word/_rels/fontTable.xml.rels") else {
        return rels;
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return rels;
    }
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

    // Phase 1: parse XML to collect embed info (drops borrow on zip before phase 2)
    let embeds = {
        let mut xml_content = String::new();
        let Ok(mut file) = zip.by_name("word/fontTable.xml") else {
            return result;
        };
        if file.read_to_string(&mut xml_content).is_err() {
            return result;
        }
        drop(file);

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
        result.insert((info.font_name.to_lowercase(), info.bold, info.italic), data);
    }

    result
}

fn parse_numbering(zip: &mut zip::ZipArchive<std::fs::File>) -> NumberingInfo {
    let mut abstract_nums: HashMap<String, HashMap<u8, LevelDef>> = HashMap::new();
    let mut num_to_abstract: HashMap<String, String> = HashMap::new();

    let mut xml_content = String::new();
    let Ok(mut file) = zip.by_name("word/numbering.xml") else {
        return NumberingInfo {
            abstract_nums,
            num_to_abstract,
        };
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return NumberingInfo {
            abstract_nums,
            num_to_abstract,
        };
    }
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

fn parse_runs(para_node: roxmltree::Node, styles: &StylesInfo, theme: &ThemeFonts) -> Vec<Run> {
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
    let style_color: Option<[u8; 3]> = para_style.and_then(|s| s.color);

    let run_nodes: Vec<_> = para_node
        .children()
        .flat_map(|child| {
            let name = child.tag_name().name();
            let is_wml = child.tag_name().namespace() == Some(WML_NS);
            if is_wml && name == "r" {
                vec![child]
            } else if is_wml && name == "hyperlink" {
                child
                    .children()
                    .filter(|n| {
                        n.tag_name().name() == "r" && n.tag_name().namespace() == Some(WML_NS)
                    })
                    .collect()
            } else {
                vec![]
            }
        })
        .collect();

    let mut runs = Vec::new();
    for run_node in run_nodes {
        let rpr = wml(run_node, "rPr");

        let font_size = rpr
            .and_then(|n| wml_attr(n, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0)
            .unwrap_or(style_font_size);

        let font_name = rpr
            .and_then(|n| wml(n, "rFonts"))
            .map(|rfonts| resolve_font_from_node(rfonts, theme, &style_font_name))
            .unwrap_or_else(|| style_font_name.clone());

        let bold = match rpr.and_then(|n| wml(n, "b")) {
            Some(n) => n
                .attribute((WML_NS, "val"))
                .is_none_or(|v| v != "0" && v != "false"),
            None => style_bold,
        };
        let italic = match rpr.and_then(|n| wml(n, "i")) {
            Some(n) => n
                .attribute((WML_NS, "val"))
                .is_none_or(|v| v != "0" && v != "false"),
            None => style_italic,
        };

        let color = rpr
            .and_then(|n| wml_attr(n, "color"))
            .and_then(parse_hex_color)
            .or(style_color);

        let text: String = run_node
            .children()
            .filter(|n| n.tag_name().name() == "t" && n.tag_name().namespace() == Some(WML_NS))
            .filter_map(|n| n.text())
            .collect();

        if !text.is_empty() {
            runs.push(Run {
                text,
                font_size,
                font_name,
                bold,
                italic,
                color,
            });
        }
    }

    // Empty paragraphs with explicit font sizing in their paragraph mark (pPr/rPr)
    // need a synthetic run so the renderer computes the correct line height.
    if runs.is_empty() {
        let mark_rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
        let has_explicit_sz = mark_rpr
            .and_then(|n| wml_attr(n, "sz"))
            .is_some();
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
                color: None,
            });
        }
    }

    runs
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

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")
        .map_err(|_| Error::InvalidDocx("missing word/document.xml (is this a DOCX file?)".into()))?
        .read_to_string(&mut xml_content)?;

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = wml(root, "body").ok_or_else(|| Error::Pdf("Missing w:body".into()))?;

    let sect = wml(body, "sectPr");
    let pg_sz = sect.and_then(|s| wml(s, "pgSz"));
    let pg_mar = sect.and_then(|s| wml(s, "pgMar"));
    let doc_grid = sect.and_then(|s| wml(s, "docGrid"));

    let page_width = pg_sz.and_then(|n| twips_attr(n, "w")).unwrap_or(612.0);
    let page_height = pg_sz.and_then(|n| twips_attr(n, "h")).unwrap_or(792.0);
    let margin_top = pg_mar.and_then(|n| twips_attr(n, "top")).unwrap_or(72.0);
    let margin_bottom = pg_mar.and_then(|n| twips_attr(n, "bottom")).unwrap_or(72.0);
    let margin_left = pg_mar.and_then(|n| twips_attr(n, "left")).unwrap_or(72.0);
    let margin_right = pg_mar.and_then(|n| twips_attr(n, "right")).unwrap_or(72.0);
    let line_pitch = doc_grid
        .and_then(|n| twips_attr(n, "linePitch"))
        .unwrap_or(styles.defaults.font_size * 1.2);

    let mut blocks = Vec::new();
    let mut counters: HashMap<(String, u8), u32> = HashMap::new();

    for node in body.children() {
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

                let mut rows = Vec::new();
                for tr in node.children().filter(|n| {
                    n.tag_name().name() == "tr" && n.tag_name().namespace() == Some(WML_NS)
                }) {
                    let mut cells = Vec::new();
                    for tc in tr.children().filter(|n| {
                        n.tag_name().name() == "tc" && n.tag_name().namespace() == Some(WML_NS)
                    }) {
                        let cell_width = wml(tc, "tcPr")
                            .and_then(|pr| wml(pr, "tcW"))
                            .and_then(|w| twips_attr(w, "w"))
                            .unwrap_or_else(|| {
                                col_widths.get(cells.len()).copied().unwrap_or(72.0)
                            });

                        let mut cell_paras = Vec::new();
                        for p in tc.children().filter(|n| {
                            n.tag_name().name() == "p" && n.tag_name().namespace() == Some(WML_NS)
                        }) {
                            let runs = parse_runs(p, &styles, &theme);
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
                            cell_paras.push(Paragraph {
                                runs,
                                space_before: 0.0,
                                space_after: 0.0,
                                content_height: 0.0,
                                alignment,
                                indent_left: 0.0,
                                indent_hanging: 0.0,
                                list_label: String::new(),
                                contextual_spacing: false,
                                keep_next: false,
                                line_spacing: Some(1.0),
                                image: None,
                                border_bottom: None,
                            });
                        }
                        cells.push(TableCell {
                            width: cell_width,
                            paragraphs: cell_paras,
                        });
                    }
                    rows.push(TableRow { cells });
                }
                blocks.push(Block::Table(Table { col_widths, rows }));
            }
            "p" => {
                let ppr = wml(node, "pPr");

                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"))
                    .unwrap_or("Normal");

                let para_style = styles.paragraph_styles.get(para_style_id);

                let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));

                let space_before = inline_spacing
                    .and_then(|n| twips_attr(n, "before"))
                    .or_else(|| para_style.map(|s| s.space_before))
                    .unwrap_or(0.0);

                let inline_bdr = ppr.and_then(parse_border_bottom);
                let inline_bdr_extra = inline_bdr
                    .as_ref()
                    .map(|b| b.space_pt + b.width_pt)
                    .unwrap_or(0.0);
                let (bdr_extra, border_bottom) = if inline_bdr.is_some() {
                    (inline_bdr_extra, inline_bdr)
                } else {
                    (
                        para_style.map(|s| s.border_bottom_extra).unwrap_or(0.0),
                        para_style.and_then(|s| s.border_bottom.clone()),
                    )
                };
                let space_after = inline_spacing
                    .and_then(|n| twips_attr(n, "after"))
                    .or_else(|| para_style.and_then(|s| s.space_after))
                    .unwrap_or(styles.defaults.space_after)
                    + bdr_extra;

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
                    .and_then(|n| n.attribute((WML_NS, "line")))
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(|val| val / 240.0)
                    .or_else(|| para_style.and_then(|s| s.line_spacing));

                let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                let (mut indent_left, mut indent_hanging, list_label) =
                    parse_list_info(num_pr, &numbering, &mut counters);

                if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                    if let Some(v) = twips_attr(ind, "left") {
                        indent_left = v;
                    }
                    if let Some(v) = twips_attr(ind, "hanging") {
                        indent_hanging = v;
                    }
                }

                let mut runs = parse_runs(node, &styles, &theme);

                // Override font defaults from style for runs that used doc defaults
                for run in &mut runs {
                    if run.color.is_none() && style_color.is_some() {
                        run.color = style_color;
                    }
                }

                let drawing = compute_drawing_info(node, &rels, &mut zip);

                blocks.push(Block::Paragraph(Paragraph {
                    runs,
                    space_before,
                    space_after,
                    content_height: drawing.height,
                    alignment,
                    indent_left,
                    indent_hanging,
                    list_label,
                    contextual_spacing,
                    keep_next,
                    line_spacing,
                    image: drawing.image,
                    border_bottom,
                }));
            }
            _ => {}
        }
    }

    Ok(Document {
        page_width,
        page_height,
        margin_top,
        margin_bottom,
        margin_left,
        margin_right,
        line_pitch,
        line_spacing: styles.defaults.line_spacing,
        blocks,
        embedded_fonts,
    })
}

fn parse_list_info(
    num_pr: Option<roxmltree::Node>,
    numbering: &NumberingInfo,
    counters: &mut HashMap<(String, u8), u32>,
) -> (f32, f32, String) {
    let Some(num_pr) = num_pr else {
        return (0.0, 0.0, String::new());
    };
    let Some(num_id) = wml_attr(num_pr, "numId") else {
        return (0.0, 0.0, String::new());
    };
    let ilvl = wml_attr(num_pr, "ilvl")
        .and_then(|v| v.parse::<u8>().ok())
        .unwrap_or(0);

    let Some(def) = numbering
        .num_to_abstract
        .get(num_id)
        .and_then(|abs_id| numbering.abstract_nums.get(abs_id))
        .and_then(|levels| levels.get(&ilvl))
    else {
        return (0.0, 0.0, String::new());
    };

    let counter = counters
        .entry((num_id.to_string(), ilvl))
        .and_modify(|c| *c += 1)
        .or_insert(1);
    let label = if def.num_fmt == "bullet" {
        "\u{2022}".to_string()
    } else {
        def.lvl_text
            .replace(&format!("%{}", ilvl + 1), &counter.to_string())
    };
    (def.indent_left, def.indent_hanging, label)
}

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn parse_relationships(zip: &mut zip::ZipArchive<std::fs::File>) -> HashMap<String, String> {
    let mut rels = HashMap::new();
    let mut xml_content = String::new();
    let Ok(mut file) = zip.by_name("word/_rels/document.xml.rels") else {
        return rels;
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return rels;
    }
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

fn jpeg_dimensions(data: &[u8]) -> Option<(u32, u32)> {
    if data.len() < 2 || data[0] != 0xFF || data[1] != 0xD8 {
        return None;
    }
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
        // SOF0, SOF1, SOF2 markers contain dimensions
        if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2) && i + 9 < data.len() {
            let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
            let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
            return Some((width, height));
        }
        i += 2 + len;
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
}

fn compute_drawing_info(
    para_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<std::fs::File>,
) -> DrawingInfo {
    let mut max_height: f32 = 0.0;
    let mut image: Option<EmbeddedImage> = None;

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
            if (name == "inline" || name == "anchor")
                && container.tag_name().namespace() == Some(WPD_NS)
            {
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
                max_height = max_height.max(display_h);

                if image.is_none()
                    && let Some(embed_id) = find_blip_embed(container)
                    && let Some(target) = rels.get(embed_id)
                {
                    let zip_path = target
                        .strip_prefix('/')
                        .map(String::from)
                        .unwrap_or_else(|| format!("word/{}", target));
                    if let Ok(mut entry) = zip.by_name(&zip_path) {
                        let mut data = Vec::new();
                        if entry.read_to_end(&mut data).is_ok()
                            && let Some((pw, ph)) = jpeg_dimensions(&data)
                        {
                            image = Some(EmbeddedImage {
                                data,
                                pixel_width: pw,
                                pixel_height: ph,
                                display_width: display_w,
                                display_height: display_h,
                            });
                        }
                    }
                }
            }
        }
    }
    DrawingInfo {
        height: max_height,
        image,
    }
}
