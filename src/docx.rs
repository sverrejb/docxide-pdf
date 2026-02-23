use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

use crate::error::Error;
use crate::model::{
    Alignment, Block, CellBorder, CellBorders, CellMargins, CellVAlign, Document, EmbeddedImage,
    FieldCode, HeaderFooter, ImageFormat, Paragraph, Run, TabAlignment, TabStop, Table, TableCell,
    TableRow, VMerge, VertAlign,
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

/// Parse a WML boolean toggle element (e.g., w:b, w:i, w:strike).
/// Present with no val or val != "0"/"false" means true.
fn wml_bool(parent: roxmltree::Node, name: &str) -> Option<bool> {
    wml(parent, name).map(|n| {
        n.attribute((WML_NS, "val"))
            .is_none_or(|v| v != "0" && v != "false")
    })
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
    space_before: Option<f32>,
    space_after: Option<f32>,
    alignment: Option<Alignment>,
    contextual_spacing: bool,
    keep_next: bool,
    line_spacing: Option<f32>, // auto line spacing factor override
    border_bottom_extra: f32,
    border_bottom: Option<crate::model::BorderBottom>,
    based_on: Option<String>,
}

struct CharacterStyle {
    font_size: Option<f32>,
    font_name: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strikethrough: Option<bool>,
    color: Option<[u8; 3]>,
}

struct TableBordersDef {
    top: CellBorder,
    bottom: CellBorder,
    left: CellBorder,
    right: CellBorder,
    inside_h: CellBorder,
    inside_v: CellBorder,
}

struct StylesInfo {
    defaults: StyleDefaults,
    paragraph_styles: HashMap<String, ParagraphStyle>,
    character_styles: HashMap<String, CharacterStyle>,
    table_border_styles: HashMap<String, TableBordersDef>,
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

    let names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
    let theme_name = names
        .iter()
        .find(|n| n.starts_with("word/theme/") && n.ends_with(".xml"));
    let Some(xml_content) = theme_name.and_then(|name| read_zip_text(zip, name)) else {
        return ThemeFonts { major, minor };
    };
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
        space_after: 0.0,
        line_spacing: 1.0,
    };
    let mut paragraph_styles = HashMap::new();
    let mut character_styles = HashMap::new();

    let Some(xml_content) = read_zip_text(zip, "word/styles.xml") else {
        return StylesInfo {
            defaults,
            paragraph_styles,
            character_styles,
            table_border_styles: HashMap::new(),
        };
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return StylesInfo {
            defaults,
            paragraph_styles,
            character_styles,
            table_border_styles: HashMap::new(),
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
        let space_before = spacing.and_then(|n| twips_attr(n, "before"));
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

        let bold = rpr.and_then(|n| wml_bool(n, "b"));
        let italic = rpr.and_then(|n| wml_bool(n, "i"));

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

    // Parse character styles (e.g., "Hyperlink")
    for style_node in root.children() {
        if style_node.tag_name().name() != "style"
            || style_node.tag_name().namespace() != Some(WML_NS)
        {
            continue;
        }
        if style_node.attribute((WML_NS, "type")) != Some("character") {
            continue;
        }
        let Some(style_id) = style_node.attribute((WML_NS, "styleId")) else {
            continue;
        };
        let Some(rpr) = wml(style_node, "rPr") else {
            continue;
        };
        let font_size = wml_attr(rpr, "sz")
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);
        let font_name = wml(rpr, "rFonts")
            .map(|rfonts| resolve_font_from_node(rfonts, theme, &defaults.font_name));
        let bold = wml_bool(rpr, "b");
        let italic = wml_bool(rpr, "i");
        let underline = wml(rpr, "u")
            .and_then(|n| n.attribute((WML_NS, "val")))
            .map(|v| v != "none");
        let strikethrough = wml_bool(rpr, "strike");
        let color = wml_attr(rpr, "color").and_then(parse_hex_color);

        character_styles.insert(
            style_id.to_string(),
            CharacterStyle {
                font_size,
                font_name,
                bold,
                italic,
                underline,
                strikethrough,
                color,
            },
        );
    }

    let mut table_border_styles = HashMap::new();
    for style_node in root.children() {
        if style_node.tag_name().name() != "style"
            || style_node.tag_name().namespace() != Some(WML_NS)
        {
            continue;
        }
        if style_node.attribute((WML_NS, "type")) != Some("table") {
            continue;
        }
        let Some(style_id) = style_node.attribute((WML_NS, "styleId")) else {
            continue;
        };
        if let Some(tbl_borders) =
            wml(style_node, "tblPr").and_then(|pr| wml(pr, "tblBorders"))
        {
            let parse_bdr = |name: &str| -> CellBorder {
                let Some(n) = wml(tbl_borders, name) else {
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
            let left = parse_bdr("left");
            let left = if left.present { left } else { parse_bdr("start") };
            let right = parse_bdr("right");
            let right = if right.present { right } else { parse_bdr("end") };
            table_border_styles.insert(
                style_id.to_string(),
                TableBordersDef {
                    top: parse_bdr("top"),
                    bottom: parse_bdr("bottom"),
                    left,
                    right,
                    inside_h: parse_bdr("insideH"),
                    inside_v: parse_bdr("insideV"),
                },
            );
        }
    }

    StylesInfo {
        defaults,
        paragraph_styles,
        character_styles,
        table_border_styles,
    }
}

fn resolve_based_on(styles: &mut HashMap<String, ParagraphStyle>) {
    let ids: Vec<String> = styles.keys().cloned().collect();
    for id in ids {
        let mut chain: Vec<String> = Vec::new();
        let mut current = id.clone();
        loop {
            if chain.contains(&current) {
                break;
            }
            chain.push(current.clone());
            match styles.get(&current).and_then(|s| s.based_on.clone()) {
                Some(parent) => current = parent,
                None => break,
            }
        }

        // Walk ancestors from furthest to closest, accumulating inherited values.
        // Each closer ancestor overrides the further one.
        macro_rules! inherit {
            ($field:ident, $inherited:expr, $s:expr) => {
                if $s.$field.is_some() {
                    $inherited = $s.$field.clone();
                }
            };
        }

        let mut inh = ParagraphStyle {
            font_size: None,
            font_name: None,
            bold: None,
            italic: None,
            color: None,
            space_before: None,
            space_after: None,
            alignment: None,
            contextual_spacing: false,
            keep_next: false,
            line_spacing: None,
            border_bottom_extra: 0.0,
            border_bottom: None,
            based_on: None,
        };

        for ancestor_id in chain.iter().rev() {
            if let Some(s) = styles.get(ancestor_id) {
                inherit!(font_name, inh.font_name, s);
                inherit!(font_size, inh.font_size, s);
                inherit!(bold, inh.bold, s);
                inherit!(italic, inh.italic, s);
                inherit!(color, inh.color, s);
                inherit!(alignment, inh.alignment, s);
                inherit!(space_before, inh.space_before, s);
                inherit!(space_after, inh.space_after, s);
                inherit!(line_spacing, inh.line_spacing, s);
            }
        }

        if let Some(s) = styles.get_mut(&id) {
            s.font_name = s.font_name.take().or(inh.font_name);
            s.font_size = s.font_size.or(inh.font_size);
            s.bold = s.bold.or(inh.bold);
            s.italic = s.italic.or(inh.italic);
            s.color = s.color.or(inh.color);
            s.alignment = s.alignment.or(inh.alignment);
            s.space_before = s.space_before.or(inh.space_before);
            s.space_after = s.space_after.or(inh.space_after);
            s.line_spacing = s.line_spacing.or(inh.line_spacing);
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

struct ParsedRuns {
    runs: Vec<Run>,
    has_page_break: bool,
    line_break_count: u32,
}

fn parse_runs(
    para_node: roxmltree::Node,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
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
    let style_color: Option<[u8; 3]> = para_style.and_then(|s| s.color);

    let run_nodes: Vec<(roxmltree::Node, Option<String>)> = para_node
        .children()
        .flat_map(|child| {
            let name = child.tag_name().name();
            let is_wml = child.tag_name().namespace() == Some(WML_NS);
            if is_wml && name == "r" {
                vec![(child, None)]
            } else if is_wml && name == "hyperlink" {
                let url = child
                    .attribute((REL_NS, "id"))
                    .and_then(|rid| rels.get(rid))
                    .cloned();
                child
                    .children()
                    .filter(|n| {
                        n.tag_name().name() == "r" && n.tag_name().namespace() == Some(WML_NS)
                    })
                    .map(move |n| (n, url.clone()))
                    .collect()
            } else {
                vec![]
            }
        })
        .collect();

    let mut runs = Vec::new();
    let mut has_page_break = false;
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

        let color = rpr
            .and_then(|n| wml_attr(n, "color"))
            .and_then(parse_hex_color)
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
                                    color,
                                    is_tab: false,
                                    vertical_align,
                                    field_code: None,
                                    hyperlink_url: hyperlink_url.clone(),
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
                                        color,
                                        is_tab: false,
                                        vertical_align: VertAlign::Baseline,
                                        field_code: Some(code),
                                        hyperlink_url: hyperlink_url.clone(),
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
                        pending_text.push_str(t);
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
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
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
                        color: None,
                        is_tab: true,
                        vertical_align: VertAlign::Baseline,
                        field_code: None,
                        hyperlink_url: None,
                    });
                }
                "br" if !in_field => {
                    if child.attribute((WML_NS, "type")) == Some("page") {
                        has_page_break = true;
                    } else {
                        line_break_count += 1;
                    }
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
                color,
                is_tab: false,
                vertical_align,
                field_code: None,
                hyperlink_url: hyperlink_url.clone(),
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
                color: None,
                is_tab: false,
                vertical_align: VertAlign::Baseline,
                field_code: None,
                hyperlink_url: None,
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
            color: None,
            is_tab: false,
            vertical_align: VertAlign::Baseline,
            field_code: None,
            hyperlink_url: None,
        });
    }

    ParsedRuns {
        runs,
        has_page_break,
        line_break_count,
    }
}

fn parse_header_footer_xml(
    xml_content: &str,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
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

        let parsed = parse_runs(node, styles, theme, rels);

        paragraphs.push(Paragraph {
            runs: parsed.runs,
            space_before: 0.0,
            space_after: 0.0,
            content_height: 0.0,
            alignment,
            indent_left: 0.0,
            indent_hanging: 0.0,
            list_label: String::new(),
            contextual_spacing: false,
            keep_next: false,
            line_spacing: None,
            image: None,
            border_bottom: None,
            page_break_before: false,
            tab_stops: vec![],
            extra_line_breaks: parsed.line_break_count,
        });
    }

    if paragraphs.is_empty() {
        None
    } else {
        Some(HeaderFooter { paragraphs })
    }
}

fn read_zip_text(zip: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Option<String> {
    let mut content = String::new();
    zip.by_name(name).ok()?.read_to_string(&mut content).ok()?;
    Some(content)
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
    let header_margin = pg_mar.and_then(|n| twips_attr(n, "header")).unwrap_or(36.0);
    let footer_margin = pg_mar.and_then(|n| twips_attr(n, "footer")).unwrap_or(36.0);
    let line_pitch = doc_grid
        .and_then(|n| twips_attr(n, "linePitch"))
        .unwrap_or(styles.defaults.font_size * 1.2);

    let different_first_page = sect.and_then(|s| wml(s, "titlePg")).is_some();

    // Parse header/footer references from sectPr
    let mut header_default_rid = None;
    let mut header_first_rid = None;
    let mut footer_default_rid = None;
    let mut footer_first_rid = None;
    if let Some(sect) = sect {
        for child in sect.children() {
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
    }

    let resolve_hf =
        |rid: Option<&str>, zip: &mut zip::ZipArchive<std::fs::File>| -> Option<HeaderFooter> {
            let target = rels.get(rid?)?;
            let zip_path = target
                .strip_prefix('/')
                .map(String::from)
                .unwrap_or_else(|| format!("word/{}", target));
            let xml_text = read_zip_text(zip, &zip_path)?;
            parse_header_footer_xml(&xml_text, &styles, &theme, &HashMap::new())
        };

    let header_default = resolve_hf(header_default_rid, &mut zip);
    let header_first = resolve_hf(header_first_rid, &mut zip);
    let footer_default = resolve_hf(footer_default_rid, &mut zip);
    let footer_first = resolve_hf(footer_first_rid, &mut zip);

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

                let tbl_rows: Vec<_> = node
                    .children()
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
                    for tc in tr.children().filter(|n| {
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
                            let parsed = parse_runs(p, &styles, &theme, &rels);
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
                                    .and_then(|n| n.attribute((WML_NS, "line")))
                                    .and_then(|v| v.parse::<f32>().ok())
                                    .map(|val| val / 240.0)
                                    .or_else(|| para_style.and_then(|s| s.line_spacing))
                                    .unwrap_or(1.0),
                            );
                            cell_paras.push(Paragraph {
                                runs: parsed.runs,
                                space_before: 0.0,
                                space_after: 0.0,
                                content_height: 0.0,
                                alignment,
                                indent_left: 0.0,
                                indent_hanging: 0.0,
                                list_label: String::new(),
                                contextual_spacing: false,
                                keep_next: false,
                                line_spacing,
                                image: None,
                                border_bottom: None,
                                page_break_before: false,
                                tab_stops: vec![],
                                extra_line_breaks: parsed.line_break_count,
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

                let space_before = inline_spacing
                    .and_then(|n| twips_attr(n, "before"))
                    .or_else(|| para_style.and_then(|s| s.space_before))
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

                let parsed = parse_runs(node, &styles, &theme, &rels);
                let mut runs = parsed.runs;

                // Override font defaults from style for runs that used doc defaults
                for run in &mut runs {
                    if run.color.is_none() && style_color.is_some() {
                        run.color = style_color;
                    }
                }

                let tab_stops = ppr.map(parse_tab_stops).unwrap_or_default();
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
                    page_break_before: parsed.has_page_break,
                    tab_stops,
                    extra_line_breaks: parsed.line_break_count,
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
        header_default,
        header_first,
        footer_default,
        footer_first,
        header_margin,
        footer_margin,
        different_first_page,
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
    let Some(xml_content) = read_zip_text(zip, "word/_rels/document.xml.rels") else {
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
                            && let Some((pw, ph, fmt)) = image_dimensions(&data)
                        {
                            image = Some(EmbeddedImage {
                                data,
                                format: fmt,
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
