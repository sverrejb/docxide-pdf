use std::collections::HashMap;

use crate::model::{Alignment, CellBorder};

use super::{
    DML_NS, WML_NS, border_bottom_extra, parse_border_bottom, parse_hex_color, read_zip_text,
    twips_attr, wml, wml_attr, wml_bool,
};

fn dml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

fn latin_typeface<'a>(node: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    dml(node, "latin")
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

pub(super) struct ThemeFonts {
    pub(super) major: String,
    pub(super) minor: String,
}

pub(super) struct StyleDefaults {
    pub(super) font_size: f32,
    pub(super) font_name: String,
    pub(super) space_after: f32,
    pub(super) line_spacing: f32, // multiplier from w:spacing @line / 240
}

pub(super) struct ParagraphStyle {
    pub(super) font_size: Option<f32>,
    pub(super) font_name: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
    pub(super) space_before: Option<f32>,
    pub(super) space_after: Option<f32>,
    pub(super) alignment: Option<Alignment>,
    pub(super) contextual_spacing: bool,
    pub(super) keep_next: bool,
    pub(super) line_spacing: Option<f32>,
    pub(super) border_bottom_extra: f32,
    pub(super) border_bottom: Option<crate::model::BorderBottom>,
    pub(super) based_on: Option<String>,
}

pub(super) struct CharacterStyle {
    pub(super) font_size: Option<f32>,
    pub(super) font_name: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) underline: Option<bool>,
    pub(super) strikethrough: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
}

pub(super) struct TableBordersDef {
    pub(super) top: CellBorder,
    pub(super) bottom: CellBorder,
    pub(super) left: CellBorder,
    pub(super) right: CellBorder,
    pub(super) inside_h: CellBorder,
    pub(super) inside_v: CellBorder,
}

pub(super) struct StylesInfo {
    pub(super) defaults: StyleDefaults,
    pub(super) paragraph_styles: HashMap<String, ParagraphStyle>,
    pub(super) character_styles: HashMap<String, CharacterStyle>,
    pub(super) table_border_styles: HashMap<String, TableBordersDef>,
}

pub(super) fn parse_alignment(val: &str) -> Alignment {
    match val {
        "center" => Alignment::Center,
        "right" | "end" => Alignment::Right,
        "both" => Alignment::Justify,
        _ => Alignment::Left,
    }
}

pub(super) fn parse_theme(zip: &mut zip::ZipArchive<std::fs::File>) -> ThemeFonts {
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

pub(super) fn resolve_font(
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

pub(super) fn resolve_font_from_node(
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

pub(super) fn parse_styles(
    zip: &mut zip::ZipArchive<std::fs::File>,
    theme: &ThemeFonts,
) -> StylesInfo {
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
