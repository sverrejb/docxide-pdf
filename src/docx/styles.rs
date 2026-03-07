use std::collections::HashMap;

use crate::model::{Alignment, CellBorder, LineSpacing, TabStop};

use super::{
    DML_NS, WML_NS, parse_hex_color, parse_paragraph_borders, parse_tab_stops, parse_text_color,
    read_zip_text, twips_attr, twips_to_pts, wml, wml_attr, wml_bool,
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

fn ea_typeface<'a>(node: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    dml(node, "ea")
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

fn script_font_typeface<'a>(
    font_group: roxmltree::Node<'a, 'a>,
    script: &str,
) -> Option<&'a str> {
    font_group
        .children()
        .find(|n| {
            n.tag_name().name() == "font"
                && n.tag_name().namespace() == Some(DML_NS)
                && n.attribute("script") == Some(script)
        })
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

fn lang_to_script(lang: &str) -> &'static str {
    if lang.starts_with("ja") {
        "Jpan"
    } else if lang.starts_with("zh") && lang.contains("TW") {
        "Hant"
    } else if lang.starts_with("zh") {
        "Hans"
    } else if lang.starts_with("ko") {
        "Hang"
    } else {
        "Jpan"
    }
}

pub(super) struct ThemeFonts {
    pub(super) major: String,
    pub(super) minor: String,
    pub(super) major_east_asia: String,
    pub(super) minor_east_asia: String,
    pub(super) colors: HashMap<String, [u8; 3]>,
}

pub(super) struct StyleDefaults {
    pub(super) font_size: f32,
    pub(super) font_name: String,
    pub(super) east_asia_font: Option<String>,
    pub(super) space_after: f32,
    pub(super) line_spacing: LineSpacing,
    pub(super) kern_threshold: Option<f32>,
    pub(super) bold: bool,
    pub(super) italic: bool,
    pub(super) caps: bool,
    pub(super) small_caps: bool,
    pub(super) vanish: bool,
    pub(super) strikethrough: bool,
    pub(super) dstrike: bool,
    pub(super) underline: bool,
    pub(super) color: Option<[u8; 3]>,
    pub(super) char_spacing: f32,
}

pub(super) struct ParagraphStyle {
    pub(super) font_size: Option<f32>,
    pub(super) font_name: Option<String>,
    pub(super) east_asia_font: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) caps: Option<bool>,
    pub(super) small_caps: Option<bool>,
    pub(super) vanish: Option<bool>,
    pub(super) underline: Option<bool>,
    pub(super) strikethrough: Option<bool>,
    pub(super) dstrike: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
    pub(super) char_spacing: Option<f32>,
    pub(super) space_before: Option<f32>,
    pub(super) space_after: Option<f32>,
    pub(super) alignment: Option<Alignment>,
    pub(super) contextual_spacing: bool,
    pub(super) keep_next: bool,
    pub(super) keep_lines: bool,
    pub(super) page_break_before: bool,
    pub(super) line_spacing: Option<LineSpacing>,
    pub(super) indent_left: Option<f32>,
    pub(super) indent_right: Option<f32>,
    pub(super) indent_hanging: Option<f32>,
    pub(super) indent_first_line: Option<f32>,
    pub(super) borders: crate::model::ParagraphBorders,
    pub(super) based_on: Option<String>,
    pub(super) kern_threshold: Option<f32>,
    pub(super) tab_stops: Vec<TabStop>,
}

pub(super) struct CharacterStyle {
    pub(super) font_size: Option<f32>,
    pub(super) font_name: Option<String>,
    pub(super) east_asia_font: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) underline: Option<bool>,
    pub(super) strikethrough: Option<bool>,
    pub(super) caps: Option<bool>,
    pub(super) small_caps: Option<bool>,
    pub(super) vanish: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
    pub(super) kern_threshold: Option<f32>,
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
    /// Maps style ID → display name (for STYLEREF resolution)
    pub(super) style_id_to_name: HashMap<String, String>,
    /// The styleId of the default paragraph style (w:default="1" w:type="paragraph").
    /// Locale-dependent: "Normal" (English), "Normalny" (Polish), "Standard" (German/LibreOffice), etc.
    pub(super) default_paragraph_style_id: String,
}

pub(super) fn parse_alignment(val: &str) -> Alignment {
    match val {
        "center" => Alignment::Center,
        "right" | "end" => Alignment::Right,
        "both" => Alignment::Justify,
        _ => Alignment::Left,
    }
}

pub(super) fn parse_theme<R: std::io::Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    east_asia_lang: Option<&str>,
) -> ThemeFonts {
    let mut major = String::from("Aptos Display");
    let mut minor = String::from("Aptos");
    let mut major_east_asia = String::new();
    let mut minor_east_asia = String::new();
    let mut colors = HashMap::new();

    let script = east_asia_lang.map(lang_to_script).unwrap_or("Jpan");

    let names: Vec<String> = zip.file_names().map(|s: &str| s.to_string()).collect();
    let theme_name = names
        .iter()
        .find(|n: &&String| n.starts_with("word/theme/") && n.ends_with(".xml"));
    let Some(xml_content) = theme_name.and_then(|name| read_zip_text(zip, name.as_str())) else {
        return ThemeFonts {
            major,
            minor,
            major_east_asia,
            minor_east_asia,
            colors,
        };
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return ThemeFonts {
            major,
            minor,
            major_east_asia,
            minor_east_asia,
            colors,
        };
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
                major_east_asia = ea_typeface(node)
                    .or_else(|| script_font_typeface(node, script))
                    .unwrap_or("")
                    .to_string();
            }
            "minorFont" => {
                if let Some(tf) = latin_typeface(node) {
                    minor = tf.to_string();
                }
                minor_east_asia = ea_typeface(node)
                    .or_else(|| script_font_typeface(node, script))
                    .unwrap_or("")
                    .to_string();
            }
            "clrScheme" => {
                for child in node.children() {
                    if child.tag_name().namespace() != Some(DML_NS) {
                        continue;
                    }
                    let scheme_name = child.tag_name().name();
                    // a:srgbClr or a:sysClr child holds the color value
                    if let Some(srgb) = dml(child, "srgbClr") {
                        if let Some(hex) = srgb.attribute("val").and_then(super::parse_hex_color) {
                            colors.insert(scheme_name.to_string(), hex);
                        }
                    } else if let Some(hex) = dml(child, "sysClr")
                        .and_then(|sys| sys.attribute("lastClr"))
                        .and_then(super::parse_hex_color)
                    {
                        colors.insert(scheme_name.to_string(), hex);
                    }
                }
            }
            _ => {}
        }
    }

    ThemeFonts {
        major,
        minor,
        major_east_asia,
        minor_east_asia,
        colors,
    }
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
    let ascii = rfonts.attribute((WML_NS, "ascii"));
    let ascii_theme = rfonts.attribute((WML_NS, "asciiTheme"));
    if ascii.is_some() || ascii_theme.is_some() {
        return resolve_font(ascii, ascii_theme, theme, default_font);
    }
    // Fall back to hAnsi/hAnsiTheme when ascii variants are absent
    resolve_font(
        rfonts.attribute((WML_NS, "hAnsi")),
        rfonts.attribute((WML_NS, "hAnsiTheme")),
        theme,
        default_font,
    )
}

pub(super) fn resolve_east_asia_font(
    east_asia: Option<&str>,
    east_asia_theme: Option<&str>,
    theme: &ThemeFonts,
) -> Option<String> {
    let from_theme = match east_asia_theme {
        Some("majorEastAsia") if !theme.major_east_asia.is_empty() => {
            Some(theme.major_east_asia.clone())
        }
        Some("minorEastAsia") if !theme.minor_east_asia.is_empty() => {
            Some(theme.minor_east_asia.clone())
        }
        _ => None,
    };
    // eastAsiaTheme overrides eastAsia per spec
    from_theme.or_else(|| east_asia.filter(|s| !s.is_empty()).map(|s| s.to_string()))
}

pub(super) fn resolve_east_asia_font_from_node(
    rfonts: roxmltree::Node,
    theme: &ThemeFonts,
) -> Option<String> {
    let east_asia = rfonts.attribute((WML_NS, "eastAsia"));
    let east_asia_theme = rfonts.attribute((WML_NS, "eastAsiaTheme"));
    resolve_east_asia_font(east_asia, east_asia_theme, theme)
}

pub(super) fn parse_line_spacing(spacing_node: roxmltree::Node, line_val: f32) -> LineSpacing {
    match spacing_node.attribute((WML_NS, "lineRule")) {
        Some("exact") => LineSpacing::Exact(line_val / 20.0),
        Some("atLeast") => LineSpacing::AtLeast(line_val / 20.0),
        _ => LineSpacing::Auto(line_val / 240.0),
    }
}

pub(super) fn parse_styles<R: std::io::Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    theme: &ThemeFonts,
) -> StylesInfo {
    let mut defaults = StyleDefaults {
        font_size: 10.0,
        font_name: theme.minor.clone(),
        east_asia_font: None,
        space_after: 0.0,
        line_spacing: LineSpacing::Auto(1.0),
        kern_threshold: None,
        bold: false,
        italic: false,
        caps: false,
        small_caps: false,
        vanish: false,
        strikethrough: false,
        dstrike: false,
        underline: false,
        color: None,
        char_spacing: 0.0,
    };
    let mut paragraph_styles = HashMap::new();
    let mut character_styles = HashMap::new();
    let mut style_id_to_name: HashMap<String, String> = HashMap::new();
    let mut default_paragraph_style_id = String::from("Normal");

    let Some(xml_content) = read_zip_text(zip, "word/styles.xml") else {
        return StylesInfo {
            defaults,
            paragraph_styles,
            character_styles,
            table_border_styles: HashMap::new(),
            style_id_to_name,
            default_paragraph_style_id,
        };
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return StylesInfo {
            defaults,
            paragraph_styles,
            character_styles,
            table_border_styles: HashMap::new(),
            style_id_to_name,
            default_paragraph_style_id,
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
                defaults.east_asia_font = resolve_east_asia_font_from_node(rfonts, theme);
            }
            defaults.kern_threshold = wml_attr(rpr, "kern")
                .and_then(|v| v.parse::<f32>().ok())
                .map(|hp| hp / 2.0);
            defaults.bold = wml_bool(rpr, "b").unwrap_or(false);
            defaults.italic = wml_bool(rpr, "i").unwrap_or(false);
            defaults.caps = wml_bool(rpr, "caps").unwrap_or(false);
            defaults.small_caps = wml_bool(rpr, "smallCaps").unwrap_or(false);
            defaults.vanish = wml_bool(rpr, "vanish").unwrap_or(false);
            defaults.strikethrough = wml_bool(rpr, "strike").unwrap_or(false);
            defaults.dstrike = wml_bool(rpr, "dstrike").unwrap_or(false);
            defaults.underline = wml(rpr, "u")
                .and_then(|u| u.attribute((WML_NS, "val")))
                .map(|v| v != "none")
                .unwrap_or(false);
            defaults.color = wml_attr(rpr, "color").and_then(parse_text_color);
            defaults.char_spacing = wml(rpr, "spacing")
                .and_then(|n| n.attribute((WML_NS, "val")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(0.0);
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
                defaults.line_spacing = parse_line_spacing(spacing, line_val);
            }
        }
    }

    for style_node in root.children() {
        if style_node.tag_name().name() != "style"
            || style_node.tag_name().namespace() != Some(WML_NS)
        {
            continue;
        }

        // Collect style ID -> display name for all style types (used by STYLEREF)
        if let Some(id) = style_node.attribute((WML_NS, "styleId"))
            && let Some(name) = wml(style_node, "name").and_then(|n| n.attribute((WML_NS, "val")))
        {
            style_id_to_name.insert(id.to_string(), name.to_string());
        }

        if style_node.attribute((WML_NS, "type")) != Some("paragraph") {
            continue;
        }
        let Some(style_id) = style_node.attribute((WML_NS, "styleId")) else {
            continue;
        };

        if style_node.attribute((WML_NS, "default")) == Some("1") {
            default_paragraph_style_id = style_id.to_string();
        }

        let ppr = wml(style_node, "pPr");
        let spacing = ppr.and_then(|n| wml(n, "spacing"));
        let space_before = spacing.and_then(|n| twips_attr(n, "before"));
        let space_after = spacing.and_then(|n| twips_attr(n, "after"));
        let borders = ppr.map(parse_paragraph_borders).unwrap_or_default();

        let rpr = wml(style_node, "rPr");

        let font_size = rpr
            .and_then(|n| wml_attr(n, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);

        let rfonts_node = rpr.and_then(|n| wml(n, "rFonts"));
        let font_name =
            rfonts_node.map(|rfonts| resolve_font_from_node(rfonts, theme, &defaults.font_name));
        let east_asia_font = rfonts_node.and_then(|rfonts| resolve_east_asia_font_from_node(rfonts, theme));

        let bold = rpr.and_then(|n| wml_bool(n, "b"));
        let italic = rpr.and_then(|n| wml_bool(n, "i"));
        let caps = rpr.and_then(|n| wml_bool(n, "caps"));
        let small_caps = rpr.and_then(|n| wml_bool(n, "smallCaps"));
        let vanish = rpr.and_then(|n| wml_bool(n, "vanish"));
        let underline = rpr.and_then(|n| {
            wml(n, "u")
                .and_then(|u| u.attribute((WML_NS, "val")))
                .map(|v| v != "none")
        });
        let strikethrough = rpr.and_then(|n| wml_bool(n, "strike"));
        let dstrike = rpr.and_then(|n| wml_bool(n, "dstrike"));
        let char_spacing = rpr
            .and_then(|n| wml(n, "spacing"))
            .and_then(|n| n.attribute((WML_NS, "val")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts);
        let kern_threshold = rpr
            .and_then(|n| wml_attr(n, "kern"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);

        let color = rpr
            .and_then(|n| wml_attr(n, "color"))
            .and_then(parse_text_color);

        let alignment = ppr.and_then(|ppr| wml_attr(ppr, "jc")).map(parse_alignment);

        let contextual_spacing = ppr
            .and_then(|ppr| wml_bool(ppr, "contextualSpacing"))
            .unwrap_or(false);

        let keep_next = ppr
            .and_then(|ppr| wml_bool(ppr, "keepNext"))
            .unwrap_or(false);
        let keep_lines = ppr
            .and_then(|ppr| wml_bool(ppr, "keepLines"))
            .unwrap_or(false);
        let page_break_before = ppr
            .and_then(|ppr| wml_bool(ppr, "pageBreakBefore"))
            .unwrap_or(false);

        let line_spacing = spacing.and_then(|n| {
            n.attribute((WML_NS, "line"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|line_val| parse_line_spacing(n, line_val))
        });

        let ind = ppr.and_then(|n| wml(n, "ind"));
        let indent_left =
            ind.and_then(|n| twips_attr(n, "start").or_else(|| twips_attr(n, "left")));
        let indent_right =
            ind.and_then(|n| twips_attr(n, "end").or_else(|| twips_attr(n, "right")));
        let indent_hanging = ind.and_then(|n| twips_attr(n, "hanging"));
        let indent_first_line = ind.and_then(|n| twips_attr(n, "firstLine"));

        let tab_stops = ppr.map(parse_tab_stops).unwrap_or_default();

        let based_on = wml(style_node, "basedOn")
            .and_then(|n| n.attribute((WML_NS, "val")))
            .map(|s| s.to_string());

        paragraph_styles.insert(
            style_id.to_string(),
            ParagraphStyle {
                font_size,
                font_name,
                east_asia_font,
                bold,
                italic,
                caps,
                small_caps,
                vanish,
                underline,
                strikethrough,
                dstrike,
                color,
                char_spacing,
                space_before,
                space_after,
                alignment,
                contextual_spacing,
                keep_next,
                keep_lines,
                page_break_before,
                line_spacing,
                indent_left,
                indent_right,
                indent_hanging,
                indent_first_line,
                borders,
                based_on,
                kern_threshold,
                tab_stops,
            },
        );
    }

    resolve_based_on(&mut paragraph_styles);

    // The default paragraph style (w:default="1") may carry properties like w:kern
    // that aren't in docDefaults. Merge kern_threshold into defaults if missing.
    if defaults.kern_threshold.is_none() {
        if let Some(default_para) = paragraph_styles.get(&default_paragraph_style_id) {
            if default_para.kern_threshold.is_some() {
                defaults.kern_threshold = default_para.kern_threshold;
            }
        }
    }

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
        let char_rfonts_node = wml(rpr, "rFonts");
        let font_name =
            char_rfonts_node.map(|rfonts| resolve_font_from_node(rfonts, theme, &defaults.font_name));
        let east_asia_font = char_rfonts_node.and_then(|rfonts| resolve_east_asia_font_from_node(rfonts, theme));
        let bold = wml_bool(rpr, "b");
        let italic = wml_bool(rpr, "i");
        let underline = wml(rpr, "u")
            .and_then(|n| n.attribute((WML_NS, "val")))
            .map(|v| v != "none");
        let strikethrough = wml_bool(rpr, "strike");
        let caps = wml_bool(rpr, "caps");
        let small_caps = wml_bool(rpr, "smallCaps");
        let vanish = wml_bool(rpr, "vanish");
        let color = wml_attr(rpr, "color").and_then(parse_text_color);
        let kern_threshold = wml_attr(rpr, "kern")
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);

        character_styles.insert(
            style_id.to_string(),
            CharacterStyle {
                font_size,
                font_name,
                east_asia_font,
                bold,
                italic,
                underline,
                strikethrough,
                caps,
                small_caps,
                vanish,
                color,
                kern_threshold,
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
        if let Some(tbl_borders) = wml(style_node, "tblPr").and_then(|pr| wml(pr, "tblBorders")) {
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
        style_id_to_name,
        default_paragraph_style_id,
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
            east_asia_font: None,
            bold: None,
            italic: None,
            caps: None,
            small_caps: None,
            vanish: None,
            underline: None,
            strikethrough: None,
            dstrike: None,
            color: None,
            char_spacing: None,
            space_before: None,
            space_after: None,
            alignment: None,
            contextual_spacing: false,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            line_spacing: None,
            indent_left: None,
            indent_right: None,
            indent_hanging: None,
            indent_first_line: None,
            borders: crate::model::ParagraphBorders::default(),
            based_on: None,
            kern_threshold: None,
            tab_stops: Vec::new(),
        };

        for ancestor_id in chain.iter().rev() {
            if let Some(s) = styles.get(ancestor_id) {
                inherit!(font_name, inh.font_name, s);
                inherit!(east_asia_font, inh.east_asia_font, s);
                inherit!(font_size, inh.font_size, s);
                inherit!(bold, inh.bold, s);
                inherit!(italic, inh.italic, s);
                inherit!(caps, inh.caps, s);
                inherit!(small_caps, inh.small_caps, s);
                inherit!(vanish, inh.vanish, s);
                inherit!(underline, inh.underline, s);
                inherit!(strikethrough, inh.strikethrough, s);
                inherit!(dstrike, inh.dstrike, s);
                inherit!(color, inh.color, s);
                inherit!(char_spacing, inh.char_spacing, s);
                inherit!(alignment, inh.alignment, s);
                inherit!(space_before, inh.space_before, s);
                inherit!(space_after, inh.space_after, s);
                inherit!(line_spacing, inh.line_spacing, s);
                inherit!(indent_left, inh.indent_left, s);
                inherit!(indent_right, inh.indent_right, s);
                inherit!(indent_hanging, inh.indent_hanging, s);
                inherit!(indent_first_line, inh.indent_first_line, s);
                inherit!(kern_threshold, inh.kern_threshold, s);
                // Tab stops are additive: accumulate from ancestors, child overrides at same pos
                for ts in &s.tab_stops {
                    if let Some(existing) = inh
                        .tab_stops
                        .iter_mut()
                        .find(|t| (t.position - ts.position).abs() < 0.5)
                    {
                        *existing = ts.clone();
                    } else {
                        inh.tab_stops.push(ts.clone());
                    }
                }
            }
        }
        inh.tab_stops
            .sort_by(|a, b| a.position.total_cmp(&b.position));

        if let Some(s) = styles.get_mut(&id) {
            s.font_name = s.font_name.take().or(inh.font_name);
            s.east_asia_font = s.east_asia_font.take().or(inh.east_asia_font);
            s.font_size = s.font_size.or(inh.font_size);
            s.bold = s.bold.or(inh.bold);
            s.italic = s.italic.or(inh.italic);
            s.caps = s.caps.or(inh.caps);
            s.small_caps = s.small_caps.or(inh.small_caps);
            s.vanish = s.vanish.or(inh.vanish);
            s.color = s.color.or(inh.color);
            s.alignment = s.alignment.or(inh.alignment);
            s.space_before = s.space_before.or(inh.space_before);
            s.space_after = s.space_after.or(inh.space_after);
            s.line_spacing = s.line_spacing.or(inh.line_spacing);
            s.indent_left = s.indent_left.or(inh.indent_left);
            s.indent_right = s.indent_right.or(inh.indent_right);
            s.indent_hanging = s.indent_hanging.or(inh.indent_hanging);
            s.indent_first_line = s.indent_first_line.or(inh.indent_first_line);
            s.kern_threshold = s.kern_threshold.or(inh.kern_threshold);
            if s.tab_stops.is_empty() {
                s.tab_stops = inh.tab_stops;
            }
        }
    }
}
