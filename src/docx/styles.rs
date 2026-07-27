use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use crate::model::{
    Alignment, CellBorder, LineSpacing, ParagraphBorders, TabStop, TextFill, TextGlow,
    TextOutline, TextShadow,
};

pub(super) use super::color::{ColorTransforms, parse_color_transforms};
use super::wordart::{parse_text_fill, parse_text_glow, parse_text_outline, parse_text_shadow};
use super::{
    DML_NS, WML_NS, dml, extract_indents, highlight_color, parse_cell_border,
    parse_cell_border_left, parse_cell_border_right, parse_hex_color, parse_one_border,
    merge_tab_stops, parse_on_off, parse_paragraph_borders, parse_run_shd,
    parse_tab_stops_with_clears, parse_text_color, read_zip_text, twips_attr, twips_to_pts, wml,
    wml_attr, wml_bool,
};

fn dml_typeface<'a>(node: roxmltree::Node<'a, 'a>, element: &str) -> Option<&'a str> {
    dml(node, element)
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

fn script_font_typeface<'a>(font_group: roxmltree::Node<'a, 'a>, script: &str) -> Option<&'a str> {
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

pub(super) struct ThemeGradientStop {
    pub(super) position: f32,
    pub(super) transforms: ColorTransforms,
}

pub(super) enum ThemeFillStyle {
    Solid,
    Gradient {
        stops: Vec<ThemeGradientStop>,
        angle_deg: f32,
    },
}

pub(super) struct ThemeFonts {
    pub(super) major: String,
    pub(super) minor: String,
    pub(super) major_east_asia: String,
    pub(super) minor_east_asia: String,
    pub(super) colors: HashMap<String, [u8; 3]>,
    pub(super) fill_styles: Vec<ThemeFillStyle>,
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
    pub(super) double_underline: bool,
    pub(super) color: Option<[u8; 3]>,
    pub(super) char_spacing: f32,
    pub(super) widow_control: bool,
    pub(super) indent_left: f32,
    pub(super) indent_right: f32,
    pub(super) indent_hanging: f32,
    pub(super) indent_first_line: f32,
    pub(super) alignment: Alignment,
}

#[derive(Default)]
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
    pub(super) double_underline: Option<bool>,
    pub(super) strikethrough: Option<bool>,
    pub(super) dstrike: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
    pub(super) char_spacing: Option<f32>,
    pub(super) space_before: Option<f32>,
    pub(super) space_after: Option<f32>,
    pub(super) space_before_autospacing: Option<bool>,
    pub(super) space_after_autospacing: Option<bool>,
    pub(super) alignment: Option<Alignment>,
    pub(super) contextual_spacing: bool,
    pub(super) keep_next: bool,
    pub(super) keep_lines: bool,
    pub(super) widow_control: Option<bool>,
    pub(super) page_break_before: bool,
    pub(super) line_spacing: Option<LineSpacing>,
    pub(super) indent_left: Option<f32>,
    pub(super) indent_right: Option<f32>,
    pub(super) indent_hanging: Option<f32>,
    pub(super) indent_first_line: Option<f32>,
    pub(super) borders: ParagraphBorders,
    pub(super) shading: Option<[u8; 3]>,
    pub(super) based_on: Option<String>,
    pub(super) kern_threshold: Option<f32>,
    pub(super) tab_stops: Vec<TabStop>,
    pub(super) clear_tab_positions: Vec<f32>,
    pub(super) num_id: Option<String>,
    pub(super) num_ilvl: Option<u8>,
    pub(super) outline_level: Option<u8>,
    pub(super) snap_to_grid: Option<bool>,
    pub(super) auto_space_de: Option<bool>,
    pub(super) auto_space_dn: Option<bool>,
    pub(super) suppress_auto_hyphens: Option<bool>,
    pub(super) text_outline: Option<TextOutline>,
    pub(super) text_fill: Option<TextFill>,
    pub(super) text_shadow: Option<TextShadow>,
    pub(super) text_glow: Option<TextGlow>,
}

pub(super) struct CharacterStyle {
    pub(super) font_size: Option<f32>,
    pub(super) font_name: Option<String>,
    pub(super) east_asia_font: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) underline: Option<bool>,
    pub(super) double_underline: Option<bool>,
    pub(super) strikethrough: Option<bool>,
    pub(super) caps: Option<bool>,
    pub(super) small_caps: Option<bool>,
    pub(super) vanish: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
    pub(super) highlight: Option<[u8; 3]>,
    /// Run-level shading from `w:rPr/w:shd`. `None` means "no explicit color
    /// here"; callers inherit from basedOn. Word treats `w:val="nil"` and
    /// `w:fill="auto"` as "no color," not as a clearing override.
    pub(super) shading: Option<[u8; 3]>,
    pub(super) border: Option<crate::model::ParagraphBorder>,
    pub(super) kern_threshold: Option<f32>,
    pub(super) text_outline: Option<TextOutline>,
    pub(super) text_fill: Option<TextFill>,
    pub(super) text_shadow: Option<TextShadow>,
    pub(super) text_glow: Option<TextGlow>,
}

pub(super) struct TableBordersDef {
    pub(super) top: CellBorder,
    pub(super) bottom: CellBorder,
    pub(super) left: CellBorder,
    pub(super) right: CellBorder,
    pub(super) inside_h: CellBorder,
    pub(super) inside_v: CellBorder,
}

/// Conditional formatting for a specific table region (e.g. firstRow, band1Horz).
pub(super) struct TableConditionalFormat {
    pub(super) borders: Option<TableBordersDef>,
    pub(super) shading: Option<[u8; 3]>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
    pub(super) font_size: Option<f32>,
    pub(super) font_name: Option<String>,
}

/// Full table style definition including conditional overrides.
pub(super) struct TableStyleDef {
    pub(super) base_borders: Option<TableBordersDef>,
    pub(super) base_font_size: Option<f32>,
    pub(super) base_font_name: Option<String>,
    pub(super) base_bold: Option<bool>,
    pub(super) base_italic: Option<bool>,
    /// Conditional overrides keyed by type: "firstRow", "lastRow", "firstCol",
    /// "lastCol", "band1Horz", "band2Horz", "band1Vert", "band2Vert",
    /// "nwCell", "neCell", "swCell", "seCell"
    pub(super) conditionals: HashMap<String, TableConditionalFormat>,
}

pub(super) struct StylesInfo {
    pub(super) defaults: StyleDefaults,
    pub(super) paragraph_styles: HashMap<String, ParagraphStyle>,
    pub(super) character_styles: HashMap<String, CharacterStyle>,
    pub(super) table_styles: HashMap<String, TableStyleDef>,
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
        // Kashida variants are Arabic justification flavors: without shaping we
        // can't elongate glyphs, but plain justify beats falling back to left.
        //
        // thaiDistribute is Thai-specific: case77's reference shows Word leaving
        // a Latin thaiDistribute line at its natural width while stretching a
        // plain `distribute` line of the same shape, so it behaves as ordinary
        // justify here. Whether Thai script triggers real distribution is
        // untested — no Thai fixture yet.
        "both" | "mediumKashida" | "highKashida" | "lowKashida" | "thaiDistribute" => {
            Alignment::Justify
        }
        "distribute" => Alignment::Distribute,
        _ => Alignment::Left,
    }
}

pub(super) fn parse_font_size(rpr: roxmltree::Node) -> Option<f32> {
    wml_attr(rpr, "sz")
        .and_then(|v| v.parse::<f32>().ok())
        .map(|hp| hp / 2.0)
}

fn parse_kern(rpr: roxmltree::Node) -> Option<f32> {
    wml_attr(rpr, "kern")
        .and_then(|v| v.parse::<f32>().ok())
        .map(|hp| hp / 2.0)
}

/// rFonts ascii (falling back to hAnsi) typeface name. Deliberately ignores
/// theme fonts — callers needing theme resolution use resolve_font_from_node.
pub(super) fn rfonts_ascii_name(rpr: roxmltree::Node) -> Option<String> {
    wml(rpr, "rFonts")
        .and_then(|rf| {
            rf.attribute((WML_NS, "ascii"))
                .or_else(|| rf.attribute((WML_NS, "hAnsi")))
        })
        .map(|s| s.to_string())
}

// Underline state comes from the w:val attribute, not the presence of <w:u>.
// A bare <w:u> with no val is "no underline applied" in Word (inherit), so
// returning None lets basedOn/defaults resolve it instead of forcing it on.
fn parse_underline(rpr: roxmltree::Node) -> Option<bool> {
    wml_attr(rpr, "u").map(|v| v != "none")
}

fn parse_double_underline(rpr: roxmltree::Node) -> Option<bool> {
    wml_attr(rpr, "u").map(|v| v == "double")
}

pub(super) fn parse_char_spacing(rpr: roxmltree::Node) -> Option<f32> {
    wml_attr(rpr, "spacing")
        .and_then(|v| v.parse::<f32>().ok())
        .map(twips_to_pts)
}

pub(super) fn parse_theme<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    east_asia_lang: Option<&str>,
) -> ThemeFonts {
    let mut major = String::from("Aptos Display");
    let mut minor = String::from("Aptos");
    let mut major_east_asia = String::new();
    let mut minor_east_asia = String::new();
    let mut colors = HashMap::new();
    let mut fill_styles = Vec::new();

    let script = east_asia_lang.map(lang_to_script).unwrap_or("Jpan");

    let names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
    let xml_content = names
        .iter()
        .find(|n| n.starts_with("word/theme/") && n.ends_with(".xml"))
        .and_then(|name| read_zip_text(zip, name));

    if let Some(xml_content) = xml_content
        && let Ok(xml) = roxmltree::Document::parse(&xml_content)
    {
        for node in xml.descendants() {
            if node.tag_name().namespace() != Some(DML_NS) {
                continue;
            }
            match node.tag_name().name() {
                "majorFont" => {
                    if let Some(tf) = dml_typeface(node, "latin") {
                        major = tf.to_string();
                    }
                    major_east_asia = dml_typeface(node, "ea")
                        .or_else(|| script_font_typeface(node, script))
                        .unwrap_or("")
                        .to_string();
                }
                "minorFont" => {
                    if let Some(tf) = dml_typeface(node, "latin") {
                        minor = tf.to_string();
                    }
                    minor_east_asia = dml_typeface(node, "ea")
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
                        if let Some(srgb) = dml(child, "srgbClr") {
                            if let Some(hex) =
                                srgb.attribute("val").and_then(parse_hex_color)
                            {
                                colors.insert(scheme_name.to_string(), hex);
                            }
                        } else if let Some(hex) = dml(child, "sysClr")
                            .and_then(|sys| sys.attribute("lastClr"))
                            .and_then(parse_hex_color)
                        {
                            colors.insert(scheme_name.to_string(), hex);
                        }
                    }
                }
                "fillStyleLst" => {
                    for child in node
                        .children()
                        .filter(|n| n.tag_name().namespace() == Some(DML_NS))
                    {
                        match child.tag_name().name() {
                            "solidFill" => fill_styles.push(ThemeFillStyle::Solid),
                            "gradFill" => {
                                if let Some(gs_lst) = dml(child, "gsLst") {
                                    let stops = parse_theme_gradient_stops(gs_lst);
                                    let angle_deg = dml(child, "lin")
                                        .and_then(|lin| lin.attribute("ang"))
                                        .and_then(|v| v.parse::<f32>().ok())
                                        .map(|v| v / 60_000.0)
                                        .unwrap_or(0.0);
                                    fill_styles.push(ThemeFillStyle::Gradient { stops, angle_deg });
                                }
                            }
                            _ => {}
                        }
                    }
                }
                _ => {}
            }
        }
    }

    ThemeFonts {
        major,
        minor,
        major_east_asia,
        minor_east_asia,
        colors,
        fill_styles,
    }
}

fn parse_theme_gradient_stops(gs_lst: roxmltree::Node) -> Vec<ThemeGradientStop> {
    gs_lst
        .children()
        .filter(|n| n.tag_name().name() == "gs" && n.tag_name().namespace() == Some(DML_NS))
        .map(|gs| {
            let position = gs
                .attribute("pos")
                .and_then(|v| v.parse::<f32>().ok())
                .map(|v| v / 100_000.0)
                .unwrap_or(0.0);
            let transforms = gs
                .descendants()
                .find(|n| {
                    n.tag_name().name() == "schemeClr" && n.tag_name().namespace() == Some(DML_NS)
                })
                .map(parse_color_transforms)
                .unwrap_or_default();
            ThemeGradientStop {
                position,
                transforms,
            }
        })
        .collect()
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
    resolve_font_from_node_opt(rfonts, theme).unwrap_or_else(|| default_font.to_string())
}

/// Returns Some only when rFonts actually specifies an ascii/hAnsi font or theme.
/// When only cstheme/eastAsia variants are set (e.g. Heading1 with `<w:rFonts w:cstheme="minorHAnsi"/>`),
/// returns None so that parent-style font inheritance applies instead of falling back to docDefaults.
pub(super) fn resolve_font_from_node_opt(
    rfonts: roxmltree::Node,
    theme: &ThemeFonts,
) -> Option<String> {
    let ascii = rfonts.attribute((WML_NS, "ascii"));
    let ascii_theme = rfonts.attribute((WML_NS, "asciiTheme"));
    if ascii.is_some() || ascii_theme.is_some() {
        return Some(resolve_font(ascii, ascii_theme, theme, ""));
    }
    let hansi = rfonts.attribute((WML_NS, "hAnsi"));
    let hansi_theme = rfonts.attribute((WML_NS, "hAnsiTheme"));
    if hansi.is_some() || hansi_theme.is_some() {
        return Some(resolve_font(hansi, hansi_theme, theme, ""));
    }
    None
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
        Some("exact") => LineSpacing::Exact(twips_to_pts(line_val)),
        Some("atLeast") => LineSpacing::AtLeast(twips_to_pts(line_val)),
        _ => LineSpacing::Auto(line_val / 240.0),
    }
}

pub(super) fn parse_styles<R: Read + Seek>(
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
        double_underline: false,
        color: None,
        char_spacing: 0.0,
        widow_control: true,
        indent_left: 0.0,
        indent_right: 0.0,
        indent_hanging: 0.0,
        indent_first_line: 0.0,
        alignment: Alignment::Left,
    };
    let mut paragraph_styles = HashMap::new();
    let mut character_styles = HashMap::new();
    let mut style_id_to_name = HashMap::new();
    let mut default_paragraph_style_id = String::from("Normal");

    let Some(xml_content) = read_zip_text(zip, "word/styles.xml") else {
        return StylesInfo {
            defaults,
            paragraph_styles,
            character_styles,
            table_styles: HashMap::new(),
            style_id_to_name,
            default_paragraph_style_id,
        };
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return StylesInfo {
            defaults,
            paragraph_styles,
            character_styles,
            table_styles: HashMap::new(),
            style_id_to_name,
            default_paragraph_style_id,
        };
    };

    let root = xml.root_element();

    if let Some(doc_defaults) = wml(root, "docDefaults") {
        if let Some(rpr) = wml(doc_defaults, "rPrDefault").and_then(|n| wml(n, "rPr")) {
            if let Some(fs) = parse_font_size(rpr) {
                defaults.font_size = fs;
            }
            if let Some(rfonts) = wml(rpr, "rFonts") {
                defaults.font_name = resolve_font_from_node(rfonts, theme, &theme.minor);
                defaults.east_asia_font = resolve_east_asia_font_from_node(rfonts, theme);
            }
            defaults.kern_threshold = parse_kern(rpr);
            defaults.bold = wml_bool(rpr, "b").unwrap_or(false);
            defaults.italic = wml_bool(rpr, "i").unwrap_or(false);
            defaults.caps = wml_bool(rpr, "caps").unwrap_or(false);
            defaults.small_caps = wml_bool(rpr, "smallCaps").unwrap_or(false);
            defaults.vanish = wml_bool(rpr, "vanish").unwrap_or(false);
            defaults.strikethrough = wml_bool(rpr, "strike").unwrap_or(false);
            defaults.dstrike = wml_bool(rpr, "dstrike").unwrap_or(false);
            defaults.underline = parse_underline(rpr).unwrap_or(false);
            defaults.double_underline = parse_double_underline(rpr).unwrap_or(false);
            defaults.color = wml_attr(rpr, "color").and_then(parse_text_color);
            defaults.char_spacing = parse_char_spacing(rpr).unwrap_or(0.0);
        }
        let default_ppr = wml(doc_defaults, "pPrDefault").and_then(|n| wml(n, "pPr"));
        if let Some(wc) = default_ppr.and_then(|ppr| wml_bool(ppr, "widowControl")) {
            defaults.widow_control = wc;
        }
        let default_spacing = default_ppr.and_then(|n| wml(n, "spacing"));
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
        if let Some(ind) = default_ppr.and_then(|n| wml(n, "ind")) {
            let (left, right, hanging, first) = extract_indents(ind, None);
            if let Some(v) = left {
                defaults.indent_left = v;
            }
            if let Some(v) = right {
                defaults.indent_right = v;
            }
            if let Some(v) = hanging {
                defaults.indent_hanging = v;
            }
            if let Some(v) = first {
                defaults.indent_first_line = v;
            }
        }
        if let Some(jc) = default_ppr.and_then(|ppr| wml_attr(ppr, "jc")) {
            defaults.alignment = parse_alignment(jc);
        }
    }

    let mut table_styles = HashMap::new();

    for style_node in root.children() {
        if style_node.tag_name().name() != "style"
            || style_node.tag_name().namespace() != Some(WML_NS)
        {
            continue;
        }

        // Collect style ID -> display name for all style types (used by STYLEREF)
        if let Some(id) = style_node.attribute((WML_NS, "styleId"))
            && let Some(name) = wml_attr(style_node, "name")
        {
            style_id_to_name.insert(id.to_string(), name.to_string());
        }

        let Some(style_id) = style_node.attribute((WML_NS, "styleId")) else {
            continue;
        };

        match style_node.attribute((WML_NS, "type")) {
            Some("paragraph") => {
                if style_node.attribute((WML_NS, "default")) == Some("1") {
                    default_paragraph_style_id = style_id.to_string();
                }

                let ppr = wml(style_node, "pPr");
                let spacing = ppr.and_then(|n| wml(n, "spacing"));
                let space_before = spacing.and_then(|n| twips_attr(n, "before"));
                let space_after = spacing.and_then(|n| twips_attr(n, "after"));
                let space_before_autospacing = spacing
                    .and_then(|n| n.attribute((WML_NS, "beforeAutospacing")).map(parse_on_off));
                let space_after_autospacing = spacing
                    .and_then(|n| n.attribute((WML_NS, "afterAutospacing")).map(parse_on_off));
                let borders = ppr.and_then(parse_paragraph_borders).unwrap_or_default();
                let shading = ppr
                    .and_then(|n| wml(n, "shd"))
                    .and_then(|shd| shd.attribute((WML_NS, "fill")))
                    .and_then(parse_hex_color);

                let rpr = wml(style_node, "rPr");

                let font_size = rpr.and_then(parse_font_size);
                let rfonts_node = rpr.and_then(|n| wml(n, "rFonts"));
                let font_name =
                    rfonts_node.and_then(|rfonts| resolve_font_from_node_opt(rfonts, theme));
                let east_asia_font =
                    rfonts_node.and_then(|rfonts| resolve_east_asia_font_from_node(rfonts, theme));

                let bold = rpr.and_then(|n| wml_bool(n, "b"));
                let italic = rpr.and_then(|n| wml_bool(n, "i"));
                let caps = rpr.and_then(|n| wml_bool(n, "caps"));
                let small_caps = rpr.and_then(|n| wml_bool(n, "smallCaps"));
                let vanish = rpr.and_then(|n| wml_bool(n, "vanish"));
                let underline = rpr.and_then(parse_underline);
                let double_underline = rpr.and_then(parse_double_underline);
                let strikethrough = rpr.and_then(|n| wml_bool(n, "strike"));
                let dstrike = rpr.and_then(|n| wml_bool(n, "dstrike"));
                let char_spacing = rpr.and_then(parse_char_spacing);
                let kern_threshold = rpr.and_then(parse_kern);
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
                let widow_control = ppr.and_then(|ppr| wml_bool(ppr, "widowControl"));
                let page_break_before = ppr
                    .and_then(|ppr| wml_bool(ppr, "pageBreakBefore"))
                    .unwrap_or(false);

                let line_spacing = spacing.and_then(|n| {
                    n.attribute((WML_NS, "line"))
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|line_val| parse_line_spacing(n, line_val))
                });

                let (indent_left, indent_right, indent_hanging, indent_first_line) = ppr
                    .and_then(|n| wml(n, "ind"))
                    .map(|ind| extract_indents(ind, None))
                    .unwrap_or_default();

                let (tab_stops, clear_tab_positions) = ppr
                    .map(parse_tab_stops_with_clears)
                    .unwrap_or_default();

                let style_num_pr = ppr.and_then(|p| wml(p, "numPr"));
                let num_id = style_num_pr
                    .and_then(|np| wml_attr(np, "numId"))
                    .map(|s| s.to_string());
                let num_ilvl = style_num_pr
                    .and_then(|np| wml_attr(np, "ilvl"))
                    .and_then(|v| v.parse::<u8>().ok());

                let outline_level = ppr
                    .and_then(|p| wml_attr(p, "outlineLvl"))
                    .and_then(|v| v.parse::<u8>().ok())
                    .filter(|&lvl| lvl <= 8);

                let snap_to_grid = ppr.and_then(|ppr| wml_bool(ppr, "snapToGrid"));
                let auto_space_de = ppr.and_then(|ppr| wml_bool(ppr, "autoSpaceDE"));
                let auto_space_dn = ppr.and_then(|ppr| wml_bool(ppr, "autoSpaceDN"));
                let suppress_auto_hyphens =
                    ppr.and_then(|ppr| wml_bool(ppr, "suppressAutoHyphens"));

                let based_on = wml_attr(style_node, "basedOn").map(|s| s.to_string());

                let text_outline = rpr.and_then(|n| parse_text_outline(n, theme));
                let text_fill = rpr.and_then(|n| parse_text_fill(n, theme));
                let text_shadow = rpr.and_then(|n| parse_text_shadow(n, theme));
                let text_glow = rpr.and_then(|n| parse_text_glow(n, theme));

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
                        double_underline,
                        strikethrough,
                        dstrike,
                        color,
                        char_spacing,
                        space_before,
                        space_after,
                        space_before_autospacing,
                        space_after_autospacing,
                        alignment,
                        contextual_spacing,
                        keep_next,
                        keep_lines,
                        widow_control,
                        page_break_before,
                        line_spacing,
                        indent_left,
                        indent_right,
                        indent_hanging,
                        indent_first_line,
                        borders,
                        shading,
                        based_on,
                        kern_threshold,
                        tab_stops,
                        clear_tab_positions,
                        num_id,
                        num_ilvl,
                        outline_level,
                        snap_to_grid,
                        auto_space_de,
                        auto_space_dn,
                        suppress_auto_hyphens,
                        text_outline,
                        text_fill,
                        text_shadow,
                        text_glow,
                    },
                );
            }
            Some("character") => {
                let Some(rpr) = wml(style_node, "rPr") else {
                    continue;
                };
                let font_size = parse_font_size(rpr);
                let rfonts_node = wml(rpr, "rFonts");
                let font_name =
                    rfonts_node.and_then(|rfonts| resolve_font_from_node_opt(rfonts, theme));
                let east_asia_font =
                    rfonts_node.and_then(|rfonts| resolve_east_asia_font_from_node(rfonts, theme));
                let bold = wml_bool(rpr, "b");
                let italic = wml_bool(rpr, "i");
                let underline = parse_underline(rpr);
                let double_underline = parse_double_underline(rpr);
                let strikethrough = wml_bool(rpr, "strike");
                let caps = wml_bool(rpr, "caps");
                let small_caps = wml_bool(rpr, "smallCaps");
                let vanish = wml_bool(rpr, "vanish");
                let color = wml_attr(rpr, "color").and_then(parse_text_color);
                let highlight = wml_attr(rpr, "highlight").and_then(highlight_color);
                let shading = parse_run_shd(rpr);
                let border = wml(rpr, "bdr").and_then(parse_one_border);
                let kern_threshold = parse_kern(rpr);
                let text_outline = parse_text_outline(rpr, theme);
                let text_fill = parse_text_fill(rpr, theme);
                let text_shadow = parse_text_shadow(rpr, theme);
                let text_glow = parse_text_glow(rpr, theme);

                character_styles.insert(
                    style_id.to_string(),
                    CharacterStyle {
                        font_size,
                        font_name,
                        east_asia_font,
                        bold,
                        italic,
                        underline,
                        double_underline,
                        strikethrough,
                        caps,
                        small_caps,
                        vanish,
                        color,
                        highlight,
                        shading,
                        border,
                        kern_threshold,
                        text_outline,
                        text_fill,
                        text_shadow,
                        text_glow,
                    },
                );
            }
            Some("table") => {
                let base_borders = wml(style_node, "tblPr")
                    .and_then(|pr| wml(pr, "tblBorders"))
                    .map(|tbl_borders| TableBordersDef {
                        top: parse_cell_border(tbl_borders, "top"),
                        bottom: parse_cell_border(tbl_borders, "bottom"),
                        left: parse_cell_border_left(tbl_borders),
                        right: parse_cell_border_right(tbl_borders),
                        inside_h: parse_cell_border(tbl_borders, "insideH"),
                        inside_v: parse_cell_border(tbl_borders, "insideV"),
                    });

                // Parse base rPr from the table style
                let base_rpr = wml(style_node, "rPr");
                let base_font_size = base_rpr.and_then(parse_font_size);
                let base_font_name = base_rpr.and_then(rfonts_ascii_name);
                let base_bold = base_rpr.and_then(|rpr| wml_bool(rpr, "b"));
                let base_italic = base_rpr.and_then(|rpr| wml_bool(rpr, "i"));

                let mut conditionals = HashMap::new();
                for child in style_node.children() {
                    if !child.has_tag_name((WML_NS, "tblStylePr")) {
                        continue;
                    }
                    let Some(cond_type) = child.attribute((WML_NS, "type")) else {
                        continue;
                    };
                    let cond_borders = wml(child, "tcPr")
                        .and_then(|tc| wml(tc, "tcBorders"))
                        .map(|b| TableBordersDef {
                            top: parse_cell_border(b, "top"),
                            bottom: parse_cell_border(b, "bottom"),
                            left: parse_cell_border_left(b),
                            right: parse_cell_border_right(b),
                            inside_h: parse_cell_border(b, "insideH"),
                            inside_v: parse_cell_border(b, "insideV"),
                        });
                    let cond_shading = wml(child, "tcPr")
                        .and_then(|tc| wml(tc, "shd"))
                        .and_then(|shd| shd.attribute((WML_NS, "fill")))
                        .and_then(parse_hex_color);
                    let cond_rpr = wml(child, "rPr");
                    let cond_bold = cond_rpr.and_then(|rpr| wml_bool(rpr, "b"));
                    let cond_italic = cond_rpr.and_then(|rpr| wml_bool(rpr, "i"));
                    let cond_color = cond_rpr
                        .and_then(|rpr| wml_attr(rpr, "color"))
                        .and_then(parse_text_color);
                    let cond_font_size = cond_rpr.and_then(|rpr| parse_font_size(rpr));
                    let cond_font_name = cond_rpr.and_then(rfonts_ascii_name);
                    if cond_borders.is_some()
                        || cond_shading.is_some()
                        || cond_bold.is_some()
                        || cond_color.is_some()
                        || cond_font_size.is_some()
                        || cond_font_name.is_some()
                        || cond_italic.is_some()
                    {
                        conditionals.insert(cond_type.to_string(), TableConditionalFormat {
                            borders: cond_borders,
                            shading: cond_shading,
                            bold: cond_bold,
                            italic: cond_italic,
                            color: cond_color,
                            font_size: cond_font_size,
                            font_name: cond_font_name,
                        });
                    }
                }

                if base_borders.is_some()
                    || !conditionals.is_empty()
                    || base_font_size.is_some()
                    || base_font_name.is_some()
                {
                    table_styles.insert(
                        style_id.to_string(),
                        TableStyleDef {
                            base_borders,
                            base_font_size,
                            base_font_name,
                            base_bold,
                            base_italic,
                            conditionals,
                        },
                    );
                }
            }
            _ => {}
        }
    }

    resolve_based_on(&mut paragraph_styles);

    // The default paragraph style (w:default="1") may carry properties like w:kern
    // that aren't in docDefaults. Merge kern_threshold into defaults if missing.
    if defaults.kern_threshold.is_none() {
        if let Some(default_para) = paragraph_styles.get(&default_paragraph_style_id) {
            defaults.kern_threshold = default_para.kern_threshold;
        }
    }

    StylesInfo {
        defaults,
        paragraph_styles,
        character_styles,
        table_styles,
        style_id_to_name,
        default_paragraph_style_id,
    }
}

fn resolve_based_on(styles: &mut HashMap<String, ParagraphStyle>) {
    let ids: Vec<String> = styles.keys().cloned().collect();
    for id in ids {
        let mut visited: HashSet<String> = HashSet::new();
        let mut chain: Vec<String> = Vec::new();
        let mut current = id.clone();
        loop {
            if !visited.insert(current.clone()) {
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
            ($dst:expr, $src:expr, $($field:ident),+ $(,)?) => {
                $(if $src.$field.is_some() { $dst.$field = $src.$field.clone(); })+
            };
        }

        let mut inh = ParagraphStyle::default();

        for ancestor_id in chain.iter().rev() {
            if let Some(s) = styles.get(ancestor_id) {
                inherit!(
                    inh,
                    s,
                    font_name,
                    east_asia_font,
                    font_size,
                    bold,
                    italic,
                    caps,
                    small_caps,
                    vanish,
                    underline,
                    double_underline,
                    strikethrough,
                    dstrike,
                    color,
                    char_spacing,
                    alignment,
                    space_before,
                    space_after,
                    space_before_autospacing,
                    space_after_autospacing,
                    line_spacing,
                    indent_left,
                    indent_right,
                    indent_hanging,
                    indent_first_line,
                    kern_threshold,
                    widow_control,
                    num_id,
                    num_ilvl,
                    outline_level,
                    snap_to_grid,
                    auto_space_de,
                    auto_space_dn,
                    suppress_auto_hyphens,
                    shading,
                    text_outline,
                    text_fill,
                    text_shadow,
                    text_glow,
                );
                // Tab stops are additive: accumulate from ancestors, child overrides at same pos
                // Clear tabs remove inherited tabs at matching positions
                merge_tab_stops(
                    &mut inh.tab_stops,
                    &s.clear_tab_positions,
                    s.tab_stops.clone(),
                );
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
            s.underline = s.underline.or(inh.underline);
            s.double_underline = s.double_underline.or(inh.double_underline);
            s.strikethrough = s.strikethrough.or(inh.strikethrough);
            s.dstrike = s.dstrike.or(inh.dstrike);
            s.color = s.color.or(inh.color);
            s.char_spacing = s.char_spacing.or(inh.char_spacing);
            s.alignment = s.alignment.or(inh.alignment);
            s.space_before = s.space_before.or(inh.space_before);
            s.space_after = s.space_after.or(inh.space_after);
            s.space_before_autospacing = s.space_before_autospacing.or(inh.space_before_autospacing);
            s.space_after_autospacing = s.space_after_autospacing.or(inh.space_after_autospacing);
            s.line_spacing = s.line_spacing.or(inh.line_spacing);
            s.indent_left = s.indent_left.or(inh.indent_left);
            s.indent_right = s.indent_right.or(inh.indent_right);
            s.indent_hanging = s.indent_hanging.or(inh.indent_hanging);
            s.indent_first_line = s.indent_first_line.or(inh.indent_first_line);
            s.kern_threshold = s.kern_threshold.or(inh.kern_threshold);
            s.widow_control = s.widow_control.or(inh.widow_control);
            s.num_id = s.num_id.take().or(inh.num_id);
            s.num_ilvl = s.num_ilvl.or(inh.num_ilvl);
            s.outline_level = s.outline_level.or(inh.outline_level);
            s.snap_to_grid = s.snap_to_grid.or(inh.snap_to_grid);
            s.auto_space_de = s.auto_space_de.or(inh.auto_space_de);
            s.auto_space_dn = s.auto_space_dn.or(inh.auto_space_dn);
            s.suppress_auto_hyphens = s.suppress_auto_hyphens.or(inh.suppress_auto_hyphens);
            s.shading = s.shading.or(inh.shading);
            s.text_outline = s.text_outline.take().or(inh.text_outline);
            s.text_fill = s.text_fill.take().or(inh.text_fill);
            s.text_shadow = s.text_shadow.take().or(inh.text_shadow);
            s.text_glow = s.text_glow.take().or(inh.text_glow);
            s.tab_stops = inh.tab_stops;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_alignment() {
        assert_eq!(parse_alignment("center"), Alignment::Center);
        assert_eq!(parse_alignment("right"), Alignment::Right);
        assert_eq!(parse_alignment("end"), Alignment::Right);
        assert_eq!(parse_alignment("both"), Alignment::Justify);
        assert_eq!(parse_alignment("distribute"), Alignment::Distribute);
        // Word leaves Latin thaiDistribute at natural width (case77 reference).
        assert_eq!(parse_alignment("thaiDistribute"), Alignment::Justify);
        // Kashida elongation needs shaping we don't have; justify is the
        // closest renderable behavior.
        assert_eq!(parse_alignment("mediumKashida"), Alignment::Justify);
        assert_eq!(parse_alignment("highKashida"), Alignment::Justify);
        assert_eq!(parse_alignment("lowKashida"), Alignment::Justify);
        assert_eq!(parse_alignment("left"), Alignment::Left);
        assert_eq!(parse_alignment("start"), Alignment::Left); // unknown → Left
        assert_eq!(parse_alignment(""), Alignment::Left);
    }
}
