use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use crate::model::{
    Alignment, CellBorder, LineSpacing, ParagraphBorders, TabStop, TextFill, TextGlow, TextOutline,
    TextShadow,
};

pub(super) use super::color::{ColorTransforms, parse_color_transforms};
use super::wordart::{parse_text_fill, parse_text_glow, parse_text_outline, parse_text_shadow};
use super::{
    DML_NS, WML_NS, angle_attr, dml, emu_attr, extract_indents, frac_attr, frame_attrs,
    highlight_color, merge_tab_stops, parse_cell_border, parse_cell_border_left,
    parse_cell_border_right, parse_hex_color, parse_on_off, parse_one_border,
    parse_paragraph_borders, parse_run_shd, parse_tab_stops_with_clears, parse_text_color,
    read_zip_text, twips_attr, twips_to_pts, wml, wml_attr, wml_bool,
};

fn dml_typeface<'a>(node: roxmltree::Node<'a, 'a>, element: &str) -> Option<&'a str> {
    dml(node, element)
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

/// A theme font group's `ea`/`cs` typeface; an empty one defers to the
/// group's font for `script`.
fn group_typeface<'a>(
    font_group: roxmltree::Node<'a, 'a>,
    element: &str,
    script: Option<&str>,
) -> String {
    dml_typeface(font_group, element)
        .or_else(|| script.and_then(|s| script_font_typeface(font_group, s)))
        .unwrap_or("")
        .to_string()
}

fn script_font_typeface<'a>(font_group: roxmltree::Node<'a, 'a>, script: &str) -> Option<&'a str> {
    font_group
        .children()
        .find(|n| n.has_tag_name((DML_NS, "font")) && n.attribute("script") == Some(script))
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

/// `w:lang`: the languages of Latin (`val`) and East Asian (`eastAsia`) text.
/// Only well-formed tags: Word writes "x-none" for "no language".
// ponytail: w:bidi (complex-script text) ignored until RTL runs are tagged
pub(super) fn parse_lang(rpr: roxmltree::Node) -> (Option<String>, Option<String>) {
    let Some(lang) = wml(rpr, "lang") else {
        return (None, None);
    };
    let tag = |attr| {
        lang.attribute((WML_NS, attr))
            .filter(|v| is_lang_tag(v))
            .map(str::to_string)
    };
    (tag("val"), tag("eastAsia"))
}

/// A BCP 47-shaped language tag: a 2–3 letter language, then alphanumeric subtags.
fn is_lang_tag(tag: &str) -> bool {
    let mut parts = tag.split('-');
    parts
        .next()
        .is_some_and(|p| (2..=3).contains(&p.len()) && p.chars().all(|c| c.is_ascii_alphabetic()))
        && parts.all(|p| (1..=8).contains(&p.len()) && p.chars().all(|c| c.is_ascii_alphanumeric()))
}

/// The theme `a:font @script` for a language.
// ponytail: CJK and right-to-left scripts only; add Indic/Thai languages when
// a document's themeFontLang bidi uses one
fn lang_to_script(lang: &str) -> Option<&'static str> {
    Some(match lang.split('-').next()? {
        "ja" => "Jpan",
        "zh" if lang.contains("TW") => "Hant",
        "zh" => "Hans",
        "ko" => "Hang",
        "ar" | "fa" | "ur" | "ps" | "sd" | "ug" => "Arab",
        "he" | "yi" => "Hebr",
        _ => return None,
    })
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
    /// Complex-script fonts (`majorBidi`/`minorBidi`).
    pub(super) major_cs: String,
    pub(super) minor_cs: String,
    pub(super) colors: HashMap<String, [u8; 3]>,
    pub(super) fill_styles: Vec<ThemeFillStyle>,
    /// `a:lnStyleLst` widths in points, which a shape's `lnRef idx` (1-3) picks.
    pub(super) line_widths: Vec<f32>,
}

impl ThemeFonts {
    /// The font a theme slot (`minorHAnsi`, `majorEastAsia`, `minorBidi`, …)
    /// names; None when unknown or empty.
    fn slot(&self, name: &str) -> Option<&str> {
        let font = match name {
            "majorHAnsi" => &self.major,
            "minorHAnsi" => &self.minor,
            "majorEastAsia" => &self.major_east_asia,
            "minorEastAsia" => &self.minor_east_asia,
            "majorBidi" => &self.major_cs,
            "minorBidi" => &self.minor_cs,
            _ => return None,
        };
        Some(font.as_str()).filter(|f| !f.is_empty())
    }
}

pub(super) struct StyleDefaults {
    pub(super) font_size: f32,
    pub(super) font_name: String,
    pub(super) east_asia_font: Option<String>,
    pub(super) cs_font: Option<String>,
    pub(super) space_before: f32,
    pub(super) space_after: f32,
    pub(super) before_autospacing: bool,
    pub(super) after_autospacing: bool,
    pub(super) line_spacing: LineSpacing,
    pub(super) kern_threshold: Option<f32>,
    pub(super) position: Option<f32>,
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
    pub(super) lang: Option<String>,
    pub(super) lang_east_asia: Option<String>,
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
    pub(super) cs_font: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) caps: Option<bool>,
    pub(super) small_caps: Option<bool>,
    pub(super) lang: Option<String>,
    pub(super) lang_east_asia: Option<String>,
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
    pub(super) contextual_spacing: Option<bool>,
    pub(super) keep_next: Option<bool>,
    pub(super) keep_lines: Option<bool>,
    pub(super) widow_control: Option<bool>,
    pub(super) page_break_before: Option<bool>,
    pub(super) line_spacing: Option<LineSpacing>,
    pub(super) indent_left: Option<f32>,
    pub(super) indent_right: Option<f32>,
    pub(super) indent_hanging: Option<f32>,
    pub(super) indent_first_line: Option<f32>,
    pub(super) borders: ParagraphBorders,
    pub(super) shading: Option<[u8; 3]>,
    pub(super) based_on: Option<String>,
    pub(super) kern_threshold: Option<f32>,
    pub(super) position: Option<f32>,
    pub(super) tab_stops: Vec<TabStop>,
    pub(super) clear_tab_positions: Vec<f32>,
    pub(super) num_id: Option<String>,
    pub(super) num_ilvl: Option<u8>,
    /// The style's indentation sits above its numbering's: walking up from
    /// this style, an own w:ind comes before (or with) an own w:numPr.
    pub(super) ind_over_numbering: bool,
    pub(super) outline_level: Option<u8>,
    pub(super) snap_to_grid: Option<bool>,
    pub(super) auto_space_de: Option<bool>,
    pub(super) auto_space_dn: Option<bool>,
    pub(super) text_outline: Option<TextOutline>,
    pub(super) text_fill: Option<TextFill>,
    pub(super) text_shadow: Option<TextShadow>,
    pub(super) text_glow: Option<TextGlow>,
    /// The style's framePr attributes; ponytail: basedOn passes the closest
    /// one on whole, merge per attribute if a chain ever splits them.
    pub(super) frame_attrs: Option<super::FrameAttrs>,
}

/// Run properties read from one `w:rPr`; every field is `None` when the
/// element doesn't set it. Shared by docDefaults, paragraph styles,
/// character styles (which are exactly this) and inline runs.
#[derive(Clone, Default)]
pub(super) struct RunProps {
    pub(super) font_size: Option<f32>,
    pub(super) font_name: Option<String>,
    pub(super) east_asia_font: Option<String>,
    pub(super) cs_font: Option<String>,
    pub(super) bold: Option<bool>,
    pub(super) italic: Option<bool>,
    pub(super) underline: Option<bool>,
    pub(super) double_underline: Option<bool>,
    pub(super) strikethrough: Option<bool>,
    pub(super) dstrike: Option<bool>,
    pub(super) caps: Option<bool>,
    pub(super) small_caps: Option<bool>,
    pub(super) lang: Option<String>,
    pub(super) lang_east_asia: Option<String>,
    pub(super) vanish: Option<bool>,
    pub(super) color: Option<[u8; 3]>,
    pub(super) highlight: Option<[u8; 3]>,
    /// Run-level shading from `w:rPr/w:shd`. `None` means "no explicit color
    /// here"; callers inherit from basedOn. Word treats `w:val="nil"` and
    /// `w:fill="auto"` as "no color," not as a clearing override.
    pub(super) shading: Option<[u8; 3]>,
    pub(super) border: Option<crate::model::ParagraphBorder>,
    pub(super) char_spacing: Option<f32>,
    pub(super) kern_threshold: Option<f32>,
    /// `w:position`: points the run is raised (negative: lowered).
    pub(super) position: Option<f32>,
    pub(super) text_outline: Option<TextOutline>,
    pub(super) text_fill: Option<TextFill>,
    pub(super) text_shadow: Option<TextShadow>,
    pub(super) text_glow: Option<TextGlow>,
}

pub(super) fn parse_run_props(rpr: roxmltree::Node, theme: &ThemeFonts) -> RunProps {
    let rfonts = wml(rpr, "rFonts");
    let (lang, lang_east_asia) = parse_lang(rpr);
    RunProps {
        font_size: parse_font_size(rpr),
        font_name: rfonts.and_then(|rf| resolve_font_from_node_opt(rf, theme)),
        east_asia_font: rfonts.and_then(|rf| resolve_east_asia_font_from_node(rf, theme)),
        cs_font: rfonts.and_then(|rf| resolve_cs_font_from_node(rf, theme)),
        bold: wml_bool(rpr, "b"),
        italic: wml_bool(rpr, "i"),
        underline: parse_underline(rpr),
        double_underline: parse_double_underline(rpr),
        strikethrough: wml_bool(rpr, "strike"),
        dstrike: wml_bool(rpr, "dstrike"),
        caps: wml_bool(rpr, "caps"),
        small_caps: wml_bool(rpr, "smallCaps"),
        lang,
        lang_east_asia,
        vanish: wml_bool(rpr, "vanish"),
        color: wml_attr(rpr, "color").and_then(parse_text_color),
        highlight: wml_attr(rpr, "highlight").and_then(highlight_color),
        shading: parse_run_shd(rpr),
        border: wml(rpr, "bdr").and_then(parse_one_border),
        char_spacing: parse_char_spacing(rpr),
        kern_threshold: parse_kern(rpr),
        position: half_points(rpr, "position"),
        text_outline: parse_text_outline(rpr, theme),
        text_fill: parse_text_fill(rpr, theme),
        text_shadow: parse_text_shadow(rpr, theme),
        text_glow: parse_text_glow(rpr, theme),
    }
}

impl RunProps {
    /// Takes every property this style leaves unset from its basedOn parent.
    fn inherit(&mut self, parent: &RunProps) {
        macro_rules! fill {
            ($($f:ident),+ $(,)?) => {{
                // Exhaustive on purpose: a new field fails to compile until it is listed here.
                let RunProps { $($f),+ } = parent;
                $(if self.$f.is_none() { self.$f = $f.clone(); })+
            }};
        }
        fill!(
            font_size,
            font_name,
            east_asia_font,
            cs_font,
            bold,
            italic,
            underline,
            double_underline,
            strikethrough,
            dstrike,
            caps,
            small_caps,
            lang,
            lang_east_asia,
            vanish,
            color,
            highlight,
            shading,
            border,
            char_spacing,
            kern_threshold,
            position,
            text_outline,
            text_fill,
            text_shadow,
            text_glow,
        );
    }
}

#[derive(Clone, Copy)]
pub(super) struct TableBordersDef {
    pub(super) top: CellBorder,
    pub(super) bottom: CellBorder,
    pub(super) left: CellBorder,
    pub(super) right: CellBorder,
    pub(super) inside_h: CellBorder,
    pub(super) inside_v: CellBorder,
}

/// A `w:tblBorders` or `w:tcBorders` set.
pub(super) fn parse_table_borders_def(bdr_node: roxmltree::Node) -> TableBordersDef {
    TableBordersDef {
        top: parse_cell_border(bdr_node, "top"),
        bottom: parse_cell_border(bdr_node, "bottom"),
        left: parse_cell_border_left(bdr_node),
        right: parse_cell_border_right(bdr_node),
        inside_h: parse_cell_border(bdr_node, "insideH"),
        inside_v: parse_cell_border(bdr_node, "insideV"),
    }
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
    /// `w:tblPr/w:tblCellMar`, top/left/bottom/right, unset sides `None`.
    pub(super) cell_margins: [Option<f32>; 4],
    /// The style's `w:pPr/w:spacing` before/after and line rule: cell
    /// paragraphs take them over docDefaults.
    pub(super) space_before: Option<f32>,
    pub(super) space_after: Option<f32>,
    pub(super) line_spacing: Option<LineSpacing>,
    pub(super) based_on: Option<String>,
}

pub(super) struct StylesInfo {
    pub(super) defaults: StyleDefaults,
    pub(super) paragraph_styles: HashMap<String, ParagraphStyle>,
    pub(super) character_styles: HashMap<String, RunProps>,
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

/// A half-point child value (`w:sz`, `w:kern`) in points.
pub(super) fn half_points(rpr: roxmltree::Node, name: &str) -> Option<f32> {
    wml_attr(rpr, name)
        .and_then(|v| v.parse::<f32>().ok())
        .map(|hp| hp / 2.0)
}

pub(super) fn parse_font_size(rpr: roxmltree::Node) -> Option<f32> {
    half_points(rpr, "sz")
}

fn parse_kern(rpr: roxmltree::Node) -> Option<f32> {
    half_points(rpr, "kern")
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
    wml(rpr, "spacing").and_then(|n| twips_attr(n, "val"))
}

pub(super) fn parse_theme<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    east_asia_lang: Option<&str>,
    bidi_lang: Option<&str>,
) -> ThemeFonts {
    let mut major = String::from("Aptos Display");
    let mut minor = String::from("Aptos");
    let mut major_east_asia = String::new();
    let mut minor_east_asia = String::new();
    let mut major_cs = String::new();
    let mut minor_cs = String::new();
    let mut colors = HashMap::new();
    let mut fill_styles = Vec::new();
    let mut line_widths = Vec::new();

    let script = east_asia_lang.and_then(lang_to_script).unwrap_or("Jpan");
    let bidi_script = bidi_lang.and_then(lang_to_script);

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
                    major_east_asia = group_typeface(node, "ea", Some(script));
                    major_cs = group_typeface(node, "cs", bidi_script);
                }
                "minorFont" => {
                    if let Some(tf) = dml_typeface(node, "latin") {
                        minor = tf.to_string();
                    }
                    minor_east_asia = group_typeface(node, "ea", Some(script));
                    minor_cs = group_typeface(node, "cs", bidi_script);
                }
                "clrScheme" => {
                    for child in node.children() {
                        if child.tag_name().namespace() != Some(DML_NS) {
                            continue;
                        }
                        let scheme_name = child.tag_name().name();
                        if let Some(srgb) = dml(child, "srgbClr") {
                            if let Some(hex) = srgb.attribute("val").and_then(parse_hex_color) {
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
                                        .and_then(|lin| angle_attr(lin, "ang"))
                                        .unwrap_or(0.0);
                                    fill_styles.push(ThemeFillStyle::Gradient { stops, angle_deg });
                                }
                            }
                            _ => {}
                        }
                    }
                }
                "lnStyleLst" => {
                    line_widths = node
                        .children()
                        .filter(|n| n.has_tag_name((DML_NS, "ln")))
                        .map(|ln| emu_attr(ln, "w"))
                        .collect();
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
        major_cs,
        minor_cs,
        colors,
        fill_styles,
        line_widths,
    }
}

fn parse_theme_gradient_stops(gs_lst: roxmltree::Node) -> Vec<ThemeGradientStop> {
    gs_lst
        .children()
        .filter(|n| n.has_tag_name((DML_NS, "gs")))
        .map(|gs| {
            let position = frac_attr(gs, "pos").unwrap_or(0.0);
            let transforms = gs
                .descendants()
                .find(|n| n.has_tag_name((DML_NS, "schemeClr")))
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
    ascii_theme
        .and_then(|t| theme.slot(t))
        .unwrap_or(default_font)
        .to_string()
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
    // ponytail: East Asian slots only; eastAsiaTheme="minorHAnsi" (997 runs in
    // the corpus) still falls through to w:eastAsia
    let from_theme = east_asia_theme
        .filter(|t| t.ends_with("EastAsia"))
        .and_then(|t| theme.slot(t))
        .map(str::to_string);
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

/// `w:cs`/`w:cstheme`, the complex-script font. As for East Asian text, only
/// the Bidi theme slots resolve.
pub(super) fn resolve_cs_font_from_node(
    rfonts: roxmltree::Node,
    theme: &ThemeFonts,
) -> Option<String> {
    rfonts
        .attribute((WML_NS, "cstheme"))
        .filter(|t| t.ends_with("Bidi"))
        .and_then(|t| theme.slot(t))
        .map(str::to_string)
        .or_else(|| {
            rfonts
                .attribute((WML_NS, "cs"))
                .filter(|s| !s.is_empty())
                .map(str::to_string)
        })
}

/// `w:spacing @line/@lineRule`, or None when `@line` is absent.
pub(super) fn parse_line_spacing(spacing_node: roxmltree::Node) -> Option<LineSpacing> {
    let line_val = spacing_node
        .attribute((WML_NS, "line"))?
        .parse::<f32>()
        .ok()?;
    Some(match spacing_node.attribute((WML_NS, "lineRule")) {
        Some("exact") => LineSpacing::Exact(twips_to_pts(line_val)),
        Some("atLeast") => LineSpacing::AtLeast(twips_to_pts(line_val)),
        _ => LineSpacing::Auto(line_val / 240.0),
    })
}

/// Word gives its built-in "heading N" styles outline level N−1 even when the
/// style omits `w:outlineLvl` (it tags and bookmarks them as headings).
fn builtin_heading_level(name: &str) -> Option<u8> {
    let level = name
        .to_ascii_lowercase()
        .strip_prefix("heading ")?
        .parse::<u8>()
        .ok()?;
    (1..=9).contains(&level).then(|| level - 1)
}

/// The stock Normal.dotm's docDefaults: theme fonts, 12pt, kerning from 1pt,
/// 8pt after, 278 auto lines.
const NORMAL_TEMPLATE_DEFAULTS: &str = concat!(
    r#"<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" "#,
    r#"w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>"#,
    r#"<w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>"#,
    r#"<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/>"#,
    r#"</w:pPr></w:pPrDefault></w:docDefaults>"#,
);

/// styles.xml after Word refreshed it from Normal.dotm: the template's
/// docDefaults and a default paragraph style with no formatting of its own
/// (Word then sets 12pt text under an 11pt Normal and steps 278 lines under
/// a 259 Normal). The template's other styles are Word's
/// built-in ones, which the file already carries.
fn with_normal_template(xml: &str) -> String {
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return xml.to_string();
    };
    let root = doc.root_element();
    let defaults = template_defaults_for(root);
    let mut edits: Vec<(std::ops::Range<usize>, String)> = Vec::new();
    match wml(root, "docDefaults") {
        Some(n) => edits.push((n.range(), defaults)),
        None => edits.push((doc_defaults_insertion_point(root), defaults)),
    }
    if let Some(normal) = root.children().find(|n| {
        n.tag_name().name() == "style"
            && n.attribute((WML_NS, "type")) == Some("paragraph")
            && n.attribute((WML_NS, "default"))
                .is_some_and(super::parse_on_off)
    }) {
        for pr in ["pPr", "rPr"]
            .into_iter()
            .filter_map(|name| wml(normal, name))
        {
            edits.push((pr.range(), String::new()));
        }
    }
    edits.sort_by_key(|(r, _)| std::cmp::Reverse(r.start));
    let mut out = xml.to_string();
    for (range, text) in edits {
        out.replace_range(range, &text);
    }
    out
}

/// NORMAL_TEMPLATE_DEFAULTS written with the styles part's own WML prefix.
fn template_defaults_for(root: roxmltree::Node) -> String {
    let prefix = match root.lookup_prefix(WML_NS) {
        Some("") | None => String::new(),
        Some(p) => format!("{p}:"),
    };
    NORMAL_TEMPLATE_DEFAULTS.replace("w:", &prefix)
}

fn doc_defaults_insertion_point(root: roxmltree::Node) -> std::ops::Range<usize> {
    let at = root
        .first_child()
        .map_or(root.range().end, |c| c.range().start);
    at..at
}

/// Without docDefaults (or without a styles part at all) Word lays the
/// document out with Normal.dotm's: 12pt theme minor font, 8pt after, 278
/// auto lines (a bare 16pt Calibri Light paragraph steps 30.48 = 19.5 ×
/// 1.158 + 8, not 19.5).
fn with_template_doc_defaults(xml: Option<String>) -> String {
    let Some(xml) = xml else {
        return format!(r#"<w:styles xmlns:w="{WML_NS}">{NORMAL_TEMPLATE_DEFAULTS}</w:styles>"#);
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return xml;
    };
    let root = doc.root_element();
    if wml(root, "docDefaults").is_some() {
        return xml;
    }
    let mut out = xml.clone();
    out.replace_range(
        doc_defaults_insertion_point(root),
        &template_defaults_for(root),
    );
    out
}

pub(super) fn parse_styles<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    theme: &ThemeFonts,
    from_normal_template: bool,
) -> StylesInfo {
    let mut defaults = StyleDefaults {
        font_size: 10.0,
        font_name: theme.minor.clone(),
        east_asia_font: None,
        cs_font: None,
        space_before: 0.0,
        space_after: 0.0,
        before_autospacing: false,
        after_autospacing: false,
        line_spacing: LineSpacing::Auto(1.0),
        kern_threshold: None,
        position: None,
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
        lang: None,
        lang_east_asia: None,
        widow_control: true,
        indent_left: 0.0,
        indent_right: 0.0,
        indent_hanging: 0.0,
        indent_first_line: 0.0,
        alignment: Alignment::Left,
    };
    let mut paragraph_styles = HashMap::new();
    let mut character_styles = HashMap::new();
    let mut character_parents = HashMap::new();
    let mut style_id_to_name = HashMap::new();
    let mut default_paragraph_style_id = String::from("Normal");

    let xml_content = read_zip_text(zip, "word/styles.xml").map(|xml| {
        if from_normal_template {
            with_normal_template(&xml)
        } else {
            xml
        }
    });
    let xml_content = with_template_doc_defaults(xml_content);
    let Some(xml) = roxmltree::Document::parse(&xml_content).ok() else {
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
            let r = parse_run_props(rpr, theme);
            defaults.font_size = r.font_size.unwrap_or(defaults.font_size);
            defaults.font_name = r.font_name.unwrap_or_else(|| theme.minor.clone());
            defaults.east_asia_font = r.east_asia_font;
            defaults.cs_font = r.cs_font;
            defaults.kern_threshold = r.kern_threshold;
            defaults.position = r.position;
            defaults.bold = r.bold.unwrap_or(false);
            defaults.italic = r.italic.unwrap_or(false);
            defaults.caps = r.caps.unwrap_or(false);
            defaults.small_caps = r.small_caps.unwrap_or(false);
            defaults.vanish = r.vanish.unwrap_or(false);
            defaults.strikethrough = r.strikethrough.unwrap_or(false);
            defaults.dstrike = r.dstrike.unwrap_or(false);
            defaults.underline = r.underline.unwrap_or(false);
            defaults.double_underline = r.double_underline.unwrap_or(false);
            defaults.color = r.color;
            defaults.char_spacing = r.char_spacing.unwrap_or(0.0);
            (defaults.lang, defaults.lang_east_asia) = (r.lang, r.lang_east_asia);
        }
        // A docDefaults with no pPrDefault at all (PHPWord writes these) takes
        // Word's built-in paragraph defaults, 8pt after and line 278 auto:
        // 12pt Arial steps 24.0 and 10pt 21.36, i.e.
        // 13.79 × 1.158 + 8 and 11.49 × 1.158 + 8. An empty pPrDefault keeps
        // the OOXML defaults (single, nothing after).
        if wml(doc_defaults, "pPrDefault").is_none() {
            defaults.space_after = 8.0;
            defaults.line_spacing = LineSpacing::Auto(278.0 / 240.0);
        }
        let default_ppr = wml(doc_defaults, "pPrDefault").and_then(|n| wml(n, "pPr"));
        if let Some(wc) = default_ppr.and_then(|ppr| wml_bool(ppr, "widowControl")) {
            defaults.widow_control = wc;
        }
        let default_spacing = default_ppr.and_then(|n| wml(n, "spacing"));
        if let Some(spacing) = default_spacing {
            if let Some(before_val) = twips_attr(spacing, "before") {
                defaults.space_before = before_val;
            }
            if let Some(after_val) = twips_attr(spacing, "after") {
                defaults.space_after = after_val;
            }
            let auto = |attr| spacing.attribute((WML_NS, attr)).is_some_and(parse_on_off);
            defaults.before_autospacing = auto("beforeAutospacing");
            defaults.after_autospacing = auto("afterAutospacing");
            if let Some(ls) = parse_line_spacing(spacing) {
                defaults.line_spacing = ls;
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
        if !style_node.has_tag_name((WML_NS, "style")) {
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
                if style_node
                    .attribute((WML_NS, "default"))
                    .is_some_and(super::parse_on_off)
                {
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
                let shading = ppr.and_then(|n| wml(n, "shd")).and_then(super::shd_color);

                let RunProps {
                    font_size,
                    font_name,
                    east_asia_font,
                    cs_font,
                    bold,
                    italic,
                    caps,
                    small_caps,
                    lang,
                    lang_east_asia,
                    vanish,
                    underline,
                    double_underline,
                    strikethrough,
                    dstrike,
                    char_spacing,
                    kern_threshold,
                    position,
                    color,
                    text_outline,
                    text_fill,
                    text_shadow,
                    text_glow,
                    ..
                } = wml(style_node, "rPr")
                    .map(|rpr| parse_run_props(rpr, theme))
                    .unwrap_or_default();

                let alignment = ppr.and_then(|ppr| wml_attr(ppr, "jc")).map(parse_alignment);

                // Kept as Option so a basedOn chain inherits them and an explicit
                // w:val="0" still turns an ancestor's off.
                let contextual_spacing = ppr.and_then(|ppr| wml_bool(ppr, "contextualSpacing"));
                let keep_next = ppr.and_then(|ppr| wml_bool(ppr, "keepNext"));
                let keep_lines = ppr.and_then(|ppr| wml_bool(ppr, "keepLines"));
                let widow_control = ppr.and_then(|ppr| wml_bool(ppr, "widowControl"));
                let page_break_before = ppr.and_then(|ppr| wml_bool(ppr, "pageBreakBefore"));

                let line_spacing = spacing.and_then(parse_line_spacing);

                let (indent_left, indent_right, indent_hanging, indent_first_line) = ppr
                    .and_then(|n| wml(n, "ind"))
                    .map(|ind| extract_indents(ind, None))
                    .unwrap_or_default();

                let (tab_stops, clear_tab_positions) =
                    ppr.map(parse_tab_stops_with_clears).unwrap_or_default();

                let style_num_pr = ppr.and_then(|p| wml(p, "numPr"));
                let num_id = style_num_pr
                    .and_then(|np| wml_attr(np, "numId"))
                    .map(|s| s.to_string());
                let num_ilvl = style_num_pr
                    .and_then(|np| wml_attr(np, "ilvl"))
                    .and_then(|v| v.parse::<u8>().ok());

                // 9 is kept: it is explicit body text and must stop a basedOn
                // heading's level (TOC Heading is basedOn Heading 1 with level 9).
                let outline_level = ppr
                    .and_then(|p| wml_attr(p, "outlineLvl"))
                    .and_then(|v| v.parse::<u8>().ok())
                    .filter(|&lvl| lvl <= 9)
                    .or_else(|| builtin_heading_level(wml_attr(style_node, "name")?));

                let snap_to_grid = ppr.and_then(|ppr| wml_bool(ppr, "snapToGrid"));
                let auto_space_de = ppr.and_then(|ppr| wml_bool(ppr, "autoSpaceDE"));
                let auto_space_dn = ppr.and_then(|ppr| wml_bool(ppr, "autoSpaceDN"));
                let frame_attrs = ppr.and_then(|ppr| wml(ppr, "framePr")).map(frame_attrs);

                let based_on = wml_attr(style_node, "basedOn").map(|s| s.to_string());

                paragraph_styles.insert(
                    style_id.to_string(),
                    ParagraphStyle {
                        font_size,
                        font_name,
                        east_asia_font,
                        cs_font,
                        bold,
                        italic,
                        caps,
                        small_caps,
                        lang,
                        lang_east_asia,
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
                        position,
                        tab_stops,
                        clear_tab_positions,
                        num_id,
                        num_ilvl,
                        ind_over_numbering: false,
                        outline_level,
                        snap_to_grid,
                        auto_space_de,
                        auto_space_dn,
                        text_outline,
                        text_fill,
                        text_shadow,
                        text_glow,
                        frame_attrs,
                    },
                );
            }
            Some("character") => {
                let props = wml(style_node, "rPr")
                    .map(|rpr| parse_run_props(rpr, theme))
                    .unwrap_or_default();
                character_styles.insert(style_id.to_string(), props);
                if let Some(parent) = wml_attr(style_node, "basedOn") {
                    character_parents.insert(style_id.to_string(), parent.to_string());
                }
            }
            Some("table") => {
                let base_borders = wml(style_node, "tblPr")
                    .and_then(|pr| wml(pr, "tblBorders"))
                    .map(parse_table_borders_def);

                // Parse base rPr from the table style
                let base_rpr = wml(style_node, "rPr");
                let base_font_size = base_rpr.and_then(parse_font_size);
                let base_font_name = base_rpr.and_then(rfonts_ascii_name);
                let base_bold = base_rpr.and_then(|rpr| wml_bool(rpr, "b"));
                let base_italic = base_rpr.and_then(|rpr| wml_bool(rpr, "i"));

                let cell_mar = wml(style_node, "tblPr").and_then(|pr| wml(pr, "tblCellMar"));
                let side =
                    |a: &str, b: &str| cell_mar.and_then(|m| super::tables::margin_twips(m, a, b));
                let cell_margins = [
                    side("top", "top"),
                    side("left", "start"),
                    side("bottom", "bottom"),
                    side("right", "end"),
                ];
                let style_spacing = wml(style_node, "pPr").and_then(|p| wml(p, "spacing"));
                let style_line_spacing = style_spacing.and_then(parse_line_spacing);

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
                        .map(parse_table_borders_def);
                    let cond_shading = wml(child, "tcPr")
                        .and_then(|tc| wml(tc, "shd"))
                        .and_then(super::shd_color);
                    let cond_rpr = wml(child, "rPr");
                    let cond_bold = cond_rpr.and_then(|rpr| wml_bool(rpr, "b"));
                    let cond_italic = cond_rpr.and_then(|rpr| wml_bool(rpr, "i"));
                    let cond_color = cond_rpr
                        .and_then(|rpr| wml_attr(rpr, "color"))
                        .and_then(parse_text_color);
                    let cond_font_size = cond_rpr.and_then(parse_font_size);
                    let cond_font_name = cond_rpr.and_then(rfonts_ascii_name);
                    if cond_borders.is_some()
                        || cond_shading.is_some()
                        || cond_bold.is_some()
                        || cond_color.is_some()
                        || cond_font_size.is_some()
                        || cond_font_name.is_some()
                        || cond_italic.is_some()
                    {
                        conditionals.insert(
                            cond_type.to_string(),
                            TableConditionalFormat {
                                borders: cond_borders,
                                shading: cond_shading,
                                bold: cond_bold,
                                italic: cond_italic,
                                color: cond_color,
                                font_size: cond_font_size,
                                font_name: cond_font_name,
                            },
                        );
                    }
                }

                table_styles.insert(
                    style_id.to_string(),
                    TableStyleDef {
                        base_borders,
                        base_font_size,
                        base_font_name,
                        base_bold,
                        base_italic,
                        conditionals,
                        cell_margins,
                        space_before: style_spacing.and_then(|n| twips_attr(n, "before")),
                        space_after: style_spacing.and_then(|n| twips_attr(n, "after")),
                        line_spacing: style_line_spacing,
                        based_on: wml_attr(style_node, "basedOn").map(str::to_string),
                    },
                );
            }
            _ => {}
        }
    }

    resolve_based_on(&mut paragraph_styles);
    resolve_character_based_on(&mut character_styles, &character_parents);

    // The default paragraph style (w:default="1") may carry properties like w:kern
    // that aren't in docDefaults. Merge kern_threshold into defaults if missing.
    if defaults.kern_threshold.is_none()
        && let Some(default_para) = paragraph_styles.get(&default_paragraph_style_id)
    {
        defaults.kern_threshold = default_para.kern_threshold;
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

fn resolve_character_based_on(
    styles: &mut HashMap<String, RunProps>,
    parents: &HashMap<String, String>,
) {
    let resolved: Vec<(String, RunProps)> = parents
        .keys()
        .filter_map(|id| {
            let mut props = styles.get(id)?.clone();
            let mut seen = HashSet::from([id.as_str()]);
            let mut current = id.as_str();
            while let Some(parent) = parents.get(current).map(String::as_str) {
                if !seen.insert(parent) {
                    break;
                }
                if let Some(p) = styles.get(parent) {
                    props.inherit(p);
                }
                current = parent;
            }
            Some((id.clone(), props))
        })
        .collect();
    styles.extend(resolved);
}

fn resolve_based_on(styles: &mut HashMap<String, ParagraphStyle>) {
    let ids: Vec<String> = styles.keys().cloned().collect();
    // Some(true) = own w:ind, Some(false) = own w:numPr only.
    let own_ind_or_num: HashMap<String, Option<bool>> = styles
        .iter()
        .map(|(id, s)| {
            let ind = s.indent_left.is_some()
                || s.indent_hanging.is_some()
                || s.indent_first_line.is_some();
            let num = s.num_id.is_some() || s.num_ilvl.is_some();
            (
                id.clone(),
                if ind {
                    Some(true)
                } else if num {
                    Some(false)
                } else {
                    None
                },
            )
        })
        .collect();
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
            (@fields $dst:expr, $src:expr, $($field:ident),+ $(,)?) => {
                $(if $src.$field.is_some() { $dst.$field = $src.$field.clone(); })+
            };
            ($dst:expr, $src:expr) => {
                inherit!(
                    @fields
                    $dst,
                    $src,
                    font_name,
                    east_asia_font,
                    cs_font,
                    font_size,
                    bold,
                    italic,
                    caps,
                    small_caps,
                    lang,
                    lang_east_asia,
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
                    position,
                    widow_control,
                    contextual_spacing,
                    keep_next,
                    keep_lines,
                    page_break_before,
                    num_id,
                    num_ilvl,
                    outline_level,
                    snap_to_grid,
                    auto_space_de,
                    auto_space_dn,
                    shading,
                    text_outline,
                    text_fill,
                    text_shadow,
                    text_glow,
                    frame_attrs,
                )
            };
        }

        let mut inh = ParagraphStyle::default();

        for ancestor_id in chain.iter().rev() {
            if let Some(s) = styles.get(ancestor_id) {
                inherit!(inh, s);
                // firstLine and hanging are one value: a style setting either
                // replaces both of its parent's.
                if s.indent_hanging.is_some() || s.indent_first_line.is_some() {
                    inh.indent_hanging = s.indent_hanging;
                    inh.indent_first_line = s.indent_first_line;
                }
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

        let ind_over_numbering = chain
            .iter()
            .find_map(|c| own_ind_or_num.get(c).copied().flatten())
            .unwrap_or(false);
        if let Some(s) = styles.get_mut(&id) {
            s.ind_over_numbering = ind_over_numbering;
            // The chain starts with the style itself, so `inh` already holds
            // its own values: one field list serves both directions.
            inherit!(s, inh);
            s.tab_stops = inh.tab_stops;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn missing_doc_defaults_take_the_template_ones() {
        let has_template_defaults = |xml: &str| {
            let doc = roxmltree::Document::parse(xml).unwrap();
            let spacing = wml(doc.root_element(), "docDefaults")
                .and_then(|d| wml(d, "pPrDefault"))
                .and_then(|p| wml(p, "pPr"))
                .and_then(|p| wml(p, "spacing"));
            spacing.and_then(|s| s.attribute((WML_NS, "after"))) == Some("160")
        };
        assert!(has_template_defaults(&with_template_doc_defaults(None)));
        let bare =
            format!(r#"<x:styles xmlns:x="{WML_NS}"><x:style x:styleId="Normal"/></x:styles>"#);
        assert!(has_template_defaults(&with_template_doc_defaults(Some(
            bare
        ))));
        let own = format!(r#"<w:styles xmlns:w="{WML_NS}"><w:docDefaults/></w:styles>"#);
        assert_eq!(with_template_doc_defaults(Some(own.clone())), own);
    }

    #[test]
    fn normal_template_replaces_defaults_and_empties_normal() {
        let xml = format!(
            r#"<w:styles xmlns:w="{WML_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="259"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:rPr><w:sz w:val="56"/></w:rPr></w:style></w:styles>"#
        );
        let out = with_normal_template(&xml);
        assert!(roxmltree::Document::parse(&out).is_ok());
        assert!(out.contains(r#"<w:sz w:val="24"/>"#) && out.contains(r#"w:line="278""#));
        assert!(!out.contains(r#"w:val="20""#) && !out.contains(r#"w:val="22""#));
        assert!(!out.contains(r#"w:line="259""#) && out.contains(r#"<w:sz w:val="56"/>"#));
    }

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

    #[test]
    fn builtin_heading_styles_have_outline_levels() {
        assert_eq!(builtin_heading_level("heading 1"), Some(0));
        assert_eq!(builtin_heading_level("Heading 9"), Some(8));
        assert_eq!(builtin_heading_level("heading 10"), None);
        assert_eq!(builtin_heading_level("Heading"), None);
        assert_eq!(builtin_heading_level("TOC Heading"), None);
    }

    #[test]
    fn east_asian_and_bidi_theme_slots_resolve() {
        let theme = ThemeFonts {
            major: "Calibri Light".into(),
            minor: "Calibri".into(),
            major_east_asia: "MS Gothic".into(),
            minor_east_asia: "맑은 고딕".into(),
            major_cs: "Times New Roman".into(),
            minor_cs: "Arial".into(),
            colors: HashMap::new(),
            fill_styles: Vec::new(),
            line_widths: Vec::new(),
        };
        let font = |slot| resolve_font(None, Some(slot), &theme, "");
        assert_eq!(font("minorEastAsia"), "맑은 고딕");
        assert_eq!(font("majorEastAsia"), "MS Gothic");
        assert_eq!(font("minorBidi"), "Arial");
        assert_eq!(font("majorBidi"), "Times New Roman");
        assert_eq!(
            resolve_font(
                None,
                Some("minorBidi"),
                &ThemeFonts {
                    minor_cs: String::new(),
                    ..theme
                },
                "Calibri"
            ),
            "Calibri"
        );
    }

    #[test]
    fn languages_map_to_theme_scripts() {
        assert_eq!(lang_to_script("ja-JP"), Some("Jpan"));
        assert_eq!(lang_to_script("zh-TW"), Some("Hant"));
        assert_eq!(lang_to_script("zh-CN"), Some("Hans"));
        assert_eq!(lang_to_script("ar-SA"), Some("Arab"));
        assert_eq!(lang_to_script("he-IL"), Some("Hebr"));
        assert_eq!(lang_to_script("th-TH"), None);
    }

    #[test]
    fn based_on_styles_inherit_language() {
        let mut styles = HashMap::new();
        styles.insert(
            "Normal".to_string(),
            ParagraphStyle {
                lang: Some("lt-LT".into()),
                lang_east_asia: Some("ja-JP".into()),
                ..Default::default()
            },
        );
        styles.insert(
            "BodyText3".to_string(),
            ParagraphStyle {
                based_on: Some("Normal".into()),
                ..Default::default()
            },
        );
        resolve_based_on(&mut styles);
        assert_eq!(styles["BodyText3"].lang.as_deref(), Some("lt-LT"));
        assert_eq!(styles["BodyText3"].lang_east_asia.as_deref(), Some("ja-JP"));
    }

    #[test]
    fn character_styles_inherit_through_based_on_chain() {
        let mut styles = HashMap::from([
            (
                "Base".to_string(),
                RunProps {
                    font_name: Some("Georgia".into()),
                    bold: Some(true),
                    font_size: Some(14.0),
                    ..Default::default()
                },
            ),
            (
                "Mid".to_string(),
                RunProps {
                    italic: Some(true),
                    font_size: Some(12.0),
                    ..Default::default()
                },
            ),
            (
                "Leaf".to_string(),
                RunProps {
                    underline: Some(true),
                    ..Default::default()
                },
            ),
        ]);
        let parents = HashMap::from([
            ("Mid".to_string(), "Base".to_string()),
            ("Leaf".to_string(), "Mid".to_string()),
            ("Base".to_string(), "Leaf".to_string()), // a cycle must not hang
        ]);
        resolve_character_based_on(&mut styles, &parents);
        let leaf = &styles["Leaf"];
        assert_eq!(leaf.font_name.as_deref(), Some("Georgia"));
        assert_eq!(
            (leaf.bold, leaf.italic, leaf.underline),
            (Some(true), Some(true), Some(true))
        );
        assert_eq!(leaf.font_size, Some(12.0), "the closer ancestor wins");
    }
}
