//! The document theme (fonts, colour scheme, fill and line styles) and the rFonts theme slots.

use std::collections::HashMap;
use std::io::{Read, Seek};

use super::color::{ColorTransforms, parse_color_transforms};
use super::{DML_NS, WML_NS, angle_attr, dml, emu_attr, frac_attr, parse_hex_color, read_zip_text};

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

#[derive(Default)]
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
    // No themeFontLang bidi: Word's default complex-script language is Arabic
    // (a Chinese template's majorBidi runs print in the theme's Arab font,
    // Times New Roman, not in a missing font's substitute).
    let bidi_script = bidi_lang.and_then(lang_to_script).or(Some("Arab"));

    let names: Vec<String> = zip
        .file_names()
        .filter_map(|s| s.ok().map(|s| s.into_owned()))
        .collect();
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
