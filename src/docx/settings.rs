use std::io::{Read, Seek};

use super::{WML_NS, read_zip_text, twips_to_pts, wml, wml_attr, wml_bool};

pub(super) struct DocumentSettings {
    pub even_and_odd_headers: bool,
    pub default_tab_stop: f32,
    #[allow(dead_code)]
    pub mirror_margins: bool,
    pub gutter_at_top: bool,
    pub east_asia_lang: Option<String>,
    pub auto_hyphenation: bool,
    pub default_lang: Option<String>,
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            even_and_odd_headers: false,
            default_tab_stop: 36.0, // 0.5 inches = 720 twips = 36pt
            mirror_margins: false,
            gutter_at_top: false,
            east_asia_lang: None,
            auto_hyphenation: false,
            default_lang: None,
        }
    }
}

pub(super) fn parse_settings<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> DocumentSettings {
    let Some(xml_text) = read_zip_text(zip, "word/settings.xml") else {
        return DocumentSettings::default();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml_text) else {
        return DocumentSettings::default();
    };
    let root = doc.root_element();

    let default_tab_stop = wml_attr(root, "defaultTabStop")
        .and_then(|v| v.parse::<f32>().ok())
        .map(twips_to_pts)
        .unwrap_or(36.0);

    let theme_font_lang = wml(root, "themeFontLang");
    let east_asia_lang = theme_font_lang
        .and_then(|n| n.attribute((WML_NS, "eastAsia")))
        .map(|s| s.to_string());
    let default_lang = theme_font_lang
        .and_then(|n| n.attribute((WML_NS, "val")))
        .map(|s| s.to_string());

    DocumentSettings {
        even_and_odd_headers: wml_bool(root, "evenAndOddHeaders").unwrap_or(false),
        default_tab_stop,
        mirror_margins: wml_bool(root, "mirrorMargins").unwrap_or(false),
        gutter_at_top: wml_bool(root, "gutterAtTop").unwrap_or(false),
        east_asia_lang,
        auto_hyphenation: wml_bool(root, "autoHyphenation").unwrap_or(false),
        default_lang,
    }
}
