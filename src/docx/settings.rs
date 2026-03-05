use std::io::Read;

use super::{WML_NS, read_zip_text, twips_to_pts};

pub(super) struct DocumentSettings {
    pub even_and_odd_headers: bool,
    pub default_tab_stop: f32,
    pub mirror_margins: bool,
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            even_and_odd_headers: false,
            default_tab_stop: 36.0, // 0.5 inches = 720 twips = 36pt
            mirror_margins: false,
        }
    }
}

pub(super) fn parse_settings<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> DocumentSettings {
    let Some(xml_text) = read_zip_text(zip, "word/settings.xml") else {
        return DocumentSettings::default();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml_text) else {
        return DocumentSettings::default();
    };
    let root = doc.root_element();

    let wml_bool_el = |name: &str| -> bool {
        root.children().any(|n| {
            n.tag_name().namespace() == Some(WML_NS)
                && n.tag_name().name() == name
                && n.attribute((WML_NS, "val"))
                    .is_none_or(|v| v != "0" && v != "false")
        })
    };

    let even_and_odd_headers = wml_bool_el("evenAndOddHeaders");
    let mirror_margins = wml_bool_el("mirrorMargins");

    let default_tab_stop = root
        .children()
        .find(|n| {
            n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "defaultTabStop"
        })
        .and_then(|n| n.attribute((WML_NS, "val")))
        .and_then(|v| v.parse::<f32>().ok())
        .map(twips_to_pts)
        .unwrap_or(36.0);

    DocumentSettings {
        even_and_odd_headers,
        default_tab_stop,
        mirror_margins,
    }
}
