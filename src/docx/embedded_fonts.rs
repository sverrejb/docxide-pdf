use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{FontFamily, FontTable, FontTableEntry};

use super::relationships::parse_part_relationships;
use super::{REL_NS, WML_NS, read_zip_text, wml};

const EMBED_VARIANTS: &[(&str, bool, bool)] = &[
    ("embedRegular", false, false),
    ("embedBold", true, false),
    ("embedItalic", false, true),
    ("embedBoldItalic", true, true),
];

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

fn deobfuscate_font(data: &mut [u8], key: &[u8; 16]) {
    for (i, byte) in data.iter_mut().take(32).enumerate() {
        *byte ^= key[i % 16];
    }
}

struct EmbedInfo {
    font_name: String,
    bold: bool,
    italic: bool,
    rel_id: String,
    font_key: Option<String>,
}

pub(super) struct FontTableResult {
    pub embedded_fonts: HashMap<(String, bool, bool), Vec<u8>>,
    pub font_table: FontTable,
}

fn empty_result(font_table: FontTable) -> FontTableResult {
    FontTableResult {
        embedded_fonts: HashMap::new(),
        font_table,
    }
}

fn parse_font_family(val: &str) -> FontFamily {
    match val {
        "roman" => FontFamily::Roman,
        "swiss" => FontFamily::Swiss,
        "modern" => FontFamily::Modern,
        "script" => FontFamily::Script,
        "decorative" => FontFamily::Decorative,
        _ => FontFamily::Auto,
    }
}

pub(super) fn parse_font_table<R: Read + Seek>(zip: &mut zip::ZipArchive<R>) -> FontTableResult {
    let mut font_table = FontTable::new();

    let embeds = {
        let Some(xml_content) = read_zip_text(zip, "word/fontTable.xml") else {
            return empty_result(font_table);
        };
        let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
            return empty_result(font_table);
        };

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

            let alt_name = wml(font_node, "altName")
                .and_then(|n| n.attribute((WML_NS, "val")))
                .map(|s| s.to_string());
            let family = wml(font_node, "family")
                .and_then(|n| n.attribute((WML_NS, "val")))
                .map(parse_font_family)
                .unwrap_or(FontFamily::Auto);
            font_table.insert(font_name.to_string(), FontTableEntry { alt_name, family });

            for &(embed_tag, bold, italic) in EMBED_VARIANTS {
                let Some(embed_node) = wml(font_node, embed_tag) else {
                    continue;
                };
                let Some(r_id) = embed_node.attribute((REL_NS, "id")) else {
                    continue;
                };

                embeds.push(EmbedInfo {
                    font_name: font_name.to_string(),
                    bold,
                    italic,
                    rel_id: r_id.to_string(),
                    font_key: embed_node
                        .attribute((WML_NS, "fontKey"))
                        .map(|s| s.to_string()),
                });
            }
        }
        embeds
    };

    if embeds.is_empty() {
        return empty_result(font_table);
    }

    // Phase 2: resolve relationships and extract font data
    let font_rels = parse_part_relationships(zip, "word/fontTable.xml");
    let mut embedded_fonts = HashMap::new();

    for info in embeds {
        let Some(target) = font_rels.get(&info.rel_id) else {
            continue;
        };

        let zip_path = match target.strip_prefix('/') {
            Some(absolute) => absolute.to_string(),
            None => format!("word/{}", target),
        };

        let mut data = Vec::new();
        {
            let Ok(mut entry) = zip.by_name(&zip_path) else {
                continue;
            };
            if entry.read_to_end(&mut data).is_err() {
                continue;
            }
        }

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
        embedded_fonts.insert(
            (info.font_name.to_lowercase(), info.bold, info.italic),
            data,
        );
    }

    FontTableResult {
        embedded_fonts,
        font_table,
    }
}
