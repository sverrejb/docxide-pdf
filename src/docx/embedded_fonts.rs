use std::collections::HashMap;
use std::io::Read;

use super::{WML_NS, REL_NS, read_zip_text, wml};

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
    for i in 0..16.min(data.len()) {
        data[i] ^= key[i];
    }
    for i in 16..32.min(data.len()) {
        data[i] ^= key[i - 16];
    }
}

fn parse_font_table_rels<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> HashMap<String, String> {
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

pub(super) fn parse_font_table<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
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
        result.insert(
            (info.font_name.to_lowercase(), info.bold, info.italic),
            data,
        );
    }

    result
}
