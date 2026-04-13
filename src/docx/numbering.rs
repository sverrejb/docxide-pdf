use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use super::{WML_NS, parse_hex_color, twips_attr, wml, wml_attr, wml_bool};

#[derive(Clone)]
pub(super) struct LevelDef {
    pub(super) num_fmt: String,
    pub(super) lvl_text: String,
    pub(super) indent_left: f32,
    pub(super) indent_hanging: f32,
    pub(super) tab_stop: Option<f32>,
    pub(super) start: u32,
    pub(super) bullet_font: Option<String>,
    pub(super) label_font_size: Option<f32>,
    pub(super) label_bold: bool,
    pub(super) label_color: Option<[u8; 3]>,
    pub(super) suff: String,
    pub(super) label_font: Option<String>,
}

#[derive(Default)]
pub(super) struct ListLabelInfo {
    pub(super) indent_left: f32,
    pub(super) indent_hanging: f32,
    pub(super) tab_stop: Option<f32>,
    pub(super) label: String,
    pub(super) font: Option<String>,
    pub(super) font_size: Option<f32>,
    pub(super) bold: bool,
    pub(super) color: Option<[u8; 3]>,
    pub(super) suff: String,
}

#[derive(Default)]
pub(super) struct NumberingInfo {
    pub(super) abstract_nums: HashMap<String, HashMap<u8, LevelDef>>,
    pub(super) num_to_abstract: HashMap<String, String>,
    pub(super) start_overrides: HashMap<String, HashMap<u8, u32>>,
}

pub(super) fn parse_numbering<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> NumberingInfo {
    let Some(xml_content) = super::read_zip_text(zip, "word/numbering.xml") else {
        return NumberingInfo::default();
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return NumberingInfo::default();
    };

    let mut abstract_nums = HashMap::new();
    let mut num_to_abstract = HashMap::new();
    let mut num_style_link = HashMap::new();
    let mut style_link_target = HashMap::new();
    let mut start_overrides = HashMap::new();

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
                let mut levels = HashMap::new();
                for lvl in node.children().filter(|n| n.has_tag_name((WML_NS, "lvl"))) {
                    let Some(ilvl) = lvl
                        .attribute((WML_NS, "ilvl"))
                        .and_then(|v| v.parse::<u8>().ok())
                    else {
                        continue;
                    };
                    let num_fmt = wml_attr(lvl, "numFmt").unwrap_or("bullet").to_string();
                    let lvl_text = wml_attr(lvl, "lvlText").unwrap_or("").to_string();
                    let start = wml_attr(lvl, "start")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1);
                    let ppr = wml(lvl, "pPr");
                    let ind = ppr.and_then(|p| wml(p, "ind"));
                    let indent_left = ind
                        .and_then(|n| twips_attr(n, "start").or_else(|| twips_attr(n, "left")))
                        .unwrap_or(0.0);
                    let indent_hanging = ind.and_then(|n| twips_attr(n, "hanging")).unwrap_or(0.0);
                    let tab_stop = ppr
                        .and_then(|p| wml(p, "tabs"))
                        .and_then(|tabs| {
                            tabs.children()
                                .filter(|n| n.has_tag_name((WML_NS, "tab")))
                                .find_map(|t| twips_attr(t, "pos"))
                        });
                    let rpr = wml(lvl, "rPr");
                    let rpr_font = rpr
                        .and_then(|r| wml(r, "rFonts"))
                        .and_then(|rf| {
                            rf.attribute((WML_NS, "ascii"))
                                .or_else(|| rf.attribute((WML_NS, "hAnsi")))
                        })
                        .map(|s| s.to_string());
                    let label_font_size = rpr
                        .and_then(|r| wml_attr(r, "sz"))
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|hp| hp / 2.0);
                    let label_bold = rpr.and_then(|r| wml_bool(r, "b")).unwrap_or(false);
                    let label_color = rpr
                        .and_then(|r| wml_attr(r, "color"))
                        .and_then(parse_hex_color);
                    let suff = wml_attr(lvl, "suff").unwrap_or("tab").to_string();
                    levels.insert(
                        ilvl,
                        LevelDef {
                            num_fmt,
                            lvl_text,
                            indent_left,
                            indent_hanging,
                            tab_stop,
                            start,
                            bullet_font: rpr_font.clone(),
                            label_font_size,
                            label_bold,
                            label_color,
                            suff,
                            label_font: rpr_font,
                        },
                    );
                }
                abstract_nums.insert(abs_id.to_string(), levels);
                if let Some(link) = wml_attr(node, "numStyleLink") {
                    num_style_link.insert(abs_id.to_string(), link.to_string());
                }
                if let Some(link) = wml_attr(node, "styleLink") {
                    style_link_target.insert(link.to_string(), abs_id.to_string());
                }
            }
            "num" => {
                let Some(num_id) = node.attribute((WML_NS, "numId")) else {
                    continue;
                };
                let Some(abs_id) = wml_attr(node, "abstractNumId") else {
                    continue;
                };
                num_to_abstract.insert(num_id.to_string(), abs_id.to_string());
                let overrides: HashMap<u8, u32> = node
                    .children()
                    .filter(|n| n.has_tag_name((WML_NS, "lvlOverride")))
                    .filter_map(|ovr| {
                        let ilvl = ovr
                            .attribute((WML_NS, "ilvl"))
                            .and_then(|v| v.parse::<u8>().ok())?;
                        let val =
                            wml_attr(ovr, "startOverride").and_then(|v| v.parse::<u32>().ok())?;
                        Some((ilvl, val))
                    })
                    .collect();
                if !overrides.is_empty() {
                    start_overrides.insert(num_id.to_string(), overrides);
                }
            }
            _ => {}
        }
    }

    // Resolve numStyleLink → styleLink chains
    for (abs_id, style_name) in &num_style_link {
        let Some(target_abs_id) = style_link_target.get(style_name) else {
            continue;
        };
        let Some(source_levels) = abstract_nums.get(target_abs_id).cloned() else {
            continue;
        };
        if source_levels.is_empty() {
            continue;
        }
        let entry = abstract_nums.entry(abs_id.clone()).or_default();
        if entry.is_empty() {
            *entry = source_levels;
        }
    }

    NumberingInfo {
        abstract_nums,
        num_to_abstract,
        start_overrides,
    }
}

fn to_roman(mut n: u32) -> String {
    const TABLE: &[(u32, &str)] = &[
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut result = String::new();
    for &(value, numeral) in TABLE {
        while n >= value {
            result.push_str(numeral);
            n -= value;
        }
    }
    result
}

fn to_letter(value: u32, base: u8) -> String {
    if value == 0 {
        return String::new();
    }
    let mut n = value - 1;
    let mut result = String::new();
    loop {
        result.insert(0, (base + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

pub(crate) fn format_number(value: u32, num_fmt: &str) -> String {
    match num_fmt {
        "decimal" => value.to_string(),
        "decimalZero" => format!("{value:02}"),
        "lowerLetter" => to_letter(value, b'a'),
        "upperLetter" => to_letter(value, b'A'),
        "lowerRoman" => to_roman(value),
        "upperRoman" => to_roman(value).to_uppercase(),
        "none" => String::new(),
        _ => value.to_string(),
    }
}

fn normalize_bullet_text(text: &str) -> String {
    text.chars()
        .map(|c| {
            let cp = c as u32;
            if (0xF000..=0xF0FF).contains(&cp) {
                symbol_pua_to_unicode(cp).unwrap_or(c)
            } else {
                c
            }
        })
        .collect()
}

fn symbol_pua_to_unicode(cp: u32) -> Option<char> {
    let sym = cp - 0xF000;
    let mapped = match sym {
        0xB7 => '\u{2022}', // bullet •
        0xA7 => '\u{25A0}', // black square ■ (Wingdings §)
        0xA8 => '\u{25CB}', // white circle ○
        0xD8 => '\u{2666}', // diamond ◆
        0x76 => '\u{221A}', // check mark √
        _ => return char::from_u32(sym),
    };
    Some(mapped)
}

pub(super) fn parse_list_info(
    num_pr: Option<roxmltree::Node>,
    style_num_id: Option<&str>,
    style_num_ilvl: Option<u8>,
    numbering: &NumberingInfo,
    counters: &mut HashMap<(u32, u8), u32>,
    last_seen_level: &mut HashMap<u32, u8>,
    applied_overrides: &mut HashSet<(u32, u8)>,
) -> ListLabelInfo {
    let (num_id, ilvl) = if let Some(np) = num_pr {
        let nid = wml_attr(np, "numId");
        let il = wml_attr(np, "ilvl")
            .and_then(|v| v.parse::<u8>().ok())
            .unwrap_or(0);
        (nid, il)
    } else if let Some(sn) = style_num_id {
        (Some(sn), style_num_ilvl.unwrap_or(0))
    } else {
        return ListLabelInfo::default();
    };
    let Some(num_id_str) = num_id else {
        return ListLabelInfo::default();
    };
    if num_id_str == "0" {
        return ListLabelInfo::default();
    }
    let Some(abs_id) = numbering.num_to_abstract.get(num_id_str) else {
        return ListLabelInfo::default();
    };
    let Some(levels) = numbering.abstract_nums.get(abs_id.as_str()) else {
        return ListLabelInfo::default();
    };
    let Some(def) = levels.get(&ilvl) else {
        return ListLabelInfo::default();
    };

    // Key counters by abstractNumId so all numIds sharing the same
    // abstract definition share a single counter stream.
    let abs_key: u32 = abs_id.parse().unwrap_or(0);

    // Reset deeper-level counters when returning to a higher level
    if let Some(&prev) = last_seen_level.get(&abs_key) {
        for deeper in (ilvl + 1)..=prev {
            counters.remove(&(abs_key, deeper));
        }
    }
    last_seen_level.insert(abs_key, ilvl);

    // Apply startOverride: when a numId carries a restart directive,
    // force-reset the shared counter on first encounter.
    let override_start = numbering
        .start_overrides
        .get(num_id_str)
        .and_then(|m| m.get(&ilvl))
        .copied();
    let num_key_original: u32 = num_id_str.parse().unwrap_or(0);
    if let Some(restart) = override_start {
        if applied_overrides.insert((num_key_original, ilvl)) {
            // First time seeing this numId+ilvl override — reset the
            // shared counter so the next increment produces restart.
            counters.insert((abs_key, ilvl), restart - 1);
        }
    }

    let start = override_start.unwrap_or(def.start);
    let current_counter = *counters
        .entry((abs_key, ilvl))
        .and_modify(|c| *c += 1)
        .or_insert(start);

    let is_bullet = def.num_fmt == "bullet";
    let original_had_pua = is_bullet
        && def
            .lvl_text
            .chars()
            .any(|c| (0xF000..=0xF0FF).contains(&(c as u32)));
    let label = if is_bullet {
        if original_had_pua && def.bullet_font.is_some() {
            // Keep PUA chars for symbol fonts — their cmaps expect PUA encoding
            def.lvl_text.clone()
        } else {
            let text = normalize_bullet_text(&def.lvl_text);
            if text.is_empty() {
                "\u{2022}".to_string()
            } else {
                text
            }
        }
    } else {
        let mut label = def.lvl_text.clone();
        for lvl_idx in 0..9u8 {
            let placeholder = format!("%{}", lvl_idx + 1);
            if label.contains(&placeholder) {
                let lvl_counter = if lvl_idx == ilvl {
                    current_counter
                } else {
                    counters
                        .get(&(abs_key, lvl_idx))
                        .copied()
                        .unwrap_or(levels.get(&lvl_idx).map(|d| d.start).unwrap_or(1))
                };
                let lvl_fmt = levels
                    .get(&lvl_idx)
                    .map(|d| d.num_fmt.as_str())
                    .unwrap_or("decimal");
                label = label.replace(&placeholder, &format_number(lvl_counter, lvl_fmt));
            }
        }
        label
    };
    ListLabelInfo {
        indent_left: def.indent_left,
        indent_hanging: def.indent_hanging,
        tab_stop: def.tab_stop,
        label,
        font: if is_bullet {
            def.bullet_font.clone()
        } else if def.suff == "nothing" {
            // suff=nothing: label flows inline with text and needs
            // its own font for the synthetic Run we prepend.
            def.label_font.clone()
        } else {
            None
        },
        font_size: def.label_font_size,
        bold: def.label_bold,
        color: def.label_color,
        suff: def.suff.clone(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // --- Decimal ---

    #[test]
    fn test_format_number_decimal() {
        assert_eq!(format_number(1, "decimal"), "1");
        assert_eq!(format_number(10, "decimal"), "10");
        assert_eq!(format_number(999, "decimal"), "999");
    }

    #[test]
    fn test_format_number_decimal_zero() {
        assert_eq!(format_number(1, "decimalZero"), "01");
        assert_eq!(format_number(9, "decimalZero"), "09");
        assert_eq!(format_number(10, "decimalZero"), "10");
        assert_eq!(format_number(99, "decimalZero"), "99");
    }

    // --- Letters ---

    #[test]
    fn test_format_number_lower_letter() {
        assert_eq!(format_number(1, "lowerLetter"), "a");
        assert_eq!(format_number(2, "lowerLetter"), "b");
        assert_eq!(format_number(26, "lowerLetter"), "z");
        assert_eq!(format_number(27, "lowerLetter"), "aa");
        assert_eq!(format_number(28, "lowerLetter"), "ab");
    }

    #[test]
    fn test_format_number_upper_letter() {
        assert_eq!(format_number(1, "upperLetter"), "A");
        assert_eq!(format_number(26, "upperLetter"), "Z");
        assert_eq!(format_number(27, "upperLetter"), "AA");
    }

    // --- Roman numerals ---

    #[test]
    fn test_format_number_lower_roman() {
        assert_eq!(format_number(1, "lowerRoman"), "i");
        assert_eq!(format_number(2, "lowerRoman"), "ii");
        assert_eq!(format_number(3, "lowerRoman"), "iii");
        assert_eq!(format_number(4, "lowerRoman"), "iv");
        assert_eq!(format_number(5, "lowerRoman"), "v");
        assert_eq!(format_number(9, "lowerRoman"), "ix");
        assert_eq!(format_number(10, "lowerRoman"), "x");
        assert_eq!(format_number(14, "lowerRoman"), "xiv");
        assert_eq!(format_number(40, "lowerRoman"), "xl");
        assert_eq!(format_number(99, "lowerRoman"), "xcix");
    }

    #[test]
    fn test_format_number_upper_roman() {
        assert_eq!(format_number(1, "upperRoman"), "I");
        assert_eq!(format_number(4, "upperRoman"), "IV");
        assert_eq!(format_number(14, "upperRoman"), "XIV");
        assert_eq!(format_number(1999, "upperRoman"), "MCMXCIX");
    }

    // --- Edge cases ---

    #[test]
    fn test_format_number_none() {
        assert_eq!(format_number(5, "none"), "");
    }

    #[test]
    fn test_format_number_unknown_falls_back_to_decimal() {
        assert_eq!(format_number(42, "unknownFormat"), "42");
    }

    #[test]
    fn test_to_letter_zero() {
        assert_eq!(to_letter(0, b'a'), "");
    }

    #[test]
    fn test_to_roman_large() {
        assert_eq!(to_roman(2024), "mmxxiv");
        assert_eq!(to_roman(3999), "mmmcmxcix");
    }
}
