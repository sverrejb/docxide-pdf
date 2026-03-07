use std::collections::HashMap;

use super::{WML_NS, twips_attr, wml, wml_attr};

pub(super) struct LevelDef {
    pub(super) num_fmt: String,
    pub(super) lvl_text: String,
    pub(super) indent_left: f32,
    pub(super) indent_hanging: f32,
    pub(super) start: u32,
    pub(super) bullet_font: Option<String>,
}

#[derive(Default)]
pub(super) struct NumberingInfo {
    pub(super) abstract_nums: HashMap<String, HashMap<u8, LevelDef>>,
    pub(super) num_to_abstract: HashMap<String, String>,
}

pub(super) fn parse_numbering<R: std::io::Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> NumberingInfo {
    let Some(xml_content) = super::read_zip_text(zip, "word/numbering.xml") else {
        return NumberingInfo::default();
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return NumberingInfo::default();
    };

    let mut abstract_nums: HashMap<String, HashMap<u8, LevelDef>> = HashMap::new();
    let mut num_to_abstract: HashMap<String, String> = HashMap::new();

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
                let mut levels: HashMap<u8, LevelDef> = HashMap::new();
                for lvl in node.children() {
                    if lvl.tag_name().name() != "lvl" || lvl.tag_name().namespace() != Some(WML_NS)
                    {
                        continue;
                    }
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
                    let ind = wml(lvl, "pPr").and_then(|ppr| wml(ppr, "ind"));
                    let indent_left = ind.and_then(|n| twips_attr(n, "left")).unwrap_or(0.0);
                    let indent_hanging = ind.and_then(|n| twips_attr(n, "hanging")).unwrap_or(0.0);
                    let bullet_font = wml(lvl, "rPr")
                        .and_then(|rpr| wml(rpr, "rFonts"))
                        .and_then(|rf| {
                            rf.attribute((WML_NS, "ascii"))
                                .or_else(|| rf.attribute((WML_NS, "hAnsi")))
                        })
                        .map(|s| s.to_string());
                    levels.insert(
                        ilvl,
                        LevelDef {
                            num_fmt,
                            lvl_text,
                            indent_left,
                            indent_hanging,
                            start,
                            bullet_font,
                        },
                    );
                }
                abstract_nums.insert(abs_id.to_string(), levels);
            }
            "num" => {
                let Some(num_id) = node.attribute((WML_NS, "numId")) else {
                    continue;
                };
                let Some(abs_id) = wml_attr(node, "abstractNumId") else {
                    continue;
                };
                num_to_abstract.insert(num_id.to_string(), abs_id.to_string());
            }
            _ => {}
        }
    }

    NumberingInfo {
        abstract_nums,
        num_to_abstract,
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

fn format_number(value: u32, num_fmt: &str) -> String {
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

fn normalize_bullet_text(text: &str, bullet_font: Option<&str>) -> String {
    let is_symbol_font = bullet_font.is_some_and(|f| {
        let lower = f.to_lowercase();
        lower.contains("wingdings") || lower.contains("symbol") || lower.contains("webdings")
    });
    text.chars()
        .map(|c| {
            let cp = c as u32;
            if (0xF000..=0xF0FF).contains(&cp) {
                if is_symbol_font {
                    c
                } else {
                    symbol_pua_to_unicode(cp).unwrap_or(c)
                }
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
    numbering: &NumberingInfo,
    counters: &mut HashMap<(String, u8), u32>,
    last_seen_level: &mut HashMap<String, u8>,
) -> (f32, f32, String, Option<String>) {
    let Some(num_pr) = num_pr else {
        return (0.0, 0.0, String::new(), None);
    };
    let Some(num_id) = wml_attr(num_pr, "numId") else {
        return (0.0, 0.0, String::new(), None);
    };
    if num_id == "0" {
        return (0.0, 0.0, String::new(), None);
    }
    let ilvl = wml_attr(num_pr, "ilvl")
        .and_then(|v| v.parse::<u8>().ok())
        .unwrap_or(0);

    let Some(abs_id) = numbering.num_to_abstract.get(num_id) else {
        return (0.0, 0.0, String::new(), None);
    };
    let Some(levels) = numbering.abstract_nums.get(abs_id.as_str()) else {
        return (0.0, 0.0, String::new(), None);
    };
    let Some(def) = levels.get(&ilvl) else {
        return (0.0, 0.0, String::new(), None);
    };

    // Reset deeper-level counters when returning to a higher level
    let prev_level = last_seen_level.get(num_id).copied();
    if let Some(prev) = prev_level {
        if ilvl <= prev {
            for deeper in (ilvl + 1)..=prev {
                counters.remove(&(num_id.to_string(), deeper));
            }
        }
    }
    last_seen_level.insert(num_id.to_string(), ilvl);

    // Increment or initialize counter using the level's start value
    let start = def.start;
    let current_counter = *counters
        .entry((num_id.to_string(), ilvl))
        .and_modify(|c| *c += 1)
        .or_insert(start);

    let label = if def.num_fmt == "bullet" {
        let text = normalize_bullet_text(&def.lvl_text, def.bullet_font.as_deref());
        if text.is_empty() {
            "\u{2022}".to_string()
        } else {
            text
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
                        .get(&(num_id.to_string(), lvl_idx))
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
    let bullet_font = if def.num_fmt == "bullet" {
        def.bullet_font.clone()
    } else {
        None
    };
    (def.indent_left, def.indent_hanging, label, bullet_font)
}
