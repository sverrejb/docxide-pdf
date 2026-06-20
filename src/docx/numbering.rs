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
    /// §17.9.4 isLgl: render every %N reference in this level's lvlText as
    /// decimal regardless of the referenced level's own numFmt.
    pub(super) is_lgl: bool,
    /// §17.9.10 lvlRestart: 1-based level whose use restarts this level (that
    /// level "or any earlier"). `Some(0)` = never restart; `None` = default
    /// (restart on any earlier level).
    pub(super) lvl_restart: Option<u32>,
    /// §17.9.23 pStyle: styleId this level is associated with. A paragraph using
    /// that style picks THIS level regardless of the style's numPr ilvl.
    pub(super) pstyle: Option<String>,
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
    /// Full per-level redefinitions from w:num/w:lvlOverride/w:lvl (§17.9.5),
    /// keyed by numId. These replace the abstract level entirely.
    pub(super) level_overrides: HashMap<String, HashMap<u8, LevelDef>>,
}

fn parse_level_def(lvl: roxmltree::Node) -> Option<(u8, LevelDef)> {
    let ilvl = lvl
        .attribute((WML_NS, "ilvl"))
        .and_then(|v| v.parse::<u8>().ok())?;
    let lvl_text = wml_attr(lvl, "lvlText").unwrap_or("").to_string();
    // §17.9.17: an omitted numFmt means decimal. Keep the bullet fallback
    // only for symbol-style lvlText without %N placeholders, where decimal
    // would drop the bullet glyph's symbol font.
    let num_fmt = wml_attr(lvl, "numFmt")
        .map(|s| s.to_string())
        .unwrap_or_else(|| {
            let has_placeholder = lvl_text.as_bytes().windows(2).any(|w| {
                w[0] == b'%' && w[1].is_ascii_digit()
            });
            if has_placeholder { "decimal" } else { "bullet" }.to_string()
        });
    let start = wml_attr(lvl, "start")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1);
    let ppr = wml(lvl, "pPr");
    let ind = ppr.and_then(|p| wml(p, "ind"));
    let indent_left = ind
        .and_then(|n| twips_attr(n, "start").or_else(|| twips_attr(n, "left")))
        .unwrap_or(0.0);
    let indent_hanging = ind.and_then(|n| twips_attr(n, "hanging")).unwrap_or(0.0);
    let tab_stop = ppr.and_then(|p| wml(p, "tabs")).and_then(|tabs| {
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
    let is_lgl = wml_bool(lvl, "isLgl").unwrap_or(false);
    let lvl_restart = wml_attr(lvl, "lvlRestart").and_then(|v| v.parse::<u32>().ok());
    let pstyle = wml_attr(lvl, "pStyle").map(|s| s.to_string());
    Some((
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
            is_lgl,
            lvl_restart,
            pstyle,
        },
    ))
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
    let mut level_overrides = HashMap::new();

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
                let levels: HashMap<u8, LevelDef> = node
                    .children()
                    .filter(|n| n.has_tag_name((WML_NS, "lvl")))
                    .filter_map(parse_level_def)
                    .collect();
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
                let mut starts: HashMap<u8, u32> = HashMap::new();
                let mut lvl_defs: HashMap<u8, LevelDef> = HashMap::new();
                for ovr in node
                    .children()
                    .filter(|n| n.has_tag_name((WML_NS, "lvlOverride")))
                {
                    let ovr_ilvl = ovr
                        .attribute((WML_NS, "ilvl"))
                        .and_then(|v| v.parse::<u8>().ok());
                    if let (Some(ilvl), Some(val)) = (
                        ovr_ilvl,
                        wml_attr(ovr, "startOverride").and_then(|v| v.parse::<u32>().ok()),
                    ) {
                        starts.insert(ilvl, val);
                    }
                    // Full level redefinition (§17.9.5) — replaces the
                    // abstract level entirely.
                    if let Some((ilvl, def)) =
                        wml(ovr, "lvl").and_then(parse_level_def)
                    {
                        lvl_defs.insert(ilvl, def);
                    }
                }
                if !starts.is_empty() {
                    start_overrides.insert(num_id.to_string(), starts);
                }
                if !lvl_defs.is_empty() {
                    level_overrides.insert(num_id.to_string(), lvl_defs);
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
        level_overrides,
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
    effective_style_id: Option<&str>,
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
    // lvlOverride level redefinitions on the numId replace abstract levels.
    let num_lvl_overrides = numbering.level_overrides.get(num_id_str);
    let lookup_level = |lvl: u8| -> Option<&LevelDef> {
        num_lvl_overrides
            .and_then(|m| m.get(&lvl))
            .or_else(|| levels.get(&lvl))
    };
    let Some(def) = lookup_level(ilvl) else {
        return ListLabelInfo::default();
    };

    // §17.9.23: a level's pStyle names the single paragraph style that
    // auto-numbers at that level. When numbering is inherited from a style (no
    // direct paragraph numPr) and the resolved level is owned by a DIFFERENT
    // style, Word suppresses the number — the paragraph keeps the list indent
    // and sits at the hanging position, with no label and no counter advance.
    // ponytail: matches the level's named style directly, not styles basedOn it.
    if num_pr.is_none()
        && let Some(owner) = def.pstyle.as_deref()
        && effective_style_id != Some(owner)
    {
        return ListLabelInfo {
            indent_left: def.indent_left,
            indent_hanging: def.indent_hanging,
            tab_stop: def.tab_stop,
            ..ListLabelInfo::default()
        };
    }

    // Key counters by abstractNumId so all numIds sharing the same
    // abstract definition share a single counter stream.
    let abs_key: u32 = abs_id.parse().unwrap_or(0);

    // Reset deeper-level counters when returning to a higher level, honoring
    // each deeper level's §17.9.10 lvlRestart. The trigger is the level we just
    // moved to (`ilvl`). Default (None) restarts on any earlier level; Some(0)
    // never restarts; Some(k) restarts only when the trigger level is `k` "or
    // any earlier" (0-based ilvl < k), and is ignored when k names a level
    // deeper than the deeper level itself (k > deeper+1) → default restart.
    if let Some(&prev) = last_seen_level.get(&abs_key) {
        for deeper in (ilvl + 1)..=prev {
            let restarts = match lookup_level(deeper).and_then(|d| d.lvl_restart) {
                Some(0) => false,
                Some(k) if k <= (deeper as u32) + 1 => (ilvl as u32) < k,
                _ => true,
            };
            if restarts {
                counters.remove(&(abs_key, deeper));
            }
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
                let ref_def = lookup_level(lvl_idx);
                let lvl_counter = if lvl_idx == ilvl {
                    current_counter
                } else {
                    counters
                        .get(&(abs_key, lvl_idx))
                        .copied()
                        .unwrap_or(ref_def.map(|d| d.start).unwrap_or(1))
                };
                // §17.9.4: isLgl forces every referenced level to decimal.
                let lvl_fmt = if def.is_lgl {
                    "decimal"
                } else {
                    ref_def.map(|d| d.num_fmt.as_str()).unwrap_or("decimal")
                };
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

    // --- Level parsing ---

    fn parse_lvl(xml: &str) -> (u8, LevelDef) {
        let doc = roxmltree::Document::parse(xml).unwrap();
        parse_level_def(doc.root_element()).unwrap()
    }

    #[test]
    fn test_missing_numfmt_with_placeholder_defaults_to_decimal() {
        let (ilvl, def) = parse_lvl(
            r#"<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="0"><w:lvlText w:val="%1."/></w:lvl>"#,
        );
        assert_eq!(ilvl, 0);
        assert_eq!(def.num_fmt, "decimal");
    }

    #[test]
    fn test_missing_numfmt_without_placeholder_stays_bullet() {
        let (_, def) = parse_lvl(
            r#"<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="0"><w:lvlText w:val="&#xF0B7;"/></w:lvl>"#,
        );
        assert_eq!(def.num_fmt, "bullet");
    }

    #[test]
    fn test_lvl_override_replaces_abstract_level() {
        let mut numbering = NumberingInfo::default();
        let abstract_def = LevelDef {
            num_fmt: "decimal".into(),
            lvl_text: "%1.".into(),
            indent_left: 36.0,
            indent_hanging: 18.0,
            tab_stop: None,
            start: 1,
            bullet_font: None,
            label_font_size: None,
            label_bold: false,
            label_color: None,
            suff: "tab".into(),
            label_font: None,
            is_lgl: false,
            lvl_restart: None,
            pstyle: None,
        };
        let override_def = LevelDef {
            num_fmt: "upperLetter".into(),
            lvl_text: "%1)".into(),
            indent_left: 10.0,
            indent_hanging: 5.0,
            ..abstract_def.clone()
        };
        numbering
            .abstract_nums
            .insert("0".into(), HashMap::from([(0u8, abstract_def)]));
        numbering.num_to_abstract.insert("5".into(), "0".into());
        numbering
            .level_overrides
            .insert("5".into(), HashMap::from([(0u8, override_def)]));

        let mut counters = HashMap::new();
        let mut last_seen = HashMap::new();
        let mut applied = HashSet::new();
        let info = parse_list_info(
            None,
            Some("5"),
            Some(0),
            None,
            &numbering,
            &mut counters,
            &mut last_seen,
            &mut applied,
        );
        assert_eq!(info.label, "A)");
        assert_eq!(info.indent_left, 10.0);
        assert_eq!(info.indent_hanging, 5.0);
    }

    #[test]
    fn test_islgl_forces_referenced_levels_decimal() {
        // §17.9.4: lvl0 upperRoman, lvl1 decimal+isLgl with lvlText "%1.%2.".
        // isLgl must force the %1 (upperRoman) reference to decimal → "1.1.".
        let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let (_, lvl0) = parse_lvl(&format!(
            r#"<w:lvl xmlns:w="{ns}" w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>"#
        ));
        let (_, lvl1) = parse_lvl(&format!(
            r#"<w:lvl xmlns:w="{ns}" w:ilvl="1"><w:numFmt w:val="decimal"/><w:isLgl/><w:lvlText w:val="%1.%2."/></w:lvl>"#
        ));
        assert!(lvl1.is_lgl);

        let mut numbering = NumberingInfo::default();
        numbering
            .abstract_nums
            .insert("0".into(), HashMap::from([(0u8, lvl0), (1u8, lvl1)]));
        numbering.num_to_abstract.insert("100".into(), "0".into());

        let mut counters = HashMap::new();
        let mut last_seen = HashMap::new();
        let mut applied = HashSet::new();
        let lvl0_info = parse_list_info(
            None, Some("100"), Some(0), None,
            &numbering, &mut counters, &mut last_seen, &mut applied,
        );
        assert_eq!(lvl0_info.label, "I.");
        let lvl1_info = parse_list_info(
            None, Some("100"), Some(1), None,
            &numbering, &mut counters, &mut last_seen, &mut applied,
        );
        // Without isLgl this would be "I.1."; isLgl forces %1 to decimal.
        assert_eq!(lvl1_info.label, "1.1.");
    }

    #[test]
    fn test_lvlrestart_zero_keeps_continuous_numbering() {
        // §17.9.10: lvl1 carries lvlRestart=0, so its counter must NOT reset
        // when lvl0 advances — sub-items run (1)(2)(3)(4) across the headings.
        let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let (_, lvl0) = parse_lvl(&format!(
            r#"<w:lvl xmlns:w="{ns}" w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>"#
        ));
        let (_, lvl1) = parse_lvl(&format!(
            r#"<w:lvl xmlns:w="{ns}" w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlRestart w:val="0"/><w:lvlText w:val="(%2)"/></w:lvl>"#
        ));
        assert_eq!(lvl1.lvl_restart, Some(0));

        let mut numbering = NumberingInfo::default();
        numbering
            .abstract_nums
            .insert("0".into(), HashMap::from([(0u8, lvl0), (1u8, lvl1)]));
        numbering.num_to_abstract.insert("100".into(), "0".into());

        let mut counters = HashMap::new();
        let mut last_seen = HashMap::new();
        let mut applied = HashSet::new();
        let mut label = |ilvl: u8| {
            parse_list_info(
                None, Some("100"), Some(ilvl), None,
                &numbering, &mut counters, &mut last_seen, &mut applied,
            )
            .label
        };
        assert_eq!(label(0), "1."); // First section
        assert_eq!(label(1), "(1)");
        assert_eq!(label(1), "(2)");
        assert_eq!(label(0), "2."); // Second section — lvl1 must NOT reset
        assert_eq!(label(1), "(3)");
        assert_eq!(label(1), "(4)");
    }

    #[test]
    fn test_pstyle_gates_number_to_owning_style() {
        // §17.9.23: lvl0 is owned (pStyle) by "PHeadingA". A paragraph whose
        // effective style is "PHeadingA" auto-numbers ("I."); a "PHeadingB"
        // paragraph resolving to the same lvl0 (its style numPr has no ilvl)
        // gets the lvl0 indent but NO number, and must not advance the counter.
        let ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        let lvl0_xml = format!(
            r#"<w:lvl xmlns:w="{ns}" w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:pStyle w:val="PHeadingA"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="432"/></w:pPr></w:lvl>"#
        );
        let mut numbering = NumberingInfo::default();
        numbering.abstract_nums.insert(
            "100".into(),
            HashMap::from([(0u8, parse_lvl(&lvl0_xml).1)]),
        );
        numbering.num_to_abstract.insert("100".into(), "100".into());

        let mut counters = HashMap::new();
        let mut last_seen = HashMap::new();
        let mut applied = HashSet::new();
        let mut label = |style: &str| {
            parse_list_info(
                None, Some("100"), None, Some(style),
                &numbering, &mut counters, &mut last_seen, &mut applied,
            )
        };
        // Owning style numbers.
        assert_eq!(label("PHeadingA").label, "I.");
        // Non-owning style: indent kept, no label, counter not advanced.
        let gated = label("PHeadingB");
        assert_eq!(gated.label, "");
        assert_eq!(gated.indent_left, 36.0);
        assert_eq!(gated.indent_hanging, 21.6);
        // Next owning paragraph is "II.", proving the gated one didn't increment.
        assert_eq!(label("PHeadingA").label, "II.");
    }
}
