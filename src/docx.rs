use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

use crate::error::Error;
use crate::model::{Alignment, Document, Paragraph, Run};

struct LevelDef {
    num_fmt: String,
    lvl_text: String,
    indent_left: f32,
    indent_hanging: f32,
}

struct NumberingInfo {
    abstract_nums: HashMap<String, HashMap<u8, LevelDef>>,
    num_to_abstract: HashMap<String, String>,
}

const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPD_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

fn parse_hex_color(val: &str) -> Option<[u8; 3]> {
    if val == "auto" || val.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&val[0..2], 16).ok()?;
    let g = u8::from_str_radix(&val[2..4], 16).ok()?;
    let b = u8::from_str_radix(&val[4..6], 16).ok()?;
    Some([r, g, b])
}

fn wml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(WML_NS))
}

fn wml_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    wml(node, child).and_then(|n| n.attribute((WML_NS, "val")))
}

fn twips_attr(node: roxmltree::Node, attr: &str) -> Option<f32> {
    node.attribute((WML_NS, attr))
        .and_then(|v| v.parse::<f32>().ok())
        .map(twips_to_pts)
}

fn dml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

fn latin_typeface<'a>(node: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    dml(node, "latin")
        .and_then(|n| n.attribute("typeface"))
        .filter(|tf| !tf.is_empty())
}

struct ThemeFonts {
    major: String,
    minor: String,
}

struct StyleDefaults {
    font_size: f32,
    font_name: String,
    space_after: f32,
    line_spacing: f32, // multiplier from w:spacing @line / 240
}

struct ParagraphStyle {
    font_size: Option<f32>,
    font_name: Option<String>,
    color: Option<[u8; 3]>,
    space_before: f32,
    space_after: Option<f32>,
    alignment: Option<Alignment>,
    contextual_spacing: bool,
}

struct StylesInfo {
    defaults: StyleDefaults,
    paragraph_styles: HashMap<String, ParagraphStyle>,
}

fn parse_alignment(val: &str) -> Alignment {
    match val {
        "center" => Alignment::Center,
        "right" | "end" => Alignment::Right,
        _ => Alignment::Left,
    }
}

fn parse_theme(zip: &mut zip::ZipArchive<std::fs::File>) -> ThemeFonts {
    let mut major = String::from("Aptos Display");
    let mut minor = String::from("Aptos");

    let mut xml_content = String::new();
    let names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
    let Some(theme_name) = names.iter().find(|n| n.starts_with("word/theme/") && n.ends_with(".xml")) else {
        return ThemeFonts { major, minor };
    };
    let theme_name = theme_name.clone();
    let Ok(mut file) = zip.by_name(&theme_name) else {
        return ThemeFonts { major, minor };
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return ThemeFonts { major, minor };
    }
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return ThemeFonts { major, minor };
    };

    for node in xml.descendants() {
        if node.tag_name().namespace() != Some(DML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "majorFont" => {
                if let Some(tf) = latin_typeface(node) {
                    major = tf.to_string();
                }
            }
            "minorFont" => {
                if let Some(tf) = latin_typeface(node) {
                    minor = tf.to_string();
                }
            }
            _ => {}
        }
    }

    ThemeFonts { major, minor }
}

fn resolve_font(
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

fn resolve_font_from_node(
    rfonts: roxmltree::Node,
    theme: &ThemeFonts,
    default_font: &str,
) -> String {
    resolve_font(
        rfonts.attribute((WML_NS, "ascii")),
        rfonts.attribute((WML_NS, "asciiTheme")),
        theme,
        default_font,
    )
}

fn parse_styles(
    zip: &mut zip::ZipArchive<std::fs::File>,
    theme: &ThemeFonts,
) -> StylesInfo {
    let mut defaults = StyleDefaults {
        font_size: 12.0,
        font_name: theme.minor.clone(),
        space_after: 8.0,
        line_spacing: 1.2,
    };
    let mut paragraph_styles = HashMap::new();

    let mut xml_content = String::new();
    let Ok(mut file) = zip.by_name("word/styles.xml") else {
        return StylesInfo { defaults, paragraph_styles };
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return StylesInfo { defaults, paragraph_styles };
    }
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return StylesInfo { defaults, paragraph_styles };
    };

    let root = xml.root_element();

    if let Some(doc_defaults) = wml(root, "docDefaults") {
        if let Some(rpr) = wml(doc_defaults, "rPrDefault").and_then(|n| wml(n, "rPr")) {
            if let Some(sz_val) = wml_attr(rpr, "sz").and_then(|v| v.parse::<f32>().ok()) {
                defaults.font_size = sz_val / 2.0;
            }
            if let Some(rfonts) = wml(rpr, "rFonts") {
                defaults.font_name = resolve_font_from_node(rfonts, theme, &theme.minor);
            }
        }
        let default_spacing = wml(doc_defaults, "pPrDefault")
            .and_then(|n| wml(n, "pPr"))
            .and_then(|n| wml(n, "spacing"));
        if let Some(spacing) = default_spacing {
            if let Some(after_val) = twips_attr(spacing, "after") {
                defaults.space_after = after_val;
            }
            if let Some(line_val) = spacing
                .attribute((WML_NS, "line"))
                .and_then(|v| v.parse::<f32>().ok())
            {
                defaults.line_spacing = line_val / 240.0;
            }
        }
    }

    for style_node in root.children() {
        if style_node.tag_name().name() != "style"
            || style_node.tag_name().namespace() != Some(WML_NS)
        {
            continue;
        }
        if style_node.attribute((WML_NS, "type")) != Some("paragraph") {
            continue;
        }
        let Some(style_id) = style_node.attribute((WML_NS, "styleId")) else {
            continue;
        };

        let ppr = wml(style_node, "pPr");
        let spacing = ppr.and_then(|n| wml(n, "spacing"));
        let space_before = spacing
            .and_then(|n| twips_attr(n, "before"))
            .unwrap_or(0.0);
        let space_after = spacing.and_then(|n| twips_attr(n, "after"));

        let rpr = wml(style_node, "rPr");

        let font_size = rpr
            .and_then(|n| wml_attr(n, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);

        let font_name = rpr
            .and_then(|n| wml(n, "rFonts"))
            .map(|rfonts| resolve_font_from_node(rfonts, theme, &defaults.font_name));

        let color = rpr
            .and_then(|n| wml_attr(n, "color"))
            .and_then(parse_hex_color);

        let alignment = ppr
            .and_then(|ppr| wml_attr(ppr, "jc"))
            .map(parse_alignment);

        let contextual_spacing = ppr
            .and_then(|ppr| wml(ppr, "contextualSpacing"))
            .is_some();

        paragraph_styles.insert(
            style_id.to_string(),
            ParagraphStyle {
                font_size,
                font_name,
                color,
                space_before,
                space_after,
                alignment,
                contextual_spacing,
            },
        );
    }

    StylesInfo { defaults, paragraph_styles }
}

fn parse_numbering(zip: &mut zip::ZipArchive<std::fs::File>) -> NumberingInfo {
    let mut abstract_nums: HashMap<String, HashMap<u8, LevelDef>> = HashMap::new();
    let mut num_to_abstract: HashMap<String, String> = HashMap::new();

    let mut xml_content = String::new();
    let Ok(mut file) = zip.by_name("word/numbering.xml") else {
        return NumberingInfo { abstract_nums, num_to_abstract };
    };
    if file.read_to_string(&mut xml_content).is_err() {
        return NumberingInfo { abstract_nums, num_to_abstract };
    }
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return NumberingInfo { abstract_nums, num_to_abstract };
    };

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
                    let ind = wml(lvl, "pPr").and_then(|ppr| wml(ppr, "ind"));
                    let indent_left = ind
                        .and_then(|n| twips_attr(n, "left"))
                        .unwrap_or(0.0);
                    let indent_hanging = ind
                        .and_then(|n| twips_attr(n, "hanging"))
                        .unwrap_or(0.0);
                    levels.insert(ilvl, LevelDef { num_fmt, lvl_text, indent_left, indent_hanging });
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

    NumberingInfo { abstract_nums, num_to_abstract }
}

pub fn parse(path: &Path) -> Result<Document, Error> {
    let file = std::fs::File::open(path)?;
    let mut zip = zip::ZipArchive::new(file)?;

    let theme = parse_theme(&mut zip);
    let styles = parse_styles(&mut zip, &theme);
    let numbering = parse_numbering(&mut zip);

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")?
        .read_to_string(&mut xml_content)?;

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = wml(root, "body")
        .ok_or_else(|| Error::Pdf("Missing w:body".into()))?;

    let sect = wml(body, "sectPr");
    let pg_sz = sect.and_then(|s| wml(s, "pgSz"));
    let pg_mar = sect.and_then(|s| wml(s, "pgMar"));
    let doc_grid = sect.and_then(|s| wml(s, "docGrid"));

    let page_width = pg_sz
        .and_then(|n| twips_attr(n, "w"))
        .unwrap_or(612.0);
    let page_height = pg_sz
        .and_then(|n| twips_attr(n, "h"))
        .unwrap_or(792.0);
    let margin_top = pg_mar
        .and_then(|n| twips_attr(n, "top"))
        .unwrap_or(72.0);
    let margin_bottom = pg_mar
        .and_then(|n| twips_attr(n, "bottom"))
        .unwrap_or(72.0);
    let margin_left = pg_mar
        .and_then(|n| twips_attr(n, "left"))
        .unwrap_or(72.0);
    let margin_right = pg_mar
        .and_then(|n| twips_attr(n, "right"))
        .unwrap_or(72.0);
    let line_pitch = doc_grid
        .and_then(|n| twips_attr(n, "linePitch"))
        .unwrap_or(styles.defaults.font_size * 1.2);

    let mut paragraphs = Vec::new();
    let mut counters: HashMap<(String, u8), u32> = HashMap::new();

    for node in body.children() {
        match node.tag_name().name() {
            "tbl" if node.tag_name().namespace() == Some(WML_NS) => {
                let row_count = node
                    .children()
                    .filter(|n| {
                        n.tag_name().name() == "tr" && n.tag_name().namespace() == Some(WML_NS)
                    })
                    .count();
                let estimated_height = row_count as f32 * (10.5 * 1.6 + 4.0);
                paragraphs.push(Paragraph {
                    runs: vec![],
                    space_before: 0.0,
                    space_after: 4.0,
                    content_height: estimated_height,
                    alignment: Alignment::Left,
                    indent_left: 0.0,
                    indent_hanging: 0.0,
                    list_label: String::new(),
                    contextual_spacing: false,
                });
            }
            "p" if node.tag_name().namespace() == Some(WML_NS) => {
                let ppr = wml(node, "pPr");

                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"));

                let para_style = para_style_id.and_then(|id| styles.paragraph_styles.get(id));

                let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));

                let space_before = inline_spacing
                    .and_then(|n| twips_attr(n, "before"))
                    .or_else(|| para_style.map(|s| s.space_before))
                    .unwrap_or(0.0);

                let space_after = inline_spacing
                    .and_then(|n| twips_attr(n, "after"))
                    .or_else(|| para_style.and_then(|s| s.space_after))
                    .unwrap_or(styles.defaults.space_after);

                let style_font_size = para_style
                    .and_then(|s| s.font_size)
                    .unwrap_or(styles.defaults.font_size);

                let style_font_name = para_style
                    .and_then(|s| s.font_name.as_deref())
                    .unwrap_or(&styles.defaults.font_name)
                    .to_string();

                let style_color: Option<[u8; 3]> = para_style.and_then(|s| s.color);

                let alignment = ppr
                    .and_then(|ppr| wml_attr(ppr, "jc"))
                    .map(parse_alignment)
                    .or_else(|| para_style.and_then(|s| s.alignment))
                    .unwrap_or(Alignment::Left);

                let contextual_spacing = ppr
                    .and_then(|ppr| wml(ppr, "contextualSpacing"))
                    .is_some()
                    || para_style.is_some_and(|s| s.contextual_spacing);

                let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
                let (mut indent_left, mut indent_hanging, list_label) =
                    parse_list_info(num_pr, &numbering, &mut counters);

                // Paragraph-level w:ind overrides level def
                if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                    if let Some(v) = twips_attr(ind, "left") {
                        indent_left = v;
                    }
                    if let Some(v) = twips_attr(ind, "hanging") {
                        indent_hanging = v;
                    }
                }

                let mut runs = Vec::new();

                for run_node in node.children() {
                    if run_node.tag_name().name() != "r"
                        || run_node.tag_name().namespace() != Some(WML_NS)
                    {
                        continue;
                    }

                    let rpr = wml(run_node, "rPr");

                    let font_size = rpr
                        .and_then(|n| wml_attr(n, "sz"))
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|hp| hp / 2.0)
                        .unwrap_or(style_font_size);

                    let font_name = rpr
                        .and_then(|n| wml(n, "rFonts"))
                        .map(|rfonts| resolve_font_from_node(rfonts, &theme, &style_font_name))
                        .unwrap_or_else(|| style_font_name.clone());

                    let bold = rpr
                        .and_then(|n| wml(n, "b"))
                        .is_some();

                    let italic = rpr
                        .and_then(|n| wml(n, "i"))
                        .is_some();

                    let color = rpr
                        .and_then(|n| wml_attr(n, "color"))
                        .and_then(parse_hex_color)
                        .or(style_color);

                    let text: String = run_node
                        .children()
                        .filter(|n| {
                            n.tag_name().name() == "t" && n.tag_name().namespace() == Some(WML_NS)
                        })
                        .filter_map(|n| n.text())
                        .collect();

                    if !text.is_empty() {
                        runs.push(Run { text, font_size, font_name, bold, italic, color });
                    }
                }

                let drawing_height = compute_drawing_height(node);

                paragraphs.push(Paragraph {
                    runs,
                    space_before,
                    space_after,
                    content_height: drawing_height,
                    alignment,
                    indent_left,
                    indent_hanging,
                    list_label,
                    contextual_spacing,
                });
            }
            _ => {}
        }
    }

    Ok(Document {
        page_width,
        page_height,
        margin_top,
        margin_bottom,
        margin_left,
        margin_right,
        line_pitch,
        line_spacing: styles.defaults.line_spacing,
        paragraphs,
    })
}

fn parse_list_info(
    num_pr: Option<roxmltree::Node>,
    numbering: &NumberingInfo,
    counters: &mut HashMap<(String, u8), u32>,
) -> (f32, f32, String) {
    let Some(num_pr) = num_pr else {
        return (0.0, 0.0, String::new());
    };
    let Some(num_id) = wml_attr(num_pr, "numId") else {
        return (0.0, 0.0, String::new());
    };
    let ilvl = wml_attr(num_pr, "ilvl")
        .and_then(|v| v.parse::<u8>().ok())
        .unwrap_or(0);

    let Some(def) = numbering
        .num_to_abstract
        .get(num_id)
        .and_then(|abs_id| numbering.abstract_nums.get(abs_id))
        .and_then(|levels| levels.get(&ilvl))
    else {
        return (0.0, 0.0, String::new());
    };

    let counter = counters
        .entry((num_id.to_string(), ilvl))
        .and_modify(|c| *c += 1)
        .or_insert(1);
    let label = if def.num_fmt == "bullet" {
        "\u{2022}".to_string()
    } else {
        def.lvl_text.replace(&format!("%{}", ilvl + 1), &counter.to_string())
    };
    (def.indent_left, def.indent_hanging, label)
}

fn compute_drawing_height(para_node: roxmltree::Node) -> f32 {
    let mut max_height: f32 = 0.0;
    for child in para_node.children() {
        let drawing_node = if child.tag_name().name() == "drawing"
            && child.tag_name().namespace() == Some(WML_NS)
        {
            Some(child)
        } else if child.tag_name().name() == "r"
            && child.tag_name().namespace() == Some(WML_NS)
        {
            wml(child, "drawing")
        } else {
            None
        };

        let Some(drawing) = drawing_node else { continue };
        for container in drawing.children() {
            let name = container.tag_name().name();
            if (name == "inline" || name == "anchor")
                && container.tag_name().namespace() == Some(WPD_NS)
                && let Some(cy) = container
                    .children()
                    .find(|n| {
                        n.tag_name().name() == "extent"
                            && n.tag_name().namespace() == Some(WPD_NS)
                    })
                    .and_then(|n| n.attribute("cy"))
                    .and_then(|v| v.parse::<f32>().ok())
            {
                max_height = max_height.max(cy / 12700.0);
            }
        }
    }
    max_height
}
