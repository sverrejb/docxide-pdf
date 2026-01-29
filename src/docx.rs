use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

use crate::error::Error;
use crate::model::{Alignment, Document, Paragraph, Run};

const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPD_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

fn wml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(WML_NS))
}

fn dml<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

struct ThemeFonts {
    major: String,
    minor: String,
}

struct StyleDefaults {
    font_size: f32,
    font_name: String,
    space_after: f32,
}

struct ParagraphStyle {
    font_size: Option<f32>,
    font_name: Option<String>,
    space_before: f32,
    space_after: Option<f32>,
    alignment: Option<Alignment>,
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
    // Theme file may be at word/theme/theme1.xml or similar
    let names: Vec<String> = zip.file_names().map(|s| s.to_string()).collect();
    let theme_name = names.iter().find(|n| n.starts_with("word/theme/") && n.ends_with(".xml"));
    let Some(theme_name) = theme_name else {
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

    // Walk the XML looking for a:majorFont/a:latin and a:minorFont/a:latin
    for node in xml.descendants() {
        let ns = node.tag_name().namespace();
        if ns != Some(DML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "majorFont" => {
                if let Some(latin) = dml(node, "latin") {
                    if let Some(tf) = latin.attribute("typeface") {
                        if !tf.is_empty() {
                            major = tf.to_string();
                        }
                    }
                }
            }
            "minorFont" => {
                if let Some(latin) = dml(node, "latin") {
                    if let Some(tf) = latin.attribute("typeface") {
                        if !tf.is_empty() {
                            minor = tf.to_string();
                        }
                    }
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

fn parse_styles(
    zip: &mut zip::ZipArchive<std::fs::File>,
    theme: &ThemeFonts,
) -> StylesInfo {
    let mut defaults =
        StyleDefaults { font_size: 12.0, font_name: theme.minor.clone(), space_after: 8.0 };
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
            if let Some(sz_val) = wml(rpr, "sz")
                .and_then(|n| n.attribute((WML_NS, "val")))
                .and_then(|v| v.parse::<f32>().ok())
            {
                defaults.font_size = sz_val / 2.0;
            }
            if let Some(rfonts) = wml(rpr, "rFonts") {
                let ascii = rfonts.attribute((WML_NS, "ascii"));
                let ascii_theme = rfonts.attribute((WML_NS, "asciiTheme"));
                defaults.font_name = resolve_font(ascii, ascii_theme, theme, &theme.minor);
            }
        }
        if let Some(after_val) = wml(doc_defaults, "pPrDefault")
            .and_then(|n| wml(n, "pPr"))
            .and_then(|n| wml(n, "spacing"))
            .and_then(|n| n.attribute((WML_NS, "after")))
            .and_then(|v| v.parse::<f32>().ok())
        {
            defaults.space_after = twips_to_pts(after_val);
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

        let spacing = wml(style_node, "pPr").and_then(|n| wml(n, "spacing"));
        let space_before = spacing
            .and_then(|n| n.attribute((WML_NS, "before")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .unwrap_or(0.0);
        let space_after = spacing
            .and_then(|n| n.attribute((WML_NS, "after")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts);

        let rpr = wml(style_node, "rPr");

        let font_size = rpr
            .and_then(|n| wml(n, "sz"))
            .and_then(|n| n.attribute((WML_NS, "val")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);

        let font_name = rpr.and_then(|n| wml(n, "rFonts")).map(|rfonts| {
            let ascii = rfonts.attribute((WML_NS, "ascii"));
            let ascii_theme = rfonts.attribute((WML_NS, "asciiTheme"));
            resolve_font(ascii, ascii_theme, theme, &defaults.font_name)
        });

        let alignment = wml(style_node, "pPr")
            .and_then(|ppr| wml(ppr, "jc"))
            .and_then(|n| n.attribute((WML_NS, "val")))
            .map(parse_alignment);

        paragraph_styles.insert(
            style_id.to_string(),
            ParagraphStyle { font_size, font_name, space_before, space_after, alignment },
        );
    }

    StylesInfo { defaults, paragraph_styles }
}

pub fn parse(path: &Path) -> Result<Document, Error> {
    let file = std::fs::File::open(path)?;
    let mut zip = zip::ZipArchive::new(file)?;

    let theme = parse_theme(&mut zip);
    let styles = parse_styles(&mut zip, &theme);

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")?
        .read_to_string(&mut xml_content)?;

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = root
        .children()
        .find(|n| n.tag_name().name() == "body" && n.tag_name().namespace() == Some(WML_NS))
        .ok_or_else(|| Error::Pdf("Missing w:body".into()))?;

    let sect_pr = body
        .children()
        .find(|n| n.tag_name().name() == "sectPr" && n.tag_name().namespace() == Some(WML_NS));

    let (page_width, page_height, margin_top, margin_bottom, margin_left, margin_right, line_pitch) =
        if let Some(sect) = sect_pr {
            let pg_sz = sect.children().find(|n| {
                n.tag_name().name() == "pgSz" && n.tag_name().namespace() == Some(WML_NS)
            });
            let pg_mar = sect.children().find(|n| {
                n.tag_name().name() == "pgMar" && n.tag_name().namespace() == Some(WML_NS)
            });
            let doc_grid = sect.children().find(|n| {
                n.tag_name().name() == "docGrid" && n.tag_name().namespace() == Some(WML_NS)
            });

            let width = pg_sz
                .and_then(|n| n.attribute((WML_NS, "w")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(612.0);
            let height = pg_sz
                .and_then(|n| n.attribute((WML_NS, "h")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(792.0);
            let top = pg_mar
                .and_then(|n| n.attribute((WML_NS, "top")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);
            let bottom = pg_mar
                .and_then(|n| n.attribute((WML_NS, "bottom")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);
            let left = pg_mar
                .and_then(|n| n.attribute((WML_NS, "left")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);
            let right = pg_mar
                .and_then(|n| n.attribute((WML_NS, "right")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);
            let pitch = doc_grid
                .and_then(|n| n.attribute((WML_NS, "linePitch")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(styles.defaults.font_size * 1.2);

            (width, height, top, bottom, left, right, pitch)
        } else {
            (612.0, 792.0, 72.0, 72.0, 72.0, 72.0, styles.defaults.font_size * 1.2)
        };

    let mut paragraphs = Vec::new();

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
                });
            }
            "p" if node.tag_name().namespace() == Some(WML_NS) => {
                let ppr = node.children().find(|n| {
                    n.tag_name().name() == "pPr" && n.tag_name().namespace() == Some(WML_NS)
                });

                let para_style_id = ppr
                    .and_then(|ppr| wml(ppr, "pStyle"))
                    .and_then(|n| n.attribute((WML_NS, "val")));

                let para_style = para_style_id.and_then(|id| styles.paragraph_styles.get(id));

                let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));

                let space_before = inline_spacing
                    .and_then(|n| n.attribute((WML_NS, "before")))
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(twips_to_pts)
                    .or_else(|| para_style.map(|s| s.space_before))
                    .unwrap_or(0.0);

                let space_after = inline_spacing
                    .and_then(|n| n.attribute((WML_NS, "after")))
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(twips_to_pts)
                    .or_else(|| para_style.and_then(|s| s.space_after))
                    .unwrap_or(styles.defaults.space_after);

                let style_font_size = para_style
                    .and_then(|s| s.font_size)
                    .unwrap_or(styles.defaults.font_size);

                let style_font_name = para_style
                    .and_then(|s| s.font_name.as_deref())
                    .unwrap_or(&styles.defaults.font_name)
                    .to_string();

                let alignment = ppr
                    .and_then(|ppr| wml(ppr, "jc"))
                    .and_then(|n| n.attribute((WML_NS, "val")))
                    .map(parse_alignment)
                    .or_else(|| para_style.and_then(|s| s.alignment.as_ref()).map(|a| match a {
                        Alignment::Center => Alignment::Center,
                        Alignment::Right => Alignment::Right,
                        Alignment::Left => Alignment::Left,
                    }))
                    .unwrap_or(Alignment::Left);

                let mut runs = Vec::new();

                for run_node in node.children() {
                    if run_node.tag_name().name() != "r"
                        || run_node.tag_name().namespace() != Some(WML_NS)
                    {
                        continue;
                    }

                    let rpr = run_node.children().find(|n| {
                        n.tag_name().name() == "rPr" && n.tag_name().namespace() == Some(WML_NS)
                    });

                    let font_size = rpr
                        .and_then(|n| {
                            n.children().find(|c| {
                                c.tag_name().name() == "sz"
                                    && c.tag_name().namespace() == Some(WML_NS)
                            })
                        })
                        .and_then(|n| n.attribute((WML_NS, "val")))
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|hp| hp / 2.0)
                        .unwrap_or(style_font_size);

                    let font_name = rpr
                        .and_then(|n| {
                            n.children().find(|c| {
                                c.tag_name().name() == "rFonts"
                                    && c.tag_name().namespace() == Some(WML_NS)
                            })
                        })
                        .map(|rfonts| {
                            let ascii = rfonts.attribute((WML_NS, "ascii"));
                            let ascii_theme = rfonts.attribute((WML_NS, "asciiTheme"));
                            resolve_font(ascii, ascii_theme, &theme, &style_font_name)
                        })
                        .unwrap_or_else(|| style_font_name.clone());

                    let bold = rpr
                        .map(|n| {
                            n.children().any(|c| {
                                c.tag_name().name() == "b"
                                    && c.tag_name().namespace() == Some(WML_NS)
                            })
                        })
                        .unwrap_or(false);

                    let italic = rpr
                        .map(|n| {
                            n.children().any(|c| {
                                c.tag_name().name() == "i"
                                    && c.tag_name().namespace() == Some(WML_NS)
                            })
                        })
                        .unwrap_or(false);

                    let text: String = run_node
                        .children()
                        .filter(|n| {
                            n.tag_name().name() == "t" && n.tag_name().namespace() == Some(WML_NS)
                        })
                        .filter_map(|n| n.text())
                        .collect();

                    if !text.is_empty() {
                        runs.push(Run { text, font_size, font_name, bold, italic });
                    }
                }

                // Compute drawing height from wp:inline/wp:anchor > wp:extent @cy
                let mut drawing_height: f32 = 0.0;
                for child in node.children() {
                    let drawing_node = if child.tag_name().name() == "drawing"
                        && child.tag_name().namespace() == Some(WML_NS)
                    {
                        Some(child)
                    } else if child.tag_name().name() == "r"
                        && child.tag_name().namespace() == Some(WML_NS)
                    {
                        child.children().find(|n| {
                            n.tag_name().name() == "drawing"
                                && n.tag_name().namespace() == Some(WML_NS)
                        })
                    } else {
                        None
                    };

                    if let Some(drawing) = drawing_node {
                        for container in drawing.children() {
                            if (container.tag_name().name() == "inline"
                                || container.tag_name().name() == "anchor")
                                && container.tag_name().namespace() == Some(WPD_NS)
                            {
                                if let Some(extent) = container.children().find(|n| {
                                    n.tag_name().name() == "extent"
                                        && n.tag_name().namespace() == Some(WPD_NS)
                                }) {
                                    if let Some(cy) =
                                        extent.attribute("cy").and_then(|v| v.parse::<f32>().ok())
                                    {
                                        drawing_height = drawing_height.max(cy / 12700.0);
                                    }
                                }
                            }
                        }
                    }
                }

                paragraphs.push(Paragraph {
                    runs,
                    space_before,
                    space_after,
                    content_height: drawing_height,
                    alignment,
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
        paragraphs,
    })
}
