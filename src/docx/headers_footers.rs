use std::collections::HashMap;
use std::io::Read;

use crate::model::{Alignment, Block, Footnote, HeaderFooter, LineSpacing, Paragraph};

use super::numbering::NumberingInfo;
use super::parse_table_node;
use super::runs::parse_runs;
use super::styles::{ParagraphStyle, StylesInfo, ThemeFonts, parse_alignment};
use super::{
    WML_NS, parse_paragraph_borders, parse_paragraph_spacing, parse_tab_stops, wml, wml_attr,
};

fn is_wml_element(node: roxmltree::Node, name: &str) -> bool {
    node.tag_name().namespace() == Some(WML_NS) && node.tag_name().name() == name
}

fn is_block_element(node: roxmltree::Node) -> bool {
    node.tag_name().namespace() == Some(WML_NS) && matches!(node.tag_name().name(), "p" | "tbl")
}

fn resolve_alignment(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
) -> Alignment {
    ppr.and_then(|ppr| wml_attr(ppr, "jc"))
        .map(parse_alignment)
        .or_else(|| para_style.and_then(|s| s.alignment))
        .unwrap_or(Alignment::Left)
}

pub(super) fn parse_header_footer_xml<R: Read + std::io::Seek>(
    xml_content: &str,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> Option<HeaderFooter> {
    let xml = roxmltree::Document::parse(xml_content).ok()?;
    let root = xml.root_element();
    let mut blocks = Vec::new();

    let mut top_nodes: Vec<roxmltree::Node> = Vec::new();
    for child in root.children() {
        if child.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        if is_wml_element(child, "sdt") {
            if let Some(content) = wml(child, "sdtContent") {
                top_nodes.extend(content.children().filter(|n| is_block_element(*n)));
            }
        } else if is_block_element(child) {
            top_nodes.push(child);
        }
    }

    let numbering = NumberingInfo::default();
    let mut counters = HashMap::new();
    let mut last_seen_level = HashMap::new();

    for node in top_nodes {
        match node.tag_name().name() {
            "tbl" => {
                let table = parse_table_node(
                    node,
                    styles,
                    theme,
                    rels,
                    zip,
                    &numbering,
                    &mut counters,
                    &mut last_seen_level,
                );
                blocks.push(Block::Table(table));
            }
            "p" => {
                let ppr = wml(node, "pPr");
                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"))
                    .unwrap_or(&styles.default_paragraph_style_id);
                let para_style = styles.paragraph_styles.get(para_style_id);

                let alignment = resolve_alignment(ppr, para_style);
                let (sp_before, sp_after, line_spacing) = parse_paragraph_spacing(ppr, para_style);
                let parsed = parse_runs(node, styles, theme, rels, zip, &numbering);

                blocks.push(Block::Paragraph(Paragraph {
                    runs: parsed.runs,
                    alignment,
                    line_spacing,
                    space_before: sp_before.unwrap_or(0.0),
                    space_after: sp_after.unwrap_or(0.0),
                    borders: ppr.map(parse_paragraph_borders).unwrap_or_default(),
                    tab_stops: ppr.map(parse_tab_stops).unwrap_or_default(),
                    floating_images: parsed.floating_images,
                    textboxes: parsed.textboxes,
                    ..Paragraph::default()
                }));
            }
            _ => {}
        }
    }

    (!blocks.is_empty()).then(|| HeaderFooter { blocks })
}

pub(super) fn parse_footnotes<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
) -> HashMap<u32, Footnote> {
    let mut footnotes = HashMap::new();
    let Some(xml_text) = super::read_zip_text(zip, "word/footnotes.xml") else {
        return footnotes;
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_text) else {
        return footnotes;
    };
    let root = xml.root_element();
    let empty_rels = HashMap::new();
    let numbering = NumberingInfo::default();

    for node in root.children() {
        if !is_wml_element(node, "footnote") {
            continue;
        }
        // Skip separator/continuationSeparator footnotes (type attribute, IDs 0 and 1)
        if node.attribute((WML_NS, "type")).is_some() {
            continue;
        }
        let Some(id) = node
            .attribute((WML_NS, "id"))
            .and_then(|v| v.parse::<u32>().ok())
        else {
            continue;
        };

        let mut paragraphs = Vec::new();
        for p in node.children().filter(|n| is_wml_element(*n, "p")) {
            let ppr = wml(p, "pPr");
            let para_style_id = ppr
                .and_then(|ppr| wml_attr(ppr, "pStyle"))
                .unwrap_or("FootnoteText");
            let para_style = styles.paragraph_styles.get(para_style_id);

            let alignment = resolve_alignment(ppr, para_style);
            let parsed = parse_runs(p, styles, theme, &empty_rels, zip, &numbering);
            let (sp_before, sp_after, ls) = parse_paragraph_spacing(ppr, para_style);

            paragraphs.push(Paragraph {
                runs: parsed.runs,
                space_before: sp_before.unwrap_or(0.0),
                space_after: sp_after.unwrap_or(0.0),
                alignment,
                line_spacing: ls.or(Some(LineSpacing::Auto(1.0))),
                ..Paragraph::default()
            });
        }

        if !paragraphs.is_empty() {
            footnotes.insert(id, Footnote { paragraphs });
        }
    }

    footnotes
}
