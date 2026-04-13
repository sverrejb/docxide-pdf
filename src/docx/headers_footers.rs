use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use crate::model::{Alignment, Block, Footnote, HeaderFooter, LineSpacing, Paragraph};

use super::numbering::NumberingInfo;
use super::parse_table_node;
use super::runs::parse_runs;
use super::styles::{ParagraphStyle, StylesInfo, ThemeFonts, parse_alignment};
use super::{
    ParseContext, WML_NS, collect_block_nodes, parse_paragraph_spacing, wml, wml_attr,
};

fn is_wml_element(node: roxmltree::Node, name: &str) -> bool {
    node.tag_name().namespace() == Some(WML_NS) && node.tag_name().name() == name
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

pub(super) fn parse_header_footer_xml<R: Read + Seek>(
    xml_content: &str,
    ctx: &mut ParseContext<'_, R>,
) -> Option<HeaderFooter> {
    let xml = roxmltree::Document::parse(xml_content).ok()?;
    let root = xml.root_element();
    let mut blocks = Vec::new();

    let top_nodes = collect_block_nodes(root);

    let mut counters = HashMap::new();
    let mut last_seen_level = HashMap::new();
    let mut applied_overrides = HashSet::new();

    for node in top_nodes {
        if node.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "tbl" => {
                let table = parse_table_node(
                    node,
                    ctx,
                    &mut counters,
                    &mut last_seen_level,
                    &mut applied_overrides,
                );
                blocks.push(Block::Table(table));
            }
            "p" => {
                let para = super::paragraph::build_paragraph(
                    node, ctx, &mut counters, &mut last_seen_level,
                    &mut applied_overrides,
                    &super::paragraph::ParagraphOptions::default(),
                );
                blocks.push(Block::Paragraph(para));
            }
            _ => {}
        }
    }

    (!blocks.is_empty()).then(|| HeaderFooter { blocks })
}

pub(super) fn parse_footnotes<R: Read + Seek>(
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

    let mut fn_ctx = ParseContext {
        styles,
        theme,
        rels: &empty_rels,
        zip,
        numbering: &numbering,
    };

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
            let para_style = fn_ctx.styles.paragraph_styles.get(para_style_id);

            let alignment = resolve_alignment(ppr, para_style);
            let parsed = parse_runs(p, &mut fn_ctx);
            let (sp_before, sp_after, ls) = parse_paragraph_spacing(ppr, para_style, None);

            paragraphs.push(Paragraph {
                runs: parsed.runs,
                space_before: sp_before.unwrap_or(0.0),
                space_after: sp_after.unwrap_or(0.0),
                alignment,
                line_spacing: ls.or(Some(LineSpacing::Auto(1.0))),
                snap_to_grid: true,
                ..Paragraph::default()
            });
        }

        if !paragraphs.is_empty() {
            footnotes.insert(id, Footnote { paragraphs });
        }
    }

    footnotes
}
