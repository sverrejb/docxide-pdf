use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use crate::model::{Alignment, Block, Footnote, HeaderFooter, LineSpacing, Paragraph};

use super::numbering::NumberingInfo;
use super::parse_table_node;
use super::relationships::parse_part_relationships;
use super::runs::parse_runs;
use super::styles::{ParagraphStyle, StylesInfo, ThemeFonts, parse_alignment};
use super::{
    ParseContext, WML_NS, collect_block_nodes, extract_indents, is_wml, parse_paragraph_spacing,
    wml, wml_attr,
};

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
    _numbering: &NumberingInfo,
) -> HashMap<u32, Footnote> {
    // Footnotes use the simple paragraph builder for backwards compatibility
    // with rendering tuned against the existing corpus. Bullets/lists inside
    // footnotes are not exercised by current fixtures.
    parse_notes_simple(zip, styles, theme, "word/footnotes.xml", "footnote", "FootnoteText")
}

pub(super) fn parse_endnotes<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> HashMap<u32, Footnote> {
    parse_notes_rich(zip, styles, theme, numbering, "word/endnotes.xml", "endnote")
}

/// Simple parsing: paragraph runs and indents only, no list-numbering.
/// Matches the original footnote rendering behavior.
fn parse_notes_simple<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    zip_path: &str,
    element_name: &str,
    default_style_id: &str,
) -> HashMap<u32, Footnote> {
    let mut footnotes = HashMap::new();
    let Some(xml_text) = super::read_zip_text(zip, zip_path) else {
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
        if !is_wml(node, element_name) {
            continue;
        }
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
        for p in node.children().filter(|n| is_wml(*n, "p")) {
            let ppr = wml(p, "pPr");
            let para_style_id = ppr
                .and_then(|ppr| wml_attr(ppr, "pStyle"))
                .unwrap_or(default_style_id);
            let para_style = fn_ctx.styles.paragraph_styles.get(para_style_id);

            let alignment = resolve_alignment(ppr, para_style);
            let parsed = parse_runs(p, &mut fn_ctx);
            let (sp_before, sp_after, ls) = parse_paragraph_spacing(ppr, para_style, None);

            // Indents: inline w:ind overrides the style, missing attributes
            // fall back to the (basedOn-resolved) style — same merge as body
            // paragraphs, minus list-numbering interplay.
            let char_fs = para_style
                .and_then(|s| s.font_size)
                .unwrap_or(styles.defaults.font_size);
            let (left, right, hanging, first) =
                if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                    let (l, r, h, f) = extract_indents(ind, Some(char_fs / 2.0));
                    if let Some(s) = para_style {
                        (
                            l.or(s.indent_left),
                            r.or(s.indent_right),
                            h.or(s.indent_hanging),
                            f.or(s.indent_first_line),
                        )
                    } else {
                        (l, r, h, f)
                    }
                } else if let Some(s) = para_style {
                    (
                        s.indent_left,
                        s.indent_right,
                        s.indent_hanging,
                        s.indent_first_line,
                    )
                } else {
                    (None, None, None, None)
                };

            paragraphs.push(Paragraph {
                runs: parsed.runs,
                space_before: sp_before.unwrap_or(0.0),
                space_after: sp_after.unwrap_or(0.0),
                alignment,
                line_spacing: ls.or(Some(LineSpacing::Auto(1.0))),
                snap_to_grid: true,
                indent_left: left.unwrap_or(0.0),
                indent_right: right.unwrap_or(0.0),
                indent_hanging: hanging.unwrap_or(0.0),
                indent_first_line: first.unwrap_or(0.0),
                ..Paragraph::default()
            });
        }

        if !paragraphs.is_empty() {
            footnotes.insert(id, Footnote { paragraphs });
        }
    }

    footnotes
}

/// Rich parsing: full paragraph builder with numbering, list labels,
/// hyperlinks resolved against the notes' own relationships file. Used for
/// endnotes which can contain bulleted lists and hyperlinked URLs.
fn parse_notes_rich<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
    zip_path: &str,
    element_name: &str,
) -> HashMap<u32, Footnote> {
    let mut footnotes = HashMap::new();
    let Some(xml_text) = super::read_zip_text(zip, zip_path) else {
        return footnotes;
    };
    let rels = parse_part_relationships(zip, zip_path);
    let Ok(xml) = roxmltree::Document::parse(&xml_text) else {
        return footnotes;
    };
    let root = xml.root_element();

    let mut fn_ctx = ParseContext {
        styles,
        theme,
        rels: &rels,
        zip,
        numbering,
    };

    let mut counters = HashMap::new();
    let mut last_seen_level = HashMap::new();
    let mut applied_overrides = HashSet::new();

    for node in root.children() {
        if !is_wml(node, element_name) {
            continue;
        }
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
        for p in node.children().filter(|n| is_wml(*n, "p")) {
            let mut para = super::paragraph::build_paragraph(
                p,
                &mut fn_ctx,
                &mut counters,
                &mut last_seen_level,
                &mut applied_overrides,
                &super::paragraph::ParagraphOptions::default(),
            );
            // Match the simple-path default so endnote line-heights stay
            // tight and aren't expanded by Word's body line-spacing default.
            if para.line_spacing.is_none() {
                para.line_spacing = Some(LineSpacing::Auto(1.0));
            }
            para.snap_to_grid = true;
            paragraphs.push(para);
        }

        if !paragraphs.is_empty() {
            footnotes.insert(id, Footnote { paragraphs });
        }
    }

    footnotes
}
