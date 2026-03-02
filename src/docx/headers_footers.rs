use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    Alignment, Footnote, HeaderFooter, LineSpacing, Paragraph, ParagraphBorders,
};

use super::{WML_NS, twips_attr, wml, wml_attr};
use super::numbering::NumberingInfo;
use super::runs::parse_runs;
use super::styles::{StylesInfo, ThemeFonts, parse_alignment, parse_line_spacing};

pub(super) fn parse_header_footer_xml<R: Read + std::io::Seek>(
    xml_content: &str,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> Option<HeaderFooter> {
    let xml = roxmltree::Document::parse(xml_content).ok()?;
    let root = xml.root_element();
    let mut paragraphs = Vec::new();

    for node in root.children() {
        if node.tag_name().namespace() != Some(WML_NS) || node.tag_name().name() != "p" {
            continue;
        }
        let ppr = wml(node, "pPr");
        let para_style_id = ppr
            .and_then(|ppr| wml_attr(ppr, "pStyle"))
            .unwrap_or("Normal");
        let para_style = styles.paragraph_styles.get(para_style_id);

        let alignment = ppr
            .and_then(|ppr| wml_attr(ppr, "jc"))
            .map(parse_alignment)
            .or_else(|| para_style.and_then(|s| s.alignment))
            .unwrap_or(Alignment::Left);

        let parsed = parse_runs(node, styles, theme, rels, zip, &NumberingInfo::default());

        paragraphs.push(Paragraph {
            runs: parsed.runs,
            space_before: 0.0,
            space_after: 0.0,
            content_height: 0.0,
            alignment,
            indent_left: 0.0,
            indent_right: 0.0,
            indent_hanging: 0.0,
            indent_first_line: 0.0,
            list_label: String::new(),
            list_label_font: None,
            contextual_spacing: false,
            keep_next: false,
            keep_lines: false,
            line_spacing: None,
            image: None,
            borders: ParagraphBorders::default(),
            shading: None,
            page_break_before: false,
            column_break_before: false,
            tab_stops: vec![],
            extra_line_breaks: parsed.line_break_count,
            floating_images: parsed.floating_images,
            textboxes: parsed.textboxes,
        });
    }

    if paragraphs.is_empty() {
        None
    } else {
        Some(HeaderFooter { paragraphs })
    }
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

    for node in root.children() {
        if node.tag_name().namespace() != Some(WML_NS) || node.tag_name().name() != "footnote" {
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
        let empty_rels = HashMap::new();
        for p in node.children() {
            if p.tag_name().namespace() != Some(WML_NS) || p.tag_name().name() != "p" {
                continue;
            }
            let ppr = wml(p, "pPr");
            let para_style_id = ppr
                .and_then(|ppr| wml_attr(ppr, "pStyle"))
                .unwrap_or("FootnoteText");
            let para_style = styles.paragraph_styles.get(para_style_id);

            let alignment = ppr
                .and_then(|ppr| wml_attr(ppr, "jc"))
                .map(parse_alignment)
                .or_else(|| para_style.and_then(|s| s.alignment))
                .unwrap_or(Alignment::Left);

            let parsed = parse_runs(p, styles, theme, &empty_rels, zip, &NumberingInfo::default());

            let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
            let space_before = inline_spacing
                .and_then(|n| twips_attr(n, "before"))
                .or_else(|| para_style.and_then(|s| s.space_before))
                .unwrap_or(0.0);
            let space_after = inline_spacing
                .and_then(|n| twips_attr(n, "after"))
                .or_else(|| para_style.and_then(|s| s.space_after))
                .unwrap_or(0.0);
            let line_spacing = inline_spacing
                .and_then(|n| {
                    n.attribute((WML_NS, "line"))
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|line_val| parse_line_spacing(n, line_val))
                })
                .or_else(|| para_style.and_then(|s| s.line_spacing))
                .or(Some(LineSpacing::Auto(1.0)));

            paragraphs.push(Paragraph {
                runs: parsed.runs,
                space_before,
                space_after,
                content_height: 0.0,
                alignment,
                indent_left: 0.0,
                indent_right: 0.0,
                indent_hanging: 0.0,
                indent_first_line: 0.0,
                list_label: String::new(),
                list_label_font: None,
                contextual_spacing: false,
                keep_next: false,
                keep_lines: false,
                line_spacing,
                image: None,
                borders: ParagraphBorders::default(),
                shading: None,
                page_break_before: false,
                column_break_before: false,
                tab_stops: vec![],
                extra_line_breaks: parsed.line_break_count,
                floating_images: vec![],
                textboxes: vec![],
            });
        }

        if !paragraphs.is_empty() {
            footnotes.insert(id, Footnote { paragraphs });
        }
    }

    footnotes
}
