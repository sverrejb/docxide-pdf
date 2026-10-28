use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{Alignment, Block, Footnote, HeaderFooter, LineSpacing, Paragraph, Run};

use super::numbering::{ListCounters, NumberingInfo};
use super::parse_table_node;
use super::relationships::parse_part_relationships;
use super::runs::parse_runs;
use super::styles::{StylesInfo, ThemeFonts};
use super::{
    ParseContext, WML_NS, collect_block_nodes, extract_indents, parse_paragraph_spacing, wml,
    wml_attr,
};

pub(super) fn parse_header_footer_xml<R: Read + Seek>(
    xml_content: &str,
    ctx: &mut ParseContext<'_, R>,
) -> Option<HeaderFooter> {
    let xml = roxmltree::Document::parse(xml_content).ok()?;
    let root = xml.root_element();
    let mut blocks = Vec::new();

    let top_nodes = collect_block_nodes(root);

    let mut lists = ListCounters::default();

    for node in top_nodes {
        if node.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "tbl" => {
                let table = parse_table_node(node, ctx, &mut lists);
                blocks.push(Block::Table(table));
            }
            "p" => {
                let mut para = super::paragraph::build_paragraph(
                    node,
                    ctx,
                    &mut lists,
                    &super::paragraph::ParagraphOptions::default(),
                );
                // Auto spacing opens no story, as at the top of the body
                // (a header line with beforeAutospacing sits at the header
                // distance).
                if blocks.is_empty() && para.space_before_auto {
                    para.space_before = 0.0;
                }
                blocks.push(Block::Paragraph(para));
            }
            _ => {}
        }
    }

    (!blocks.is_empty()).then_some(HeaderFooter { blocks })
}

pub(super) fn parse_footnotes<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> (HashMap<u32, Footnote>, Option<Paragraph>) {
    // Footnotes use the simple paragraph builder for backwards compatibility
    // with rendering tuned against the existing corpus; only list paragraphs
    // take the full builder.
    parse_notes_simple(
        zip,
        styles,
        theme,
        numbering,
        "word/footnotes.xml",
        "footnote",
        "FootnoteText",
    )
}

pub(super) fn parse_endnotes<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> (HashMap<u32, Footnote>, Option<Paragraph>) {
    parse_notes_rich(
        zip,
        styles,
        theme,
        numbering,
        "word/endnotes.xml",
        "endnote",
    )
}

fn parse_note_separator<R: Read + Seek>(
    node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    lists: &mut ListCounters,
) -> Option<Paragraph> {
    node.children()
        .find(|c| c.has_tag_name((WML_NS, "p")))
        .map(|p| {
            super::paragraph::build_paragraph(
                p,
                ctx,
                lists,
                &super::paragraph::ParagraphOptions::default(),
            )
        })
}

/// Simple parsing: paragraph runs and indents only, except that list
/// paragraphs take the full builder for their numbering (croatian_grant's
/// bulleted footnote). Matches the original footnote rendering behavior.
fn parse_notes_simple<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
    zip_path: &str,
    element_name: &str,
    default_style_id: &str,
) -> (HashMap<u32, Footnote>, Option<Paragraph>) {
    let mut footnotes = HashMap::new();
    let mut separator = None;
    let Some(xml_text) = super::read_zip_text(zip, zip_path) else {
        return (footnotes, separator);
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_text) else {
        return (footnotes, separator);
    };
    let root = xml.root_element();
    let rels = parse_part_relationships(zip, zip_path);
    let mut fn_ctx = ParseContext {
        styles,
        theme,
        rels: &rels,
        zip,
        numbering,
        comments: Default::default(),
        fields: Default::default(),
        grid_char_delta: 0.0,
    };
    let mut lists = ListCounters::default();

    for node in root.children() {
        if !node.has_tag_name((WML_NS, element_name)) {
            continue;
        }
        match node.attribute((WML_NS, "type")) {
            // A body paragraph in all but place: no pStyle means Normal.
            Some("separator") => {
                separator = parse_note_separator(node, &mut fn_ctx, &mut lists);
                continue;
            }
            Some(_) => continue,
            None => {}
        }
        let Some(id) = node
            .attribute((WML_NS, "id"))
            .and_then(|v| v.parse::<u32>().ok())
        else {
            continue;
        };

        let mut parse_para = |p: roxmltree::Node, fn_ctx: &mut ParseContext<'_, R>| -> Paragraph {
            let ppr = wml(p, "pPr");
            let para_style_id = ppr
                .and_then(|ppr| wml_attr(ppr, "pStyle"))
                .unwrap_or(default_style_id);
            let numbered = ppr.is_some_and(|ppr| wml(ppr, "numPr").is_some())
                || fn_ctx
                    .styles
                    .paragraph_styles
                    .get(para_style_id)
                    .is_some_and(|s| s.num_id.is_some());
            if numbered {
                let mut para = super::paragraph::build_paragraph(
                    p,
                    fn_ctx,
                    &mut lists,
                    &super::paragraph::ParagraphOptions::default(),
                );
                para.line_spacing = para.line_spacing.or(Some(LineSpacing::Auto(1.0)));
                para.snap_to_grid = true;
                return para;
            }
            // A style the document doesn't define falls back to its default
            // paragraph style, as in Word (a python-docx file has no
            // FootnoteText: its notes take Normal's 1.15 lines and space after).
            let para_style = fn_ctx
                .styles
                .paragraph_styles
                .get(para_style_id)
                .or_else(|| {
                    fn_ctx
                        .styles
                        .paragraph_styles
                        .get(&fn_ctx.styles.default_paragraph_style_id)
                });

            let alignment =
                super::runs::paragraph_alignment(p, ppr, para_style).unwrap_or(Alignment::Left);
            let parsed = parse_runs(p, fn_ctx);
            let (sp_before, sp_after, ls) =
                parse_paragraph_spacing(ppr, para_style, &fn_ctx.styles.defaults);

            // Indents: inline w:ind overrides the style, missing attributes
            // fall back to the (basedOn-resolved) style — same merge as body
            // paragraphs, minus list-numbering interplay.
            let (left, right, hanging, first) =
                if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
                    let (l, r, h, f) = extract_indents(ind, None);
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
            let mut indents = (
                left.unwrap_or(0.0),
                right.unwrap_or(0.0),
                hanging.unwrap_or(0.0),
                first.unwrap_or(0.0),
            );
            let direct_chars = ppr
                .and_then(|ppr| wml(ppr, "ind"))
                .map(super::extract_indent_chars)
                .unwrap_or_default();
            let style_chars = para_style.map(|s| s.indent_chars).unwrap_or_default();
            super::apply_indent_chars(
                &mut indents,
                super::IndentChars {
                    left: direct_chars.left.or(style_chars.left),
                    right: direct_chars.right.or(style_chars.right),
                    hanging: direct_chars.hanging.or(style_chars.hanging),
                    first_line: direct_chars.first_line.or(style_chars.first_line),
                },
                super::paragraph::indent_char_width(
                    &parsed.runs,
                    para_style,
                    &styles.defaults,
                    fn_ctx.grid_char_delta,
                ),
            );
            let (left, right, hanging, first) = indents;
            // uk_commercial_lease's FootnoteText hangs 567 twips with no tab
            // stops: Word's "1<tab>" text starts at that implicit stop.
            let mut tab_stops = super::resolve_tab_stops(ppr, para_style);
            super::paragraph::add_hanging_tab_stop(&mut tab_stops, left, hanging);

            Paragraph {
                runs: parsed.runs,
                space_before: sp_before.unwrap_or(styles.defaults.space_before),
                space_after: sp_after.unwrap_or(styles.defaults.space_after),
                tab_stops,
                alignment,
                line_spacing: ls.or(Some(styles.defaults.line_spacing)),
                snap_to_grid: true,
                indent_left: left,
                indent_right: right,
                indent_hanging: hanging,
                indent_first_line: first,
                ..Paragraph::default()
            }
        };

        let mut paragraphs = Vec::new();
        for child in node.children() {
            if child.has_tag_name((WML_NS, "p")) {
                paragraphs.push(parse_para(child, &mut fn_ctx));
            } else if child.has_tag_name((WML_NS, "tbl")) {
                // Some templates (FAR/ST4) lay a footnote out as a table: the reference
                // mark in a narrow first cell, the text in the second. Dropping the table
                // dropped the whole footnote (auditor_regulatory_report_template p2).
                // ponytail: each row becomes one paragraph with the cells joined by a
                // space; real column geometry needs Block support in Footnote.
                for tr in child.children().filter(|n| n.has_tag_name((WML_NS, "tr"))) {
                    let mut cell_paras = tr
                        .descendants()
                        .filter(|n| n.has_tag_name((WML_NS, "p")))
                        .map(|p| parse_para(p, &mut fn_ctx));
                    let Some(mut merged) = cell_paras.next() else {
                        continue;
                    };
                    for para in cell_paras {
                        if let Some(first) = para.runs.first() {
                            merged.runs.push(Run {
                                text: " ".to_string(),
                                font_size: first.font_size,
                                font_name: first.font_name.clone(),
                                east_asia_font_name: first.east_asia_font_name.clone(),
                                bold: first.bold,
                                italic: first.italic,
                                ..Run::default()
                            });
                        }
                        merged.runs.extend(para.runs);
                    }
                    paragraphs.push(merged);
                }
            }
        }

        if !paragraphs.is_empty() {
            footnotes.insert(id, Footnote { paragraphs });
        }
    }

    (footnotes, separator)
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
) -> (HashMap<u32, Footnote>, Option<Paragraph>) {
    let mut footnotes = HashMap::new();
    let mut separator = None;
    let Some(xml_text) = super::read_zip_text(zip, zip_path) else {
        return (footnotes, separator);
    };
    let rels = parse_part_relationships(zip, zip_path);
    let Ok(xml) = roxmltree::Document::parse(&xml_text) else {
        return (footnotes, separator);
    };
    let root = xml.root_element();

    let mut fn_ctx = ParseContext {
        styles,
        theme,
        rels: &rels,
        zip,
        numbering,
        comments: Default::default(),
        fields: Default::default(),
        grid_char_delta: 0.0,
    };

    let mut lists = ListCounters::default();

    for node in root.children() {
        if !node.has_tag_name((WML_NS, element_name)) {
            continue;
        }
        match node.attribute((WML_NS, "type")) {
            Some("separator") => {
                separator = parse_note_separator(node, &mut fn_ctx, &mut lists);
                continue;
            }
            Some(_) => continue,
            None => {}
        }
        let Some(id) = node
            .attribute((WML_NS, "id"))
            .and_then(|v| v.parse::<u32>().ok())
        else {
            continue;
        };

        let mut paragraphs = Vec::new();
        for p in node.children().filter(|n| n.has_tag_name((WML_NS, "p"))) {
            let mut para = super::paragraph::build_paragraph(
                p,
                &mut fn_ctx,
                &mut lists,
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

    (footnotes, separator)
}
