use crate::model::{Paragraph, TabAlignment, TabStop};

use super::images::compute_drawing_info;
use super::numbering::{ListCounters, ListLabelInfo, parse_list_info};
use super::runs::{parse_runs, push_textbox};
use super::styles::{
    ParagraphStyle, StyleDefaults, half_points, parse_alignment, parse_font_size,
    resolve_font_from_node_opt,
};
use super::textbox::collect_textboxes_from_paragraph;
use super::{
    ParseContext, WML_NS, extract_indents, parse_frame_props, parse_paragraph_borders,
    parse_paragraph_spacing, wml, wml_attr, wml_bool,
};

/// Options controlling which paragraph features to resolve.
#[derive(Default)]
pub(super) struct ParagraphOptions {
    /// Whether to resolve bookmarks from the node
    pub resolve_bookmarks: bool,
    /// Whether to resolve outline_level
    pub resolve_outline_level: bool,
    /// Whether to resolve drawing info (images/charts/smartart) from the node
    pub resolve_drawings: bool,
    /// Whether to collect additional textboxes from the paragraph node
    pub collect_extra_textboxes: bool,
    /// Style-level numbering ID fallback (from paragraph style)
    pub style_num_id: Option<String>,
    /// Style-level numbering ilvl fallback
    pub style_num_ilvl: Option<u8>,
}

pub(super) fn build_paragraph<R: std::io::Read + std::io::Seek>(
    node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    lists: &mut ListCounters,
    opts: &ParagraphOptions,
) -> Paragraph {
    let ppr = wml(node, "pPr");

    let ppr_rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
    let paragraph_mark_vanish = ppr_rpr
        .and_then(|rpr| wml_bool(rpr, "vanish"))
        .unwrap_or(false);

    let para_style_id = ppr
        .and_then(|ppr| wml_attr(ppr, "pStyle"))
        .unwrap_or(&ctx.styles.default_paragraph_style_id);

    let para_style = ctx.styles.paragraph_styles.get(para_style_id);

    let mark_style = ppr_rpr
        .and_then(|rpr| wml_attr(rpr, "rStyle"))
        .and_then(|id| ctx.styles.character_styles.get(id));

    // The mark inherits like any run: its own rPr, then the paragraph style,
    // then docDefaults.
    let paragraph_mark_font_size = ppr_rpr
        .and_then(parse_font_size)
        .or_else(|| mark_style.and_then(|s| s.font_size))
        .or_else(|| para_style.and_then(|s| s.font_size))
        .or(Some(ctx.styles.defaults.font_size));
    let paragraph_mark_font_name = ppr_rpr
        .and_then(|rpr| wml(rpr, "rFonts"))
        .and_then(|rf| resolve_font_from_node_opt(rf, ctx.theme))
        .or_else(|| mark_style.and_then(|s| s.font_name.clone()))
        .or_else(|| para_style.and_then(|s| s.font_name.clone()))
        .or_else(|| Some(ctx.styles.defaults.font_name.clone()));
    let paragraph_mark_position = paragraph_mark_position(ppr_rpr, para_style);

    // A paragraph-level pBdr element overrides the style borders even when
    // all individual borders are set to val="none" (parsed as None).
    let borders = ppr
        .and_then(parse_paragraph_borders)
        .unwrap_or_else(|| para_style.map(|s| s.borders.clone()).unwrap_or_default());
    let (sp_before, sp_after, line_spacing) =
        parse_paragraph_spacing(ppr, para_style, &ctx.styles.defaults);
    let space_before = sp_before.unwrap_or(ctx.styles.defaults.space_before);
    let space_after = sp_after.unwrap_or(ctx.styles.defaults.space_after);

    let inline_shd_node = ppr.and_then(|ppr| wml(ppr, "shd"));
    let para_shading = if inline_shd_node.is_some() {
        // Inline w:shd present — use it even if fill="auto" (None), don't inherit
        inline_shd_node.and_then(super::shd_color)
    } else {
        // No inline w:shd — inherit from paragraph style
        para_style.and_then(|s| s.shading)
    };

    let style_color = para_style.and_then(|s| s.color);

    let alignment = ppr
        .and_then(|ppr| wml_attr(ppr, "jc"))
        .map(parse_alignment)
        .or_else(|| super::runs::display_math_alignment(node))
        .or_else(|| para_style.and_then(|s| s.alignment))
        .unwrap_or(ctx.styles.defaults.alignment);

    let contextual_spacing = contextual_spacing(ppr, para_style);

    let keep_next = ppr
        .and_then(|ppr| wml_bool(ppr, "keepNext"))
        .unwrap_or_else(|| para_style.and_then(|s| s.keep_next).unwrap_or(false));

    let keep_lines = ppr
        .and_then(|ppr| wml_bool(ppr, "keepLines"))
        .unwrap_or_else(|| para_style.and_then(|s| s.keep_lines).unwrap_or(false));

    let widow_control = ppr
        .and_then(|ppr| wml_bool(ppr, "widowControl"))
        .or_else(|| para_style.and_then(|s| s.widow_control))
        .unwrap_or(ctx.styles.defaults.widow_control);

    let snap_to_grid = ppr
        .and_then(|ppr| wml_bool(ppr, "snapToGrid"))
        .or_else(|| para_style.and_then(|s| s.snap_to_grid))
        .unwrap_or(true);

    let auto_space_de = ppr
        .and_then(|ppr| wml_bool(ppr, "autoSpaceDE"))
        .or_else(|| para_style.and_then(|s| s.auto_space_de))
        .unwrap_or(true);

    let auto_space_dn = ppr
        .and_then(|ppr| wml_bool(ppr, "autoSpaceDN"))
        .or_else(|| para_style.and_then(|s| s.auto_space_dn))
        .unwrap_or(true);

    let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
    let style_num = opts.style_num_id.as_deref();
    let style_ilvl = opts.style_num_ilvl;
    let numbering = parse_list_info(
        num_pr,
        style_num,
        style_ilvl,
        Some(para_style_id),
        &ctx.styles.paragraph_styles,
        ctx.numbering,
        lists,
    );
    let (indent_left, indent_right, indent_hanging, indent_first_line) =
        resolve_indents(ppr, para_style, &numbering, &ctx.styles.defaults);
    let ListLabelInfo {
        indent_left: _,
        indent_hanging: _,
        tab_stop: mut num_tab_stop,
        label: list_label,
        font: list_label_font,
        font_size: list_label_font_size,
        bold: list_label_bold,
        color: list_label_color,
        suff: list_label_suff,
        jc: list_label_jc,
        item: list_item,
    } = numbering;
    // Paragraph-level `<w:tab val="num" pos="..."/>` overrides the numbering
    // level's num tab (paired with a `clear` of the inherited value when
    // Word-authored). `pos="0"` is a Word sentinel meaning "disable the num
    // tab for this paragraph" — ignore it so we don't collapse text onto the
    // label.
    if let Some(tabs) = ppr.and_then(|ppr| wml(ppr, "tabs")) {
        for t in tabs.children().filter(|n| n.has_tag_name((WML_NS, "tab"))) {
            if t.attribute((WML_NS, "val")) == Some("num")
                && let Some(pos) = super::twips_attr(t, "pos")
                && pos > 0.0
            {
                num_tab_stop = Some(pos);
            }
        }
    }

    let parsed = parse_runs(node, ctx);
    let mut runs = parsed.runs;

    if let Some(color) = style_color {
        for run in &mut runs {
            run.color.get_or_insert(color);
        }
    }

    let mut tab_stops = super::resolve_tab_stops(ppr, para_style);
    // Add the numbering level's explicit tab stop so the label-text
    // gap matches Word (which uses this instead of the implicit
    // hanging-indent tab when it is closer).
    if let Some(nts) = num_tab_stop
        && !tab_stops.iter().any(|t| (t.position - nts).abs() < 0.5)
    {
        tab_stops.push(TabStop {
            position: nts,
            alignment: TabAlignment::Left,
            leader: None,
        });
        tab_stops.sort_by(|a, b| a.position.total_cmp(&b.position));
    }
    add_hanging_tab_stop(&mut tab_stops, indent_left, indent_hanging);

    let has_text = runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
    let inline_image_count = runs.iter().filter(|r| r.inline_image.is_some()).count();
    let has_inline_images = inline_image_count > 0;

    let floating_images = parsed.floating_images;

    let (para_image, mut content_height) = if !opts.resolve_drawings {
        (None, 0.0)
    } else if inline_image_count == 1 && !has_text {
        let img_run_idx = runs.iter().position(|r| r.inline_image.is_some());
        let img = img_run_idx.and_then(|i| runs[i].inline_image.take());
        let h = img
            .as_ref()
            .map(|i| i.display_height + i.layout_extra_height)
            .unwrap_or(0.0);
        (img, h)
    } else if has_inline_images && !has_text {
        // Multi-image-only paragraph: keep images in runs so the
        // line builder can lay them out side-by-side. Expose the
        // tallest image height for vertical sizing.
        let max_h = runs
            .iter()
            .filter_map(|r| {
                r.inline_image
                    .as_ref()
                    .map(|i| i.display_height + i.layout_extra_height)
            })
            .fold(0.0f32, f32::max);
        (None, max_h)
    } else if has_inline_images {
        (None, 0.0)
    } else {
        let drawing = compute_drawing_info(node, ctx.rels, ctx.zip);
        (drawing.image, drawing.height)
    };

    if let Some(ref ic) = parsed.inline_chart {
        content_height = content_height.max(ic.display_height);
    }
    for sa in &parsed.smartart {
        content_height = content_height.max(sa.display_height);
    }

    let outline_level = if opts.resolve_outline_level {
        ppr.and_then(|ppr| wml_attr(ppr, "outlineLvl"))
            .and_then(|v| v.parse::<u8>().ok())
            .filter(|&lvl| lvl <= 9)
            .or_else(|| para_style.and_then(|s| s.outline_level))
            // Level 9 is body text: it overrides the style's heading level.
            .filter(|&lvl| lvl <= 8)
    } else {
        None
    };

    let bookmarks: Vec<String> = if opts.resolve_bookmarks {
        node.children()
            .filter(|n| n.has_tag_name((WML_NS, "bookmarkStart")))
            .filter_map(|n| n.attribute((WML_NS, "name")).map(|s| s.to_string()))
            .collect()
    } else {
        vec![]
    };

    let textboxes = {
        let mut tbs = parsed.textboxes;
        if opts.collect_extra_textboxes {
            for tb in collect_textboxes_from_paragraph(node, ctx) {
                push_textbox(&floating_images, &mut tbs, tb);
            }
        }
        tbs
    };

    let (space_before_auto, space_after_auto) =
        super::autospacing(ppr, para_style, &ctx.styles.defaults);
    Paragraph {
        runs,
        style_id: Some(para_style_id.to_string()),
        space_before,
        space_after,
        space_before_auto,
        space_after_auto,
        content_height,
        alignment,
        indent_left,
        indent_right,
        indent_hanging,
        indent_first_line,
        list_label,
        list_label_font,
        list_label_font_size,
        list_label_bold,
        list_label_color,
        list_label_jc,
        list_label_suff: super::numbering::label_suffix(&list_label_suff),
        list_item,
        starts_toc_field: node.descendants().any(|n| {
            n.has_tag_name((WML_NS, "instrText"))
                && n.text().is_some_and(|t| t.trim_start().starts_with("TOC"))
        }),
        num_level_tab_stop: num_tab_stop,
        contextual_spacing,
        keep_next,
        keep_lines,
        widow_control,
        line_spacing,
        image: para_image,
        borders,
        shading: para_shading,
        // A direct pageBreakBefore, on or off, overrides the style's
        // (wa_child's Part 1 heading turns its Heading2 break off).
        page_break_before: parsed.has_explicit_page_break_before
            || ppr
                .and_then(|ppr| wml_bool(ppr, "pageBreakBefore"))
                .unwrap_or_else(|| {
                    para_style
                        .and_then(|s| s.page_break_before)
                        .unwrap_or(false)
                }),
        page_break_before_explicit: parsed.has_explicit_page_break_before,
        page_break_after: parsed.has_page_break_after,
        page_break_at: parsed.page_break_at,
        column_break_before: parsed.has_column_break,
        // split_at_page_break hands a mid-paragraph one to the continuation
        column_break_after: parsed.column_break_at,
        clears_floats: parsed.has_clear_break,
        tab_stops,
        floating_images,
        textboxes,
        connectors: parsed.connectors,
        inline_chart: parsed.inline_chart,
        smartart: parsed.smartart,
        horizontal_rule: parsed.horizontal_rule,
        is_section_break: false,
        bookmarks,
        outline_level,
        paragraph_mark_vanish,
        paragraph_mark_font_size,
        paragraph_mark_font_name,
        paragraph_mark_position,
        snap_to_grid,
        auto_space_de,
        auto_space_dn,
        frame_props: parse_frame_props(ppr, para_style.and_then(|s| s.frame_attrs.as_ref())),
    }
}

/// OOXML 17.3.1.38: a hanging indent implicitly sets a tab stop at the indent.
pub(super) fn add_hanging_tab_stop(tab_stops: &mut Vec<TabStop>, indent_left: f32, hanging: f32) {
    if hanging > 0.0
        && !tab_stops
            .iter()
            .any(|t| (t.position - indent_left).abs() < 0.5)
    {
        tab_stops.push(TabStop {
            position: indent_left,
            alignment: TabAlignment::Left,
            leader: None,
        });
        tab_stops.sort_by(|a, b| a.position.total_cmp(&b.position));
    }
}

/// Word moves the text after a page break inside a paragraph to the next page
/// and lays it out as the paragraph's continuation: no list label, no
/// first-line indent, no space before. The paragraph's space after, keep-next
/// and section break belong to its end, i.e. the continuation.
pub(super) fn split_at_page_break(para: &mut Paragraph) -> Option<Paragraph> {
    let at = para.page_break_at.take()?;
    let runs = para.runs.split_off(at);
    let rest = Paragraph {
        runs,
        column_break_before: std::mem::take(&mut para.column_break_after),
        style_id: para.style_id.clone(),
        space_after: std::mem::take(&mut para.space_after),
        space_after_auto: std::mem::take(&mut para.space_after_auto),
        alignment: para.alignment,
        indent_left: para.indent_left,
        indent_right: para.indent_right,
        contextual_spacing: para.contextual_spacing,
        keep_next: std::mem::take(&mut para.keep_next),
        keep_lines: para.keep_lines,
        widow_control: para.widow_control,
        line_spacing: para.line_spacing,
        borders: para.borders.clone(),
        shading: para.shading,
        tab_stops: para.tab_stops.clone(),
        paragraph_mark_vanish: para.paragraph_mark_vanish,
        paragraph_mark_font_size: para.paragraph_mark_font_size,
        paragraph_mark_font_name: para.paragraph_mark_font_name.clone(),
        paragraph_mark_position: para.paragraph_mark_position,
        snap_to_grid: para.snap_to_grid,
        auto_space_de: para.auto_space_de,
        auto_space_dn: para.auto_space_dn,
        ..Paragraph::default()
    };
    Some(rest)
}

/// The paragraph mark's `w:position`, from its own rPr or the paragraph style.
/// Shared by body and table-cell paragraphs.
pub(super) fn paragraph_mark_position(
    ppr_rpr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
) -> f32 {
    ppr_rpr
        .and_then(|rpr| half_points(rpr, "position"))
        .or_else(|| para_style.and_then(|s| s.position))
        .unwrap_or(0.0)
}

/// `w:contextualSpacing`, direct or from the style. Shared by body and
/// table-cell paragraphs.
pub(super) fn contextual_spacing(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
) -> bool {
    ppr.and_then(|ppr| wml_bool(ppr, "contextualSpacing"))
        .unwrap_or_else(|| {
            para_style
                .and_then(|s| s.contextual_spacing)
                .unwrap_or(false)
        })
}

/// A paragraph's indents (left, right, hanging, first line) from its direct
/// `w:ind`, its style and its numbering level, over the document defaults.
/// Shared by body and table-cell paragraphs.
pub(super) fn resolve_indents(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
    numbering: &ListLabelInfo,
    defaults: &StyleDefaults,
) -> (f32, f32, f32, f32) {
    let mut indent_left = numbering.indent_left;
    let mut indent_hanging = numbering.indent_hanging;
    let mut indent_first_line = defaults.indent_first_line;
    let mut indent_right = defaults.indent_right;
    let char_width_fs = para_style
        .and_then(|s| s.font_size)
        .unwrap_or(defaults.font_size);
    // Numbering-level ind (already in indent_left/indent_hanging) outranks
    // style ind (§17.9.27); only directly-specified attributes override it.
    // Numbering that comes from the paragraph style sits below an ind set on
    // the same style or one below it (§17.7.2): an AC Bullet style setting
    // 340/340 over its list level's 153/360 indents by 340 in Word; an ind
    // only on a style above the numPr stays under the level's.
    let style_ind_wins = ppr.and_then(|ppr| wml(ppr, "numPr")).is_none()
        && para_style.is_some_and(|s| s.ind_over_numbering);
    let numbering_ind = !style_ind_wins && (indent_left != 0.0 || indent_hanging != 0.0);
    let (left, right, hanging, first) = if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
        let (l, r, h, f) = extract_indents(ind, Some(char_width_fs / 2.0));
        // firstLine and hanging are one value: a direct either replaces the
        // style's both (indonesian's title, ind left=281 firstLine=0 over
        // Heading1's hanging=543, starts at 281 in Word).
        let first_hanging_direct = h.is_some() || f.is_some();
        // Merge: inline w:ind attributes override style, but missing
        // attributes fall back to the paragraph style values.
        if let Some(s) = para_style {
            if numbering_ind {
                let f = if first_hanging_direct {
                    f
                } else {
                    s.indent_first_line
                };
                (l, r.or(s.indent_right), h, f)
            } else if first_hanging_direct {
                (l.or(s.indent_left), r.or(s.indent_right), h, f)
            } else {
                (
                    l.or(s.indent_left),
                    r.or(s.indent_right),
                    s.indent_hanging,
                    s.indent_first_line,
                )
            }
        } else {
            (l, r, h, f)
        }
    } else if (numbering.label.is_empty() || style_ind_wins)
        && let Some(s) = para_style
    {
        (
            s.indent_left,
            s.indent_right,
            s.indent_hanging,
            s.indent_first_line,
        )
    } else {
        (None, None, None, None)
    };
    if let Some(v) = left {
        indent_left = v;
    } else if indent_left == 0.0 {
        indent_left = defaults.indent_left;
    }
    if let Some(v) = right {
        indent_right = v;
    }
    if let Some(v) = hanging {
        indent_hanging = v;
    } else if first.is_some() {
        indent_hanging = 0.0;
    } else if indent_hanging == 0.0 {
        indent_hanging = defaults.indent_hanging;
    }
    if let Some(v) = first {
        indent_first_line = v;
    } else if numbering_ind && hanging.is_none() && numbering.indent_hanging != 0.0 {
        // firstLine and hanging are one value (§17.3.1.12): the level's hanging
        // replaces an inherited first-line indent (CV's docDefaults firstLine=360
        // under a 284/284 bullet level puts its label at the margin in Word).
        indent_first_line = 0.0;
    }
    (indent_left, indent_right, indent_hanging, indent_first_line)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Run;

    #[test]
    fn hanging_indent_adds_one_stop_at_the_indent() {
        let mut stops = Vec::new();
        add_hanging_tab_stop(&mut stops, 28.35, 28.35);
        add_hanging_tab_stop(&mut stops, 28.35, 28.35);
        assert_eq!(stops.len(), 1);
        assert_eq!(stops[0].position, 28.35);
        add_hanging_tab_stop(&mut stops, 50.0, 0.0);
        assert_eq!(stops.len(), 1);
    }

    #[test]
    fn paragraph_mark_inherits_character_style_with_direct_overrides() {
        use std::io::{Cursor, Write};
        use zip::write::SimpleFileOptions;

        for (mark, font, size) in [
            (r#"<w:rStyle w:val="Child"/>"#, "Times New Roman", 11.0),
            (
                r#"<w:rStyle w:val="Child"/><w:sz w:val="28"/>"#,
                "Times New Roman",
                14.0,
            ),
            (
                r#"<w:rStyle w:val="Child"/><w:rFonts w:ascii="Arial"/>"#,
                "Arial",
                11.0,
            ),
            (r#"<w:rStyle w:val="Missing"/>"#, "Calibri", 12.0),
        ] {
            let document = format!(
                r#"<w:document xmlns:w="{WML_NS}"><w:body><w:p><w:pPr><w:rPr>{mark}</w:rPr></w:pPr></w:p><w:sectPr/></w:body></w:document>"#
            );
            let styles = format!(
                r#"<w:styles xmlns:w="{WML_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="character" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="character" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Base"/></w:style></w:styles>"#
            );
            let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
            for (name, xml) in [("word/document.xml", document), ("word/styles.xml", styles)] {
                zip.start_file(name, SimpleFileOptions::default()).unwrap();
                zip.write_all(xml.as_bytes()).unwrap();
            }
            let bytes = zip.finish().unwrap().into_inner();
            let parsed = crate::docx::parse_bytes(&bytes).unwrap();
            let crate::model::Block::Paragraph(para) = &parsed.sections[0].blocks[0] else {
                panic!("expected a paragraph");
            };
            assert_eq!(para.paragraph_mark_font_name.as_deref(), Some(font));
            assert_eq!(para.paragraph_mark_font_size, Some(size));
            assert_eq!(para.runs[0].font_name, font);
            assert_eq!(para.runs[0].font_size, size);
        }
    }

    #[test]
    fn column_break_split_starts_the_continuation_in_the_next_column() {
        let run = |t: &str| Run {
            text: t.into(),
            ..Run::default()
        };
        let mut para = Paragraph {
            runs: vec![run("before"), run("after")],
            page_break_at: Some(1),
            column_break_after: true,
            ..Paragraph::default()
        };
        let rest = split_at_page_break(&mut para).unwrap();
        assert!(!para.column_break_after && !para.page_break_after);
        assert!(rest.column_break_before && rest.runs[0].text == "after");
    }

    #[test]
    fn page_break_split_moves_the_paragraph_end_to_the_continuation() {
        let run = |t: &str| Run {
            text: t.into(),
            ..Run::default()
        };
        let mut para = Paragraph {
            runs: vec![run("before"), run("after")],
            page_break_at: Some(1),
            page_break_after: true,
            space_before: 6.0,
            space_after: 10.0,
            keep_next: true,
            indent_left: 18.0,
            indent_hanging: 18.0,
            list_label: "1.".into(),
            ..Paragraph::default()
        };
        let rest = split_at_page_break(&mut para).unwrap();
        assert_eq!(para.runs.len(), 1);
        assert_eq!(rest.runs[0].text, "after");
        assert!(para.page_break_after && !para.keep_next && para.space_after == 0.0);
        assert!(rest.keep_next && rest.space_after == 10.0 && rest.space_before == 0.0);
        assert_eq!((rest.indent_left, rest.indent_hanging), (18.0, 0.0));
        assert!(rest.list_label.is_empty() && !rest.page_break_after);
        assert!(split_at_page_break(&mut para).is_none());
    }
}
