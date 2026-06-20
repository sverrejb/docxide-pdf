use std::collections::{HashMap, HashSet};

use crate::model::{Paragraph, Run, TabAlignment, TabStop};

use super::images::compute_drawing_info;
use super::numbering::{ListLabelInfo, parse_list_info};
use super::runs::parse_runs;
use super::styles::{parse_alignment, parse_font_size, rfonts_ascii_name};
use super::textbox::collect_textboxes_from_paragraph;
use super::{
    ParseContext, WML_NS, extract_indents, parse_frame_props, parse_hex_color,
    merge_tab_stops, parse_paragraph_borders, parse_paragraph_spacing, parse_tab_stops_with_clears,
    wml, wml_attr,
    wml_bool,
};

/// Options controlling which paragraph features to resolve.
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

impl Default for ParagraphOptions {
    fn default() -> Self {
        Self {
            resolve_bookmarks: false,
            resolve_outline_level: false,
            resolve_drawings: false,
            collect_extra_textboxes: false,
            style_num_id: None,
            style_num_ilvl: None,
        }
    }
}

pub(super) fn build_paragraph<R: std::io::Read + std::io::Seek>(
    node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    counters: &mut HashMap<(u32, u8), u32>,
    last_seen_level: &mut HashMap<u32, u8>,
    applied_overrides: &mut HashSet<(u32, u8)>,
    opts: &ParagraphOptions,
) -> Paragraph {
    let ppr = wml(node, "pPr");

    let ppr_rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
    let paragraph_mark_vanish = ppr_rpr
        .and_then(|rpr| wml_bool(rpr, "vanish"))
        .unwrap_or(false);
    let paragraph_mark_font_size = ppr_rpr.and_then(parse_font_size);
    let paragraph_mark_font_name = ppr_rpr.and_then(rfonts_ascii_name);

    let para_style_id = ppr
        .and_then(|ppr| wml_attr(ppr, "pStyle"))
        .unwrap_or(&ctx.styles.default_paragraph_style_id);

    let para_style = ctx.styles.paragraph_styles.get(para_style_id);

    // A paragraph-level pBdr element overrides the style borders even when
    // all individual borders are set to val="none" (parsed as None).
    let borders = ppr
        .and_then(parse_paragraph_borders)
        .unwrap_or_else(|| para_style.map(|s| s.borders.clone()).unwrap_or_default());
    let (sp_before, sp_after, line_spacing) = parse_paragraph_spacing(ppr, para_style, None);
    let space_before = sp_before.unwrap_or(0.0);
    let space_after = sp_after.unwrap_or(ctx.styles.defaults.space_after);

    let inline_shd_node = ppr.and_then(|ppr| wml(ppr, "shd"));
    let para_shading = if inline_shd_node.is_some() {
        // Inline w:shd present — use it even if fill="auto" (None), don't inherit
        inline_shd_node
            .and_then(|shd| shd.attribute((WML_NS, "fill")))
            .and_then(parse_hex_color)
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

    let contextual_spacing = ppr
        .and_then(|ppr| wml_bool(ppr, "contextualSpacing"))
        .unwrap_or_else(|| para_style.is_some_and(|s| s.contextual_spacing));

    let keep_next = ppr
        .and_then(|ppr| wml_bool(ppr, "keepNext"))
        .unwrap_or_else(|| para_style.is_some_and(|s| s.keep_next));

    let keep_lines = ppr
        .and_then(|ppr| wml_bool(ppr, "keepLines"))
        .unwrap_or_else(|| para_style.is_some_and(|s| s.keep_lines));

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

    let suppress_auto_hyphens = ppr
        .and_then(|ppr| wml_bool(ppr, "suppressAutoHyphens"))
        .or_else(|| para_style.and_then(|s| s.suppress_auto_hyphens))
        .unwrap_or(false);

    let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
    let style_num = opts.style_num_id.as_deref();
    let style_ilvl = opts.style_num_ilvl;
    let ListLabelInfo {
        mut indent_left,
        mut indent_hanging,
        tab_stop: mut num_tab_stop,
        label: mut list_label,
        font: mut list_label_font,
        font_size: mut list_label_font_size,
        bold: mut list_label_bold,
        color: mut list_label_color,
        suff: list_label_suff,
    } = parse_list_info(
        num_pr,
        style_num,
        style_ilvl,
        Some(para_style_id),
        ctx.numbering,
        counters,
        last_seen_level,
        applied_overrides,
    );
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

    let mut indent_first_line = ctx.styles.defaults.indent_first_line;
    let mut indent_right = ctx.styles.defaults.indent_right;
    let char_width_fs = para_style
        .and_then(|s| s.font_size)
        .unwrap_or(ctx.styles.defaults.font_size);
    let (left, right, hanging, first) =
        if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
            let (l, r, h, f) = extract_indents(ind, Some(char_width_fs / 2.0));
            // Merge: inline w:ind attributes override style, but missing
            // attributes fall back to the paragraph style values.
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
        } else if list_label.is_empty()
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
        indent_left = ctx.styles.defaults.indent_left;
    }
    if let Some(v) = right {
        indent_right = v;
    }
    if let Some(v) = hanging {
        indent_hanging = v;
    } else if first.is_some() {
        indent_hanging = 0.0;
    } else if indent_hanging == 0.0 {
        indent_hanging = ctx.styles.defaults.indent_hanging;
    }
    if let Some(v) = first {
        indent_first_line = v;
    }

    let parsed = parse_runs(node, ctx);
    let mut runs = parsed.runs;

    // When w:suff="nothing", the label is inline with text -- prepend
    // it as a run rather than rendering it separately in the margin.
    if !list_label.is_empty() && list_label_suff == "nothing" {
        let first_run = runs.first();
        let label_run = Run {
            text: list_label.clone(),
            font_size: list_label_font_size
                .unwrap_or_else(|| first_run.map(|r| r.font_size).unwrap_or(10.0)),
            font_name: list_label_font
                .clone()
                .unwrap_or_else(|| first_run.map(|r| r.font_name.clone()).unwrap_or_default()),
            bold: list_label_bold,
            color: list_label_color,
            ..Run::default()
        };
        runs.insert(0, label_run);
        list_label = String::new();
        list_label_font = None;
        list_label_font_size = None;
        list_label_bold = false;
        list_label_color = None;
    }

    if let Some(color) = style_color {
        for run in &mut runs {
            run.color.get_or_insert(color);
        }
    }

    let mut tab_stops = if let Some(s) = para_style {
        s.tab_stops.clone()
    } else {
        vec![]
    };
    let (para_tabs, para_clears) = ppr
        .map(parse_tab_stops_with_clears)
        .unwrap_or_default();
    if !para_tabs.is_empty() || !para_clears.is_empty() {
        merge_tab_stops(&mut tab_stops, &para_clears, para_tabs);
        tab_stops.sort_by(|a, b| a.position.total_cmp(&b.position));
    }
    // Add the numbering level's explicit tab stop so the label-text
    // gap matches Word (which uses this instead of the implicit
    // hanging-indent tab when it is closer).
    if let Some(nts) = num_tab_stop {
        if !tab_stops.iter().any(|t| (t.position - nts).abs() < 0.5) {
            tab_stops.push(TabStop {
                position: nts,
                alignment: TabAlignment::Left,
                leader: None,
            });
            tab_stops.sort_by(|a, b| a.position.total_cmp(&b.position));
        }
    }
    // OOXML 17.3.1.38: hanging indent implicitly creates a tab stop
    if indent_hanging > 0.0 {
        let hang_pos = indent_left;
        if !tab_stops
            .iter()
            .any(|t| (t.position - hang_pos).abs() < 0.5)
        {
            tab_stops.push(TabStop {
                position: hang_pos,
                alignment: TabAlignment::Left,
                leader: None,
            });
            tab_stops.sort_by(|a, b| a.position.total_cmp(&b.position));
        }
    }

    let has_text = runs.iter().any(|r| !r.text.is_empty() || r.is_tab);
    let inline_image_count = runs.iter().filter(|r| r.inline_image.is_some()).count();
    let has_inline_images = inline_image_count > 0;

    let mut floating_images = parsed.floating_images;

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
        floating_images.extend(drawing.floating_images);
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
            .filter(|&lvl| lvl <= 8)
            .or_else(|| para_style.and_then(|s| s.outline_level))
    } else {
        None
    };

    let bookmarks: Vec<String> = if opts.resolve_bookmarks {
        node.children()
            .filter(|n| {
                n.tag_name().namespace() == Some(WML_NS)
                    && n.tag_name().name() == "bookmarkStart"
            })
            .filter_map(|n| n.attribute((WML_NS, "name")).map(|s| s.to_string()))
            .collect()
    } else {
        vec![]
    };

    let textboxes = {
        let mut tbs = parsed.textboxes;
        if opts.collect_extra_textboxes {
            tbs.extend(collect_textboxes_from_paragraph(node, ctx));
        }
        tbs
    };

    Paragraph {
        runs,
        style_id: Some(para_style_id.to_string()),
        space_before,
        space_after,
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
        num_level_tab_stop: num_tab_stop,
        contextual_spacing,
        keep_next,
        keep_lines,
        widow_control,
        line_spacing,
        image: para_image,
        borders,
        shading: para_shading,
        page_break_before: parsed.has_page_break_before
            || para_style.is_some_and(|s| s.page_break_before),
        page_break_before_explicit: parsed.has_explicit_page_break_before,
        page_break_after: parsed.has_page_break_after,
        column_break_before: parsed.has_column_break,
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
        snap_to_grid,
        auto_space_de,
        auto_space_dn,
        suppress_auto_hyphens,
        frame_props: ppr.and_then(parse_frame_props),
    }
}
