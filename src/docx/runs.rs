use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    FieldCode, FloatingImage, Run, Textbox, VertAlign,
};

use super::{WML_NS, highlight_color, parse_text_color, twips_to_pts, wml, wml_attr, wml_bool};
use super::images::{parse_run_drawing, RunDrawingResult};
use super::numbering::NumberingInfo;
use super::styles::{StylesInfo, ThemeFonts, resolve_font_from_node};
use super::textbox::parse_textbox_from_vml;

pub(super) struct ParsedRuns {
    pub(super) runs: Vec<Run>,
    pub(super) has_page_break: bool,
    pub(super) has_column_break: bool,
    pub(super) line_break_count: u32,
    pub(super) floating_images: Vec<FloatingImage>,
    pub(super) textboxes: Vec<Textbox>,
}

pub(super) fn parse_runs<R: Read + std::io::Seek>(
    para_node: roxmltree::Node,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    numbering: &NumberingInfo,
) -> ParsedRuns {
    let ppr = wml(para_node, "pPr");
    let para_style_id = ppr
        .and_then(|ppr| wml_attr(ppr, "pStyle"))
        .unwrap_or("Normal");
    let para_style = styles.paragraph_styles.get(para_style_id);

    let style_font_size = para_style
        .and_then(|s| s.font_size)
        .unwrap_or(styles.defaults.font_size);
    let style_font_name = para_style
        .and_then(|s| s.font_name.as_deref())
        .unwrap_or(&styles.defaults.font_name)
        .to_string();
    let style_bold = para_style.and_then(|s| s.bold).unwrap_or(false);
    let style_italic = para_style.and_then(|s| s.italic).unwrap_or(false);
    let style_caps = para_style.and_then(|s| s.caps).unwrap_or(false);
    let style_small_caps = para_style.and_then(|s| s.small_caps).unwrap_or(false);
    let style_vanish = para_style.and_then(|s| s.vanish).unwrap_or(false);
    let style_color: Option<[u8; 3]> = para_style.and_then(|s| s.color);

    const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn collect_run_nodes<'a>(
        parent: roxmltree::Node<'a, 'a>,
        rels: &HashMap<String, String>,
        out: &mut Vec<(roxmltree::Node<'a, 'a>, Option<String>)>,
    ) {
        const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        for child in parent.children() {
            let name = child.tag_name().name();
            let ns = child.tag_name().namespace();
            let is_wml = ns == Some(WML_NS);
            if is_wml && name == "r" {
                out.push((child, None));
            } else if is_wml && name == "hyperlink" {
                let url = child
                    .attribute((REL_NS, "id"))
                    .and_then(|rid| rels.get(rid))
                    .cloned();
                for n in child.children().filter(|n| {
                    n.tag_name().name() == "r" && n.tag_name().namespace() == Some(WML_NS)
                }) {
                    out.push((n, url.clone()));
                }
            } else if is_wml && name == "sdt" {
                if let Some(content) = wml(child, "sdtContent") {
                    collect_run_nodes(content, rels, out);
                }
            } else if ns == Some(MC_NS) && name == "AlternateContent" {
                let choice = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS) && n.tag_name().name() == "Choice"
                });
                let fallback = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS) && n.tag_name().name() == "Fallback"
                });
                let branch = choice.or(fallback);
                if let Some(branch) = branch {
                    collect_run_nodes(branch, rels, out);
                }
            }
        }
    }
    let mut run_nodes: Vec<(roxmltree::Node, Option<String>)> = Vec::new();
    collect_run_nodes(para_node, rels, &mut run_nodes);

    let mut runs = Vec::new();
    let mut floating_images: Vec<FloatingImage> = Vec::new();
    let mut textboxes: Vec<Textbox> = Vec::new();
    let mut has_page_break = false;
    let mut has_column_break = false;
    let mut line_break_count: u32 = 0;
    let mut in_field = false;
    let mut field_instr = String::new();

    for (run_node, hyperlink_url) in run_nodes {
        let rpr = wml(run_node, "rPr");

        let char_style = rpr
            .and_then(|n| wml_attr(n, "rStyle"))
            .and_then(|id| styles.character_styles.get(id));

        let font_size = rpr
            .and_then(|n| wml_attr(n, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0)
            .or_else(|| char_style.and_then(|cs| cs.font_size))
            .unwrap_or(style_font_size);

        let font_name = rpr
            .and_then(|n| wml(n, "rFonts"))
            .map(|rfonts| resolve_font_from_node(rfonts, theme, &style_font_name))
            .or_else(|| char_style.and_then(|cs| cs.font_name.clone()))
            .unwrap_or_else(|| style_font_name.clone());

        let bold = rpr
            .and_then(|n| wml_bool(n, "b"))
            .or_else(|| char_style.and_then(|cs| cs.bold))
            .unwrap_or(style_bold);
        let italic = rpr
            .and_then(|n| wml_bool(n, "i"))
            .or_else(|| char_style.and_then(|cs| cs.italic))
            .unwrap_or(style_italic);
        let underline = rpr
            .and_then(|n| {
                wml(n, "u")
                    .and_then(|u| u.attribute((WML_NS, "val")))
                    .map(|v| v != "none")
            })
            .or_else(|| char_style.and_then(|cs| cs.underline))
            .unwrap_or(false);
        let strikethrough = rpr
            .and_then(|n| wml_bool(n, "strike"))
            .or_else(|| char_style.and_then(|cs| cs.strikethrough))
            .unwrap_or(false);
        let dstrike = rpr
            .and_then(|n| wml_bool(n, "dstrike"))
            .unwrap_or(false);
        let char_spacing = rpr
            .and_then(|n| wml(n, "spacing"))
            .and_then(|n| n.attribute((WML_NS, "val")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .unwrap_or(0.0);
        let text_scale = rpr
            .and_then(|n| wml_attr(n, "w"))
            .and_then(|v| v.trim_end_matches('%').parse::<f32>().ok())
            .unwrap_or(100.0);
        let caps = rpr
            .and_then(|n| wml_bool(n, "caps"))
            .or_else(|| char_style.and_then(|cs| cs.caps))
            .unwrap_or(style_caps);
        let small_caps = rpr
            .and_then(|n| wml_bool(n, "smallCaps"))
            .or_else(|| char_style.and_then(|cs| cs.small_caps))
            .unwrap_or(style_small_caps);
        let vanish = rpr
            .and_then(|n| wml_bool(n, "vanish"))
            .or_else(|| char_style.and_then(|cs| cs.vanish))
            .unwrap_or(style_vanish);

        let color = rpr
            .and_then(|n| wml_attr(n, "color"))
            .and_then(parse_text_color)
            .or_else(|| char_style.and_then(|cs| cs.color))
            .or(style_color);

        let vertical_align = rpr
            .and_then(|n| wml_attr(n, "vertAlign"))
            .map(|v| match v {
                "superscript" => VertAlign::Superscript,
                "subscript" => VertAlign::Subscript,
                _ => VertAlign::Baseline,
            })
            .unwrap_or(VertAlign::Baseline);

        let highlight = rpr
            .and_then(|n| wml_attr(n, "highlight"))
            .and_then(highlight_color);

        // Iterate children in document order to handle w:t, w:tab, w:br, w:fldChar, w:instrText
        let mut pending_text = String::new();
        for child in run_node.children() {
            let child_ns = child.tag_name().namespace();
            // Handle mc:AlternateContent inside runs (e.g. textboxes)
            if child_ns == Some(MC_NS) && child.tag_name().name() == "AlternateContent" {
                let choice = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS) && n.tag_name().name() == "Choice"
                });
                let fallback = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS) && n.tag_name().name() == "Fallback"
                });
                if let Some(branch) = choice {
                    for drawing in branch.children().filter(|n| {
                        n.tag_name().namespace() == Some(WML_NS)
                            && n.tag_name().name() == "drawing"
                    }) {
                        match parse_run_drawing(drawing, rels, zip, styles, theme, numbering) {
                            Some(RunDrawingResult::Inline(img)) => {
                                runs.push(Run {
                                    text: String::new(),
                                    font_size,
                                    font_name: font_name.clone(),
                                    bold: false,
                                    italic: false,
                                    underline: false,
                                    strikethrough: false,
                                    dstrike: false,
                                    char_spacing: 0.0,
                                    text_scale: 100.0,
                                    caps: false,
                                    small_caps: false,
                                    vanish: false,
                                    color: None,
                                    is_tab: false,
                                    vertical_align: VertAlign::Baseline,
                                    field_code: None,
                                    hyperlink_url: None,
                                    highlight: None,
                                    inline_image: Some(img),
                                    footnote_id: None,
                                    is_footnote_ref_mark: false,
                                });
                            }
                            Some(RunDrawingResult::Floating(fi)) => {
                                floating_images.push(fi);
                            }
                            Some(RunDrawingResult::TextBox(tb)) => {
                                textboxes.push(tb);
                            }
                            None => {}
                        }
                    }
                } else if let Some(branch) = fallback {
                    for pict in branch.descendants().filter(|n| {
                        n.tag_name().namespace() == Some(WML_NS)
                            && n.tag_name().name() == "pict"
                    }) {
                        if let Some(tb) =
                            parse_textbox_from_vml(pict, rels, zip, styles, theme, numbering)
                        {
                            textboxes.push(tb);
                        }
                    }
                }
                continue;
            }
            if child_ns != Some(WML_NS) {
                continue;
            }
            match child.tag_name().name() {
                "fldChar" => {
                    match child.attribute((WML_NS, "fldCharType")) {
                        Some("begin") => {
                            // Flush pending text before entering field
                            if !pending_text.is_empty() {
                                runs.push(Run {
                                    text: std::mem::take(&mut pending_text),
                                    font_size,
                                    font_name: font_name.clone(),
                                    bold,
                                    italic,
                                    underline,
                                    strikethrough,
                                    dstrike,
                                    char_spacing,
                                    text_scale,
                                    caps,
                                    small_caps,
                                    vanish,
                                    color,
                                    is_tab: false,
                                    vertical_align,
                                    field_code: None,
                                    hyperlink_url: hyperlink_url.clone(),
                                    highlight,
                                    inline_image: None,
                                    footnote_id: None,
                                    is_footnote_ref_mark: false,
                                });
                            }
                            in_field = true;
                            field_instr.clear();
                        }
                        Some("end") => {
                            if in_field {
                                let trimmed = field_instr.trim();
                                let fc = if trimmed.eq_ignore_ascii_case("PAGE") {
                                    Some(FieldCode::Page)
                                } else if trimmed.eq_ignore_ascii_case("NUMPAGES") {
                                    Some(FieldCode::NumPages)
                                } else {
                                    None
                                };
                                if let Some(code) = fc {
                                    runs.push(Run {
                                        text: String::new(),
                                        font_size,
                                        font_name: font_name.clone(),
                                        bold,
                                        italic,
                                        underline: false,
                                        strikethrough: false,
                                        dstrike: false,
                                        char_spacing: 0.0,
                                        text_scale: 100.0,
                                        caps: false,
                                        small_caps: false,
                                        vanish: false,
                                        color,
                                        is_tab: false,
                                        vertical_align: VertAlign::Baseline,
                                        field_code: Some(code),
                                        hyperlink_url: hyperlink_url.clone(),
                                        highlight: None,
                                        inline_image: None,
                                        footnote_id: None,
                                        is_footnote_ref_mark: false,
                                    });
                                }
                                in_field = false;
                                field_instr.clear();
                            }
                        }
                        _ => {}
                    }
                }
                "instrText" if in_field => {
                    if let Some(t) = child.text() {
                        field_instr.push_str(t);
                    }
                }
                "t" if !in_field => {
                    if let Some(t) = child.text() {
                        // Word treats newlines in w:t as whitespace; only w:br creates line breaks
                        let normalized = t.replace('\n', " ");
                        pending_text.push_str(&normalized);
                    }
                }
                "tab" if !in_field => {
                    // Flush any pending text before the tab
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            dstrike,
                            char_spacing,
                            text_scale,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    // Insert tab marker run
                    runs.push(Run {
                        text: String::new(),
                        font_size,
                        font_name: font_name.clone(),
                        bold: false,
                        italic: false,
                        underline: false,
                        strikethrough: false,
                        dstrike: false,
                        char_spacing: 0.0,
                        text_scale: 100.0,
                        caps: false,
                        small_caps: false,
                        vanish: false,
                        color: None,
                        is_tab: true,
                        vertical_align: VertAlign::Baseline,
                        field_code: None,
                        hyperlink_url: None,
                        highlight: None,
                        inline_image: None,
                        footnote_id: None,
                        is_footnote_ref_mark: false,
                    });
                }
                "br" if !in_field => {
                    match child.attribute((WML_NS, "type")) {
                        Some("page") => has_page_break = true,
                        Some("column") => has_column_break = true,
                        _ => line_break_count += 1,
                    }
                }
                "drawing" if in_field => {}
                "drawing" => {
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            dstrike,
                            char_spacing,
                            text_scale,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    match parse_run_drawing(child, rels, zip, styles, theme, numbering) {
                        Some(RunDrawingResult::Inline(img)) => {
                            runs.push(Run {
                                text: String::new(),
                                font_size,
                                font_name: font_name.clone(),
                                bold: false,
                                italic: false,
                                underline: false,
                                strikethrough: false,
                                dstrike: false,
                                char_spacing: 0.0,
                                text_scale: 100.0,
                                caps: false,
                                small_caps: false,
                                vanish: false,
                                color: None,
                                is_tab: false,
                                vertical_align: VertAlign::Baseline,
                                field_code: None,
                                hyperlink_url: None,
                                highlight: None,
                                inline_image: Some(img),
                                footnote_id: None,
                                is_footnote_ref_mark: false,
                            });
                        }
                        Some(RunDrawingResult::Floating(fi)) => {
                            floating_images.push(fi);
                        }
                        Some(RunDrawingResult::TextBox(tb)) => {
                            textboxes.push(tb);
                        }
                        None => {}
                    }
                }
                "pict" if !in_field => {
                    if let Some(tb) =
                        parse_textbox_from_vml(child, rels, zip, styles, theme, numbering)
                    {
                        textboxes.push(tb);
                    }
                }
                "footnoteReference" if !in_field => {
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            dstrike,
                            char_spacing,
                            text_scale,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    if let Some(id) = child
                        .attribute((WML_NS, "id"))
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        runs.push(Run {
                            text: String::new(),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline: false,
                            strikethrough: false,
                            dstrike: false,
                            char_spacing: 0.0,
                            text_scale: 100.0,
                            caps: false,
                            small_caps: false,
                            vanish: false,
                            color,
                            is_tab: false,
                            vertical_align: VertAlign::Superscript,
                            field_code: None,
                            hyperlink_url: None,
                            highlight: None,
                            inline_image: None,
                            footnote_id: Some(id),
                            is_footnote_ref_mark: false,
                        });
                    }
                }
                "footnoteRef" if !in_field => {
                    if !pending_text.is_empty() {
                        runs.push(Run {
                            text: std::mem::take(&mut pending_text),
                            font_size,
                            font_name: font_name.clone(),
                            bold,
                            italic,
                            underline,
                            strikethrough,
                            dstrike,
                            char_spacing,
                            text_scale,
                            caps,
                            small_caps,
                            vanish,
                            color,
                            is_tab: false,
                            vertical_align,
                            field_code: None,
                            hyperlink_url: hyperlink_url.clone(),
                            highlight,
                            inline_image: None,
                            footnote_id: None,
                            is_footnote_ref_mark: false,
                        });
                    }
                    runs.push(Run {
                        text: String::new(),
                        font_size,
                        font_name: font_name.clone(),
                        bold,
                        italic,
                        underline: false,
                        strikethrough: false,
                        dstrike: false,
                        char_spacing: 0.0,
                        text_scale: 100.0,
                        caps: false,
                        small_caps: false,
                        vanish: false,
                        color,
                        is_tab: false,
                        vertical_align: VertAlign::Superscript,
                        field_code: None,
                        hyperlink_url: None,
                        highlight: None,
                        inline_image: None,
                        footnote_id: None,
                        is_footnote_ref_mark: true,
                    });
                }
                _ => {}
            }
        }
        // Flush remaining text
        if !pending_text.is_empty() {
            runs.push(Run {
                text: pending_text,
                font_size,
                font_name,
                bold,
                italic,
                underline,
                strikethrough,
                dstrike,
                char_spacing,
                text_scale,
                caps,
                small_caps,
                vanish,
                color,
                is_tab: false,
                vertical_align,
                field_code: None,
                hyperlink_url: hyperlink_url.clone(),
                highlight,
                inline_image: None,
                footnote_id: None,
                is_footnote_ref_mark: false,
            });
        }
    }

    if ppr
        .and_then(|ppr| wml_bool(ppr, "pageBreakBefore"))
        .unwrap_or(false)
    {
        has_page_break = true;
    }

    // Empty paragraphs with explicit font sizing in their paragraph mark (pPr/rPr)
    // need a synthetic run so the renderer computes the correct line height.
    if runs.is_empty() && !has_page_break {
        let mark_rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
        let has_explicit_sz = mark_rpr.and_then(|n| wml_attr(n, "sz")).is_some();
        if has_explicit_sz {
            let mark_font_size = mark_rpr
                .and_then(|n| wml_attr(n, "sz"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|hp| hp / 2.0)
                .unwrap_or(style_font_size);
            let mark_font_name = mark_rpr
                .and_then(|n| wml(n, "rFonts"))
                .map(|rfonts| resolve_font_from_node(rfonts, theme, &style_font_name))
                .unwrap_or_else(|| style_font_name.clone());
            runs.push(Run {
                text: String::new(),
                font_size: mark_font_size,
                font_name: mark_font_name,
                bold: style_bold,
                italic: style_italic,
                underline: false,
                strikethrough: false,
                dstrike: false,
                char_spacing: 0.0,
                text_scale: 100.0,
                caps: false,
                small_caps: false,
                vanish: false,
                color: None,
                highlight: None,
                is_tab: false,
                vertical_align: VertAlign::Baseline,
                field_code: None,
                hyperlink_url: None,
                inline_image: None,
                footnote_id: None,
                is_footnote_ref_mark: false,
            });
        }
    }

    // Word's paragraph mark (¶) uses the paragraph style's font even in empty
    // paragraphs; ensure we carry that font info so line height is correct.
    if runs.is_empty() {
        runs.push(Run {
            text: String::new(),
            font_size: style_font_size,
            font_name: style_font_name.clone(),
            bold: style_bold,
            italic: style_italic,
            underline: false,
            strikethrough: false,
            dstrike: false,
            char_spacing: 0.0,
            text_scale: 100.0,
            caps: false,
            small_caps: false,
            vanish: false,
            color: None,
            highlight: None,
            is_tab: false,
            vertical_align: VertAlign::Baseline,
            field_code: None,
            hyperlink_url: None,
            inline_image: None,
            footnote_id: None,
            is_footnote_ref_mark: false,
        });
    }

    ParsedRuns {
        runs,
        has_page_break,
        has_column_break,
        line_break_count,
        floating_images,
        textboxes,
    }
}
