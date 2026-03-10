use std::collections::HashMap;
use std::io::Read;

use crate::model::{FieldCode, FloatingImage, InlineChart, Run, SmartArtDiagram, Textbox, VertAlign};

use super::images::{RunDrawingResult, parse_run_drawing};
use super::is_east_asian_char;
use super::numbering::NumberingInfo;
use super::styles::{
    StylesInfo, ThemeFonts, resolve_east_asia_font_from_node, resolve_font_from_node,
};
use super::textbox::parse_textbox_from_vml;
use super::{WML_NS, highlight_color, parse_text_color, twips_to_pts, wml, wml_attr, wml_bool};

fn is_dynamic_field(instr: &str) -> bool {
    let keyword = instr.split_whitespace().next().unwrap_or("");
    keyword.eq_ignore_ascii_case("PAGE")
        || keyword.eq_ignore_ascii_case("NUMPAGES")
        || keyword.eq_ignore_ascii_case("STYLEREF")
}

fn parse_styleref_arg(instr: &str) -> Option<String> {
    let trimmed = instr.trim();
    let kw = trimmed.split_whitespace().next()?;
    if !kw.eq_ignore_ascii_case("styleref") {
        return None;
    }
    let rest = trimmed[kw.len()..].trim();
    if let Some(quoted) = rest.strip_prefix('"') {
        let end = quoted.find('"')?;
        Some(quoted[..end].to_string())
    } else {
        Some(rest.split_whitespace().next()?.to_string())
    }
}

pub(super) struct ParsedRuns {
    pub(super) runs: Vec<Run>,
    pub(super) has_page_break: bool,
    pub(super) has_column_break: bool,
    pub(super) line_break_count: u32,
    pub(super) floating_images: Vec<FloatingImage>,
    pub(super) textboxes: Vec<Textbox>,
    pub(super) inline_chart: Option<InlineChart>,
    pub(super) smartart: Option<SmartArtDiagram>,
}

/// Resolved formatting for the current run, used to build Run structs concisely.
struct RunFormat {
    font_size: f32,
    font_name: String,
    east_asia_font_name: Option<String>,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    dstrike: bool,
    char_spacing: f32,
    text_scale: f32,
    caps: bool,
    small_caps: bool,
    vanish: bool,
    color: Option<[u8; 3]>,
    vertical_align: VertAlign,
    highlight: Option<[u8; 3]>,
    kern_threshold: Option<f32>,
    char_style_id: Option<String>,
}

impl RunFormat {
    /// Build a text run with the full formatting applied.
    fn text_run(&self, text: String, hyperlink_url: Option<String>) -> Run {
        Run {
            text,
            font_size: self.font_size,
            font_name: self.font_name.clone(),
            east_asia_font_name: self.east_asia_font_name.clone(),
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
            strikethrough: self.strikethrough,
            dstrike: self.dstrike,
            char_spacing: self.char_spacing,
            text_scale: self.text_scale,
            caps: self.caps,
            small_caps: self.small_caps,
            vanish: self.vanish,
            color: self.color,
            vertical_align: self.vertical_align,
            highlight: self.highlight,
            kern_threshold: self.kern_threshold,
            char_style_id: self.char_style_id.clone(),
            hyperlink_url,
            ..Run::default()
        }
    }

    /// Build a minimal run that only carries font identity (for images, tabs, field codes).
    fn minimal_run(&self) -> Run {
        Run {
            font_size: self.font_size,
            font_name: self.font_name.clone(),
            ..Run::default()
        }
    }
}

fn split_run_by_script(run: Run) -> Vec<Run> {
    let ea_font = match &run.east_asia_font_name {
        Some(f) if f != &run.font_name => f.clone(),
        _ => return vec![run],
    };

    let text = &run.text;
    if text.is_empty() {
        return vec![run];
    }

    let mut result: Vec<Run> = Vec::new();
    let mut segment_start = 0;
    let mut in_ea = false;
    let mut first = true;

    for (i, ch) in text.char_indices() {
        let ch_is_ea = if ch.is_whitespace() {
            // Whitespace inherits current script context
            in_ea
        } else {
            is_east_asian_char(ch)
        };

        if first {
            in_ea = ch_is_ea;
            first = false;
            continue;
        }

        if ch_is_ea != in_ea {
            let segment = &text[segment_start..i];
            if !segment.is_empty() {
                let mut sub = run.clone();
                sub.text = segment.to_string();
                if in_ea {
                    sub.font_name = ea_font.clone();
                }
                sub.east_asia_font_name = None;
                result.push(sub);
            }
            segment_start = i;
            in_ea = ch_is_ea;
        }
    }

    let segment = &text[segment_start..];
    if !segment.is_empty() {
        let mut sub = run.clone();
        sub.text = segment.to_string();
        if in_ea {
            sub.font_name = ea_font;
        }
        sub.east_asia_font_name = None;
        result.push(sub);
    }

    if result.is_empty() {
        let mut r = run;
        r.east_asia_font_name = None;
        vec![r]
    } else {
        result
    }
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
        .unwrap_or(&styles.default_paragraph_style_id);
    let para_style = styles.paragraph_styles.get(para_style_id);

    let style_font_size = para_style
        .and_then(|s| s.font_size)
        .unwrap_or(styles.defaults.font_size);
    let style_font_name = para_style
        .and_then(|s| s.font_name.as_deref())
        .unwrap_or(&styles.defaults.font_name)
        .to_string();
    let style_bold = para_style
        .and_then(|s| s.bold)
        .unwrap_or(styles.defaults.bold);
    let style_italic = para_style
        .and_then(|s| s.italic)
        .unwrap_or(styles.defaults.italic);
    let style_caps = para_style
        .and_then(|s| s.caps)
        .unwrap_or(styles.defaults.caps);
    let style_small_caps = para_style
        .and_then(|s| s.small_caps)
        .unwrap_or(styles.defaults.small_caps);
    let style_vanish = para_style
        .and_then(|s| s.vanish)
        .unwrap_or(styles.defaults.vanish);
    let style_underline = para_style
        .and_then(|s| s.underline)
        .unwrap_or(styles.defaults.underline);
    let style_strikethrough = para_style
        .and_then(|s| s.strikethrough)
        .unwrap_or(styles.defaults.strikethrough);
    let style_dstrike = para_style
        .and_then(|s| s.dstrike)
        .unwrap_or(styles.defaults.dstrike);
    let style_color: Option<[u8; 3]> = para_style.and_then(|s| s.color).or(styles.defaults.color);
    let style_char_spacing = para_style
        .and_then(|s| s.char_spacing)
        .unwrap_or(styles.defaults.char_spacing);
    let style_kern_threshold: Option<f32> = para_style
        .and_then(|s| s.kern_threshold)
        .or(styles.defaults.kern_threshold);
    let style_east_asia_font: Option<&str> = para_style
        .and_then(|s| s.east_asia_font.as_deref())
        .or(styles.defaults.east_asia_font.as_deref());

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
            } else if is_wml && name == "ins" {
                collect_run_nodes(child, rels, out);
            } else if is_wml && name == "del" {
                // Final mode: skip deleted content entirely
            } else if is_wml && name == "smartTag" {
                collect_run_nodes(child, rels, out);
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
                if let Some(branch) = choice.or(fallback) {
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
    let mut inline_chart: Option<InlineChart> = None;
    let mut smartart: Option<SmartArtDiagram> = None;
    let mut has_page_break = false;
    let mut has_column_break = false;
    let mut line_break_count: u32 = 0;
    let mut in_field = false;
    let mut in_field_result = false;
    let mut field_instr = String::new();
    let mut field_result_text = String::new();

    for (run_node, hyperlink_url) in run_nodes {
        let rpr = wml(run_node, "rPr");

        let char_style_id_str = rpr.and_then(|n| wml_attr(n, "rStyle"));
        let char_style = char_style_id_str.and_then(|id| styles.character_styles.get(id));

        let fmt = RunFormat {
            font_size: rpr
                .and_then(|n| wml_attr(n, "sz"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|hp| hp / 2.0)
                .or_else(|| char_style.and_then(|cs| cs.font_size))
                .unwrap_or(style_font_size),
            font_name: {
                let rfonts_node = rpr.and_then(|n| wml(n, "rFonts"));
                rfonts_node
                    .map(|rfonts| resolve_font_from_node(rfonts, theme, &style_font_name))
                    .or_else(|| char_style.and_then(|cs| cs.font_name.clone()))
                    .unwrap_or_else(|| style_font_name.clone())
            },
            east_asia_font_name: {
                let rfonts_node = rpr.and_then(|n| wml(n, "rFonts"));
                rfonts_node
                    .and_then(|rfonts| resolve_east_asia_font_from_node(rfonts, theme))
                    .or_else(|| char_style.and_then(|cs| cs.east_asia_font.clone()))
                    .or_else(|| style_east_asia_font.map(|s| s.to_string()))
            },
            bold: rpr
                .and_then(|n| wml_bool(n, "b"))
                .or_else(|| char_style.and_then(|cs| cs.bold))
                .unwrap_or(style_bold),
            italic: rpr
                .and_then(|n| wml_bool(n, "i"))
                .or_else(|| char_style.and_then(|cs| cs.italic))
                .unwrap_or(style_italic),
            underline: rpr
                .and_then(|n| {
                    wml(n, "u")
                        .and_then(|u| u.attribute((WML_NS, "val")))
                        .map(|v| v != "none")
                })
                .or_else(|| char_style.and_then(|cs| cs.underline))
                .unwrap_or(style_underline),
            strikethrough: rpr
                .and_then(|n| wml_bool(n, "strike"))
                .or_else(|| char_style.and_then(|cs| cs.strikethrough))
                .unwrap_or(style_strikethrough),
            dstrike: rpr
                .and_then(|n| wml_bool(n, "dstrike"))
                .unwrap_or(style_dstrike),
            char_spacing: rpr
                .and_then(|n| wml(n, "spacing"))
                .and_then(|n| n.attribute((WML_NS, "val")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(style_char_spacing),
            text_scale: rpr
                .and_then(|n| wml_attr(n, "w"))
                .and_then(|v| v.trim_end_matches('%').parse::<f32>().ok())
                .unwrap_or(100.0),
            caps: rpr
                .and_then(|n| wml_bool(n, "caps"))
                .or_else(|| char_style.and_then(|cs| cs.caps))
                .unwrap_or(style_caps),
            small_caps: rpr
                .and_then(|n| wml_bool(n, "smallCaps"))
                .or_else(|| char_style.and_then(|cs| cs.small_caps))
                .unwrap_or(style_small_caps),
            vanish: rpr
                .and_then(|n| wml_bool(n, "vanish"))
                .or_else(|| char_style.and_then(|cs| cs.vanish))
                .unwrap_or(style_vanish),
            color: rpr
                .and_then(|n| wml_attr(n, "color"))
                .and_then(parse_text_color)
                .or_else(|| char_style.and_then(|cs| cs.color))
                .or(style_color),
            vertical_align: rpr
                .and_then(|n| wml_attr(n, "vertAlign"))
                .map(|v| match v {
                    "superscript" => VertAlign::Superscript,
                    "subscript" => VertAlign::Subscript,
                    _ => VertAlign::Baseline,
                })
                .unwrap_or(VertAlign::Baseline),
            highlight: rpr
                .and_then(|n| wml_attr(n, "highlight"))
                .and_then(highlight_color),
            kern_threshold: rpr
                .and_then(|n| wml_attr(n, "kern"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|hp| hp / 2.0)
                .or_else(|| char_style.and_then(|cs| cs.kern_threshold))
                .or(style_kern_threshold),
            char_style_id: char_style_id_str.map(|s| s.to_string()),
        };

        let flush_pending = |pending: &mut String, runs: &mut Vec<Run>| {
            if !pending.is_empty() {
                let run = fmt.text_run(std::mem::take(pending), hyperlink_url.clone());
                runs.extend(split_run_by_script(run));
            }
        };

        let mut pending_text = String::new();
        for child in run_node.children() {
            let child_ns = child.tag_name().namespace();
            if child_ns == Some(MC_NS) && child.tag_name().name() == "AlternateContent" {
                let choice = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS) && n.tag_name().name() == "Choice"
                });
                let fallback = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS) && n.tag_name().name() == "Fallback"
                });
                if let Some(branch) = choice {
                    for drawing in branch.children().filter(|n| {
                        n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "drawing"
                    }) {
                        match parse_run_drawing(drawing, rels, zip, styles, theme, numbering) {
                            Some(RunDrawingResult::Inline(img)) => {
                                runs.push(Run {
                                    inline_image: Some(img),
                                    ..fmt.minimal_run()
                                });
                            }
                            Some(RunDrawingResult::Floating(fi)) => {
                                floating_images.push(fi);
                            }
                            Some(RunDrawingResult::TextBox(tb)) => {
                                textboxes.push(tb);
                            }
                            Some(RunDrawingResult::Chart(ic)) => {
                                inline_chart = Some(ic);
                            }
                            Some(RunDrawingResult::SmartArt(diagram)) => {
                                smartart = Some(diagram);
                            }
                            None => {}
                        }
                    }
                } else if let Some(branch) = fallback {
                    for pict in branch.descendants().filter(|n| {
                        n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "pict"
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
                "fldChar" => match child.attribute((WML_NS, "fldCharType")) {
                    Some("begin") => {
                        flush_pending(&mut pending_text, &mut runs);
                        in_field = true;
                        in_field_result = false;
                        field_instr.clear();
                        field_result_text.clear();
                    }
                    Some("separate") => {
                        in_field_result = true;
                    }
                    Some("end") => {
                        if in_field {
                            let keyword = field_instr.split_whitespace().next().unwrap_or("");
                            let fc = if keyword.eq_ignore_ascii_case("PAGE") {
                                Some(FieldCode::Page)
                            } else if keyword.eq_ignore_ascii_case("NUMPAGES") {
                                Some(FieldCode::NumPages)
                            } else if keyword.eq_ignore_ascii_case("STYLEREF") {
                                parse_styleref_arg(&field_instr).map(FieldCode::StyleRef)
                            } else {
                                None
                            };
                            if let Some(code) = fc {
                                let text = std::mem::take(&mut field_result_text);
                                runs.push(Run {
                                    text,
                                    font_size: fmt.font_size,
                                    font_name: fmt.font_name.clone(),
                                    bold: fmt.bold,
                                    italic: fmt.italic,
                                    color: fmt.color,
                                    field_code: Some(code),
                                    hyperlink_url: hyperlink_url.clone(),
                                    ..Run::default()
                                });
                            }
                            in_field = false;
                            in_field_result = false;
                            field_instr.clear();
                        }
                    }
                    _ => {}
                },
                "instrText" if in_field && !in_field_result => {
                    if let Some(t) = child.text() {
                        field_instr.push_str(t);
                    }
                }
                "t" if !in_field || (in_field_result && !is_dynamic_field(&field_instr)) => {
                    if let Some(t) = child.text() {
                        let normalized = t.replace('\n', " ");
                        pending_text.push_str(&normalized);
                    }
                }
                "t" if in_field_result && is_dynamic_field(&field_instr) => {
                    if let Some(t) = child.text() {
                        field_result_text.push_str(t);
                    }
                }
                "tab" if !in_field => {
                    flush_pending(&mut pending_text, &mut runs);
                    runs.push(Run {
                        is_tab: true,
                        ..fmt.minimal_run()
                    });
                }
                "br" if !in_field => match child.attribute((WML_NS, "type")) {
                    Some("page") => has_page_break = true,
                    Some("column") => has_column_break = true,
                    _ => line_break_count += 1,
                },
                "drawing" if in_field => {}
                "drawing" => {
                    flush_pending(&mut pending_text, &mut runs);
                    match parse_run_drawing(child, rels, zip, styles, theme, numbering) {
                        Some(RunDrawingResult::Inline(img)) => {
                            runs.push(Run {
                                inline_image: Some(img),
                                ..fmt.minimal_run()
                            });
                        }
                        Some(RunDrawingResult::Floating(fi)) => {
                            floating_images.push(fi);
                        }
                        Some(RunDrawingResult::TextBox(tb)) => {
                            textboxes.push(tb);
                        }
                        Some(RunDrawingResult::Chart(ic)) => {
                            inline_chart = Some(ic);
                        }
                        Some(RunDrawingResult::SmartArt(diagram)) => {
                            smartart = Some(diagram);
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
                    flush_pending(&mut pending_text, &mut runs);
                    if let Some(id) = child
                        .attribute((WML_NS, "id"))
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        runs.push(Run {
                            font_size: fmt.font_size,
                            font_name: fmt.font_name.clone(),
                            bold: fmt.bold,
                            italic: fmt.italic,
                            color: fmt.color,
                            vertical_align: VertAlign::Superscript,
                            footnote_id: Some(id),
                            ..Run::default()
                        });
                    }
                }
                "footnoteRef" if !in_field => {
                    flush_pending(&mut pending_text, &mut runs);
                    runs.push(Run {
                        font_size: fmt.font_size,
                        font_name: fmt.font_name.clone(),
                        bold: fmt.bold,
                        italic: fmt.italic,
                        color: fmt.color,
                        vertical_align: VertAlign::Superscript,
                        is_footnote_ref_mark: true,
                        ..Run::default()
                    });
                }
                _ => {}
            }
        }
        if !pending_text.is_empty() {
            let run = fmt.text_run(pending_text, hyperlink_url.clone());
            runs.extend(split_run_by_script(run));
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
        let mark_font_size = mark_rpr
            .and_then(|n| wml_attr(n, "sz"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hp| hp / 2.0);
        if let Some(mark_font_size) = mark_font_size {
            let mark_font_name = mark_rpr
                .and_then(|n| wml(n, "rFonts"))
                .map(|rfonts| resolve_font_from_node(rfonts, theme, &style_font_name))
                .unwrap_or_else(|| style_font_name.clone());
            runs.push(Run {
                font_size: mark_font_size,
                font_name: mark_font_name,
                bold: style_bold,
                italic: style_italic,
                ..Run::default()
            });
        }
    }

    // Word's paragraph mark uses the paragraph style's font even in empty
    // paragraphs; ensure we carry that font info so line height is correct.
    if runs.is_empty() {
        runs.push(Run {
            font_size: style_font_size,
            font_name: style_font_name.clone(),
            bold: style_bold,
            italic: style_italic,
            ..Run::default()
        });
    }

    ParsedRuns {
        runs,
        has_page_break,
        has_column_break,
        line_break_count,
        floating_images,
        textboxes,
        inline_chart,
        smartart,
    }
}
