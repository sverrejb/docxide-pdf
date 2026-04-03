use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{
    ConnectorShape, FieldCode, FloatingImage, HorizontalRule, InlineChart, Run, SmartArtDiagram,
    TextFill, TextGlow, TextOutline, TextShadow, Textbox, VertAlign,
};

use super::images::{RunDrawingResult, parse_run_drawing};
use super::is_east_asian_char;
use super::numbering::NumberingInfo;
use super::styles::{
    CharacterStyle, ParagraphStyle, StyleDefaults, StylesInfo, ThemeFonts,
    resolve_east_asia_font_from_node, resolve_font_from_node,
};
use super::textbox::parse_textbox_from_vml;
use super::wordart::{parse_text_fill, parse_text_glow, parse_text_outline, parse_text_shadow};
use super::{
    MC_NS_TOP, REL_NS, VML_NS, WML_NS, highlight_color, parse_hex_color, parse_text_color,
    twips_to_pts, wml, wml_attr, wml_bool,
};

fn is_dynamic_field(instr: &str) -> bool {
    let keyword = instr.split_whitespace().next().unwrap_or("");
    keyword.eq_ignore_ascii_case("PAGE")
        || keyword.eq_ignore_ascii_case("NUMPAGES")
        || keyword.eq_ignore_ascii_case("STYLEREF")
        || keyword.eq_ignore_ascii_case("PAGEREF")
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

fn mc_choice_or_fallback<'a>(node: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    let mut fallback: Option<roxmltree::Node<'a, 'a>> = None;
    for n in node.children() {
        if n.tag_name().namespace() != Some(MC_NS_TOP) {
            continue;
        }
        match n.tag_name().name() {
            "Choice" => return Some(n),
            "Fallback" if fallback.is_none() => fallback = Some(n),
            _ => {}
        }
    }
    fallback
}

pub(super) struct ParsedRuns {
    pub(super) runs: Vec<Run>,
    pub(super) has_page_break_before: bool,
    pub(super) has_page_break_after: bool,
    pub(super) has_column_break: bool,
    pub(super) floating_images: Vec<FloatingImage>,
    pub(super) textboxes: Vec<Textbox>,
    pub(super) connectors: Vec<ConnectorShape>,
    pub(super) inline_chart: Option<InlineChart>,
    pub(super) smartart: Option<SmartArtDiagram>,
    pub(super) horizontal_rule: Option<HorizontalRule>,
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
    text_outline: Option<TextOutline>,
    text_fill: Option<TextFill>,
    text_shadow: Option<TextShadow>,
    text_glow: Option<TextGlow>,
    lang: Option<String>,
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
            text_outline: self.text_outline.clone(),
            text_fill: self.text_fill.clone(),
            text_shadow: self.text_shadow.clone(),
            text_glow: self.text_glow.clone(),
            lang: self.lang.clone(),
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

    fn styled_run(&self) -> Run {
        Run {
            font_size: self.font_size,
            font_name: self.font_name.clone(),
            bold: self.bold,
            italic: self.italic,
            color: self.color,
            highlight: self.highlight,
            ..Run::default()
        }
    }

    fn superscript_run(&self) -> Run {
        Run {
            vertical_align: VertAlign::Superscript,
            ..self.styled_run()
        }
    }
}

/// Paragraph-level formatting defaults resolved from the paragraph style chain
/// and document defaults. Used as fallbacks when run-level properties are absent.
struct ParagraphRunDefaults {
    font_size: f32,
    font_name: String,
    bold: bool,
    italic: bool,
    caps: bool,
    small_caps: bool,
    vanish: bool,
    underline: bool,
    strikethrough: bool,
    dstrike: bool,
    color: Option<[u8; 3]>,
    char_spacing: f32,
    kern_threshold: Option<f32>,
    east_asia_font: Option<String>,
}

impl ParagraphRunDefaults {
    fn from_style(para_style: Option<&ParagraphStyle>, defaults: &StyleDefaults) -> Self {
        Self {
            font_size: para_style
                .and_then(|s| s.font_size)
                .unwrap_or(defaults.font_size),
            font_name: para_style
                .and_then(|s| s.font_name.as_deref())
                .unwrap_or(&defaults.font_name)
                .to_string(),
            bold: para_style
                .and_then(|s| s.bold)
                .unwrap_or(defaults.bold),
            italic: para_style
                .and_then(|s| s.italic)
                .unwrap_or(defaults.italic),
            caps: para_style
                .and_then(|s| s.caps)
                .unwrap_or(defaults.caps),
            small_caps: para_style
                .and_then(|s| s.small_caps)
                .unwrap_or(defaults.small_caps),
            vanish: para_style
                .and_then(|s| s.vanish)
                .unwrap_or(defaults.vanish),
            underline: para_style
                .and_then(|s| s.underline)
                .unwrap_or(defaults.underline),
            strikethrough: para_style
                .and_then(|s| s.strikethrough)
                .unwrap_or(defaults.strikethrough),
            dstrike: para_style
                .and_then(|s| s.dstrike)
                .unwrap_or(defaults.dstrike),
            color: para_style.and_then(|s| s.color).or(defaults.color),
            char_spacing: para_style
                .and_then(|s| s.char_spacing)
                .unwrap_or(defaults.char_spacing),
            kern_threshold: para_style
                .and_then(|s| s.kern_threshold)
                .or(defaults.kern_threshold),
            east_asia_font: para_style
                .and_then(|s| s.east_asia_font.clone())
                .or_else(|| defaults.east_asia_font.clone()),
        }
    }

    fn resolve_run_format(
        &self,
        rpr: Option<roxmltree::Node>,
        char_style: Option<&CharacterStyle>,
        char_style_id_str: Option<&str>,
        theme: &ThemeFonts,
    ) -> RunFormat {
        let rfonts_node = rpr.and_then(|n| wml(n, "rFonts"));
        RunFormat {
            font_size: rpr
                .and_then(|n| wml_attr(n, "sz"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|hp| hp / 2.0)
                .or_else(|| char_style.and_then(|cs| cs.font_size))
                .unwrap_or(self.font_size),
            font_name: rfonts_node
                .map(|rfonts| resolve_font_from_node(rfonts, theme, &self.font_name))
                .or_else(|| char_style.and_then(|cs| cs.font_name.clone()))
                .unwrap_or_else(|| self.font_name.clone()),
            east_asia_font_name: rfonts_node
                .and_then(|rfonts| resolve_east_asia_font_from_node(rfonts, theme))
                .or_else(|| char_style.and_then(|cs| cs.east_asia_font.clone()))
                .or_else(|| self.east_asia_font.clone()),
            bold: rpr
                .and_then(|n| wml_bool(n, "b"))
                .or_else(|| char_style.and_then(|cs| cs.bold))
                .unwrap_or(self.bold),
            italic: rpr
                .and_then(|n| wml_bool(n, "i"))
                .or_else(|| char_style.and_then(|cs| cs.italic))
                .unwrap_or(self.italic),
            underline: rpr
                .and_then(|n| {
                    wml(n, "u")
                        .and_then(|u| u.attribute((WML_NS, "val")))
                        .map(|v| v != "none")
                })
                .or_else(|| char_style.and_then(|cs| cs.underline))
                .unwrap_or(self.underline),
            strikethrough: rpr
                .and_then(|n| wml_bool(n, "strike"))
                .or_else(|| char_style.and_then(|cs| cs.strikethrough))
                .unwrap_or(self.strikethrough),
            dstrike: rpr
                .and_then(|n| wml_bool(n, "dstrike"))
                .unwrap_or(self.dstrike),
            char_spacing: rpr
                .and_then(|n| wml(n, "spacing"))
                .and_then(|n| n.attribute((WML_NS, "val")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(self.char_spacing),
            text_scale: rpr
                .and_then(|n| wml_attr(n, "w"))
                .and_then(|v| v.trim_end_matches('%').parse::<f32>().ok())
                .unwrap_or(100.0),
            caps: rpr
                .and_then(|n| wml_bool(n, "caps"))
                .or_else(|| char_style.and_then(|cs| cs.caps))
                .unwrap_or(self.caps),
            small_caps: rpr
                .and_then(|n| wml_bool(n, "smallCaps"))
                .or_else(|| char_style.and_then(|cs| cs.small_caps))
                .unwrap_or(self.small_caps),
            vanish: rpr
                .and_then(|n| wml_bool(n, "vanish"))
                .or_else(|| char_style.and_then(|cs| cs.vanish))
                .unwrap_or(self.vanish),
            color: rpr
                .and_then(|n| wml_attr(n, "color"))
                .and_then(parse_text_color)
                .or_else(|| char_style.and_then(|cs| cs.color))
                .or(self.color),
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
                .and_then(highlight_color)
                .or_else(|| {
                    rpr.and_then(|n| wml(n, "shd"))
                        .and_then(|shd| shd.attribute((WML_NS, "fill")))
                        .filter(|f| *f != "none" && *f != "auto")
                        .and_then(parse_hex_color)
                })
                .or_else(|| char_style.and_then(|cs| cs.highlight)),
            kern_threshold: rpr
                .and_then(|n| wml_attr(n, "kern"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|hp| hp / 2.0)
                .or_else(|| char_style.and_then(|cs| cs.kern_threshold))
                .or(self.kern_threshold),
            char_style_id: char_style_id_str.map(|s| s.to_string()),
            text_outline: rpr.and_then(parse_text_outline),
            text_fill: rpr.and_then(parse_text_fill),
            text_shadow: rpr.and_then(parse_text_shadow),
            text_glow: rpr.and_then(parse_text_glow),
            lang: rpr
                .and_then(|n| wml(n, "lang"))
                .and_then(|n| n.attribute((WML_NS, "val")))
                .map(|s| s.to_string()),
        }
    }
}

/// Create synthetic runs for empty paragraphs so the renderer computes the
/// correct line height from the paragraph mark's formatting.
fn ensure_nonempty_paragraph(
    runs: &mut Vec<Run>,
    ppr: Option<roxmltree::Node>,
    defaults: &ParagraphRunDefaults,
    theme: &ThemeFonts,
    has_page_break_before: bool,
) {
    if !runs.is_empty() || has_page_break_before {
        return;
    }
    let mark_rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
    let mark_font_size = mark_rpr
        .and_then(|n| wml_attr(n, "sz"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|hp| hp / 2.0);
    if let Some(mark_font_size) = mark_font_size {
        let mark_font_name = mark_rpr
            .and_then(|n| wml(n, "rFonts"))
            .map(|rfonts| resolve_font_from_node(rfonts, theme, &defaults.font_name))
            .unwrap_or_else(|| defaults.font_name.clone());
        runs.push(Run {
            font_size: mark_font_size,
            font_name: mark_font_name,
            bold: defaults.bold,
            italic: defaults.italic,
            ..Run::default()
        });
    }
    if runs.is_empty() {
        runs.push(Run {
            font_size: defaults.font_size,
            font_name: defaults.font_name.clone(),
            bold: defaults.bold,
            italic: defaults.italic,
            ..Run::default()
        });
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

fn collect_run_nodes<'a>(
    parent: roxmltree::Node<'a, 'a>,
    rels: &HashMap<String, String>,
    out: &mut Vec<(roxmltree::Node<'a, 'a>, Option<String>, bool)>,
) {
    for child in parent.children() {
        let name = child.tag_name().name();
        let ns = child.tag_name().namespace();
        let is_wml = ns == Some(WML_NS);
        if is_wml && name == "r" {
            out.push((child, None, false));
        } else if is_wml && name == "hyperlink" {
            let has_rid = child.attribute((REL_NS, "id")).is_some();
            let has_anchor = child.attribute((WML_NS, "anchor")).is_some();
            let is_anchor_only = has_anchor && !has_rid;
            let url = if is_anchor_only {
                child.attribute((WML_NS, "anchor")).map(|a| format!("#{a}"))
            } else {
                child
                    .attribute((REL_NS, "id"))
                    .and_then(|rid| rels.get(rid))
                    .cloned()
            };
            for n in child
                .children()
                .filter(|n| n.tag_name().name() == "r" && n.tag_name().namespace() == Some(WML_NS))
            {
                out.push((n, url.clone(), is_anchor_only));
            }
        } else if is_wml && matches!(name, "ins" | "smartTag") {
            collect_run_nodes(child, rels, out);
        } else if is_wml && name == "del" {
            // Final mode: skip deleted content entirely
        } else if is_wml && name == "sdt" {
            if let Some(content) = wml(child, "sdtContent") {
                collect_run_nodes(content, rels, out);
            }
        } else if ns == Some(MC_NS_TOP) && name == "AlternateContent" {
            if let Some(branch) = mc_choice_or_fallback(child) {
                collect_run_nodes(branch, rels, out);
            }
        }
    }
}

macro_rules! handle_drawing_result {
    ($result:expr, $fmt:expr, $runs:expr, $floating_images:expr, $textboxes:expr,
     $inline_chart:expr, $smartart:expr, $connectors:expr) => {
        match $result {
            Some(RunDrawingResult::Inline(img)) => {
                $runs.push(Run {
                    inline_image: Some(img),
                    ..$fmt.minimal_run()
                });
            }
            Some(RunDrawingResult::Floating(fi)) => $floating_images.push(fi),
            Some(RunDrawingResult::TextBox(tb)) => $textboxes.push(tb),
            Some(RunDrawingResult::Chart(ic)) => $inline_chart = Some(ic),
            Some(RunDrawingResult::SmartArt(diagram)) => $smartart = Some(diagram),
            Some(RunDrawingResult::Connector(c)) => $connectors.push(c),
            None => {}
        }
    };
}

/// Merge consecutive text runs with identical visual properties.
/// Word often splits a single word across multiple `w:r` elements for revision
/// tracking (different rsidR). Without merging, the layout engine treats each
/// run's text independently, allowing line breaks mid-word.
fn merge_compatible_runs(runs: Vec<Run>) -> Vec<Run> {
    if runs.len() <= 1 {
        return runs;
    }
    let mut result: Vec<Run> = Vec::with_capacity(runs.len());
    for run in runs {
        let can_merge = result.last().is_some_and(|prev| {
            !prev.is_tab
                && !run.is_tab
                && !prev.is_line_break
                && !run.is_line_break
                && prev.inline_image.is_none()
                && run.inline_image.is_none()
                && prev.footnote_id.is_none()
                && run.footnote_id.is_none()
                && !prev.is_footnote_ref_mark
                && !run.is_footnote_ref_mark
                && prev.field_code.is_none()
                && run.field_code.is_none()
                && prev.font_name == run.font_name
                && prev.east_asia_font_name == run.east_asia_font_name
                && prev.font_size == run.font_size
                && prev.bold == run.bold
                && prev.italic == run.italic
                && prev.underline == run.underline
                && prev.strikethrough == run.strikethrough
                && prev.dstrike == run.dstrike
                && prev.char_spacing == run.char_spacing
                && prev.text_scale == run.text_scale
                && prev.caps == run.caps
                && prev.small_caps == run.small_caps
                && prev.vanish == run.vanish
                && prev.color == run.color
                && prev.highlight == run.highlight
                && prev.vertical_align == run.vertical_align
                && prev.kern_threshold == run.kern_threshold
                && prev.hyperlink_url == run.hyperlink_url
                && prev.text_outline == run.text_outline
                && prev.text_fill == run.text_fill
                && prev.text_shadow == run.text_shadow
                && prev.text_glow == run.text_glow
                && prev.lang == run.lang
        });
        if can_merge {
            result.last_mut().unwrap().text.push_str(&run.text);
        } else {
            result.push(run);
        }
    }
    result
}

pub(super) fn parse_runs<R: Read + Seek>(
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
    let defaults = ParagraphRunDefaults::from_style(para_style, &styles.defaults);

    let mut run_nodes: Vec<(roxmltree::Node, Option<String>, bool)> = Vec::new();
    collect_run_nodes(para_node, rels, &mut run_nodes);

    let mut runs = Vec::new();
    let mut floating_images: Vec<FloatingImage> = Vec::new();
    let mut textboxes: Vec<Textbox> = Vec::new();
    let mut connectors: Vec<ConnectorShape> = Vec::new();
    let mut inline_chart: Option<InlineChart> = None;
    let mut smartart: Option<SmartArtDiagram> = None;
    let mut horizontal_rule: Option<HorizontalRule> = None;
    let mut has_page_break_after = false;
    let mut page_break_before_content = false;
    let mut has_column_break = false;
    let mut in_field = false;
    let mut in_field_result = false;
    let mut field_instr = String::new();
    let mut field_result_text = String::new();

    for (run_node, hyperlink_url, is_anchor_hyperlink) in run_nodes {
        let rpr = wml(run_node, "rPr");

        let char_style_id_str = rpr.and_then(|n| wml_attr(n, "rStyle"));
        let char_style = if is_anchor_hyperlink {
            None
        } else {
            char_style_id_str.and_then(|id| styles.character_styles.get(id))
        };

        let fmt = defaults.resolve_run_format(rpr, char_style, char_style_id_str, theme);

        let flush_pending = |pending: &mut String, runs: &mut Vec<Run>| {
            if !pending.is_empty() {
                let run = fmt.text_run(std::mem::take(pending), hyperlink_url.clone());
                runs.extend(split_run_by_script(run));
            }
        };

        let mut pending_text = String::new();
        for child in run_node.children() {
            let child_ns = child.tag_name().namespace();
            if child_ns == Some(MC_NS_TOP) && child.tag_name().name() == "AlternateContent" {
                let choice = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS_TOP) && n.tag_name().name() == "Choice"
                });
                if let Some(branch) = choice {
                    for drawing in branch.children().filter(|n| {
                        n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "drawing"
                    }) {
                        let result =
                            parse_run_drawing(drawing, rels, zip, styles, theme, numbering);
                        handle_drawing_result!(
                            result,
                            fmt,
                            runs,
                            floating_images,
                            textboxes,
                            inline_chart,
                            smartart,
                            connectors
                        );
                    }
                } else if let Some(branch) = child.children().find(|n| {
                    n.tag_name().namespace() == Some(MC_NS_TOP) && n.tag_name().name() == "Fallback"
                }) {
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
                            } else if keyword.eq_ignore_ascii_case("PAGEREF") {
                                field_instr.split_whitespace().nth(1)
                                    .map(|s| FieldCode::PageRef(s.to_string()))
                            } else {
                                None
                            };
                            if let Some(code) = fc {
                                let text = std::mem::take(&mut field_result_text);
                                runs.push(Run {
                                    text,
                                    field_code: Some(code),
                                    hyperlink_url: hyperlink_url.clone(),
                                    ..fmt.styled_run()
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
                        pending_text.push_str(&t.replace('\n', " "));
                    }
                }
                "t" if in_field_result && is_dynamic_field(&field_instr) => {
                    if let Some(t) = child.text() {
                        field_result_text.push_str(t);
                    }
                }
                "noBreakHyphen" => {
                    pending_text.push('-');
                }
                "tab" if !in_field
                    || (in_field_result && !is_dynamic_field(&field_instr)) =>
                {
                    flush_pending(&mut pending_text, &mut runs);
                    runs.push(Run {
                        is_tab: true,
                        ..fmt.minimal_run()
                    });
                }
                "br" if !in_field => match child.attribute((WML_NS, "type")) {
                    Some("page") => {
                        if runs.is_empty() && pending_text.is_empty() {
                            page_break_before_content = true;
                        } else {
                            has_page_break_after = true;
                        }
                    }
                    Some("column") => has_column_break = true,
                    _ => {
                        flush_pending(&mut pending_text, &mut runs);
                        runs.push(Run {
                            is_line_break: true,
                            ..fmt.minimal_run()
                        });
                    }
                },
                "drawing" if in_field => {}
                "drawing" => {
                    flush_pending(&mut pending_text, &mut runs);
                    let result = parse_run_drawing(child, rels, zip, styles, theme, numbering);
                    handle_drawing_result!(
                        result,
                        fmt,
                        runs,
                        floating_images,
                        textboxes,
                        inline_chart,
                        smartart,
                        connectors
                    );
                }
                "pict" if !in_field => {
                    if let Some(hr) = parse_vml_horizontal_rule(child) {
                        horizontal_rule = Some(hr);
                    } else if let Some(tb) =
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
                            footnote_id: Some(id),
                            ..fmt.superscript_run()
                        });
                    }
                }
                "footnoteRef" if !in_field => {
                    flush_pending(&mut pending_text, &mut runs);
                    runs.push(Run {
                        is_footnote_ref_mark: true,
                        ..fmt.superscript_run()
                    });
                }
                "sym" if !in_field => {
                    flush_pending(&mut pending_text, &mut runs);
                    let sym_font = child.attribute((WML_NS, "font")).unwrap_or(&fmt.font_name);
                    if let Some(ch) = child
                        .attribute((WML_NS, "char"))
                        .and_then(|hex| u32::from_str_radix(hex, 16).ok())
                        .and_then(char::from_u32)
                    {
                        runs.push(Run {
                            text: ch.to_string(),
                            font_name: sym_font.to_string(),
                            font_size: fmt.font_size,
                            bold: fmt.bold,
                            italic: fmt.italic,
                            color: fmt.color,
                            underline: fmt.underline,
                            strikethrough: fmt.strikethrough,
                            char_spacing: fmt.char_spacing,
                            ..Run::default()
                        });
                    }
                }
                _ => {}
            }
        }
        if !pending_text.is_empty() {
            let run = fmt.text_run(pending_text, hyperlink_url.clone());
            runs.extend(split_run_by_script(run));
        }
    }

    let has_page_break_before = ppr
        .and_then(|ppr| wml_bool(ppr, "pageBreakBefore"))
        .unwrap_or(false)
        || page_break_before_content;

    ensure_nonempty_paragraph(&mut runs, ppr, &defaults, theme, has_page_break_before);

    let runs = merge_compatible_runs(runs);

    ParsedRuns {
        runs,
        has_page_break_before,
        has_page_break_after,
        has_column_break,
        floating_images,
        textboxes,
        connectors,
        inline_chart,
        smartart,
        horizontal_rule,
    }
}

const OFFICE_NS: &str = "urn:schemas-microsoft-com:office:office";

fn parse_vml_horizontal_rule(pict_node: roxmltree::Node) -> Option<HorizontalRule> {
    let shape = pict_node.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS)
            && matches!(n.tag_name().name(), "rect" | "shape")
    })?;

    let is_hr = shape
        .attribute((OFFICE_NS, "hr"))
        .is_some_and(|v| v == "t" || v == "true");
    if !is_hr {
        return None;
    }

    let style_str = shape.attribute("style").unwrap_or("");
    let mut height_pt = 1.5_f32;
    for part in style_str.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            if key.trim() == "height" {
                height_pt = val.trim().trim_end_matches("pt").parse().unwrap_or(1.5);
            }
        }
    }

    let fill_color = shape
        .attribute("fillcolor")
        .and_then(|c| {
            let hex = c.strip_prefix('#').unwrap_or(c);
            parse_hex_color(hex)
        })
        .unwrap_or([0xa0, 0xa0, 0xa0]);

    let hrpct = shape
        .attribute((OFFICE_NS, "hrpct"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 10.0);
    let width_pct = hrpct.unwrap_or(100.0);

    let is_standard = shape
        .attribute((OFFICE_NS, "hrstd"))
        .is_some_and(|v| v == "t" || v == "true");

    Some(HorizontalRule {
        height_pt,
        fill_color,
        width_pct,
        is_standard,
    })
}
