use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{
    ConnectorShape, FieldCode, FloatingImage, FormCheckbox, HorizontalRule, IfPart, InlineChart,
    Run, SmartArtDiagram, TabAlignment, TextFill, TextGlow, TextOutline, TextShadow, Textbox,
    VertAlign,
};

use super::images::{
    RunDrawingResult, is_vml_picture, parse_object_floating_image, parse_object_inline_image,
    parse_run_drawing,
};
use super::styles::{
    ParagraphStyle, RunProps, StyleDefaults, ThemeFonts, half_points, parse_font_size,
    parse_run_props, resolve_font_from_node_opt,
};
use super::textbox::parse_textbox_from_vml;
use super::{
    MATH_NS, MC_NS_TOP, OFFICE_NS, ParseContext, REL_NS, VML_NS, WML_NS, math_child, math_run_text,
    math_val, parse_hex_color, parse_pt, wml, wml_attr, wml_bool,
};
use super::{is_complex_script_char, is_east_asian_char};

/// A field the renderer re-evaluates per page, from its instruction.
fn parse_field_code(instr: &str) -> Option<FieldCode> {
    let keyword = instr.split_whitespace().next()?;
    if keyword.eq_ignore_ascii_case("PAGE") {
        Some(FieldCode::Page)
    } else if keyword.eq_ignore_ascii_case("NUMPAGES") {
        Some(FieldCode::NumPages)
    } else if keyword.eq_ignore_ascii_case("STYLEREF") {
        parse_styleref_arg(instr).map(|name| FieldCode::StyleRef {
            name,
            number: has_switch(instr, "\\n"),
        })
    } else if keyword.eq_ignore_ascii_case("PAGEREF") {
        instr
            .split_whitespace()
            .nth(1)
            .map(|s| FieldCode::PageRef(s.to_string()))
    } else {
        None
    }
}

fn has_switch(instr: &str, switch: &str) -> bool {
    instr
        .split_whitespace()
        .any(|s| s.eq_ignore_ascii_case(switch))
}

/// One open complex field while parsing runs. Word fields nest: a field that
/// begins inside another field's *instruction* region (before its `separate`)
/// is an argument to the parent (e.g. a `PAGE` field used inside an `IF`
/// expression) and must never display on its own — Word substitutes the
/// parent's evaluated result. A field that begins inside a parent's *result*
/// region (after `separate`, e.g. a `PAGEREF` inside a `TOC` result) is visible
/// content and is processed normally. `visible` records which case applies.
struct FieldFrame {
    seen_sep: bool,
    visible: bool,
    instr: String,
    result: String,
    /// The instruction as text and nested fields, for `FieldCode::If`.
    parts: Vec<IfPart>,
    /// A nested field we cannot evaluate: an `IF` keeps its cached result.
    opaque: bool,
    checkbox: Option<FormCheckbox>,
}

impl FieldFrame {
    fn new(visible: bool) -> Self {
        Self {
            seen_sep: false,
            visible,
            instr: String::new(),
            result: String::new(),
            parts: Vec::new(),
            opaque: false,
            checkbox: None,
        }
    }

    /// A field nested in this one's instruction.
    fn nest(&mut self, instr: &str) {
        match parse_field_code(instr) {
            None | Some(FieldCode::PageRef(_)) => self.opaque = true,
            Some(code) => self.parts.push(IfPart::Field(code)),
        }
    }

    /// Instruction text; a nested field's cached value only joins `instr`.
    fn push_instr(&mut self, t: &str, nested_value: bool) {
        self.instr.push_str(t);
        if nested_value {
            return;
        }
        match self.parts.last_mut() {
            Some(IfPart::Text(s)) => s.push_str(t),
            _ => self.parts.push(IfPart::Text(t.to_string())),
        }
    }

    fn evaluable_if(&self) -> bool {
        !self.opaque
            && self
                .instr
                .split_whitespace()
                .next()
                .is_some_and(|k| k.eq_ignore_ascii_case("IF"))
            && self.parts.iter().any(|p| matches!(p, IfPart::Field(_)))
    }

    /// The cached result is replaced at render time.
    fn dynamic(&self) -> bool {
        self.evaluable_if() || parse_field_code(&self.instr).is_some()
    }
}

/// A legacy check box form field (`w:ffData/w:checkBox` on its begin
/// `fldChar`); `w:checked` overrides `w:default`.
fn parse_checkbox(fld_char: roxmltree::Node, run_size: f32) -> Option<FormCheckbox> {
    let cb = wml(wml(fld_char, "ffData")?, "checkBox")?;
    let size = half_points(cb, "size").unwrap_or(run_size);
    let checked = wml_bool(cb, "checked")
        .or_else(|| wml_bool(cb, "default"))
        .unwrap_or(false);
    Some(FormCheckbox { size, checked })
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
    /// True only when the break came from `<w:br w:type="page"/>` at the
    /// start of the paragraph, not from the `<w:pageBreakBefore/>` style
    /// property. The renderer treats these differently — see
    /// `Paragraph.page_break_before_explicit`.
    pub(super) has_explicit_page_break_before: bool,
    pub(super) has_page_break_after: bool,
    /// Index into `runs` of the first run after a `<w:br w:type="page"/>`
    /// that has visible content after it in the same paragraph.
    /// Mid-paragraph break index (`page_break_at`) is a column break.
    pub(super) column_break_at: bool,
    pub(super) page_break_at: Option<usize>,
    pub(super) has_column_break: bool,
    pub(super) has_clear_break: bool,
    pub(super) floating_images: Vec<FloatingImage>,
    pub(super) textboxes: Vec<Textbox>,
    pub(super) connectors: Vec<ConnectorShape>,
    pub(super) inline_chart: Option<InlineChart>,
    pub(super) smartart: Vec<SmartArtDiagram>,
    pub(super) horizontal_rule: Option<HorizontalRule>,
}

/// A `Run` doubles as the resolved formatting of one `w:r`: every run the
/// element yields is built from that template with struct update syntax.
impl Run {
    fn text_run(&self, text: String, hyperlink_url: Option<String>) -> Run {
        Run {
            text,
            hyperlink_url,
            ..self.clone()
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

    /// Build a tab run that retains underline + color so a `<w:tab/>` with
    /// `<w:u w:val="single"/>` renders an underlined span across the tab gap —
    /// Word uses this to draw horizontal rules in signature lines etc.
    fn tab_run(&self) -> Run {
        Run {
            is_tab: true,
            underline: self.underline,
            double_underline: self.double_underline,
            color: self.color,
            ..self.minimal_run()
        }
    }

    /// Build a positional-tab run (`w:ptab`). It is a tab (splits segments) that
    /// carries its own alignment, resolved against the margin box in layout.
    fn ptab_run(&self, alignment: TabAlignment) -> Run {
        Run {
            is_tab: true,
            ptab_alignment: Some(alignment),
            color: self.color,
            ..self.minimal_run()
        }
    }

    fn styled_run(&self) -> Run {
        Run {
            bold: self.bold,
            italic: self.italic,
            color: self.color,
            highlight: self.highlight,
            shading: self.shading,
            ..self.minimal_run()
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
    /// True when font_size came from doc defaults, not from the paragraph style
    font_size_is_doc_default: bool,
    /// True when font_name came from doc defaults, not from the paragraph style
    font_name_is_doc_default: bool,
    bold: bool,
    italic: bool,
    caps: bool,
    small_caps: bool,
    vanish: bool,
    underline: bool,
    double_underline: bool,
    strikethrough: bool,
    dstrike: bool,
    color: Option<[u8; 3]>,
    char_spacing: f32,
    kern_threshold: Option<f32>,
    position: Option<f32>,
    east_asia_font: Option<String>,
    cs_font: Option<String>,
    text_outline: Option<TextOutline>,
    text_fill: Option<TextFill>,
    text_shadow: Option<TextShadow>,
    text_glow: Option<TextGlow>,
    lang: Option<String>,
    lang_east_asia: Option<String>,
}

impl ParagraphRunDefaults {
    fn from_style(para_style: Option<&ParagraphStyle>, defaults: &StyleDefaults) -> Self {
        let para_font_size = para_style.and_then(|s| s.font_size);
        let para_font_name = para_style.and_then(|s| s.font_name.as_deref());
        let style_or = |f: fn(&ParagraphStyle) -> Option<bool>, default: bool| {
            para_style.and_then(f).unwrap_or(default)
        };
        let style_or_clone = |f: fn(&ParagraphStyle) -> Option<&String>,
                              default: &Option<String>| {
            para_style.and_then(f).or(default.as_ref()).cloned()
        };
        Self {
            font_size: para_font_size.unwrap_or(defaults.font_size),
            font_name: para_font_name.unwrap_or(&defaults.font_name).to_string(),
            font_size_is_doc_default: para_font_size.is_none(),
            font_name_is_doc_default: para_font_name.is_none(),
            bold: style_or(|s| s.bold, defaults.bold),
            italic: style_or(|s| s.italic, defaults.italic),
            caps: style_or(|s| s.caps, defaults.caps),
            small_caps: style_or(|s| s.small_caps, defaults.small_caps),
            vanish: style_or(|s| s.vanish, defaults.vanish),
            underline: style_or(|s| s.underline, defaults.underline),
            double_underline: style_or(|s| s.double_underline, defaults.double_underline),
            strikethrough: style_or(|s| s.strikethrough, defaults.strikethrough),
            dstrike: style_or(|s| s.dstrike, defaults.dstrike),
            color: para_style.and_then(|s| s.color).or(defaults.color),
            char_spacing: para_style
                .and_then(|s| s.char_spacing)
                .unwrap_or(defaults.char_spacing),
            lang: style_or_clone(|s| s.lang.as_ref(), &defaults.lang),
            lang_east_asia: style_or_clone(|s| s.lang_east_asia.as_ref(), &defaults.lang_east_asia),
            kern_threshold: para_style
                .and_then(|s| s.kern_threshold)
                .or(defaults.kern_threshold),
            position: para_style.and_then(|s| s.position).or(defaults.position),
            east_asia_font: style_or_clone(|s| s.east_asia_font.as_ref(), &defaults.east_asia_font),
            cs_font: style_or_clone(|s| s.cs_font.as_ref(), &defaults.cs_font),
            text_outline: para_style.and_then(|s| s.text_outline.clone()),
            text_fill: para_style.and_then(|s| s.text_fill.clone()),
            text_shadow: para_style.and_then(|s| s.text_shadow.clone()),
            text_glow: para_style.and_then(|s| s.text_glow.clone()),
        }
    }

    /// The formatting template for one `w:r`: its own `rpr`, then the
    /// character style, then this paragraph's defaults.
    fn resolve_run_format(
        &self,
        rpr: Option<roxmltree::Node>,
        char_style: Option<&RunProps>,
        char_style_id_str: Option<&str>,
        theme: &ThemeFonts,
    ) -> Run {
        let own = rpr.map(|n| parse_run_props(n, theme)).unwrap_or_default();
        let cs = |f: fn(&RunProps) -> Option<bool>| char_style.and_then(f);
        let char_style_font_size = char_style.and_then(|cs| cs.font_size);
        let char_style_font_name = char_style.and_then(|cs| cs.font_name.clone());
        // True only when font_size/name came from doc defaults — not from inline rPr,
        // character style, OR paragraph style.  Table style overrides apply only here.
        let font_size_from_default = own.font_size.is_none()
            && char_style_font_size.is_none()
            && self.font_size_is_doc_default;
        let font_name_from_default = own.font_name.is_none()
            && char_style_font_name.is_none()
            && self.font_name_is_doc_default;

        // Legacy Word-97 run text-effect toggles (§17.3.2.23/.31/.13/.18). These
        // are independent of the modern w14 DrawingML effects parsed below: a
        // plain <w:outline/>/<w:shadow/>/<w:emboss/>/<w:imprint/> draws as hollow
        // / dropshadowed / raised / engraved text. `w:effect` (animated shimmer)
        // has no print form, so we deliberately don't read it — base text only.
        let font_size = own
            .font_size
            .or(char_style_font_size)
            .unwrap_or(self.font_size);
        let color = own
            .color
            .or_else(|| char_style.and_then(|cs| cs.color))
            .or(self.color);
        let legacy_outline = rpr.and_then(|n| wml_bool(n, "outline")).unwrap_or(false);
        let legacy_shadow = rpr.and_then(|n| wml_bool(n, "shadow")).unwrap_or(false);
        let legacy_emboss = rpr.and_then(|n| wml_bool(n, "emboss")).unwrap_or(false);
        let legacy_imprint = rpr.and_then(|n| wml_bool(n, "imprint")).unwrap_or(false);

        Run {
            font_size,
            font_name: own
                .font_name
                .or(char_style_font_name)
                .unwrap_or_else(|| self.font_name.clone()),
            east_asia_font_name: own
                .east_asia_font
                .or_else(|| char_style.and_then(|cs| cs.east_asia_font.clone()))
                .or_else(|| self.east_asia_font.clone()),
            cs_font_name: own
                .cs_font
                .or_else(|| char_style.and_then(|cs| cs.cs_font.clone()))
                .or_else(|| self.cs_font.clone()),
            bold: own.bold.or_else(|| cs(|c| c.bold)).unwrap_or(self.bold),
            italic: own
                .italic
                .or_else(|| cs(|c| c.italic))
                .unwrap_or(self.italic),
            // Decide underline from the w:val attribute, not the mere presence
            // of <w:u>. Word treats a bare <w:u> with no val (e.g.
            // <w:u w:color="000000"/>) as "no underline applied" and inherits,
            // so keying off presence wrongly underlines such runs.
            underline: own
                .underline
                .or_else(|| cs(|c| c.underline))
                .unwrap_or(self.underline),
            double_underline: own
                .double_underline
                .or_else(|| cs(|c| c.double_underline))
                .unwrap_or(self.double_underline),
            strikethrough: own
                .strikethrough
                .or_else(|| cs(|c| c.strikethrough))
                .unwrap_or(self.strikethrough),
            dstrike: own.dstrike.unwrap_or(self.dstrike),
            char_spacing: own.char_spacing.unwrap_or(self.char_spacing),
            text_scale: rpr
                .and_then(|n| wml_attr(n, "w"))
                .and_then(|v| v.trim_end_matches('%').parse::<f32>().ok())
                .unwrap_or(100.0),
            caps: own.caps.or_else(|| cs(|c| c.caps)).unwrap_or(self.caps),
            small_caps: own
                .small_caps
                .or_else(|| cs(|c| c.small_caps))
                .unwrap_or(self.small_caps),
            vanish: own
                .vanish
                .or_else(|| cs(|c| c.vanish))
                .unwrap_or(self.vanish),
            color,
            vertical_align: rpr
                .and_then(|n| wml_attr(n, "vertAlign"))
                .map(|v| match v {
                    "superscript" => VertAlign::Superscript,
                    "subscript" => VertAlign::Subscript,
                    _ => VertAlign::Baseline,
                })
                .unwrap_or(VertAlign::Baseline),
            highlight: own
                .highlight
                .or_else(|| char_style.and_then(|cs| cs.highlight)),
            shading: own.shading.or_else(|| char_style.and_then(|cs| cs.shading)),
            border: own
                .border
                .or_else(|| char_style.and_then(|cs| cs.border.clone())),
            kern_threshold: own
                .kern_threshold
                .or_else(|| char_style.and_then(|cs| cs.kern_threshold))
                .or(self.kern_threshold),
            position: own
                .position
                .or_else(|| char_style.and_then(|cs| cs.position))
                .or(self.position)
                .unwrap_or(0.0),
            char_style_id: char_style_id_str.map(|s| s.to_string()),
            text_outline: own
                .text_outline
                .or_else(|| {
                    legacy_outline.then(|| TextOutline {
                        // Thin hairline: Word's outline antialiases to light gray,
                        // which only happens with a sub-pixel stroke width.
                        width_pt: (font_size * 0.014).max(0.3),
                        color: color.unwrap_or([0, 0, 0]),
                    })
                })
                .or_else(|| char_style.and_then(|cs| cs.text_outline.clone()))
                .or_else(|| self.text_outline.clone()),
            text_fill: own
                .text_fill
                // <w:outline/> hollows the glyphs: stroke only, no fill.
                .or_else(|| legacy_outline.then_some(TextFill::NoFill))
                .or_else(|| char_style.and_then(|cs| cs.text_fill.clone()))
                .or_else(|| self.text_fill.clone()),
            text_shadow: own
                .text_shadow
                .or_else(|| {
                    legacy_text_shadow(legacy_shadow, legacy_emboss, legacy_imprint, font_size)
                })
                .or_else(|| char_style.and_then(|cs| cs.text_shadow.clone()))
                .or_else(|| self.text_shadow.clone()),
            text_glow: own
                .text_glow
                .or_else(|| char_style.and_then(|cs| cs.text_glow.clone()))
                .or_else(|| self.text_glow.clone()),
            lang: rpr.and_then(|n| wml_attr(n, "lang")).map(|s| s.to_string()),
            text_lang: own
                .lang
                .or_else(|| char_style.and_then(|cs| cs.lang.clone()))
                .or_else(|| self.lang.clone())
                .map(Into::into),
            text_lang_east_asia: own
                .lang_east_asia
                .or_else(|| char_style.and_then(|cs| cs.lang_east_asia.clone()))
                .or_else(|| self.lang_east_asia.clone())
                .map(Into::into),
            font_size_from_default,
            font_name_from_default,
            bold_is_direct: own.bold.is_some(),
            italic_is_direct: own.italic.is_some(),
            ..Run::default()
        }
    }
}

/// Map the legacy Word-97 run toggles (mutually exclusive per spec) to a
/// drop-shadow approximation. Word's PDF export renders all three as a
/// down-right gray drop shadow behind the (still-colored) glyph face;
/// emboss/imprint use a slightly smaller offset than plain shadow.
fn legacy_text_shadow(
    shadow: bool,
    emboss: bool,
    imprint: bool,
    font_size: f32,
) -> Option<TextShadow> {
    // Drop shadow casts down-right. Word renders both emboss (raised) and
    // imprint (engraved) on a white page as a prominent down-right gray drop
    // shadow behind the glyph face, with the white background showing between
    // as the raised/engraved ridge. Empirically the relief reads about as heavy
    // as the plain-shadow line in Word's PDF export, so use the same offset
    // (not a reduced one).
    // ponytail: a single offset gray copy — Word's exact multi-pass face/
    // highlight/shadow antialiasing isn't replicated (a rare legacy effect).
    let d = (font_size * 0.035).max(0.4);
    (shadow || emboss || imprint).then_some(TextShadow {
        color: [128, 128, 128],
        offset_x: d,
        offset_y: -d,
        alpha: 1.0,
    })
}

/// Create synthetic runs for empty paragraphs so the renderer computes the
/// correct line height from the paragraph mark's formatting.
fn ensure_nonempty_paragraph(
    runs: &mut Vec<Run>,
    ppr: Option<roxmltree::Node>,
    defaults: &ParagraphRunDefaults,
    theme: &ThemeFonts,
    has_page_break_before: bool,
    character_styles: &HashMap<String, RunProps>,
) {
    if !runs.is_empty() || has_page_break_before {
        return;
    }
    // The mark's own font and size each override the style's on their own
    // (a Calibri mark with no w:sz in a Times New Roman style).
    // Like a real run, an unset size or font inherits and so takes a table
    // style's (empty 10pt Table Grid cell paragraphs).
    let mark_rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
    let mark_style = mark_rpr
        .and_then(|rpr| wml_attr(rpr, "rStyle"))
        .and_then(|id| character_styles.get(id));
    let mark_size = mark_rpr
        .and_then(parse_font_size)
        .or_else(|| mark_style.and_then(|s| s.font_size));
    let mark_font = mark_rpr
        .and_then(|n| wml(n, "rFonts"))
        .and_then(|rfonts| resolve_font_from_node_opt(rfonts, theme))
        .or_else(|| mark_style.and_then(|s| s.font_name.clone()));
    runs.push(Run {
        font_size: mark_size.unwrap_or(defaults.font_size),
        font_size_from_default: mark_size.is_none() && defaults.font_size_is_doc_default,
        font_name_from_default: mark_font.is_none() && defaults.font_name_is_doc_default,
        font_name: mark_font.unwrap_or_else(|| defaults.font_name.clone()),
        // The mark's w:b picks the bold face's metrics: nabl's 39 empty Arial
        // Narrow Bold marks are 0.14pt taller each than regular ones.
        bold: mark_rpr
            .and_then(|r| wml_bool(r, "b"))
            .unwrap_or(defaults.bold),
        italic: mark_rpr
            .and_then(|r| wml_bool(r, "i"))
            .unwrap_or(defaults.italic),
        ..Run::default()
    });
}

fn split_run_by_script(run: Run) -> Vec<Run> {
    let ea_font = run
        .east_asia_font_name
        .clone()
        .filter(|f| f != &run.font_name);
    let has_cs = run.text.chars().any(is_complex_script_char);
    if run.text.is_empty() || (ea_font.is_none() && !has_cs) {
        return vec![run];
    }
    // Word draws complex-script letters in the cs font, in Arial when the
    // document names none (arabic_rice_benefits_article has no styles part).
    let cs_font = run.cs_font_name.clone().unwrap_or_else(|| "Arial".into());
    // None: the run's own font.
    let script_font = |ch: char| {
        if is_complex_script_char(ch) {
            Some(&cs_font)
        } else if is_east_asian_char(ch) {
            ea_font.as_ref()
        } else {
            None
        }
    };
    let segment = |range: std::ops::Range<usize>, font: Option<&String>| {
        let mut sub = run.clone();
        sub.text = run.text[range].to_string();
        if let Some(f) = font {
            sub.font_name = f.clone();
        }
        sub.east_asia_font_name = None;
        sub.cs_font_name = None;
        sub
    };

    let mut result = Vec::new();
    let mut segment_start = 0;
    let mut current = None;
    for (i, ch) in run.text.char_indices() {
        // Whitespace inherits the current script context.
        let font = if ch.is_whitespace() && i > 0 {
            current
        } else {
            script_font(ch)
        };
        if i > 0 && font != current {
            result.push(segment(segment_start..i, current));
            segment_start = i;
        }
        current = font;
    }
    result.push(segment(segment_start..run.text.len(), current));
    result
}

fn is_comment_reference_run(node: roxmltree::Node) -> bool {
    node.children()
        .any(|c| c.has_tag_name((WML_NS, "commentReference")))
}

fn collect_run_nodes<'a>(
    parent: roxmltree::Node<'a, 'a>,
    rels: &HashMap<String, String>,
    out: &mut Vec<(roxmltree::Node<'a, 'a>, Option<String>, bool, Vec<u32>)>,
    active_comments: &mut Vec<u32>,
) {
    for child in parent.children() {
        let name = child.tag_name().name();
        let ns = child.tag_name().namespace();
        let is_wml = ns == Some(WML_NS);
        if is_wml && name == "commentRangeStart" {
            if let Some(id) = child
                .attribute((WML_NS, "id"))
                .and_then(|v| v.parse::<u32>().ok())
            {
                active_comments.push(id);
            }
            continue;
        }
        if is_wml && name == "commentRangeEnd" {
            if let Some(id) = child
                .attribute((WML_NS, "id"))
                .and_then(|v| v.parse::<u32>().ok())
                && let Some(pos) = active_comments.iter().rposition(|x| *x == id)
            {
                active_comments.remove(pos);
            }
            continue;
        }
        if is_wml && name == "r" {
            if is_comment_reference_run(child) {
                continue;
            }
            out.push((child, None, false, active_comments.clone()));
        } else if is_wml && name == "br" {
            // A w:br straight under w:p is malformed, but Word breaks there
            // (a citation's URL then starts the next line).
            out.push((child, None, false, active_comments.clone()));
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
            // Everything inside the link takes it, nested links and wrappers
            // included: croatian_grant's portal URL sits in a hyperlink
            // inside a hyperlink, and the inner one wins.
            let first = out.len();
            collect_run_nodes(child, rels, out, active_comments);
            for (_, run_url, anchor_only, _) in &mut out[first..] {
                if run_url.is_none() {
                    *run_url = url.clone();
                    *anchor_only = is_anchor_only;
                }
            }
        } else if is_wml && matches!(name, "ins" | "moveTo" | "smartTag" | "customXml") {
            // w:customXml inline-wraps runs transparently, like w:smartTag.
            // Final mode shows moved text at its destination (moveTo).
            collect_run_nodes(child, rels, out, active_comments);
        } else if is_wml && matches!(name, "del" | "moveFrom") {
            // Final mode: skip deleted content and the source of moved text
        } else if is_wml && name == "sdt" {
            if let Some(content) = wml(child, "sdtContent") {
                collect_run_nodes(content, rels, out, active_comments);
            }
        } else if is_wml && name == "fldSimple" {
            // Simple field: <w:fldSimple w:instr="..."><w:r>cached</w:r></w:fldSimple>.
            // Emit the cached display runs. Dynamic field codes (PAGE/NUMPAGES/etc.)
            // are not rewritten here — the cached text matches Word's online converter
            // output for the static-PDF case.
            collect_run_nodes(child, rels, out, active_comments);
        } else if ns == Some(MC_NS_TOP) && name == "AlternateContent" {
            if let Some(branch) = mc_choice_or_fallback(child) {
                collect_run_nodes(branch, rels, out, active_comments);
            }
        } else if ns == Some(MATH_NS) && name == "oMathPara" {
            // A math paragraph wraps one or more m:oMath; emit each in order.
            for om in child
                .children()
                .filter(|n| n.has_tag_name((MATH_NS, "oMath")))
            {
                out.push((om, None, false, active_comments.clone()));
            }
        } else if ns == Some(MATH_NS) && name == "oMath" {
            out.push((child, None, false, active_comments.clone()));
        }
    }
}

pub(super) fn push_floating(fis: &mut Vec<FloatingImage>, tbs: &[Textbox], mut fi: FloatingImage) {
    fi.anchor_seq = (fis.len() + tbs.len()) as u32;
    fis.push(fi);
}

pub(super) fn push_textbox(fis: &[FloatingImage], tbs: &mut Vec<Textbox>, mut tb: Textbox) {
    tb.anchor_seq = (fis.len() + tbs.len()) as u32;
    tbs.push(tb);
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
            Some(RunDrawingResult::Floating(fi)) => {
                push_floating(&mut $floating_images, &$textboxes, fi)
            }
            Some(RunDrawingResult::TextBox(tb)) => {
                push_textbox(&$floating_images, &mut $textboxes, tb)
            }
            Some(RunDrawingResult::Chart(ic)) => $inline_chart = Some(ic),
            Some(RunDrawingResult::SmartArt(diagram)) => $smartart.push(diagram),
            Some(RunDrawingResult::Connector(c)) => $connectors.push(c),
            Some(RunDrawingResult::Group(items)) => {
                // Flattened groups contain only leaf shapes, never nested Groups
                for item in items {
                    match item {
                        RunDrawingResult::Inline(img) => {
                            $runs.push(Run {
                                inline_image: Some(img),
                                ..$fmt.minimal_run()
                            });
                        }
                        RunDrawingResult::Floating(fi) => {
                            push_floating(&mut $floating_images, &$textboxes, fi)
                        }
                        RunDrawingResult::TextBox(tb) => {
                            push_textbox(&$floating_images, &mut $textboxes, tb)
                        }
                        RunDrawingResult::Chart(ic) => $inline_chart = Some(ic),
                        RunDrawingResult::SmartArt(diagram) => $smartart.push(diagram),
                        RunDrawingResult::Connector(c) => $connectors.push(c),
                        RunDrawingResult::Group(_) => {}
                    }
                }
            }
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
                && prev.endnote_id.is_none()
                && run.endnote_id.is_none()
                && !prev.is_endnote_ref_mark
                && !run.is_endnote_ref_mark
                && prev.field_code.is_none()
                && run.field_code.is_none()
                && prev.checkbox.is_none()
                && run.checkbox.is_none()
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
                && prev.shading == run.shading
                && prev.border == run.border
                && prev.vertical_align == run.vertical_align
                && prev.kern_threshold == run.kern_threshold
                && prev.position == run.position
                && prev.hyperlink_url == run.hyperlink_url
                && prev.text_outline == run.text_outline
                && prev.text_fill == run.text_fill
                && prev.text_shadow == run.text_shadow
                && prev.text_glow == run.text_glow
                && prev.lang == run.lang
                && prev.char_style_id == run.char_style_id
                && prev.comment_ids == run.comment_ids
                // One Formula per math zone, never merged into the text around it.
                && match (&prev.formula, &run.formula) {
                    (None, None) => true,
                    (Some(a), Some(b)) => std::sync::Arc::ptr_eq(a, b),
                    _ => false,
                }
        });
        if can_merge {
            result.last_mut().unwrap().text.push_str(&run.text);
        } else {
            result.push(run);
        }
    }
    result
}

/// Default alignment for a paragraph that contains a display-math block
/// (`m:oMathPara`). OOXML centers display math by default (`m:oMathParaPr/m:jc`,
/// default `centerGroup`), which is why e.g. table-header equation cells appear
/// centered even with no paragraph `w:jc`. Returns `None` for paragraphs with no
/// `m:oMathPara` (inline `m:oMath` flows with the surrounding text and is not
/// treated as display math).
pub(super) fn display_math_alignment(para: roxmltree::Node) -> Option<crate::model::Alignment> {
    use crate::model::Alignment;
    let omp = para
        .children()
        .find(|n| n.has_tag_name((MATH_NS, "oMathPara")))?;
    let jc = math_val(omp, "oMathParaPr", "jc");
    Some(match jc {
        Some("left") => Alignment::Left,
        Some("right") => Alignment::Right,
        // center / centerGroup / absent
        _ => Alignment::Center,
    })
}

/// Linearize an Office Math (OMML `m:oMath`) tree into ordinary text runs,
/// reusing the normal run pipeline (fonts, sub/superscript). Stacked constructs
/// are approximated inline — fractions as `num/den`, radicals as `√(radicand)`,
/// delimiters as `(…)` — rather than laid out two-dimensionally. This is not
/// full equation layout, but it renders the symbols and sub/superscripts that
/// were previously dropped entirely (e.g. table headers `T₁₀±∆T₁₀` and inline
/// formulas), since the run parser never descended into the math namespace.
fn omath_to_runs(
    node: roxmltree::Node,
    defaults: &ParagraphRunDefaults,
    theme: &ThemeFonts,
    vert: VertAlign,
    out: &mut Vec<Run>,
) {
    let push_lit = |out: &mut Vec<Run>, s: &str| {
        let fmt = defaults.resolve_run_format(None, None, None, theme);
        let mut r = fmt.text_run(s.to_string(), None);
        r.vertical_align = vert;
        r.is_math = true;
        out.push(r);
    };
    for child in node.children() {
        if child.tag_name().namespace() != Some(MATH_NS) {
            continue;
        }
        match child.tag_name().name() {
            "r" => {
                let text = math_run_text(child);
                if !text.is_empty() {
                    // Math runs carry a w:rPr (fonts/size); m:rPr math styling is ignored.
                    let fmt = defaults.resolve_run_format(wml(child, "rPr"), None, None, theme);
                    let mut r = fmt.text_run(text, None);
                    // In math, sub/superscript is conveyed by the structure
                    // (m:sSub/m:sSup), not the run's w:vertAlign. Word ignores a
                    // run-level vertAlign here, so always apply the structural
                    // position — otherwise a cell whose runs all carry a spurious
                    // w:vertAlign="subscript" renders entirely shrunken.
                    r.vertical_align = vert;
                    r.is_math = true;
                    out.extend(split_run_by_script(r).into_iter().map(|mut sr| {
                        sr.is_math = true;
                        sr
                    }));
                }
            }
            "sSub" => {
                if let Some(e) = math_child(child, "e") {
                    omath_to_runs(e, defaults, theme, vert, out);
                }
                if let Some(s) = math_child(child, "sub") {
                    omath_to_runs(s, defaults, theme, VertAlign::Subscript, out);
                }
            }
            "sSup" => {
                if let Some(e) = math_child(child, "e") {
                    omath_to_runs(e, defaults, theme, vert, out);
                }
                if let Some(s) = math_child(child, "sup") {
                    omath_to_runs(s, defaults, theme, VertAlign::Superscript, out);
                }
            }
            "sSubSup" => {
                if let Some(e) = math_child(child, "e") {
                    omath_to_runs(e, defaults, theme, vert, out);
                }
                if let Some(s) = math_child(child, "sub") {
                    omath_to_runs(s, defaults, theme, VertAlign::Subscript, out);
                }
                if let Some(s) = math_child(child, "sup") {
                    omath_to_runs(s, defaults, theme, VertAlign::Superscript, out);
                }
            }
            "f" => {
                if let Some(n) = math_child(child, "num") {
                    omath_to_runs(n, defaults, theme, vert, out);
                }
                push_lit(out, "/");
                if let Some(d) = math_child(child, "den") {
                    omath_to_runs(d, defaults, theme, vert, out);
                }
            }
            "rad" => {
                push_lit(out, "√");
                if let Some(e) = math_child(child, "e") {
                    push_lit(out, "(");
                    omath_to_runs(e, defaults, theme, vert, out);
                    push_lit(out, ")");
                }
            }
            "d" => {
                push_lit(out, "(");
                if let Some(e) = math_child(child, "e") {
                    omath_to_runs(e, defaults, theme, vert, out);
                }
                push_lit(out, ")");
            }
            "nary" => {
                let chr = math_val(child, "naryPr", "chr").unwrap_or("∫");
                push_lit(out, chr);
                if let Some(s) = math_child(child, "sub") {
                    omath_to_runs(s, defaults, theme, VertAlign::Subscript, out);
                }
                if let Some(s) = math_child(child, "sup") {
                    omath_to_runs(s, defaults, theme, VertAlign::Superscript, out);
                }
                if let Some(e) = math_child(child, "e") {
                    omath_to_runs(e, defaults, theme, vert, out);
                }
            }
            // Structural containers reached directly (e.g. m:e, m:func, m:limLow,
            // m:groupChr) — recurse to collect their text in document order.
            _ => omath_to_runs(child, defaults, theme, vert, out),
        }
    }
}

pub(super) fn parse_runs<R: Read + Seek>(
    para_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> ParsedRuns {
    let ppr = wml(para_node, "pPr");
    let para_style_id = ppr
        .and_then(|ppr| wml_attr(ppr, "pStyle"))
        .unwrap_or(&ctx.styles.default_paragraph_style_id);
    let para_style = ctx.styles.paragraph_styles.get(para_style_id);
    let defaults = ParagraphRunDefaults::from_style(para_style, &ctx.styles.defaults);

    let mut run_nodes: Vec<(roxmltree::Node, Option<String>, bool, Vec<u32>)> = Vec::new();
    let mut active_comments: Vec<u32> = Vec::new();
    collect_run_nodes(para_node, ctx.rels, &mut run_nodes, &mut active_comments);

    let mut runs = Vec::new();
    let mut floating_images: Vec<FloatingImage> = Vec::new();
    let mut textboxes: Vec<Textbox> = Vec::new();
    let mut connectors: Vec<ConnectorShape> = Vec::new();
    let mut inline_chart: Option<InlineChart> = None;
    let mut smartart: Vec<SmartArtDiagram> = Vec::new();
    let mut horizontal_rule: Option<HorizontalRule> = None;
    let mut has_page_break_after = false;
    let mut page_break_before_content = false;
    let mut page_break_at: Option<usize> = None;
    let mut column_break_at = false;
    let mut has_column_break = false;
    let mut has_clear_break = false;
    let mut field_stack: Vec<FieldFrame> = Vec::new();

    for (run_node, hyperlink_url, is_anchor_hyperlink, comment_ids) in run_nodes {
        if run_node.has_tag_name((MATH_NS, "oMath")) {
            let first = runs.len();
            omath_to_runs(
                run_node,
                &defaults,
                ctx.theme,
                VertAlign::Baseline,
                &mut runs,
            );
            let spoken = super::math_speech::speak(run_node);
            let formula = (!spoken.is_empty()).then(|| std::sync::Arc::<str>::from(spoken));
            for r in &mut runs[first..] {
                r.formula = formula.clone();
            }
            continue;
        }
        let runs_before = runs.len();
        // A fldSimple inside an open field's instruction is an argument to it;
        // its runs carry the nested field's cached value as instrText.
        let simple = run_node
            .parent()
            .filter(|p| p.has_tag_name((WML_NS, "fldSimple")));
        let simple_instr = simple.and_then(|fs| fs.attribute((WML_NS, "instr")));
        // The fldSimple's first run stands for the field; the rest are its value.
        let simple_first_run = simple.is_some_and(|fs| {
            fs.children().find(|n| n.has_tag_name((WML_NS, "r"))) == Some(run_node)
        });
        if simple_first_run
            && let Some(f) = field_stack.last_mut()
            && !f.seen_sep
        {
            f.nest(simple_instr.unwrap_or(""));
        }
        let rpr = wml(run_node, "rPr");

        let char_style_id_str = rpr.and_then(|n| wml_attr(n, "rStyle"));
        let char_style = if is_anchor_hyperlink {
            None
        } else {
            char_style_id_str.and_then(|id| ctx.styles.character_styles.get(id))
        };

        let fmt = defaults.resolve_run_format(rpr, char_style, char_style_id_str, ctx.theme);

        // A STYLEREF fldSimple on its own is one field run, re-evaluated per
        // page like the complex form.
        if field_stack.is_empty()
            && let Some(code @ FieldCode::StyleRef { .. }) = simple_instr.and_then(parse_field_code)
        {
            if simple_first_run {
                let cached: String = simple
                    .into_iter()
                    .flat_map(|fs| fs.descendants())
                    .filter(|n| n.has_tag_name((WML_NS, "t")))
                    .filter_map(|n| n.text())
                    .collect();
                runs.push(Run {
                    text: cached,
                    field_code: Some(code),
                    hyperlink_url: hyperlink_url.clone(),
                    ..fmt.styled_run()
                });
            }
            continue;
        }

        let flush_pending = |pending: &mut String, runs: &mut Vec<Run>| {
            if !pending.is_empty() {
                let run = fmt.text_run(std::mem::take(pending), hyperlink_url.clone());
                runs.extend(split_run_by_script(run));
            }
        };

        let mut pending_text = String::new();
        let bare_break = (run_node.tag_name().name() == "br").then_some(run_node);
        for child in bare_break.into_iter().chain(run_node.children()) {
            let child_ns = child.tag_name().namespace();
            if child.has_tag_name((MC_NS_TOP, "AlternateContent")) {
                let choice = child
                    .children()
                    .find(|n| n.has_tag_name((MC_NS_TOP, "Choice")));
                if let Some(branch) = choice {
                    for drawing in branch
                        .children()
                        .filter(|n| n.has_tag_name((WML_NS, "drawing")))
                    {
                        let result = parse_run_drawing(drawing, ctx);
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
                } else if let Some(branch) = child
                    .children()
                    .find(|n| n.has_tag_name((MC_NS_TOP, "Fallback")))
                {
                    for pict in branch
                        .descendants()
                        .filter(|n| n.has_tag_name((WML_NS, "pict")))
                    {
                        if let Some(tb) = parse_textbox_from_vml(pict, ctx) {
                            push_textbox(&floating_images, &mut textboxes, tb);
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
                        // A field is visible content unless the innermost open
                        // field is still in its instruction region — in which
                        // case this is a field argument and never displays.
                        let parent_visible = field_stack.last().is_none_or(|f| f.seen_sep);
                        if field_stack.is_empty() {
                            flush_pending(&mut pending_text, &mut runs);
                        }
                        field_stack.push(FieldFrame {
                            checkbox: parse_checkbox(child, fmt.font_size),
                            ..FieldFrame::new(parent_visible)
                        });
                    }
                    Some("separate") => {
                        if let Some(f) = field_stack.last_mut() {
                            f.seen_sep = true;
                        }
                    }
                    Some("end") => {
                        let Some(mut f) = field_stack.pop() else {
                            continue;
                        };
                        if !f.visible {
                            if let Some(parent) = field_stack.last_mut()
                                && !parent.seen_sep
                            {
                                parent.nest(&f.instr);
                            }
                            continue;
                        }
                        if let Some(cb) = f.checkbox {
                            runs.push(Run {
                                text: FormCheckbox::TEXT.to_string(),
                                checkbox: Some(cb),
                                hyperlink_url: hyperlink_url.clone(),
                                ..fmt.styled_run()
                            });
                            continue;
                        }
                        let fc = if f.evaluable_if() {
                            Some(FieldCode::If(std::mem::take(&mut f.parts)))
                        } else {
                            parse_field_code(&f.instr)
                        };
                        if let Some(code) = fc {
                            // PAGEREF \h is a hyperlink to its bookmark (TOC page
                            // numbers): Word tags it as the TOCI's Link.
                            let url = match &code {
                                FieldCode::PageRef(bookmark)
                                    if hyperlink_url.is_none() && has_switch(&f.instr, "\\h") =>
                                {
                                    Some(format!("#{bookmark}"))
                                }
                                _ => hyperlink_url.clone(),
                            };
                            runs.push(Run {
                                text: f.result,
                                field_code: Some(code),
                                hyperlink_url: url,
                                ..fmt.styled_run()
                            });
                        }
                    }
                    _ => {}
                },
                "instrText" => {
                    // Instruction text belongs to the innermost open field that
                    // has not yet reached its separator.
                    if let Some(f) = field_stack.last_mut()
                        && !f.seen_sep
                        && let Some(t) = child.text()
                    {
                        f.push_instr(t, simple.is_some());
                    }
                }
                "t" => {
                    let visible = field_stack.last().is_none_or(|f| f.seen_sep);
                    let dyn_result = field_stack
                        .last()
                        .is_some_and(|f| f.seen_sep && f.dynamic());
                    if dyn_result {
                        // Cached result of a dynamic field — keep it as the
                        // field run's placeholder text (re-evaluated at render).
                        if let Some(t) = child.text() {
                            field_stack.last_mut().unwrap().result.push_str(t);
                        }
                    } else if visible && let Some(t) = child.text() {
                        pending_text.push_str(&t.replace('\n', " "));
                    }
                }
                "noBreakHyphen" => {
                    pending_text.push('-');
                }
                "tab" => {
                    let visible = field_stack.last().is_none_or(|f| f.seen_sep);
                    let dyn_result = field_stack
                        .last()
                        .is_some_and(|f| f.seen_sep && f.dynamic());
                    if visible && !dyn_result {
                        flush_pending(&mut pending_text, &mut runs);
                        runs.push(fmt.tab_run());
                    }
                }
                "ptab" => {
                    // Positional tab: alignment positions the FOLLOWING text against the
                    // margin box (left/center/right), independent of paragraph tab stops.
                    let visible = field_stack.last().is_none_or(|f| f.seen_sep);
                    let dyn_result = field_stack
                        .last()
                        .is_some_and(|f| f.seen_sep && f.dynamic());
                    if visible && !dyn_result {
                        let alignment = match child.attribute((WML_NS, "alignment")) {
                            Some("center") => TabAlignment::Center,
                            Some("right") => TabAlignment::Right,
                            _ => TabAlignment::Left,
                        };
                        flush_pending(&mut pending_text, &mut runs);
                        runs.push(fmt.ptab_run(alignment));
                    }
                }
                "br" if field_stack.is_empty() => match child.attribute((WML_NS, "type")) {
                    Some("page") => {
                        // Floats anchored before the break belong to the page it
                        // ends (flyer templates put every text box in front of it).
                        let floats_before = !floating_images.is_empty()
                            || !textboxes.is_empty()
                            || !connectors.is_empty()
                            || !smartart.is_empty();
                        if runs.is_empty() && pending_text.is_empty() && !floats_before {
                            page_break_before_content = true;
                        } else {
                            flush_pending(&mut pending_text, &mut runs);
                            page_break_at.get_or_insert(runs.len());
                            has_page_break_after = true;
                        }
                    }
                    // Text after a column break moves on, like after a page break
                    Some("column") => {
                        if runs.is_empty() && pending_text.is_empty() {
                            has_column_break = true;
                        } else {
                            flush_pending(&mut pending_text, &mut runs);
                            if page_break_at.is_none() {
                                page_break_at = Some(runs.len());
                                column_break_at = true;
                            }
                        }
                    }
                    _ => {
                        if child.attribute((WML_NS, "clear")) == Some("all") {
                            has_clear_break = true;
                        }
                        flush_pending(&mut pending_text, &mut runs);
                        runs.push(Run {
                            is_line_break: true,
                            ..fmt.minimal_run()
                        });
                    }
                },
                // A carriage return breaks the line like a text-wrapping br (§17.3.3.4).
                "cr" if field_stack.is_empty() => {
                    flush_pending(&mut pending_text, &mut runs);
                    runs.push(Run {
                        is_line_break: true,
                        ..fmt.minimal_run()
                    });
                }
                "drawing" if !field_stack.is_empty() => {}
                "drawing" => {
                    flush_pending(&mut pending_text, &mut runs);
                    let result = parse_run_drawing(child, ctx);
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
                "pict" if field_stack.is_empty() => {
                    if let Some(hr) = parse_vml_horizontal_rule(child) {
                        horizontal_rule = Some(hr);
                    } else if is_vml_picture(child) {
                        flush_pending(&mut pending_text, &mut runs);
                        if let Some(fi) = parse_object_floating_image(child, ctx) {
                            push_floating(&mut floating_images, &textboxes, fi);
                        } else if let Some(img) = parse_object_inline_image(child, ctx) {
                            runs.push(Run {
                                inline_image: Some(img),
                                ..fmt.minimal_run()
                            });
                        }
                    } else if let Some(tb) = parse_textbox_from_vml(child, ctx) {
                        push_textbox(&floating_images, &mut textboxes, tb);
                    }
                }
                "object" if field_stack.is_empty() => {
                    flush_pending(&mut pending_text, &mut runs);
                    if let Some(fi) = parse_object_floating_image(child, ctx) {
                        push_floating(&mut floating_images, &textboxes, fi);
                    } else if let Some(img) = parse_object_inline_image(child, ctx) {
                        runs.push(Run {
                            inline_image: Some(img),
                            ..fmt.minimal_run()
                        });
                    }
                }
                "footnoteReference" if field_stack.is_empty() => {
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
                "footnoteRef" if field_stack.is_empty() => {
                    flush_pending(&mut pending_text, &mut runs);
                    runs.push(Run {
                        is_footnote_ref_mark: true,
                        ..fmt.superscript_run()
                    });
                }
                "endnoteReference" if field_stack.is_empty() => {
                    flush_pending(&mut pending_text, &mut runs);
                    if let Some(id) = child
                        .attribute((WML_NS, "id"))
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        runs.push(Run {
                            endnote_id: Some(id),
                            ..fmt.superscript_run()
                        });
                    }
                }
                "endnoteRef" if field_stack.is_empty() => {
                    flush_pending(&mut pending_text, &mut runs);
                    runs.push(Run {
                        is_endnote_ref_mark: true,
                        ..fmt.superscript_run()
                    });
                }
                "sym" if field_stack.is_empty() => {
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
        if !comment_ids.is_empty() {
            for r in &mut runs[runs_before..] {
                r.comment_ids = comment_ids.clone();
            }
        }
    }

    let has_page_break_before = ppr
        .and_then(|ppr| wml_bool(ppr, "pageBreakBefore"))
        .unwrap_or(false)
        || page_break_before_content;

    ensure_nonempty_paragraph(
        &mut runs,
        ppr,
        &defaults,
        ctx.theme,
        has_page_break_before,
        &ctx.styles.character_styles,
    );

    // Merge each side of a mid-paragraph page break on its own so the split
    // index stays valid; a break with nothing visible after it stays a plain
    // break after the paragraph.
    let tail = page_break_at.map(|at| runs.split_off(at.min(runs.len())));
    let mut runs = merge_compatible_runs(runs);
    let page_break_at = tail.and_then(|tail| {
        let at = runs.len();
        let has_content = tail
            .iter()
            .any(|r| !r.text.trim().is_empty() || r.is_tab || r.inline_image.is_some());
        runs.extend(merge_compatible_runs(tail));
        // After a column break the paragraph mark still takes a line at the
        // top of the next column (Word probe: "A<br column/>" lays out like A
        // followed by a paragraph holding only the break).
        (has_content || column_break_at).then_some(at)
    });

    ParsedRuns {
        runs,
        has_explicit_page_break_before: page_break_before_content,
        has_page_break_after,
        column_break_at,
        page_break_at,
        has_column_break,
        has_clear_break,
        floating_images,
        textboxes,
        connectors,
        inline_chart,
        smartart,
        horizontal_rule,
    }
}

fn parse_vml_horizontal_rule(pict_node: roxmltree::Node) -> Option<HorizontalRule> {
    let shape = pict_node.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS) && matches!(n.tag_name().name(), "rect" | "shape")
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
        if let Some((key, val)) = part.trim().split_once(':')
            && key.trim() == "height"
        {
            height_pt = parse_pt(val).unwrap_or(1.5);
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

#[cfg(test)]
mod tests {
    use super::*;

    fn run(text: &str, shading: Option<[u8; 3]>, highlight: Option<[u8; 3]>) -> Run {
        Run {
            text: text.to_string(),
            font_name: "Arial".into(),
            font_size: 11.0,
            shading,
            highlight,
            ..Run::default()
        }
    }

    #[test]
    fn checkbox_size_and_state() {
        let parse = |check_box: &str| {
            let xml = format!(
                r#"<w:fldChar xmlns:w="{WML_NS}" w:fldCharType="begin"><w:ffData><w:checkBox>{check_box}</w:checkBox></w:ffData></w:fldChar>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_checkbox(doc.root_element(), 11.0)
        };
        let cb = |size, checked| Some(FormCheckbox { size, checked });
        assert_eq!(
            parse(r#"<w:size w:val="20"/><w:default w:val="0"/>"#),
            cb(10.0, false)
        );
        assert_eq!(
            parse(r#"<w:sizeAuto/><w:default w:val="1"/>"#),
            cb(11.0, true)
        );
        assert_eq!(
            parse(r#"<w:sizeAuto/><w:default w:val="1"/><w:checked w:val="0"/>"#),
            cb(11.0, false)
        );
        assert_eq!(parse(r#"<w:sizeAuto/><w:checked/>"#), cb(11.0, true));
        let text_field = format!(
            r#"<w:fldChar xmlns:w="{WML_NS}" w:fldCharType="begin"><w:ffData><w:textInput/></w:ffData></w:fldChar>"#
        );
        let doc = roxmltree::Document::parse(&text_field).unwrap();
        assert_eq!(parse_checkbox(doc.root_element(), 11.0), None);
    }

    #[test]
    fn complex_script_letters_take_the_cs_font() {
        let split = |cs: Option<&str>| {
            let mut r = run("rice الأرز 150 x", None, None);
            r.font_name = "Aptos".into();
            r.cs_font_name = cs.map(Into::into);
            split_run_by_script(r)
                .into_iter()
                .map(|r| (r.text, r.font_name))
                .collect::<Vec<_>>()
        };
        let seg = |t: &str, f: &str| (t.to_string(), f.to_string());
        // Spaces follow the letters before them; digits keep the run's font.
        let expect = |cs| {
            vec![
                seg("rice ", "Aptos"),
                seg("الأرز ", cs),
                seg("150 x", "Aptos"),
            ]
        };
        assert_eq!(split(None), expect("Arial"));
        assert_eq!(split(Some("Tahoma")), expect("Tahoma"));
    }

    #[test]
    fn merge_compatible_runs_preserves_shading_boundary() {
        // Runs with different shading must NOT merge, or the later run's
        // shading silently disappears into the former's formatting.
        let runs = vec![
            run("before ", None, None),
            run("shaded", Some([255, 244, 163]), None),
            run(" after", None, None),
        ];
        let merged = merge_compatible_runs(runs);
        assert_eq!(merged.len(), 3, "runs differing in shading must not merge");
        assert_eq!(merged[1].shading, Some([255, 244, 163]));
    }

    #[test]
    fn collect_run_nodes_unwraps_fldsimple() {
        // w:fldSimple wraps cached display runs (e.g. for FILLIN fields). The
        // run inside must be reachable so its text reaches the output.
        let ns = WML_NS;
        let xml = format!(
            r#"<w:p xmlns:w="{ns}">
              <w:fldSimple w:instr=" FILLIN &quot;x&quot; \d OFFICIAL ">
                <w:r><w:t>OFFICIAL</w:t></w:r>
              </w:fldSimple>
            </w:p>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let rels = HashMap::new();
        let mut out = Vec::new();
        let mut active = Vec::new();
        collect_run_nodes(doc.root_element(), &rels, &mut out, &mut active);
        assert_eq!(out.len(), 1, "fldSimple's inner w:r must be collected");
        let t_text = out[0]
            .0
            .children()
            .find(|n| n.has_tag_name((WML_NS, "t")))
            .and_then(|n| n.text());
        assert_eq!(t_text, Some("OFFICIAL"));
    }

    #[test]
    fn collect_run_nodes_keeps_text_of_nested_hyperlinks() {
        let ns = WML_NS;
        let xml = format!(
            r#"<w:p xmlns:w="{ns}" xmlns:r="{REL_NS}">
              <w:hyperlink r:id="a"><w:hyperlink r:id="b"><w:r><w:t>inner</w:t></w:r></w:hyperlink></w:hyperlink>
            </w:p>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let rels = HashMap::from([("a".into(), "outer".into()), ("b".into(), "inner".into())]);
        let mut out = Vec::new();
        collect_run_nodes(doc.root_element(), &rels, &mut out, &mut Vec::new());
        assert_eq!(out.len(), 1);
        assert_eq!(out[0].1.as_deref(), Some("inner"));
    }

    #[test]
    fn collect_run_nodes_keeps_moved_text_at_its_destination() {
        let ns = WML_NS;
        let xml = format!(
            r#"<w:p xmlns:w="{ns}">
              <w:moveFrom w:id="1" w:author="a"><w:r><w:t>old</w:t></w:r></w:moveFrom>
              <w:moveTo w:id="2" w:author="a"><w:r><w:t>new</w:t></w:r></w:moveTo>
            </w:p>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let mut out = Vec::new();
        collect_run_nodes(
            doc.root_element(),
            &HashMap::new(),
            &mut out,
            &mut Vec::new(),
        );
        let texts: Vec<_> = out
            .iter()
            .filter_map(|(r, ..)| r.descendants().find(|n| n.has_tag_name((WML_NS, "t"))))
            .filter_map(|t| t.text())
            .collect();
        assert_eq!(texts, ["new"]);
    }

    #[test]
    fn merge_compatible_runs_merges_same_shading() {
        // Adjacent runs with identical shading SHOULD merge (the common case
        // of a single shaded word split across multiple w:r by revision
        // tracking).
        let c = Some([200, 250, 204]);
        let runs = vec![run("green ", c, None), run("continues", c, None)];
        let merged = merge_compatible_runs(runs);
        assert_eq!(merged.len(), 1);
        assert_eq!(merged[0].text, "green continues");
        assert_eq!(merged[0].shading, c);
    }
}
