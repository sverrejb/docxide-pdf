use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

use crate::model::{
    ConnectorShape, FieldCode, FloatingImage, FormCheckbox, HorizontalRule, IfPart, InlineChart,
    Ruby, Run, SmartArtDiagram, TabAlignment, TextFill, TextGlow, TextOutline, TextShadow, Textbox,
    VertAlign,
};

use super::images::{
    RunDrawingResult, is_vml_picture, parse_object_floating_image, parse_object_inline_image,
    parse_run_drawing,
};
use super::styles::{
    ComplexProps, ParagraphStyle, RunProps, StyleDefaults, ThemeFonts, half_points,
    parse_font_size, parse_run_props, resolve_font_from_node_opt,
};
use super::textbox::{parse_connector_from_vml, parse_textbox_from_vml};
use super::{
    MATH_NS, MC_NS_TOP, OFFICE_NS, ParseContext, REL_NS, VML_NS, WML_NS, math_run_text, math_val,
    parse_hex_color, parse_pt, wml, wml_attr, wml_bool,
};
use super::{is_complex_script_char, is_east_asian_char};

fn custom_note_mark(note_ref: roxmltree::Node) -> bool {
    note_ref
        .attribute((WML_NS, "customMarkFollows"))
        .is_some_and(super::parse_on_off)
}

/// A field the renderer re-evaluates per page, from its instruction.
fn parse_field_code(instr: &str) -> Option<FieldCode> {
    let keyword = instr.split_whitespace().next()?;
    if keyword.eq_ignore_ascii_case("PAGE") {
        Some(FieldCode::Page(number_format_switch(instr)))
    } else if keyword.eq_ignore_ascii_case("NUMPAGES") {
        Some(FieldCode::NumPages(number_format_switch(instr)))
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

/// `EQ \o\ac(□,√)` — Word's overstrike equation, how 方案模板 draw a checked
/// checkbox: the arguments are painted on top of each other and the first sets
/// the advance. The field carries **no cached result** (Word computes it when
/// the document is laid out), so without this the whole checkbox disappeared.
///
/// Only the overstrike switch (`\o`) is understood — fractions (`\f`), radicals
/// (`\r`) and the rest keep no result, same as before. Checkbox shapes map to
/// their precomposed characters (`□`/`☐` + `√` → `☑`); any other overstrike
/// keeps the base character, which is the part that stays legible when the
/// accent is dropped.
fn parse_eq_overstrike(instr: &str) -> Option<String> {
    let keyword = instr.split_whitespace().next()?;
    if !keyword.eq_ignore_ascii_case("EQ") {
        return None;
    }
    let lower = instr.to_ascii_lowercase();
    let over = lower.find("\\o")?;
    let open = instr[over..].find('(')? + over;
    let close = instr.rfind(')')?;
    if close <= open {
        return None;
    }
    let parts: Vec<&str> = instr[open + 1..close].split(',').map(str::trim).collect();
    let base = parts.first().copied().unwrap_or("");
    let marks: String = parts.iter().skip(1).copied().collect();
    if base.chars().count() == 1 && matches!(base, "□" | "☐") && marks.contains('√') {
        return Some("☑".to_string());
    }
    (!base.is_empty()).then(|| base.to_string())
}

fn has_switch(instr: &str, switch: &str) -> bool {
    instr
        .split_whitespace()
        .any(|s| s.eq_ignore_ascii_case(switch))
}

/// `\* ROMAN`, `\* alphabetic` …: the numFmt a numeric field is shown in.
/// Like Word, the case of the format word picks upper or lower case.
fn number_format_switch(instr: &str) -> Option<&'static str> {
    let words: Vec<&str> = instr.split_whitespace().collect();
    words.windows(2).filter(|w| w[0] == "\\*").find_map(|w| {
        let upper = w[1].chars().all(|c| c.is_ascii_uppercase());
        match w[1].to_ascii_lowercase().as_str() {
            "roman" if upper => Some("upperRoman"),
            "roman" => Some("lowerRoman"),
            "alphabetic" if upper => Some("upperLetter"),
            "alphabetic" => Some("lowerLetter"),
            "arabic" => Some("decimal"),
            _ => None,
        }
    })
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
    /// A MACROBUTTON: its value is what follows the macro name, with no separator.
    macrobutton: bool,
}

/// Fields still open when a paragraph ends, carried into the next one: a
/// field's code (or result) may span paragraphs. Paragraph marks inside the
/// code are hidden with it (geneva_waste's six-paragraph IF/REF letterhead).
#[derive(Default)]
pub(in crate::docx) struct FieldCarry {
    frames: Vec<FieldFrame>,
    /// Paragraphs the open code has spanned so far.
    age: usize,
}

// ponytail: a stray begin in a broken document would hide everything after
// it; the carry is dropped after this many paragraphs instead.
const MAX_FIELD_CODE_PARAGRAPHS: usize = 12;

/// Visible content: every open field has reached its result part.
fn fields_visible(stack: &[FieldFrame]) -> bool {
    stack.iter().all(|f| f.seen_sep)
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
            macrobutton: false,
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

/// The display text of a `MACROBUTTON` instruction once its macro name is
/// complete: everything after the name. The field has no separator — the text
/// or graphic after the macro name is its value (ISO/IEC 29500-1 §17.16.5.34).
fn macrobutton_display(instr: &str) -> Option<&str> {
    let (keyword, rest) = instr.trim_start().split_once(char::is_whitespace)?;
    if !keyword.eq_ignore_ascii_case("MACROBUTTON") {
        return None;
    }
    rest.trim_start()
        .split_once(char::is_whitespace)
        .map(|(_, display)| display)
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
    /// The paragraph mark lies inside a field's code: hidden, the next
    /// paragraph continues this one.
    pub(super) mark_in_field_code: bool,
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
    lang_bidi: Option<String>,
    complex: ComplexProps,
}

/// The paragraph mark's formatting, resolved like any run: its rPr, its
/// rStyle, the paragraph style, then docDefaults.
pub(super) fn paragraph_mark_run<R: Read + Seek>(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
    ctx: &ParseContext<'_, R>,
) -> Run {
    let rpr = ppr.and_then(|ppr| wml(ppr, "rPr"));
    let char_style = rpr
        .and_then(|rpr| wml_attr(rpr, "rStyle"))
        .and_then(|id| ctx.styles.character_styles.get(id));
    ParagraphRunDefaults::from_style(para_style, &ctx.styles.defaults)
        .resolve_run_format(rpr, char_style, None, ctx.theme)
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
            lang_bidi: style_or_clone(|s| s.lang_bidi.as_ref(), &defaults.lang_bidi),
            complex: para_style.map_or(defaults.complex, |s| s.complex.or(defaults.complex)),
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
        let complex = own
            .complex
            .or(char_style.map_or(ComplexProps::default(), |c| c.complex))
            .or(self.complex);
        let cs_font_name = own
            .cs_font
            .or_else(|| char_style.and_then(|cs| cs.cs_font.clone()))
            .or_else(|| self.cs_font.clone());
        let (font_name, font_size, bold, italic) = if complex.applies() {
            (
                cs_font_name
                    .clone()
                    .unwrap_or_else(|| DEFAULT_CS_FONT.into()),
                // ponytail: no szCs anywhere keeps sz; Word may use its own 10pt default
                complex.size.unwrap_or(font_size),
                complex.bold.unwrap_or(false),
                complex.italic.unwrap_or(false),
            )
        } else {
            (
                own.font_name
                    .or(char_style_font_name)
                    .unwrap_or_else(|| self.font_name.clone()),
                font_size,
                own.bold.or_else(|| cs(|c| c.bold)).unwrap_or(self.bold),
                own.italic
                    .or_else(|| cs(|c| c.italic))
                    .unwrap_or(self.italic),
            )
        };

        Run {
            font_size,
            font_name,
            east_asia_font_name: own
                .east_asia_font
                .or_else(|| char_style.and_then(|cs| cs.east_asia_font.clone()))
                .or_else(|| self.east_asia_font.clone()),
            // ponytail: run and character style only; a paragraph style's hint
            // needs a fixture that shows it.
            east_asia_hint: own
                .east_asia_hint
                .or_else(|| char_style.and_then(|cs| cs.east_asia_hint))
                .unwrap_or(false),
            cs_font_name,
            bold,
            italic,
            rtl: complex.rtl == Some(true),
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
            text_lang_bidi: own
                .lang_bidi
                .or_else(|| char_style.and_then(|cs| cs.lang_bidi.clone()))
                .or_else(|| self.lang_bidi.clone())
                .map(Into::into),
            font_size_from_default,
            font_name_from_default,
            bold_is_direct: own.bold.is_some(),
            italic_is_direct: own.italic.is_some(),
            ..Run::default()
        }
    }
}

/// Map the legacy Word-97 run toggles (mutually exclusive per spec) to the
/// gray copies Word's PDF export draws (online export probe, Arial 10/20/40pt,
/// and case66 at 28pt). Shadow: a C0C0C0 copy down-right by size/24 on a
/// 0.25pt grid. Emboss/imprint: a 0.6 gray copy 1pt down-right/up-left of the
/// face at every size, plus a black one 1pt the other way.
fn legacy_text_shadow(
    shadow: bool,
    emboss: bool,
    imprint: bool,
    font_size: f32,
) -> Option<TextShadow> {
    let (color, d) = if emboss || imprint {
        ([153, 153, 153], if imprint { -1.0 } else { 1.0 })
    } else {
        ([192, 192, 192], (font_size / 24.0 * 4.0).round() / 4.0)
    };
    (shadow || emboss || imprint).then_some(TextShadow {
        color,
        offset_x: d,
        offset_y: -d,
        alpha: 1.0,
        relief: emboss || imprint,
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

/// The cs font of a document that names none (see `split_run_by_script`).
const DEFAULT_CS_FONT: &str = "Arial";

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
    let cs_font = run
        .cs_font_name
        .clone()
        .unwrap_or_else(|| DEFAULT_CS_FONT.into());
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
    let mut after_cs = false;
    for (i, ch) in run.text.char_indices() {
        // Whitespace inherits an East Asian context. Between complex-script
        // letters it keeps the run's font: arabic_rice's Arabic words step by
        // Aptos's 2.44pt space, not Arial's 3.33pt (its lines stay Arial-tall,
        // as whitespace never sizes a line).
        let font = if ch.is_whitespace() && i > 0 && !after_cs {
            current
        } else {
            after_cs = is_complex_script_char(ch);
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

/// Hangs each pending comment anchor whose reference follows at most `built`
/// run nodes on the last run built so far.
fn attach_comment_anchors(runs: &mut [Run], pending: &mut Vec<(usize, u32)>, built: usize) {
    let Some(last) = runs.last_mut() else { return };
    pending.retain(|&(k, id)| {
        let due = k <= built;
        if due {
            last.comment_anchor_ids.push(id);
        }
        !due
    });
}

fn is_comment_reference_run(node: roxmltree::Node) -> bool {
    node.children()
        .any(|c| c.has_tag_name((WML_NS, "commentReference")))
}

/// A run node with what its wrappers give it: hyperlink URL (and whether it
/// is an in-document anchor), open comments, and the ruby whose base it is.
type RunNode<'a> = (
    roxmltree::Node<'a, 'a>,
    Option<String>,
    bool,
    Vec<u32>,
    Option<Arc<Ruby>>,
);

fn collect_run_nodes<'a>(
    parent: roxmltree::Node<'a, 'a>,
    rels: &HashMap<String, String>,
    out: &mut Vec<RunNode<'a>>,
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
            // A ruby's base text stands in the line, its runs sharing the guide.
            if let Some(ruby) = wml(child, "ruby") {
                let first = out.len();
                if let Some(base) = wml(ruby, "rubyBase") {
                    collect_run_nodes(base, rels, out, active_comments);
                }
                let guide = ruby_of(ruby);
                for (.., r) in &mut out[first..] {
                    *r = Some(guide.clone());
                }
                continue;
            }
            out.push((child, None, false, active_comments.clone(), None));
        } else if is_wml && name == "br" {
            // A w:br straight under w:p is malformed, but Word breaks there
            // (a citation's URL then starts the next line).
            out.push((child, None, false, active_comments.clone(), None));
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
            for (_, run_url, anchor_only, ..) in &mut out[first..] {
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
                out.push((om, None, false, active_comments.clone(), None));
            }
        } else if ns == Some(MATH_NS) && name == "oMath" {
            out.push((child, None, false, active_comments.clone(), None));
        }
    }
}

fn anchor_precedes_text(runs: &[Run]) -> bool {
    runs.iter().all(|run| {
        run.text.trim().is_empty()
            && !run.is_line_break
            && !run.is_tab
            && run.inline_image.is_none()
    })
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
     $inline_chart:expr, $smartart:expr, $connectors:expr, $pending:expr) => {
        match $result {
            Some(RunDrawingResult::Inline(img)) => {
                $runs.push(Run {
                    inline_image: Some(img),
                    ..$fmt.minimal_run()
                });
            }
            Some(RunDrawingResult::Floating(mut fi)) => {
                fi.anchor_at_start = $pending.is_empty() && anchor_precedes_text(&$runs);
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
                        RunDrawingResult::Floating(mut fi) => {
                            fi.anchor_at_start =
                                $pending.is_empty() && anchor_precedes_text(&$runs);
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
/// Both absent, or the same shared value.
fn same_arc<T: ?Sized>(a: &Option<Arc<T>>, b: &Option<Arc<T>>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => Arc::ptr_eq(a, b),
        _ => false,
    }
}

/// A `w:ruby`'s guide text and its placement.
/// ponytail: the guide is drawn in its base text's font; read `w:rt`'s run
/// formatting when a document's guide uses another face.
fn ruby_of(ruby: roxmltree::Node) -> Arc<Ruby> {
    let pr = wml(ruby, "rubyPr");
    let text = wml(ruby, "rt")
        .into_iter()
        .flat_map(|rt| rt.descendants())
        .filter(|n| n.has_tag_name((WML_NS, "t")))
        .filter_map(|t| t.text())
        .collect();
    Arc::new(Ruby {
        text,
        font_size: pr.and_then(|pr| half_points(pr, "hps")).unwrap_or(5.0),
        raise: pr.and_then(|pr| half_points(pr, "hpsRaise")).unwrap_or(10.0),
    })
}

/// Word often splits a single word across multiple `w:r` elements for revision
/// tracking (different rsidR). Without merging, the layout engine treats each
/// run's text independently, allowing line breaks mid-word.
fn merge_compatible_runs(runs: Vec<Run>) -> Vec<Run> {
    if runs.len() <= 1 {
        return runs;
    }
    let mut result: Vec<Run> = Vec::with_capacity(runs.len());
    for mut run in runs {
        // customMarkFollows: the text after the reference is its mark, so it
        // carries the note (and its tag) itself.
        if let Some(prev) = result.last()
            && prev.custom_note_mark
            && prev.text.is_empty()
            && !run.text.is_empty()
            && run.footnote_id.is_none()
            && run.endnote_id.is_none()
        {
            run.footnote_id = prev.footnote_id;
            run.endnote_id = prev.endnote_id;
            run.custom_note_mark = true;
            result.pop();
        }
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
                && prev.math.is_none()
                && run.math.is_none()
                && prev.font_name == run.font_name
                && prev.east_asia_font_name == run.east_asia_font_name
                // The hint only tells the ambiguous characters apart; merging
                // across it otherwise keeps autoSpaceDE gaps at run boundaries.
                && (prev.east_asia_hint == run.east_asia_hint
                    || !(prev.text.chars().any(super::is_hint_ambiguous)
                        || run.text.chars().any(super::is_hint_ambiguous)))
                && prev.rtl == run.rtl
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
                && prev.comment_anchor_ids.is_empty()
                && run.comment_anchor_ids.is_empty()
                // One Formula per math zone, never merged into the text around it.
                && same_arc(&prev.formula, &run.formula)
                // One ruby's base runs merge only with each other.
                && same_arc(&prev.ruby, &run.ruby)
        });
        if can_merge {
            let prev = result.last_mut().unwrap();
            let base = prev.text.len();
            if prev.rsid != run.rsid {
                prev.source_breaks.push(base);
            }
            prev.rsid = run.rsid;
            prev.source_breaks.extend(run.source_breaks.iter().map(|b| base + b));
            prev.text.push_str(&run.text);
        } else {
            result.push(run);
        }
    }
    result
}

/// A paragraph's alignment, when it sets one: its display math's own, else
/// `w:jc`, else its style's.
pub(super) fn paragraph_alignment(
    para: roxmltree::Node,
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
) -> Option<crate::model::Alignment> {
    display_math_alignment(para)
        .or_else(|| {
            ppr.and_then(|ppr| wml_attr(ppr, "jc"))
                .map(super::styles::parse_alignment)
        })
        .or_else(|| para_style.and_then(|s| s.alignment))
}

/// The alignment of a paragraph holding display math (`m:oMathPara`): its
/// own `m:oMathParaPr/m:jc`, centred by default (`centerGroup`), whatever the
/// paragraph's `w:jc` says. A math paragraph has its own justification (ISO
/// 29500 Annex L); Word centres a `w:jc="both"` cell's equation. `None` for
/// paragraphs without one.
fn display_math_alignment(para: roxmltree::Node) -> Option<crate::model::Alignment> {
    use crate::model::Alignment;
    let omp = display_math(para)?;
    let jc = math_val(omp, "oMathParaPr", "jc");
    Some(match jc {
        Some("left") => Alignment::Left,
        Some("right") => Alignment::Right,
        // center / centerGroup / absent
        _ => Alignment::Center,
    })
}

/// The paragraph's display math (`m:oMathPara`), if it holds any.
pub(super) fn display_math<'a>(para: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    para.children()
        .find(|n| n.has_tag_name((MATH_NS, "oMathPara")))
}

/// One `m:oMath` as a single run: the zone laid out as a box, its linear text
/// for extraction, its spoken form for the Formula tag.
fn math_zone_run(
    omath: roxmltree::Node,
    defaults: &ParagraphRunDefaults,
    theme: &ThemeFonts,
) -> Option<Run> {
    let parser = super::math::MathParser {
        format: &|rpr| defaults.resolve_run_format(rpr, None, None, theme),
    };
    let root = parser.row(omath);
    if root.is_empty() {
        return None;
    }
    let text: String = omath
        .descendants()
        .filter(|n| n.has_tag_name((MATH_NS, "r")))
        .map(math_run_text)
        .collect();
    // The zone takes its first run's size (and the paragraph's otherwise).
    let first_rpr = omath
        .descendants()
        .find(|n| n.has_tag_name((MATH_NS, "r")))
        .and_then(|r| wml(r, "rPr"));
    let mut run = (parser.format)(first_rpr).text_run(text, None);
    let spoken = super::math_speech::speak(omath);
    run.formula = (!spoken.is_empty()).then(|| std::sync::Arc::<str>::from(spoken));
    run.math = Some(std::sync::Arc::new(crate::model::MathZone {
        root,
        display: omath
            .parent()
            .is_some_and(|p| p.has_tag_name((MATH_NS, "oMathPara"))),
    }));
    Some(run)
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

    let mut run_nodes: Vec<RunNode> = Vec::new();
    let mut active_comments: Vec<u32> = ctx
        .comments
        .open_at
        .get(&para_node.range().start)
        .cloned()
        .unwrap_or_default();
    collect_run_nodes(para_node, ctx.rels, &mut run_nodes, &mut active_comments);
    // Word shows only comments the text references; a range without its
    // w:commentReference gets no tint and no balloon.
    for (_, _, _, ids, _) in &mut run_nodes {
        ids.retain(|id| ctx.comments.shown.contains(id));
    }
    // A referenced comment whose range holds no text (empty, or missing)
    // anchors at its reference: after the run nodes that precede it.
    let mut ref_anchors: Vec<(usize, u32)> = para_node
        .descendants()
        .filter(|n| n.has_tag_name((WML_NS, "commentReference")))
        .filter_map(|n| {
            let id = n.attribute((WML_NS, "id"))?.parse::<u32>().ok()?;
            if !ctx.comments.textless.contains(&id) {
                return None;
            }
            let start = n.range().start;
            Some((
                run_nodes
                    .iter()
                    .filter(|(r, ..)| r.range().start < start)
                    .count(),
                id,
            ))
        })
        .collect();

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
    let mut field_stack: Vec<FieldFrame> = std::mem::take(&mut ctx.fields.frames);

    for (node_idx, (run_node, hyperlink_url, is_anchor_hyperlink, comment_ids, ruby)) in
        run_nodes.into_iter().enumerate()
    {
        attach_comment_anchors(&mut runs, &mut ref_anchors, node_idx);
        if run_node.has_tag_name((MATH_NS, "oMath")) {
            runs.extend(math_zone_run(run_node, &defaults, ctx.theme));
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

        let mut fmt = defaults.resolve_run_format(rpr, char_style, char_style_id_str, ctx.theme);
        fmt.rsid = ["rsidR", "rsidRPr"].iter().fold(0, |acc, a| {
            let v = run_node.attribute((WML_NS, *a));
            acc << 32 | v.and_then(|v| u32::from_str_radix(v, 16).ok()).unwrap_or(0) as u64
        });

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
                            connectors,
                            pending_text
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
                        if let Some(connector) = parse_connector_from_vml(pict) {
                            connectors.push(connector);
                        } else if let Some(tb) = parse_textbox_from_vml(pict, ctx) {
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
                        let parent_visible = fields_visible(&field_stack);
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
                                // The box is a glyph of its own `w:size`: a 12pt
                                // check box lifts a 10pt line to 12pt metrics
                                // (geneva_waste's Non/Oui lines, +1.5pt).
                                font_size: cb.size,
                                hyperlink_url: hyperlink_url.clone(),
                                ..fmt.styled_run()
                            });
                            continue;
                        }
                        if f.macrobutton {
                            let text = f.result.trim();
                            if !text.is_empty() {
                                runs.push(Run {
                                    text: text.to_string(),
                                    hyperlink_url: hyperlink_url.clone(),
                                    ..fmt.styled_run()
                                });
                            }
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
                        } else if f.result.is_empty()
                            && let Some(text) = parse_eq_overstrike(&f.instr)
                        {
                            // EQ fields cache no result — Word computes them at
                            // layout time. Derive the visible text from the
                            // instruction so a checkbox does not vanish.
                            runs.push(Run {
                                text,
                                hyperlink_url: hyperlink_url.clone(),
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
                        && let Some(t) = child.text()
                    {
                        if f.macrobutton {
                            f.result.push_str(t);
                        } else if !f.seen_sep {
                            f.push_instr(t, simple.is_some());
                            if let Some(display) = macrobutton_display(&f.instr) {
                                // The macro name ends the code: what follows (text
                                // here, a picture in later runs) is the value.
                                f.result = display.to_string();
                                f.seen_sep = true;
                                f.macrobutton = true;
                            }
                        }
                    }
                }
                "t" => {
                    let visible = fields_visible(&field_stack);
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
                    let visible = fields_visible(&field_stack);
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
                    let visible = fields_visible(&field_stack);
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
                "drawing" if !fields_visible(&field_stack) => {}
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
                        connectors,
                        pending_text
                    );
                }
                "pict" if fields_visible(&field_stack) => {
                    if let Some(result) = super::group::parse_vml_group(child, ctx) {
                        handle_drawing_result!(
                            Some(result),
                            fmt,
                            runs,
                            floating_images,
                            textboxes,
                            inline_chart,
                            smartart,
                            connectors,
                            pending_text
                        );
                    } else if let Some(hr) = parse_vml_horizontal_rule(child) {
                        horizontal_rule = Some(hr);
                    } else if is_vml_picture(child) {
                        flush_pending(&mut pending_text, &mut runs);
                        if let Some(mut fi) = parse_object_floating_image(child, ctx) {
                            fi.anchor_at_start = anchor_precedes_text(&runs);
                            push_floating(&mut floating_images, &textboxes, fi);
                        } else if let Some(img) = parse_object_inline_image(child, ctx) {
                            runs.push(Run {
                                inline_image: Some(img),
                                ..fmt.minimal_run()
                            });
                        }
                    } else if let Some(connector) = parse_connector_from_vml(child) {
                        connectors.push(connector);
                    } else if let Some(tb) = parse_textbox_from_vml(child, ctx) {
                        push_textbox(&floating_images, &mut textboxes, tb);
                    }
                }
                "object" if fields_visible(&field_stack) => {
                    flush_pending(&mut pending_text, &mut runs);
                    if let Some(mut fi) = parse_object_floating_image(child, ctx) {
                        fi.anchor_at_start = anchor_precedes_text(&runs);
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
                            custom_note_mark: custom_note_mark(child),
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
                            custom_note_mark: custom_note_mark(child),
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
        if ruby.is_some() {
            for r in &mut runs[runs_before..] {
                r.ruby = ruby.clone();
            }
        }
    }

    attach_comment_anchors(&mut runs, &mut ref_anchors, usize::MAX);

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
    // After merging: Word splits words over runs that look alike (rsid), and
    // a letter's form depends on its neighbours.
    // ponytail: no joining across runs that differ (a bold letter in a word)
    for run in &mut runs {
        if run.text.chars().any(is_complex_script_char) {
            run.text = super::arabic::shape(&run.text);
        }
    }

    if field_stack.is_empty() || ctx.fields.age >= MAX_FIELD_CODE_PARAGRAPHS {
        field_stack.clear();
        ctx.fields.age = 0;
    } else {
        ctx.fields.age += 1;
    }
    let mark_in_field_code = !fields_visible(&field_stack);
    ctx.fields.frames = field_stack;

    ParsedRuns {
        runs,
        mark_in_field_code,
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

    #[test]
    fn eq_overstrike_checkbox_and_other_fields() {
        assert_eq!(parse_eq_overstrike(r"EQ \o\ac(□,√)").as_deref(), Some("☑"));
        assert_eq!(parse_eq_overstrike(r" eq \O\ac( ☐ , √ )").as_deref(), Some("☑"));
        assert_eq!(parse_eq_overstrike(r"EQ \o(a,´)").as_deref(), Some("a"));
        assert_eq!(parse_eq_overstrike(r"EQ \f(1,2)"), None);
        assert_eq!(parse_eq_overstrike(r"PAGE \* MERGEFORMAT"), None);
        assert_eq!(parse_eq_overstrike(r"EQ \o()"), None);
    }

    #[test]
    fn macrobutton_value_follows_the_macro_name() {
        assert_eq!(macrobutton_display(" MACROBUTTON AutoNew "), Some(""));
        assert_eq!(
            macrobutton_display(" MACROBUTTON NoMacro [Click here] "),
            Some("[Click here] ")
        );
        assert_eq!(macrobutton_display(" MACROBUTTON AutoNew"), None);
        assert_eq!(macrobutton_display(r" PAGE \* MERGEFORMAT "), None);
    }

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
        // Spaces and digits keep the run's font.
        let expect = |cs| {
            vec![
                seg("rice ", "Aptos"),
                seg("الأرز", cs),
                seg(" 150 x", "Aptos"),
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

    #[test]
    fn merge_compatible_runs_keeps_rsid_boundaries() {
        let with_rsid = |text: &str, rsid| Run {
            rsid,
            ..run(text, None, None)
        };
        let runs = vec![
            with_rsid("one ", 1),
            with_rsid("two ", 1),
            with_rsid("thr", 2),
            with_rsid("ee", 1),
        ];
        let merged = merge_compatible_runs(runs);
        assert_eq!(merged.len(), 1);
        assert_eq!(merged[0].source_breaks, vec![8, 11]);
    }

    #[test]
    fn imprint_casts_up_left_and_emboss_down_right() {
        let imprint = legacy_text_shadow(false, false, true, 28.0).unwrap();
        assert_eq!((imprint.offset_x, imprint.offset_y), (-1.0, 1.0));
        assert!(imprint.relief);
        let emboss = legacy_text_shadow(false, true, false, 28.0).unwrap();
        assert_eq!((emboss.offset_x, emboss.offset_y), (1.0, -1.0));
        assert!(emboss.relief);
        // Word's shadow offsets at 10/20/28/40pt.
        for (size, d) in [(10.0, 0.5), (20.0, 0.75), (28.0, 1.25), (40.0, 1.75)] {
            let shadow = legacy_text_shadow(true, false, false, size).unwrap();
            assert_eq!(
                (shadow.offset_x, shadow.offset_y, shadow.relief),
                (d, -d, false)
            );
        }
    }
}

#[cfg(test)]
mod leading_anchor_tests {
    use super::*;

    #[test]
    fn leading_anchor_ignores_blank_text_but_not_flow_content() {
        assert!(anchor_precedes_text(&[]));
        assert!(anchor_precedes_text(&[Run {
            text: "  ".into(),
            ..Run::default()
        }]));
        assert!(!anchor_precedes_text(&[Run {
            text: "Text".into(),
            ..Run::default()
        }]));
        assert!(!anchor_precedes_text(&[Run {
            is_line_break: true,
            ..Run::default()
        }]));
        assert!(!anchor_precedes_text(&[Run {
            is_tab: true,
            ..Run::default()
        }]));
    }
}
