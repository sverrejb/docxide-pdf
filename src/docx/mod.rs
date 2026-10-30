mod alt_chunk;
pub(crate) mod arabic;
mod charts;
mod color;
mod comments;
mod embedded_fonts;
pub(crate) mod emf;
mod group;
mod headers_footers;
mod images;
mod math;
mod math_speech;
pub(crate) mod numbering;
mod paragraph;
mod runs;
mod sections;
mod settings;
pub(crate) mod smartart;
mod styles;
pub(crate) mod svg;
mod tables;
mod textbox;
mod wmf;
mod wordart;

use std::collections::HashMap;
use std::io::Read;

use crate::error::Error;
use crate::model::{
    Block, BorderStyle, CellBorder, DocGridType, Document, DropCap, FrameProperties, HRelativeFrom,
    HorizontalPosition, LineSpacing, ParagraphBorder, ParagraphBorders, Section, SectionBreakType,
    SectionProperties, TabAlignment, TabStop, VRelativeFrom, VerticalPosition,
};

use styles::{ParagraphStyle, StyleDefaults, parse_line_spacing, parse_styles, parse_theme};

use embedded_fonts::parse_font_table;
use headers_footers::{parse_endnotes, parse_footnotes};
use numbering::parse_numbering;
use relationships::parse_relationships;
use sections::parse_section_properties;
use settings::parse_settings;
use tables::parse_table_node;

pub(super) const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const WPD_NS: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
pub(super) const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
pub(super) const REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const MC_NS_TOP: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
pub(super) const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub(super) const DSP_NS: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";
pub(super) const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
pub(super) const VML_NS: &str = "urn:schemas-microsoft-com:vml";
pub(super) const W10_NS: &str = "urn:schemas-microsoft-com:office:word";
pub(super) const OFFICE_NS: &str = "urn:schemas-microsoft-com:office:office";
pub(super) const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
pub(super) const DIAGRAM_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
pub(super) const MATH_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
pub(super) const WPC_NS: &str =
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
pub(super) const WPG_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

pub(super) fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

pub(super) fn emu_to_pts(emu: f32) -> f32 {
    emu / 12700.0
}

/// A numeric attribute; `None` when absent or unparsable.
pub(super) fn f32_attr<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    node.attribute(attr).and_then(|v| v.parse::<f32>().ok())
}

/// An EMU attribute in points; `None` when absent.
pub(super) fn emu_attr_opt<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    f32_attr(node, attr).map(emu_to_pts)
}

/// An EMU attribute in points; 0 when absent.
pub(super) fn emu_attr(node: roxmltree::Node, attr: &str) -> f32 {
    emu_attr_opt(node, attr).unwrap_or(0.0)
}

/// A DrawingML angle attribute (60000ths of a degree) in degrees.
pub(super) fn angle_attr<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    f32_attr(node, attr).map(|v| v / 60_000.0)
}

/// A DrawingML percentage attribute (`ST_Percentage` as 1/1000 of a percent)
/// as a fraction.
pub(super) fn frac_attr<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    f32_attr(node, attr).map(|v| v / 100_000.0)
}

/// ST_OnOff truthiness (§17.17.4): "1", "true", and "on" all mean on.
pub(super) fn parse_on_off(v: &str) -> bool {
    matches!(v, "1" | "true" | "on")
}

/// Parse a VML/CSS length like "12.5pt" — the "pt" suffix and surrounding
/// whitespace are optional.
pub(super) fn parse_pt(s: &str) -> Option<f32> {
    s.trim().trim_end_matches("pt").trim().parse::<f32>().ok()
}

/// Find a child element by name and namespace.
pub(super) fn find_child<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
    namespace: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    node.children().find(|n| n.has_tag_name((namespace, name)))
}

pub(super) fn dml<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, DML_NS)
}

/// All child elements with the given name and namespace.
pub(super) fn find_children<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
    namespace: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    node.children()
        .filter(move |n| n.has_tag_name((namespace, name)))
}

pub(super) fn dml_children<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    find_children(parent, name, DML_NS)
}

pub(super) fn wpd<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, WPD_NS)
}

pub(super) fn wps<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, WPS_NS)
}

pub(super) fn dsp<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, DSP_NS)
}

pub(super) fn chart_ns<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, CHART_NS)
}

pub(super) fn chart_ns_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    chart_ns(node, child).and_then(|n| n.attribute("val"))
}

pub(super) fn chart_ns_children<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    find_children(parent, name, CHART_NS)
}

pub(in crate::docx) struct ParseContext<'a, R: std::io::Read + std::io::Seek> {
    pub(in crate::docx) styles: &'a styles::StylesInfo,
    pub(in crate::docx) theme: &'a styles::ThemeFonts,
    pub(in crate::docx) rels: &'a HashMap<String, String>,
    pub(in crate::docx) zip: &'a mut zip::ZipArchive<R>,
    pub(in crate::docx) numbering: &'a numbering::NumberingInfo,
    /// The body's comment ranges (`comment_ranges`); empty for headers,
    /// footers and notes.
    pub(in crate::docx) comments: CommentRanges,
    /// Fields left open by the previous paragraph (`runs::FieldCarry`).
    pub(in crate::docx) fields: runs::FieldCarry,
    /// What the current section's character grid adds to one character
    /// (`sections::grid_char_delta`), the unit of `w:ind` `*Chars` values.
    pub(in crate::docx) grid_char_delta: f32,
}

#[derive(Default)]
pub(in crate::docx) struct CommentRanges {
    /// Comment ids Word shows a balloon for: referenced by a
    /// `w:commentReference` in the body.
    pub(in crate::docx) shown: std::collections::HashSet<u32>,
    /// Comment ids whose range holds no text (or that have no range).
    pub(in crate::docx) textless: std::collections::HashSet<u32>,
    /// Comment ids open where a paragraph starts, by its byte offset.
    pub(in crate::docx) open_at: HashMap<usize, Vec<u32>>,
    /// Every `w:commentReference` id in document order.
    pub(in crate::docx) reference_order: Vec<u32>,
}

/// Comment ranges as Word reads them, over the whole body (a range may start
/// between blocks and span paragraphs and tables).
fn comment_ranges(root: roxmltree::Node) -> CommentRanges {
    let id = |n: roxmltree::Node| n.attribute((WML_NS, "id"))?.parse::<u32>().ok();
    let mut open: Vec<u32> = Vec::new();
    let mut open_at = HashMap::new();
    let mut with_text = std::collections::HashSet::new();
    let mut reference_order = Vec::new();
    let mut unballooned = std::collections::HashSet::new();
    for n in root.descendants() {
        if n.has_tag_name((WML_NS, "commentRangeStart")) {
            open.extend(id(n));
        } else if n.has_tag_name((WML_NS, "commentRangeEnd")) {
            open.retain(|o| Some(*o) != id(n));
            // Word probes: a range over text ending before any text of its
            // paragraph, or between blocks, gets no balloon wherever its reference is
            // (after a table or a paragraph alike); one ending after text,
            // in a cell's last paragraph too, gets one.
            let text_before = n
                .ancestors()
                .find(|a| a.has_tag_name((WML_NS, "p")))
                .is_some_and(|p| {
                    p.descendants().take_while(|d| *d != n).any(|d| {
                        d.has_tag_name((WML_NS, "t")) && d.text().is_some_and(|t| !t.is_empty())
                    })
                });
            if !text_before {
                unballooned.extend(id(n));
            }
        } else if n.has_tag_name((WML_NS, "commentReference")) {
            reference_order.extend(id(n));
        } else if n.has_tag_name((WML_NS, "t")) && n.text().is_some_and(|t| !t.is_empty()) {
            with_text.extend(open.iter().copied());
        } else if n.has_tag_name((WML_NS, "p")) && !open.is_empty() {
            open_at.insert(n.range().start, open.clone());
        }
    }
    // An empty range there keeps its balloon (it hangs at its reference).
    let referenced: std::collections::HashSet<u32> = reference_order.iter().copied().collect();
    let shown = &referenced - &(&unballooned & &with_text);
    CommentRanges {
        textless: &shown - &with_text,
        shown,
        open_at,
        reference_order,
    }
}

pub(crate) fn is_east_asian_char(ch: char) -> bool {
    matches!(ch as u32,
        0x2E80..=0x2EFF   // CJK Radicals Supplement
        | 0x2F00..=0x2FDF // Kangxi Radicals
        | 0x2FF0..=0x2FFF // Ideographic Description Characters
        | 0x3000..=0x303F // CJK Symbols and Punctuation
        | 0x3040..=0x309F // Hiragana
        | 0x30A0..=0x30FF // Katakana
        | 0x3100..=0x312F // Bopomofo
        | 0x3130..=0x318F // Hangul Compatibility Jamo
        | 0x31A0..=0x31BF // Bopomofo Extended
        | 0x31F0..=0x31FF // Katakana Phonetic Extensions
        | 0x3200..=0x32FF // Enclosed CJK Letters and Months
        | 0x3300..=0x33FF // CJK Compatibility
        | 0x3400..=0x4DBF // CJK Unified Ideographs Extension A
        | 0x4E00..=0x9FFF // CJK Unified Ideographs
        | 0xAC00..=0xD7AF // Hangul Syllables
        | 0xF900..=0xFAFF // CJK Compatibility Ideographs
        | 0xFE30..=0xFE4F // CJK Compatibility Forms
        | 0xFF00..=0xFFEF // Halfwidth and Fullwidth Forms
        | 0x1100..=0x11FF // Hangul Jamo
        | 0x20000..=0x2A6DF // CJK Unified Ideographs Extension B
        | 0x2A700..=0x2B73F // CJK Unified Ideographs Extension C
        | 0x2B740..=0x2B81F // CJK Unified Ideographs Extension D
    )
}

/// Characters that `w:rFonts@hint` assigns to the East Asian font: Word then
/// draws them full-width and treats them as ideographs for spacing and
/// punctuation compression (in a hinted run 在“测 gets no autoSpaceDE gap and
/// both quotes squeeze like commas, issue #139).
// ponytail: the quotes, dashes, bullets and ellipsis of General Punctuation
// plus the Latin-1 symbols; §17.3.2.26 lists many more (Greek, Cyrillic,
// box drawing, …), add them with a fixture that shows one.
pub(crate) fn is_hint_ambiguous(c: char) -> bool {
    matches!(c as u32,
        0x2013..=0x2016 // – — ‖
        | 0x2018..=0x201D // ‘ ’ “ ”
        | 0x2020..=0x2022 | 0x2025..=0x2026 | 0x2030 // † ‡ • ‥ … ‰
        | 0x2032..=0x2033 | 0x2035 | 0x203B // ′ ″ ‵ ※
        | 0x00A7 | 0x00A8 | 0x00B0 | 0x00B1 | 0x00B4 | 0x00B6 | 0x00B7 | 0x00D7 | 0x00F7
    )
}

/// Letters Word draws in the complex-script (`w:cs`) font.
// ponytail: right-to-left scripts only; Thai and Indic are complex script in
// Word too, add them with a fixture that shows their fallback font.
pub(crate) fn is_complex_script_char(ch: char) -> bool {
    matches!(ch as u32,
        0x0590..=0x07BF   // Hebrew, Arabic, Syriac, Arabic Supplement, Thaana
        | 0x08A0..=0x08FF // Arabic Extended-A
        | 0xFB1D..=0xFDFF // Hebrew and Arabic Presentation Forms-A
        | 0xFE70..=0xFEFE // Arabic Presentation Forms-B (U+FEFF is the BOM)
    )
}

pub(super) fn parse_hex_color(val: &str) -> Option<[u8; 3]> {
    if val == "auto" || val.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&val[0..2], 16).ok()?;
    let g = u8::from_str_radix(&val[2..4], 16).ok()?;
    let b = u8::from_str_radix(&val[4..6], 16).ok()?;
    Some([r, g, b])
}

pub(super) fn parse_text_color(val: &str) -> Option<[u8; 3]> {
    if val == "auto" {
        return Some([0, 0, 0]);
    }
    parse_hex_color(val)
}

/// Parse a `w:shd` child element on a `w:rPr` (or similar) node.
/// Returns `Some(rgb)` when an explicit color is present, else `None`
/// (caller should inherit from style/basedOn). In Word, `w:val="nil"` and
/// `w:fill="auto"` are *not* treated as overrides that clear inherited
/// shading — they simply mean "no color specified here."
///
/// `w:themeFill` resolution is TODO.
pub(super) fn parse_run_shd(parent: roxmltree::Node) -> Option<[u8; 3]> {
    wml(parent, "shd").and_then(shd_color)
}

/// The solid color a `w:shd` paints: its `solid`/`pctNN` pattern of `w:color`
/// (auto = black) laid over `w:fill` (auto = white), §17.18.78. A `solid`
/// shd paints `w:color` alone (CC99FF over fill auto is CC99FF).
pub(super) fn shd_color(shd: roxmltree::Node) -> Option<[u8; 3]> {
    let fill = shd
        .attribute((WML_NS, "fill"))
        .filter(|f| *f != "auto" && *f != "none")
        .and_then(parse_hex_color);
    let coverage = match shd.attribute((WML_NS, "val")) {
        Some("nil") => return None,
        Some("solid") => 1.0,
        Some(v) => v
            .strip_prefix("pct")
            .and_then(|p| p.parse::<f32>().ok())
            .map_or(0.0, |p| (p / 100.0).clamp(0.0, 1.0)),
        None => 0.0,
    };
    if coverage == 0.0 {
        return fill;
    }
    let ink = shd
        .attribute((WML_NS, "color"))
        .and_then(parse_hex_color)
        .unwrap_or([0, 0, 0]);
    let bg = fill.unwrap_or([255, 255, 255]);
    Some(std::array::from_fn(|i| {
        (bg[i] as f32 * (1.0 - coverage) + ink[i] as f32 * coverage).round() as u8
    }))
}

pub(super) fn highlight_color(name: &str) -> Option<[u8; 3]> {
    match name {
        "yellow" => Some([255, 255, 0]),
        "green" => Some([0, 255, 0]),
        "cyan" => Some([0, 255, 255]),
        "magenta" => Some([255, 0, 255]),
        "red" => Some([255, 0, 0]),
        "blue" => Some([0, 0, 255]),
        "darkYellow" => Some([128, 128, 0]),
        "darkGreen" => Some([0, 128, 0]),
        "darkCyan" => Some([0, 128, 128]),
        "darkMagenta" => Some([128, 0, 128]),
        "darkRed" => Some([128, 0, 0]),
        "darkBlue" => Some([0, 0, 128]),
        "lightGray" => Some([192, 192, 192]),
        "darkGray" => Some([128, 128, 128]),
        "black" => Some([0, 0, 0]),
        "white" => Some([255, 255, 255]),
        _ => None,
    }
}

pub(super) fn wml_bool(parent: roxmltree::Node, name: &str) -> Option<bool> {
    wml(parent, name).map(|n| {
        n.attribute((WML_NS, "val"))
            .is_none_or(|v| v != "0" && v != "false")
    })
}

pub(super) fn wml<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, WML_NS)
}

pub(super) fn wml_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    wml(node, child).and_then(|n| n.attribute((WML_NS, "val")))
}

/// First math-namespace child element with the given local name.
pub(super) fn math_child<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(parent, name, MATH_NS)
}

/// A math property's `m:val`: `<pr><name m:val="…"/></pr>` under `node`.
pub(super) fn math_val<'a>(node: roxmltree::Node<'a, 'a>, pr: &str, name: &str) -> Option<&'a str> {
    math_child(math_child(node, pr)?, name)?.attribute((MATH_NS, "val"))
}

/// A math property flag (`<pr><name m:val="…"/></pr>`): None when absent; an
/// element without `m:val` is on (ST_OnOff).
pub(super) fn math_bool(node: roxmltree::Node, pr: &str, name: &str) -> Option<bool> {
    let n = math_child(math_child(node, pr)?, name)?;
    Some(n.attribute((MATH_NS, "val")).is_none_or(parse_on_off))
}

/// A math run's text (its `m:t` children).
pub(super) fn math_run_text(r: roxmltree::Node) -> String {
    find_children(r, "t", MATH_NS)
        .filter_map(|t| t.text())
        .collect()
}

/// A WordprocessingML twips attribute in points. ST_TwipsMeasure also
/// takes a universal measure such as "10mm" (§22.9.2.15): a page written
/// with `w:left="10mm"` lost its margins to the 1in default.
pub(super) fn twips_attr(node: roxmltree::Node, attr: &str) -> Option<f32> {
    let v = node.attribute((WML_NS, attr))?;
    v.parse::<f32>()
        .ok()
        .map(twips_to_pts)
        .or_else(|| universal_measure_pt(v))
}

fn universal_measure_pt(v: &str) -> Option<f32> {
    let per_unit = [
        ("mm", 72.0 / 25.4),
        ("cm", 72.0 / 2.54),
        ("in", 72.0),
        ("pt", 1.0),
        ("pc", 12.0),
        ("pi", 12.0),
    ];
    per_unit
        .iter()
        .find_map(|(unit, k)| Some(v.strip_suffix(unit)?.parse::<f32>().ok()? * k))
}

pub(super) fn parse_one_border(node: roxmltree::Node) -> Option<ParagraphBorder> {
    let val = node.attribute((WML_NS, "val")).unwrap_or("none");
    if val == "none" || val == "nil" {
        return None;
    }
    let width_pt = node
        .attribute((WML_NS, "sz"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 8.0)
        .unwrap_or(0.5);
    let space_pt = node
        .attribute((WML_NS, "space"))
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0);
    let color = node
        .attribute((WML_NS, "color"))
        .and_then(parse_hex_color)
        .unwrap_or([0, 0, 0]);
    Some(ParagraphBorder {
        width_pt,
        space_pt,
        color,
        style: border_style(val),
    })
}

/// A border's `w:val` as drawn. ponytail: the thin/thick pairs' small,
/// medium and large gaps draw alike (medium measured); `triple`, `thick`,
/// 3-D and inset/outset borders draw as one line.
pub(super) fn border_style(val: &str) -> BorderStyle {
    match val {
        "dotted" => BorderStyle::Dotted,
        "dashed" => BorderStyle::Dashed,
        "dashSmallGap" => BorderStyle::DashSmallGap,
        "dotDash" | "dashDotStroked" | "dashDot" => BorderStyle::DashDot,
        "dotDotDash" | "dashDotDot" => BorderStyle::DashDotDot,
        "double" => BorderStyle::Double,
        "wave" | "doubleWave" => BorderStyle::Wave,
        v if v.starts_with("thinThickThin") => BorderStyle::ThinThickThin,
        v if v.starts_with("thinThick") => BorderStyle::ThinThick,
        v if v.starts_with("thickThin") => BorderStyle::ThickThin,
        _ => BorderStyle::Single,
    }
}

/// The pBdr sides one level of the style hierarchy sets: `None` = not set
/// here, `Some(None)` = set to none/nil. Word merges them per side through
/// basedOn and from direct formatting onto the style (probe 2026-10-30).
#[derive(Clone, Debug, Default)]
pub(super) struct ParagraphBorderSides {
    top: Option<Option<ParagraphBorder>>,
    bottom: Option<Option<ParagraphBorder>>,
    left: Option<Option<ParagraphBorder>>,
    right: Option<Option<ParagraphBorder>>,
    between: Option<Option<ParagraphBorder>>,
}

impl ParagraphBorderSides {
    /// The sides set here, the rest from `base`.
    pub(super) fn over(&self, base: &Self) -> Self {
        let side = |own: &Option<_>, base: &Option<_>| own.clone().or_else(|| base.clone());
        Self {
            top: side(&self.top, &base.top),
            bottom: side(&self.bottom, &base.bottom),
            left: side(&self.left, &base.left),
            right: side(&self.right, &base.right),
            between: side(&self.between, &base.between),
        }
    }

    pub(super) fn resolve(self) -> ParagraphBorders {
        ParagraphBorders {
            top: self.top.flatten(),
            bottom: self.bottom.flatten(),
            left: self.left.flatten(),
            right: self.right.flatten(),
            between: self.between.flatten(),
        }
    }
}

pub(super) fn parse_paragraph_borders(ppr: roxmltree::Node) -> ParagraphBorderSides {
    let Some(pbdr) = wml(ppr, "pBdr") else {
        return ParagraphBorderSides::default();
    };
    let side = |name| wml(pbdr, name).map(parse_one_border);
    ParagraphBorderSides {
        top: side("top"),
        bottom: side("bottom"),
        left: side("left"),
        right: side("right"),
        between: side("between"),
    }
}

pub(super) fn parse_cell_border(parent: roxmltree::Node, name: &str) -> CellBorder {
    let Some(n) = wml(parent, name) else {
        return CellBorder::default();
    };
    let val = n.attribute((WML_NS, "val")).unwrap_or("none");
    if val == "nil" || val == "none" {
        return CellBorder {
            is_override: true,
            ..CellBorder::default()
        };
    }
    let width = n
        .attribute((WML_NS, "sz"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 8.0)
        .unwrap_or(0.5);
    let color = n.attribute((WML_NS, "color")).and_then(parse_hex_color);
    CellBorder::visible(color, width, border_style(val))
}

fn parse_cell_border_with_fallback(
    parent: roxmltree::Node,
    primary: &str,
    fallback: &str,
) -> CellBorder {
    let border = parse_cell_border(parent, primary);
    if border.present || border.is_override {
        border
    } else {
        parse_cell_border(parent, fallback)
    }
}

/// Parse left border with "start" fallback per OOXML bidi naming.
pub(super) fn parse_cell_border_left(parent: roxmltree::Node) -> CellBorder {
    parse_cell_border_with_fallback(parent, "left", "start")
}

/// Parse right border with "end" fallback per OOXML bidi naming.
pub(super) fn parse_cell_border_right(parent: roxmltree::Node) -> CellBorder {
    parse_cell_border_with_fallback(parent, "right", "end")
}

/// A framePr's attributes, kept so a paragraph can lay its own over its style's.
pub(super) type FrameAttrs = HashMap<String, String>;

pub(super) fn frame_attrs(fp: roxmltree::Node) -> FrameAttrs {
    fp.attributes()
        .filter(|a| a.namespace() == Some(WML_NS))
        .map(|a| (a.name().to_string(), a.value().to_string()))
        .collect()
}

/// The frame a paragraph sits in: its own framePr over its style's, attribute
/// by attribute (massachusetts' "Exec Office" lines leave hAnchor and hSpace to
/// their style and so share the title line's frame).
pub(super) fn parse_frame_props(
    ppr: Option<roxmltree::Node>,
    style_attrs: Option<&FrameAttrs>,
) -> Option<FrameProperties> {
    let direct = ppr.and_then(|p| wml(p, "framePr"));
    if direct.is_none() && style_attrs.is_none() {
        return None;
    }
    let mut attrs = style_attrs.cloned().unwrap_or_default();
    attrs.extend(direct.map(frame_attrs).unwrap_or_default());
    let attr = |name: &str| attrs.get(name).map(String::as_str);
    // An inline frame sits at its place in the flow and takes its height
    // there, like any paragraph (bulgarian_farmland's three header lines).
    // ponytail: every corpus one has no width or position; give a sized
    // inline frame its own path if one turns up.
    if attr("yAlign") == Some("inline") {
        return None;
    }
    // A framePr setting nothing but its wrap (a Normal style's
    // `<w:framePr w:wrap="auto"/>`) places nothing: Word flows such
    // paragraphs as plain ones.
    if attrs.keys().all(|k| k == "wrap") {
        return None;
    }
    let twips = |name| {
        attr(name)
            .and_then(|v| v.parse::<f32>().ok())
            .map_or(0.0, twips_to_pts)
    };
    let h_anchor = match attr("hAnchor").unwrap_or("text") {
        "margin" => HRelativeFrom::Margin,
        "page" => HRelativeFrom::Page,
        _ => HRelativeFrom::Column,
    };
    let h_position = if let Some(xa) = attr("xAlign") {
        match xa {
            "center" => HorizontalPosition::AlignCenter,
            "right" | "outside" => HorizontalPosition::AlignRight,
            _ => HorizontalPosition::AlignLeft,
        }
    } else {
        HorizontalPosition::Offset(twips("x"))
    };
    // Word probe: a frame without vAnchor sits w:y below the top margin, however
    // many paragraphs come before it.
    let drop_cap = match attr("dropCap") {
        Some(kind @ ("drop" | "margin")) => Some(DropCap {
            lines: attr("lines")
                .and_then(|v| v.parse().ok())
                .unwrap_or(1)
                .max(1),
            in_margin: kind == "margin",
        }),
        _ => None,
    };
    // The spec ignores a drop cap's vertical position: it sits at the top of
    // the next paragraph.
    let v_anchor = match attr("vAnchor") {
        _ if drop_cap.is_some() => VRelativeFrom::Paragraph,
        Some("page") => VRelativeFrom::Page,
        Some("text") => VRelativeFrom::Paragraph,
        _ => VRelativeFrom::Margin,
    };
    Some(FrameProperties {
        h_relative_from: h_anchor,
        h_position,
        v_relative_from: v_anchor,
        v_position: match attr("yAlign") {
            _ if drop_cap.is_some() => VerticalPosition::Offset(0.0),
            Some("top" | "inside") => VerticalPosition::AlignTop,
            Some("center") => VerticalPosition::AlignCenter,
            Some("bottom" | "outside") => VerticalPosition::AlignBottom,
            _ => VerticalPosition::Offset(twips("y")),
        },
        width: twips("w"),
        height: twips("h"),
        height_exact: attr("hRule") == Some("exact"),
        text_below: matches!(attr("wrap"), Some("none") | Some("notBeside")),
        h_space: twips("hSpace"),
        v_space: twips("vSpace"),
        drop_cap,
    })
}

/// A paragraph's tab stops: its style's, which its own `w:tabs` clear and
/// override, in order.
pub(in crate::docx) fn resolve_tab_stops(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
) -> Vec<TabStop> {
    let mut tab_stops = para_style.map(|s| s.tab_stops.clone()).unwrap_or_default();
    let (para_tabs, para_clears) = ppr.map(parse_tab_stops_with_clears).unwrap_or_default();
    if !para_tabs.is_empty() || !para_clears.is_empty() {
        merge_tab_stops(&mut tab_stops, &para_clears, para_tabs);
        tab_stops.sort_by(|a, b| a.position.total_cmp(&b.position));
    }
    tab_stops
}

/// Merge `incoming` tab stops into `dst`: clear positions remove matching stops
/// (0.5pt tolerance), then each incoming stop overrides one at the same position
/// or is appended. Caller is responsible for sorting.
pub(super) fn merge_tab_stops(dst: &mut Vec<TabStop>, clears: &[f32], incoming: Vec<TabStop>) {
    for &clear_pos in clears {
        dst.retain(|t| (t.position - clear_pos).abs() >= 0.5);
    }
    for ts in incoming {
        if let Some(existing) = dst
            .iter_mut()
            .find(|t| (t.position - ts.position).abs() < 0.5)
        {
            *existing = ts;
        } else {
            dst.push(ts);
        }
    }
}

pub(super) fn parse_tab_stops_with_clears(ppr: roxmltree::Node) -> (Vec<TabStop>, Vec<f32>) {
    let Some(tabs) = wml(ppr, "tabs") else {
        return (vec![], vec![]);
    };
    let mut stops = Vec::new();
    let mut clears = Vec::new();
    for n in tabs.children().filter(|n| n.has_tag_name((WML_NS, "tab"))) {
        let Some(pos) = twips_attr(n, "pos") else {
            continue;
        };
        let val = n.attribute((WML_NS, "val")).unwrap_or("left");
        if val == "clear" {
            clears.push(pos);
            continue;
        }
        let alignment = match val {
            "center" => TabAlignment::Center,
            "right" => TabAlignment::Right,
            "decimal" => TabAlignment::Decimal,
            _ => TabAlignment::Left,
        };
        let leader = n.attribute((WML_NS, "leader")).and_then(|l| match l {
            "dot" => Some('.'),
            "hyphen" => Some('-'),
            "underscore" => Some('_'),
            _ => None,
        });
        stops.push(TabStop {
            position: pos,
            alignment,
            leader,
        });
    }
    stops.sort_by(|a, b| a.position.total_cmp(&b.position));
    (stops, clears)
}

pub(super) fn resolve_theme_color_key(scheme_name: &str) -> &str {
    match scheme_name {
        "dk1" | "lt1" | "dk2" | "lt2" => scheme_name,
        "tx1" => "dk1",
        "tx2" => "dk2",
        "bg1" => "lt1",
        "bg2" => "lt2",
        other => other,
    }
}

/// HTML auto spacing (`w:beforeAutospacing` / `w:afterAutospacing`) on each
/// side: an inline value overrides the style's, so an inline "0" switches it off.
pub(in crate::docx) fn autospacing(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
    defaults: &StyleDefaults,
) -> (bool, bool) {
    // Word sums the written 5pt after and 5pt before of auto-spaced list
    // items under doNotUseHTMLParagraphAutoSpacing (23.76pt steps for a 13.8
    // line), where the flags alone would give 14pt and drop it inside a list.
    if defaults.no_html_auto_spacing {
        return (false, false);
    }
    let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
    let side = |attr: &str, style_val: Option<bool>, default: bool| {
        inline_spacing
            .and_then(|n| n.attribute((WML_NS, attr)).map(parse_on_off))
            .or(style_val)
            .unwrap_or(default)
    };
    (
        side(
            "beforeAutospacing",
            para_style.and_then(|s| s.space_before_autospacing),
            defaults.before_autospacing,
        ),
        side(
            "afterAutospacing",
            para_style.and_then(|s| s.space_after_autospacing),
            defaults.after_autospacing,
        ),
    )
}

/// Word's auto spacing: 14pt whatever the font, replacing w:before/w:after
/// (Normal (Web) paragraphs step 13.8 + 14 where their twips say 5pt).
const AUTO_SPACING: f32 = 14.0;

pub(in crate::docx) fn parse_paragraph_spacing(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
    defaults: &StyleDefaults,
) -> (Option<f32>, Option<f32>, Option<LineSpacing>) {
    let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
    let (before_auto, after_auto) = autospacing(ppr, para_style, defaults);
    let space_before = if before_auto {
        Some(AUTO_SPACING)
    } else {
        inline_spacing
            .and_then(|n| twips_attr(n, "before"))
            .or_else(|| para_style.and_then(|s| s.space_before))
    };
    let space_after = if after_auto {
        Some(AUTO_SPACING)
    } else {
        inline_spacing
            .and_then(|n| twips_attr(n, "after"))
            .or_else(|| para_style.and_then(|s| s.space_after))
    };

    let line_spacing = inline_spacing
        .and_then(parse_line_spacing)
        .or_else(|| para_style.and_then(|s| s.line_spacing));
    (space_before, space_after, line_spacing)
}

/// Convert a character-unit indent attribute (hundredths of character width) to points.
fn chars_to_pts(ind: roxmltree::Node, attr: &str, char_width: f32) -> Option<f32> {
    ind.attribute((WML_NS, attr))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|hundredths| hundredths / 100.0 * char_width)
}

/// A `w:ind`'s character-unit attributes, in characters (the attributes are
/// hundredths). An explicit zero switches the value off, also over a style's
/// (taiwanese_education_fraud_ruling: `leftChars="0" left="0"` under a
/// List Paragraph style's 200 starts at the margin); Word then uses the twips.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub(super) struct IndentChars {
    pub(super) left: Option<f32>,
    pub(super) right: Option<f32>,
    pub(super) hanging: Option<f32>,
    pub(super) first_line: Option<f32>,
}

pub(super) fn extract_indent_chars(ind: roxmltree::Node) -> IndentChars {
    let chars = |attrs: &[&str]| {
        attrs
            .iter()
            .find_map(|a| ind.attribute((WML_NS, *a)))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|hundredths| hundredths / 100.0)
    };
    IndentChars {
        left: chars(&["startChars", "leftChars"]),
        right: chars(&["endChars", "rightChars"]),
        hanging: chars(&["hangingChars"]),
        first_line: chars(&["firstLineChars"]),
    }
}

/// Word's character-unit indents, measured (online export probes 2026-10-28,
/// issue #143): a non-zero `*Chars` value wins over the twips next to it, even
/// a style's over a paragraph's direct `w:hanging`; `char_widths` is one
/// character for left/right and for first line/hanging (`indent_char_widths`);
/// and with `leftChars` or `hangingChars` in play the left
/// value is the first line's position and the hanging (chars or twips) is
/// added for the body, so left=480 hanging=200 hangingChars=200 and
/// leftChars=200 hangingChars=200 start at 0/24 and 24/48pt. `firstLineChars`
/// alone keeps the twips left.
pub(super) fn apply_indent_chars(
    ind: &mut (f32, f32, f32, f32),
    chars: IndentChars,
    char_widths: (f32, f32),
) {
    let (side_width, char_width) = char_widths;
    let on = |v: Option<f32>| v.filter(|v| *v != 0.0);
    let chars = IndentChars {
        left: on(chars.left),
        right: on(chars.right),
        hanging: on(chars.hanging),
        first_line: on(chars.first_line),
    };
    let (left, right, hanging, first) = ind;
    if let Some(r) = chars.right {
        *right = r * side_width;
    }
    let first_chars = chars.first_line.map(|f| f * char_width);
    if chars.left.is_some() || chars.hanging.is_some() {
        let base = chars.left.map_or(0.0, |l| l * side_width);
        let hang = match (chars.hanging, first_chars) {
            (Some(h), _) => h * char_width,
            (None, Some(_)) => 0.0,
            (None, None) => *hanging,
        };
        *left = base + hang;
        *hanging = hang;
        *first = first_chars.unwrap_or(0.0);
    } else if let Some(f) = first_chars {
        *first = f;
        *hanging = 0.0;
    }
}

pub(super) fn extract_indents(
    ind: roxmltree::Node,
    char_width: Option<f32>,
) -> (Option<f32>, Option<f32>, Option<f32>, Option<f32>) {
    // One character is the font size, Latin or not (Word probe: a style's
    // firstLineChars=200 indents 24pt at 12pt, 20pt at 10pt). Styles and
    // paragraphs pass no width and resolve their `*Chars` values with the
    // first run's size instead (`apply_indent_chars`); the document defaults
    // keep this fallback.
    let side = |chars: &str, twips: &[&str]| {
        twips
            .iter()
            .find_map(|t| twips_attr(ind, t))
            .or_else(|| char_width.and_then(|cw| chars_to_pts(ind, chars, cw)))
    };
    (
        side("leftChars", &["start", "left"]),
        side("rightChars", &["end", "right"]),
        side("hangingChars", &["hanging"]),
        side("firstLineChars", &["firstLine"]),
    )
}

// Final view omits an empty paragraph whose paragraph mark was deleted.
// Do not discard surviving runs, bookmarks, or section boundaries: deleting
// the mark alone merges text with the next paragraph rather than deleting it.
fn is_fully_deleted_paragraph(node: roxmltree::Node) -> bool {
    if !node.has_tag_name((WML_NS, "p")) {
        return false;
    }
    let Some(ppr) = wml(node, "pPr") else {
        return false;
    };
    if wml(ppr, "sectPr").is_some() || wml(ppr, "rPr").and_then(|rpr| wml(rpr, "del")).is_none() {
        return false;
    }
    node.children().filter(|n| n.is_element()).all(|n| {
        n.tag_name().namespace() == Some(WML_NS)
            && matches!(n.tag_name().name(), "pPr" | "del" | "moveFrom" | "proofErr")
    })
}

pub(super) fn collect_block_nodes<'a>(
    parent: roxmltree::Node<'a, 'a>,
) -> Vec<roxmltree::Node<'a, 'a>> {
    let mut nodes = Vec::new();
    for child in parent.children() {
        if is_fully_deleted_paragraph(child) {
            continue;
        }
        if child.has_tag_name((WML_NS, "sdt")) {
            if let Some(content) = wml(child, "sdtContent") {
                nodes.extend(collect_block_nodes(content));
            }
        } else if child.has_tag_name((WML_NS, "customXml")) {
            // w:customXml is a transparent wrapper (block/row/cell): its children
            // ARE the content. Same descent drives tbl→tr and tr→tc unwrapping.
            nodes.extend(collect_block_nodes(child));
        } else if child.has_tag_name((MC_NS_TOP, "AlternateContent")) {
            // mc:AlternateContent wraps block-level content in mc:Choice/mc:Fallback.
            // Use mc:Fallback for compatibility (it avoids newer namespace requirements).
            let fallback = child
                .children()
                .find(|n| n.has_tag_name((MC_NS_TOP, "Fallback")));
            if let Some(fb) = fallback {
                nodes.extend(collect_block_nodes(fb));
            }
        } else {
            nodes.push(child);
        }
    }
    nodes
}

pub(super) fn read_zip_text<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    name: &str,
) -> Option<String> {
    let mut content = String::new();
    zip.by_name(name).ok()?.read_to_string(&mut content).ok()?;
    Some(strict_to_transitional(content))
}

/// Strict OOXML (ISO/IEC 29500 Strict) names its namespaces and relationship
/// types `http://purl.oclc.org/ooxml/<area>/<name>` where Transitional uses
/// `http://schemas.openxmlformats.org/<area>/2006/<name>`. Rewriting them on
/// load lets one parser read both; the package namespaces are shared.
fn strict_to_transitional(xml: String) -> String {
    if !xml.contains("http://purl.oclc.org/ooxml/") {
        return xml;
    }
    ["wordprocessingml", "drawingml", "officeDocument"]
        .iter()
        .fold(xml, |acc, area| {
            acc.replace(
                &format!("http://purl.oclc.org/ooxml/{area}/"),
                &format!("http://schemas.openxmlformats.org/{area}/2006/"),
            )
        })
}

pub(super) fn read_zip_bytes<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    name: &str,
) -> Option<Vec<u8>> {
    let mut data = Vec::new();
    zip.by_name(name).ok()?.read_to_end(&mut data).ok()?;
    Some(data)
}

/// Zip path of a relationship target: absolute (`/word/media/x.png`) as is,
/// otherwise relative to `word/`.
pub(super) fn part_path(target: &str) -> String {
    match target.strip_prefix('/') {
        Some(absolute) => absolute.to_string(),
        None => format!("word/{target}"),
    }
}

mod relationships {
    use std::collections::HashMap;
    use std::io::Read;

    use super::read_zip_text;

    fn parse_rels_xml(xml_content: &str) -> HashMap<String, String> {
        let Ok(xml) = roxmltree::Document::parse(xml_content) else {
            return HashMap::new();
        };
        xml.root_element()
            .children()
            .filter(|n| n.tag_name().name() == "Relationship")
            .filter_map(|n| {
                Some((
                    n.attribute("Id")?.to_string(),
                    n.attribute("Target")?.to_string(),
                ))
            })
            .collect()
    }

    pub(in crate::docx) fn parse_relationships<R: Read + std::io::Seek>(
        zip: &mut zip::ZipArchive<R>,
    ) -> HashMap<String, String> {
        let Some(xml_content) = read_zip_text(zip, "word/_rels/document.xml.rels") else {
            return HashMap::new();
        };
        parse_rels_xml(&xml_content)
    }

    pub(in crate::docx) fn parse_part_relationships<R: Read + std::io::Seek>(
        zip: &mut zip::ZipArchive<R>,
        part_path: &str,
    ) -> HashMap<String, String> {
        let (dir, file) = match part_path.rsplit_once('/') {
            Some((d, f)) => (d, f),
            None => ("", part_path),
        };
        let rels_path = if dir.is_empty() {
            format!("_rels/{}.rels", file)
        } else {
            format!("{}/_rels/{}.rels", dir, file)
        };
        let Some(xml_content) = read_zip_text(zip, &rels_path) else {
            return HashMap::new();
        };
        parse_rels_xml(&xml_content)
    }
}

pub fn parse(path: &std::path::Path) -> Result<Document, Error> {
    let file = std::fs::File::open(path).map_err(|e| match e.kind() {
        std::io::ErrorKind::NotFound | std::io::ErrorKind::PermissionDenied => Error::Io(
            std::io::Error::new(e.kind(), format!("{}: {}", e, path.display())),
        ),
        _ => Error::Io(e),
    })?;

    let mut zip = zip::ZipArchive::new(file)
        .map_err(|_| Error::InvalidDocx("file is not a ZIP archive".into()))?;

    parse_zip(&mut zip)
}

pub fn parse_bytes(bytes: &[u8]) -> Result<Document, Error> {
    let cursor = std::io::Cursor::new(bytes);
    let mut zip = zip::ZipArchive::new(cursor)
        .map_err(|_| Error::InvalidDocx("data is not a valid ZIP/DOCX archive".into()))?;

    parse_zip(&mut zip)
}

fn parse_core_props<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> (
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
) {
    let Some(xml_content) = read_zip_text(zip, "docProps/core.xml") else {
        return (None, None, None, None);
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return (None, None, None, None);
    };

    const DC_NS: &str = "http://purl.org/dc/elements/1.1/";
    const CP_NS: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    let root = xml.root_element();
    let prop = |ns, name| {
        find_children(root, name, ns)
            .filter_map(|n| n.text())
            .last()
            .map(str::to_string)
    };
    (
        prop(DC_NS, "title"),
        prop(DC_NS, "creator"),
        prop(DC_NS, "subject"),
        prop(CP_NS, "keywords"),
    )
}

fn parse_zip<R: Read + std::io::Seek>(zip: &mut zip::ZipArchive<R>) -> Result<Document, Error> {
    let settings = parse_settings(zip);
    let theme = parse_theme(
        zip,
        settings.east_asia_lang.as_deref(),
        settings.bidi_lang.as_deref(),
    );
    let mut styles = parse_styles(zip, &theme, settings.styles_from_normal_template);
    styles.defaults.no_html_auto_spacing = settings.no_html_auto_spacing;
    styles.defaults.no_indent_as_num_tab = settings.no_indent_as_num_tab;
    let numbering = parse_numbering(zip);
    let rels = parse_relationships(zip);
    let ft = parse_font_table(zip);
    let (embedded_fonts, font_table) = (ft.embedded_fonts, ft.font_table);
    let (footnotes, footnote_separator) = parse_footnotes(zip, &styles, &theme, &numbering);
    let (endnotes, endnote_separator) = parse_endnotes(zip, &styles, &theme, &numbering);
    let mut comments = comments::parse_comments(zip);
    let (title, author, subject, keywords) = parse_core_props(zip);

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")
        .map_err(|_| Error::InvalidDocx("missing word/document.xml (is this a DOCX file?)".into()))?
        .read_to_string(&mut xml_content)?;
    let xml_content = strict_to_transitional(xml_content);

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = wml(root, "body").ok_or_else(|| Error::InvalidDocx("Missing w:body".into()))?;

    let default_line_pitch = styles.defaults.font_size * 1.2;

    let comment_ranges = comment_ranges(root);
    // Word numbers balloons in the order of their reference marks.
    let mut next_index = 0;
    for id in &comment_ranges.reference_order {
        if let Some(c) = comments.get_mut(id) {
            next_index += 1;
            c.display_index = next_index;
        }
    }
    // A sectPr closes its section, so the grid its paragraphs' chars indents
    // measure by is read ahead (tokyo_welfare: firstLineChars=1000 is 2217
    // twips at 11pt under charSpace 339; Word probes 2026-10-28: 13pt per
    // character under charSpace 4096 at 12pt, 12 under a lines-only grid).
    let grid_deltas: Vec<f32> = body
        .descendants()
        .filter(|n| n.has_tag_name((WML_NS, "sectPr")))
        .map(sections::grid_char_delta)
        .collect();
    let mut sect_idx = 0;
    let mut ctx = ParseContext {
        styles: &styles,
        theme: &theme,
        rels: &rels,
        zip,
        numbering: &numbering,
        comments: comment_ranges,
        fields: Default::default(),
        grid_char_delta: grid_deltas.first().copied().unwrap_or(0.0),
    };

    let mut sections: Vec<Section> = Vec::new();
    let mut blocks = Vec::new();
    let mut lists = numbering::ListCounters::default();
    // numId of the previous body block when it was a list paragraph.
    let mut prev_list_num_id: Option<String> = None;

    for node in collect_block_nodes(body) {
        if node.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "tbl" => {
                ctx.fields = Default::default();
                let table = parse_table_node(node, &mut ctx, &mut lists);
                blocks.push(Block::Table(table));
                prev_list_num_id = None;
            }
            "p" => {
                let ppr = wml(node, "pPr");
                let para_style_id = ppr
                    .and_then(|ppr| wml_attr(ppr, "pStyle"))
                    .unwrap_or(&styles.default_paragraph_style_id);
                let para_style = styles.paragraph_styles.get(para_style_id);

                let opts = paragraph::ParagraphOptions {
                    resolve_bookmarks: true,
                    resolve_outline_level: true,
                    resolve_drawings: true,
                    collect_extra_textboxes: true,
                    style_num_id: para_style.and_then(|s| s.num_id.clone()),
                    style_num_ilvl: para_style.and_then(|s| s.num_ilvl),
                };
                let mut para = paragraph::build_paragraph(node, &mut ctx, &mut lists, &opts);

                // HTML auto spacing drops between items of one list:
                // russian_university's auto-spaced list items step a plain
                // line apart.
                let num_id = ppr
                    .and_then(|ppr| wml(ppr, "numPr"))
                    .and_then(|np| wml_attr(np, "numId"))
                    .map(str::to_string)
                    .or_else(|| para_style.and_then(|s| s.num_id.clone()))
                    .filter(|id| id != "0");
                if num_id.is_some() && num_id == prev_list_num_id {
                    if para.space_before_auto {
                        para.space_before = 0.0;
                    }
                    if let Some(Block::Paragraph(prev)) = blocks.last_mut()
                        && prev.space_after_auto
                    {
                        prev.space_after = 0.0;
                    }
                }
                prev_list_num_id = num_id;

                let continuation = paragraph::split_at_page_break(&mut para);
                blocks.push(Block::Paragraph(para));
                blocks.extend(continuation.map(Block::Paragraph));
                paragraph::merge_field_code_paragraphs(&mut blocks);

                // Mid-document section break: sectPr inside pPr ends the current section
                if let Some(sect_node) = ppr.and_then(|ppr| wml(ppr, "sectPr")) {
                    if let Some(Block::Paragraph(last_para)) = blocks.last_mut() {
                        last_para.is_section_break = true;
                        // A page break ending the section's last paragraph is dropped:
                        // the section break alone decides (online export probes,
                        // 2026-10-03: no new page before a continuous section, one
                        // before a next-page section; stem_partnerships' cover).
                        last_para.page_break_after = false;
                    }
                    let props = parse_section_properties(
                        sect_node,
                        &mut ctx,
                        default_line_pitch,
                        settings.gutter_at_top,
                    );
                    sections.push(Section {
                        properties: props,
                        blocks: std::mem::take(&mut blocks),
                    });
                    sect_idx += 1;
                    ctx.grid_char_delta = grid_deltas.get(sect_idx).copied().unwrap_or(0.0);
                }
            }
            "altChunk" => {
                if let Some(id) = node.attribute((REL_NS, "id")) {
                    blocks.extend(alt_chunk::parse_alt_chunk(id, ctx.rels, ctx.zip));
                    prev_list_num_id = None;
                }
            }
            _ => {}
        }
    }

    // Word always lays out an implicit final paragraph mark after a
    // document-ending table; when it doesn't fit below the table it spills
    // onto a fresh (blank) page. Model it as a real empty paragraph.
    if matches!(blocks.last(), Some(Block::Table(_))) {
        blocks.push(Block::Paragraph(crate::model::Paragraph::default()));
    }

    // Final section: body-level sectPr
    let final_props = if let Some(sect_node) = wml(body, "sectPr") {
        parse_section_properties(
            sect_node,
            &mut ctx,
            default_line_pitch,
            settings.gutter_at_top,
        )
    } else {
        SectionProperties {
            page_width: 612.0,
            page_height: 792.0,
            margin_top: 72.0,
            margin_bottom: 72.0,
            margin_top_fixed: false,
            margin_bottom_fixed: false,
            margin_left: 72.0,
            margin_right: 72.0,
            header_margin: 36.0,
            footer_margin: 36.0,
            header_default: None,
            header_first: None,
            header_even: None,
            footer_default: None,
            footer_first: None,
            footer_even: None,
            different_first_page: false,
            line_pitch: default_line_pitch,
            grid_type: DocGridType::Default,
            break_type: SectionBreakType::NextPage,
            columns: None,
            page_num_start: None,
            page_num_format: None,
            page_borders: None,
            vertical_align: crate::model::PageVerticalAlign::Top,
            line_numbering: None,
            footnote_num_fmt: None,
            endnote_num_fmt: None,
        }
    };
    sections.push(Section {
        properties: final_props,
        blocks,
    });
    expand_toc_fields(&mut sections, &mut ctx);
    // HTML auto spacing never opens the document, from a w:p or an altChunk
    // alike: case42's first Normal (Web) paragraph sits at the top margin.
    if let Some(Block::Paragraph(first)) = sections[0].blocks.first_mut()
        && first.space_before_auto
    {
        first.space_before = 0.0;
    }
    // hideMark never hides the document's first paragraph mark: a table
    // opening the body keeps a line for an empty hideMark first cell, the
    // same row collapses after a page break or a paragraph (Word probes).
    // Nested layout tables put that mark in the innermost first cell.
    let mut first = sections[0].blocks.first_mut();
    while let Some(Block::Table(table)) = first {
        let Some(cell) = table.rows.first_mut().and_then(|r| r.cells.first_mut()) else {
            break;
        };
        if cell.content.len() == 1 {
            cell.hide_mark = false;
        }
        first = cell.content.first_mut();
    }
    // A continuous break can't change the paper mid-page: Word starts a new
    // page when the size or orientation changes (transition_to_work's
    // landscape Annexure B1, whose page would otherwise take the portrait
    // text above it off the top).
    for i in 1..sections.len() {
        let prev = &sections[i - 1].properties;
        let size = (prev.page_width, prev.page_height);
        let sp = &mut sections[i].properties;
        if sp.break_type == SectionBreakType::Continuous && (sp.page_width, sp.page_height) != size
        {
            sp.break_type = SectionBreakType::NextPage;
        }
    }

    // Word's PDF export scales the body content via a wrapping `cm` operator
    // when comments are present, so glyphs render at ~76% and a column opens
    // up on the right for the comment pane. We apply the same scaling at
    // content-stream assembly time; the body keeps its natural layout here.

    // A textbox sized as a share of the text width (pctWidth) takes it from
    // its section: 40% of a 451pt column draws a 180.25pt box in Word, not
    // the 187pt extent saved beside it.
    for section in &mut sections {
        let text_width = section.properties.text_width();
        for block in &mut section.blocks {
            if let Block::Paragraph(p) = block {
                for tb in &mut p.textboxes {
                    if let Some(f) = tb.width_pct_margin {
                        tb.width_pt = f * text_width;
                    }
                }
            }
        }
    }

    // Word's print export sets a balloon in the Balloon Text style (Times New
    // Roman 9pt when the document has none; 70 references): the label in its
    // face, bold, the comment text in its own face, both at its size.
    let balloon_style = styles.paragraph_styles.get("BalloonText");
    let comment_label_font = balloon_style
        .and_then(|s| s.font_name.clone())
        .unwrap_or_else(|| "Times New Roman".into());
    let comment_font_size = balloon_style.and_then(|s| s.font_size).unwrap_or(9.0);
    // With no CommentText style the text takes Normal's face, as the built-in
    // CommentText would: case63's Aptos Normal over a Cambria theme draws its
    // comments in Aptos (whose "✓" Cambria lacks).
    let comment_body_font = ["CommentText", styles.default_paragraph_style_id.as_str()]
        .iter()
        .find_map(|id| styles.paragraph_styles.get(*id)?.font_name.clone())
        .unwrap_or_else(|| styles.defaults.font_name.clone());

    let mut doc = Document {
        sections,
        line_spacing: styles.defaults.line_spacing,
        embedded_fonts,
        footnotes,
        footnote_separator,
        endnotes,
        endnote_separator,
        comments,
        comment_label_font,
        comment_body_font,
        comment_font_size,
        font_table,
        even_and_odd_headers: settings.even_and_odd_headers,
        mirror_margins: settings.mirror_margins,
        default_tab_stop: settings.default_tab_stop,
        style_id_to_name: styles.style_id_to_name,
        theme_minor_font: theme.minor.clone(),
        title,
        author,
        subject,
        keywords,
        default_lang: styles.defaults.lang.clone().or(settings.default_lang),
        compress_punctuation: settings.compress_punctuation,
        compat_mode: settings.compat_mode,
        open_type_features: settings.open_type_features,
        east_asian_lang: styles
            .defaults
            .lang_east_asia
            .as_deref()
            .or(settings.east_asia_lang.as_deref())
            .is_some_and(|l| ["zh", "ja", "ko"].iter().any(|p| l.starts_with(p)))
            && styles
                .defaults
                .east_asia_font
                .as_deref()
                .is_some_and(crate::fonts::is_cjk_family),
        do_not_expand_shift_return: settings.do_not_expand_shift_return,
        wp_justification: settings.wp_justification,
        additive_paragraph_spacing: settings.no_html_auto_spacing,
        adjust_line_height_in_table: settings.adjust_line_height_in_table,
        math: settings.math,
    };
    // Captions get their numbers before a STYLEREF quotes them.
    resolve_seq_fields(&mut doc);
    resolve_body_stylerefs(&mut doc);
    link_textbox_stories(&mut doc);
    Ok(doc)
}

/// Number the body's SEQ fields in document order, one counter per name
/// (issue #145: a stale "old" cache shows 1). ponytail: `\s` (restart under
/// a heading level) is not read, and text-box captions are not counted.
fn resolve_seq_fields(doc: &mut Document) {
    use crate::model::{FieldCode, SeqStep};
    let mut paras = Vec::new();
    for section in &mut doc.sections {
        body_paragraphs(&mut section.blocks, &mut paras);
    }
    let mut counters: HashMap<String, u32> = HashMap::new();
    for run in paras.iter_mut().flat_map(|p| p.runs.iter_mut()) {
        let Some(FieldCode::Seq {
            name,
            format,
            step,
            hidden,
        }) = &run.field_code
        else {
            continue;
        };
        let counter = counters.entry(name.clone()).or_insert(0);
        match step {
            SeqStep::Next => *counter += 1,
            // A repeat before any number shows the first.
            SeqStep::Repeat => *counter = (*counter).max(1),
            SeqStep::Restart(n) => *counter = *n,
        }
        run.text = if *hidden {
            String::new()
        } else {
            numbering::format_number(*counter, format.unwrap_or("decimal"))
        };
    }
}

/// What a TOC field collects: `\o "a-b"` (and `\u`) the heading outline
/// levels, `\t "Style,level,…"` paragraphs of a style, `\c "Name"` paragraphs
/// holding a SEQ of that name; `\n` drops the page numbers. Without any of
/// them Word takes headings 1–3.
#[derive(Debug, PartialEq)]
struct TocSpec {
    outline: Option<std::ops::RangeInclusive<u8>>,
    styles: Vec<(String, u8)>,
    seq: Option<String>,
    page_numbers: bool,
}

impl TocSpec {
    fn parse(instr: &str) -> TocSpec {
        let has = |s| runs::has_switch(instr, s);
        let mut outline = toc_switch_arg(instr, "\\o")
            .and_then(|r| r.split_once('-'))
            .and_then(|(a, b)| Some(a.trim().parse().ok()?..=b.trim().parse().ok()?))
            .or_else(|| (has("\\o") || has("\\u")).then_some(1..=9));
        let styles: Vec<(String, u8)> = toc_switch_arg(instr, "\\t")
            .map(|list| {
                let items: Vec<&str> = list.split([',', ';']).map(str::trim).collect();
                items
                    .chunks(2)
                    .filter_map(|c| Some((c.first()?.to_string(), c.get(1)?.parse().ok()?)))
                    .collect()
            })
            .unwrap_or_default();
        let seq = toc_switch_arg(instr, "\\c").map(str::to_string);
        if outline.is_none() && styles.is_empty() && seq.is_none() {
            outline = Some(1..=3);
        }
        TocSpec {
            outline,
            styles,
            seq,
            page_numbers: !has("\\n"),
        }
    }

    /// The TOC level `p` enters at, if any.
    fn level_of(&self, p: &crate::model::Paragraph, names: &HashMap<String, String>) -> Option<u8> {
        use crate::model::FieldCode;
        if let (Some(range), Some(level)) = (&self.outline, p.outline_level)
            && range.contains(&(level + 1))
        {
            return Some(level + 1);
        }
        if let Some(id) = &p.style_id
            && let Some((_, level)) = self.styles.iter().find(|(name, _)| {
                id.eq_ignore_ascii_case(name)
                    || names.get(id).is_some_and(|n| n.eq_ignore_ascii_case(name))
            })
        {
            return Some(*level);
        }
        let seq = self.seq.as_deref()?;
        p.runs
            .iter()
            .any(|r| matches!(&r.field_code, Some(FieldCode::Seq { name, .. }) if name.eq_ignore_ascii_case(seq)))
            .then_some(1)
    }
}

/// The argument after a TOC switch (`\o "1-3"`), if the switch is present.
fn toc_switch_arg<'a>(instr: &'a str, switch: &str) -> Option<&'a str> {
    let lower = instr.to_ascii_lowercase();
    let after = lower
        .match_indices(&switch.to_ascii_lowercase())
        .map(|(start, m)| (start, start + m.len()))
        .find(|&(start, end)| {
            (start == 0 || lower.as_bytes()[start - 1].is_ascii_whitespace())
                && lower[end..].chars().next().is_none_or(char::is_whitespace)
        })
        .map(|(_, end)| end)?;
    runs::quoted_or_word(&instr[after..])
}

fn xml_text(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('"', "&quot;")
}

/// Regenerate the TOC fields whose cache is a single paragraph (issue #145's
/// "Update field in Word."), as Word's field update does: one "toc N"
/// paragraph per entry with its text, a dot-leader right tab 10 twips inside
/// the margin and a PAGEREF to the entry's paragraph (a bookmark added here),
/// then the field's own paragraph left empty. Later TOCs are expanded first
/// so the earlier ones' block indices hold. ponytail: `\p` separators, `\b`
/// scopes and TC fields are ignored; without a "toc N" style a level N entry
/// is indented 11pt × (N − 1).
fn expand_toc_fields<R: Read + std::io::Seek>(
    sections: &mut [Section],
    ctx: &mut ParseContext<'_, R>,
) {
    let mut hosts = Vec::new();
    for (s, section) in sections.iter().enumerate() {
        for (b, block) in section.blocks.iter().enumerate() {
            if matches!(block, Block::Paragraph(p) if p.toc_field.is_some()) {
                hosts.push((s, b));
            }
        }
    }
    let names = &ctx.styles.style_id_to_name;
    let mut bookmarks = 0;
    for (s, b) in hosts.into_iter().rev() {
        let Block::Paragraph(host) = &mut sections[s].blocks[b] else {
            continue;
        };
        let spec = TocSpec::parse(&host.toc_field.take().unwrap_or_default());
        host.runs.clear();
        host.starts_toc_field = false;

        // (level, bookmark, list label, text) of each entry, in document order.
        let mut entries = Vec::new();
        let mut paras = Vec::new();
        for section in sections.iter_mut() {
            body_paragraphs(&mut section.blocks, &mut paras);
        }
        for p in paras {
            if p.starts_toc_field || p.toc_field.is_some() {
                continue;
            }
            let Some(level) = spec.level_of(p, names) else {
                continue;
            };
            let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
            let bookmark = format!("_Toc145_{bookmarks}");
            bookmarks += 1;
            p.bookmarks.push(bookmark.clone());
            let label = p.list_label.clone();
            entries.push((level, bookmark, label, text.trim().to_string()));
        }

        let sp = &sections[s].properties;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;
        let tab_pos = (text_width * 20.0).round() as i32 - 10;
        let toc_style = |level: u8| {
            let name = format!("toc {level}");
            names.iter().find(|(_, n)| n.eq_ignore_ascii_case(&name)).map(|(id, _)| id.clone())
        };
        let mut xmls: Vec<String> = entries
            .iter()
            .map(|(level, bookmark, label, text)| {
                let mut ppr = match toc_style(*level) {
                    Some(id) => format!(r#"<w:pStyle w:val="{}"/>"#, xml_text(&id)),
                    None if *level > 1 => format!(r#"<w:ind w:left="{}"/>"#, 220 * (*level as i32 - 1)),
                    None => String::new(),
                };
                let mut body = String::new();
                if !label.is_empty() {
                    body += &format!(r#"<w:r><w:t xml:space="preserve">{}</w:t></w:r><w:r><w:tab/></w:r>"#, xml_text(label));
                }
                body += &format!(r#"<w:r><w:t xml:space="preserve">{}</w:t></w:r>"#, xml_text(text));
                if spec.page_numbers {
                    ppr += &format!(r#"<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{tab_pos}"/></w:tabs>"#);
                    body += &format!(
                        r#"<w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF {bookmark} \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
                    );
                }
                format!(r#"<w:p xmlns:w="{WML_NS}"><w:pPr>{ppr}</w:pPr>{body}</w:p>"#)
            })
            .collect();
        if xmls.is_empty() {
            xmls.push(format!(
                r#"<w:p xmlns:w="{WML_NS}"><w:r><w:t>No table of contents entries found.</w:t></w:r></w:p>"#
            ));
        }
        let opts = paragraph::ParagraphOptions {
            resolve_bookmarks: false,
            resolve_outline_level: false,
            resolve_drawings: false,
            collect_extra_textboxes: false,
            style_num_id: None,
            style_num_ilvl: None,
        };
        let mut generated: Vec<Block> = xmls
            .iter()
            .filter_map(|xml| roxmltree::Document::parse(xml).ok())
            .map(|doc| {
                ctx.fields = Default::default();
                let mut lists = numbering::ListCounters::default();
                Block::Paragraph(paragraph::build_paragraph(doc.root_element(), ctx, &mut lists, &opts))
            })
            .collect();
        if let Some(Block::Paragraph(first)) = generated.first_mut() {
            first.starts_toc_field = true;
        }
        sections[s].blocks.splice(b..b, generated);
    }
}

/// Every body paragraph, table cells included, in document order.
fn body_paragraphs<'a>(blocks: &'a mut [Block], out: &mut Vec<&'a mut crate::model::Paragraph>) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => out.push(p),
            Block::Table(t) => {
                for cell in t.rows.iter_mut().flat_map(|r| r.cells.iter_mut()) {
                    body_paragraphs(&mut cell.content, out);
                }
            }
        }
    }
}

/// Each box of a linked chain shares the story its first box holds and
/// learns the text areas of the boxes before it, from which the renderer works
/// out the part it shows. ponytail: body text boxes only, not headers or groups.
fn link_textbox_stories(doc: &mut Document) {
    let mut paras = Vec::new();
    for section in &mut doc.sections {
        body_paragraphs(&mut section.blocks, &mut paras);
    }
    let mut boxes: Vec<(u32, u32, &mut crate::model::Textbox)> = paras
        .into_iter()
        .flat_map(|p| p.textboxes.iter_mut())
        .filter_map(|tb| Some((tb.link.as_ref()?.id, tb.link.as_ref()?.seq, tb)))
        .collect();
    boxes.sort_by_key(|&(id, seq, _)| (id, seq));
    for chain in boxes.chunk_by_mut(|a, b| a.0 == b.0) {
        let story = std::sync::Arc::new(std::mem::take(&mut chain[0].2.paragraphs));
        let mut before = Vec::new();
        for (_, _, tb) in chain.iter_mut() {
            let area = tb.text_area(tb.height_pt);
            if let Some(link) = &mut tb.link {
                link.story = story.clone();
                link.before = before.clone();
                link.continues = true;
            }
            before.push(area);
        }
        if let Some(link) = chain.last_mut().and_then(|(_, _, tb)| tb.link.as_mut()) {
            link.continues = false;
        }
    }
}

/// Body STYLEREF fields show the nearest paragraph in the named style before
/// the field, else the first one after it (§17.16.5.59). Word updates them on
/// export, so the cached text is stale (issue #142: the body copy of a header's
/// STYLEREF kept "Heading pending"). Headers resolve theirs per page instead.
fn resolve_body_stylerefs(doc: &mut Document) {
    use crate::model::{FieldCode, styleref_key};
    let names = std::mem::take(&mut doc.style_id_to_name);
    let mut paras = Vec::new();
    for section in &mut doc.sections {
        body_paragraphs(&mut section.blocks, &mut paras);
    }
    // Per paragraph, the keys it defines (paragraph style and character
    // styles, by id and by name) and their values.
    let defined: Vec<Vec<(String, String)>> = paras
        .iter()
        .map(|p| {
            let mut defs = Vec::new();
            let mut define = |id: &str, text: String| {
                for name in std::iter::once(id).chain(names.get(id).map(String::as_str)) {
                    for (form, value) in p.styleref_values(&text) {
                        defs.push((styleref_key(name, form), value.to_string()));
                    }
                }
            };
            if let Some(sid) = &p.style_id {
                let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                if !text.is_empty() {
                    define(sid, text);
                }
            }
            for group in p.runs.chunk_by(|a, b| a.char_style_id == b.char_style_id) {
                if let Some(csid) = &group[0].char_style_id {
                    let text: String = group.iter().map(|r| r.text.as_str()).collect();
                    if !text.is_empty() {
                        define(csid, text);
                    }
                }
            }
            defs
        })
        .collect();
    let lookup = |key: &str, at: usize| {
        let before = (0..at).rev();
        let after = at + 1..defined.len();
        before
            .chain(after)
            .find_map(|i| defined[i].iter().find(|(k, _)| k == key).map(|(_, v)| v.clone()))
    };
    for (i, p) in paras.iter_mut().enumerate() {
        for run in &mut p.runs {
            if let Some(FieldCode::StyleRef { name, form }) = run.field_code.clone()
                && let Some(text) = lookup(&styleref_key(&name, form), i)
            {
                run.text = text;
            }
        }
    }
    doc.style_id_to_name = names;
}

#[cfg(test)]
mod tests {
    #[test]
    fn border_styles_parse_and_size_their_bands() {
        use crate::model::BorderStyle;
        let style = super::border_style;
        assert_eq!(style("thinThickMediumGap"), BorderStyle::ThinThick);
        assert_eq!(style("thickThinSmallGap"), BorderStyle::ThickThin);
        assert_eq!(style("dotDash"), BorderStyle::DashDot);
        assert_eq!(style("triple"), BorderStyle::Single);
        let xml = r#"<w:bottom xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="double" w:sz="12"/>"#;
        let b = super::parse_one_border(roxmltree::Document::parse(xml).unwrap().root_element())
            .unwrap();
        assert_eq!(
            (b.style, b.width_pt, b.extent()),
            (BorderStyle::Double, 1.5, 4.5)
        );
    }

    #[test]
    fn toc_spec_reads_its_switches() {
        let spec = super::TocSpec::parse(r#" TOC \o "1-3" \h \z \u "#);
        assert_eq!((spec.outline, spec.page_numbers), (Some(1..=3), true));
        let spec = super::TocSpec::parse(r#" TOC \t "Figure Caption,1,Table Caption,2" \n "#);
        let styles = [("Figure Caption".to_string(), 1), ("Table Caption".to_string(), 2)];
        assert_eq!(spec.styles, styles);
        assert!(!spec.page_numbers && spec.outline.is_none());
        let spec = super::TocSpec::parse(r#" TOC \c "Figure" "#);
        assert_eq!(spec.seq.as_deref(), Some("Figure"));
        assert_eq!(super::TocSpec::parse(" TOC \\h ").outline, Some(1..=3));
    }

    #[test]
    fn twips_take_universal_measures() {
        let doc = roxmltree::Document::parse(
            r#"<w:m xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:a="1440" w:b="10mm" w:c="1in" w:d="-2.5pt"/>"#,
        )
        .unwrap();
        let n = doc.root_element();
        assert_eq!(super::twips_attr(n, "a"), Some(72.0));
        assert!((super::twips_attr(n, "b").unwrap() - 28.346).abs() < 0.01);
        assert_eq!(super::twips_attr(n, "c"), Some(72.0));
        assert_eq!(super::twips_attr(n, "d"), Some(-2.5));
    }

    use super::*;

    #[test]
    fn paragraph_borders_merge_per_side() {
        let sides = |pbdr: &str| {
            let xml = format!(r#"<w:pPr xmlns:w="{WML_NS}"><w:pBdr>{pbdr}</w:pBdr></w:pPr>"#);
            parse_paragraph_borders(roxmltree::Document::parse(&xml).unwrap().root_element())
        };
        let parent = sides(r#"<w:top w:val="single"/><w:left w:val="single"/>"#);
        let child = sides(r#"<w:left w:val="nil"/><w:bottom w:val="single"/>"#);
        let b = child.over(&parent).resolve();
        assert!(b.top.is_some() && b.bottom.is_some());
        assert!(b.left.is_none() && b.right.is_none());
    }

    #[test]
    fn strict_namespaces_read_as_transitional() {
        let strict = r#"<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><w:body/></w:document>"#;
        let xml = strict_to_transitional(strict.to_string());
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert!(wml(doc.root_element(), "body").is_some());
        assert!(
            xml.contains("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        );
    }

    #[test]
    fn frame_y_align_overrides_y_and_inline_stays_in_flow() {
        let parse = |fp: &str| {
            let xml = format!(r#"<w:pPr xmlns:w="{WML_NS}">{fp}</w:pPr>"#);
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_frame_props(Some(doc.root_element()), None)
        };
        let bottom =
            parse(r#"<w:framePr w:w="2659" w:hAnchor="page" w:x="8971" w:yAlign="bottom"/>"#)
                .unwrap();
        assert_eq!(bottom.v_position, VerticalPosition::AlignBottom);
        assert_eq!(bottom.v_relative_from, VRelativeFrom::Margin);
        let y = parse(r#"<w:framePr w:vAnchor="page" w:y="3222"/>"#).unwrap();
        assert_eq!(y.v_position, VerticalPosition::Offset(161.1));
        assert!(parse(r#"<w:framePr w:yAlign="inline"/>"#).is_none());
    }

    // --- Pure math / conversion ---

    #[test]
    fn deleted_paragraph_marks_hide_only_fully_deleted_blocks() {
        let cases = [
            (
                "<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr><w:del><w:r><w:delText>old</w:delText></w:r></w:del></w:p>",
                true,
            ),
            ("<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr></w:p>", true),
            (
                "<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr><w:r><w:t>kept</w:t></w:r></w:p>",
                false,
            ),
            (
                "<w:p><w:pPr><w:rPr><w:del/></w:rPr><w:sectPr/></w:pPr></w:p>",
                false,
            ),
            (
                "<w:p><w:pPr><w:rPr><w:del/></w:rPr></w:pPr><w:bookmarkStart w:id=\"1\"/></w:p>",
                false,
            ),
            (
                "<w:p><w:del><w:r><w:delText>old</w:delText></w:r></w:del></w:p>",
                false,
            ),
            ("<w:p/>", false),
        ];
        for (paragraph, deleted) in cases {
            let xml = format!("<w:body xmlns:w=\"{WML_NS}\">{paragraph}</w:body>");
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let blocks = collect_block_nodes(doc.root_element());
            assert_eq!(blocks.len(), usize::from(!deleted), "{paragraph}");
        }
    }



    #[test]
    fn test_twips_to_pts() {
        assert_eq!(twips_to_pts(0.0), 0.0);
        assert_eq!(twips_to_pts(20.0), 1.0);
        assert_eq!(twips_to_pts(360.0), 18.0);
        assert_eq!(twips_to_pts(1440.0), 72.0); // 1 inch
    }

    #[test]
    fn test_emu_to_pts() {
        assert_eq!(emu_to_pts(0.0), 0.0);
        assert_eq!(emu_to_pts(12700.0), 1.0);
        assert_eq!(emu_to_pts(914400.0), 72.0); // 1 inch
    }

    // --- Color parsing ---

    #[test]
    fn test_parse_hex_color() {
        assert_eq!(parse_hex_color("FF0000"), Some([255, 0, 0]));
        assert_eq!(parse_hex_color("00FF00"), Some([0, 255, 0]));
        assert_eq!(parse_hex_color("0000FF"), Some([0, 0, 255]));
        assert_eq!(parse_hex_color("4472C4"), Some([68, 114, 196]));
        assert_eq!(parse_hex_color("auto"), None);
        assert_eq!(parse_hex_color("red"), None); // wrong length
        assert_eq!(parse_hex_color(""), None);
        assert_eq!(parse_hex_color("ZZZZZZ"), None); // invalid hex
    }

    #[test]
    fn test_parse_text_color() {
        assert_eq!(parse_text_color("auto"), Some([0, 0, 0])); // auto → black
        assert_eq!(parse_text_color("FF0000"), Some([255, 0, 0]));
        assert_eq!(parse_text_color("invalid"), None);
    }

    #[test]
    fn test_parse_run_shd() {
        let ns = WML_NS;

        // No w:shd child → None (inherit)
        let xml = format!(r#"<w:rPr xmlns:w="{ns}"><w:b/></w:rPr>"#);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), None);

        // val="clear" fill="FFFF00" → Some(yellow)
        let xml = format!(
            r#"<w:rPr xmlns:w="{ns}"><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:rPr>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), Some([0xFF, 0xFF, 0x00]));

        // val="nil" → None (not an override: inherit from style)
        let xml = format!(r#"<w:rPr xmlns:w="{ns}"><w:shd w:val="nil"/></w:rPr>"#);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), None);

        // fill="auto" → None (inherit)
        let xml = format!(r#"<w:rPr xmlns:w="{ns}"><w:shd w:val="clear" w:fill="auto"/></w:rPr>"#);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), None);

        // Missing fill → None (inherit)
        let xml = format!(r#"<w:rPr xmlns:w="{ns}"><w:shd w:val="clear"/></w:rPr>"#);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), None);

        // pct25 of auto (black) over CCCCCC
        let xml =
            format!(r#"<w:rPr xmlns:w="{ns}"><w:shd w:val="pct25" w:fill="CCCCCC"/></w:rPr>"#);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), Some([153, 153, 153]));

        // solid paints the pattern color over an auto fill
        let xml = format!(
            r#"<w:rPr xmlns:w="{ns}"><w:shd w:val="solid" w:color="CC99FF" w:fill="auto"/></w:rPr>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), Some([0xCC, 0x99, 0xFF]));
    }

    #[test]
    fn test_highlight_color() {
        assert_eq!(highlight_color("yellow"), Some([255, 255, 0]));
        assert_eq!(highlight_color("red"), Some([255, 0, 0]));
        assert_eq!(highlight_color("blue"), Some([0, 0, 255]));
        assert_eq!(highlight_color("darkGreen"), Some([0, 128, 0]));
        assert_eq!(highlight_color("lightGray"), Some([192, 192, 192]));
        assert_eq!(highlight_color("black"), Some([0, 0, 0]));
        assert_eq!(highlight_color("white"), Some([255, 255, 255]));
        assert_eq!(highlight_color("unknown"), None);
    }

    // --- East Asian character detection ---

    #[test]
    fn test_is_east_asian_char() {
        assert!(is_east_asian_char('漢')); // CJK Unified Ideograph
        assert!(is_east_asian_char('あ')); // Hiragana
        assert!(is_east_asian_char('ア')); // Katakana
        assert!(is_east_asian_char('한')); // Hangul
        assert!(is_east_asian_char('、')); // CJK Symbols
        assert!(is_east_asian_char('\u{FF01}')); // Fullwidth !
        assert!(!is_east_asian_char('A'));
        assert!(!is_east_asian_char('1'));
        assert!(!is_east_asian_char(' '));
        assert!(!is_east_asian_char('é'));
    }

    // --- XML-based utility tests ---

    #[test]
    fn test_wml_bool() {
        let xml = format!(
            r#"<w:pPr xmlns:w="{}">
                <w:b/>
                <w:i w:val="0"/>
                <w:caps w:val="true"/>
                <w:vanish w:val="false"/>
            </w:pPr>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        // Present with no val attribute → true
        assert_eq!(wml_bool(node, "b"), Some(true));
        // val="0" → false
        assert_eq!(wml_bool(node, "i"), Some(false));
        // val="true" → true
        assert_eq!(wml_bool(node, "caps"), Some(true));
        // val="false" → false
        assert_eq!(wml_bool(node, "vanish"), Some(false));
        // Missing element → None
        assert_eq!(wml_bool(node, "strike"), None);
    }

    #[test]
    fn test_wml_attr() {
        let xml = format!(
            r#"<w:pPr xmlns:w="{}">
                <w:jc w:val="center"/>
                <w:sz w:val="24"/>
            </w:pPr>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        assert_eq!(wml_attr(node, "jc"), Some("center"));
        assert_eq!(wml_attr(node, "sz"), Some("24"));
        assert_eq!(wml_attr(node, "missing"), None);
    }

    #[test]
    fn test_twips_attr() {
        let xml = format!(
            r#"<w:ind xmlns:w="{}" w:left="720" w:right="360"/>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        assert_eq!(twips_attr(node, "left"), Some(36.0)); // 720/20=36pt
        assert_eq!(twips_attr(node, "right"), Some(18.0)); // 360/20=18pt
        assert_eq!(twips_attr(node, "hanging"), None);
    }

    // --- Tab stops ---

    #[test]
    fn test_parse_tab_stops() {
        let xml = format!(
            r#"<w:pPr xmlns:w="{}">
                <w:tabs>
                    <w:tab w:val="center" w:pos="4320"/>
                    <w:tab w:val="right" w:pos="8640" w:leader="dot"/>
                    <w:tab w:val="left" w:pos="1440"/>
                </w:tabs>
            </w:pPr>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        let stops = parse_tab_stops_with_clears(node).0;
        assert_eq!(stops.len(), 3);
        // Sorted by position
        assert_eq!(stops[0].position, 72.0); // 1440/20
        assert_eq!(stops[0].alignment, TabAlignment::Left);
        assert_eq!(stops[0].leader, None);
        assert_eq!(stops[1].position, 216.0); // 4320/20
        assert_eq!(stops[1].alignment, TabAlignment::Center);
        assert_eq!(stops[2].position, 432.0); // 8640/20
        assert_eq!(stops[2].alignment, TabAlignment::Right);
        assert_eq!(stops[2].leader, Some('.'));
    }

    #[test]
    fn test_parse_tab_stops_with_clears() {
        let xml = format!(
            r#"<w:pPr xmlns:w="{}">
                <w:tabs>
                    <w:tab w:val="clear" w:pos="720"/>
                    <w:tab w:val="right" w:pos="9360"/>
                </w:tabs>
            </w:pPr>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        let (stops, clears) = parse_tab_stops_with_clears(node);
        assert_eq!(stops.len(), 1);
        assert_eq!(stops[0].alignment, TabAlignment::Right);
        assert_eq!(clears.len(), 1);
        assert_eq!(clears[0], 36.0); // 720/20
    }

    // --- Indents ---

    #[test]
    fn test_extract_indents() {
        let xml = format!(
            r#"<w:ind xmlns:w="{}" w:left="720" w:right="360" w:hanging="360" w:firstLine="0"/>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        let (left, right, hanging, first_line) = extract_indents(node, None);
        assert_eq!(left, Some(36.0)); // 720/20
        assert_eq!(right, Some(18.0)); // 360/20
        assert_eq!(hanging, Some(18.0)); // 360/20
        assert_eq!(first_line, Some(0.0));
    }

    #[test]
    fn test_extract_indents_start_end() {
        // w:start/w:end are the modern equivalents of w:left/w:right
        let xml = format!(
            r#"<w:ind xmlns:w="{}" w:start="1440" w:end="720"/>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        let (left, right, hanging, first_line) = extract_indents(node, None);
        assert_eq!(left, Some(72.0)); // 1440/20
        assert_eq!(right, Some(36.0)); // 720/20
        assert_eq!(hanging, None);
        assert_eq!(first_line, None);
    }

    #[test]
    fn chars_indents_follow_word_probes() {
        // (left, hanging, hangingChars, leftChars, firstLineChars) → (first line x, body x)
        let case = |twips: (f32, f32), chars: IndentChars, cw: f32| {
            let mut ind = (twips.0, 0.0, twips.1, 0.0);
            apply_indent_chars(&mut ind, chars, (cw, cw));
            (ind.0 - ind.2 + ind.3, ind.0)
        };
        let c = |l, h, f| IndentChars { left: l, right: None, hanging: h, first_line: f };
        assert_eq!(case((10.0, 10.0), c(None, Some(2.0), None), 12.0), (0.0, 24.0));
        assert_eq!(case((24.0, 10.0), c(None, Some(2.0), None), 12.0), (0.0, 24.0));
        assert_eq!(case((0.0, 0.0), c(Some(2.0), Some(2.0), None), 12.0), (24.0, 48.0));
        assert_eq!(case((0.0, 20.0), c(Some(2.0), None, None), 12.0), (24.0, 44.0));
        assert_eq!(case((48.0, 0.0), c(None, None, Some(2.0)), 12.0), (72.0, 48.0));
        assert_eq!(case((0.0, 0.0), c(Some(2.0), None, Some(1.0)), 12.0), (36.0, 24.0));
        assert_eq!(case((10.0, 10.0), c(None, Some(2.0), None), 18.0), (0.0, 36.0));
        assert_eq!(case((10.0, 10.0), IndentChars::default(), 12.0), (0.0, 10.0));
        // Issue #139 probes: Normal 10.5pt, 7pt runs, leftChars=400 hangingChars=200
        let mut ind = (0.0, 0.0, 0.0, 0.0);
        apply_indent_chars(&mut ind, c(Some(4.0), Some(2.0), None), (10.5, 7.0));
        assert_eq!((ind.0 - ind.2, ind.0), (42.0, 56.0));
    }

    #[test]
    fn test_extract_indents_chars() {
        let parse = |attrs: &str| {
            let xml = format!(r#"<w:ind xmlns:w="{}" {attrs}/>"#, WML_NS);
            let doc = roxmltree::Document::parse(&xml).unwrap();
            extract_indents(doc.root_element(), Some(12.0))
        };
        assert_eq!(parse(r#"w:leftChars="200" w:left="100""#).0, Some(5.0));
        assert_eq!(parse(r#"w:firstLineChars="200""#).3, Some(24.0));
        assert_eq!(
            parse(r#"w:firstLineChars="0" w:firstLine="720""#).3,
            Some(36.0)
        );
        assert_eq!(parse(r#"w:leftChars="0""#).0, Some(0.0));
    }

    // --- Theme color key resolution ---

    #[test]
    fn test_resolve_theme_color_key() {
        assert_eq!(resolve_theme_color_key("dk1"), "dk1");
        assert_eq!(resolve_theme_color_key("tx1"), "dk1");
        assert_eq!(resolve_theme_color_key("tx2"), "dk2");
        assert_eq!(resolve_theme_color_key("accent1"), "accent1");
        assert_eq!(resolve_theme_color_key("accent6"), "accent6");
        assert_eq!(resolve_theme_color_key("bg1"), "lt1");
        assert_eq!(resolve_theme_color_key("bg2"), "lt2");
    }
}
