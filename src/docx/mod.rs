mod alt_chunk;
pub(crate) mod arabic;
mod chartex;
mod charts;
mod color;
mod comments;
mod doc_fields;
mod embedded_fonts;
pub(crate) mod emf;
mod field_code;
mod group;
mod headers_footers;
mod images;
mod math;
mod math_speech;
mod media;
pub(crate) mod numbering;
mod paragraph;
mod runs;
mod sections;
mod settings;
pub(crate) mod smartart;
mod styles;
pub(crate) mod svg;
mod table_borders;
mod tables;
mod textbox;
mod theme;
mod vml_object;
mod vml_shape;
mod wmf;
mod wordart;
mod xml;

use std::collections::HashMap;
use std::io::Read;

use crate::error::Error;
use crate::model::{
    Block, BorderStyle, CellBorder, Document, DropCap, FrameProperties, HRelativeFrom,
    HorizontalPosition, LineSpacing, ParagraphBorder, ParagraphBorders, Section, SectionBreakType,
    TabAlignment, TabStop, VRelativeFrom, VerticalPosition,
};

use styles::{ParagraphStyle, StyleDefaults, parse_line_spacing, parse_styles, parse_theme};

use doc_fields::{
    double_wide_char_spacing, expand_toc_fields, grid_char_pitch, link_textbox_stories,
    resolve_body_stylerefs, resolve_seq_fields,
};
use embedded_fonts::parse_font_table;
use headers_footers::{parse_endnotes, parse_footnotes};
use numbering::parse_numbering;
use relationships::parse_part_relationships;
use sections::parse_section_properties;
use settings::parse_settings;
use tables::parse_table_node;
use xml::*;

/// How one document's tracked changes print. Every part (body, headers,
/// notes) shares one, so an author keeps one colour across the document.
#[derive(Default)]
pub(in crate::docx) struct Revisions {
    /// Print tracked changes as Word's All Markup view (`RevisionView::Markup`).
    pub(in crate::docx) markup: bool,
    authors: std::cell::RefCell<Vec<String>>,
}

impl Revisions {
    /// Word colours each author's revisions by the order the authors first
    /// appear (Word for Mac probes 2026-11-05; a machine that has seen an author
    /// before keeps that author's earlier colour, which no converter can know).
    /// Moves print in Word's move colour whatever the author (11 corpus
    /// redlines with eleven author colours, every move 2C6234).
    pub(in crate::docx) fn revision_color(&self, revision: roxmltree::Node) -> [u8; 3] {
        if is_move(revision) {
            return [0x2C, 0x62, 0x34];
        }
        let author = revision.attribute((WML_NS, "author")).unwrap_or("");
        let mut authors = self.authors.borrow_mut();
        let index = authors.iter().position(|n| n == author).unwrap_or_else(|| {
            authors.push(author.to_string());
            authors.len() - 1
        });
        runs::AUTHOR_COLORS[index % runs::AUTHOR_COLORS.len()]
    }
}

pub(in crate::docx) struct ParseContext<'a, R: std::io::Read + std::io::Seek> {
    pub(in crate::docx) styles: &'a styles::StylesInfo,
    pub(in crate::docx) theme: &'a styles::ThemeFonts,
    pub(in crate::docx) rels: &'a HashMap<String, String>,
    pub(in crate::docx) zip: &'a mut zip::ZipArchive<R>,
    pub(in crate::docx) media: &'a mut media::MediaParts,
    pub(in crate::docx) numbering: &'a numbering::NumberingInfo,
    pub(in crate::docx) revisions: &'a Revisions,
    /// The body's comment ranges (`comment_ranges`); empty for headers,
    /// footers and notes.
    pub(in crate::docx) comments: CommentRanges,
    /// Fields left open by the previous paragraph (`runs::FieldCarry`).
    pub(in crate::docx) fields: runs::FieldCarry,
    /// What the current section's character grid adds to one character
    /// (`sections::grid_char_delta`), the unit of `w:ind` `*Chars` values.
    pub(in crate::docx) grid_char_delta: f32,
    /// The current section's text width (`sections::text_width`), read ahead
    /// like `grid_char_delta`; 0 where no kashida reserve applies.
    pub(in crate::docx) text_width: f32,
    /// Parsing a text box's story (`images::parse_run_drawing`).
    pub(in crate::docx) in_text_box: bool,
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

fn is_text_node(n: roxmltree::Node) -> bool {
    n.has_tag_name((WML_NS, "t")) || n.has_tag_name((WML_NS, "delText"))
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
            // Deleted text is text here too (Word for Mac probes p21/p21b:
            // a comment on a deleted cell's text keeps its balloon, one whose
            // range opens before a wholly deleted table and closes after it,
            // with its reference in a deleted paragraph, loses it in the
            // markup view as well; four external corpus documents print no
            // pane for that reason).
            let text_before = n
                .ancestors()
                .find(|a| a.has_tag_name((WML_NS, "p")))
                .is_some_and(|p| {
                    p.descendants()
                        .take_while(|d| *d != n)
                        .any(|d| is_text_node(d) && d.text().is_some_and(|t| !t.is_empty()))
                });
            if !text_before {
                unballooned.extend(id(n));
            }
        } else if n.has_tag_name((WML_NS, "commentReference")) {
            reference_order.extend(id(n));
        } else if is_text_node(n) && n.text().is_some_and(|t| !t.is_empty()) {
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
    let v = u32::from_str_radix(val, 16).ok()?;
    Some([(v >> 16) as u8, (v >> 8) as u8, v as u8])
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

/// A paragraph mark's tracked change of `kind` ("ins" or "del"); a moved
/// mark counts as the insertion (`moveTo`) or deletion (`moveFrom`) it is.
pub(super) fn mark_revision<'a>(
    p: roxmltree::Node<'a, 'a>,
    kind: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    wml(p, "pPr")
        .and_then(|ppr| wml(ppr, "rPr"))
        .and_then(|rpr| revised_mark(rpr, kind))
}

/// The `w:ins`/`w:moveTo` or `w:del`/`w:moveFrom` in a mark's `w:rPr`.
pub(super) fn revised_mark<'a>(
    rpr: roxmltree::Node<'a, 'a>,
    kind: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    let moved = if kind == "del" { "moveFrom" } else { "moveTo" };
    wml(rpr, kind).or_else(|| wml(rpr, moved))
}

pub(super) fn is_move(revision: roxmltree::Node) -> bool {
    matches!(revision.tag_name().name(), "moveFrom" | "moveTo")
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
            // Word's export draws a heavy leader with plain underscores.
            "underscore" | "heavy" => Some('_'),
            "middleDot" => Some('\u{00B7}'),
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
            // The `*Chars` attributes are hundredths of a character width.
            .or_else(|| {
                let hundredths: f32 = ind.attribute((WML_NS, chars))?.parse().ok()?;
                Some(hundredths / 100.0 * char_width?)
            })
    };
    (
        side("leftChars", &["start", "left"]),
        side("rightChars", &["end", "right"]),
        side("hangingChars", &["hanging"]),
        side("firstLineChars", &["firstLine"]),
    )
}

// Final view omits an empty paragraph whose paragraph mark was deleted.
// Do not discard surviving runs or bookmarks: deleting the mark alone merges
// text with the next paragraph rather than deleting it. A section break in
// the deleted mark goes with it (Word online export probe 2026-11-05: three
// sections whose break marks are deleted print as one, with the last header).
fn is_fully_deleted_paragraph(node: roxmltree::Node) -> bool {
    if !node.has_tag_name((WML_NS, "p")) || mark_revision(node, "del").is_none() {
        return false;
    }
    node.children().filter(|n| n.is_element()).all(|n| {
        n.tag_name().namespace() == Some(WML_NS)
            && matches!(n.tag_name().name(), "pPr" | "del" | "moveFrom" | "proofErr")
    })
}

pub(super) fn collect_block_nodes<'a>(
    parent: roxmltree::Node<'a, 'a>,
    markup: bool,
) -> Vec<roxmltree::Node<'a, 'a>> {
    let mut nodes = Vec::new();
    for child in parent.children() {
        // The markup view keeps a deleted paragraph, struck through.
        if !markup && is_fully_deleted_paragraph(child) {
            continue;
        }
        if child.has_tag_name((WML_NS, "sdt")) {
            if let Some(content) = wml(child, "sdtContent") {
                nodes.extend(collect_block_nodes(content, markup));
            }
        } else if child.has_tag_name((WML_NS, "customXml")) {
            // w:customXml is a transparent wrapper (block/row/cell): its children
            // ARE the content. Same descent drives tbl→tr and tr→tc unwrapping.
            nodes.extend(collect_block_nodes(child, markup));
        } else if child.has_tag_name((MC_NS_TOP, "AlternateContent")) {
            // mc:AlternateContent wraps block-level content in mc:Choice/mc:Fallback.
            // Use mc:Fallback for compatibility (it avoids newer namespace requirements).
            let fallback = child
                .children()
                .find(|n| n.has_tag_name((MC_NS_TOP, "Fallback")));
            if let Some(fb) = fallback {
                nodes.extend(collect_block_nodes(fb, markup));
            }
        } else {
            nodes.push(child);
        }
    }
    nodes
}

/// A document parsed from `(zip path, XML)` parts.
#[cfg(test)]
pub(crate) fn parse_parts(parts: &[(&str, impl AsRef<[u8]>)]) -> Document {
    use std::io::Write;
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, xml) in parts {
        zip.start_file(*name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(xml.as_ref()).unwrap();
    }
    parse_bytes(
        &zip.finish().unwrap().into_inner(),
        crate::RevisionView::Final,
    )
    .unwrap()
}

/// `parse_parts` with only a word/document.xml.
#[cfg(test)]
pub(crate) fn parse_document_xml(xml: &str) -> Document {
    parse_parts(&[("word/document.xml", xml)])
}

pub(super) fn read_zip_text<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    name: &str,
) -> Option<String> {
    let mut content = String::new();
    zip.by_name(name).ok()?.read_to_string(&mut content).ok()?;
    Some(join_split_run_properties(strict_to_transitional(content)))
}

/// Word reads a run's properties split over two adjacent `w:rPr` elements as
/// one: french_sexual_health's URLs carry the Hyperlink style in the first
/// and `sz` 18 in the second, and Word draws them blue, underlined and 9pt.
fn join_split_run_properties(xml: String) -> String {
    if xml.contains("</w:rPr><w:rPr>") {
        xml.replace("</w:rPr><w:rPr>", "")
    } else {
        xml
    }
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

    pub(in crate::docx) fn parse_part_relationships<R: Read + std::io::Seek>(
        zip: &mut zip::ZipArchive<R>,
        part_path: &str,
    ) -> HashMap<String, String> {
        let Some(xml_content) = read_zip_text(zip, &rels_path(part_path)) else {
            return HashMap::new();
        };
        parse_rels_xml(&xml_content)
    }

    fn rels_path(part_path: &str) -> String {
        match part_path.rsplit_once('/') {
            Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
            None => format!("_rels/{part_path}.rels"),
        }
    }

    /// The zip path of the part `part_path` relates to with a relationship
    /// type ending `/type_name`.
    pub(in crate::docx) fn related_part<R: Read + std::io::Seek>(
        zip: &mut zip::ZipArchive<R>,
        part_path: &str,
        type_name: &str,
    ) -> Option<String> {
        let xml_content = read_zip_text(zip, &rels_path(part_path))?;
        let xml = roxmltree::Document::parse(&xml_content).ok()?;
        let target = xml.root_element().children().find_map(|n| {
            let ty = n.attribute("Type")?;
            (ty.rsplit('/').next() == Some(type_name)).then(|| n.attribute("Target"))?
        })?;
        Some(match target.strip_prefix('/') {
            Some(absolute) => absolute.to_string(),
            None => match part_path.rsplit_once('/') {
                Some((dir, _)) => format!("{dir}/{target}"),
                None => target.to_string(),
            },
        })
    }
}

pub fn parse(path: &std::path::Path, view: crate::RevisionView) -> Result<Document, Error> {
    let file = std::fs::File::open(path).map_err(|e| match e.kind() {
        std::io::ErrorKind::NotFound | std::io::ErrorKind::PermissionDenied => Error::Io(
            std::io::Error::new(e.kind(), format!("{}: {}", e, path.display())),
        ),
        _ => Error::Io(e),
    })?;

    let mut zip = zip::ZipArchive::new(file)
        .map_err(|_| Error::InvalidDocx("file is not a ZIP archive".into()))?;

    parse_zip(&mut zip, view)
}

pub fn parse_bytes(bytes: &[u8], view: crate::RevisionView) -> Result<Document, Error> {
    let cursor = std::io::Cursor::new(bytes);
    let mut zip = zip::ZipArchive::new(cursor)
        .map_err(|_| Error::InvalidDocx("data is not a valid ZIP/DOCX archive".into()))?;

    parse_zip(&mut zip, view)
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

fn parse_zip<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    view: crate::RevisionView,
) -> Result<Document, Error> {
    let markup = view == crate::RevisionView::Markup;
    let revisions = Revisions {
        markup,
        ..Default::default()
    };
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
    let rels = parse_part_relationships(zip, "word/document.xml");
    let ft = parse_font_table(zip);
    let (embedded_fonts, font_table) = (ft.embedded_fonts, ft.font_table);
    let mut media = media::MediaParts::default();
    let (footnotes, footnote_separator) =
        parse_footnotes(zip, &mut media, &styles, &theme, &numbering, &revisions);
    let (endnotes, endnote_separator) =
        parse_endnotes(zip, &mut media, &styles, &theme, &numbering, &revisions);
    let mut comments = comments::parse_comments(zip);
    let (title, author, subject, keywords) = parse_core_props(zip);

    let xml_content = read_zip_text(zip, "word/document.xml").ok_or_else(|| {
        Error::InvalidDocx("missing word/document.xml (is this a DOCX file?)".into())
    })?;

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
    let sect_prs: Vec<_> = body
        .descendants()
        .filter(|n| n.has_tag_name((WML_NS, "sectPr")))
        .collect();
    let grid_deltas: Vec<f32> = sect_prs
        .iter()
        .map(|&n| sections::grid_char_delta(n))
        .collect();
    let text_widths: Vec<f32> = sect_prs.iter().map(|&n| sections::text_width(n)).collect();
    let mut sect_idx = 0;
    let mut ctx = ParseContext {
        styles: &styles,
        theme: &theme,
        rels: &rels,
        zip,
        media: &mut media,
        numbering: &numbering,
        revisions: &revisions,
        comments: comment_ranges,
        fields: Default::default(),
        grid_char_delta: grid_deltas.first().copied().unwrap_or(0.0),
        text_width: text_widths.first().copied().unwrap_or(0.0),
        in_text_box: false,
    };

    let mut sections: Vec<Section> = Vec::new();
    let mut blocks = Vec::new();
    let mut lists = numbering::ListCounters::default();
    // numId of the previous body block when it was a list paragraph.
    let mut prev_list_num_id: Option<String> = None;

    for node in collect_block_nodes(body, markup) {
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
                let (_, para_style) = styles.paragraph_style_of(ppr);

                let opts = paragraph::ParagraphOptions {
                    resolve_bookmarks: true,
                    resolve_outline_level: true,
                    resolve_drawings: true,
                    collect_extra_textboxes: true,
                    style_num_id: para_style.and_then(|s| s.num_id.clone()),
                    style_num_ilvl: para_style.and_then(|s| s.num_ilvl),
                };
                let mut para = paragraph::build_paragraph(node, &mut ctx, &mut lists, &opts);
                let mark_deleted = para.mark_deleted;

                // Final view: a deleted paragraph mark joins its text to the
                // next paragraph, which keeps its own properties (Word online
                // export probe 2026-11-05: "Para C, deleted mark after it.Para D.").
                if !markup
                    && matches!(blocks.last(), Some(Block::Paragraph(prev)) if prev.mark_deleted)
                    && let Some(Block::Paragraph(prev)) = blocks.pop()
                {
                    // Appending to prev's runs keeps a chain of merges linear.
                    let mut runs = prev.runs;
                    runs.append(&mut para.runs);
                    para.runs = runs;
                }

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

                // Mid-document section break: sectPr inside pPr ends the current section.
                // A deleted paragraph mark takes its section break with it in the
                // final view; the markup view keeps the break (Word probes
                // 2026-11-05: online export one page with the last section's
                // header, inline markup export three pages, one header each).
                if let Some(sect_node) = ppr
                    .and_then(|ppr| wml(ppr, "sectPr"))
                    .filter(|_| markup || !mark_deleted)
                {
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
                    ctx.text_width = text_widths.get(sect_idx).copied().unwrap_or(0.0);
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

    // Final section: body-level sectPr. A body without one takes what an
    // empty sectPr parses to, so the defaults live in one place.
    let (empty_xml, empty_sect_pr);
    let sect_node = match wml(body, "sectPr") {
        Some(node) => node,
        None => {
            empty_xml = format!(r#"<w:sectPr xmlns:w="{WML_NS}"/>"#);
            empty_sect_pr = roxmltree::Document::parse(&empty_xml).expect("constant XML");
            empty_sect_pr.root_element()
        }
    };
    let final_props = parse_section_properties(
        sect_node,
        &mut ctx,
        default_line_pitch,
        settings.gutter_at_top,
    );
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
        .and_then(|s| s.run.font_name.clone())
        .unwrap_or_else(|| "Times New Roman".into());
    let comment_font_size = balloon_style.and_then(|s| s.run.font_size).unwrap_or(9.0);
    // With no CommentText style the text takes Normal's face, as the built-in
    // CommentText would: case63's Aptos Normal over a Cambria theme draws its
    // comments in Aptos (whose "✓" Cambria lacks).
    let comment_body_font = ["CommentText", styles.default_paragraph_style_id.as_str()]
        .iter()
        .find_map(|id| styles.paragraph_styles.get(*id)?.run.font_name.clone())
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
        ul_trail_space: settings.ul_trail_space,
        wp_justification: settings.wp_justification,
        additive_paragraph_spacing: settings.no_html_auto_spacing,
        adjust_line_height_in_table: settings.adjust_line_height_in_table,
        math: settings.math,
    };
    // Captions get their numbers before a STYLEREF quotes them.
    resolve_seq_fields(&mut doc);
    resolve_body_stylerefs(&mut doc);
    link_textbox_stories(&mut doc);
    if settings.balance_sbdb_width {
        double_wide_char_spacing(&mut doc);
    }
    grid_char_pitch(&mut doc, &grid_deltas);
    Ok(doc)
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
    fn deleted_text_counts_for_a_comment_range() {
        // Word for Mac probes p21/p21b (2026-11-05): a comment whose range
        // closes after deleted text in a paragraph keeps its balloon; one
        // whose range opens before a wholly deleted table and closes between
        // blocks loses it, in the markup view too.
        let xml = format!(
            r#"<w:body xmlns:w="{WML_NS}">
<w:p><w:commentRangeStart w:id="1"/><w:del w:id="9" w:author="a"><w:r><w:delText>gone</w:delText></w:r></w:del><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>
<w:commentRangeStart w:id="2"/><w:tbl><w:tr><w:tc><w:p><w:del w:id="8" w:author="a"><w:r><w:delText>cell</w:delText></w:r></w:del></w:p></w:tc></w:tr></w:tbl><w:commentRangeEnd w:id="2"/>
<w:p><w:r><w:commentReference w:id="2"/></w:r></w:p>
<w:p><w:commentRangeStart w:id="3"/><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>
</w:body>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let ranges = super::comment_ranges(doc.root_element());
        assert!(
            ranges.shown.contains(&1),
            "deleted text before the end keeps the balloon"
        );
        assert!(
            !ranges.shown.contains(&2),
            "a range over a deleted table between blocks loses it"
        );
        assert!(
            ranges.shown.contains(&3),
            "an empty range keeps its balloon"
        );
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
                true,
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
            let blocks = collect_block_nodes(doc.root_element(), false);
            assert_eq!(blocks.len(), usize::from(!deleted), "{paragraph}");
        }
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
        let c = |l, h, f| IndentChars {
            left: l,
            right: None,
            hanging: h,
            first_line: f,
        };
        assert_eq!(
            case((10.0, 10.0), c(None, Some(2.0), None), 12.0),
            (0.0, 24.0)
        );
        assert_eq!(
            case((24.0, 10.0), c(None, Some(2.0), None), 12.0),
            (0.0, 24.0)
        );
        assert_eq!(
            case((0.0, 0.0), c(Some(2.0), Some(2.0), None), 12.0),
            (24.0, 48.0)
        );
        assert_eq!(
            case((0.0, 20.0), c(Some(2.0), None, None), 12.0),
            (24.0, 44.0)
        );
        assert_eq!(
            case((48.0, 0.0), c(None, None, Some(2.0)), 12.0),
            (72.0, 48.0)
        );
        assert_eq!(
            case((0.0, 0.0), c(Some(2.0), None, Some(1.0)), 12.0),
            (36.0, 24.0)
        );
        assert_eq!(
            case((10.0, 10.0), c(None, Some(2.0), None), 18.0),
            (0.0, 36.0)
        );
        assert_eq!(
            case((10.0, 10.0), IndentChars::default(), 12.0),
            (0.0, 10.0)
        );
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
