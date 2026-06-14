mod alt_chunk;
mod charts;
mod color;
mod comments;
mod embedded_fonts;
mod group;
mod headers_footers;
mod images;
pub(crate) mod numbering;
mod paragraph;
mod runs;
mod sections;
mod settings;
pub(crate) mod smartart;
mod styles;
mod tables;
pub(crate) mod emf;
mod textbox;
mod wmf;
mod wordart;

use std::collections::{HashMap, HashSet};
use std::io::Read;

use crate::error::Error;
use crate::model::{
    Block, BorderStyle, CellBorder, DocGridType, Document, FrameProperties,
    HRelativeFrom, HorizontalPosition, LineSpacing, ParagraphBorder, ParagraphBorders,
    Section, SectionBreakType, SectionProperties, TabAlignment, TabStop, VRelativeFrom,
};

use styles::{ParagraphStyle, parse_line_spacing, parse_styles, parse_theme};

use embedded_fonts::parse_font_table;
use headers_footers::{parse_endnotes, parse_footnotes};
use numbering::parse_numbering;
use relationships::parse_relationships;
use sections::parse_section_properties;
use settings::parse_settings;
use tables::parse_table_node;

pub(super) const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const WPD_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
pub(super) const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
pub(super) const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const MC_NS_TOP: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
pub(super) const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub(super) const DSP_NS: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";
pub(super) const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
pub(super) const VML_NS: &str = "urn:schemas-microsoft-com:vml";

pub(super) fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

pub(super) fn emu_to_pts(emu: f32) -> f32 {
    emu / 12700.0
}

pub(super) fn emu_attr(node: roxmltree::Node, attr: &str) -> f32 {
    node.attribute(attr)
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0)
        / 12700.0
}

/// Find a child element by name and namespace.
pub(super) fn find_child<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
    namespace: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(namespace))
}

pub(super) fn dml<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, DML_NS)
}

pub(super) fn dml_children<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    parent
        .children()
        .filter(move |n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
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
    parent
        .children()
        .filter(move |n| n.tag_name().name() == name && n.tag_name().namespace() == Some(CHART_NS))
}

pub(in crate::docx) struct ParseContext<'a, R: std::io::Read + std::io::Seek> {
    pub(in crate::docx) styles: &'a styles::StylesInfo,
    pub(in crate::docx) theme: &'a styles::ThemeFonts,
    pub(in crate::docx) rels: &'a HashMap<String, String>,
    pub(in crate::docx) zip: &'a mut zip::ZipArchive<R>,
    pub(in crate::docx) numbering: &'a numbering::NumberingInfo,
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
/// Pattern values other than `nil`/`clear`/`solid` (e.g. `pct25`, `diagStripe`)
/// are approximated by returning the `w:fill` color as a solid — TODO: real
/// pattern rendering. `w:themeFill` resolution is also TODO.
pub(super) fn parse_run_shd(parent: roxmltree::Node) -> Option<[u8; 3]> {
    let shd = wml(parent, "shd")?;
    if shd.attribute((WML_NS, "val")) == Some("nil") {
        return None;
    }
    let fill = shd.attribute((WML_NS, "fill"))?;
    if fill == "auto" || fill == "none" {
        return None;
    }
    parse_hex_color(fill)
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
    node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(WML_NS))
}

pub(super) fn wml_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    wml(node, child).and_then(|n| n.attribute((WML_NS, "val")))
}

pub(super) fn twips_attr(node: roxmltree::Node, attr: &str) -> Option<f32> {
    node.attribute((WML_NS, attr))
        .and_then(|v| v.parse::<f32>().ok())
        .map(twips_to_pts)
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
    })
}

/// Returns `None` when the pBdr element is absent, `Some` when present
/// (even if all individual borders are none/nil).
pub(super) fn parse_paragraph_borders(ppr: roxmltree::Node) -> Option<ParagraphBorders> {
    let pbdr = wml(ppr, "pBdr")?;
    Some(ParagraphBorders {
        top: wml(pbdr, "top").and_then(parse_one_border),
        bottom: wml(pbdr, "bottom").and_then(parse_one_border),
        left: wml(pbdr, "left").and_then(parse_one_border),
        right: wml(pbdr, "right").and_then(parse_one_border),
        between: wml(pbdr, "between").and_then(parse_one_border),
    })
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
    let style = match val {
        "dotted" => BorderStyle::Dotted,
        "dashed" => BorderStyle::Dashed,
        "dashSmallGap" => BorderStyle::DashSmallGap,
        "dashDotStroked" | "dashDot" => BorderStyle::DashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "double" => BorderStyle::Double,
        _ => BorderStyle::Single,
    };
    CellBorder::visible(color, width, style)
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

pub(super) fn parse_frame_props(ppr: roxmltree::Node) -> Option<FrameProperties> {
    let fp = wml(ppr, "framePr")?;
    let attr = |name| fp.attribute((WML_NS, name));
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
        let x_twips: f32 = attr("x")
            .and_then(|v| v.parse().ok())
            .unwrap_or(0.0);
        HorizontalPosition::Offset(x_twips / 20.0)
    };
    let v_anchor = match attr("vAnchor").unwrap_or("text") {
        "margin" => VRelativeFrom::Margin,
        "page" => VRelativeFrom::Page,
        _ => VRelativeFrom::Paragraph,
    };
    let y_pts = attr("y")
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0)
        / 20.0;
    let width = attr("w")
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0)
        / 20.0;
    Some(FrameProperties {
        h_relative_from: h_anchor,
        h_position,
        v_relative_from: v_anchor,
        y_offset: y_pts,
        width,
    })
}

#[allow(dead_code)]
pub(super) fn parse_tab_stops(ppr: roxmltree::Node) -> Vec<TabStop> {
    let (stops, _) = parse_tab_stops_with_clears(ppr);
    stops
}

pub(super) fn parse_tab_stops_with_clears(ppr: roxmltree::Node) -> (Vec<TabStop>, Vec<f32>) {
    let Some(tabs) = wml(ppr, "tabs") else {
        return (vec![], vec![]);
    };
    let mut stops = Vec::new();
    let mut clears = Vec::new();
    for n in tabs
        .children()
        .filter(|n| n.tag_name().name() == "tab" && n.tag_name().namespace() == Some(WML_NS))
    {
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

pub(in crate::docx) fn parse_paragraph_spacing(
    ppr: Option<roxmltree::Node>,
    para_style: Option<&ParagraphStyle>,
    autospacing_font_size: Option<f32>,
) -> (Option<f32>, Option<f32>, Option<LineSpacing>) {
    let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));

    let (space_before, space_after) = if let Some(fs) = autospacing_font_size {
        // Auto-spacing (beforeAutospacing/afterAutospacing="1"): when set, Word
        // uses the font's em-size (≈ font_size) as spacing. Inline "0" disables.
        let before_auto = inline_spacing
            .and_then(|n| n.attribute((WML_NS, "beforeAutospacing")).map(|v| v == "1" || v == "true"))
            .or_else(|| para_style.and_then(|s| s.space_before_autospacing))
            .unwrap_or(false);
        let after_auto = inline_spacing
            .and_then(|n| n.attribute((WML_NS, "afterAutospacing")).map(|v| v == "1" || v == "true"))
            .or_else(|| para_style.and_then(|s| s.space_after_autospacing))
            .unwrap_or(false);

        let effective_fs = para_style.and_then(|s| s.font_size).unwrap_or(fs);

        let sb = if before_auto {
            Some(effective_fs)
        } else {
            inline_spacing
                .and_then(|n| twips_attr(n, "before"))
                .or_else(|| para_style.and_then(|s| s.space_before))
        };
        let sa = if after_auto {
            Some(effective_fs)
        } else {
            inline_spacing
                .and_then(|n| twips_attr(n, "after"))
                .or_else(|| para_style.and_then(|s| s.space_after))
        };
        (sb, sa)
    } else {
        let sb = inline_spacing
            .and_then(|n| twips_attr(n, "before"))
            .or_else(|| para_style.and_then(|s| s.space_before));
        let sa = inline_spacing
            .and_then(|n| twips_attr(n, "after"))
            .or_else(|| para_style.and_then(|s| s.space_after));
        (sb, sa)
    };

    let line_spacing = inline_spacing
        .and_then(|n| {
            n.attribute((WML_NS, "line"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|line_val| parse_line_spacing(n, line_val))
        })
        .or_else(|| para_style.and_then(|s| s.line_spacing));
    (space_before, space_after, line_spacing)
}

/// Convert a character-unit indent attribute (hundredths of character width) to points.
fn chars_to_pts(ind: roxmltree::Node, attr: &str, char_width: f32) -> Option<f32> {
    ind.attribute((WML_NS, attr))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|hundredths| hundredths / 100.0 * char_width)
}

pub(super) fn extract_indents(
    ind: roxmltree::Node,
    char_width: Option<f32>,
) -> (Option<f32>, Option<f32>, Option<f32>, Option<f32>) {
    // Word always writes both twip and *Chars values consistently; when both
    // are present we prefer the twip value since it's Word's pre-resolved
    // measurement. Our char_width approximation (font_size/2) is correct for
    // Latin text but wrong for CJK text (where a character is a full em).
    let cw = char_width.unwrap_or(0.0);
    let has_cw = char_width.is_some();
    (
        twips_attr(ind, "start")
            .or_else(|| twips_attr(ind, "left"))
            .or_else(|| has_cw.then(|| chars_to_pts(ind, "leftChars", cw)).flatten()),
        twips_attr(ind, "end")
            .or_else(|| twips_attr(ind, "right"))
            .or_else(|| has_cw.then(|| chars_to_pts(ind, "rightChars", cw)).flatten()),
        twips_attr(ind, "hanging")
            .or_else(|| has_cw.then(|| chars_to_pts(ind, "hangingChars", cw)).flatten()),
        twips_attr(ind, "firstLine")
            .or_else(|| has_cw.then(|| chars_to_pts(ind, "firstLineChars", cw)).flatten()),
    )
}

pub(super) fn collect_block_nodes<'a>(
    parent: roxmltree::Node<'a, 'a>,
) -> Vec<roxmltree::Node<'a, 'a>> {
    let mut nodes = Vec::new();
    for child in parent.children() {
        if child.tag_name().name() == "sdt" && child.tag_name().namespace() == Some(WML_NS) {
            if let Some(content) = wml(child, "sdtContent") {
                nodes.extend(collect_block_nodes(content));
            }
        } else if child.tag_name().namespace() == Some(MC_NS_TOP)
            && child.tag_name().name() == "AlternateContent"
        {
            // mc:AlternateContent wraps block-level content in mc:Choice/mc:Fallback.
            // Use mc:Fallback for compatibility (it avoids newer namespace requirements).
            let fallback = child.children().find(|n| {
                n.tag_name().namespace() == Some(MC_NS_TOP)
                    && n.tag_name().name() == "Fallback"
            });
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
    Some(content)
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
) -> (Option<String>, Option<String>, Option<String>, Option<String>) {
    let Some(xml_content) = read_zip_text(zip, "docProps/core.xml") else {
        return (None, None, None, None);
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_content) else {
        return (None, None, None, None);
    };

    let root = xml.root_element();

    let mut title = None;
    let mut author = None;
    let mut subject = None;
    let mut keywords = None;

    for child in root.children() {
        if child.tag_name().name() == "title" && child.tag_name().namespace() == Some("http://purl.org/dc/elements/1.1/") {
            if let Some(text) = child.text() {
                title = Some(text.to_string());
            }
        } else if child.tag_name().name() == "creator" && child.tag_name().namespace() == Some("http://purl.org/dc/elements/1.1/") {
            if let Some(text) = child.text() {
                author = Some(text.to_string());
            }
        } else if child.tag_name().name() == "subject" && child.tag_name().namespace() == Some("http://purl.org/dc/elements/1.1/") {
            if let Some(text) = child.text() {
                subject = Some(text.to_string());
            }
        } else if child.tag_name().name() == "keywords" && child.tag_name().namespace() == Some("http://schemas.openxmlformats.org/package/2006/metadata/core-properties") {
            if let Some(text) = child.text() {
                keywords = Some(text.to_string());
            }
        }
    }

    (title, author, subject, keywords)
}

fn parse_zip<R: Read + std::io::Seek>(zip: &mut zip::ZipArchive<R>) -> Result<Document, Error> {
    let settings = parse_settings(zip);
    let theme = parse_theme(zip, settings.east_asia_lang.as_deref());
    let styles = parse_styles(zip, &theme);
    let numbering = parse_numbering(zip);
    let rels = parse_relationships(zip);
    let ft = parse_font_table(zip);
    let (embedded_fonts, font_table) = (ft.embedded_fonts, ft.font_table);
    let footnotes = parse_footnotes(zip, &styles, &theme, &numbering);
    let endnotes = parse_endnotes(zip, &styles, &theme, &numbering);
    let comments = comments::parse_comments(zip);
    let (title, author, subject, keywords) = parse_core_props(zip);

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")
        .map_err(|_| Error::InvalidDocx("missing word/document.xml (is this a DOCX file?)".into()))?
        .read_to_string(&mut xml_content)?;

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = wml(root, "body").ok_or_else(|| Error::InvalidDocx("Missing w:body".into()))?;

    let default_line_pitch = styles.defaults.font_size * 1.2;

    let mut ctx = ParseContext {
        styles: &styles,
        theme: &theme,
        rels: &rels,
        zip,
        numbering: &numbering,
    };

    let mut sections: Vec<Section> = Vec::new();
    let mut blocks = Vec::new();
    let mut counters: HashMap<(u32, u8), u32> = HashMap::new();
    let mut last_seen_level: HashMap<u32, u8> = HashMap::new();
    let mut applied_overrides: HashSet<(u32, u8)> = HashSet::new();

    for node in collect_block_nodes(body) {
        if node.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        match node.tag_name().name() {
            "tbl" => {
                let table = parse_table_node(
                    node,
                    &mut ctx,
                    &mut counters,
                    &mut last_seen_level,
                    &mut applied_overrides,
                );
                blocks.push(Block::Table(table));
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
                let para = paragraph::build_paragraph(
                    node, &mut ctx, &mut counters, &mut last_seen_level, &mut applied_overrides, &opts,
                );

                blocks.push(Block::Paragraph(para));

                // Mid-document section break: sectPr inside pPr ends the current section
                if let Some(sect_node) = ppr.and_then(|ppr| wml(ppr, "sectPr")) {
                    if let Some(Block::Paragraph(last_para)) = blocks.last_mut() {
                        last_para.is_section_break = true;
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
                }
            }
            "altChunk" => {
                if let Some(id) = node.attribute((REL_NS, "id")) {
                    blocks.extend(alt_chunk::parse_alt_chunk(id, ctx.rels, ctx.zip));
                }
            }
            _ => {}
        }
    }

    // Final section: body-level sectPr
    let final_props = if let Some(sect_node) = wml(body, "sectPr") {
        parse_section_properties(sect_node, &mut ctx, default_line_pitch, settings.gutter_at_top)
    } else {
        SectionProperties {
            page_width: 612.0,
            page_height: 792.0,
            margin_top: 72.0,
            margin_bottom: 72.0,
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
        }
    };
    sections.push(Section {
        properties: final_props,
        blocks,
    });

    // Word's PDF export scales the body content via a wrapping `cm` operator
    // when comments are present, so glyphs render at ~76% and a column opens
    // up on the right for the comment pane. We apply the same scaling at
    // content-stream assembly time; the body keeps its natural layout here.

    Ok(Document {
        sections,
        line_spacing: styles.defaults.line_spacing,
        embedded_fonts,
        footnotes,
        endnotes,
        comments,
        font_table,
        even_and_odd_headers: settings.even_and_odd_headers,
        default_tab_stop: settings.default_tab_stop,
        style_id_to_name: styles.style_id_to_name,
        chart_font_name: theme.minor.clone(),
        title,
        author,
        subject,
        keywords,
        auto_hyphenation: settings.auto_hyphenation,
        default_lang: settings.default_lang,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    // --- Pure math / conversion ---

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

        // Pattern value (pct25) with fill → approximated as solid fill
        let xml =
            format!(r#"<w:rPr xmlns:w="{ns}"><w:shd w:val="pct25" w:fill="CCCCCC"/></w:rPr>"#);
        let doc = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(parse_run_shd(doc.root_element()), Some([0xCC, 0xCC, 0xCC]));
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
        let stops = parse_tab_stops(node);
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
    fn test_extract_indents_left_chars() {
        // When both twips and *Chars are present, twips takes priority
        // (Word's pre-resolved value is authoritative).
        let xml = format!(
            r#"<w:ind xmlns:w="{}" w:leftChars="200" w:left="480"/>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        let (left, _right, _hanging, _first_line) = extract_indents(node, Some(5.5));
        assert_eq!(left, Some(24.0)); // w:left="480" twips → 24pt
        // When only *Chars is present, use it with char_width approximation
        let xml2 = format!(r#"<w:ind xmlns:w="{}" w:leftChars="200"/>"#, WML_NS);
        let doc2 = roxmltree::Document::parse(&xml2).unwrap();
        let node2 = doc2.root_element();
        let (left2, _, _, _) = extract_indents(node2, Some(5.5));
        assert_eq!(left2, Some(11.0)); // 200/100 * 5.5 = 11.0pt
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
