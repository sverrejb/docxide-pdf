//! VML shapes: text boxes, preset outlines, lines and their style strings.

use std::io::Read;

use crate::model::{
    ArrowEnd, ConnectorShape, ConnectorType, HRelativeFrom, HorizontalPosition, ShapeFill,
    ShapeGeometry, Textbox, VRelativeFrom, VerticalPosition,
};

use super::images::{extent_dimensions, place_anchored};
use super::textbox::{TxbxContent, parse_textbox_from_wsp, parse_txbx_content};
use super::vml_object::{vml_frame_fill_image, vml_z_order};
use super::{MC_NS_TOP, OFFICE_NS, ParseContext, VML_NS, WML_NS, WPD_NS, parse_pt};

pub(super) struct VmlBoxStyle {
    pub(super) width: f32,
    pub(super) height: f32,
    pub(super) margin_left: f32,
    pub(super) margin_top: f32,
    pub(super) h_relative: HRelativeFrom,
    pub(super) v_relative: VRelativeFrom,
    /// Raw VML z-index; `vml_z_order` turns it into the layer and stacking order.
    pub(super) z_index: i64,
    /// `mso-position-horizontal:center` / `mso-position-vertical:center`.
    pub(super) h_center: bool,
    pub(super) v_center: bool,
    /// `rotation:` in degrees, clockwise.
    pub(super) rotation_deg: f32,
}

/// Parse a VML shape `style` attribute (semicolon-separated `key:val`) into box
/// geometry: dimensions/margins in points and the position-relative anchors.
pub(super) fn parse_vml_box_style(style_str: &str) -> VmlBoxStyle {
    let mut s = VmlBoxStyle {
        width: 0.0,
        height: 0.0,
        margin_left: 0.0,
        margin_top: 0.0,
        h_relative: HRelativeFrom::Column,
        v_relative: VRelativeFrom::Paragraph,
        z_index: 0,
        h_center: false,
        v_center: false,
        rotation_deg: 0.0,
    };
    for part in style_str.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            let val = val.trim();
            match key.trim() {
                "width" => s.width = parse_pt(val).unwrap_or(0.0),
                "height" => s.height = parse_pt(val).unwrap_or(0.0),
                "margin-left" => s.margin_left = parse_pt(val).unwrap_or(0.0),
                "margin-top" => s.margin_top = parse_pt(val).unwrap_or(0.0),
                "z-index" => s.z_index = val.parse().unwrap_or(0),
                "rotation" => {
                    s.rotation_deg = val.trim_end_matches("deg").parse().unwrap_or(0.0);
                }
                "mso-position-horizontal" => s.h_center = val == "center",
                "mso-position-vertical" => s.v_center = val == "center",
                "mso-position-horizontal-relative" => {
                    s.h_relative = match val {
                        "page" => HRelativeFrom::Page,
                        "margin" => HRelativeFrom::Margin,
                        _ => HRelativeFrom::Column,
                    };
                }
                "mso-position-vertical-relative" => {
                    s.v_relative = match val {
                        "page" => VRelativeFrom::Page,
                        "margin" => VRelativeFrom::Margin,
                        "top-margin-area" => VRelativeFrom::TopMargin,
                        "bottom-margin-area" => VRelativeFrom::BottomMargin,
                        _ => VRelativeFrom::Paragraph,
                    };
                }
                _ => {}
            }
        }
    }
    s
}

/// Parse a VML color value: `#RRGGBB`, a named color, or a theme token like
/// `white [3212]` (the `[idx]` suffix is the resolved theme color we ignore in
/// favor of the spelled-out name). Returns None for unrecognized values so the
/// caller can fall back to a default.
pub(super) fn parse_vml_color(val: &str) -> Option<[u8; 3]> {
    let val = val.trim().split('[').next().unwrap_or("").trim();
    if let Some(hex) = val.strip_prefix('#') {
        if hex.len() == 6 {
            return super::parse_hex_color(hex);
        }
        if hex.len() == 3 {
            let expand = |c: &str| u8::from_str_radix(&c.repeat(2), 16).ok();
            return Some([
                expand(&hex[0..1])?,
                expand(&hex[1..2])?,
                expand(&hex[2..3])?,
            ]);
        }
        return None;
    }
    match val.to_ascii_lowercase().as_str() {
        "black" => Some([0, 0, 0]),
        "white" => Some([255, 255, 255]),
        "red" => Some([255, 0, 0]),
        "green" => Some([0, 128, 0]),
        "blue" => Some([0, 0, 255]),
        "yellow" => Some([255, 255, 0]),
        "gray" | "grey" => Some([128, 128, 128]),
        "silver" => Some([192, 192, 192]),
        _ => None,
    }
}

/// Map a VML shape type code (`o:spt`, also encoded as the `t<N>` suffix of a
/// shapetype id like `#_x0000_t87`) to a DrawingML preset geometry name for the
/// stroke-only shapes we can render. Only the brace/bracket family is mapped —
/// these are the common stroke-only VML presets that carry no textbox.
fn vml_spt_to_preset(spt: u32) -> Option<&'static str> {
    match spt {
        85 => Some("leftBracket"),
        86 => Some("rightBracket"),
        87 => Some("leftBrace"),
        88 => Some("rightBrace"),
        _ => None,
    }
}

/// Resolve a VML shape's preset geometry from its `o:spt` attribute or its
/// `type="#_x0000_t<N>"` reference.
fn vml_shape_preset(shape: roxmltree::Node) -> Option<&'static str> {
    vml_spt_to_preset(vml_spt(shape)?)
}

fn vml_spt(shape: roxmltree::Node) -> Option<u32> {
    shape
        .attribute((OFFICE_NS, "spt"))
        .and_then(|v| v.trim().parse::<f64>().ok().map(|f| f as u32))
        .or_else(|| {
            shape
                .attribute("type")
                .and_then(|t| t.rsplit_once('t'))
                .and_then(|(_, n)| n.parse::<u32>().ok())
        })
}

/// Build a paragraph-less stroked textbox for a VML preset shape that has no
/// textbox/textpath (e.g. a curly brace). Returns None for unrecognized presets
/// so such shapes remain dropped rather than rendered as a stray rectangle.
fn parse_vml_geometry_shape(shape: roxmltree::Node) -> Option<Textbox> {
    // A plain v:rect / v:oval / v:roundrect names its geometry by element.
    let element_preset = match shape.tag_name().name() {
        "rect" => Some("rect"),
        "roundrect" => Some("roundRect"),
        "oval" => Some("ellipse"),
        _ => None,
    };
    let spt_preset = vml_shape_preset(shape);
    let preset = spt_preset.or(element_preset)?;

    // VML fills white by default; only a written fillcolor is drawn
    // (ponytail: the implicit white fill would cover text we place behind).
    // A page-wide v:rect filled #2c3e50 behind a book cover's white title.
    let fill = (!matches!(shape.attribute("filled"), Some("f") | Some("false")))
        .then(|| shape.attribute("fillcolor").and_then(parse_vml_color))
        .flatten()
        .map(ShapeFill::Solid);
    // A shape named only by its element draws the default black outline
    // only when the stroke is written out, like the fill.
    let stroked = !matches!(shape.attribute("stroked"), Some("f") | Some("false"))
        && (spt_preset.is_some()
            || shape.attribute("stroked").is_some()
            || shape.attribute("strokecolor").is_some());
    if !stroked && fill.is_none() {
        return None;
    }
    let (stroke_color, stroke_width) = vml_stroke(shape);

    let VmlBoxStyle {
        width,
        height,
        margin_left,
        margin_top,
        h_relative,
        v_relative,
        z_index,
        ..
    } = parse_vml_box_style(shape.attribute("style").unwrap_or(""));
    let (behind_doc, z_index) = vml_z_order(z_index);
    if width <= 0.0 || height <= 0.0 {
        return None;
    }

    Some(Textbox {
        width_pt: width,
        height_pt: height,
        h_position: HorizontalPosition::Offset(margin_left),
        h_relative_from: h_relative,
        v_offset_pt: margin_top,
        v_position: VerticalPosition::Offset(margin_top),
        v_relative_from: v_relative,
        shape_type: ShapeGeometry {
            preset: Some(preset.to_string()),
            adjustments: Vec::new(),
            custom: None,
        },
        stroke_color: stroked.then_some(stroke_color),
        stroke_width,
        fill,
        behind_doc,
        z_index,
        no_text_wrap: true,
        ..Textbox::default()
    })
}

/// VML textboxes have an opaque white solid fill unless explicitly disabled.
fn parse_vml_textbox_fill(shape: roxmltree::Node) -> Option<ShapeFill> {
    let fill = shape.children().find(|n| n.has_tag_name((VML_NS, "fill")));
    if matches!(shape.attribute("filled"), Some("f" | "false" | "0"))
        || fill.is_some_and(|n| matches!(n.attribute("on"), Some("f" | "false" | "0")))
    {
        return None;
    }
    // Other VML fill types need their own renderer, not an opaque approximation.
    if fill.is_some_and(|n| n.attribute("type").is_some_and(|t| t != "solid")) {
        return None;
    }
    let color = shape
        .attribute("fillcolor")
        .or_else(|| fill.and_then(|n| n.attribute("color")))
        .and_then(parse_vml_color)
        .unwrap_or([255, 255, 255]);
    Some(ShapeFill::Solid(color))
}

/// Stroke-only legacy line shapes are connectors, including zero-height
/// horizontal rules. Do not turn unrelated VML presets into rectangles.
pub(super) fn parse_connector_from_vml(pict: roxmltree::Node) -> Option<ConnectorShape> {
    let shape = pict
        .children()
        .find(|n| n.has_tag_name((VML_NS, "shape")))?;
    let mut line = parse_vml_line(shape)?;
    let style = shape.attribute("style").unwrap_or("");
    if vml_style_prop(style, "visibility") == Some("hidden") {
        return None;
    }
    let b = parse_vml_box_style(style);
    if !b.width.is_finite()
        || !b.height.is_finite()
        || b.width < 0.0
        || b.height < 0.0
        || (b.width == 0.0 && b.height == 0.0)
    {
        return None;
    }
    (line.behind_doc, line.z_index) = vml_z_order(b.z_index);
    line.h_position = HorizontalPosition::Offset(b.margin_left);
    line.h_relative_from = b.h_relative;
    line.v_position = VerticalPosition::Offset(b.margin_top);
    line.v_relative_from = b.v_relative;
    line.width = b.width;
    line.height = b.height;
    Some(line)
}

/// A VML shape's outline colour and width, black 0.75pt unless written out.
fn vml_stroke(shape: roxmltree::Node) -> ([u8; 3], f32) {
    let color = shape.attribute("strokecolor").and_then(parse_vml_color);
    let width = shape
        .attribute("strokeweight")
        .and_then(parse_pt)
        .filter(|w| *w > 0.0);
    (color.unwrap_or([0, 0, 0]), width.unwrap_or(0.75))
}

/// A stroked VML line (`o:spt` 32) with zeroed position and size; the caller
/// supplies the placement.
pub(super) fn parse_vml_line(shape: roxmltree::Node) -> Option<ConnectorShape> {
    if vml_spt(shape) != Some(32) || matches!(shape.attribute("stroked"), Some("f" | "false")) {
        return None;
    }
    let flip = vml_style_prop(shape.attribute("style").unwrap_or(""), "flip").unwrap_or("");
    let (stroke_color, stroke_width) = vml_stroke(shape);
    Some(ConnectorShape {
        behind_doc: false,
        h_position: HorizontalPosition::Offset(0.0),
        h_relative_from: HRelativeFrom::default(),
        v_position: VerticalPosition::Offset(0.0),
        v_relative_from: VRelativeFrom::default(),
        width: 0.0,
        height: 0.0,
        stroke_color,
        stroke_width: if stroke_width.is_finite() {
            stroke_width
        } else {
            0.75
        },
        connector_type: ConnectorType::Line {
            flip_h: flip.contains('x'),
            flip_v: flip.contains('y'),
        },
        head_end: ArrowEnd::None,
        tail_end: ArrowEnd::None,
        z_index: 0,
        v_from_line: false,
        on_line: false,
    })
}

pub(super) fn vml_style_prop<'a>(style: &'a str, key: &str) -> Option<&'a str> {
    style
        .split(';')
        .filter_map(|p| p.trim().split_once(':'))
        .find_map(|(k, v)| (k.trim() == key).then_some(v.trim()))
}

pub(super) fn parse_textbox_from_vml<R: Read + std::io::Seek>(
    pict_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<Textbox> {
    let shape = pict_node.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS) && matches!(n.tag_name().name(), "shape" | "rect")
    })?;

    // VML WordArt uses v:textpath instead of v:textbox
    let Some(textbox_node) = shape
        .children()
        .find(|n| n.has_tag_name((VML_NS, "textbox")))
    else {
        if let Some(tp) = shape
            .children()
            .find(|n| n.has_tag_name((VML_NS, "textpath")))
        {
            return super::wordart::parse_vml_wordart(shape, tp, ctx.styles);
        }
        // Stroke-only VML preset shapes (e.g. a curly brace, o:spt 87) have
        // neither a textbox nor a textpath. They were previously dropped
        // entirely (annotation #174 — the blue "{" wrapping the paragraphs).
        // Build a paragraph-less stroked shape from the preset geometry.
        return parse_vml_geometry_shape(shape);
    };
    let txbx_content = textbox_node
        .children()
        .find(|n| n.has_tag_name((WML_NS, "txbxContent")))?;

    let style_str = shape.attribute("style").unwrap_or("");
    let VmlBoxStyle {
        width,
        height,
        margin_left,
        margin_top,
        h_relative,
        v_relative,
        z_index,
        ..
    } = parse_vml_box_style(style_str);
    let (behind_doc, z_index) = vml_z_order(z_index);

    // VML shapes are stroked by default (black, 0.75pt) unless stroked="f".
    // Word renders this border, so honor it — previously every VML textbox was
    // forced borderless, dropping e.g. the "Nota Bene" box outline.
    let stroked = !matches!(shape.attribute("stroked"), Some("f") | Some("false"));
    let (stroke_color, stroke_width) = if stroked {
        let (color, width) = vml_stroke(shape);
        (Some(color), width)
    } else {
        (None, 0.0)
    };

    let TxbxContent { paragraphs, tables } = parse_txbx_content(txbx_content, ctx);
    if paragraphs.is_empty() {
        return None;
    }
    Some(Textbox {
        paragraphs,
        tables,
        width_pt: width,
        height_pt: height,
        h_position: HorizontalPosition::Offset(margin_left),
        h_relative_from: h_relative,
        v_offset_pt: margin_top,
        v_position: VerticalPosition::Offset(margin_top),
        v_relative_from: v_relative,
        stroke_color,
        stroke_width,
        behind_doc,
        z_index,
        fill: parse_vml_textbox_fill(shape),
        picture_fill: vml_frame_fill_image(shape, ctx, width, height),
        margin_left: 7.2,
        margin_right: 7.2,
        margin_top: 3.6,
        margin_bottom: 3.6,
        ..Textbox::default()
    })
}

pub(super) fn collect_textboxes_from_paragraph<R: Read + std::io::Seek>(
    para_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Vec<Textbox> {
    let mut textboxes = Vec::new();

    for child in para_node.children() {
        let ns = child.tag_name().namespace();
        let name = child.tag_name().name();
        if ns == Some(MC_NS_TOP) && name == "AlternateContent" {
            let choice = child
                .children()
                .find(|n| n.has_tag_name((MC_NS_TOP, "Choice")));
            let fallback = child
                .children()
                .find(|n| n.has_tag_name((MC_NS_TOP, "Fallback")));

            if let Some(branch) = choice {
                for drawing in branch
                    .children()
                    .filter(|n| n.has_tag_name((WML_NS, "drawing")))
                {
                    for container in drawing
                        .children()
                        .filter(|n| n.has_tag_name((WPD_NS, "anchor")))
                    {
                        if let Some(wsp) = parse_textbox_from_wsp(container, ctx) {
                            let size = extent_dimensions(container);
                            // This path keeps a line-relative box at its
                            // paragraph's top, as it always has.
                            textboxes.push(Textbox {
                                v_from_line: false,
                                ..place_anchored(container, size, wsp)
                            });
                        }
                    }
                }
            } else if let Some(branch) = fallback {
                for pict in branch
                    .children()
                    .filter(|n| n.has_tag_name((WML_NS, "pict")))
                {
                    if let Some(tb) = parse_textbox_from_vml(pict, ctx) {
                        textboxes.push(tb);
                    }
                }
                for r in branch.children().filter(|n| n.has_tag_name((WML_NS, "r"))) {
                    for pict in r.children().filter(|n| n.has_tag_name((WML_NS, "pict"))) {
                        if let Some(tb) = parse_textbox_from_vml(pict, ctx) {
                            textboxes.push(tb);
                        }
                    }
                }
            }
        }
    }
    textboxes
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn legacy_zero_height_line_keeps_style_and_rejects_non_lines() {
        let xml = format!(
            r##"<w:pict xmlns:w="{WML_NS}" xmlns:v="{VML_NS}"><v:shape type="#_x0000_t32" style="width:240pt;height:0;margin-left:8pt;margin-top:4pt;z-index:42;mso-position-vertical-relative:page" strokecolor="#e36c0a [2409]" strokeweight="1.75pt"/></w:pict>"##
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let line = parse_connector_from_vml(doc.root_element()).unwrap();
        assert_eq!((line.width, line.height), (240.0, 0.0));
        assert_eq!(line.stroke_color, [227, 108, 10]);
        assert_eq!(line.stroke_width, 1.75);
        assert_eq!(line.z_index, 42);
        assert!(matches!(line.v_relative_from, VRelativeFrom::Page));
        for change in [
            xml.replace("_t32", "_t87"),
            xml.replace("height:0", "height:-1pt"),
            xml.replace("z-index:42", "visibility:hidden"),
            xml.replace("strokeweight=", "stroked=\"f\" strokeweight="),
        ] {
            let doc = roxmltree::Document::parse(&change).unwrap();
            assert!(parse_connector_from_vml(doc.root_element()).is_none());
        }
    }

    #[test]
    fn legacy_textbox_fill_is_white_unless_disabled_or_colored() {
        for (attrs, child, expected) in [
            ("", "", Some([255, 255, 255])),
            ("filled=\"f\"", "", None),
            ("", "<v:fill on=\"false\"/>", None),
            ("fillcolor=\"#fff080\"", "", Some([255, 240, 128])),
            ("", "<v:fill type=\"gradient\"/>", None),
        ] {
            let xml = format!(r#"<v:shape xmlns:v="{VML_NS}" {attrs}>{child}</v:shape>"#);
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let actual = match parse_vml_textbox_fill(doc.root_element()) {
                Some(ShapeFill::Solid(color)) => Some(color),
                None => None,
                _ => panic!("expected a solid fill"),
            };
            assert_eq!(actual, expected);
        }
    }
}
