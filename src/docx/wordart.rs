use crate::model::{
    Alignment, AutoFit, HRelativeFrom, HorizontalPosition, Paragraph, Run, TextFill, TextGlow,
    TextOutline, TextShadow, TextWarp, Textbox, VRelativeFrom, VerticalPosition, WrapType,
};

use super::styles::{StylesInfo, ThemeFonts};
use super::textbox::parse_avlst;
use super::{DML_NS, W14_NS, dml as find_dml, find_child, parse_hex_color};

fn find_w14<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(parent, name, W14_NS)
}

/// Parsed WordArt-specific properties from wps:bodyPr.
pub(super) struct WordArtBodyProps {
    pub is_wordart: bool,
    pub text_warp: Option<TextWarp>,
    pub auto_fit: AutoFit,
}

/// Parse WordArt-related attributes from a wps:bodyPr node.
pub(super) fn parse_wordart_body_pr(body_pr: roxmltree::Node) -> WordArtBodyProps {
    let is_wordart = body_pr
        .attribute("fromWordArt")
        .is_some_and(|v| v == "1" || v == "true");

    let text_warp = find_dml(body_pr, "prstTxWarp").and_then(|prst| {
        let preset = prst.attribute("prst")?.to_string();
        let adjustments = parse_avlst(prst);
        Some(TextWarp {
            preset,
            adjustments,
        })
    });

    let auto_fit = if find_dml(body_pr, "spAutoFit").is_some() {
        AutoFit::Shape
    } else if let Some(norm) = find_dml(body_pr, "normAutofit") {
        let font_scale = norm
            .attribute("fontScale")
            .and_then(|v| v.parse::<f32>().ok())
            .map(|v| v / 100_000.0);
        let line_space_reduction = norm
            .attribute("lnSpcReduction")
            .and_then(|v| v.parse::<f32>().ok())
            .map(|v| v / 100_000.0);
        AutoFit::Normal {
            font_scale,
            line_space_reduction,
        }
    } else {
        AutoFit::None
    };

    WordArtBodyProps {
        is_wordart,
        text_warp,
        auto_fit,
    }
}

/// Parse w14:textOutline from a w:rPr node.
pub(super) fn parse_text_outline(rpr: roxmltree::Node) -> Option<TextOutline> {
    let outline = find_w14(rpr, "textOutline")?;

    let width_pt = outline
        .attribute((W14_NS, "w"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(super::emu_to_pts)
        .unwrap_or(0.75);

    let color = find_w14_solid_fill_color(outline)?;
    Some(TextOutline { width_pt, color })
}

/// Parse w14:textFill from a w:rPr node.
pub(super) fn parse_text_fill(rpr: roxmltree::Node) -> Option<TextFill> {
    let text_fill = find_w14(rpr, "textFill")?;

    if find_w14(text_fill, "noFill").is_some() {
        return Some(TextFill::NoFill);
    }

    if let Some(color) = find_w14(text_fill, "solidFill").and_then(resolve_w14_color) {
        return Some(TextFill::Solid(color));
    }

    if let Some(grad) = find_w14(text_fill, "gradFill") {
        return parse_w14_gradient(grad);
    }

    None
}

/// Parse w14:shadow from a w:rPr node.
pub(super) fn parse_text_shadow(rpr: roxmltree::Node) -> Option<TextShadow> {
    let shadow = find_w14(rpr, "shadow")?;

    let color = find_w14_solid_fill_color(shadow).unwrap_or([128, 128, 128]);

    let blur_rad = shadow
        .attribute((W14_NS, "blurRad"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(super::emu_to_pts)
        .unwrap_or(0.0);

    let dist = shadow
        .attribute((W14_NS, "dist"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(super::emu_to_pts)
        .unwrap_or(blur_rad.max(1.5));

    // Direction in 60000ths of a degree, clockwise from right
    let dir = shadow
        .attribute((W14_NS, "dir"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 60_000.0)
        .unwrap_or(225.0);

    let alpha = find_w14(shadow, "srgbClr")
        .and_then(|srgb| find_w14(srgb, "alpha"))
        .and_then(|a| a.attribute((W14_NS, "val")))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 100_000.0)
        .unwrap_or(0.6);

    // Convert polar (dist, dir) to cartesian offsets
    let dir_rad = dir.to_radians();
    let offset_x = dist * dir_rad.cos();
    let offset_y = dist * dir_rad.sin();

    Some(TextShadow {
        color,
        offset_x,
        offset_y,
        alpha,
    })
}

/// Parse w14:glow from a w:rPr node.
pub(super) fn parse_text_glow(rpr: roxmltree::Node) -> Option<TextGlow> {
    let glow = find_w14(rpr, "glow")?;

    let radius_pt = glow
        .attribute((W14_NS, "rad"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(super::emu_to_pts)
        .unwrap_or(0.0);

    if radius_pt <= 0.0 {
        return None;
    }

    let color = resolve_w14_color(glow).unwrap_or([255, 255, 0]);

    Some(TextGlow { color, radius_pt })
}

/// Parse VML WordArt when v:textpath is present instead of v:textbox.
pub(super) fn parse_vml_wordart(
    shape: roxmltree::Node,
    textpath: roxmltree::Node,
    styles: &StylesInfo,
    _theme: &ThemeFonts,
) -> Option<Textbox> {
    let text = textpath.attribute("string")?;
    if text.is_empty() {
        return None;
    }

    let style_str = textpath.attribute("style").unwrap_or("");
    let (font_name, font_size) = parse_vml_font_style(style_str);

    let fill_color = shape
        .attribute("fillcolor")
        .and_then(|v| parse_hex_color(v.trim_start_matches('#')));

    let run = Run {
        text: text.to_string(),
        font_name: font_name.unwrap_or_else(|| styles.defaults.font_name.clone()),
        font_size: font_size.unwrap_or(36.0),
        bold: true,
        color: fill_color,
        ..Run::default()
    };

    let para = Paragraph {
        runs: vec![run],
        alignment: Alignment::Center,
        ..Paragraph::default()
    };

    let shape_style = shape.attribute("style").unwrap_or("");
    let (width, height) = parse_vml_dimensions(shape_style);

    Some(Textbox {
        paragraphs: vec![para],
        width_pt: width,
        height_pt: height,
        h_position: HorizontalPosition::Offset(0.0),
        h_relative_from: HRelativeFrom::Column,
        v_offset_pt: 0.0,
        v_position: VerticalPosition::Offset(0.0),
        v_relative_from: VRelativeFrom::Paragraph,
        fill: None,
        shape_type: Default::default(),
        stroke_color: None,
        stroke_width: 0.0,
        text_anchor: Default::default(),
        margin_left: 0.0,
        margin_right: 0.0,
        margin_top: 0.0,
        margin_bottom: 0.0,
        wrap_type: WrapType::None,
        dist_top: 0.0,
        dist_bottom: 0.0,
        behind_doc: false,
        no_text_wrap: true,
        is_wordart: true,
        text_warp: None,
        auto_fit: AutoFit::Shape,
    })
}

fn resolve_w14_color(fill_node: roxmltree::Node) -> Option<[u8; 3]> {
    if let Some(srgb) = find_w14(fill_node, "srgbClr") {
        return srgb.attribute((W14_NS, "val")).and_then(parse_hex_color);
    }
    if let Some(srgb) = find_child(fill_node, "srgbClr", DML_NS) {
        return srgb.attribute("val").and_then(parse_hex_color);
    }
    None
}

fn find_w14_solid_fill_color(parent: roxmltree::Node) -> Option<[u8; 3]> {
    find_w14(parent, "solidFill")
        .and_then(resolve_w14_color)
        .or_else(|| find_dml(parent, "solidFill").and_then(resolve_w14_color))
}

fn parse_w14_gradient(grad: roxmltree::Node) -> Option<TextFill> {
    let gs_lst = find_w14(grad, "gsLst")?;

    let mut stops = Vec::new();
    for gs in gs_lst
        .children()
        .filter(|n| n.tag_name().name() == "gs" && n.tag_name().namespace() == Some(W14_NS))
    {
        let pos = gs
            .attribute((W14_NS, "pos"))
            .and_then(|v| v.parse::<f32>().ok())
            .map(|v| v / 100_000.0)
            .unwrap_or(0.0);
        if let Some(color) = resolve_w14_color(gs) {
            stops.push((color, pos));
        }
    }

    if stops.is_empty() {
        return None;
    }

    let angle_deg = find_w14(grad, "lin")
        .and_then(|lin| lin.attribute((W14_NS, "ang")))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 60_000.0)
        .unwrap_or(0.0);

    Some(TextFill::Gradient { stops, angle_deg })
}

fn parse_vml_font_style(style: &str) -> (Option<String>, Option<f32>) {
    let mut font_name = None;
    let mut font_size = None;

    for part in style.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            let val = val.trim().trim_matches('"').trim_matches('\'');
            match key.trim() {
                "font-family" => font_name = Some(val.to_string()),
                "font-size" => {
                    font_size = val.trim_end_matches("pt").trim().parse::<f32>().ok();
                }
                _ => {}
            }
        }
    }

    (font_name, font_size)
}

fn parse_vml_dimensions(style: &str) -> (f32, f32) {
    let mut width = 0.0_f32;
    let mut height = 0.0_f32;
    let parse_pt =
        |s: &str| -> f32 { s.trim_end_matches("pt").trim().parse::<f32>().unwrap_or(0.0) };

    for part in style.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            match key.trim() {
                "width" => width = parse_pt(val.trim()),
                "height" => height = parse_pt(val.trim()),
                _ => {}
            }
        }
    }

    (width, height)
}
