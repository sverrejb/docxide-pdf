use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    EmbeddedImage, FloatingImage, HRelativeFrom, HorizontalPosition, ImageFormat, ImageShadow,
    InlineChart, SmartArtDiagram, VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use super::charts::parse_chart_from_zip;
use super::numbering::NumberingInfo;
use super::smartart::{has_diagram_ref, parse_smartart_drawing};
use super::styles::{StylesInfo, ThemeFonts};
use super::textbox::{parse_connector_from_wsp, parse_textbox_from_wsp};
use super::{DML_NS, REL_NS, WML_NS, WPD_NS, emu_attr, emu_to_pts, find_child, parse_hex_color, wml};

const CHART_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

fn parse_emu_text(text: Option<&str>) -> f32 {
    emu_to_pts(text.unwrap_or("0").parse::<f32>().unwrap_or(0.0))
}

fn wpd<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, WPD_NS)
}

fn wpd_child_text<'a>(parent: Option<roxmltree::Node<'a, 'a>>, name: &str) -> Option<&'a str> {
    parent
        .and_then(|n| n.children().find(|c| c.tag_name().name() == name))
        .and_then(|n| n.text())
}

/// Extra vertical space for an inline image: (total, top_portion).
/// Total = effectExtent top+bottom + distT+distB.
/// Top = effectExtent.t + distT.
fn inline_extra_height(container: roxmltree::Node) -> (f32, f32) {
    let ee = wpd(container, "effectExtent");
    let ee_t = ee.map(|n| emu_attr(n, "t")).unwrap_or(0.0);
    let ee_b = ee.map(|n| emu_attr(n, "b")).unwrap_or(0.0);
    let dist_t = emu_attr(container, "distT");
    let dist_b = emu_attr(container, "distB");
    (ee_t + ee_b + dist_t + dist_b, ee_t + dist_t)
}

pub(super) fn extent_dimensions(container: roxmltree::Node) -> (f32, f32) {
    let extent = wpd(container, "extent");
    let cx = extent
        .and_then(|n| n.attribute("cx"))
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0);
    let cy = extent
        .and_then(|n| n.attribute("cy"))
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0);
    (emu_to_pts(cx), emu_to_pts(cy))
}

pub(super) fn image_dimensions(data: &[u8]) -> Option<(u32, u32, ImageFormat, u8)> {
    if data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8 {
        return parse_jpeg_dimensions(data);
    }

    if data.len() >= 24 && data[0..4] == [0x89, 0x50, 0x4E, 0x47] {
        let width = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
        let height = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
        return Some((width, height, ImageFormat::Png, 3));
    }

    if data.len() >= 26 && data[0] == b'B' && data[1] == b'M' {
        let width = u32::from_le_bytes([data[18], data[19], data[20], data[21]]);
        let height = u32::from_le_bytes([data[22], data[23], data[24], data[25]]);
        return Some((width, height, ImageFormat::Bmp, 3));
    }

    None
}

fn parse_jpeg_dimensions(data: &[u8]) -> Option<(u32, u32, ImageFormat, u8)> {
    let mut i = 2;
    while i + 4 < data.len() {
        if data[i] != 0xFF {
            return None;
        }
        let marker = data[i + 1];
        if marker == 0xD9 {
            break;
        }
        let len = u16::from_be_bytes([data[i + 2], data[i + 3]]) as usize;
        if matches!(marker, 0xC0 | 0xC1 | 0xC2) && i + 9 < data.len() {
            let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
            let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
            let components = data[i + 9];
            return Some((width, height, ImageFormat::Jpeg, components));
        }
        i += 2 + len;
    }
    None
}

/// Parse outline stroke from `pic:spPr/a:ln` inside a pic:pic element.
fn parse_pic_outline(container: roxmltree::Node) -> (Option<[u8; 3]>, f32) {
    let pic_ns = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    let pic_node = container
        .descendants()
        .find(|n| n.tag_name().name() == "pic" && n.tag_name().namespace() == Some(pic_ns));
    // pic:spPr uses the pic namespace, but a:ln inside it uses DML_NS
    let sp_pr = pic_node.and_then(|p| {
        p.children()
            .find(|c| c.tag_name().name() == "spPr" && c.tag_name().namespace() == Some(pic_ns))
    });
    let ln = sp_pr.and_then(|s| {
        s.children()
            .find(|c| c.tag_name().name() == "ln" && c.tag_name().namespace() == Some(DML_NS))
    });
    let Some(ln) = ln else {
        return (None, 0.0);
    };
    let width = ln
        .attribute("w")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_pts)
        .unwrap_or(0.75); // default 0.75pt
    let color = ln
        .descendants()
        .find(|n| n.tag_name().name() == "srgbClr" && n.tag_name().namespace() == Some(DML_NS))
        .and_then(|n| n.attribute("val"))
        .and_then(parse_hex_color);
    if color.is_some() {
        (color, width)
    } else {
        (None, 0.0)
    }
}

/// Parse drop shadow from `pic:spPr/a:effectLst/a:outerShdw`.
/// Returns color pre-blended with white at the shadow's alpha.
fn parse_pic_shadow(container: roxmltree::Node) -> Option<ImageShadow> {
    let pic_ns = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    let pic_node = container
        .descendants()
        .find(|n| n.tag_name().name() == "pic" && n.tag_name().namespace() == Some(pic_ns));
    let sp_pr = pic_node.and_then(|p| {
        p.children()
            .find(|c| c.tag_name().name() == "spPr" && c.tag_name().namespace() == Some(pic_ns))
    });
    let effect_lst = sp_pr.and_then(|s| {
        s.children()
            .find(|c| c.tag_name().name() == "effectLst" && c.tag_name().namespace() == Some(DML_NS))
    });
    let outer_shdw = effect_lst.and_then(|e| {
        e.children()
            .find(|c| c.tag_name().name() == "outerShdw" && c.tag_name().namespace() == Some(DML_NS))
    })?;

    let blur_radius = outer_shdw
        .attribute("blurRad")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_pts)
        .unwrap_or(0.0);
    let dist = outer_shdw
        .attribute("dist")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_pts)
        .unwrap_or(0.0);
    // Direction in 60000ths of a degree, clockwise from east
    let dir_deg = outer_shdw
        .attribute("dir")
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0)
        / 60000.0;
    let dir_rad = dir_deg.to_radians();
    let offset_x = dist * dir_rad.cos();
    let offset_y = dist * dir_rad.sin();

    let color_node = outer_shdw
        .descendants()
        .find(|n| n.tag_name().name() == "srgbClr" && n.tag_name().namespace() == Some(DML_NS));
    let rgb = color_node
        .and_then(|n| n.attribute("val"))
        .and_then(parse_hex_color)
        .unwrap_or([0, 0, 0]);
    // Alpha in thousandths of percent (e.g. 65000 = 65%)
    let alpha = color_node
        .and_then(|n| {
            n.children()
                .find(|c| c.tag_name().name() == "alpha" && c.tag_name().namespace() == Some(DML_NS))
        })
        .and_then(|a| a.attribute("val"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 100000.0)
        .unwrap_or(1.0);

    // Pre-blend with white: blended = alpha * color + (1-alpha) * 255
    let blended = [
        (alpha * rgb[0] as f32 + (1.0 - alpha) * 255.0) as u8,
        (alpha * rgb[1] as f32 + (1.0 - alpha) * 255.0) as u8,
        (alpha * rgb[2] as f32 + (1.0 - alpha) * 255.0) as u8,
    ];

    Some(ImageShadow {
        offset_x,
        offset_y,
        blur_radius,
        color: blended,
    })
}

pub(super) fn read_image_from_zip<R: Read + std::io::Seek>(
    embed_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    display_w: f32,
    display_h: f32,
) -> Option<EmbeddedImage> {
    read_image_from_zip_extra(embed_id, rels, zip, display_w, display_h, 0.0, 0.0)
}

pub(super) fn read_image_from_zip_extra<R: Read + std::io::Seek>(
    embed_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    display_w: f32,
    display_h: f32,
    layout_extra_height: f32,
    layout_extra_top: f32,
) -> Option<EmbeddedImage> {
    let target = rels.get(embed_id)?;
    let zip_path = target
        .strip_prefix('/')
        .map(String::from)
        .unwrap_or_else(|| format!("word/{}", target));
    let mut entry = zip.by_name(&zip_path).ok()?;
    let mut data = Vec::new();
    entry.read_to_end(&mut data).ok()?;
    let (pw, ph, fmt, components) = image_dimensions(&data)?;
    Some(EmbeddedImage {
        data: std::sync::Arc::new(data),
        format: fmt,
        pixel_width: pw,
        pixel_height: ph,
        display_width: display_w,
        display_height: display_h,
        jpeg_components: components,
        layout_extra_height,
        layout_extra_top,
        stroke_color: None,
        stroke_width: 0.0,
        shadow: None,
    })
}

pub(super) fn find_blip_embed<'a>(container: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    container
        .descendants()
        .find(|n| n.tag_name().name() == "blip" && n.tag_name().namespace() == Some(DML_NS))
        .and_then(|n| n.attribute((REL_NS, "embed")))
}

pub(super) struct DrawingInfo {
    pub(super) height: f32,
    pub(super) image: Option<EmbeddedImage>,
    pub(super) floating_images: Vec<FloatingImage>,
}

pub(super) fn parse_anchor_position(
    container: roxmltree::Node,
) -> (
    HorizontalPosition,
    HRelativeFrom,
    VerticalPosition,
    VRelativeFrom,
) {
    let pos_h = wpd(container, "positionH");
    let h_relative = match pos_h.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => HRelativeFrom::Page,
        Some("margin") => HRelativeFrom::Margin,
        _ => HRelativeFrom::Column,
    };
    let h_position = if let Some(text) = wpd_child_text(pos_h, "align") {
        match text {
            "center" => HorizontalPosition::AlignCenter,
            "right" => HorizontalPosition::AlignRight,
            _ => HorizontalPosition::AlignLeft,
        }
    } else if let Some(text) = wpd_child_text(pos_h, "posOffset") {
        HorizontalPosition::Offset(parse_emu_text(Some(text)))
    } else {
        HorizontalPosition::AlignLeft
    };

    let pos_v = wpd(container, "positionV");
    let v_relative = match pos_v.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => VRelativeFrom::Page,
        Some("margin") => VRelativeFrom::Margin,
        Some("topMargin") => VRelativeFrom::TopMargin,
        _ => VRelativeFrom::Paragraph,
    };
    let v_position = if let Some(text) = wpd_child_text(pos_v, "align") {
        match text {
            "bottom" => VerticalPosition::AlignBottom,
            "center" => VerticalPosition::AlignCenter,
            _ => VerticalPosition::AlignTop,
        }
    } else if let Some(text) = wpd_child_text(pos_v, "posOffset") {
        VerticalPosition::Offset(parse_emu_text(Some(text)))
    } else {
        VerticalPosition::Offset(0.0)
    };

    (h_position, h_relative, v_position, v_relative)
}

pub(super) fn parse_wrap_type(
    container: roxmltree::Node,
) -> (WrapType, WrapText, Option<Vec<(i32, i32)>>) {
    for child in container.children() {
        if child.tag_name().namespace() != Some(WPD_NS) {
            continue;
        }
        let wrap_text = match child.attribute("wrapText") {
            Some("bothSides") => WrapText::BothSides,
            Some("left") => WrapText::Left,
            Some("right") => WrapText::Right,
            Some("largest") => WrapText::Largest,
            _ => WrapText::BothSides,
        };
        match child.tag_name().name() {
            "wrapSquare" => return (WrapType::Square, wrap_text, None),
            "wrapTight" | "wrapThrough" => {
                let polygon = parse_wrap_polygon(child);
                let wt = if child.tag_name().name() == "wrapTight" {
                    WrapType::Tight
                } else {
                    WrapType::Through
                };
                return (wt, wrap_text, polygon);
            }
            "wrapTopAndBottom" => {
                return (WrapType::TopAndBottom, WrapText::BothSides, None)
            }
            "wrapNone" => return (WrapType::None, WrapText::BothSides, None),
            _ => {}
        }
    }
    (WrapType::None, WrapText::BothSides, None)
}

fn parse_wrap_polygon(wrap_elem: roxmltree::Node) -> Option<Vec<(i32, i32)>> {
    let poly = wrap_elem
        .children()
        .find(|c| c.tag_name().name() == "wrapPolygon" && c.tag_name().namespace() == Some(WPD_NS))?;
    let mut vertices = Vec::new();
    for child in poly.children() {
        if child.tag_name().namespace() != Some(WPD_NS) {
            continue;
        }
        match child.tag_name().name() {
            "start" | "lineTo" => {
                let x = child
                    .attribute("x")
                    .and_then(|v| v.parse::<i32>().ok())
                    .unwrap_or(0);
                let y = child
                    .attribute("y")
                    .and_then(|v| v.parse::<i32>().ok())
                    .unwrap_or(0);
                vertices.push((x, y));
            }
            _ => {}
        }
    }
    if vertices.is_empty() {
        None
    } else {
        Some(vertices)
    }
}

pub(super) enum RunDrawingResult {
    Inline(EmbeddedImage),
    Floating(FloatingImage),
    TextBox(crate::model::Textbox),
    Connector(crate::model::ConnectorShape),
    Chart(InlineChart),
    SmartArt(SmartArtDiagram),
}

fn is_wpd_drawing(node: roxmltree::Node, expected: &str) -> bool {
    node.tag_name().name() == expected && node.tag_name().namespace() == Some(WPD_NS)
}

pub(super) fn parse_run_drawing<R: Read + std::io::Seek>(
    drawing_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> Option<RunDrawingResult> {
    for container in drawing_node.children() {
        let is_inline = is_wpd_drawing(container, "inline");
        let is_anchor = is_wpd_drawing(container, "anchor");
        if !is_inline && !is_anchor {
            continue;
        }

        let (display_w, display_h) = extent_dimensions(container);

        if is_anchor {
            if let Some(wsp) =
                parse_textbox_from_wsp(container, rels, zip, styles, theme, numbering)
            {
                let (h_position, h_relative, v_pos, v_relative) = parse_anchor_position(container);
                let v_offset = match v_pos {
                    VerticalPosition::Offset(o) => o,
                    _ => 0.0,
                };
                let (wrap_type, _, _) = parse_wrap_type(container);
                let behind_doc = container.attribute("behindDoc") == Some("1");
                return Some(RunDrawingResult::TextBox(crate::model::Textbox {
                    paragraphs: wsp.paragraphs,
                    width_pt: display_w,
                    height_pt: display_h,
                    h_position,
                    h_relative_from: h_relative,
                    v_offset_pt: v_offset,
                    v_relative_from: v_relative,
                    fill: wsp.fill,
                    shape_type: wsp.shape_type,
                    stroke_color: wsp.stroke_color,
                    stroke_width: wsp.stroke_width,
                    text_anchor: wsp.text_anchor,
                    margin_left: wsp.margin_left,
                    margin_right: wsp.margin_right,
                    margin_top: wsp.margin_top,
                    margin_bottom: wsp.margin_bottom,
                    wrap_type,
                    dist_top: emu_attr(container, "distT"),
                    dist_bottom: emu_attr(container, "distB"),
                    behind_doc,
                    no_text_wrap: wsp.no_text_wrap,
                    is_wordart: wsp.is_wordart,
                    text_warp: wsp.text_warp,
                    auto_fit: wsp.auto_fit,
                }));
            }
            if let Some(conn) = parse_connector_from_wsp(container, theme) {
                return Some(RunDrawingResult::Connector(conn));
            }
            if let Some(embed_id) = find_blip_embed(container) {
                if let Some(mut img) = read_image_from_zip(embed_id, rels, zip, display_w, display_h) {
                    let (stroke_color, stroke_width) = parse_pic_outline(container);
                    img.stroke_color = stroke_color;
                    img.stroke_width = stroke_width;
                    img.shadow = parse_pic_shadow(container);
                    let (h_position, h_relative, v_position, v_relative) =
                        parse_anchor_position(container);
                    let (wrap_type, wrap_text, wrap_polygon) = parse_wrap_type(container);
                    let behind_doc = container.attribute("behindDoc") == Some("1");
                    return Some(RunDrawingResult::Floating(FloatingImage {
                        image: img,
                        h_position,
                        h_relative_from: h_relative,
                        v_position,
                        v_relative_from: v_relative,
                        wrap_type,
                        wrap_text,
                        wrap_polygon,
                        behind_doc,
                        dist_top: emu_attr(container, "distT"),
                        dist_bottom: emu_attr(container, "distB"),
                        dist_left: emu_attr(container, "distL"),
                        dist_right: emu_attr(container, "distR"),
                    }));
                }
            }
            continue;
        }

        // Inline textbox: wp:inline containing wps:wsp with text content
        if let Some(wsp) =
            parse_textbox_from_wsp(container, rels, zip, styles, theme, numbering)
        {
            // Treat inline textbox as a floating textbox at paragraph position
            // with TopAndBottom wrap so it acts as a block element
            return Some(RunDrawingResult::TextBox(crate::model::Textbox {
                paragraphs: wsp.paragraphs,
                width_pt: display_w,
                height_pt: display_h,
                h_position: HorizontalPosition::Offset(0.0),
                h_relative_from: HRelativeFrom::Column,
                v_offset_pt: 0.0,
                v_relative_from: crate::model::VRelativeFrom::Paragraph,
                fill: wsp.fill,
                shape_type: wsp.shape_type,
                stroke_color: wsp.stroke_color,
                stroke_width: wsp.stroke_width,
                text_anchor: wsp.text_anchor,
                margin_left: wsp.margin_left,
                margin_right: wsp.margin_right,
                margin_top: wsp.margin_top,
                margin_bottom: wsp.margin_bottom,
                wrap_type: crate::model::WrapType::TopAndBottom,
                dist_top: 0.0,
                dist_bottom: 0.0,
                behind_doc: false,
                no_text_wrap: wsp.no_text_wrap,
                is_wordart: wsp.is_wordart,
                text_warp: wsp.text_warp,
                auto_fit: wsp.auto_fit,
            }));
        }

        if let Some(embed_id) = find_blip_embed(container) {
            let (extra_h, extra_top) = inline_extra_height(container);
            if let Some(mut img) =
                read_image_from_zip_extra(embed_id, rels, zip, display_w, display_h, extra_h, extra_top)
            {
                let (stroke_color, stroke_width) = parse_pic_outline(container);
                img.stroke_color = stroke_color;
                img.stroke_width = stroke_width;
                img.shadow = parse_pic_shadow(container);
                return Some(RunDrawingResult::Inline(img));
            }
        }

        if let Some(chart_rid) = find_chart_ref(container) {
            let accent_colors: Vec<[u8; 3]> = (1..=6)
                .filter_map(|i| theme.colors.get(&format!("accent{i}")).copied())
                .collect();
            if let Some(ic) =
                parse_chart_from_zip(chart_rid, rels, zip, display_w, display_h, accent_colors)
            {
                return Some(RunDrawingResult::Chart(ic));
            }
        }

        if display_h > 0.0 && has_diagram_ref(container) {
            let diagram = parse_smartart_drawing(rels, zip, theme, display_w, display_h);
            return Some(RunDrawingResult::SmartArt(diagram));
        }
    }
    None
}

fn find_chart_ref<'a>(container: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    container
        .descendants()
        .find(|n| {
            n.tag_name().name() == "graphicData"
                && n.tag_name().namespace() == Some(DML_NS)
                && n.attribute("uri") == Some(CHART_URI)
        })
        .and_then(|gd| {
            gd.children()
                .find(|n| n.tag_name().name() == "chart")
                .and_then(|c| c.attribute((REL_NS, "id")))
        })
}

pub(super) fn compute_drawing_info<R: Read + std::io::Seek>(
    para_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> DrawingInfo {
    let mut max_height: f32 = 0.0;
    let mut image: Option<EmbeddedImage> = None;

    for child in para_node.children() {
        let is_wml = child.tag_name().namespace() == Some(WML_NS);
        let drawing_node = match child.tag_name().name() {
            "drawing" if is_wml => Some(child),
            "r" if is_wml => wml(child, "drawing"),
            _ => None,
        };

        let Some(drawing) = drawing_node else {
            continue;
        };
        for container in drawing.children() {
            if !is_wpd_drawing(container, "inline") {
                continue;
            }

            let (display_w, display_h) = extent_dimensions(container);
            let (extra_h, extra_top) = inline_extra_height(container);
            max_height = max_height.max(display_h + extra_h);

            if image.is_none() {
                if let Some(embed_id) = find_blip_embed(container) {
                    image =
                        read_image_from_zip_extra(embed_id, rels, zip, display_w, display_h, extra_h, extra_top);
                    if let Some(ref mut img) = image {
                        let (stroke_color, stroke_width) = parse_pic_outline(container);
                        img.stroke_color = stroke_color;
                        img.stroke_width = stroke_width;
                        img.shadow = parse_pic_shadow(container);
                    }
                }
            }
        }
    }
    DrawingInfo {
        height: max_height,
        image,
        floating_images: Vec::new(),
    }
}
