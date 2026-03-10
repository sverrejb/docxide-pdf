use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    EmbeddedImage, FloatingImage, HorizontalPosition, ImageFormat, InlineChart, ShapeType,
    SmartArtDiagram, SmartArtShape, WrapType,
};

use super::charts::parse_chart_from_zip;
use super::numbering::NumberingInfo;
use super::styles::{StylesInfo, ThemeFonts};
use super::textbox::{parse_solid_fill, parse_textbox_from_wsp};
use super::{DML_NS, REL_NS, WML_NS, WPD_NS, read_zip_text, wml};

const CHART_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DIAGRAM_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

/// Extract display dimensions (in points) from a wp:inline or wp:anchor element.
pub(super) fn extent_dimensions(container: roxmltree::Node) -> (f32, f32) {
    let extent = container
        .children()
        .find(|n| n.tag_name().name() == "extent" && n.tag_name().namespace() == Some(WPD_NS));
    let cx = extent
        .and_then(|n| n.attribute("cx"))
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0);
    let cy = extent
        .and_then(|n| n.attribute("cy"))
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0);
    (cx / 12700.0, cy / 12700.0)
}

pub(super) fn image_dimensions(data: &[u8]) -> Option<(u32, u32, ImageFormat)> {
    // JPEG: starts with FF D8
    if data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8 {
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
            if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2) && i + 9 < data.len() {
                let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
                let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
                return Some((width, height, ImageFormat::Jpeg));
            }
            i += 2 + len;
        }
        return None;
    }

    // PNG: starts with 89 50 4E 47, dimensions in IHDR chunk at bytes 16-23
    if data.len() >= 24 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
    {
        let width = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
        let height = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
        return Some((width, height, ImageFormat::Png));
    }

    None
}

pub(super) fn read_image_from_zip<R: Read + std::io::Seek>(
    embed_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    display_w: f32,
    display_h: f32,
) -> Option<EmbeddedImage> {
    let target = rels.get(embed_id)?;
    let zip_path = target
        .strip_prefix('/')
        .map(String::from)
        .unwrap_or_else(|| format!("word/{}", target));
    let mut entry = zip.by_name(&zip_path).ok()?;
    let mut data = Vec::new();
    entry.read_to_end(&mut data).ok()?;
    let (pw, ph, fmt) = image_dimensions(&data)?;
    Some(EmbeddedImage {
        data: std::sync::Arc::new(data),
        format: fmt,
        pixel_width: pw,
        pixel_height: ph,
        display_width: display_w,
        display_height: display_h,
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
) -> (HorizontalPosition, &'static str, f32, &'static str) {
    let pos_h = container
        .children()
        .find(|n| n.tag_name().name() == "positionH" && n.tag_name().namespace() == Some(WPD_NS));
    let h_relative = match pos_h.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => "page",
        Some("margin") => "margin",
        _ => "column",
    };
    let h_position = if let Some(align_node) =
        pos_h.and_then(|n| n.children().find(|c| c.tag_name().name() == "align"))
    {
        match align_node.text().unwrap_or("") {
            "center" => HorizontalPosition::AlignCenter,
            "right" => HorizontalPosition::AlignRight,
            _ => HorizontalPosition::AlignLeft,
        }
    } else if let Some(offset_node) =
        pos_h.and_then(|n| n.children().find(|c| c.tag_name().name() == "posOffset"))
    {
        let emu = offset_node
            .text()
            .unwrap_or("0")
            .parse::<f32>()
            .unwrap_or(0.0);
        HorizontalPosition::Offset(emu / 12700.0)
    } else {
        HorizontalPosition::AlignLeft
    };

    let pos_v = container
        .children()
        .find(|n| n.tag_name().name() == "positionV" && n.tag_name().namespace() == Some(WPD_NS));
    let v_relative = match pos_v.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => "page",
        Some("margin") => "margin",
        Some("topMargin") => "topMargin",
        _ => "paragraph",
    };
    let v_offset = if let Some(offset_node) =
        pos_v.and_then(|n| n.children().find(|c| c.tag_name().name() == "posOffset"))
    {
        offset_node
            .text()
            .unwrap_or("0")
            .parse::<f32>()
            .unwrap_or(0.0)
            / 12700.0
    } else {
        0.0
    };

    (h_position, h_relative, v_offset, v_relative)
}

pub(super) fn parse_wrap_type(container: roxmltree::Node) -> WrapType {
    for child in container.children() {
        if child.tag_name().namespace() != Some(WPD_NS) {
            continue;
        }
        match child.tag_name().name() {
            "wrapSquare" => return WrapType::Square,
            "wrapTight" => return WrapType::Tight,
            "wrapThrough" => return WrapType::Through,
            "wrapTopAndBottom" => return WrapType::TopAndBottom,
            "wrapNone" => return WrapType::None,
            _ => {}
        }
    }
    WrapType::None
}

#[allow(dead_code)]
pub(super) enum RunDrawingResult {
    Inline(EmbeddedImage),
    Floating(FloatingImage),
    TextBox(crate::model::Textbox),
    Chart(InlineChart),
    SmartArt(SmartArtDiagram),
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
        let name = container.tag_name().name();
        if name != "inline" && name != "anchor" {
            continue;
        }
        if container.tag_name().namespace() != Some(WPD_NS) {
            continue;
        }

        let (display_w, display_h) = extent_dimensions(container);

        if name == "anchor" {
            if let Some(wsp) =
                parse_textbox_from_wsp(container, rels, zip, styles, theme, numbering)
            {
                let (h_position, h_relative, v_offset, v_relative) =
                    parse_anchor_position(container);
                let wrap_type = parse_wrap_type(container);
                let dist_top = container
                    .attribute("distT")
                    .and_then(|v| v.parse::<f32>().ok())
                    .unwrap_or(0.0)
                    / 12700.0;
                let dist_bottom = container
                    .attribute("distB")
                    .and_then(|v| v.parse::<f32>().ok())
                    .unwrap_or(0.0)
                    / 12700.0;
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
                    margin_left: wsp.margin_left,
                    margin_right: wsp.margin_right,
                    margin_top: wsp.margin_top,
                    margin_bottom: wsp.margin_bottom,
                    wrap_type,
                    dist_top,
                    dist_bottom,
                }));
            }
            if let Some(embed_id) = find_blip_embed(container) {
                if let Some(img) = read_image_from_zip(embed_id, rels, zip, display_w, display_h) {
                    let (h_position, h_relative, v_offset, v_relative) =
                        parse_anchor_position(container);
                    let wrap_type = parse_wrap_type(container);
                    return Some(RunDrawingResult::Floating(FloatingImage {
                        image: img,
                        h_position,
                        h_relative_from: h_relative,
                        v_offset_pt: v_offset,
                        v_relative_from: v_relative,
                        wrap_type,
                    }));
                }
            }
            continue;
        }

        if let Some(embed_id) = find_blip_embed(container) {
            if let Some(img) = read_image_from_zip(embed_id, rels, zip, display_w, display_h) {
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
            let diagram =
                parse_smartart_drawing(rels, zip, theme, display_w, display_h);
            return Some(RunDrawingResult::SmartArt(diagram));
        }
    }
    None
}

fn has_diagram_ref(container: roxmltree::Node) -> bool {
    container.descendants().any(|n| {
        n.tag_name().name() == "graphicData"
            && n.tag_name().namespace() == Some(DML_NS)
            && n.attribute("uri") == Some(DIAGRAM_URI)
    })
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
            let name = container.tag_name().name();
            if name != "inline" && name != "anchor" {
                continue;
            }
            if container.tag_name().namespace() != Some(WPD_NS) {
                continue;
            }

            let (display_w, display_h) = extent_dimensions(container);

            // Anchored images are handled by parse_runs() — skip them here
            // to avoid duplication. Only process inline drawings.
            if name == "anchor" {
                continue;
            }

            max_height = max_height.max(display_h);

            if image.is_none() {
                if let Some(embed_id) = find_blip_embed(container) {
                    image = read_image_from_zip(embed_id, rels, zip, display_w, display_h);
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

const DSP_NS: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";

fn parse_smartart_drawing<R: Read + std::io::Seek>(
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    theme: &ThemeFonts,
    display_w: f32,
    display_h: f32,
) -> SmartArtDiagram {
    let mut shapes = Vec::new();

    let drawing_target = rels
        .values()
        .find(|t| t.contains("diagrams/drawing"));

    if let Some(target) = drawing_target {
        let zip_path = target
            .strip_prefix('/')
            .map(String::from)
            .unwrap_or_else(|| format!("word/{}", target));

        if let Some(xml) = read_zip_text(zip, &zip_path) {
            if let Ok(doc) = roxmltree::Document::parse(&xml) {
                let sp_tree = doc
                    .root()
                    .children()
                    .find(|n| n.tag_name().name() == "drawing" && n.tag_name().namespace() == Some(DSP_NS))
                    .and_then(|d| {
                        d.children().find(|n| {
                            n.tag_name().name() == "spTree"
                                && n.tag_name().namespace() == Some(DSP_NS)
                        })
                    });

                if let Some(tree) = sp_tree {
                    for sp in tree.children().filter(|n| {
                        n.tag_name().name() == "sp"
                            && n.tag_name().namespace() == Some(DSP_NS)
                    }) {
                        if let Some(shape) = parse_dsp_shape(sp, theme) {
                            shapes.push(shape);
                        }
                    }
                }
            }
        }
    }

    SmartArtDiagram {
        display_width: display_w,
        display_height: display_h,
        shapes,
    }
}

fn parse_dsp_shape(sp: roxmltree::Node, theme: &ThemeFonts) -> Option<SmartArtShape> {
    let sp_pr = sp.children().find(|n| {
        n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(DSP_NS)
    })?;

    let xfrm = sp_pr.children().find(|n| {
        n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(DML_NS)
    })?;

    let off = xfrm.children().find(|n| {
        n.tag_name().name() == "off" && n.tag_name().namespace() == Some(DML_NS)
    })?;
    let ext = xfrm.children().find(|n| {
        n.tag_name().name() == "ext" && n.tag_name().namespace() == Some(DML_NS)
    })?;

    let x = off.attribute("x").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
    let y = off.attribute("y").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
    let w = ext.attribute("cx").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
    let h = ext.attribute("cy").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;

    let has_no_fill = sp_pr.children().any(|n| {
        n.tag_name().name() == "noFill" && n.tag_name().namespace() == Some(DML_NS)
    });

    let fill = if has_no_fill {
        None
    } else {
        parse_solid_fill(sp_pr, theme)
    };

    let (text, font_size) = parse_dsp_text(sp);

    if fill.is_none() && text.is_empty() {
        return None;
    }

    let prst = sp_pr
        .children()
        .find(|n| {
            n.tag_name().name() == "prstGeom" && n.tag_name().namespace() == Some(DML_NS)
        })
        .and_then(|n| n.attribute("prst"))
        .unwrap_or("rect");

    let shape_type = match prst {
        "ellipse" => ShapeType::Ellipse,
        _ => ShapeType::Rect,
    };

    Some(SmartArtShape {
        x: x as f32,
        y: y as f32,
        width: w as f32,
        height: h as f32,
        shape_type,
        fill,
        text,
        font_size,
    })
}

fn parse_dsp_text(sp: roxmltree::Node) -> (String, f32) {
    let tx_body = sp.children().find(|n| {
        n.tag_name().name() == "txBody" && n.tag_name().namespace() == Some(DSP_NS)
    });

    let Some(body) = tx_body else {
        return (String::new(), 0.0);
    };

    let mut lines = Vec::new();
    let mut font_size = 0.0_f32;

    for p in body.children().filter(|n| {
        n.tag_name().name() == "p" && n.tag_name().namespace() == Some(DML_NS)
    }) {
        let mut line_text = String::new();
        for r in p.children().filter(|n| {
            n.tag_name().name() == "r" && n.tag_name().namespace() == Some(DML_NS)
        }) {
            if let Some(rpr) = r.children().find(|n| {
                n.tag_name().name() == "rPr" && n.tag_name().namespace() == Some(DML_NS)
            }) {
                if let Some(sz) = rpr.attribute("sz").and_then(|v| v.parse::<f32>().ok()) {
                    if font_size == 0.0 {
                        font_size = sz / 100.0;
                    }
                }
            }
            if let Some(t) = r.children().find(|n| {
                n.tag_name().name() == "t" && n.tag_name().namespace() == Some(DML_NS)
            }) {
                if let Some(text) = t.text() {
                    line_text.push_str(text);
                }
            }
        }
        if !line_text.is_empty() {
            lines.push(line_text);
        }
    }

    (lines.join("\n"), font_size)
}
