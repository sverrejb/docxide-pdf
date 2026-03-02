use std::collections::HashMap;
use std::io::Read;

use crate::model::{EmbeddedImage, FloatingImage, HorizontalPosition, ImageFormat};

use super::{DML_NS, REL_NS, WML_NS, WPD_NS, wml};
use super::styles::{StylesInfo, ThemeFonts};
use super::numbering::NumberingInfo;
use super::textbox::parse_textbox_from_wsp;

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
        data,
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
    let pos_h = container.children().find(|n| {
        n.tag_name().name() == "positionH" && n.tag_name().namespace() == Some(WPD_NS)
    });
    let h_relative = match pos_h.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => "page",
        Some("margin") => "margin",
        _ => "column",
    };
    let h_position = if let Some(align_node) = pos_h.and_then(|n| {
        n.children().find(|c| c.tag_name().name() == "align")
    }) {
        match align_node.text().unwrap_or("") {
            "center" => HorizontalPosition::AlignCenter,
            "right" => HorizontalPosition::AlignRight,
            _ => HorizontalPosition::AlignLeft,
        }
    } else if let Some(offset_node) = pos_h.and_then(|n| {
        n.children().find(|c| c.tag_name().name() == "posOffset")
    }) {
        let emu = offset_node
            .text()
            .unwrap_or("0")
            .parse::<f32>()
            .unwrap_or(0.0);
        HorizontalPosition::Offset(emu / 12700.0)
    } else {
        HorizontalPosition::AlignLeft
    };

    let pos_v = container.children().find(|n| {
        n.tag_name().name() == "positionV" && n.tag_name().namespace() == Some(WPD_NS)
    });
    let v_relative = match pos_v.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => "page",
        Some("margin") => "margin",
        Some("topMargin") => "topMargin",
        _ => "paragraph",
    };
    let v_offset = if let Some(offset_node) = pos_v.and_then(|n| {
        n.children().find(|c| c.tag_name().name() == "posOffset")
    }) {
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

pub(super) enum RunDrawingResult {
    Inline(EmbeddedImage),
    Floating(FloatingImage),
    TextBox(crate::model::Textbox),
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

        let extent = container.children().find(|n| {
            n.tag_name().name() == "extent" && n.tag_name().namespace() == Some(WPD_NS)
        });
        let cx = extent
            .and_then(|n| n.attribute("cx"))
            .and_then(|v| v.parse::<f32>().ok())
            .unwrap_or(0.0);
        let cy = extent
            .and_then(|n| n.attribute("cy"))
            .and_then(|v| v.parse::<f32>().ok())
            .unwrap_or(0.0);
        let display_w = cx / 12700.0;
        let display_h = cy / 12700.0;

        if name == "anchor" {
            // Check for textbox/shape content (any wrap mode)
            if let Some(wsp) =
                parse_textbox_from_wsp(container, rels, zip, styles, theme, numbering)
            {
                let (h_position, h_relative, v_offset, v_relative) =
                    parse_anchor_position(container);
                return Some(RunDrawingResult::TextBox(crate::model::Textbox {
                    paragraphs: wsp.paragraphs,
                    width_pt: display_w,
                    height_pt: display_h,
                    h_position,
                    h_relative_from: h_relative,
                    v_offset_pt: v_offset,
                    v_relative_from: v_relative,
                    fill_color: wsp.fill_color,
                    margin_left: wsp.margin_left,
                    margin_right: wsp.margin_right,
                    margin_top: wsp.margin_top,
                    margin_bottom: wsp.margin_bottom,
                }));
            }
            if let Some(embed_id) = find_blip_embed(container) {
                if let Some(img) = read_image_from_zip(embed_id, rels, zip, display_w, display_h) {
                    let (h_position, h_relative, v_offset, v_relative) =
                        parse_anchor_position(container);
                    return Some(RunDrawingResult::Floating(FloatingImage {
                        image: img,
                        h_position,
                        h_relative_from: h_relative,
                        v_offset_pt: v_offset,
                        v_relative_from: v_relative,
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
    }
    None
}

pub(super) fn compute_drawing_info<R: Read + std::io::Seek>(
    para_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> DrawingInfo {
    let mut max_height: f32 = 0.0;
    let mut image: Option<EmbeddedImage> = None;
    let floating_images: Vec<FloatingImage> = Vec::new();

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

            let extent = container.children().find(|n| {
                n.tag_name().name() == "extent" && n.tag_name().namespace() == Some(WPD_NS)
            });
            let cx = extent
                .and_then(|n| n.attribute("cx"))
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0);
            let cy = extent
                .and_then(|n| n.attribute("cy"))
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0);
            let display_w = cx / 12700.0;
            let display_h = cy / 12700.0;

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
        floating_images,
    }
}
