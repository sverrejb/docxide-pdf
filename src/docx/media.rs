//! Media parts: reading pictures from the package and decoding them.

use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{EmbeddedImage, ImageFormat};

use super::{DML_NS, REL_NS, emu_attr_opt, frac_attr, part_path, read_zip_bytes};

/// GIF and TIFF pictures are re-encoded as PNG for the PDF writer.
fn gif_or_tiff_to_png(data: &[u8]) -> Option<Vec<u8>> {
    let fmt = match image::guess_format(data).ok()? {
        f @ (image::ImageFormat::Gif | image::ImageFormat::Tiff) => f,
        _ => return None,
    };
    encode_png(image::load_from_memory_with_format(data, fmt).ok()?)
}

/// `a:tile`: the picture repeated at `sx`/`sy` of its size (96 dpi) from the
/// frame's top-left corner, offset by `tx`/`ty`, as a PNG of the frame. None
/// (the picture stretched instead) past 4096 pixels a side.
/// ponytail: `flip` and `algn` other than the top-left are ignored; Word
/// writes a tiling pattern, which would keep the tile's size in the file.
pub(super) fn tile_picture(
    data: &[u8],
    tile: roxmltree::Node,
    frame_w: f32,
    frame_h: f32,
) -> Option<(Vec<u8>, u32, u32)> {
    const MAX_SIDE: f32 = 4096.0;
    let pic = image::load_from_memory(data).ok()?.into_rgba8();
    // Pixels per point at the tile's size, so each tile keeps its pixels.
    let px = |axis| 96.0 / 72.0 / frac_attr(tile, axis).unwrap_or(1.0);
    let (w, h) = (frame_w * px("sx"), frame_h * px("sy"));
    if !(1.0..=MAX_SIDE).contains(&w) || !(1.0..=MAX_SIDE).contains(&h) {
        return None;
    }
    let offset = |attr, axis, size: u32| {
        let o = (emu_attr_opt(tile, attr).unwrap_or(0.0) * px(axis)) as i64;
        o.rem_euclid(size as i64) - size as i64
    };
    let mut canvas = image::RgbaImage::new(w.ceil() as u32, h.ceil() as u32);
    let (tw, th) = (pic.width() as usize, pic.height() as usize);
    for y in (offset("ty", "sy", pic.height())..canvas.height() as i64).step_by(th) {
        for x in (offset("tx", "sx", pic.width())..canvas.width() as i64).step_by(tw) {
            image::imageops::replace(&mut canvas, &pic, x, y);
        }
    }
    let (cw, ch) = canvas.dimensions();
    Some((encode_png(image::DynamicImage::ImageRgba8(canvas))?, cw, ch))
}

fn encode_png(img: image::DynamicImage) -> Option<Vec<u8>> {
    let mut png = Vec::new();
    img.write_to(&mut std::io::Cursor::new(&mut png), image::ImageFormat::Png)
        .ok()?;
    Some(png)
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
        // A negative height marks a top-down DIB (common in EMF-wrapped bitmaps).
        let height = i32::from_le_bytes([data[22], data[23], data[24], data[25]]).unsigned_abs();
        return Some((width, height, ImageFormat::Bmp, 3));
    }

    if super::emf::is_emf(data) {
        let (bw, bh) = super::emf::parse_header(data)?.bounds_size();
        // Pixel dimensions don't strictly apply to vector EMFs — use the
        // bounds-rect logical size so downstream code that expects them gets
        // a sensible aspect ratio.
        let pw = bw.max(1) as u32;
        let ph = bh.max(1) as u32;
        return Some((pw, ph, ImageFormat::Emf, 0));
    }

    if super::svg::looks_like_svg(data) {
        let drawing = super::svg::parse(data)?;
        let px = |v: f32| v.round().max(1.0) as u32;
        return Some((
            px(drawing.view_box[2]),
            px(drawing.view_box[3]),
            ImageFormat::Svg,
            0,
        ));
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
        if matches!(marker, 0xC0..=0xC2) && i + 9 < data.len() {
            let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
            let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
            let components = data[i + 9];
            return Some((width, height, ImageFormat::Jpeg, components));
        }
        i += 2 + len;
    }
    None
}

#[derive(Default)]
pub(super) struct MediaParts {
    parts: HashMap<String, Option<EmbeddedImage>>,
    pictures: usize,
}

pub(super) fn read_image_from_zip<R: Read + Seek>(
    embed_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    media: &mut MediaParts,
    display_w: f32,
    display_h: f32,
) -> Option<EmbeddedImage> {
    let part = media
        .parts
        .entry(part_path(rels.get(embed_id)?))
        .or_insert_with_key(|path| read_media_part(zip, path))
        .clone()?;
    media.pictures += 1;
    Some(EmbeddedImage {
        id: media.pictures,
        display_width: display_w,
        display_height: display_h,
        ..part
    })
}

fn read_media_part<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Option<EmbeddedImage> {
    let mut data = read_zip_bytes(zip, path)?;
    if super::wmf::is_wmf(&data) {
        data = match super::wmf::embedded_emf(&data) {
            Some(emf) => emf,
            None => super::wmf::wmf_to_raster(&data).or_else(|| super::wmf::wmf_to_emf(&data))?,
        };
    }
    if let Some(bmp) = super::emf::emf_to_raster(&data) {
        data = bmp;
    } else if let Some(png) = gif_or_tiff_to_png(&data) {
        data = png;
    }
    let (pw, ph, fmt, components) = image_dimensions(&data)?;
    Some(EmbeddedImage {
        id: 0,
        is_ole_preview: false,
        alt: None,
        decorative: false,
        data: std::sync::Arc::new(data),
        format: fmt,
        pixel_width: pw,
        pixel_height: ph,
        display_width: 0.0,
        display_height: 0.0,
        trimmed_width: 0.0,
        jpeg_components: components,
        layout_extra_height: 0.0,
        layout_extra_top: 0.0,
        rotation_deg: 0.0,
        stroke_color: None,
        stroke_width: 0.0,
        shadow: None,
        soft_edge: None,
        glow: None,
        inner_shadow: None,
        reflection: None,
        clip_geometry: None,
        src_rect: None,
        lum: None,
        recolor: None,
        alpha: None,
    })
}

pub(super) fn find_blip<'a>(container: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    container
        .descendants()
        .find(|n| n.has_tag_name((DML_NS, "blip")))
}

pub(super) fn find_blip_embed<'a>(container: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    find_blip(container)?.attribute((REL_NS, "embed"))
}

const ASVG_NS: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

/// The picture's image parts, vector first: the `asvg:svgBlip` extension (Office
/// 2016+), then the `a:blip` raster that doubles as its fallback. Readers try
/// them in order, so an SVG we cannot draw still leaves the PNG.
pub(super) fn find_blip_embeds<'a>(
    container: roxmltree::Node<'a, 'a>,
) -> impl Iterator<Item = &'a str> {
    let blip = find_blip(container);
    let svg = blip
        .and_then(|b| {
            b.descendants()
                .find(|n| n.has_tag_name((ASVG_NS, "svgBlip")))
        })
        .and_then(|n| n.attribute((REL_NS, "embed")));
    svg.into_iter()
        .chain(blip.and_then(|b| b.attribute((REL_NS, "embed"))))
}
