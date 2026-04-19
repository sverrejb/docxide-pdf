use std::collections::HashMap;

use pdf_writer::{Filter, Pdf, Ref};

use crate::model::{
    Block, Document, EmbeddedImage, HeaderFooter, ImageFormat, ImageShadow, SectionProperties,
    Table,
};

/// Per-image effect XObject names (shadow, glow, inner shadow, reflection).
/// Soft edge has no XObject — it modifies the image's own SMask.
#[derive(Default, Clone)]
pub(crate) struct EffectXObjs {
    pub shadow: Option<String>,
    pub glow: Option<String>,
    pub inner_shadow: Option<String>,
    pub reflection: Option<String>,
}

impl EffectXObjs {
    fn has_any(&self) -> bool {
        self.shadow.is_some() || self.glow.is_some()
            || self.inner_shadow.is_some() || self.reflection.is_some()
    }
}

pub(super) struct EmbeddedImages {
    pub(super) image_pdf_names: HashMap<usize, String>,
    pub(super) inline_image_pdf_names: HashMap<(usize, usize), String>,
    pub(super) floating_image_pdf_names: HashMap<(usize, usize), String>,
    pub(super) image_xobjects: Vec<(String, Ref)>,
    pub(super) hf_image_names: HashMap<(usize, u8, usize), String>,
    pub(super) hf_inline_image_names: HashMap<(usize, u8, usize, usize), String>,
    pub(super) hf_floating_image_names: HashMap<(usize, u8, usize, usize), String>,
    /// Images in table cell paragraphs, keyed by Arc data pointer address.
    pub(super) table_cell_image_names: HashMap<usize, String>,
    /// SmartArt image fills, keyed by Arc data pointer address.
    pub(super) smartart_image_names: HashMap<usize, String>,
    // Effect XObject names (parallel to image name maps)
    pub(super) effect_names: HashMap<usize, EffectXObjs>,
    pub(super) effect_floating_names: HashMap<(usize, usize), EffectXObjs>,
    pub(super) effect_inline_names: HashMap<(usize, usize), EffectXObjs>,
    pub(super) effect_hf_names: HashMap<(usize, u8, usize), EffectXObjs>,
    #[allow(dead_code)]
    pub(super) effect_hf_inline_names: HashMap<(usize, u8, usize, usize), EffectXObjs>,
    pub(super) effect_hf_floating_names: HashMap<(usize, u8, usize, usize), EffectXObjs>,
    pub(super) effect_table_names: HashMap<usize, EffectXObjs>,
}

const DOWNSAMPLE_DPI_THRESHOLD: f32 = 200.0;
const DOWNSAMPLE_DPI_TARGET: f32 = 150.0;
const JPEG_QUALITY: u8 = 85;

/// If pixel dimensions far exceed what's needed at the display size, return
/// target pixel dimensions for downscaling. Returns `None` if no downscaling needed.
fn downscale_target(
    pixel_w: u32,
    pixel_h: u32,
    display_w: f32,
    display_h: f32,
) -> Option<(u32, u32)> {
    if display_w <= 0.0 || display_h <= 0.0 {
        return None;
    }
    let eff_dpi_x = pixel_w as f32 / (display_w / 72.0);
    let eff_dpi_y = pixel_h as f32 / (display_h / 72.0);
    let eff_dpi = eff_dpi_x.max(eff_dpi_y);

    if eff_dpi <= DOWNSAMPLE_DPI_THRESHOLD {
        return None;
    }

    let target_w = ((display_w / 72.0) * DOWNSAMPLE_DPI_TARGET).round() as u32;
    let target_h = ((display_h / 72.0) * DOWNSAMPLE_DPI_TARGET).round() as u32;

    if target_w >= pixel_w || target_h >= pixel_h || target_w < 4 || target_h < 4 {
        return None;
    }

    Some((target_w, target_h))
}

/// Encode an RGB image as JPEG, returning the bytes. Returns `None` on failure.
fn encode_jpeg(rgb: &image::RgbImage) -> Option<Vec<u8>> {
    let mut buf = Vec::new();
    let mut encoder = image::codecs::jpeg::JpegEncoder::new_with_quality(&mut buf, JPEG_QUALITY);
    match encoder.encode_image(rgb) {
        Ok(()) => Some(buf),
        Err(e) => {
            log::warn!("JPEG encode failed: {e}");
            None
        }
    }
}

/// Fallback PNG decoder using the `png` crate directly. The `image` crate's
/// format-pinned reader can fail on interlaced or unusual PNGs; this handles
/// those cases by decoding with EXPAND+ALPHA transformations.
fn decode_png_raw(data: &[u8]) -> Option<(u32, u32, Vec<u8>)> {
    let mut decoder = png::Decoder::new(std::io::Cursor::new(data));
    decoder.set_transformations(png::Transformations::EXPAND | png::Transformations::ALPHA);
    let mut reader = decoder.read_info().ok()?;
    let buf_size = reader.output_buffer_size().unwrap_or(reader.info().raw_bytes() as usize);
    let mut buf = vec![0u8; buf_size];
    let info = reader.next_frame(&mut buf).ok()?;
    let (w, h) = (info.width, info.height);
    let rgba = match (info.color_type, info.bit_depth) {
        (png::ColorType::Rgba, png::BitDepth::Eight) => {
            buf.truncate(info.buffer_size());
            buf
        }
        (png::ColorType::Rgb, png::BitDepth::Eight) => {
            let rgb = &buf[..info.buffer_size()];
            let mut rgba = Vec::with_capacity((w * h * 4) as usize);
            for chunk in rgb.chunks_exact(3) {
                rgba.extend_from_slice(chunk);
                rgba.push(255);
            }
            rgba
        }
        (png::ColorType::Grayscale, png::BitDepth::Eight) => {
            let gray = &buf[..info.buffer_size()];
            let mut rgba = Vec::with_capacity((w * h * 4) as usize);
            for &g in gray {
                rgba.extend_from_slice(&[g, g, g, 255]);
            }
            rgba
        }
        _ => return None,
    };
    Some((w, h, rgba))
}

/// Write an image XObject into the PDF with the given filter and data.
/// Optionally attaches a soft mask (alpha channel) reference.
fn write_image_xobject(
    pdf: &mut Pdf,
    xobj_ref: Ref,
    data: &[u8],
    filter: Filter,
    w: u32,
    h: u32,
    smask: Option<Ref>,
) {
    let mut xobj = pdf.image_xobject(xobj_ref, data);
    xobj.filter(filter);
    xobj.width(w as i32);
    xobj.height(h as i32);
    xobj.color_space().device_rgb();
    xobj.bits_per_component(8);
    xobj.interpolate(true);
    if let Some(mask_ref) = smask {
        xobj.s_mask(mask_ref);
    }
}

fn embed_single_image(
    img: &EmbeddedImage,
    image_xobjects: &mut Vec<(String, Ref)>,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> String {
    let xobj_ref = alloc();
    let pdf_name = format!("Im{}", image_xobjects.len() + 1);
    let target = downscale_target(
        img.pixel_width,
        img.pixel_height,
        img.display_width,
        img.display_height,
    );

    // Helper: create soft-edge SMask for a given pixel size
    let make_soft_edge_smask = |w: u32, h: u32, pdf: &mut Pdf, alloc: &mut dyn FnMut() -> Ref| -> Option<Ref> {
        let se = img.soft_edge.as_ref()?;
        let radius_px = se.radius * w as f32 / img.display_width;
        let mask_data = super::color::generate_soft_edge_mask(w, h, radius_px);
        let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&mask_data, 6);
        let mask_ref = alloc();
        let mut mask = pdf.image_xobject(mask_ref, &compressed);
        mask.filter(Filter::FlateDecode);
        mask.width(w as i32);
        mask.height(h as i32);
        mask.color_space().device_gray();
        mask.bits_per_component(8);
        mask.interpolate(true);
        Some(mask_ref)
    };

    match img.format {
        ImageFormat::Jpeg => {
            if let Some((tw, th)) = target {
                // Decode -> resize -> re-encode as JPEG
                let cursor = std::io::Cursor::new(img.data.as_slice());
                let reader = image::ImageReader::with_format(
                    std::io::BufReader::new(cursor),
                    image::ImageFormat::Jpeg,
                );
                if let Ok(decoded) = reader.decode() {
                    let resized =
                        decoded.resize_exact(tw, th, image::imageops::FilterType::Lanczos3);
                    if let Some(jpeg_buf) = encode_jpeg(&resized.to_rgb8()) {
                        log::debug!(
                            "Downscaled JPEG {}x{} -> {}x{} ({} -> {} bytes)",
                            img.pixel_width,
                            img.pixel_height,
                            tw,
                            th,
                            img.data.len(),
                            jpeg_buf.len()
                        );
                        let se_mask = make_soft_edge_smask(tw, th, pdf, alloc);
                        write_image_xobject(
                            pdf, xobj_ref, &jpeg_buf, Filter::DctDecode, tw, th, se_mask,
                        );
                        image_xobjects.push((pdf_name.clone(), xobj_ref));
                        return pdf_name;
                    }
                }
                // Fall through to raw embed on failure
            }
            // Passthrough: embed original JPEG bytes directly
            let se_mask = make_soft_edge_smask(img.pixel_width, img.pixel_height, pdf, alloc);
            let mut xobj = pdf.image_xobject(xobj_ref, &*img.data);
            xobj.filter(Filter::DctDecode);
            xobj.width(img.pixel_width as i32);
            xobj.height(img.pixel_height as i32);
            match img.jpeg_components {
                1 => xobj.color_space().device_gray(),
                4 => xobj.color_space().device_cmyk(),
                _ => xobj.color_space().device_rgb(),
            };
            xobj.bits_per_component(8);
            xobj.interpolate(true);
            if let Some(mask_ref) = se_mask {
                xobj.s_mask(mask_ref);
            }
        }
        ImageFormat::Png | ImageFormat::Bmp => {
            let img_fmt = match img.format {
                ImageFormat::Bmp => image::ImageFormat::Bmp,
                _ => image::ImageFormat::Png,
            };
            let cursor = std::io::Cursor::new(img.data.as_slice());
            let reader =
                image::ImageReader::with_format(std::io::BufReader::new(cursor), img_fmt);
            let decoded = match reader.decode() {
                Ok(d) => d,
                Err(_) => {
                    // Fallback: use png crate directly for PNGs the image
                    // crate's format-pinned reader can't decode
                    match decode_png_raw(&img.data) {
                        Some((w, h, rgba_data)) => {
                            image::DynamicImage::ImageRgba8(
                                image::RgbaImage::from_raw(w, h, rgba_data)
                                    .expect("RGBA data size matches dimensions"),
                            )
                        }
                        None => {
                            log::warn!("PNG decode failed — writing 1x1 placeholder");
                            let mut xobj = pdf.image_xobject(xobj_ref, &[255, 255, 255]);
                            xobj.width(1);
                            xobj.height(1);
                            xobj.color_space().device_rgb();
                            xobj.bits_per_component(8);
                            image_xobjects.push((pdf_name.clone(), xobj_ref));
                            return pdf_name;
                        }
                    }
                }
            };

            let decoded = if let Some((tw, th)) = target {
                log::debug!(
                    "Downscaling PNG/BMP {}x{} -> {}x{}",
                    decoded.width(),
                    decoded.height(),
                    tw,
                    th
                );
                decoded.resize_exact(tw, th, image::imageops::FilterType::Lanczos3)
            } else {
                decoded
            };

            let rgba: image::RgbaImage = decoded.to_rgba8();
            let (w, h) = (rgba.width(), rgba.height());
            let pixel_count = (w * h) as usize;

            // Single pass: split RGBA into RGB + alpha, detect transparency
            let mut rgb_data = Vec::with_capacity(pixel_count * 3);
            let mut alpha_data = Vec::with_capacity(pixel_count);
            let mut has_alpha = false;
            for p in rgba.pixels() {
                rgb_data.push(p.0[0]);
                rgb_data.push(p.0[1]);
                rgb_data.push(p.0[2]);
                let a = p.0[3];
                alpha_data.push(a);
                if a < 255 {
                    has_alpha = true;
                }
            }

            // Apply soft-edge mask: multiply with existing alpha or create new alpha
            if let Some(se) = img.soft_edge.as_ref() {
                let radius_px = se.radius * w as f32 / img.display_width;
                let se_mask = super::color::generate_soft_edge_mask(w, h, radius_px);
                if has_alpha {
                    for (a, &se) in alpha_data.iter_mut().zip(se_mask.iter()) {
                        *a = ((*a as u16 * se as u16) / 255) as u8;
                    }
                } else {
                    alpha_data = se_mask;
                    has_alpha = true;
                }
            }

            let smask_ref = if has_alpha {
                let compressed_alpha = miniz_oxide::deflate::compress_to_vec_zlib(&alpha_data, 6);
                let mask_ref = alloc();
                let mut mask = pdf.image_xobject(mask_ref, &compressed_alpha);
                mask.filter(Filter::FlateDecode);
                mask.width(w as i32);
                mask.height(h as i32);
                mask.color_space().device_gray();
                mask.bits_per_component(8);
                Some(mask_ref)
            } else {
                None
            };

            // Try both JPEG and Flate, pick whichever is smaller.
            // JPEG wins for photographic content; Flate wins for synthetic/chart images.
            let compressed_rgb = miniz_oxide::deflate::compress_to_vec_zlib(&rgb_data, 6);
            let rgb = image::RgbImage::from_raw(w, h, rgb_data)
                .expect("RGB data size matches dimensions");
            let jpeg_buf = encode_jpeg(&rgb);

            let use_jpeg = jpeg_buf
                .as_ref()
                .is_some_and(|j| j.len() < compressed_rgb.len());

            let (filter, data) = if use_jpeg {
                (Filter::DctDecode, jpeg_buf.unwrap())
            } else {
                (Filter::FlateDecode, compressed_rgb)
            };
            write_image_xobject(pdf, xobj_ref, &data, filter, w, h, smask_ref);
        }
    }

    image_xobjects.push((pdf_name.clone(), xobj_ref));
    pdf_name
}

/// Embed a shadow as a 1x1 solid-color Image XObject whose SMask is a
/// full-resolution Gaussian-blur grayscale mask. Returns the XObject name.
fn embed_shadow(
    shadow: &ImageShadow,
    display_width: f32,
    display_height: f32,
    image_xobjects: &mut Vec<(String, Ref)>,
    shadow_counter: &mut usize,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> String {
    let (mask_pixels, mask_w, mask_h) =
        super::color::generate_shadow_mask(display_width, display_height, shadow.blur_radius, shadow.alpha);

    // Compress the grayscale mask
    let compressed_mask = miniz_oxide::deflate::compress_to_vec_zlib(&mask_pixels, 6);

    // Create the grayscale mask Image XObject
    let mask_ref = alloc();
    let mut mask = pdf.image_xobject(mask_ref, &compressed_mask);
    mask.filter(Filter::FlateDecode);
    mask.width(mask_w as i32);
    mask.height(mask_h as i32);
    mask.color_space().device_gray();
    mask.bits_per_component(8);
    mask.interpolate(true);
    drop(mask);

    // Create a 1x1 solid-color RGB image with the mask as SMask
    let color_data = [shadow.color[0], shadow.color[1], shadow.color[2]];
    let color_ref = alloc();
    let mut color_img = pdf.image_xobject(color_ref, &color_data);
    color_img.width(1);
    color_img.height(1);
    color_img.color_space().device_rgb();
    color_img.bits_per_component(8);
    color_img.interpolate(true);
    color_img.s_mask(mask_ref);
    drop(color_img);

    *shadow_counter += 1;
    let name = format!("Sh{}", *shadow_counter);
    image_xobjects.push((name.clone(), color_ref));
    name
}

/// Embed a reflection: re-embed the image data with a vertical gradient SMask.
/// Returns the XObject name, or None if the image can't be re-embedded.
fn embed_reflection(
    img: &EmbeddedImage,
    refl: &crate::model::ImageReflection,
    image_xobjects: &mut Vec<(String, Ref)>,
    effect_counter: &mut usize,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> Option<String> {
    // Generate vertical gradient mask: 1 pixel wide, image pixel height tall.
    // The rendering flips the image vertically, so image row 0 (top of original)
    // becomes the bottom of the reflection (far from image), and the last row
    // (bottom of original) becomes the top of the reflection (near image).
    // Only the bottom `end_pos` fraction of the image rows are visible; the rest
    // are transparent. Within the visible band, alpha goes from endA at the cutoff
    // edge to stA at the last row (which appears nearest the original image).
    let grad_h = img.pixel_height.max(1);
    let visible_start = ((1.0 - refl.end_pos) * grad_h as f32).round() as u32;
    let visible_span = (grad_h - visible_start).max(1);
    let mut grad_data = Vec::with_capacity(grad_h as usize);
    for y in 0..grad_h {
        if y < visible_start {
            grad_data.push(0);
        } else {
            let t = (y - visible_start) as f32 / (visible_span - 1).max(1) as f32;
            // t=0 at cutoff edge (far from image after flip) → endA
            // t=1 at last row (near image after flip) → stA
            let a = refl.end_alpha * (1.0 - t) + refl.start_alpha * t;
            grad_data.push((a * 255.0).round().min(255.0) as u8);
        }
    }
    let compressed_grad = miniz_oxide::deflate::compress_to_vec_zlib(&grad_data, 6);
    let grad_ref = alloc();
    let mut grad = pdf.image_xobject(grad_ref, &compressed_grad);
    grad.filter(Filter::FlateDecode);
    grad.width(1);
    grad.height(grad_h as i32);
    grad.color_space().device_gray();
    grad.bits_per_component(8);
    grad.interpolate(true);
    drop(grad);

    // Re-embed the same image data as a new XObject with the gradient as SMask.
    let refl_ref = alloc();
    match img.format {
        ImageFormat::Jpeg => {
            let mut xobj = pdf.image_xobject(refl_ref, &*img.data);
            xobj.filter(Filter::DctDecode);
            xobj.width(img.pixel_width as i32);
            xobj.height(img.pixel_height as i32);
            match img.jpeg_components {
                1 => xobj.color_space().device_gray(),
                4 => xobj.color_space().device_cmyk(),
                _ => xobj.color_space().device_rgb(),
            };
            xobj.bits_per_component(8);
            xobj.interpolate(true);
            xobj.s_mask(grad_ref);
        }
        ImageFormat::Png | ImageFormat::Bmp => {
            // Decode and re-encode as RGB (strip any existing alpha — the gradient replaces it)
            let img_fmt = match img.format {
                ImageFormat::Bmp => image::ImageFormat::Bmp,
                _ => image::ImageFormat::Png,
            };
            let cursor = std::io::Cursor::new(img.data.as_slice());
            let reader = image::ImageReader::with_format(std::io::BufReader::new(cursor), img_fmt);
            let decoded = match reader.decode() {
                Ok(d) => d,
                Err(_) => return None,
            };
            let rgb = decoded.to_rgb8();
            let (w, h) = (rgb.width(), rgb.height());
            let rgb_data: Vec<u8> = rgb.into_raw();

            if let Some(jpeg_buf) = encode_jpeg(
                &image::RgbImage::from_raw(w, h, rgb_data.clone())
                    .expect("RGB data size matches"),
            ) {
                let mut xobj = pdf.image_xobject(refl_ref, &jpeg_buf);
                xobj.filter(Filter::DctDecode);
                xobj.width(w as i32);
                xobj.height(h as i32);
                xobj.color_space().device_rgb();
                xobj.bits_per_component(8);
                xobj.interpolate(true);
                xobj.s_mask(grad_ref);
            } else {
                let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&rgb_data, 6);
                write_image_xobject(pdf, refl_ref, &compressed, Filter::FlateDecode, w, h, Some(grad_ref));
            }
        }
    }

    *effect_counter += 1;
    let name = format!("Sh{}", *effect_counter);
    image_xobjects.push((name.clone(), refl_ref));
    Some(name)
}

/// Build all effect XObjects for one image (shadow, glow, etc.).
fn embed_image_effects(
    img: &EmbeddedImage,
    image_xobjects: &mut Vec<(String, Ref)>,
    effect_counter: &mut usize,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> EffectXObjs {
    let mut fx = EffectXObjs::default();
    if let Some(shadow) = img.shadow.as_ref() {
        if shadow.blur_radius > 0.0 {
            fx.shadow = Some(embed_shadow(
                shadow, img.display_width, img.display_height,
                image_xobjects, effect_counter, pdf, alloc,
            ));
        }
    }
    if let Some(glow) = img.glow.as_ref() {
        // Glow reuses shadow embedding: centered blur (zero offset) with glow color
        let shadow_equiv = crate::model::ImageShadow {
            offset_x: 0.0,
            offset_y: 0.0,
            blur_radius: glow.radius,
            color: glow.color,
            alpha: glow.alpha,
        };
        fx.glow = Some(embed_shadow(
            &shadow_equiv, img.display_width, img.display_height,
            image_xobjects, effect_counter, pdf, alloc,
        ));
    }
    if let Some(inner) = img.inner_shadow.as_ref() {
        let (mask_pixels, mask_w, mask_h) = super::color::generate_inner_shadow_mask(
            img.display_width, img.display_height, inner.blur_radius, inner.alpha,
        );
        let compressed_mask = miniz_oxide::deflate::compress_to_vec_zlib(&mask_pixels, 6);
        let mask_ref = alloc();
        let mut mask = pdf.image_xobject(mask_ref, &compressed_mask);
        mask.filter(Filter::FlateDecode);
        mask.width(mask_w as i32);
        mask.height(mask_h as i32);
        mask.color_space().device_gray();
        mask.bits_per_component(8);
        mask.interpolate(true);
        drop(mask);

        let color_data = [inner.color[0], inner.color[1], inner.color[2]];
        let color_ref = alloc();
        let mut color_img = pdf.image_xobject(color_ref, &color_data);
        color_img.width(1);
        color_img.height(1);
        color_img.color_space().device_rgb();
        color_img.bits_per_component(8);
        color_img.interpolate(true);
        color_img.s_mask(mask_ref);
        drop(color_img);

        *effect_counter += 1;
        let name = format!("Sh{}", *effect_counter);
        image_xobjects.push((name.clone(), color_ref));
        fx.inner_shadow = Some(name);
    }
    if let Some(refl) = img.reflection.as_ref() {
        fx.reflection = embed_reflection(img, refl, image_xobjects, effect_counter, pdf, alloc);
    }
    fx
}

pub(super) fn embed_all_images(
    doc: &Document,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> EmbeddedImages {
    let mut image_pdf_names: HashMap<usize, String> = HashMap::new();
    let mut inline_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();
    let mut image_xobjects: Vec<(String, Ref)> = Vec::new();
    let mut floating_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();
    let mut effect_counter = 0usize;
    let mut effect_names: HashMap<usize, EffectXObjs> = HashMap::new();
    let mut effect_floating_names: HashMap<(usize, usize), EffectXObjs> = HashMap::new();
    let mut effect_inline_names: HashMap<(usize, usize), EffectXObjs> = HashMap::new();

    {
        let mut global_block_idx = 0usize;
        for section in &doc.sections {
            for block in &section.blocks {
                if let Block::Paragraph(para) = block {
                    if let Some(img) = &para.image {
                        let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                        image_pdf_names.insert(global_block_idx, name);
                        let fx = embed_image_effects(img, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                        if fx.has_any() { effect_names.insert(global_block_idx, fx); }
                    }
                    for (run_idx, run) in para.runs.iter().enumerate() {
                        if let Some(img) = &run.inline_image {
                            let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                            inline_image_pdf_names.insert((global_block_idx, run_idx), name);
                            let fx = embed_image_effects(img, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                            if fx.has_any() { effect_inline_names.insert((global_block_idx, run_idx), fx); }
                        }
                    }
                    for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                        let name = embed_single_image(&fi.image, &mut image_xobjects, pdf, alloc);
                        floating_image_pdf_names.insert((global_block_idx, fi_idx), name);
                        let fx = embed_image_effects(&fi.image, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                        if fx.has_any() { effect_floating_names.insert((global_block_idx, fi_idx), fx); }
                    }
                }
                global_block_idx += 1;
            }
        }
    }

    let mut hf_image_names: HashMap<(usize, u8, usize), String> = HashMap::new();
    let mut hf_inline_image_names: HashMap<(usize, u8, usize, usize), String> = HashMap::new();
    let mut hf_floating_image_names: HashMap<(usize, u8, usize, usize), String> = HashMap::new();
    let mut effect_hf_names: HashMap<(usize, u8, usize), EffectXObjs> = HashMap::new();
    let mut effect_hf_inline_names: HashMap<(usize, u8, usize, usize), EffectXObjs> = HashMap::new();
    let mut effect_hf_floating_names: HashMap<(usize, u8, usize, usize), EffectXObjs> = HashMap::new();
    {
        let hf_variants: [(u8, fn(&SectionProperties) -> Option<&HeaderFooter>); 6] = [
            (0, |sp| sp.header_default.as_ref()),
            (1, |sp| sp.header_first.as_ref()),
            (2, |sp| sp.footer_default.as_ref()),
            (3, |sp| sp.footer_first.as_ref()),
            (4, |sp| sp.header_even.as_ref()),
            (5, |sp| sp.footer_even.as_ref()),
        ];
        for (si, section) in doc.sections.iter().enumerate() {
            for &(hf_type, accessor) in &hf_variants {
                if let Some(hf) = accessor(&section.properties) {
                    let mut pi = 0usize;
                    for block in &hf.blocks {
                        if let Block::Paragraph(para) = block {
                            if let Some(img) = &para.image {
                                let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                hf_image_names.insert((si, hf_type, pi), name);
                                let fx = embed_image_effects(img, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                                if fx.has_any() { effect_hf_names.insert((si, hf_type, pi), fx); }
                            }
                            for (ri, run) in para.runs.iter().enumerate() {
                                if let Some(img) = &run.inline_image {
                                    let name =
                                        embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                    hf_inline_image_names.insert((si, hf_type, pi, ri), name);
                                    let fx = embed_image_effects(img, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                                    if fx.has_any() { effect_hf_inline_names.insert((si, hf_type, pi, ri), fx); }
                                }
                            }
                            for (fi, floating) in para.floating_images.iter().enumerate() {
                                let name = embed_single_image(
                                    &floating.image,
                                    &mut image_xobjects,
                                    pdf,
                                    alloc,
                                );
                                hf_floating_image_names.insert((si, hf_type, pi, fi), name);
                                let fx = embed_image_effects(&floating.image, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                                if fx.has_any() { effect_hf_floating_names.insert((si, hf_type, pi, fi), fx); }
                            }
                            pi += 1;
                        }
                    }
                }
            }
        }
    }

    let mut table_cell_image_names: HashMap<usize, String> = HashMap::new();
    let mut effect_table_names: HashMap<usize, EffectXObjs> = HashMap::new();
    {
        let mut tables: Vec<&Table> = Vec::new();
        for section in &doc.sections {
            for block in &section.blocks {
                if let Block::Table(table) = block {
                    tables.push(table);
                }
            }
            let hf_list: [Option<&HeaderFooter>; 6] = [
                section.properties.header_default.as_ref(),
                section.properties.header_first.as_ref(),
                section.properties.footer_default.as_ref(),
                section.properties.footer_first.as_ref(),
                section.properties.header_even.as_ref(),
                section.properties.footer_even.as_ref(),
            ];
            for hf_opt in hf_list {
                if let Some(hf) = hf_opt {
                    for block in &hf.blocks {
                        if let Block::Table(table) = block {
                            tables.push(table);
                        }
                    }
                }
            }
        }
        for table in tables {
            for row in &table.rows {
                for cell in &row.cells {
                    for para in cell.all_paragraphs() {
                        if let Some(img) = &para.image {
                            let key = std::sync::Arc::as_ptr(&img.data) as usize;
                            if !table_cell_image_names.contains_key(&key) {
                                let name =
                                    embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                table_cell_image_names.insert(key, name.clone());
                                let fx = embed_image_effects(img, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                                if fx.has_any() { effect_table_names.insert(key, fx); }
                            }
                        }
                        for fi in &para.floating_images {
                            let key = std::sync::Arc::as_ptr(&fi.image.data) as usize;
                            if !table_cell_image_names.contains_key(&key) {
                                let name =
                                    embed_single_image(&fi.image, &mut image_xobjects, pdf, alloc);
                                table_cell_image_names.insert(key, name.clone());
                                let fx = embed_image_effects(&fi.image, &mut image_xobjects, &mut effect_counter, pdf, alloc);
                                if fx.has_any() { effect_table_names.insert(key, fx); }
                            }
                        }
                    }
                }
            }
        }
    }

    let mut smartart_image_names: HashMap<usize, String> = HashMap::new();
    for section in &doc.sections {
        for block in &section.blocks {
            if let Block::Paragraph(para) = block {
                for diagram in &para.smartart {
                    for shape in &diagram.shapes {
                        if let Some(ref img) = shape.image_fill {
                            let key = std::sync::Arc::as_ptr(&img.data) as usize;
                            if !smartart_image_names.contains_key(&key) {
                                let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                smartart_image_names.insert(key, name);
                            }
                        }
                    }
                }
            }
        }
    }

    EmbeddedImages {
        image_pdf_names,
        inline_image_pdf_names,
        floating_image_pdf_names,
        image_xobjects,
        hf_image_names,
        hf_inline_image_names,
        hf_floating_image_names,
        table_cell_image_names,
        smartart_image_names,
        effect_names,
        effect_floating_names,
        effect_inline_names,
        effect_hf_names,
        effect_hf_inline_names,
        effect_hf_floating_names,
        effect_table_names,
    }
}
