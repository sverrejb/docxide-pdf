use std::collections::HashMap;

use pdf_writer::{Filter, Pdf, Ref};

use crate::model::{
    Block, Document, EmbeddedImage, HeaderFooter, ImageFormat, SectionProperties, Table,
};

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
                        write_image_xobject(
                            pdf, xobj_ref, &jpeg_buf, Filter::DctDecode, tw, th, None,
                        );
                        image_xobjects.push((pdf_name.clone(), xobj_ref));
                        return pdf_name;
                    }
                }
                // Fall through to raw embed on failure
            }
            // Passthrough: embed original JPEG bytes directly
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

pub(super) fn embed_all_images(
    doc: &Document,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> EmbeddedImages {
    let mut image_pdf_names: HashMap<usize, String> = HashMap::new();
    let mut inline_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();
    let mut image_xobjects: Vec<(String, Ref)> = Vec::new();
    let mut floating_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();

    {
        let mut global_block_idx = 0usize;
        for section in &doc.sections {
            for block in &section.blocks {
                if let Block::Paragraph(para) = block {
                    if let Some(img) = &para.image {
                        let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                        image_pdf_names.insert(global_block_idx, name);
                    }
                    for (run_idx, run) in para.runs.iter().enumerate() {
                        if let Some(img) = &run.inline_image {
                            let name = embed_single_image(img, &mut image_xobjects, pdf, alloc);
                            inline_image_pdf_names.insert((global_block_idx, run_idx), name);
                        }
                    }
                    for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                        let name = embed_single_image(&fi.image, &mut image_xobjects, pdf, alloc);
                        floating_image_pdf_names.insert((global_block_idx, fi_idx), name);
                    }
                }
                global_block_idx += 1;
            }
        }
    }

    let mut hf_image_names: HashMap<(usize, u8, usize), String> = HashMap::new();
    let mut hf_inline_image_names: HashMap<(usize, u8, usize, usize), String> = HashMap::new();
    let mut hf_floating_image_names: HashMap<(usize, u8, usize, usize), String> = HashMap::new();
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
                            }
                            for (ri, run) in para.runs.iter().enumerate() {
                                if let Some(img) = &run.inline_image {
                                    let name =
                                        embed_single_image(img, &mut image_xobjects, pdf, alloc);
                                    hf_inline_image_names.insert((si, hf_type, pi, ri), name);
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
                            }
                            pi += 1;
                        }
                    }
                }
            }
        }
    }

    let mut table_cell_image_names: HashMap<usize, String> = HashMap::new();
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
                                table_cell_image_names.insert(key, name);
                            }
                        }
                        for fi in &para.floating_images {
                            let key = std::sync::Arc::as_ptr(&fi.image.data) as usize;
                            if !table_cell_image_names.contains_key(&key) {
                                let name =
                                    embed_single_image(&fi.image, &mut image_xobjects, pdf, alloc);
                                table_cell_image_names.insert(key, name);
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
    }
}
