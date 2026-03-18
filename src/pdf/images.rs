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

fn embed_single_image(
    img: &EmbeddedImage,
    image_xobjects: &mut Vec<(String, Ref)>,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> String {
    let xobj_ref = alloc();
    let pdf_name = format!("Im{}", image_xobjects.len() + 1);

    match img.format {
        ImageFormat::Jpeg => {
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
        ImageFormat::Png => {
            let cursor = std::io::Cursor::new(img.data.as_slice());
            let reader = image::ImageReader::with_format(
                std::io::BufReader::new(cursor),
                image::ImageFormat::Png,
            );
            let decoded = match reader.decode() {
                Ok(d) => d,
                Err(e) => {
                    log::warn!("PNG decode failed: {e} — writing 1x1 placeholder");
                    let mut xobj = pdf.image_xobject(xobj_ref, &[255, 255, 255]);
                    xobj.width(1);
                    xobj.height(1);
                    xobj.color_space().device_rgb();
                    xobj.bits_per_component(8);
                    image_xobjects.push((pdf_name.clone(), xobj_ref));
                    return pdf_name;
                }
            };
            let rgba: image::RgbaImage = decoded.to_rgba8();
            let (w, h) = (rgba.width(), rgba.height());
            let has_alpha = rgba.pixels().any(|p| p.0[3] < 255);

            let rgb_data: Vec<u8> = rgba
                .pixels()
                .flat_map(|p| [p.0[0], p.0[1], p.0[2]])
                .collect();
            let compressed_rgb = miniz_oxide::deflate::compress_to_vec_zlib(&rgb_data, 6);

            let smask_ref = if has_alpha {
                let alpha_data: Vec<u8> = rgba.pixels().map(|p| p.0[3]).collect();
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

            let mut xobj = pdf.image_xobject(xobj_ref, &compressed_rgb);
            xobj.filter(Filter::FlateDecode);
            xobj.width(w as i32);
            xobj.height(h as i32);
            xobj.color_space().device_rgb();
            xobj.bits_per_component(8);
            xobj.interpolate(true);
            if let Some(mask_ref) = smask_ref {
                xobj.s_mask(mask_ref);
            }
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
