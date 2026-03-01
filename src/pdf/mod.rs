mod layout;
mod table;

use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str};

use crate::error::Error;
use crate::fonts::{FontEntry, encode_as_gids, font_key, register_font, to_winansi_bytes};
use crate::model::{
    Alignment, Block, Document, EmbeddedImage, FieldCode, Footnote, HeaderFooter,
    HorizontalPosition, ImageFormat, LineSpacing, Run, SectionBreakType, SectionProperties,
};

use layout::{
    LinkAnnotation,
    build_paragraph_lines, build_tabbed_line,
    font_metric, is_text_empty, render_paragraph_lines, tallest_run_metrics,
};
use table::render_table;

fn border_eq(a: &Option<crate::model::ParagraphBorder>, b: &Option<crate::model::ParagraphBorder>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => a.width_pt == b.width_pt && a.color == b.color,
        _ => false,
    }
}

fn borders_match(a: &crate::model::ParagraphBorders, b: &crate::model::ParagraphBorders) -> bool {
    border_eq(&a.top, &b.top)
        && border_eq(&a.bottom, &b.bottom)
        && border_eq(&a.left, &b.left)
        && border_eq(&a.right, &b.right)
        && border_eq(&a.between, &b.between)
}

fn resolve_line_h(ls: LineSpacing, font_size: f32, tallest_lhr: Option<f32>) -> f32 {
    match ls {
        LineSpacing::Auto(mult) => tallest_lhr
            .map(|ratio| font_size * ratio * mult)
            .unwrap_or(font_size * 1.2 * mult),
        LineSpacing::Exact(pts) => pts,
        LineSpacing::AtLeast(min_pts) => {
            let natural = tallest_lhr
                .map(|ratio| font_size * ratio)
                .unwrap_or(font_size * 1.2);
            natural.max(min_pts)
        }
    }
}

fn render_header_footer(
    content: &mut Content,
    hf: &HeaderFooter,
    seen_fonts: &HashMap<String, FontEntry>,
    sp: &SectionProperties,
    doc_line_spacing: LineSpacing,
    is_header: bool,
    page_num: usize,
    total_pages: usize,
    para_image_names: &HashMap<usize, String>,
    inline_image_names: &HashMap<(usize, usize), String>,
) {
    let text_width = sp.page_width - sp.margin_left - sp.margin_right;

    for (pi, para) in hf.paragraphs.iter().enumerate() {
        let has_para_image = para.image.is_some();
        let text_empty = is_text_empty(&para.runs);

        let substituted_runs: Vec<Run> = para
            .runs
            .iter()
            .map(|run| {
                let mut r = run.clone();
                r.field_code = None;
                if let Some(ref fc) = run.field_code {
                    r.text = match fc {
                        FieldCode::Page => page_num.to_string(),
                        FieldCode::NumPages => total_pages.to_string(),
                    };
                }
                r
            })
            .collect();

        let (font_size, _, tallest_ar) = tallest_run_metrics(&substituted_runs, seen_fonts);
        let ascender_ratio = tallest_ar.unwrap_or(0.75);

        let baseline_y = if is_header {
            sp.page_height - sp.header_margin - font_size * ascender_ratio
        } else {
            sp.footer_margin + font_size * (1.0 - ascender_ratio)
        };

        if (has_para_image || text_empty) && para.content_height > 0.0 {
            if let Some(pdf_name) = para_image_names.get(&pi) {
                let img = para.image.as_ref().unwrap();
                let y_bottom = baseline_y + font_size * ascender_ratio - img.display_height;
                let x = sp.margin_left
                    + match para.alignment {
                        Alignment::Center => (text_width - img.display_width).max(0.0) / 2.0,
                        Alignment::Right => (text_width - img.display_width).max(0.0),
                        _ => 0.0,
                    };
                content.save_state();
                content.transform([
                    img.display_width,
                    0.0,
                    0.0,
                    img.display_height,
                    x,
                    y_bottom,
                ]);
                content.x_object(Name(pdf_name.as_bytes()));
                content.restore_state();
            }
            continue;
        }

        if text_empty {
            continue;
        }

        let block_inline_images: HashMap<usize, String> = inline_image_names
            .iter()
            .filter(|((pi2, _), _)| *pi2 == pi)
            .map(|((_, ri), name)| (*ri, name.clone()))
            .collect();

        let lines = build_paragraph_lines(&substituted_runs, seen_fonts, text_width, 0.0, &block_inline_images);

        let effective_ls = para.line_spacing.unwrap_or(doc_line_spacing);
        let tallest_lhr = font_metric(&substituted_runs, seen_fonts, |e| e.line_h_ratio);
        let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);

        render_paragraph_lines(
            content,
            &lines,
            &para.alignment,
            sp.margin_left,
            text_width,
            baseline_y,
            line_h,
            lines.len(),
            0,
            &mut Vec::new(),
            0.0,
            seen_fonts,
        );
    }
}

pub fn render(doc: &Document) -> Result<Vec<u8>, Error> {
    let t0 = std::time::Instant::now();
    let mut pdf = Pdf::new();
    let mut next_id = 1i32;
    let mut alloc = || {
        let r = Ref::new(next_id);
        next_id += 1;
        r
    };

    let catalog_id = alloc();
    let pages_id = alloc();

    // Phase 1: collect unique font names (with variant) and embed them
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();

    // Collect all runs from all sections (blocks, headers/footers)
    let hf_runs = doc
        .sections
        .iter()
        .flat_map(|s| {
            [
                &s.properties.header_default,
                &s.properties.header_first,
                &s.properties.footer_default,
                &s.properties.footer_first,
            ]
            .into_iter()
            .filter_map(|hf| hf.as_ref())
            .flat_map(|hf| hf.paragraphs.iter())
            .flat_map(|p| p.runs.iter())
        });

    let footnote_runs = doc
        .footnotes
        .values()
        .flat_map(|fn_| fn_.paragraphs.iter())
        .flat_map(|p| p.runs.iter());

    fn para_runs_with_textboxes(para: &crate::model::Paragraph) -> Vec<&Run> {
        let mut out: Vec<&Run> = para.runs.iter().collect();
        for tb in &para.textboxes {
            for tp in &tb.paragraphs {
                out.extend(para_runs_with_textboxes(tp));
            }
        }
        out
    }

    let all_runs: Vec<&Run> = doc
        .sections
        .iter()
        .flat_map(|s| s.blocks.iter())
        .flat_map(|block| -> Vec<&Run> {
            match block {
                Block::Paragraph(para) => para_runs_with_textboxes(para),
                Block::Table(table) => table
                    .rows
                    .iter()
                    .flat_map(|row| row.cells.iter())
                    .flat_map(|cell| cell.paragraphs.iter())
                    .flat_map(|para| para_runs_with_textboxes(para))
                    .collect(),
            }
        })
        .chain(hf_runs)
        .chain(footnote_runs)
        .collect();

    let t_collect = t0.elapsed();

    // Collect used characters per font key for subsetting
    let mut used_chars_per_font: HashMap<String, HashSet<char>> = HashMap::new();
    for run in &all_runs {
        let key = font_key(run);
        let chars = used_chars_per_font.entry(key).or_default();
        if run.caps || run.small_caps {
            chars.extend(run.text.to_uppercase().chars());
        } else {
            chars.extend(run.text.chars());
        }
        if let Some(ref fc) = run.field_code {
            match fc {
                FieldCode::Page | FieldCode::NumPages => {
                    chars.extend('0'..='9');
                }
            }
        }
        if run.footnote_id.is_some() || run.is_footnote_ref_mark {
            chars.extend('0'..='9');
        }
    }
    // List labels and leader characters from paragraphs
    let all_paras = doc.sections.iter()
        .flat_map(|s| s.blocks.iter())
        .flat_map(|block| -> Box<dyn Iterator<Item = &crate::model::Paragraph> + '_> {
            match block {
                Block::Paragraph(p) => Box::new(std::iter::once(p)),
                Block::Table(t) => Box::new(
                    t.rows.iter()
                        .flat_map(|row| row.cells.iter())
                        .flat_map(|cell| cell.paragraphs.iter()),
                ),
            }
        });
    for para in all_paras {
        if !para.list_label.is_empty()
            && let Some(run) = para.runs.first()
        {
            let key = font_key(run);
            used_chars_per_font
                .entry(key)
                .or_default()
                .extend(para.list_label.chars());
        }
        for stop in &para.tab_stops {
            if let Some(leader_char) = stop.leader
                && let Some(run) = para.runs.first()
            {
                let key = font_key(run);
                used_chars_per_font
                    .entry(key)
                    .or_default()
                    .insert(leader_char);
            }
        }
    }
    for section in &doc.sections {
        for hf in [
            &section.properties.header_default,
            &section.properties.header_first,
            &section.properties.footer_default,
            &section.properties.footer_first,
        ]
        .into_iter()
        .flatten()
        {
            for para in &hf.paragraphs {
                for run in &para.runs {
                    let key = font_key(run);
                    let chars = used_chars_per_font.entry(key).or_default();
                    if run.caps || run.small_caps {
                        chars.extend(run.text.to_uppercase().chars());
                    } else {
                        chars.extend(run.text.chars());
                    }
                    if let Some(ref fc) = run.field_code {
                        match fc {
                            FieldCode::Page | FieldCode::NumPages => {
                                chars.extend('0'..='9');
                            }
                        }
                    }
                }
            }
        }
    }
    // Ensure space is always included
    for chars in used_chars_per_font.values_mut() {
        chars.insert(' ');
    }

    for run in &all_runs {
        let key = font_key(run);
        if !seen_fonts.contains_key(&key) {
            let pdf_name = format!("F{}", font_order.len() + 1);
            let used = used_chars_per_font.get(&key).cloned().unwrap_or_default();
            let entry = register_font(
                &mut pdf,
                &run.font_name,
                run.bold,
                run.italic,
                pdf_name,
                &mut alloc,
                &doc.embedded_fonts,
                &used,
            );
            seen_fonts.insert(key.clone(), entry);
            font_order.push(key);
        }
    }

    if seen_fonts.is_empty() {
        let pdf_name = "F1".to_string();
        let entry = register_font(
            &mut pdf,
            "Helvetica",
            false,
            false,
            pdf_name,
            &mut alloc,
            &doc.embedded_fonts,
            &HashSet::new(),
        );
        seen_fonts.insert("Helvetica".to_string(), entry);
        font_order.push("Helvetica".to_string());
    }

    let t_fonts = t0.elapsed();

    // Phase 1b: embed images
    // Keys use global_block_idx (flat index across all sections)
    let mut image_pdf_names: HashMap<usize, String> = HashMap::new();
    let mut inline_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();
    let mut image_xobjects: Vec<(String, Ref)> = Vec::new();

    let embed_image = |img: &EmbeddedImage,
                           image_xobjects: &mut Vec<(String, Ref)>,
                           pdf: &mut Pdf,
                           alloc: &mut dyn FnMut() -> Ref|
     -> String {
        let xobj_ref = alloc();
        let pdf_name = format!("Im{}", image_xobjects.len() + 1);

        match img.format {
            ImageFormat::Jpeg => {
                let mut xobj = pdf.image_xobject(xobj_ref, &img.data);
                xobj.filter(Filter::DctDecode);
                xobj.width(img.pixel_width as i32);
                xobj.height(img.pixel_height as i32);
                xobj.color_space().device_rgb();
                xobj.bits_per_component(8);
            }
            ImageFormat::Png => {
                let cursor = std::io::Cursor::new(&img.data);
                let reader = image::ImageReader::with_format(
                    std::io::BufReader::new(cursor),
                    image::ImageFormat::Png,
                );
                if let Ok(decoded) = reader.decode() {
                    let rgba: image::RgbaImage = decoded.to_rgba8();
                    let (w, h) = (rgba.width(), rgba.height());
                    let has_alpha = rgba.pixels().any(|p| p.0[3] < 255);

                    let rgb_data: Vec<u8> = rgba
                        .pixels()
                        .flat_map(|p| [p.0[0], p.0[1], p.0[2]])
                        .collect();
                    let compressed_rgb =
                        miniz_oxide::deflate::compress_to_vec_zlib(&rgb_data, 6);

                    let smask_ref = if has_alpha {
                        let alpha_data: Vec<u8> = rgba.pixels().map(|p| p.0[3]).collect();
                        let compressed_alpha =
                            miniz_oxide::deflate::compress_to_vec_zlib(&alpha_data, 6);
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
                    if let Some(mask_ref) = smask_ref {
                        xobj.s_mask(mask_ref);
                    }
                }
            }
        }

        image_xobjects.push((pdf_name.clone(), xobj_ref));
        pdf_name
    };

    // Key: (global_block_idx, floating_img_idx) → pdf xobject name
    let mut floating_image_pdf_names: HashMap<(usize, usize), String> = HashMap::new();
    {
        let mut global_block_idx = 0usize;
        for section in &doc.sections {
            for block in &section.blocks {
                if let Block::Paragraph(para) = block {
                    if let Some(img) = &para.image {
                        let name =
                            embed_image(img, &mut image_xobjects, &mut pdf, &mut alloc);
                        image_pdf_names.insert(global_block_idx, name);
                    }
                    for (run_idx, run) in para.runs.iter().enumerate() {
                        if let Some(img) = &run.inline_image {
                            let name =
                                embed_image(img, &mut image_xobjects, &mut pdf, &mut alloc);
                            inline_image_pdf_names.insert((global_block_idx, run_idx), name);
                        }
                    }
                    for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                        let name = embed_image(&fi.image, &mut image_xobjects, &mut pdf, &mut alloc);
                        floating_image_pdf_names.insert((global_block_idx, fi_idx), name);
                    }
                }
                global_block_idx += 1;
            }
        }
    }

    // Embed header/footer images
    // Key: (section_idx, hf_type, para_idx) for paragraph images
    // Key: (section_idx, hf_type, para_idx, run_idx) for inline images
    // hf_type: 0=header_default, 1=header_first, 2=footer_default, 3=footer_first
    let mut hf_image_names: HashMap<(usize, u8, usize), String> = HashMap::new();
    let mut hf_inline_image_names: HashMap<(usize, u8, usize, usize), String> = HashMap::new();
    {
        let hf_variants: [(u8, fn(&SectionProperties) -> Option<&HeaderFooter>); 4] = [
            (0, |sp| sp.header_default.as_ref()),
            (1, |sp| sp.header_first.as_ref()),
            (2, |sp| sp.footer_default.as_ref()),
            (3, |sp| sp.footer_first.as_ref()),
        ];
        for (si, section) in doc.sections.iter().enumerate() {
            for &(hf_type, accessor) in &hf_variants {
                if let Some(hf) = accessor(&section.properties) {
                    for (pi, para) in hf.paragraphs.iter().enumerate() {
                        if let Some(img) = &para.image {
                            let name = embed_image(img, &mut image_xobjects, &mut pdf, &mut alloc);
                            hf_image_names.insert((si, hf_type, pi), name);
                        }
                        for (ri, run) in para.runs.iter().enumerate() {
                            if let Some(img) = &run.inline_image {
                                let name = embed_image(img, &mut image_xobjects, &mut pdf, &mut alloc);
                                hf_inline_image_names.insert((si, hf_type, pi, ri), name);
                            }
                        }
                    }
                }
            }
        }
    }

    let t_images = t0.elapsed();

    // Pre-compute footnote display order: scan body runs for footnote_id, assign sequential numbers
    let mut footnote_display_order: HashMap<u32, u32> = HashMap::new();
    {
        let mut next_fn_num = 1u32;
        for section in &doc.sections {
            for block in &section.blocks {
                let runs: Box<dyn Iterator<Item = &Run>> = match block {
                    Block::Paragraph(p) => Box::new(p.runs.iter()),
                    Block::Table(t) => Box::new(
                        t.rows.iter()
                            .flat_map(|row| row.cells.iter())
                            .flat_map(|cell| cell.paragraphs.iter())
                            .flat_map(|p| p.runs.iter()),
                    ),
                };
                for run in runs {
                    if let Some(id) = run.footnote_id {
                        if !footnote_display_order.contains_key(&id) {
                            footnote_display_order.insert(id, next_fn_num);
                            next_fn_num += 1;
                        }
                    }
                }
            }
        }
    }

    // Phase 2: build multi-page content streams (section-aware)
    let mut all_contents: Vec<Content> = Vec::new();
    let mut current_content = Content::new();
    let mut prev_space_after: f32 = 0.0;
    let mut all_page_links: Vec<Vec<LinkAnnotation>> = Vec::new();
    let mut current_page_links: Vec<LinkAnnotation> = Vec::new();

    // Per-page footnote tracking
    let mut all_page_footnote_ids: Vec<Vec<u32>> = Vec::new();
    let mut current_page_footnote_ids: Vec<u32> = Vec::new();

    // Track which section each page belongs to, and whether it's the first page of that section
    let mut page_section_indices: Vec<(usize, bool)> = Vec::new();

    // Initialize from first section
    let first_sp = &doc.sections[0].properties;
    let mut cur_sp = first_sp;
    let mut slot_top = cur_sp.page_height - cur_sp.margin_top;
    let mut effective_margin_bottom: f32 = cur_sp.margin_bottom;
    let mut is_first_page_of_section = true;
    let mut global_block_idx: usize = 0;

    for (sect_idx, section) in doc.sections.iter().enumerate() {
        let sp = &section.properties;

        // Section break handling (not for the first section)
        if sect_idx > 0 {
            match sp.break_type {
                SectionBreakType::NextPage | SectionBreakType::OddPage | SectionBreakType::EvenPage => {
                    // Flush current page and start new page with new geometry
                    all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                    all_page_links.push(std::mem::take(&mut current_page_links));
                    all_page_footnote_ids.push(std::mem::take(&mut current_page_footnote_ids));
                    page_section_indices.push((sect_idx - 1, is_first_page_of_section));
                    slot_top = sp.page_height - sp.margin_top;
                    effective_margin_bottom = sp.margin_bottom;
                }
                SectionBreakType::Continuous => {
                    // No forced break; geometry updates on next page
                }
            }
            is_first_page_of_section = true;
        }

        cur_sp = sp;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;

        // Column geometry: vec of (x_offset, width) for each column
        let col_config = sp.columns.as_ref();
        let col_count = col_config.map(|c| c.columns.len()).unwrap_or(1);
        let col_geometry: Vec<(f32, f32)> = if let Some(cfg) = col_config {
            let mut x = sp.margin_left;
            cfg.columns
                .iter()
                .map(|col| {
                    let result = (x, col.width);
                    x += col.width + col.space;
                    result
                })
                .collect()
        } else {
            vec![(sp.margin_left, text_width)]
        };
        let mut current_col: usize = 0;

        let adjacent_para = |idx: usize| -> Option<&crate::model::Paragraph> {
            match section.blocks.get(idx)? {
                Block::Paragraph(p) => Some(p),
                Block::Table(_) => None,
            }
        };

        for (block_idx, block) in section.blocks.iter().enumerate() {
            match block {
                Block::Paragraph(para) => {
                    // Handle explicit page breaks
                    if para.page_break_before {
                        let at_top = (slot_top - (cur_sp.page_height - cur_sp.margin_top)).abs() < 1.0;
                        if !at_top {
                            all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                            all_page_links.push(std::mem::take(&mut current_page_links));
                            all_page_footnote_ids.push(std::mem::take(&mut current_page_footnote_ids));
                            page_section_indices.push((sect_idx, is_first_page_of_section));
                            slot_top = cur_sp.page_height - cur_sp.margin_top;
                            effective_margin_bottom = cur_sp.margin_bottom;
                            is_first_page_of_section = false;
                            current_col = 0;
                        }
                        prev_space_after = 0.0;
                        if is_text_empty(&para.runs) {
                            global_block_idx += 1;
                            continue;
                        }
                    }

                    // Handle explicit column breaks
                    if para.column_break_before && col_count > 1 {
                        if current_col + 1 < col_count {
                            current_col += 1;
                            slot_top = cur_sp.page_height - cur_sp.margin_top;
                            prev_space_after = 0.0;
                        } else {
                            current_col = 0;
                            all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                            all_page_links.push(std::mem::take(&mut current_page_links));
                            all_page_footnote_ids.push(std::mem::take(&mut current_page_footnote_ids));
                            page_section_indices.push((sect_idx, is_first_page_of_section));
                            slot_top = cur_sp.page_height - cur_sp.margin_top;
                            effective_margin_bottom = cur_sp.margin_bottom;
                            is_first_page_of_section = false;
                            prev_space_after = 0.0;
                        }
                    }

                    let next_para = adjacent_para(block_idx + 1);
                    let prev_para = if block_idx > 0 {
                        adjacent_para(block_idx - 1)
                    } else {
                        None
                    };

                    let effective_space_before =
                        if para.contextual_spacing && prev_para.is_some_and(|p| p.contextual_spacing) {
                            0.0
                        } else {
                            para.space_before
                        };
                    let effective_space_after =
                        if para.contextual_spacing && next_para.is_some_and(|p| p.contextual_spacing) {
                            0.0
                        } else {
                            para.space_after
                        };

                    let mut inter_gap = f32::max(prev_space_after, effective_space_before);

                    let (font_size, tallest_lhr, tallest_ar) =
                        tallest_run_metrics(&para.runs, &seen_fonts);
                    let effective_ls = para.line_spacing.unwrap_or(doc.line_spacing);
                    let line_h = resolve_line_h(effective_ls, font_size, tallest_lhr);

                    let (col_x, col_w) = col_geometry[current_col];
                    let para_text_x = col_x + para.indent_left;
                    let para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
                    let label_x = col_x + para.indent_left - para.indent_hanging;
                    let text_hanging = if !para.list_label.is_empty() {
                        0.0
                    } else if para.indent_hanging > 0.0 {
                        para.indent_hanging
                    } else {
                        -para.indent_first_line
                    };

                    // Substitute footnote reference runs with display numbers
                    let has_footnote_refs = para.runs.iter().any(|r| r.footnote_id.is_some());
                    let effective_runs: std::borrow::Cow<'_, Vec<Run>> = if has_footnote_refs {
                        let substituted: Vec<Run> = para
                            .runs
                            .iter()
                            .map(|run| {
                                if let Some(id) = run.footnote_id {
                                    let num = footnote_display_order.get(&id).copied().unwrap_or(0);
                                    let mut r = run.clone();
                                    r.text = num.to_string();
                                    r
                                } else {
                                    run.clone()
                                }
                            })
                            .collect();
                        std::borrow::Cow::Owned(substituted)
                    } else {
                        std::borrow::Cow::Borrowed(&para.runs)
                    };

                    let text_empty = is_text_empty(&effective_runs);
                    let has_tabs = effective_runs.iter().any(|r| r.is_tab);
                    let block_inline_images: HashMap<usize, String> = inline_image_pdf_names
                        .iter()
                        .filter(|((bi, _), _)| *bi == global_block_idx)
                        .map(|((_, ri), name)| (*ri, name.clone()))
                        .collect();
                    let lines = if para.image.is_some() || text_empty {
                        vec![]
                    } else if has_tabs {
                        build_tabbed_line(&effective_runs, &seen_fonts, &para.tab_stops, para.indent_left)
                    } else {
                        build_paragraph_lines(&effective_runs, &seen_fonts, para_text_width, text_hanging, &block_inline_images)
                    };

                    // For lines containing inline images, use the tallest element as line height
                    let max_inline_img_h = lines.iter()
                        .flat_map(|l| l.chunks.iter())
                        .map(|c| c.inline_image_height)
                        .fold(0.0f32, f32::max);

                    let content_h = if para.image.is_some() {
                        para.content_height.max(sp.line_pitch)
                    } else if text_empty {
                        line_h
                    } else if max_inline_img_h > 0.0 {
                        let mut h = 0.0f32;
                        for line in &lines {
                            let img_h = line.chunks.iter()
                                .map(|c| c.inline_image_height)
                                .fold(0.0f32, f32::max);
                            h += if img_h > line_h { img_h } else { line_h };
                        }
                        h
                    } else {
                        let min_lines = 1 + para.extra_line_breaks as usize;
                        lines.len().max(min_lines) as f32 * line_h
                    };

                    let bdr_top_pad = para
                        .borders
                        .top
                        .as_ref()
                        .map(|b| b.space_pt + b.width_pt / 2.0)
                        .unwrap_or(0.0);
                    let bdr_bottom_pad = para
                        .borders
                        .bottom
                        .as_ref()
                        .map(|b| b.space_pt + b.width_pt / 2.0)
                        .unwrap_or(0.0);

                    let needed = inter_gap + bdr_top_pad + content_h;
                    let at_page_top = (slot_top - (cur_sp.page_height - cur_sp.margin_top)).abs() < 1.0;

                    let keep_next_extra = if para.keep_next {
                        let mut extra = 0.0;
                        let mut prev_sa = effective_space_after;
                        let mut i = block_idx + 1;
                        while let Some(next) = adjacent_para(i) {
                            let (nfs, nlhr, _) = tallest_run_metrics(&next.runs, &seen_fonts);
                            let next_inter = f32::max(prev_sa, next.space_before);
                            let next_first_line_h = nlhr.map(|ratio| nfs * ratio).unwrap_or(nfs * 1.2);
                            if !next.keep_next {
                                let next_ls = next.line_spacing.unwrap_or(doc.line_spacing);
                                let next_line_h = resolve_line_h(next_ls, nfs, nlhr);
                                extra += next_inter + next_first_line_h + next_line_h;
                                break;
                            }
                            extra += next_inter + next_first_line_h;
                            prev_sa = next.space_after;
                            i += 1;
                        }
                        extra
                    } else {
                        0.0
                    };

                    if !at_page_top && slot_top - needed - keep_next_extra < effective_margin_bottom {
                        let available = slot_top - inter_gap - effective_margin_bottom;
                        let first_line_h = tallest_lhr
                            .map(|ratio| font_size * ratio)
                            .unwrap_or(font_size);
                        let mut lines_that_fit = if line_h > 0.0 && available >= first_line_h {
                            1 + ((available - first_line_h) / line_h).floor() as usize
                        } else {
                            0
                        };

                        // Reduce to ensure at least 2 lines remain on next page (orphan control)
                        if lines_that_fit > 0 && lines.len().saturating_sub(lines_that_fit) < 2 {
                            lines_that_fit = lines.len().saturating_sub(2);
                        }

                        // keepLines: don't split — move entire paragraph to next column/page
                        if para.keep_lines {
                            lines_that_fit = 0;
                        }

                        if lines_that_fit >= 2 && lines_that_fit < lines.len() {
                            let first_part = &lines[..lines_that_fit];
                            slot_top -= inter_gap;
                            let ascender_ratio = tallest_ar.unwrap_or(0.75);
                            let baseline_y = slot_top - font_size * ascender_ratio;

                            if !para.list_label.is_empty() {
                                let (label_font_name, label_bytes) =
                                    label_for_run(&para.runs[0], &seen_fonts, &para.list_label);
                                current_content
                                    .begin_text()
                                    .set_font(Name(label_font_name.as_bytes()), font_size)
                                    .next_line(label_x, baseline_y)
                                    .show(Str(&label_bytes))
                                    .end_text();
                            }

                            render_paragraph_lines(
                                &mut current_content,
                                first_part,
                                &para.alignment,
                                para_text_x,
                                para_text_width,
                                baseline_y,
                                line_h,
                                lines.len(),
                                0,
                                &mut current_page_links,
                                text_hanging,
                                &seen_fonts,
                            );

                            if current_col + 1 < col_count {
                                current_col += 1;
                                slot_top = cur_sp.page_height - cur_sp.margin_top;
                            } else {
                                current_col = 0;
                                all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                                all_page_links.push(std::mem::take(&mut current_page_links));
                                all_page_footnote_ids.push(std::mem::take(&mut current_page_footnote_ids));
                                page_section_indices.push((sect_idx, is_first_page_of_section));
                                slot_top = cur_sp.page_height - cur_sp.margin_top;
                                effective_margin_bottom = cur_sp.margin_bottom;
                                is_first_page_of_section = false;
                            }

                            let rest = &lines[lines_that_fit..];
                            let rest_content_h = rest.len() as f32 * line_h;
                            let baseline_y2 = slot_top - font_size * ascender_ratio;

                            let (rest_col_x, rest_col_w) = col_geometry[current_col];
                            let rest_text_x = rest_col_x + para.indent_left;
                            let rest_text_width = (rest_col_w - para.indent_left - para.indent_right).max(1.0);

                            render_paragraph_lines(
                                &mut current_content,
                                rest,
                                &para.alignment,
                                rest_text_x,
                                rest_text_width,
                                baseline_y2,
                                line_h,
                                lines.len(),
                                lines_that_fit,
                                &mut current_page_links,
                                text_hanging,
                                &seen_fonts,
                            );

                            slot_top -= rest_content_h;
                            prev_space_after = effective_space_after;
                            global_block_idx += 1;
                            continue;
                        }

                        if current_col + 1 < col_count {
                            current_col += 1;
                            slot_top = cur_sp.page_height - cur_sp.margin_top;
                        } else {
                            current_col = 0;
                            all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                            all_page_links.push(std::mem::take(&mut current_page_links));
                            all_page_footnote_ids.push(std::mem::take(&mut current_page_footnote_ids));
                            page_section_indices.push((sect_idx, is_first_page_of_section));
                            slot_top = cur_sp.page_height - cur_sp.margin_top;
                            effective_margin_bottom = cur_sp.margin_bottom;
                            is_first_page_of_section = false;
                        }
                        inter_gap = 0.0;
                    }

                    // Suppress space_before at the top of a page
                    let at_new_page_top = !all_contents.is_empty()
                        && (slot_top - (cur_sp.page_height - cur_sp.margin_top)).abs() < 1.0;
                    if at_new_page_top {
                        if is_first_page_of_section {
                            // Section break: collapse with the previous section's trailing space_after
                            inter_gap = (effective_space_before - prev_space_after).max(0.0);
                        } else {
                            inter_gap = 0.0;
                        }
                    }

                    slot_top -= inter_gap;

                    // Re-fetch column geometry (may have changed after overflow)
                    let (col_x, col_w) = col_geometry[current_col];
                    let para_text_x = col_x + para.indent_left;
                    let para_text_width = (col_w - para.indent_left - para.indent_right).max(1.0);
                    let label_x = col_x + para.indent_left - para.indent_hanging;

                    // Draw paragraph shading (background), extending outward to match borders
                    if let Some([r, g, b]) = para.shading {
                        let shd_left_outset = para.borders.left.as_ref().map(|b| b.space_pt).unwrap_or(0.0);
                        let shd_right_outset = para.borders.right.as_ref().map(|b| b.space_pt).unwrap_or(0.0);
                        let shd_left = col_x - shd_left_outset;
                        let shd_right = col_x + col_w + shd_right_outset;
                        let shd_top = slot_top;
                        let shd_bottom = slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
                        current_content.save_state();
                        current_content
                            .set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                        current_content.rect(
                            shd_left,
                            shd_bottom,
                            shd_right - shd_left,
                            shd_top - shd_bottom,
                        );
                        current_content.fill_nonzero();
                        current_content.restore_state();
                    }

                    for (fi_idx, fi) in para.floating_images.iter().enumerate() {
                        if let Some(pdf_name) = floating_image_pdf_names.get(&(global_block_idx, fi_idx)) {
                            let img = &fi.image;
                            let fi_x = match fi.h_relative_from {
                                "page" => match fi.h_position {
                                    HorizontalPosition::AlignCenter => (sp.page_width - img.display_width) / 2.0,
                                    HorizontalPosition::AlignRight => sp.page_width - img.display_width,
                                    HorizontalPosition::AlignLeft => 0.0,
                                    HorizontalPosition::Offset(o) => o,
                                },
                                "column" => match fi.h_position {
                                    HorizontalPosition::AlignCenter => {
                                        col_x + (col_w - img.display_width) / 2.0
                                    }
                                    HorizontalPosition::AlignRight => {
                                        col_x + col_w - img.display_width
                                    }
                                    HorizontalPosition::AlignLeft => col_x,
                                    HorizontalPosition::Offset(o) => col_x + o,
                                },
                                "margin" | _ => match fi.h_position {
                                    HorizontalPosition::AlignCenter => {
                                        sp.margin_left + (text_width - img.display_width) / 2.0
                                    }
                                    HorizontalPosition::AlignRight => {
                                        sp.margin_left + text_width - img.display_width
                                    }
                                    HorizontalPosition::AlignLeft => sp.margin_left,
                                    HorizontalPosition::Offset(o) => sp.margin_left + o,
                                },
                            };
                            // OOXML: positive offset = downward; PDF: Y increases upward
                            let fi_y_top = match fi.v_relative_from {
                                "page" => sp.page_height - fi.v_offset_pt,
                                "margin" | "topMargin" => sp.page_height - sp.margin_top - fi.v_offset_pt,
                                _ => slot_top - fi.v_offset_pt,
                            };
                            let fi_y_bottom = fi_y_top - img.display_height;
                            current_content.save_state();
                            current_content.transform([
                                img.display_width,
                                0.0,
                                0.0,
                                img.display_height,
                                fi_x,
                                fi_y_bottom,
                            ]);
                            current_content.x_object(Name(pdf_name.as_bytes()));
                            current_content.restore_state();
                        }
                    }

                    for tb in &para.textboxes {
                        let tb_x = match tb.h_relative_from {
                            "page" => match tb.h_position {
                                HorizontalPosition::AlignCenter => (sp.page_width - tb.width_pt) / 2.0,
                                HorizontalPosition::AlignRight => sp.page_width - tb.width_pt,
                                HorizontalPosition::AlignLeft => 0.0,
                                HorizontalPosition::Offset(o) => o,
                            },
                            "column" => match tb.h_position {
                                HorizontalPosition::AlignCenter => col_x + (col_w - tb.width_pt) / 2.0,
                                HorizontalPosition::AlignRight => col_x + col_w - tb.width_pt,
                                HorizontalPosition::AlignLeft => col_x,
                                HorizontalPosition::Offset(o) => col_x + o,
                            },
                            "margin" | _ => match tb.h_position {
                                HorizontalPosition::AlignCenter => sp.margin_left + (text_width - tb.width_pt) / 2.0,
                                HorizontalPosition::AlignRight => sp.margin_left + text_width - tb.width_pt,
                                HorizontalPosition::AlignLeft => sp.margin_left,
                                HorizontalPosition::Offset(o) => sp.margin_left + o,
                            },
                        };
                        let tb_y_top = match tb.v_relative_from {
                            "page" => sp.page_height - tb.v_offset_pt,
                            "margin" | "topMargin" => sp.page_height - sp.margin_top - tb.v_offset_pt,
                            _ => slot_top - tb.v_offset_pt,
                        };

                        let mut cursor_y = tb_y_top;
                        let empty_inline_imgs: HashMap<usize, String> = HashMap::new();
                        for tp in &tb.paragraphs {
                            let tp_ls = tp.line_spacing.unwrap_or(doc.line_spacing);
                            let has_tabs = tp.runs.iter().any(|r| r.is_tab);
                            let tb_lines = if has_tabs {
                                build_tabbed_line(
                                    &tp.runs,
                                    &seen_fonts,
                                    &tp.tab_stops,
                                    0.0,
                                )
                            } else {
                                build_paragraph_lines(
                                    &tp.runs,
                                    &seen_fonts,
                                    tb.width_pt,
                                    0.0,
                                    &empty_inline_imgs,
                                )
                            };
                            if tb_lines.is_empty() {
                                let (fs, _, _) = tallest_run_metrics(&tp.runs, &seen_fonts);
                                let lh = resolve_line_h(tp_ls, fs, None);
                                cursor_y -= tp.space_before + lh + tp.space_after;
                                continue;
                            }
                            let (tb_fs, _, tb_ar) = tallest_run_metrics(&tp.runs, &seen_fonts);
                            let tb_ascender = tb_ar.unwrap_or(0.75);
                            let tb_line_h = resolve_line_h(tp_ls, tb_fs, tb_ar);
                            let tb_baseline = cursor_y - tp.space_before - tb_fs * tb_ascender;
                            render_paragraph_lines(
                                &mut current_content,
                                &tb_lines,
                                &tp.alignment,
                                tb_x,
                                tb.width_pt,
                                tb_baseline,
                                tb_line_h,
                                tb_lines.len(),
                                0,
                                &mut current_page_links,
                                0.0,
                                &seen_fonts,
                            );
                            cursor_y -= tp.space_before + (tb_lines.len() as f32) * tb_line_h + tp.space_after;
                        }
                    }

                    if (para.image.is_some() || text_empty) && para.content_height > 0.0 {
                        if let Some(pdf_name) = image_pdf_names.get(&global_block_idx) {
                            let img = para.image.as_ref().unwrap();
                            let y_bottom = slot_top - img.display_height;
                            let x = col_x
                                + match para.alignment {
                                    Alignment::Center => {
                                        (col_w - img.display_width).max(0.0) / 2.0
                                    }
                                    Alignment::Right => (col_w - img.display_width).max(0.0),
                                    _ => 0.0,
                                };
                            current_content.save_state();
                            current_content.transform([
                                img.display_width,
                                0.0,
                                0.0,
                                img.display_height,
                                x,
                                y_bottom,
                            ]);
                            current_content.x_object(Name(pdf_name.as_bytes()));
                            current_content.restore_state();
                        } else {
                            current_content
                                .set_fill_gray(0.5)
                                .rect(col_x, slot_top - content_h, col_w, content_h)
                                .fill_nonzero()
                                .set_fill_gray(0.0);
                        }
                    } else if !lines.is_empty() {
                        let ascender_ratio = tallest_ar.unwrap_or(0.75);
                        let baseline_y = slot_top - bdr_top_pad - font_size * ascender_ratio;

                        if !para.list_label.is_empty() {
                            let (label_font_name, label_bytes) =
                                label_for_run(&para.runs[0], &seen_fonts, &para.list_label);
                            current_content
                                .begin_text()
                                .set_font(Name(label_font_name.as_bytes()), font_size)
                                .next_line(label_x, baseline_y)
                                .show(Str(&label_bytes))
                                .end_text();
                        }

                        render_paragraph_lines(
                            &mut current_content,
                            &lines,
                            &para.alignment,
                            para_text_x,
                            para_text_width,
                            baseline_y,
                            line_h,
                            lines.len(),
                            0,
                            &mut current_page_links,
                            text_hanging,
                            &seen_fonts,
                        );
                    }

                    // Draw paragraph borders — left/right borders extend outward
                    // from the text area so text inside stays aligned with text outside
                    {
                        let bdr = &para.borders;
                        let box_top = slot_top;
                        let box_bottom = slot_top - bdr_top_pad - content_h - bdr_bottom_pad;
                        let bdr_left_outset = bdr.left.as_ref().map(|b| b.space_pt + b.width_pt / 2.0).unwrap_or(0.0);
                        let bdr_right_outset = bdr.right.as_ref().map(|b| b.space_pt + b.width_pt / 2.0).unwrap_or(0.0);
                        let box_left = col_x - bdr_left_outset;
                        let box_right = col_x + col_w + bdr_right_outset;

                        let draw_h_border =
                            |content: &mut Content, b: &crate::model::ParagraphBorder, y: f32| {
                                let [r, g, b_c] = b.color;
                                content.save_state();
                                content.set_line_width(b.width_pt);
                                content.set_stroke_rgb(
                                    r as f32 / 255.0,
                                    g as f32 / 255.0,
                                    b_c as f32 / 255.0,
                                );
                                content.move_to(box_left, y);
                                content.line_to(box_right, y);
                                content.stroke();
                                content.restore_state();
                            };
                        let draw_v_border = |content: &mut Content,
                                             b: &crate::model::ParagraphBorder,
                                             x: f32| {
                            let [r, g, b_c] = b.color;
                            content.save_state();
                            content.set_line_width(b.width_pt);
                            content.set_stroke_rgb(
                                r as f32 / 255.0,
                                g as f32 / 255.0,
                                b_c as f32 / 255.0,
                            );
                            content.move_to(x, box_top);
                            content.line_to(x, box_bottom);
                            content.stroke();
                            content.restore_state();
                        };

                        let prev_has_between = prev_para.is_some_and(|pp| {
                            pp.borders.between.is_some() && borders_match(&pp.borders, &para.borders)
                        });
                        let next_has_between = next_para.is_some_and(|np| {
                            bdr.between.is_some() && borders_match(&para.borders, &np.borders)
                        });

                        if !prev_has_between {
                            if let Some(b) = &bdr.top {
                                draw_h_border(&mut current_content, b, box_top);
                            }
                        }
                        if next_has_between {
                            if let Some(b) = &bdr.between {
                                draw_h_border(&mut current_content, b, box_bottom);
                            }
                        } else if let Some(b) = &bdr.bottom {
                            draw_h_border(&mut current_content, b, box_bottom);
                        }
                        if let Some(b) = &bdr.left {
                            draw_v_border(&mut current_content, b, box_left);
                        }
                        if let Some(b) = &bdr.right {
                            draw_v_border(&mut current_content, b, box_right);
                        }
                    }

                    slot_top -= content_h + bdr_top_pad;
                    prev_space_after = effective_space_after;

                    // Track footnotes referenced on this page
                    for run in para.runs.iter() {
                        if let Some(id) = run.footnote_id {
                            if !current_page_footnote_ids.contains(&id) {
                                current_page_footnote_ids.push(id);
                                if let Some(footnote) = doc.footnotes.get(&id) {
                                    let fn_height = compute_footnote_height(
                                        footnote, &seen_fonts, text_width, doc.line_spacing,
                                    );
                                    let separator_h = if current_page_footnote_ids.len() == 1 {
                                        12.0
                                    } else {
                                        0.0
                                    };
                                    effective_margin_bottom += separator_h + fn_height;
                                }
                            }
                        }
                    }
                }

                Block::Table(table) => {
                    let override_pos = table.position.as_ref().map(|pos| {
                        let table_total_w: f32 = table.col_widths.iter().sum();
                        let x = match pos.h_anchor {
                            "page" => match pos.h_position {
                                HorizontalPosition::AlignCenter => (sp.page_width - table_total_w) / 2.0,
                                HorizontalPosition::AlignRight => sp.page_width - table_total_w,
                                HorizontalPosition::AlignLeft => 0.0,
                                HorizontalPosition::Offset(o) => o,
                            },
                            "margin" => match pos.h_position {
                                HorizontalPosition::AlignCenter => sp.margin_left + (text_width - table_total_w) / 2.0,
                                HorizontalPosition::AlignRight => sp.margin_left + text_width - table_total_w,
                                HorizontalPosition::AlignLeft => sp.margin_left,
                                HorizontalPosition::Offset(o) => sp.margin_left + o,
                            },
                            _ => {
                                let (col_x, col_w) = col_geometry[current_col];
                                match pos.h_position {
                                    HorizontalPosition::AlignCenter => col_x + (col_w - table_total_w) / 2.0,
                                    HorizontalPosition::AlignRight => col_x + col_w - table_total_w,
                                    HorizontalPosition::AlignLeft => col_x,
                                    HorizontalPosition::Offset(o) => col_x + o,
                                }
                            }
                        };
                        let y = match pos.v_anchor {
                            "page" => sp.page_height - pos.v_offset_pt,
                            "margin" => sp.page_height - sp.margin_top - pos.v_offset_pt,
                            _ => slot_top - pos.v_offset_pt,
                        };
                        (x, y)
                    });
                    render_table(
                        table,
                        sp,
                        doc.line_spacing,
                        &seen_fonts,
                        &mut current_content,
                        &mut all_contents,
                        &mut all_page_links,
                        &mut current_page_links,
                        &mut page_section_indices,
                        sect_idx,
                        &mut is_first_page_of_section,
                        &mut slot_top,
                        prev_space_after,
                        override_pos,
                    );
                    if override_pos.is_none() {
                        prev_space_after = 0.0;
                    }
                }
            }
            global_block_idx += 1;
        }
    }
    all_contents.push(current_content);
    all_page_links.push(current_page_links);
    all_page_footnote_ids.push(current_page_footnote_ids);
    page_section_indices.push((doc.sections.len() - 1, is_first_page_of_section));

    let t_layout = t0.elapsed();

    // Phase 2b: render headers and footers on each page (per-section)
    let total_pages = all_contents.len();
    // Pad section indices if table renderer added pages
    while page_section_indices.len() < total_pages {
        let last = page_section_indices.last().map(|&(si, _)| si).unwrap_or(0);
        page_section_indices.push((last, false));
    }

    let build_hf_maps = |si: usize, hf_type: u8| -> (HashMap<usize, String>, HashMap<(usize, usize), String>) {
        let para_imgs: HashMap<usize, String> = hf_image_names
            .iter()
            .filter(|((s, ht, _), _)| *s == si && *ht == hf_type)
            .map(|((_, _, pi), name)| (*pi, name.clone()))
            .collect();
        let inline_imgs: HashMap<(usize, usize), String> = hf_inline_image_names
            .iter()
            .filter(|((s, ht, _, _), _)| *s == si && *ht == hf_type)
            .map(|((_, _, pi, ri), name)| ((*pi, *ri), name.clone()))
            .collect();
        (para_imgs, inline_imgs)
    };

    for (page_idx, content) in all_contents.iter_mut().enumerate() {
        let (si, is_first) = page_section_indices[page_idx];
        let sp = &doc.sections[si].properties;
        let page_num = page_idx + 1;

        // Header
        let (header, hdr_type) = if is_first && sp.different_first_page {
            (sp.header_first.as_ref(), 1u8)
        } else {
            (sp.header_default.as_ref(), 0u8)
        };
        if let Some(hf) = header {
            let (pi_map, ii_map) = build_hf_maps(si, hdr_type);
            render_header_footer(content, hf, &seen_fonts, sp, doc.line_spacing, true, page_num, total_pages, &pi_map, &ii_map);
        }

        // Footer
        let (footer, ftr_type) = if is_first && sp.different_first_page {
            (sp.footer_first.as_ref(), 3u8)
        } else {
            (sp.footer_default.as_ref(), 2u8)
        };
        if let Some(hf) = footer {
            let (pi_map, ii_map) = build_hf_maps(si, ftr_type);
            render_header_footer(content, hf, &seen_fonts, sp, doc.line_spacing, false, page_num, total_pages, &pi_map, &ii_map);
        }

        // Column separator lines
        if let Some(cfg) = &sp.columns {
            if cfg.sep {
                let mut x = sp.margin_left;
                for (i, col) in cfg.columns.iter().enumerate() {
                    x += col.width;
                    if i < cfg.columns.len() - 1 {
                        let mid_x = x + col.space / 2.0;
                        content.save_state();
                        content.set_line_width(0.5);
                        content.move_to(mid_x, sp.margin_bottom);
                        content.line_to(mid_x, sp.page_height - sp.margin_top);
                        content.stroke();
                        content.restore_state();
                        x += col.space;
                    }
                }
            }
        }
    }

    // Phase 2c: render footnotes at page bottom
    while all_page_footnote_ids.len() < all_contents.len() {
        all_page_footnote_ids.push(Vec::new());
    }
    for (page_idx, content) in all_contents.iter_mut().enumerate() {
        let fn_ids = &all_page_footnote_ids[page_idx];
        if fn_ids.is_empty() {
            continue;
        }

        let (si, _) = page_section_indices[page_idx];
        let sp = &doc.sections[si].properties;
        let text_width = sp.page_width - sp.margin_left - sp.margin_right;

        // Compute total footnote block height
        let mut total_fn_height = 0.0f32;
        for fn_id in fn_ids {
            if let Some(footnote) = doc.footnotes.get(fn_id) {
                total_fn_height +=
                    compute_footnote_height(footnote, &seen_fonts, text_width, doc.line_spacing);
            }
        }
        let separator_gap = 12.0f32;
        let block_top = sp.margin_bottom + total_fn_height + separator_gap;

        // Draw separator line: 0.5pt black, ~1/3 page width
        let sep_y = block_top - 3.0;
        let sep_width = 144.0f32.min(text_width);
        content.save_state();
        content.set_line_width(0.5);
        content.move_to(sp.margin_left, sep_y);
        content.line_to(sp.margin_left + sep_width, sep_y);
        content.stroke();
        content.restore_state();

        // Render footnote paragraphs top-down from below separator
        let mut fn_y = sep_y - 9.0;
        for fn_id in fn_ids {
            let Some(footnote) = doc.footnotes.get(fn_id) else {
                continue;
            };
            let display_num = footnote_display_order.get(fn_id).copied().unwrap_or(0);

            for para in &footnote.paragraphs {
                let substituted_runs: Vec<Run> = para
                    .runs
                    .iter()
                    .map(|run| {
                        if run.is_footnote_ref_mark {
                            let mut r = run.clone();
                            r.text = display_num.to_string();
                            r
                        } else {
                            run.clone()
                        }
                    })
                    .collect();

                if is_text_empty(&substituted_runs) {
                    continue;
                }

                let (fs, tallest_lhr, tallest_ar) =
                    tallest_run_metrics(&substituted_runs, &seen_fonts);
                let effective_ls = para.line_spacing.unwrap_or(LineSpacing::Auto(1.0));
                let lh = resolve_line_h(effective_ls, fs, tallest_lhr);

                let lines = build_paragraph_lines(
                    &substituted_runs,
                    &seen_fonts,
                    text_width,
                    0.0,
                    &HashMap::new(),
                );

                if lines.is_empty() {
                    continue;
                }

                let ascender_ratio = tallest_ar.unwrap_or(0.75);
                let baseline_y = fn_y - fs * ascender_ratio;

                render_paragraph_lines(
                    content,
                    &lines,
                    &para.alignment,
                    sp.margin_left,
                    text_width,
                    baseline_y,
                    lh,
                    lines.len(),
                    0,
                    &mut Vec::new(),
                    0.0,
                    &seen_fonts,
                );

                fn_y -= lines.len() as f32 * lh;
            }
        }
    }

    let t_headers = t0.elapsed();

    // Phase 3: allocate page and content IDs now that page count is known
    let n = all_contents.len();
    let page_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();

    // Allocate annotation refs and write annotation objects
    let page_annot_refs: Vec<Vec<Ref>> = all_page_links
        .iter()
        .map(|links| {
            links
                .iter()
                .map(|link| {
                    let annot_ref = alloc();
                    let mut annot = pdf.annotation(annot_ref);
                    annot
                        .subtype(pdf_writer::types::AnnotationType::Link)
                        .rect(link.rect)
                        .border(0.0, 0.0, 0.0, None);
                    annot
                        .action()
                        .action_type(pdf_writer::types::ActionType::Uri)
                        .uri(Str(link.url.as_bytes()));
                    annot_ref
                })
                .collect()
        })
        .collect();

    for (i, c) in all_contents.into_iter().enumerate() {
        let raw = c.finish();
        let compressed = miniz_oxide::deflate::compress_to_vec_zlib(raw.as_slice(), 6);
        pdf.stream(content_ids[i], &compressed).filter(Filter::FlateDecode);
    }

    pdf.catalog(catalog_id).pages(pages_id);
    pdf.pages(pages_id)
        .kids(page_ids.iter().copied())
        .count(n as i32);

    let font_pairs: Vec<(String, Ref)> = font_order
        .iter()
        .map(|name| (seen_fonts[name].pdf_name.clone(), seen_fonts[name].font_ref))
        .collect();

    // Per-page media box from section properties
    for i in 0..n {
        let (si, _) = page_section_indices[i];
        let sp = &doc.sections[si].properties;
        let mut page = pdf.page(page_ids[i]);
        page.media_box(Rect::new(0.0, 0.0, sp.page_width, sp.page_height))
            .parent(pages_id)
            .contents(content_ids[i]);
        if !page_annot_refs[i].is_empty() {
            page.annotations(page_annot_refs[i].iter().copied());
        }
        {
            let mut resources = page.resources();
            {
                let mut fonts = resources.fonts();
                for (name, font_ref) in &font_pairs {
                    fonts.pair(Name(name.as_bytes()), *font_ref);
                }
            }
            if !image_xobjects.is_empty() {
                let mut xobjects = resources.x_objects();
                for (name, xobj_ref) in &image_xobjects {
                    xobjects.pair(Name(name.as_bytes()), *xobj_ref);
                }
            }
        }
    }

    let t_assembly = t0.elapsed();

    log::info!(
        "Render phases: collect_runs={:.1}ms, font_embed={:.1}ms, images={:.1}ms, layout={:.1}ms, headers={:.1}ms, assembly={:.1}ms",
        t_collect.as_secs_f64() * 1000.0,
        (t_fonts - t_collect).as_secs_f64() * 1000.0,
        (t_images - t_fonts).as_secs_f64() * 1000.0,
        (t_layout - t_images).as_secs_f64() * 1000.0,
        (t_headers - t_layout).as_secs_f64() * 1000.0,
        (t_assembly - t_headers).as_secs_f64() * 1000.0,
    );

    Ok(pdf.finish())
}

fn label_for_run<'a>(
    run: &Run,
    seen_fonts: &'a HashMap<String, FontEntry>,
    label: &str,
) -> (&'a str, Vec<u8>) {
    let key = font_key(run);
    let entry = seen_fonts.get(&key).expect("font registered");
    let bytes = match &entry.char_to_gid {
        Some(map) => encode_as_gids(label, map),
        None => to_winansi_bytes(label),
    };
    (entry.pdf_name.as_str(), bytes)
}

fn compute_footnote_height(
    footnote: &Footnote,
    seen_fonts: &HashMap<String, FontEntry>,
    text_width: f32,
    doc_line_spacing: LineSpacing,
) -> f32 {
    let mut height = 0.0f32;
    for para in &footnote.paragraphs {
        if is_text_empty(&para.runs) {
            continue;
        }
        let (fs, tallest_lhr, _) = tallest_run_metrics(&para.runs, seen_fonts);
        let effective_ls = para.line_spacing.unwrap_or(doc_line_spacing);
        let lh = resolve_line_h(effective_ls, fs, tallest_lhr);
        let lines = build_paragraph_lines(&para.runs, seen_fonts, text_width, 0.0, &HashMap::new());
        height += lines.len().max(1) as f32 * lh;
    }
    height
}
