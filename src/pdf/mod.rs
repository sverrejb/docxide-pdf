mod layout;
mod table;

use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str};

use crate::error::Error;
use crate::fonts::{FontEntry, encode_as_gids, font_key, register_font, to_winansi_bytes};
use crate::model::{
    Alignment, Block, Document, FieldCode, HeaderFooter, ImageFormat, Run,
};

use layout::{
    LinkAnnotation,
    build_paragraph_lines, build_tabbed_line,
    font_metric, is_text_empty, render_paragraph_lines, tallest_run_metrics,
};
use table::render_table;

fn render_header_footer(
    content: &mut Content,
    hf: &HeaderFooter,
    seen_fonts: &HashMap<String, FontEntry>,
    doc: &Document,
    is_header: bool,
    page_num: usize,
    total_pages: usize,
) {
    let text_width = doc.page_width - doc.margin_left - doc.margin_right;

    for para in &hf.paragraphs {
        if is_text_empty(&para.runs) {
            continue;
        }

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

        let lines = build_paragraph_lines(&substituted_runs, seen_fonts, text_width, 0.0);

        let (font_size, _, tallest_ar) = tallest_run_metrics(&substituted_runs, seen_fonts);
        let ascender_ratio = tallest_ar.unwrap_or(0.75);

        let baseline_y = if is_header {
            doc.page_height - doc.header_margin - font_size * ascender_ratio
        } else {
            doc.footer_margin + font_size * (1.0 - ascender_ratio)
        };

        let effective_ls = para.line_spacing.unwrap_or(doc.line_spacing);
        let line_h = font_metric(&substituted_runs, seen_fonts, |e| e.line_h_ratio)
            .map(|ratio| font_size * ratio * effective_ls)
            .unwrap_or(font_size * 1.2 * effective_ls);

        render_paragraph_lines(
            content,
            &lines,
            &para.alignment,
            doc.margin_left,
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

    // Collect all runs from all blocks (paragraphs, table cells, headers/footers)
    let hf_options = [
        &doc.header_default,
        &doc.header_first,
        &doc.footer_default,
        &doc.footer_first,
    ];
    let hf_runs = hf_options
        .iter()
        .filter_map(|hf| hf.as_ref())
        .flat_map(|hf| hf.paragraphs.iter())
        .flat_map(|p| p.runs.iter());

    let all_runs: Vec<&Run> = doc
        .blocks
        .iter()
        .flat_map(|block| -> Box<dyn Iterator<Item = &Run> + '_> {
            match block {
                Block::Paragraph(para) => Box::new(para.runs.iter()),
                Block::Table(table) => Box::new(
                    table
                        .rows
                        .iter()
                        .flat_map(|row| row.cells.iter())
                        .flat_map(|cell| cell.paragraphs.iter())
                        .flat_map(|para| para.runs.iter()),
                ),
            }
        })
        .chain(hf_runs)
        .collect();

    let t_collect = t0.elapsed();

    // Collect used characters per font key for subsetting
    let mut used_chars_per_font: HashMap<String, HashSet<char>> = HashMap::new();
    for run in &all_runs {
        let key = font_key(run);
        let chars = used_chars_per_font.entry(key).or_default();
        chars.extend(run.text.chars());
        if let Some(ref fc) = run.field_code {
            match fc {
                FieldCode::Page | FieldCode::NumPages => {
                    chars.extend('0'..='9');
                }
            }
        }
    }
    // List labels and leader characters from paragraphs
    let all_paras = doc.blocks.iter().flat_map(|block| -> Box<dyn Iterator<Item = &crate::model::Paragraph> + '_> {
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
    for hf in hf_options.iter().copied().flatten() {
        for para in &hf.paragraphs {
            for run in &para.runs {
                let key = font_key(run);
                let chars = used_chars_per_font.entry(key).or_default();
                chars.extend(run.text.chars());
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

    let text_width = doc.page_width - doc.margin_left - doc.margin_right;

    // Phase 1b: embed images
    let mut image_pdf_names: HashMap<usize, String> = HashMap::new();
    let mut image_xobjects: Vec<(String, Ref)> = Vec::new();
    for (block_idx, block) in doc.blocks.iter().enumerate() {
        if let Block::Paragraph(para) = block
            && let Some(img) = &para.image
        {
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
                        let rgba = decoded.to_rgba8();
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
                    } else {
                        continue;
                    }
                }
            }

            image_xobjects.push((pdf_name.clone(), xobj_ref));
            image_pdf_names.insert(block_idx, pdf_name);
        }
    }

    let t_images = t0.elapsed();

    // Phase 2: build multi-page content streams
    let mut all_contents: Vec<Content> = Vec::new();
    let mut current_content = Content::new();
    let mut slot_top = doc.page_height - doc.margin_top;
    let mut prev_space_after: f32 = 0.0;
    let mut all_page_links: Vec<Vec<LinkAnnotation>> = Vec::new();
    let mut current_page_links: Vec<LinkAnnotation> = Vec::new();

    let adjacent_para = |idx: usize| -> Option<&crate::model::Paragraph> {
        match doc.blocks.get(idx)? {
            Block::Paragraph(p) => Some(p),
            Block::Table(_) => None,
        }
    };

    for (block_idx, block) in doc.blocks.iter().enumerate() {
        match block {
            Block::Paragraph(para) => {
                // Handle explicit page breaks
                if para.page_break_before {
                    let at_top = (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;
                    if !at_top {
                        all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                        all_page_links.push(std::mem::take(&mut current_page_links));
                        slot_top = doc.page_height - doc.margin_top;
                    }
                    prev_space_after = 0.0;
                    if is_text_empty(&para.runs) {
                        continue;
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
                let effective_line_spacing = para.line_spacing.unwrap_or(doc.line_spacing);
                let line_h = tallest_lhr
                    .map(|ratio| font_size * ratio * effective_line_spacing)
                    .unwrap_or(font_size * 1.2 * effective_line_spacing);

                let para_text_x = doc.margin_left + para.indent_left;
                let para_text_width = (text_width - para.indent_left).max(1.0);
                let label_x = doc.margin_left + para.indent_left - para.indent_hanging;
                // Only apply hanging first-line shift when there's no visible label;
                // with a visible label, the hanging area is for the label only.
                let text_hanging = if para.list_label.is_empty() {
                    para.indent_hanging
                } else {
                    0.0
                };

                let text_empty = is_text_empty(&para.runs);
                let has_tabs = para.runs.iter().any(|r| r.is_tab);
                let lines = if para.image.is_some() || text_empty {
                    vec![]
                } else if has_tabs {
                    build_tabbed_line(&para.runs, &seen_fonts, &para.tab_stops, para.indent_left)
                } else {
                    build_paragraph_lines(&para.runs, &seen_fonts, para_text_width, text_hanging)
                };

                let content_h = if para.image.is_some() {
                    para.content_height.max(doc.line_pitch)
                } else if text_empty {
                    line_h
                } else {
                    let min_lines = 1 + para.extra_line_breaks as usize;
                    lines.len().max(min_lines) as f32 * line_h
                };

                let needed = inter_gap + content_h;
                let at_page_top = (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;

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
                            let next_line_h = nlhr
                                .map(|ratio| nfs * ratio * next_ls)
                                .unwrap_or(nfs * 1.2 * next_ls);
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

                if !at_page_top && slot_top - needed - keep_next_extra < doc.margin_bottom {
                    let available = slot_top - inter_gap - doc.margin_bottom;
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

                        all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                        all_page_links.push(std::mem::take(&mut current_page_links));
                        slot_top = doc.page_height - doc.margin_top;

                        let rest = &lines[lines_that_fit..];
                        let rest_content_h = rest.len() as f32 * line_h;
                        let baseline_y2 = slot_top - font_size * ascender_ratio;

                        render_paragraph_lines(
                            &mut current_content,
                            rest,
                            &para.alignment,
                            para_text_x,
                            para_text_width,
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
                        continue;
                    }

                    all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                    all_page_links.push(std::mem::take(&mut current_page_links));
                    slot_top = doc.page_height - doc.margin_top;
                    inter_gap = 0.0;
                }

                // Suppress space_before at the top of a page (after a page break, not first page)
                let at_new_page_top = !all_contents.is_empty()
                    && (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;
                if at_new_page_top {
                    inter_gap = 0.0;
                }

                slot_top -= inter_gap;

                if (para.image.is_some() || text_empty) && para.content_height > 0.0 {
                    if let Some(pdf_name) = image_pdf_names.get(&block_idx) {
                        let img = para.image.as_ref().unwrap();
                        let y_bottom = slot_top - img.display_height;
                        let x = doc.margin_left
                            + match para.alignment {
                                Alignment::Center => {
                                    (text_width - img.display_width).max(0.0) / 2.0
                                }
                                Alignment::Right => (text_width - img.display_width).max(0.0),
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
                            .rect(doc.margin_left, slot_top - content_h, text_width, content_h)
                            .fill_nonzero()
                            .set_fill_gray(0.0);
                    }
                } else if !lines.is_empty() {
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

                // Draw bottom border if present
                if let Some(bdr) = &para.border_bottom {
                    let line_y = slot_top - content_h - bdr.space_pt;
                    let [r, g, b] = bdr.color;
                    current_content
                        .set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0)
                        .rect(
                            doc.margin_left,
                            line_y - bdr.width_pt,
                            text_width,
                            bdr.width_pt,
                        )
                        .fill_nonzero()
                        .set_fill_rgb(0.0, 0.0, 0.0);
                }

                slot_top -= content_h;
                prev_space_after = effective_space_after;
            }

            Block::Table(table) => {
                render_table(
                    table,
                    doc,
                    &seen_fonts,
                    &mut current_content,
                    &mut all_contents,
                    &mut all_page_links,
                    &mut current_page_links,
                    &mut slot_top,
                    prev_space_after,
                );
                prev_space_after = 0.0;
            }
        }
    }
    all_contents.push(current_content);
    all_page_links.push(current_page_links);

    let t_layout = t0.elapsed();

    // Phase 2b: render headers and footers on each page
    let total_pages = all_contents.len();
    let has_hf = doc.header_default.is_some()
        || doc.header_first.is_some()
        || doc.footer_default.is_some()
        || doc.footer_first.is_some();

    if has_hf {
        for (page_idx, content) in all_contents.iter_mut().enumerate() {
            let is_first = page_idx == 0;
            let page_num = page_idx + 1;

            // Header
            let header = if is_first && doc.different_first_page {
                doc.header_first.as_ref()
            } else {
                doc.header_default.as_ref()
            };
            if let Some(hf) = header {
                render_header_footer(content, hf, &seen_fonts, doc, true, page_num, total_pages);
            }

            // Footer
            let footer = if is_first && doc.different_first_page {
                doc.footer_first.as_ref()
            } else {
                doc.footer_default.as_ref()
            };
            if let Some(hf) = footer {
                render_header_footer(content, hf, &seen_fonts, doc, false, page_num, total_pages);
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

    for i in 0..n {
        let mut page = pdf.page(page_ids[i]);
        page.media_box(Rect::new(0.0, 0.0, doc.page_width, doc.page_height))
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
