use std::collections::HashMap;
use std::path::{Path, PathBuf};

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str};
use ttf_parser::Face;

use crate::error::Error;
use crate::model::{Alignment, Document, Run};

struct FontEntry {
    pdf_name: String,
    font_ref: Ref,
    widths_1000: Vec<f32>, // 224 entries for WinAnsi chars 32..=255
    line_h_ratio: Option<f32>, // (ascender - descender + lineGap) / UPM; None for Type1 fallback
    ascender_ratio: Option<f32>, // ascender / UPM; used to place baseline within the slot
}

/// Search the host system for a TTF/OTF file for the given font name.
/// Checks Microsoft Office bundled fonts and system font directories.
fn find_font_file(font_name: &str) -> Option<PathBuf> {
    let normalized = font_name.replace(' ', "");

    // 1. Office CloudFonts (e.g. Aptos Display is downloaded here on Mac)
    if let Ok(home) = std::env::var("HOME") {
        let cloud_dir = PathBuf::from(&home)
            .join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts")
            .join(font_name);
        if cloud_dir.is_dir()
            && let Ok(entries) = std::fs::read_dir(&cloud_dir)
        {
            for entry in entries.flatten() {
                let p = entry.path();
                if matches!(
                    p.extension().and_then(|e| e.to_str()),
                    Some("ttf") | Some("otf")
                ) {
                    log::debug!("find_font_file({:?}) -> CloudFonts {:?}", font_name, p);
                    return Some(p);
                }
            }
        }
    }

    // 2. Word app DFonts (Mac)
    let word_fonts =
        Path::new("/Applications/Microsoft Word.app/Contents/Resources/DFonts");
    for ext in &["ttf", "otf"] {
        let p = word_fonts.join(format!("{}.{}", normalized, ext));
        if p.exists() {
            log::debug!("find_font_file({:?}) -> DFonts {:?}", font_name, p);
            return Some(p);
        }
    }

    // 3. System font directories
    let system_dirs = [
        "/Library/Fonts",
        "/Library/Fonts/Microsoft",
        "/System/Library/Fonts",
        "/System/Library/Fonts/Supplemental",
    ];
    for name_variant in &[normalized.as_str(), font_name] {
        for dir in &system_dirs {
            for ext in &["ttf", "otf"] {
                let p = Path::new(dir).join(format!("{}.{}", name_variant, ext));
                if p.exists() {
                    return Some(p);
                }
            }
        }
    }

    log::debug!("find_font_file({:?}) -> None", font_name);
    None
}

/// Windows-1252 (WinAnsi) byte to Unicode char mapping.
/// Bytes 0x80-0x9F are remapped; all others map directly to their Unicode codepoint.
fn winansi_to_char(byte: u8) -> char {
    match byte {
        0x80 => '\u{20AC}',
        0x82 => '\u{201A}',
        0x83 => '\u{0192}',
        0x84 => '\u{201E}',
        0x85 => '\u{2026}',
        0x86 => '\u{2020}',
        0x87 => '\u{2021}',
        0x88 => '\u{02C6}',
        0x89 => '\u{2030}',
        0x8A => '\u{0160}',
        0x8B => '\u{2039}',
        0x8C => '\u{0152}',
        0x8E => '\u{017D}',
        0x91 => '\u{2018}',
        0x92 => '\u{2019}',
        0x93 => '\u{201C}',
        0x94 => '\u{201D}',
        0x95 => '\u{2022}', // bullet
        0x96 => '\u{2013}',
        0x97 => '\u{2014}',
        0x98 => '\u{02DC}',
        0x99 => '\u{2122}',
        0x9A => '\u{0161}',
        0x9B => '\u{203A}',
        0x9C => '\u{0153}',
        0x9E => '\u{017E}',
        0x9F => '\u{0178}',
        _ => byte as char,
    }
}

/// Convert a UTF-8 string to WinAnsi (Windows-1252) bytes for PDF Str encoding.
fn to_winansi_bytes(s: &str) -> Vec<u8> {
    s.chars()
        .filter_map(|c| match c as u32 {
            0x0000..=0x007F => Some(c as u8),
            0x00A0..=0x00FF => Some(c as u8), // Latin-1 supplement maps directly
            0x20AC => Some(0x80),
            0x201A => Some(0x82),
            0x0192 => Some(0x83),
            0x201E => Some(0x84),
            0x2026 => Some(0x85),
            0x2020 => Some(0x86),
            0x2021 => Some(0x87),
            0x02C6 => Some(0x88),
            0x2030 => Some(0x89),
            0x0160 => Some(0x8A),
            0x2039 => Some(0x8B),
            0x0152 => Some(0x8C),
            0x017D => Some(0x8E),
            0x2018 => Some(0x91),
            0x2019 => Some(0x92),
            0x201C => Some(0x93),
            0x201D => Some(0x94),
            0x2022 => Some(0x95), // bullet
            0x2013 => Some(0x96),
            0x2014 => Some(0x97),
            0x02DC => Some(0x98),
            0x2122 => Some(0x99),
            0x0161 => Some(0x9A),
            0x203A => Some(0x9B),
            0x0153 => Some(0x9C),
            0x017E => Some(0x9E),
            0x0178 => Some(0x9F),
            _ => None,
        })
        .collect()
}

/// Approximate Helvetica widths at 1000 units/em for WinAnsi chars 32..=255.
fn helvetica_widths() -> Vec<f32> {
    (32u8..=255u8)
        .map(|b| match b {
            32 => 278.0,          // space
            33..=47 => 333.0,     // punctuation
            48..=57 => 556.0,     // digits
            58..=64 => 333.0,     // more punctuation
            73 | 74 => 278.0,     // I J (narrow uppercase)
            77 => 833.0,          // M (wide)
            65..=90 => 667.0,     // uppercase A-Z (average)
            91..=96 => 333.0,     // brackets etc.
            102 | 105 | 106 | 108 | 116 => 278.0, // narrow lowercase: f i j l t
            109 | 119 => 833.0,   // m w (wide)
            97..=122 => 556.0,    // lowercase a-z (average)
            _ => 556.0,
        })
        .collect()
}

/// Attempt to embed a TrueType/OpenType font into the PDF.
fn embed_truetype(
    pdf: &mut Pdf,
    font_ref: Ref,
    descriptor_ref: Ref,
    data_ref: Ref,
    font_name: &str,
    font_path: &Path,
) -> Option<(Vec<f32>, f32, f32)> {
    let font_data = std::fs::read(font_path).ok()?;
    let face = Face::parse(&font_data, 0).ok()?;

    let units = face.units_per_em() as f32;
    let ascent = face.ascender() as f32 / units * 1000.0;
    let descent = face.descender() as f32 / units * 1000.0;
    let cap_height = face
        .capital_height()
        .map(|h| h as f32 / units * 1000.0)
        .unwrap_or(700.0);

    let bb = face.global_bounding_box();
    let bbox = Rect::new(
        bb.x_min as f32 / units * 1000.0,
        bb.y_min as f32 / units * 1000.0,
        bb.x_max as f32 / units * 1000.0,
        bb.y_max as f32 / units * 1000.0,
    );

    let widths: Vec<f32> = (32u8..=255u8)
        .map(|byte| {
            face.glyph_index(winansi_to_char(byte))
                .and_then(|gid| face.glyph_hor_advance(gid))
                .map(|adv| adv as f32 / units * 1000.0)
                .unwrap_or(0.0)
        })
        .collect();

    let data_len = i32::try_from(font_data.len()).ok()?;
    pdf.stream(data_ref, &font_data).pair(Name(b"Length1"), data_len);

    let ps_name = font_name.replace(' ', "");

    pdf.font_descriptor(descriptor_ref)
        .name(Name(ps_name.as_bytes()))
        .flags(pdf_writer::types::FontFlags::NON_SYMBOLIC)
        .bbox(bbox)
        .italic_angle(0.0)
        .ascent(ascent)
        .descent(descent)
        .cap_height(cap_height)
        .stem_v(80.0)
        .font_file2(data_ref);

    {
        let mut d = pdf.indirect(font_ref).dict();
        d.pair(Name(b"Type"), Name(b"Font"));
        d.pair(Name(b"Subtype"), Name(b"TrueType"));
        d.pair(Name(b"BaseFont"), Name(ps_name.as_bytes()));
        d.pair(Name(b"Encoding"), Name(b"WinAnsiEncoding"));
        d.pair(Name(b"FirstChar"), 32i32);
        d.pair(Name(b"LastChar"), 255i32);
        d.pair(Name(b"FontDescriptor"), descriptor_ref);
        d.insert(Name(b"Widths")).array().items(widths.iter().copied());
    }

    let line_gap = face.line_gap() as f32;
    let line_h_ratio = (face.ascender() as f32 - face.descender() as f32 + line_gap) / units;
    let ascender_ratio = face.ascender() as f32 / units;

    Some((widths, line_h_ratio, ascender_ratio))
}

fn primary_font_name(name: &str) -> &str {
    name.split(';').next().unwrap_or(name).trim()
}

fn register_font(
    pdf: &mut Pdf,
    font_name: &str,
    pdf_name: String,
    alloc: &mut impl FnMut() -> Ref,
) -> FontEntry {
    let font_ref = alloc();
    let descriptor_ref = alloc();
    let data_ref = alloc();

    let (widths, line_h_ratio, ascender_ratio) = find_font_file(font_name)
        .and_then(|path| {
            embed_truetype(pdf, font_ref, descriptor_ref, data_ref, font_name, &path)
        })
        .map(|(w, r, ar)| (w, Some(r), Some(ar)))
        .unwrap_or_else(|| {
            pdf.type1_font(font_ref)
                .base_font(Name(b"Helvetica"))
                .encoding_predefined(Name(b"WinAnsiEncoding"));
            (helvetica_widths(), None, None)
        });

    FontEntry { pdf_name, font_ref, widths_1000: widths, line_h_ratio, ascender_ratio }
}

struct WordChunk {
    pdf_font: String,
    text: String,
    font_size: f32,
    color: Option<[u8; 3]>,
    x_offset: f32, // x relative to line start
    width: f32,
}

struct TextLine {
    chunks: Vec<WordChunk>,
    total_width: f32,
}

fn finish_line(chunks: &mut Vec<WordChunk>) -> TextLine {
    let total_width = chunks
        .last()
        .map(|c| c.x_offset + c.width)
        .unwrap_or(0.0);
    TextLine { chunks: std::mem::take(chunks), total_width }
}

/// Layout runs into wrapped lines.
fn build_paragraph_lines(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    max_width: f32,
) -> Vec<TextLine> {
    let mut lines: Vec<TextLine> = Vec::new();
    let mut current_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;

    for run in runs {
        let key = primary_font_name(&run.font_name);
        let entry = seen_fonts.get(key).expect("font registered");
        let space_w = entry.widths_1000[0] * run.font_size / 1000.0;

        for word in run.text.split_whitespace() {
            let ww: f32 = word
                .bytes()
                .filter(|&b| b >= 32)
                .map(|b| entry.widths_1000[(b - 32) as usize] * run.font_size / 1000.0)
                .sum();

            if !current_chunks.is_empty() && current_x + ww > max_width {
                lines.push(finish_line(&mut current_chunks));
                current_x = 0.0;
            }

            current_chunks.push(WordChunk {
                pdf_font: entry.pdf_name.clone(),
                text: word.to_string(),
                font_size: run.font_size,
                color: run.color,
                x_offset: current_x,
                width: ww,
            });
            current_x += ww + space_w;
        }
    }

    if !current_chunks.is_empty() {
        lines.push(finish_line(&mut current_chunks));
    }

    if lines.is_empty() {
        lines.push(TextLine { chunks: vec![], total_width: 0.0 });
    }
    lines
}

/// Render pre-built lines applying the paragraph alignment.
/// `total_line_count` is the full paragraph line count (for justify: last line stays left-aligned).
fn render_paragraph_lines(
    content: &mut Content,
    lines: &[TextLine],
    alignment: &Alignment,
    margin_left: f32,
    text_width: f32,
    first_baseline_y: f32,
    line_pitch: f32,
    total_line_count: usize,
    first_line_index: usize,
) {
    let mut current_color: Option<[u8; 3]> = None;

    let last_line_idx = total_line_count.saturating_sub(1);
    for (line_num, line) in lines.iter().enumerate() {
        let y = first_baseline_y - line_num as f32 * line_pitch;
        let global_line_idx = first_line_index + line_num;

        let is_justified = *alignment == Alignment::Justify
            && global_line_idx != last_line_idx
            && line.chunks.len() > 1;

        let line_start_x = match alignment {
            Alignment::Center => margin_left + (text_width - line.total_width) / 2.0,
            Alignment::Right => margin_left + text_width - line.total_width,
            Alignment::Left | Alignment::Justify => margin_left,
        };

        let extra_per_gap = if is_justified {
            (text_width - line.total_width) / (line.chunks.len() - 1) as f32
        } else {
            0.0
        };

        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            let x = line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap;
            if chunk.color != current_color {
                if let Some([r, g, b]) = chunk.color {
                    content.set_fill_rgb(
                        r as f32 / 255.0,
                        g as f32 / 255.0,
                        b as f32 / 255.0,
                    );
                } else {
                    content.set_fill_gray(0.0);
                }
                current_color = chunk.color;
            }
            let text_bytes = to_winansi_bytes(&chunk.text);
            content
                .begin_text()
                .set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size)
                .next_line(x, y)
                .show(Str(&text_bytes))
                .end_text();
        }
    }
    if current_color.is_some() {
        content.set_fill_gray(0.0);
    }
}

fn font_metric(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    get: impl Fn(&FontEntry) -> Option<f32>,
) -> Option<f32> {
    runs.first()
        .map(|r| primary_font_name(&r.font_name))
        .and_then(|k| seen_fonts.get(k))
        .and_then(get)
}

pub fn render(doc: &Document) -> Result<Vec<u8>, Error> {
    let mut pdf = Pdf::new();
    let mut next_id = 1i32;
    let mut alloc = || {
        let r = Ref::new(next_id);
        next_id += 1;
        r
    };

    let catalog_id = alloc();
    let pages_id = alloc();

    // Phase 1: collect unique font names and embed them
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();
    for para in &doc.paragraphs {
        for run in &para.runs {
            let key = primary_font_name(&run.font_name).to_string();
            if !seen_fonts.contains_key(&key) {
                let pdf_name = format!("F{}", font_order.len() + 1);
                let entry = register_font(&mut pdf, &key, pdf_name, &mut alloc);
                seen_fonts.insert(key.clone(), entry);
                font_order.push(key);
            }
        }
    }

    if seen_fonts.is_empty() {
        let pdf_name = "F1".to_string();
        let entry = register_font(&mut pdf, "Helvetica", pdf_name, &mut alloc);
        seen_fonts.insert("Helvetica".to_string(), entry);
        font_order.push("Helvetica".to_string());
    }

    let text_width = doc.page_width - doc.margin_left - doc.margin_right;

    // Phase 1b: embed images
    struct ImageRef {
        pdf_name: String,
    }
    let mut image_refs: HashMap<usize, ImageRef> = HashMap::new(); // para_idx -> ImageRef
    let mut image_names: Vec<(String, Ref)> = Vec::new();
    for (para_idx, para) in doc.paragraphs.iter().enumerate() {
        if let Some(img) = &para.image {
            let xobj_ref = alloc();
            let pdf_name = format!("Im{}", image_names.len() + 1);

            let mut xobj = pdf.image_xobject(xobj_ref, &img.data);
            xobj.filter(Filter::DctDecode);
            xobj.width(img.pixel_width as i32);
            xobj.height(img.pixel_height as i32);
            xobj.color_space().device_rgb();
            xobj.bits_per_component(8);

            image_names.push((pdf_name.clone(), xobj_ref));
            image_refs.insert(para_idx, ImageRef { pdf_name });
        }
    }

    // Phase 2: build multi-page content streams
    let mut all_contents: Vec<Content> = Vec::new();
    let mut current_content = Content::new();
    let mut slot_top = doc.page_height - doc.margin_top;
    // Word collapses adjacent paragraph spacing: gap = max(space_after_prev, space_before_next)
    let mut prev_space_after: f32 = 0.0;

    for (para_idx, para) in doc.paragraphs.iter().enumerate() {
        let next_para = doc.paragraphs.get(para_idx + 1);
        let prev_para = if para_idx > 0 { doc.paragraphs.get(para_idx - 1) } else { None };

        let effective_space_before = if para.contextual_spacing
            && prev_para.is_some_and(|p| p.contextual_spacing)
        {
            0.0
        } else {
            para.space_before
        };
        let effective_space_after = if para.contextual_spacing
            && next_para.is_some_and(|p| p.contextual_spacing)
        {
            0.0
        } else {
            para.space_after
        };

        let mut inter_gap = f32::max(prev_space_after, effective_space_before);

        let font_size = para.runs.first().map_or(12.0, |r| r.font_size);
        // When font metrics are available, apply the document's line-spacing factor.
        // For the Helvetica fallback (no metrics), 1.2 is already an approximation;
        // multiplying by doc.line_spacing would overcount and shift subsequent paragraphs.
        let effective_line_spacing = para.line_spacing.unwrap_or(doc.line_spacing);
        let text_line_h = font_metric(&para.runs, &seen_fonts, |e| e.line_h_ratio)
            .map(|ratio| font_size * ratio * effective_line_spacing)
            .unwrap_or(font_size * 1.2);

        let line_h = text_line_h;

        let para_text_x = doc.margin_left + para.indent_left;
        let para_text_width = (text_width - para.indent_left).max(1.0);
        let label_x = doc.margin_left + (para.indent_left - para.indent_hanging).max(0.0);

        let lines = if para.runs.is_empty() {
            vec![]
        } else {
            build_paragraph_lines(&para.runs, &seen_fonts, para_text_width)
        };

        let content_h = if para.runs.is_empty() {
            para.content_height.max(doc.line_pitch)
        } else {
            lines.len() as f32 * line_h
        };
        log::debug!("para font={} size={} lines={} line_h={:.3} inter_gap={:.1} space_a={:.1} slot_top={:.3} content_h={:.3}",
            para.runs.first().map(|r| r.font_name.as_str()).unwrap_or("?"),
            font_size, lines.len(), line_h,
            inter_gap, effective_space_after,
            slot_top, content_h);

        let needed = inter_gap + content_h + effective_space_after;
        let at_page_top = (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;

        if !at_page_top && slot_top - needed < doc.margin_bottom {
            let available = slot_top - inter_gap - doc.margin_bottom;
            let first_line_h = font_metric(&para.runs, &seen_fonts, |e| e.line_h_ratio)
                .map(|ratio| font_size * ratio)
                .unwrap_or(font_size);
            let lines_that_fit = if line_h > 0.0 && available >= first_line_h {
                1 + ((available - first_line_h) / line_h).floor() as usize
            } else {
                0
            };

            if lines_that_fit >= 2 && lines.len() > lines_that_fit + 1 {
                // Split paragraph across pages
                let first_part = &lines[..lines_that_fit];
                slot_top -= inter_gap;
                let ascender_ratio = font_metric(&para.runs, &seen_fonts, |e| e.ascender_ratio)
                    .unwrap_or(0.75);
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
                );

                // Page break
                all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                slot_top = doc.page_height - doc.margin_top;

                // Render remaining lines on new page
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
                );

                slot_top -= rest_content_h;
                prev_space_after = effective_space_after;
                continue;
            }

            all_contents.push(std::mem::replace(&mut current_content, Content::new()));
            slot_top = doc.page_height - doc.margin_top;
            inter_gap = effective_space_before;
        }

        slot_top -= inter_gap;

        if para.runs.is_empty() && para.content_height > 0.0 {
            if let Some(img_ref) = image_refs.get(&para_idx) {
                let img = para.image.as_ref().unwrap();
                let y_bottom = slot_top - img.display_height;
                let x = doc.margin_left + (text_width - img.display_width).max(0.0) / 2.0;
                current_content.save_state();
                current_content.transform([
                    img.display_width, 0.0,
                    0.0, img.display_height,
                    x, y_bottom,
                ]);
                current_content.x_object(Name(img_ref.pdf_name.as_bytes()));
                current_content.restore_state();
            } else {
                current_content
                    .set_fill_gray(0.5)
                    .rect(doc.margin_left, slot_top - content_h, text_width, content_h)
                    .fill_nonzero()
                    .set_fill_gray(0.0);
            }
        } else if !lines.is_empty() {
            let ascender_ratio = font_metric(&para.runs, &seen_fonts, |e| e.ascender_ratio)
                .unwrap_or(0.75);
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
            );
        }

        slot_top -= content_h;
        prev_space_after = effective_space_after;
    }
    all_contents.push(current_content);

    // Phase 3: allocate page and content IDs now that page count is known
    let n = all_contents.len();
    let page_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();

    for (i, c) in all_contents.into_iter().enumerate() {
        pdf.stream(content_ids[i], &c.finish());
    }

    pdf.catalog(catalog_id).pages(pages_id);
    pdf.pages(pages_id).kids(page_ids.iter().copied()).count(n as i32);

    let font_pairs: Vec<(String, Ref)> = font_order
        .iter()
        .map(|name| (seen_fonts[name].pdf_name.clone(), seen_fonts[name].font_ref))
        .collect();

    for i in 0..n {
        let mut page = pdf.page(page_ids[i]);
        page.media_box(Rect::new(0.0, 0.0, doc.page_width, doc.page_height))
            .parent(pages_id)
            .contents(content_ids[i]);
        {
            let mut resources = page.resources();
            {
                let mut fonts = resources.fonts();
                for (name, font_ref) in &font_pairs {
                    fonts.pair(Name(name.as_bytes()), *font_ref);
                }
            }
            if !image_names.is_empty() {
                let mut xobjects = resources.x_objects();
                for (name, xobj_ref) in &image_names {
                    xobjects.pair(Name(name.as_bytes()), *xobj_ref);
                }
            }
        }
    }

    Ok(pdf.finish())
}

fn label_for_run<'a>(
    run: &Run,
    seen_fonts: &'a HashMap<String, FontEntry>,
    label: &str,
) -> (&'a str, Vec<u8>) {
    let key = primary_font_name(&run.font_name);
    let entry = seen_fonts.get(key).expect("font registered");
    (entry.pdf_name.as_str(), to_winansi_bytes(label))
}
