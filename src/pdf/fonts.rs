use std::collections::{BTreeMap, HashMap, HashSet};

use pdf_writer::{Pdf, Ref};

use crate::fonts::{FontContext, FontEntry, font_key_buf, register_font};
use crate::model::{Block, Document, FieldCode, IfPart, ImageFormat, Paragraph, Run};

use super::header_footer::hf_paragraphs;
use super::list_label::label_pieces;
use super::{collect_paras, para_runs_with_textboxes};

/// The font keys the comment pane draws with: the comment text in the body
/// font, the "Commented [Rn]:" label in its bold variant. Aptos or Calibri
/// when the document has either; otherwise pick the body font from the first
/// NON-bold, NON-italic run we can find (headings are typically bold and would
/// otherwise point us at the wrong face); fall back to the first run if
/// nothing else matches. Glyph collection and the pane both ask here, so the
/// pane never draws a character its face wasn't given.
pub(super) fn comment_pane_keys(doc: &Document) -> (String, String) {
    (
        doc.comment_body_font.clone(),
        format!("{}/B", doc.comment_label_font),
    )
}

pub(super) fn collect_all_runs(doc: &Document) -> Vec<&Run> {
    let hf_runs = doc.sections.iter().flat_map(|s| {
        [
            &s.properties.header_default,
            &s.properties.header_first,
            &s.properties.header_even,
            &s.properties.footer_default,
            &s.properties.footer_first,
            &s.properties.footer_even,
        ]
        .into_iter()
        .filter_map(|hf| hf.as_ref())
        .flat_map(|hf| hf_paragraphs(hf))
        .flat_map(|p| para_runs_with_textboxes(p))
    });

    let footnote_runs = doc
        .footnotes
        .values()
        .flat_map(|fn_| fn_.paragraphs.iter())
        .flat_map(|p| para_runs_with_textboxes(p));

    let endnote_runs = doc
        .endnotes
        .values()
        .flat_map(|en| en.paragraphs.iter())
        .flat_map(|p| para_runs_with_textboxes(p));

    doc.sections
        .iter()
        .flat_map(|s| s.blocks.iter())
        .flat_map(|block| -> Vec<&Run> {
            match block {
                Block::Paragraph(para) => para_runs_with_textboxes(para),
                Block::Table(table) => table
                    .rows
                    .iter()
                    .flat_map(|row| row.cells.iter())
                    .flat_map(|cell| cell.all_paragraphs())
                    .flat_map(|para| para_runs_with_textboxes(para))
                    .collect(),
            }
        })
        .chain(hf_runs)
        .chain(footnote_runs)
        .chain(endnote_runs)
        .collect()
}

fn collect_used_chars(
    doc: &Document,
    all_runs: &[&Run],
    math_font: bool,
) -> HashMap<String, HashSet<char>> {
    let mut used: HashMap<String, HashSet<char>> = HashMap::new();
    let mut key_buf = String::new();

    // §17.11.18/.17 mark numbering formats — must mirror the render pre-pass
    // (src/pdf/mod.rs) so the subset embeds the glyphs the marks actually use.
    // Without this, a non-decimal format (e.g. upperRoman "I"/"II", lowerLetter
    // "b") whose letters never appear in body text gets dropped from the subset
    // and renders as a missing glyph.
    let fn_mark_fmt = doc
        .sections
        .iter()
        .find_map(|s| s.properties.footnote_num_fmt.as_deref())
        .unwrap_or("decimal");
    let en_mark_fmt = doc
        .sections
        .iter()
        .find_map(|s| s.properties.endnote_num_fmt.as_deref())
        .unwrap_or("lowerRoman");
    // Collect all page number formats across sections so inherited footers get
    // the right characters (e.g. footer in section 1 inherited by section 2 with lowerRoman).
    let all_page_num_formats: Vec<&str> = doc
        .sections
        .iter()
        .filter_map(|s| s.properties.page_num_format.as_deref())
        .collect();

    for ln in doc
        .sections
        .iter()
        .filter_map(|s| s.properties.line_numbering.as_ref())
    {
        used.entry(crate::fonts::font_key_for(&ln.font_name, false, false))
            .or_default()
            .extend('0'..='9');
    }
    for run in all_runs {
        let key = font_key_buf(run, &mut key_buf);
        let chars = used.entry(key.to_string()).or_default();
        // Word boundaries are drawn with the preceding word's font, which may
        // not be the font of the run that holds the space.
        chars.insert(' ');
        if math_font && run.math.is_some() {
            // Drawn from its box in the math font (m:nor text below).
        } else if run.caps || run.small_caps {
            chars.extend(run.text.to_uppercase().chars());
        } else {
            chars.extend(run.text.chars());
        }
        if let Some(ref fc) = run.field_code {
            match fc {
                FieldCode::Page(_) | FieldCode::NumPages(_) | FieldCode::PageRef(_) => {
                    chars.extend('0'..='9');
                    if let FieldCode::Page(Some(fmt)) | FieldCode::NumPages(Some(fmt)) = fc {
                        extend_chars_for_num_format(chars, fmt);
                    }
                    for fmt in &all_page_num_formats {
                        extend_chars_for_num_format(chars, fmt);
                    }
                }
                FieldCode::StyleRef { .. } | FieldCode::If(_) => {}
            }
        }
        if run.footnote_id.is_some() || run.is_footnote_ref_mark {
            chars.extend('0'..='9');
            extend_chars_for_num_format(chars, fn_mark_fmt);
        }
        if run.endnote_id.is_some() || run.is_endnote_ref_mark {
            chars.extend('0'..='9');
            extend_chars_for_num_format(chars, en_mark_fmt);
        }
    }
    // m:nor text, in its own font.
    let zones = all_runs.iter().filter_map(|r| r.math.as_deref());
    for r in zones.filter(|_| math_font).flat_map(|z| z.runs()) {
        if let Some(tf) = &r.normal {
            used.entry(crate::fonts::font_key_for(&tf.name, tf.bold, tf.italic))
                .or_default()
                .extend(r.text.chars());
        }
    }

    let mut all_paras: Vec<&Paragraph> = doc
        .sections
        .iter()
        .flat_map(|s| s.blocks.iter())
        .flat_map(|block| -> Vec<&Paragraph> {
            match block {
                Block::Paragraph(p) => collect_paras(p),
                Block::Table(t) => t
                    .rows
                    .iter()
                    .flat_map(|row| row.cells.iter())
                    .flat_map(|cell| cell.all_paragraphs())
                    .flat_map(|p| collect_paras(p))
                    .collect(),
            }
        })
        .collect();
    for note in doc.footnotes.values().chain(doc.endnotes.values()) {
        for p in &note.paragraphs {
            all_paras.extend(collect_paras(p));
        }
    }

    for para in &all_paras {
        if !para.list_label.is_empty() {
            // With the space `encode_label` ends the label with.
            for (key, text) in label_pieces(para) {
                if let Some(key) = key {
                    let chars = used.entry(key).or_default();
                    chars.extend(text.chars());
                    chars.insert(' ');
                }
            }
            // The label may fall back to the surrounding body font when the
            // labeled font (e.g. Symbol) can't render a PUA bullet char. Make
            // sure the body font's subset includes the Unicode substitute.
            if let Some(run) = para.runs.first() {
                let body_key = font_key_buf(run, &mut key_buf).to_string();
                let entry = used.entry(body_key).or_default();
                for c in para.list_label.chars() {
                    if let Some(sub) = super::list_label::symbol_pua_to_unicode(c) {
                        entry.insert(sub);
                    }
                }
            }
        }
        for stop in &para.tab_stops {
            if let Some(leader_char) = stop.leader
                && let Some(run) = para.runs.first()
            {
                let key = font_key_buf(run, &mut key_buf).to_string();
                used.entry(key).or_default().insert(leader_char);
            }
        }
    }

    // Collect SmartArt text chars into the default document font (for shapes
    // without per-run font names, which fall back to the global SmartArt font)
    if let Some(first_run) = all_runs.first() {
        let sa_key = font_key_buf(first_run, &mut key_buf).to_string();
        let chars = used.entry(sa_key).or_default();
        for para in &all_paras {
            for diagram in para.smartart.iter().chain(&para.floating_smartart) {
                for shape in &diagram.shapes {
                    for sa_para in &shape.paragraphs {
                        for run in &sa_para.runs {
                            if run.font_name.is_none() {
                                chars.extend(run.text.chars());
                            }
                        }
                        if let Some(ref b) = sa_para.bullet {
                            chars.extend(b.chars());
                        }
                    }
                }
            }
        }
    }

    if !doc.comments.is_empty() {
        let (body_key, label_key) = comment_pane_keys(doc);
        let body = used.entry(body_key).or_default();
        for comment in doc.comments.values() {
            body.extend(comment.text.chars());
        }
        // The labels exactly as drawn: one without initials says "[?1]".
        let label = used.entry(label_key).or_default();
        for comment in doc.comments.values() {
            label.extend(super::comments::format_label(comment).chars());
        }
    }

    // Collect characters that may appear in STYLEREF values by scanning
    // body paragraph text (paragraph styles) and run text (character styles).
    let mut styleref_chars: HashSet<char> = HashSet::new();
    for para in &all_paras {
        // `\n` values: the list number, or 0
        styleref_chars.extend(para.list_label.chars().chain(['0']));
        if para.style_id.is_some() {
            for run in &para.runs {
                styleref_chars.extend(run.text.chars());
            }
        }
        for run in &para.runs {
            if run.char_style_id.is_some() {
                styleref_chars.extend(run.text.chars());
            }
        }
    }

    for section in &doc.sections {
        for hf in [
            &section.properties.header_default,
            &section.properties.header_first,
            &section.properties.header_even,
            &section.properties.footer_default,
            &section.properties.footer_first,
            &section.properties.footer_even,
        ]
        .into_iter()
        .flatten()
        {
            for para in hf_paragraphs(hf) {
                for run in &para.runs {
                    let key = font_key_buf(run, &mut key_buf);
                    let chars = used.entry(key.to_string()).or_default();
                    if run.caps || run.small_caps {
                        chars.extend(run.text.to_uppercase().chars());
                    } else {
                        chars.extend(run.text.chars());
                    }
                    if let Some(ref fc) = run.field_code {
                        match fc {
                            FieldCode::Page(_) | FieldCode::NumPages(_) | FieldCode::PageRef(_) => {
                                chars.extend('0'..='9');
                                if let FieldCode::Page(Some(fmt)) | FieldCode::NumPages(Some(fmt)) =
                                    fc
                                {
                                    extend_chars_for_num_format(chars, fmt);
                                }
                                for fmt in &all_page_num_formats {
                                    extend_chars_for_num_format(chars, fmt);
                                }
                            }
                            FieldCode::StyleRef { .. } | FieldCode::If(_) => {
                                chars.extend(styleref_chars.iter());
                                // An IF shows its own text and nested PAGE digits too
                                if let FieldCode::If(parts) = fc {
                                    chars.extend('0'..='9');
                                    for part in parts {
                                        if let IfPart::Text(t) = part {
                                            chars.extend(t.chars());
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Collect characters for chart labels under the theme minor font
    {
        let mut chart_label_chars: HashSet<char> = HashSet::new();
        for para in &all_paras {
            if let Some(ref ic) = para.inline_chart {
                chart_label_chars.extend('0'..='9');
                chart_label_chars.insert('.');
                chart_label_chars.insert('-');
                chart_label_chars.insert(',');
                chart_label_chars.insert('%');
                for series in &ic.chart.series {
                    chart_label_chars.extend(series.label.chars());
                }
                if let Some(ref cat_axis) = ic.chart.cat_axis {
                    for label in &cat_axis.labels {
                        chart_label_chars.extend(label.chars());
                    }
                }
            }
        }
        if !chart_label_chars.is_empty() {
            used.entry(doc.theme_minor_font.clone())
                .or_default()
                .extend(chart_label_chars);
        }
    }

    // Text inside EMF pictures draws with the fonts the metafile selects
    // (`pdf::emf`), registered under the same keys as SmartArt fonts.
    let hf_paras = doc
        .sections
        .iter()
        .flat_map(|s| {
            let p = &s.properties;
            [
                &p.header_default,
                &p.header_first,
                &p.header_even,
                &p.footer_default,
                &p.footer_first,
                &p.footer_even,
            ]
        })
        .filter_map(|hf| hf.as_ref())
        .flat_map(hf_paragraphs)
        .flat_map(collect_paras);
    for para in all_paras.iter().copied().chain(hf_paras) {
        let images = para
            .image
            .iter()
            .chain(para.runs.iter().filter_map(|r| r.inline_image.as_ref()))
            .chain(para.floating_images.iter().map(|fi| &fi.image));
        for img in images.filter(|i| matches!(i.format, ImageFormat::Emf)) {
            for (face, bold, italic, text) in crate::docx::emf::text_with_fonts(&img.data) {
                used.entry(crate::fonts::font_key_for(&face, bold, italic))
                    .or_default()
                    .extend(text.chars());
            }
        }
    }

    for chars in used.values_mut() {
        chars.insert(' ');
    }

    used
}

fn extend_chars_for_num_format(chars: &mut HashSet<char>, fmt: &str) {
    match fmt {
        "lowerRoman" => chars.extend(['i', 'v', 'x', 'l', 'c', 'd', 'm']),
        "upperRoman" => chars.extend(['I', 'V', 'X', 'L', 'C', 'D', 'M']),
        "lowerLetter" => chars.extend('a'..='z'),
        "upperLetter" => chars.extend('A'..='Z'),
        _ => {}
    }
}

/// The `seen_fonts` key of the math font (`register_math_font`).
pub(super) const MATH_FONT_KEY: &str = "__math";

/// The document's math font (`m:mathFont`, else Cambria Math) when there are
/// zones and it has a MATH table: its name and file.
fn math_font_file<'d>(doc: &'d Document, all_runs: &[&Run]) -> Option<(&'d str, (Vec<u8>, u32))> {
    all_runs.iter().find(|r| r.math.is_some())?;
    let (name, (data, index)) = [doc.math.font.as_str(), "Cambria Math"]
        .into_iter()
        .find_map(|n| Some((n, crate::fonts::font_file(n, &doc.embedded_fonts)?)))?;
    if crate::math::MathFont::new(&data, index).is_none() {
        log::warn!("Math font {name} has no MATH table: equations stay linear");
        return None;
    }
    Some((name, (data, index)))
}

/// The ink of `text` in `entry`'s font, per em: (above the baseline, below
/// it).
fn text_ink(entry: &FontEntry, text: &str) -> Option<(f32, f32)> {
    let ink = entry.char_ink_1000.as_ref()?;
    let (top, bottom) = text
        .chars()
        .filter_map(|c| ink.get(&c))
        .fold((0.0f32, 0.0f32), |(t, b), &(a, d)| (t.max(a), b.max(d)));
    Some((top / 1000.0, bottom / 1000.0))
}

/// Lay out every Office Math zone with the math font, then embed that font
/// with the glyphs the boxes draw (by id: script-size alternates and
/// stretched variants have no characters). `m:nor` text measures in its own
/// font, registered before this.
fn register_math_font(
    doc: &Document,
    all_runs: &[&Run],
    (name, (data, index)): (&str, (Vec<u8>, u32)),
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
    seen_fonts: &mut HashMap<String, FontEntry>,
    font_order: &mut Vec<String>,
) {
    // A zone and its size (the run's), once each.
    let mut zones: Vec<(&std::sync::Arc<crate::model::MathZone>, f32)> = Vec::new();
    for run in all_runs {
        if let Some(z) = &run.math
            && !zones
                .iter()
                .any(|(seen, _)| std::sync::Arc::ptr_eq(seen, z))
        {
            zones.push((z, run.font_size));
        }
    }
    let Some(font) = crate::math::MathFont::new(&data, index) else {
        return;
    };
    // The text's ink, like any atom's: Word's math lines don't grow for the
    // text font's ascent and descent (probe `mp14_nor`).
    let measure = |tf: &crate::model::TextFont, text: &str, size: f32| {
        let e = seen_fonts.get(&crate::fonts::font_key_for(&tf.name, tf.bold, tf.italic))?;
        let (asc, desc) = text_ink(e, text).unwrap_or((e.ascender_ratio.unwrap_or(0.8), 0.2));
        Some(crate::math::TextBox {
            w: e.word_width(text, size, false),
            asc: asc * size,
            desc: desc * size,
            font: e.pdf_name.clone(),
        })
    };
    let mut boxes: Vec<crate::math::MathBox> = zones
        .iter()
        .map(|(z, size)| crate::math::layout(&z.root, z.display, *size, &font, &doc.math, &measure))
        .collect();
    let mut glyphs = BTreeMap::new();
    for item in boxes.iter().flat_map(|b| &b.items) {
        if let crate::math::MathItem::Glyph { gid, ch, .. } = item {
            glyphs.entry(*gid).or_insert(*ch);
        }
    }
    let pdf_name = format!("F{}", font_order.len() + 1);
    // A space marks the word boundary after a zone (`WordChunk::space_after`).
    let space = HashSet::from([' ']);
    let ctx = FontContext {
        embedded_fonts: &doc.embedded_fonts,
        font_table: &doc.font_table,
        used_chars: &space,
        word_text: false,
        extra_glyphs: Some(&glyphs),
    };
    let mut entry = register_font(pdf, name, false, false, pdf_name, alloc, &ctx);
    for item in boxes.iter_mut().flat_map(|b| &mut b.items) {
        if let crate::math::MathItem::Glyph { gid, .. } = item {
            *gid = entry.gid_remap.get(gid).copied().unwrap_or(0);
        }
    }
    entry.math_boxes = zones
        .into_iter()
        .zip(boxes)
        .map(|((z, _), b)| (std::sync::Arc::as_ptr(z).addr(), std::sync::Arc::new(b)))
        .collect();
    font_order.push(MATH_FONT_KEY.to_string());
    seen_fonts.insert(MATH_FONT_KEY.to_string(), entry);
}

pub(super) fn collect_and_register_fonts(
    doc: &Document,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> (HashMap<String, FontEntry>, Vec<String>) {
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();
    let all_runs = collect_all_runs(doc);
    let math_font = math_font_file(doc, &all_runs);
    let mut used_chars_per_font = collect_used_chars(doc, &all_runs, math_font.is_some());

    for section in &doc.sections {
        for block in &section.blocks {
            if let Block::Paragraph(para) = block {
                for diagram in para.smartart.iter().chain(&para.floating_smartart) {
                    for shape in &diagram.shapes {
                        for sa_para in &shape.paragraphs {
                            for (i, run) in sa_para.runs.iter().enumerate() {
                                if let Some(ref name) = run.font_name {
                                    let key =
                                        crate::fonts::font_key_for(name, run.bold, run.italic);
                                    let chars = used_chars_per_font.entry(key).or_default();
                                    chars.extend(run.text.chars());
                                    if i == 0
                                        && let Some(ref b) = sa_para.bullet
                                    {
                                        chars.extend(b.chars());
                                        chars.insert(' ');
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    let mut key_buf = String::new();

    for run in &all_runs {
        let key = font_key_buf(run, &mut key_buf);
        if !seen_fonts.contains_key(key) {
            let key_owned = key.to_string();
            let pdf_name = format!("F{}", font_order.len() + 1);
            let used = used_chars_per_font
                .get(&key_owned)
                .cloned()
                .unwrap_or_default();
            let ctx = FontContext {
                embedded_fonts: &doc.embedded_fonts,
                font_table: &doc.font_table,
                used_chars: &used,
                word_text: true,
                extra_glyphs: None,
            };
            let entry = register_font(
                pdf,
                &run.font_name,
                run.bold,
                run.italic,
                pdf_name,
                alloc,
                &ctx,
            );
            font_order.push(key_owned.clone());
            seen_fonts.insert(key_owned, entry);
        }
    }

    // Sorted so the font names and object order don't follow the per-process HashMap order.
    let mut leftover: Vec<_> = used_chars_per_font.iter().collect();
    leftover.sort_unstable_by(|a, b| a.0.cmp(b.0));
    for (key, used) in leftover {
        if seen_fonts.contains_key(key) {
            continue;
        }
        // SmartArt keys have style suffixes like /B, /I, /BI — extract the flags
        let (base, bold, italic) = if let Some((name, suffix)) = key.split_once('/') {
            let b = suffix.contains('B');
            let i = suffix.contains('I');
            (name, b, i)
        } else {
            (key.as_str(), false, false)
        };
        let pdf_name = format!("F{}", font_order.len() + 1);
        let ctx = FontContext {
            embedded_fonts: &doc.embedded_fonts,
            font_table: &doc.font_table,
            used_chars: used,
            word_text: false,
            extra_glyphs: None,
        };
        let entry = register_font(pdf, base, bold, italic, pdf_name, alloc, &ctx);
        seen_fonts.insert(key.clone(), entry);
        font_order.push(key.clone());
    }

    if let Some(math_font) = math_font {
        register_math_font(
            doc,
            &all_runs,
            math_font,
            pdf,
            alloc,
            &mut seen_fonts,
            &mut font_order,
        );
    }

    // Collect all CJK chars missing from their primary fonts and register a
    // shared CJK fallback font so the rendering code can substitute per-character.
    let mut all_missing_cjk: HashSet<char> = HashSet::new();
    for entry in seen_fonts.values() {
        all_missing_cjk.extend(&entry.missing_chars);
    }
    if !all_missing_cjk.is_empty() {
        let fallback_key = "__glyph_fallback".to_string();
        let pdf_name = format!("F{}", font_order.len() + 1);
        let fallback_font_name = crate::fonts::rescue_fonts(&all_missing_cjk);
        let ctx = FontContext {
            embedded_fonts: &doc.embedded_fonts,
            font_table: &doc.font_table,
            used_chars: &all_missing_cjk,
            word_text: false,
            extra_glyphs: None,
        };
        let entry = register_font(
            pdf,
            &fallback_font_name,
            false,
            false,
            pdf_name,
            alloc,
            &ctx,
        );
        font_order.push(fallback_key.clone());
        seen_fonts.insert(fallback_key, entry);

        // Copy fallback char widths into each primary font so layout uses correct widths
        if let Some(fb_widths) = seen_fonts
            .get("__glyph_fallback")
            .and_then(|e| e.char_widths_1000.clone())
        {
            for entry in seen_fonts.values_mut() {
                if entry.missing_chars.is_empty() {
                    continue;
                }
                let widths = entry.char_widths_1000.get_or_insert_with(HashMap::new);
                for &ch in &entry.missing_chars {
                    if let Some(&w) = fb_widths.get(&ch) {
                        widths.insert(ch, w);
                    }
                }
            }
        }
    }

    if seen_fonts.is_empty() {
        let pdf_name = "F1".to_string();
        let ctx = FontContext {
            embedded_fonts: &doc.embedded_fonts,
            font_table: &doc.font_table,
            used_chars: &HashSet::new(),
            word_text: false,
            extra_glyphs: None,
        };
        let entry = register_font(pdf, "Helvetica", false, false, pdf_name, alloc, &ctx);
        seen_fonts.insert("Helvetica".to_string(), entry);
        font_order.push("Helvetica".to_string());
    }

    (seen_fonts, font_order)
}
