use std::collections::{HashMap, HashSet};

use pdf_writer::{Pdf, Ref};

use crate::fonts::{FontEntry, font_key_buf, register_font};
use crate::model::{
    Block, Document, FieldCode, Paragraph, Run,
};

use super::header_footer::hf_paragraphs;
use super::{collect_paras, label_font_key, para_runs_with_textboxes};

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
        .collect()
}

fn collect_used_chars(doc: &Document, all_runs: &[&Run]) -> HashMap<String, HashSet<char>> {
    let mut used: HashMap<String, HashSet<char>> = HashMap::new();
    let mut key_buf = String::new();

    for run in all_runs {
        let key = font_key_buf(run, &mut key_buf);
        let chars = used.entry(key.to_string()).or_default();
        if run.caps || run.small_caps {
            chars.extend(run.text.to_uppercase().chars());
        } else {
            chars.extend(run.text.chars());
        }
        if let Some(ref fc) = run.field_code {
            match fc {
                FieldCode::Page | FieldCode::NumPages | FieldCode::PageRef(_) => {
                    chars.extend('0'..='9');
                }
                FieldCode::StyleRef(_) => {}
            }
        }
        if run.footnote_id.is_some() || run.is_footnote_ref_mark {
            chars.extend('0'..='9');
        }
    }

    let all_paras: Vec<&Paragraph> = doc
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

    for para in &all_paras {
        if !para.list_label.is_empty() {
            if let Some(key) = label_font_key(para) {
                used.entry(key)
                    .or_default()
                    .extend(para.list_label.chars());
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

    if let Some(first_run) = all_runs.first() {
        let sa_key = font_key_buf(first_run, &mut key_buf).to_string();
        let chars = used.entry(sa_key).or_default();
        for para in &all_paras {
            if let Some(ref diagram) = para.smartart {
                for shape in &diagram.shapes {
                    chars.extend(shape.text.chars());
                }
            }
        }
    }

    // Collect characters that may appear in STYLEREF values by scanning
    // body paragraph text (paragraph styles) and run text (character styles).
    let mut styleref_chars: HashSet<char> = HashSet::new();
    for para in &all_paras {
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

    // Collect all page number formats across sections so inherited footers get
    // the right characters (e.g. footer in section 1 inherited by section 2 with lowerRoman).
    let all_page_num_formats: Vec<&str> = doc
        .sections
        .iter()
        .filter_map(|s| s.properties.page_num_format.as_deref())
        .collect();

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
                            FieldCode::Page | FieldCode::NumPages | FieldCode::PageRef(_) => {
                                chars.extend('0'..='9');
                                for fmt in &all_page_num_formats {
                                    extend_chars_for_num_format(chars, fmt);
                                }
                            }
                            FieldCode::StyleRef(_) => {
                                chars.extend(styleref_chars.iter());
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
            used.entry(doc.chart_font_name.clone())
                .or_default()
                .extend(chart_label_chars);
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

pub(super) fn collect_and_register_fonts(
    doc: &Document,
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
) -> (HashMap<String, FontEntry>, Vec<String>) {
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();
    let all_runs = collect_all_runs(doc);
    let used_chars_per_font = collect_used_chars(doc, &all_runs);
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
            let entry = register_font(
                pdf,
                &run.font_name,
                run.bold,
                run.italic,
                pdf_name,
                alloc,
                &doc.embedded_fonts,
                &used,
                &doc.font_table,
            );
            font_order.push(key_owned.clone());
            seen_fonts.insert(key_owned, entry);
        }
    }

    for (key, used) in &used_chars_per_font {
        if !seen_fonts.contains_key(key) {
            let pdf_name = format!("F{}", font_order.len() + 1);
            let entry = register_font(
                pdf,
                key,
                false,
                false,
                pdf_name,
                alloc,
                &doc.embedded_fonts,
                used,
                &doc.font_table,
            );
            seen_fonts.insert(key.clone(), entry);
            font_order.push(key.clone());
        }
    }

    // Collect all CJK chars missing from their primary fonts and register a
    // shared CJK fallback font so the rendering code can substitute per-character.
    let mut all_missing_cjk: HashSet<char> = HashSet::new();
    for entry in seen_fonts.values() {
        all_missing_cjk.extend(&entry.missing_cjk_chars);
    }
    if !all_missing_cjk.is_empty() {
        let fallback_key = "__cjk_fallback".to_string();
        let pdf_name = format!("F{}", font_order.len() + 1);
        // Per-character fallback needs comprehensive CJK coverage (including
        // Japanese Kanji that Korean fonts like Malgun Gothic may lack).
        // Try comprehensive fonts first, then language-specific ones.
        #[cfg(target_os = "macos")]
        let fallback_font_name =
            "Hiragino Sans W3;Hiragino Kaku Gothic ProN W3;Arial Unicode MS;Malgun Gothic";
        #[cfg(target_os = "linux")]
        let fallback_font_name = "Noto Sans CJK SC;Noto Sans CJK KR;Noto Sans CJK JP";
        #[cfg(target_os = "windows")]
        let fallback_font_name = "Yu Gothic;Microsoft YaHei;Malgun Gothic";
        #[cfg(not(any(target_os = "macos", target_os = "linux", target_os = "windows")))]
        let fallback_font_name = "Arial Unicode MS";
        let entry = register_font(
            pdf,
            &fallback_font_name,
            false,
            false,
            pdf_name,
            alloc,
            &doc.embedded_fonts,
            &all_missing_cjk,
            &doc.font_table,
        );
        font_order.push(fallback_key.clone());
        seen_fonts.insert(fallback_key, entry);

        // Copy fallback char widths into each primary font so layout uses correct widths
        if let Some(fb_widths) = seen_fonts
            .get("__cjk_fallback")
            .and_then(|e| e.char_widths_1000.clone())
        {
            for entry in seen_fonts.values_mut() {
                if entry.missing_cjk_chars.is_empty() {
                    continue;
                }
                let widths = entry.char_widths_1000.get_or_insert_with(HashMap::new);
                for &ch in &entry.missing_cjk_chars {
                    if let Some(&w) = fb_widths.get(&ch) {
                        widths.insert(ch, w);
                    }
                }
            }
        }
    }

    if seen_fonts.is_empty() {
        let pdf_name = "F1".to_string();
        let entry = register_font(
            pdf,
            "Helvetica",
            false,
            false,
            pdf_name,
            alloc,
            &doc.embedded_fonts,
            &HashSet::new(),
            &doc.font_table,
        );
        seen_fonts.insert("Helvetica".to_string(), entry);
        font_order.push("Helvetica".to_string());
    }

    (seen_fonts, font_order)
}
