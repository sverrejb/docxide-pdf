use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, encode_as_gids, font_key, to_winansi_bytes};
use crate::model::{Paragraph, Run};

use super::color::fill_rgb;

pub(super) fn label_font_key(para: &Paragraph) -> Option<String> {
    if let Some(ref bf) = para.list_label_font {
        let mut k = bf.clone();
        if para.list_label_bold {
            k.push_str("/B");
        }
        Some(k)
    } else {
        let run = para.runs.first()?;
        let key_run = Run {
            bold: para.list_label_bold || run.bold,
            ..run.clone()
        };
        Some(font_key(&key_run))
    }
}

pub(super) fn label_for_paragraph<'a>(
    para: &Paragraph,
    seen_fonts: &'a HashMap<String, FontEntry>,
) -> (&'a str, Vec<u8>) {
    let Some(key) = label_font_key(para) else {
        return ("", vec![]);
    };
    let Some(entry) = seen_fonts.get(&key) else {
        return ("", vec![]);
    };
    let bytes = match &entry.char_to_gid {
        Some(map) => encode_as_gids(&para.list_label, map),
        None => to_winansi_bytes(&para.list_label),
    };
    (entry.pdf_name.as_str(), bytes)
}

pub(super) fn render_list_label(
    content: &mut Content,
    para: &Paragraph,
    fonts: &HashMap<String, FontEntry>,
    label_x: f32,
    baseline_y: f32,
    fallback_font_size: f32,
) {
    if para.list_label.is_empty() {
        return;
    }
    let (label_font_name, label_bytes) = label_for_paragraph(para, fonts);
    let label_color = para
        .list_label_color
        .or_else(|| para.runs.first().and_then(|r| r.color));
    if let Some(c) = label_color {
        fill_rgb(content, c);
    }
    let label_fs = para.list_label_font_size.unwrap_or(fallback_font_size);
    content
        .begin_text()
        .set_font(Name(label_font_name.as_bytes()), label_fs)
        .next_line(label_x, baseline_y)
        .show(Str(&label_bytes))
        .end_text();
    if label_color.is_some() {
        content.set_fill_gray(0.0);
    }
}

pub(super) fn para_runs_with_textboxes(para: &Paragraph) -> Vec<&Run> {
    let mut out: Vec<&Run> = para.runs.iter().collect();
    for tb in &para.textboxes {
        for tp in &tb.paragraphs {
            out.extend(para_runs_with_textboxes(tp));
        }
    }
    out
}

pub(super) fn collect_paras(para: &Paragraph) -> Vec<&Paragraph> {
    let mut out = vec![para];
    for tb in &para.textboxes {
        for tp in &tb.paragraphs {
            out.extend(collect_paras(tp));
        }
    }
    out
}
