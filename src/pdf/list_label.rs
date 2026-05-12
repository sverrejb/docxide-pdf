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

/// Symbol-font PUA bullet codepoints map to Unicode equivalents that any
/// general-purpose font can render. When the labeled font is missing or
/// can't encode the PUA char, fall back to these.
pub(super) fn symbol_pua_to_unicode(c: char) -> Option<char> {
    match c {
        '\u{F0B7}' => Some('\u{2022}'), // •  Symbol bullet
        '\u{F0A7}' => Some('\u{25AA}'), // ▪  Symbol filled square
        '\u{F0D8}' => Some('\u{2192}'), // →  Symbol right arrow
        '\u{F0FC}' => Some('\u{2713}'), // ✓  Symbol checkmark
        _ => None,
    }
}

fn map_symbol_pua(text: &str) -> Option<String> {
    if !text.chars().any(|c| symbol_pua_to_unicode(c).is_some()) {
        return None;
    }
    Some(
        text.chars()
            .map(|c| symbol_pua_to_unicode(c).unwrap_or(c))
            .collect(),
    )
}

fn encode_against(entry: &FontEntry, text: &str) -> Vec<u8> {
    match &entry.char_to_gid {
        Some(map) => encode_as_gids(text, map),
        None => to_winansi_bytes(text),
    }
}

pub(super) fn label_for_paragraph<'a>(
    para: &Paragraph,
    seen_fonts: &'a HashMap<String, FontEntry>,
) -> (&'a str, Vec<u8>) {
    let key = label_font_key(para);
    let entry = key.as_deref().and_then(|k| seen_fonts.get(k));

    if let Some(entry) = entry {
        let bytes = encode_against(entry, &para.list_label);
        if !is_all_notdef(&bytes) {
            return (entry.pdf_name.as_str(), bytes);
        }
    }

    // Either the labeled font is missing, or it produced only .notdef
    // glyphs (typical for Symbol-PUA chars when the font's cmap can't
    // round-trip them). Fall back to the surrounding body-run font with
    // Symbol-PUA chars mapped to their Unicode equivalents.
    if let Some(mapped) = map_symbol_pua(&para.list_label)
        && let Some(run) = para.runs.first()
        && let Some(body_entry) = seen_fonts.get(&font_key(run))
    {
        let bytes = encode_against(body_entry, &mapped);
        return (body_entry.pdf_name.as_str(), bytes);
    }

    ("", vec![])
}

/// `encode_as_gids` writes the .notdef glyph (gid 0) for any char missing
/// from the map. A run of pure-zero output means the font couldn't render
/// anything — treat that as a miss and fall through to substitution.
fn is_all_notdef(bytes: &[u8]) -> bool {
    !bytes.is_empty() && bytes.chunks_exact(2).all(|c| c == [0, 0])
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
