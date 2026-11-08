use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, font_key, font_key_for};
use crate::model::{Alignment, LabelSuffix, Paragraph};

use super::color::fill_rgb;

pub(super) fn label_font_key(para: &Paragraph) -> Option<String> {
    // Parsed labels carry the paragraph mark's font; paragraphs built by hand
    // without one take the first run's.
    Some(match &para.list_label_font {
        Some(f) => font_key_for(f, para.list_label_bold, para.list_label_italic),
        None => {
            let run = para.runs.first()?;
            font_key_for(&run.font_name, para.list_label_bold || run.bold, run.italic)
        }
    })
}

/// The label split by font: East Asian characters go in the paragraph mark's
/// East Asian font, like run text (Word draws "第 1 章" in SimSun around a
/// Times New Roman "1"). Spaces stay with the text before them.
pub(super) fn label_pieces(para: &Paragraph) -> Vec<(Option<String>, &str)> {
    let key = label_font_key(para);
    let label = para.list_label.as_str();
    let Some(ea) = para.list_label_east_asia_font.as_deref() else {
        return vec![(key, label)];
    };
    let ea_key = Some(font_key_for(
        ea,
        para.list_label_bold,
        para.list_label_italic,
    ));
    if ea_key == key {
        return vec![(key, label)];
    }
    let piece_key = |is_ea: bool| if is_ea { ea_key.clone() } else { key.clone() };
    let mut pieces = Vec::new();
    let (mut start, mut in_ea) = (0, false);
    // A label's space beside an East Asian character on either side takes
    // that font: Word sets both spaces of "第 %1 章" 6pt wide in 12pt text
    // (styleref-number-only), where the same text in a body run gets 3pt ones.
    let next_is_ea = |i: usize| {
        label[i..]
            .chars()
            .find(|c| !c.is_whitespace())
            .is_some_and(crate::docx::is_east_asian_char)
    };
    for (i, c) in label.char_indices() {
        let is_ea = if c.is_whitespace() {
            in_ea || next_is_ea(i)
        } else {
            crate::docx::is_east_asian_char(c)
        };
        if i > 0 && is_ea != in_ea {
            pieces.push((piece_key(in_ea), &label[start..i]));
            start = i;
        }
        in_ea = is_ea;
    }
    pieces.push((piece_key(in_ea), &label[start..]));
    pieces
}

/// Shows a label of several fonts in an open text object, ending with a
/// space like `encode_label`.
pub(super) fn show_label_pieces<S: AsRef<str>>(
    content: &mut Content,
    pieces: &[(Option<String>, S)],
    fonts: &HashMap<String, FontEntry>,
    size: f32,
) {
    for (i, (key, text)) in pieces.iter().enumerate() {
        let Some(entry) = key.as_deref().and_then(|k| fonts.get(k)) else {
            continue;
        };
        let mut bytes = entry.encode(text.as_ref());
        if i + 1 == pieces.len() && entry.has_char(' ') {
            bytes.extend(entry.encode(" "));
        }
        content
            .set_font(Name(entry.pdf_name.as_bytes()), size)
            .show(Str(&bytes));
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

/// The label's glyphs and a space, so text extraction and screen readers
/// don't run the label into the paragraph text ("1.01SECTION"); invisible,
/// as nothing follows it in its text object. Fonts without a space get none.
/// None when the font draws none of the label (all `.notdef`).
pub(super) fn encode_label(entry: &FontEntry, label: &str) -> Option<Vec<u8>> {
    let mut bytes = entry.encode(label);
    if is_all_notdef(&bytes) {
        return None;
    }
    if entry.has_char(' ') {
        bytes.extend(entry.encode(" "));
    }
    Some(bytes)
}

fn label_for_paragraph<'a>(
    para: &Paragraph,
    seen_fonts: &'a HashMap<String, FontEntry>,
) -> Option<(&'a FontEntry, Vec<u8>)> {
    let key = label_font_key(para);
    let entry = key.as_deref().and_then(|k| seen_fonts.get(k));

    if let Some(entry) = entry
        && let Some(bytes) = encode_label(entry, &para.list_label)
    {
        return Some((entry, bytes));
    }

    // Either the labeled font is missing, or it produced only .notdef
    // glyphs (typical for Symbol-PUA chars when the font's cmap can't
    // round-trip them). Fall back to the surrounding body-run font with
    // Symbol-PUA chars mapped to their Unicode equivalents.
    if let Some(mapped) = map_symbol_pua(&para.list_label)
        && let Some(run) = para.runs.first()
        && let Some(body_entry) = seen_fonts.get(&font_key(run))
    {
        let bytes = encode_label(body_entry, &mapped).unwrap_or_else(|| body_entry.encode(&mapped));
        return Some((body_entry, bytes));
    }

    None
}

/// `encode_as_gids` writes the .notdef glyph (gid 0) for any char missing
/// from the map. A run of pure-zero output means the font couldn't render
/// anything — treat that as a miss and fall through to substitution.
fn is_all_notdef(bytes: &[u8]) -> bool {
    !bytes.is_empty() && bytes.chunks_exact(2).all(|c| c == [0, 0])
}

/// Under `textAlignment="center"` Word centres the marker on the first
/// run's box: a 14pt Symbol bullet beside Times New Roman sits 0.75pt below
/// the text's baseline (center-body-list-height-center-*).
fn centred_label_drop(para: &Paragraph, fonts: &HashMap<String, FontEntry>, text_fs: f32) -> f32 {
    if para.text_alignment != crate::model::TextAlignment::Center {
        return 0.0;
    }
    // The glyph box, without the line gap our ascent carries.
    let centre = |e: &FontEntry, fs: f32| {
        let ar = e.ascender_ratio.unwrap_or(0.75);
        let descent = super::layout::descender_ratio(e.line_h_ratio, Some(ar));
        fs * (ar - e.leading_ratio.unwrap_or(0.0) - descent) / 2.0
    };
    let label = label_font_key(para).and_then(|k| fonts.get(&k));
    let text = para
        .runs
        .iter()
        .find(|r| !r.text.trim().is_empty())
        .and_then(|r| fonts.get(&crate::fonts::font_key(r)).map(|e| (e, r.font_size)));
    match (label, text) {
        (Some(l), Some((t, fs))) => {
            centre(l, para.list_label_font_size.unwrap_or(text_fs)) - centre(t, fs)
        }
        _ => 0.0,
    }
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
    let baseline_y =
        baseline_y + para.list_label_rise - centred_label_drop(para, fonts, fallback_font_size);
    let pieces = label_pieces(para);
    if pieces.len() > 1 && !para.bidi {
        let label_fs = para.list_label_font_size.unwrap_or(fallback_font_size);
        if let Some(c) = para.list_label_color {
            fill_rgb(content, c);
        }
        content
            .begin_text()
            .next_line(label_x + label_shift(para, fonts), baseline_y);
        show_label_pieces(content, &pieces, fonts, label_fs);
        content.end_text();
        if para.list_label_color.is_some() {
            content.set_fill_gray(0.0);
        }
        return;
    }
    let (entry, label_bytes) = label_for_paragraph(para, fonts).unzip();
    let label_bytes = label_bytes.unwrap_or_default();
    let label_font_name = entry.map_or("", |e| e.pdf_name.as_str());
    let label_color = para.list_label_color;
    if let Some(c) = label_color {
        fill_rgb(content, c);
    }
    let label_fs = para.list_label_font_size.unwrap_or(fallback_font_size);
    content
        .begin_text()
        .set_font(Name(label_font_name.as_bytes()), label_fs);
    if let Some(entry) = entry.filter(|_| para.bidi) {
        // Laid out right to left ("3." draws ".3"), read in logical order.
        content.next_line(label_x, baseline_y);
        super::layout::show_visual(content, entry, &para.list_label);
    } else {
        content
            .next_line(label_x + label_shift(para, fonts), baseline_y)
            .show(Str(&label_bytes));
    }
    content.end_text();
    let strike = |entry: &FontEntry| {
        super::layout::os2_strike(entry, label_fs).map(|(top, th)| (baseline_y + top - th, th))
    };
    let underline = |entry: &FontEntry| {
        super::layout::post_underline(entry, label_fs)
            .map(|(offset, th)| (baseline_y - offset - th, th))
    };
    if let (Some(deleted), Some(entry)) = (para.list_label_revision, entry) {
        // Struck or underlined as a run of the same font would be (Word for
        // Mac probe p11b: 0.72pt lines, 2.16 above / 1.92 below the baseline
        // at Calibri 11), over the label's width; the tab gap to the text
        // carries a 0.48pt strike after a deleted label and the full
        // underline after an inserted one.
        let x = label_x + label_shift(para, fonts);
        let w = label_width(para, fonts);
        let line = if deleted { strike(entry) } else { underline(entry) };
        if let Some((bottom, th)) = line {
            content.rect(x, bottom, w, th).fill_nonzero();
            let gap = para.indent_hanging - w;
            if gap > 0.5 {
                let gap_th = if deleted { 0.48 } else { th };
                content.rect(x + w, bottom, gap, gap_th).fill_nonzero();
            }
        }
    }
    if let (Some(split), Some(entry)) = (para.list_label_renumbered, entry) {
        // "4.2.": the original number struck, the final one underlined.
        let x = label_x + label_shift(para, fonts);
        let orig: String = para.list_label.chars().take(split).collect();
        let w_orig = label_text_width(para, fonts, label_font_key(para), &orig);
        let w_all = label_width(para, fonts);
        if let Some((bottom, th)) = strike(entry) {
            content.rect(x, bottom, w_orig, th).fill_nonzero();
        }
        if let Some((bottom, th)) = underline(entry) {
            content.rect(x + w_orig, bottom, w_all - w_orig, th).fill_nonzero();
        }
    }
    if label_color.is_some() {
        content.set_fill_gray(0.0);
    }
}

fn label_width(para: &Paragraph, fonts: &HashMap<String, FontEntry>) -> f32 {
    label_pieces(para)
        .into_iter()
        .map(|(key, text)| label_text_width(para, fonts, key, text))
        .sum()
}

/// The label's width as Word measures it when asking whether the number fits
/// its hanging indent: each glyph's advance on the 288 dpi device grid,
/// rounded up from 0.46 px (`glyph_grid`). "(iii)" in Courier New 12 is
/// 36.25pt there, not 36.006 (zimbabwe_broadcasting p5: Word tabs past a
/// 36pt indent); "(1)" in SimSun 12 is exactly 18 (probe p23: it fits an 18pt
/// one).
fn label_device_width(para: &Paragraph, fonts: &HashMap<String, FontEntry>) -> f32 {
    let size = label_font_size(para);
    label_pieces(para)
        .into_iter()
        .filter_map(|(key, text)| key.and_then(|k| fonts.get(&k)).map(|f| (f, text)))
        .flat_map(|(f, text)| {
            text.chars().map(move |ch| {
                let px = f.char_width_1000(ch) * size / 1000.0 * 4.0;
                (px.floor() + f32::from(px.fract() >= 0.46)) / 4.0
            })
        })
        .sum()
}

fn label_space_width(para: &Paragraph, fonts: &HashMap<String, FontEntry>) -> f32 {
    label_text_width(para, fonts, label_font_key(para), " ")
}

fn label_text_width(
    para: &Paragraph,
    fonts: &HashMap<String, FontEntry>,
    key: Option<String>,
    text: &str,
) -> f32 {
    let entry = key.and_then(|k| fonts.get(&k));
    entry.map_or(0.0, |f| f.word_width(text, label_font_size(para), false))
}

fn label_font_size(para: &Paragraph) -> f32 {
    para.list_label_font_size
        .unwrap_or_else(|| para.runs.first().map_or(11.0, |r| r.font_size))
}

/// Where the label starts relative to its position: `lvlJc` right ends it
/// there (indigenous_innovation's "ii." ends at the hanging indent), centre
/// straddles it.
pub(super) fn label_shift(para: &Paragraph, fonts: &HashMap<String, FontEntry>) -> f32 {
    match para.list_label_jc {
        Alignment::Right => -label_width(para, fonts),
        Alignment::Center => -label_width(para, fonts) / 2.0,
        _ => 0.0,
    }
}

/// Text starts after the label at the next available tab position. The
/// first-line indent also moves the label; tab gaps are not label width.
pub(super) fn text_hanging(
    para: &Paragraph,
    default_tab_stop: f32,
    fonts: &HashMap<String, FontEntry>,
) -> f32 {
    if para.list_label.is_empty() {
        return if para.indent_hanging > 0.0 {
            para.indent_hanging
        } else {
            -para.indent_first_line
        };
    }
    let start = para.start_indent();
    let label_x = label_offset(para) + label_shift(para, fonts);
    let end = label_x + label_width(para, fonts);
    // suff nothing/space: no tab, the text follows the label directly
    // (chinese_costume's cell items "（1）演示…").
    match para.list_label_suff {
        LabelSuffix::Nothing => return start - end,
        LabelSuffix::Space => return start - end - label_space_width(para, fonts),
        LabelSuffix::Tab => {}
    }
    // A label that exactly fills the hanging indent, by Word's device-grid
    // measure, keeps the text on it (Word for Mac probe p23: "(1)" in SimSun
    // 12 is 18pt wide under a 360/360 indent and the text starts at 360;
    // "(10)" overshoots by 6pt and goes to the next default stop). Explicit
    // stops need the pen strictly before them like any tab
    // (indigenous_innovation p5: "b." in Arial 12 ends on a 1280-twip stop
    // and Word tabs on to the 1440 indent).
    let fits_hanging = para.indent_hanging > 0.0
        && !para.no_indent_num_tab
        && start + 0.001 >= label_x + label_device_width(para, fonts);
    let next = para
        .tab_stops
        .iter()
        .map(|t| t.position)
        .chain(para.num_level_tab_stop)
        .filter(|p| *p > end + 0.01)
        .chain(fits_hanging.then_some(start))
        .min_by(f32::total_cmp)
        .unwrap_or_else(|| {
            if default_tab_stop > 0.0 {
                ((end / default_tab_stop).floor() + 1.0) * default_tab_stop
            } else {
                end
            }
        });
    start - next
}

/// How far the label sits in from the paragraph's start edge.
fn label_offset(para: &Paragraph) -> f32 {
    para.start_indent() - para.indent_hanging + para.indent_first_line
}

/// Where the label starts in a box from `box_x`, `box_w` wide. In a
/// right-to-left paragraph its right edge sits `label_offset` in from the
/// right (Word probe: "3." of a 360/360 list ends at the right margin, text
/// 18pt in).
// ponytail: lvlJc ignored for right-to-left labels
pub(super) fn label_x(
    para: &Paragraph,
    box_x: f32,
    box_w: f32,
    fonts: &HashMap<String, FontEntry>,
) -> f32 {
    if para.bidi {
        box_x + box_w - label_offset(para) - label_width(para, fonts)
    } else {
        box_x + label_offset(para)
    }
}
