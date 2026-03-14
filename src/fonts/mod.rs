mod cache;
mod discovery;
mod embed;
mod encoding;

use std::collections::{HashMap, HashSet};

use pdf_writer::{Name, Pdf, Ref};

use crate::model::{FontFamily, FontTable, Run};

pub(crate) use encoding::{encode_as_gids, to_winansi_bytes};

/// Metrics returned from font embedding: widths, line-height ratio, ascender ratio,
/// char-to-gid mapping, per-char widths, and kerning pairs.
pub(crate) struct FontMetrics {
    pub(crate) widths_1000: Vec<f32>,
    pub(crate) line_h_ratio: f32,
    pub(crate) ascender_ratio: f32,
    pub(crate) char_to_gid: HashMap<char, u16>,
    pub(crate) char_widths_1000: HashMap<char, f32>,
    pub(crate) kern_pairs: HashMap<(u16, u16), f32>,
    pub(crate) synthetic_bold: bool,
}

pub(crate) struct FontEntry {
    pub(crate) pdf_name: String,
    pub(crate) font_ref: Ref,
    pub(crate) widths_1000: Vec<f32>,
    pub(crate) line_h_ratio: Option<f32>,
    pub(crate) ascender_ratio: Option<f32>,
    pub(crate) char_to_gid: Option<HashMap<char, u16>>,
    pub(crate) char_widths_1000: Option<HashMap<char, f32>>,
    pub(crate) kern_pairs: Option<HashMap<(u16, u16), f32>>,
    pub(crate) synthetic_bold: bool,
}

impl FontEntry {
    /// Width of a single character in 1000-units. Uses the per-char cache (covers
    /// all Unicode chars seen in the document), falls back to the WinAnsi table.
    pub(crate) fn char_width_1000(&self, ch: char) -> f32 {
        if let Some(w) = self.char_widths_1000.as_ref().and_then(|m| m.get(&ch)) {
            return *w;
        }
        let byte = encoding::char_to_winansi(ch);
        if byte >= 32 {
            self.widths_1000[(byte - 32) as usize]
        } else {
            0.0
        }
    }

    pub(crate) fn word_width(&self, word: &str, font_size: f32, kern: bool) -> f32 {
        if !kern || self.kern_pairs.is_none() {
            return word
                .chars()
                .map(|ch| self.char_width_1000(ch) * font_size / 1000.0)
                .sum();
        }
        let scale = font_size / 1000.0;
        let mut prev: Option<char> = None;
        let mut w: f32 = 0.0;
        for ch in word.chars() {
            if let Some(p) = prev {
                w += self.kern_1000(p, ch) * scale;
            }
            w += self.char_width_1000(ch) * scale;
            prev = Some(ch);
        }
        w
    }

    fn kern_1000(&self, left: char, right: char) -> f32 {
        let (Some(pairs), Some(c2g)) = (&self.kern_pairs, &self.char_to_gid) else {
            return 0.0;
        };
        c2g.get(&left)
            .zip(c2g.get(&right))
            .and_then(|(&l, &r)| pairs.get(&(l, r)))
            .copied()
            .unwrap_or(0.0)
    }

    pub(crate) fn space_width(&self, font_size: f32) -> f32 {
        self.char_width_1000(' ') * font_size / 1000.0
    }
}

pub(crate) fn primary_font_name(name: &str) -> &str {
    name.split(';').next().unwrap_or(name).trim()
}

/// Write the font key for a run into the provided buffer, returning it as a `&str`.
/// Avoids per-call heap allocation when callers reuse the buffer.
pub(crate) fn font_key_buf<'a>(run: &Run, buf: &'a mut String) -> &'a str {
    buf.clear();
    buf.push_str(primary_font_name(&run.font_name));
    match (run.bold, run.italic) {
        (true, true) => buf.push_str("/BI"),
        (true, false) => buf.push_str("/B"),
        (false, true) => buf.push_str("/I"),
        (false, false) => {}
    }
    buf.as_str()
}

pub(crate) fn font_key(run: &Run) -> String {
    let mut buf = String::new();
    font_key_buf(run, &mut buf);
    buf
}

pub(crate) type EmbeddedFonts = HashMap<(String, bool, bool), Vec<u8>>;

fn try_font(
    pdf: &mut Pdf,
    candidate: &str,
    bold: bool,
    italic: bool,
    font_ref: Ref,
    descriptor_ref: Ref,
    data_ref: Ref,
    alloc: &mut impl FnMut() -> Ref,
    embedded_fonts: &EmbeddedFonts,
    used_chars: &HashSet<char>,
) -> Option<FontMetrics> {
    let mut embed = |data: &[u8], face_index: u32| {
        embed::embed_truetype(
            pdf,
            font_ref,
            descriptor_ref,
            data_ref,
            candidate,
            data,
            face_index,
            used_chars,
            alloc,
        )
    };

    let embedded_key = (candidate.to_lowercase(), bold, italic);
    if let Some(mut metrics) = embedded_fonts.get(&embedded_key).and_then(|d| embed(d, 0)) {
        metrics.synthetic_bold = false;
        return Some(metrics);
    }

    let (path, face_index, exact_match) = discovery::find_font_file(candidate, bold, italic)?;
    let data = std::fs::read(&path).ok()?;
    let mut metrics = embed(&data, face_index)?;
    metrics.synthetic_bold = bold && !exact_match;
    Some(metrics)
}

fn lookup_font_table<'a>(
    font_table: &'a FontTable,
    name: &str,
) -> Option<&'a crate::model::FontTableEntry> {
    font_table.get(name).or_else(|| {
        let lower = name.to_lowercase();
        font_table
            .iter()
            .find(|(k, _)| k.to_lowercase() == lower)
            .map(|(_, v)| v)
    })
}

fn family_fallback(family: FontFamily) -> Option<&'static str> {
    match family {
        FontFamily::Roman => Some("Times New Roman"),
        FontFamily::Swiss => Some("Arial"),
        FontFamily::Modern => Some("Courier New"),
        _ => None,
    }
}

pub(crate) fn register_font(
    pdf: &mut Pdf,
    font_name: &str,
    bold: bool,
    italic: bool,
    pdf_name: String,
    alloc: &mut impl FnMut() -> Ref,
    embedded_fonts: &EmbeddedFonts,
    used_chars: &HashSet<char>,
    font_table: &FontTable,
) -> FontEntry {
    let t0 = std::time::Instant::now();
    let font_ref = alloc();
    let descriptor_ref = alloc();
    let data_ref = alloc();

    let primary = primary_font_name(font_name);

    let mut try_candidate = |name: &str| {
        try_font(
            pdf,
            name,
            bold,
            italic,
            font_ref,
            descriptor_ref,
            data_ref,
            alloc,
            embedded_fonts,
            used_chars,
        )
    };

    let result = font_name
        .split(';')
        .map(|s| s.trim())
        .find_map(|c| try_candidate(c))
        .or_else(|| {
            let entry = lookup_font_table(font_table, primary)?;
            if let Some(ref alt) = entry.alt_name {
                if let Some(m) = try_candidate(alt) {
                    log::info!("Font substitution: {primary} → altName \"{alt}\"");
                    return Some(m);
                }
            }
            let fallback = family_fallback(entry.family)?;
            let m = try_candidate(fallback)?;
            log::info!(
                "Font substitution: {primary} → family {:?} fallback \"{fallback}\"",
                entry.family
            );
            Some(m)
        });

    let entry = match result {
        Some(m) => FontEntry {
            pdf_name,
            font_ref,
            widths_1000: m.widths_1000,
            line_h_ratio: Some(m.line_h_ratio),
            ascender_ratio: Some(m.ascender_ratio),
            char_to_gid: Some(m.char_to_gid),
            char_widths_1000: Some(m.char_widths_1000),
            kern_pairs: if m.kern_pairs.is_empty() {
                None
            } else {
                Some(m.kern_pairs)
            },
            synthetic_bold: m.synthetic_bold,
        },
        None => {
            log::warn!("Font not found: {font_name} bold={bold} italic={italic} — using Helvetica");
            pdf.type1_font(font_ref)
                .base_font(Name(b"Helvetica"))
                .encoding_predefined(Name(b"WinAnsiEncoding"));
            FontEntry {
                pdf_name,
                font_ref,
                widths_1000: encoding::helvetica_widths(),
                line_h_ratio: None,
                ascender_ratio: None,
                char_to_gid: None,
                char_widths_1000: None,
                kern_pairs: None,
                synthetic_bold: false,
            }
        }
    };

    log::debug!(
        "register_font: {font_name} bold={bold} italic={italic} → {:.1}ms",
        t0.elapsed().as_secs_f64() * 1000.0,
    );

    entry
}
