mod cache;
mod discovery;
mod embed;
mod encoding;

use std::collections::{HashMap, HashSet};

use pdf_writer::{Name, Pdf, Ref};

use crate::model::{FontFamily, FontTable, Run};

pub(crate) use encoding::{encode_as_gids, to_winansi_bytes};

pub(crate) struct FontEntry {
    pub(crate) pdf_name: String,
    pub(crate) font_ref: Ref,
    pub(crate) widths_1000: Vec<f32>,
    pub(crate) line_h_ratio: Option<f32>,
    pub(crate) ascender_ratio: Option<f32>,
    pub(crate) char_to_gid: Option<HashMap<char, u16>>,
    pub(crate) char_widths_1000: Option<HashMap<char, f32>>,
    pub(crate) kern_pairs: Option<HashMap<(u16, u16), f32>>,
}

impl FontEntry {
    /// Width of a single character in 1000-units. Uses the per-char cache (covers
    /// all Unicode chars seen in the document), falls back to the WinAnsi table.
    pub(crate) fn char_width_1000(&self, ch: char) -> f32 {
        if let Some(ref map) = self.char_widths_1000 {
            if let Some(&w) = map.get(&ch) {
                return w;
            }
        }
        // Fallback: WinAnsi lookup (Helvetica or chars not in used_chars)
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
        let chars: Vec<char> = word.chars().collect();
        let mut w: f32 = 0.0;
        for (i, &ch) in chars.iter().enumerate() {
            w += self.char_width_1000(ch) * font_size / 1000.0;
            if i + 1 < chars.len() {
                w += self.kern_1000(ch, chars[i + 1]) * font_size / 1000.0;
            }
        }
        w
    }

    fn kern_1000(&self, left: char, right: char) -> f32 {
        let Some(ref pairs) = self.kern_pairs else {
            return 0.0;
        };
        let Some(ref c2g) = self.char_to_gid else {
            return 0.0;
        };
        let Some(&l) = c2g.get(&left) else {
            return 0.0;
        };
        let Some(&r) = c2g.get(&right) else {
            return 0.0;
        };
        pairs.get(&(l, r)).copied().unwrap_or(0.0)
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
) -> Option<(
    Vec<f32>,
    f32,
    f32,
    HashMap<char, u16>,
    HashMap<char, f32>,
    HashMap<(u16, u16), f32>,
)> {
    let embedded_key = (candidate.to_lowercase(), bold, italic);
    let embedded_data = embedded_fonts.get(&embedded_key);

    embedded_data
        .and_then(|data| {
            embed::embed_truetype(
                pdf,
                font_ref,
                descriptor_ref,
                data_ref,
                candidate,
                data,
                0,
                used_chars,
                alloc,
            )
        })
        .or_else(|| {
            discovery::find_font_file(candidate, bold, italic).and_then(|(path, face_index)| {
                let data = std::fs::read(&path).ok()?;
                embed::embed_truetype(
                    pdf,
                    font_ref,
                    descriptor_ref,
                    data_ref,
                    candidate,
                    &data,
                    face_index,
                    used_chars,
                    alloc,
                )
            })
        })
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

    let font_candidates: Vec<&str> = font_name.split(';').map(|s| s.trim()).collect();

    let mut result = None;
    for candidate in &font_candidates {
        let found = try_font(
            pdf,
            candidate,
            bold,
            italic,
            font_ref,
            descriptor_ref,
            data_ref,
            alloc,
            embedded_fonts,
            used_chars,
        );
        if let Some(metrics) = found {
            result = Some(metrics);
            break;
        }
    }

    // Consult fontTable.xml for substitution hints
    if result.is_none() {
        let primary = primary_font_name(font_name);
        if let Some(entry) = lookup_font_table(font_table, primary) {
            // Try altName first
            if let Some(ref alt) = entry.alt_name {
                if let Some(metrics) = try_font(
                    pdf,
                    alt,
                    bold,
                    italic,
                    font_ref,
                    descriptor_ref,
                    data_ref,
                    alloc,
                    embedded_fonts,
                    used_chars,
                ) {
                    log::info!("Font substitution: {primary} → altName \"{alt}\"");
                    result = Some(metrics);
                }
            }
            // Try family-class fallback
            if result.is_none() {
                if let Some(fallback) = family_fallback(entry.family) {
                    if let Some(metrics) = try_font(
                        pdf,
                        fallback,
                        bold,
                        italic,
                        font_ref,
                        descriptor_ref,
                        data_ref,
                        alloc,
                        embedded_fonts,
                        used_chars,
                    ) {
                        log::info!(
                            "Font substitution: {primary} → family {:?} fallback \"{fallback}\"",
                            entry.family
                        );
                        result = Some(metrics);
                    }
                }
            }
        }
    }

    let (widths, line_h_ratio, ascender_ratio, char_to_gid, char_widths_1000, kern_pairs) = result
        .map(|(w, r, ar, m, cw, kp)| {
            let kp_opt = if kp.is_empty() { None } else { Some(kp) };
            (w, Some(r), Some(ar), Some(m), Some(cw), kp_opt)
        })
        .unwrap_or_else(|| {
            log::warn!("Font not found: {font_name} bold={bold} italic={italic} — using Helvetica");
            pdf.type1_font(font_ref)
                .base_font(Name(b"Helvetica"))
                .encoding_predefined(Name(b"WinAnsiEncoding"));
            (encoding::helvetica_widths(), None, None, None, None, None)
        });

    log::debug!(
        "register_font: {font_name} bold={bold} italic={italic} → {:.1}ms",
        t0.elapsed().as_secs_f64() * 1000.0,
    );

    FontEntry {
        pdf_name,
        font_ref,
        widths_1000: widths,
        line_h_ratio,
        ascender_ratio,
        char_to_gid,
        char_widths_1000,
        kern_pairs,
    }
}
