mod cache;
mod discovery;
mod embed;
mod encoding;
mod resolve;

use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::path::PathBuf;
use std::time::Instant;

use pdf_writer::{Name, Pdf, Ref};

use crate::model::{FontFamily, FontTable, Run};

pub(crate) use discovery::{FontIndex, FontSetup};
pub(crate) use encoding::{encode_as_gids, to_winansi_bytes};
pub(crate) use resolve::{
    FontContext, FontRefs, hhea_box, is_cjk_family, register_font, rescue_fonts,
};

/// Metrics extracted from a font file during embedding. Does not include font resolution
/// metadata (path, face index, synthetic bold) which is tracked separately.
pub(crate) struct FontMetrics {
    pub(crate) widths_1000: Vec<f32>,
    pub(crate) line_h_ratio: f32,
    pub(crate) ascender_ratio: f32,
    /// The height Word counts docGrid cells with (`embed::compute_line_metrics`).
    pub(crate) grid_line_ratio: Option<f32>,
    /// Latin-rule metrics without the East Asian 1.3× leading (`embed::compute_line_metrics`).
    pub(crate) plain_line_h_ratio: f32,
    pub(crate) plain_ascender_ratio: f32,
    /// External leading per em, above the win box (`embed::external_leading`).
    pub(crate) leading_ratio: f32,
    /// How far a lowered run can drop in a top-aligned line before the line
    /// grows, per em (`embed::compute_line_metrics`).
    pub(crate) drop_slack_ratio: f32,
    pub(crate) grid_baseline_shift: f32,
    /// OS/2 ySuperscriptYSize / ySubscriptYSize per em (`pdf::layout::effective_font_size`).
    pub(crate) superscript_ratio: Option<f32>,
    pub(crate) subscript_ratio: Option<f32>,
    /// post underlinePosition (below the baseline), underlineThickness and the
    /// descent the underline stays inside, per em.
    pub(crate) underline: Option<(f32, f32, f32)>,
    /// OS/2 yStrikeoutPosition (above the baseline) and yStrikeoutSize, per em.
    pub(crate) strikeout: Option<(f32, f32)>,
    pub(crate) east_asian: bool,
    /// See `FontEntry::avg_char_width`.
    pub(crate) avg_char_width: Option<(u16, u16)>,
    pub(crate) char_to_gid: HashMap<char, u16>,
    pub(crate) char_widths_1000: HashMap<char, f32>,
    /// Each used character's ink above and below the baseline, per 1000 em.
    pub(crate) char_ink_1000: HashMap<char, (f32, f32)>,
    pub(crate) kern_pairs: HashMap<(u16, u16), f32>,
    /// Subset ids of `FontContext::extra_glyphs`, by original id.
    pub(crate) gid_remap: HashMap<u16, u16>,
    /// Characters drawn with a stand-in glyph (U+2011 as the font's hyphen):
    /// a rescue font with the real glyph draws them first, as Word does.
    pub(crate) stand_ins: HashSet<char>,
}

pub(crate) struct FontEntry {
    pub(crate) pdf_name: String,
    pub(crate) font_ref: Ref,
    pub(crate) widths_1000: Vec<f32>,
    pub(crate) line_h_ratio: Option<f32>,
    pub(crate) ascender_ratio: Option<f32>,
    /// The height Word counts docGrid cells with (`embed::compute_line_metrics`).
    pub(crate) grid_line_ratio: Option<f32>,
    /// Metrics without the East Asian 1.3× leading, for whitespace-only runs and
    /// empty paragraph marks (`pdf::layout::run_line_metrics`).
    pub(crate) plain_line_h_ratio: Option<f32>,
    pub(crate) plain_ascender_ratio: Option<f32>,
    pub(crate) leading_ratio: Option<f32>,
    pub(crate) drop_slack_ratio: Option<f32>,
    /// See `embed::LineMetrics::grid_baseline_shift`.
    pub(crate) grid_baseline_shift: Option<f32>,
    /// OS/2 ySuperscriptYSize / ySubscriptYSize per em, sizing raised and
    /// lowered runs (`pdf::layout::effective_font_size`).
    pub(crate) superscript_ratio: Option<f32>,
    pub(crate) subscript_ratio: Option<f32>,
    /// post underlinePosition (below the baseline), underlineThickness and the
    /// descent the underline stays inside, per em.
    pub(crate) underline: Option<(f32, f32, f32)>,
    /// OS/2 yStrikeoutPosition (above the baseline) and yStrikeoutSize, per em.
    pub(crate) strikeout: Option<(f32, f32)>,
    /// See `embed::LineMetrics::east_asian`.
    pub(crate) east_asian: bool,
    /// Units per em and OS/2 xAvgCharWidth, what Word scales a `w:w` run's
    /// advances by (`FontEntry::scaled_word_width`).
    pub(crate) avg_char_width: Option<(u16, u16)>,
    /// The same for each character a fallback face draws, that face's.
    pub(crate) fallback_avg_char_width: HashMap<char, Option<(u16, u16)>>,
    pub(crate) char_to_gid: Option<HashMap<char, u16>>,
    pub(crate) char_widths_1000: Option<HashMap<char, f32>>,
    /// See `FontMetrics::char_ink_1000`.
    pub(crate) char_ink_1000: Option<HashMap<char, (f32, f32)>>,
    pub(crate) kern_pairs: Option<HashMap<(u16, u16), f32>>,
    pub(crate) synthetic_bold: bool,
    /// The face found has no italic: Word shears it (`pdf::layout`).
    pub(crate) synthetic_italic: bool,
    /// True when the requested font was missing and a metric-changing fallback
    /// (CJK/family/standard-14) was used — altName/alias mappings don't count.
    pub(crate) is_substituted: bool,
    /// CJK chars requested but not present in this font (need fallback rendering).
    pub(crate) missing_chars: HashSet<char>,
    /// Set when text was encoded with a character the font lacks, which draws
    /// the .notdef glyph (PDF/UA 7.21.8); the document then can't claim PDF/UA.
    pub(crate) drew_notdef: std::cell::Cell<bool>,
    /// Font file path for glyph outline extraction (text warping).
    pub(crate) font_path: Option<PathBuf>,
    pub(crate) face_index: u32,
    /// Subset ids of glyphs embedded by id (`FontContext::extra_glyphs`).
    pub(crate) gid_remap: HashMap<u16, u16>,
    /// The math font's laid-out Office Math zones, by `MathZone` address
    /// (`pdf::fonts::register_math_font`), drawing subset ids.
    pub(crate) math_boxes: HashMap<usize, std::sync::Arc<crate::math::MathBox>>,
}

#[cfg(test)]
impl FontEntry {
    /// A face with no metrics and every glyph 500 units wide, for unit tests.
    pub(crate) fn stub() -> Self {
        FontEntry {
            widths_1000: vec![500.0; 224],
            is_substituted: false,
            ..Self::unresolved("F1".into(), Ref::new(1), HashSet::new())
        }
    }
}

impl FontEntry {
    /// The entry for a font no file was found for: standard-14 Helvetica
    /// widths, no measured metrics and no per-character tables.
    fn unresolved(pdf_name: String, font_ref: Ref, missing_chars: HashSet<char>) -> Self {
        FontEntry {
            pdf_name,
            font_ref,
            widths_1000: encoding::helvetica_widths(),
            line_h_ratio: None,
            ascender_ratio: None,
            grid_line_ratio: None,
            plain_line_h_ratio: None,
            grid_baseline_shift: None,
            superscript_ratio: None,
            subscript_ratio: None,
            underline: None,
            strikeout: None,
            east_asian: false,
            avg_char_width: None,
            fallback_avg_char_width: Default::default(),
            plain_ascender_ratio: None,
            leading_ratio: None,
            drop_slack_ratio: None,
            char_to_gid: None,
            char_widths_1000: None,
            char_ink_1000: None,
            kern_pairs: None,
            synthetic_bold: false,
            synthetic_italic: false,
            is_substituted: true,
            missing_chars,
            drew_notdef: Default::default(),
            font_path: None,
            face_index: 0,
            gid_remap: HashMap::new(),
            math_boxes: HashMap::new(),
        }
    }

    /// Lay lines out with another face's box: Word measures some faces with
    /// metrics other than the file's (`register_font`, `pdf::fonts`).
    pub(crate) fn set_line_box(&mut self, ascent: f32, line_h: f32) {
        self.line_h_ratio = Some(line_h);
        self.plain_line_h_ratio = Some(line_h);
        self.ascender_ratio = Some(ascent);
        self.plain_ascender_ratio = Some(ascent);
    }

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

    /// How far the pen moves after `ch` in the PDF: an embedded font's
    /// widths are written as whole numbers, the measured width is not one.
    pub(crate) fn pen_width_1000(&self, ch: char) -> f32 {
        let width = self.char_width_1000(ch);
        if self.char_to_gid.is_some() {
            width.round()
        } else {
            width
        }
    }

    /// Encode text for a PDF show operator: glyph IDs when this font carries a
    /// char→gid map (embedded subset), else WinAnsi bytes (standard font).
    pub(crate) fn encode(&self, text: &str) -> Vec<u8> {
        match &self.char_to_gid {
            Some(map) => encoding::encode_as_gids_noting(text, map, &self.drew_notdef),
            None => encoding::to_winansi_bytes(text),
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
        let mut w = 0.0;
        for ch in word.chars() {
            if let Some(p) = prev {
                w += self.kern_1000(p, ch) * scale;
            }
            w += self.char_width_1000(ch) * scale;
            prev = Some(ch);
        }
        w
    }

    /// Width of `word` in a run Word scales horizontally (`w:w`, `text_scale`
    /// in percent). Word scales each face by its average character width
    /// rounded down to whole font units, ⌊text_scale × xAvgCharWidth / 100⌋ /
    /// xAvgCharWidth, and rounds each scaled advance to whole font units.
    /// Word examples (60-glyph words, left-aligned, 0.0025pt steps):
    /// Arial "n" at 66, 80, 85, 90, 95, 120 and 150% is 751, 911, 968, 1024,
    /// 1081, 1366 and 1709 units wide, Calibri at 85 and 90% 914 and 968,
    /// Times New Roman 869 and 920, the same at 8, 11 and 14pt and in bold
    /// and italic faces; scaling the advance alone is up to 1.15 units off.
    /// Arial Bold (xAvgCharWidth 980) at 90 and 95% scales by exactly 882 and
    /// 931, and a fallback glyph by its own face's metrics: Arial "⇒" from
    /// Cambria Math at 85% is 1507 units, not Arial's 1506 (examples).
    pub(crate) fn scaled_word_width(
        &self,
        word: &str,
        font_size: f32,
        kern: bool,
        text_scale: f32,
    ) -> f32 {
        if text_scale == 100.0 {
            return self.word_width(word, font_size, kern);
        }
        self.scaled_advances(word, font_size, kern, text_scale)
            .map(|(advance, _)| advance)
            .sum()
    }

    /// Each glyph's advance in points in a run scaled to `text_scale`
    /// percent (`scaled_word_width`) and the kerning in it, a kerned pair's
    /// adjustment on its first glyph.
    pub(crate) fn scaled_advances<'a>(
        &'a self,
        word: &'a str,
        font_size: f32,
        kern: bool,
        text_scale: f32,
    ) -> impl Iterator<Item = (f32, f32)> + 'a {
        let scale = text_scale / 100.0;
        let percent = text_scale.round() as u64;
        let mut chars = word.chars().peekable();
        std::iter::from_fn(move || {
            let ch = chars.next()?;
            let k = match chars.peek() {
                Some(&next) if kern => self.kern_1000(ch, next),
                _ => 0.0,
            };
            let kerning = k * (font_size / 1000.0);
            let metrics = if text_scale != 100.0 {
                self.fallback_avg_char_width
                    .get(&ch)
                    .copied()
                    .unwrap_or(self.avg_char_width)
            } else {
                None
            };
            let advance = match metrics {
                Some((units, avg)) => {
                    let avg = u64::from(avg);
                    let scaled_avg = percent * avg / 100;
                    let adv = (f64::from(self.char_width_1000(ch)) * f64::from(units) / 1000.0)
                        .round() as u64;
                    let scaled = (2 * adv * scaled_avg + avg) / (2 * avg);
                    scaled as f32 * font_size / f32::from(units) + kerning * scale
                }
                None => (self.char_width_1000(ch) + k) * (font_size / 1000.0) * scale,
            };
            Some((advance, kerning))
        })
    }

    pub(crate) fn kern_1000(&self, left: char, right: char) -> f32 {
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

    /// Whether `encode` draws `ch` (embedded fonts map only what they hold).
    pub(crate) fn has_char(&self, ch: char) -> bool {
        self.char_to_gid
            .as_ref()
            .is_none_or(|m| m.contains_key(&ch))
    }
}

/// The bytes and face index of a font the document embeds or the system has,
/// for reading tables the PDF side lays out with (Office Math).
pub(crate) fn font_file(
    name: &str,
    embedded: &EmbeddedFonts,
    font_index: &FontIndex,
) -> Option<(Vec<u8>, u32)> {
    if let Some(data) = embedded.get(&(name.to_lowercase(), false, false)) {
        return Some((data.clone(), 0));
    }
    let (path, index, ..) = font_index.find_font_file(name, false, false)?;
    Some((std::fs::read(path).ok()?, index))
}

pub(crate) fn primary_font_name(name: &str) -> &str {
    first_css_name(name.split(';').next().unwrap_or(name))
}

/// Web editors write CSS font lists into `w:rFonts` ("Verdana, Geneva,
/// sans-serif"); Word draws the name before the first comma, quotes and all,
/// so `"Calibri", Arial` is a missing font to it.
fn first_css_name(name: &str) -> &str {
    name.split(',').next().unwrap_or(name).trim()
}

/// Write the font key for a run into the provided buffer, returning it as a `&str`.
/// Avoids per-call heap allocation when callers reuse the buffer.
pub(crate) fn font_key_buf<'a>(run: &Run, buf: &'a mut String) -> &'a str {
    buf.clear();
    push_font_key(buf, &run.font_name, run.bold, run.italic);
    buf.as_str()
}

fn push_font_key(buf: &mut String, font_name: &str, bold: bool, italic: bool) {
    buf.push_str(primary_font_name(font_name));
    match (bold, italic) {
        (true, true) => buf.push_str("/BI"),
        (true, false) => buf.push_str("/B"),
        (false, true) => buf.push_str("/I"),
        (false, false) => {}
    }
}

pub(crate) fn font_key(run: &Run) -> String {
    font_key_for(&run.font_name, run.bold, run.italic)
}

/// The key of a face named outside any run (list labels, SmartArt, EMF).
pub(crate) fn font_key_for(font_name: &str, bold: bool, italic: bool) -> String {
    let mut buf = String::new();
    push_font_key(&mut buf, font_name, bold, italic);
    buf
}

pub(crate) type EmbeddedFonts = HashMap<(String, bool, bool), Vec<u8>>;

pub(crate) type FontFiles = RefCell<HashMap<Vec<u8>, Ref>>;

#[cfg(test)]
mod tests {

    use super::*;

    #[test]
    fn scaled_widths_match_word_probes() {
        // "n" advances (units per 2048 em) and OS/2 xAvgCharWidth of Arial,
        // Calibri, Times New Roman and Arial Bold; Word's scaled advances
        // from examples.
        for (adv, avg, scale, word_units) in [
            (1139.0, 904, 66.0, 751),
            (1139.0, 904, 80.0, 911),
            (1139.0, 904, 85.0, 968),
            (1139.0, 904, 90.0, 1024),
            (1139.0, 904, 95.0, 1081),
            (1139.0, 904, 120.0, 1366),
            (1139.0, 904, 150.0, 1709),
            (1076.0, 1067, 85.0, 914),
            (1076.0, 1067, 90.0, 968),
            (1024.0, 821, 85.0, 869),
            (1024.0, 821, 90.0, 920),
            (1251.0, 980, 90.0, 1126),
            (1251.0, 980, 95.0, 1188),
        ] {
            let entry = FontEntry {
                char_widths_1000: Some(HashMap::from([('n', adv * 1000.0 / 2048.0)])),
                avg_char_width: Some((2048, avg)),
                ..FontEntry::stub()
            };
            let w = entry.scaled_word_width("nnnn", 11.0, false, scale);
            assert!(
                (w - 4.0 * word_units as f32 * 11.0 / 2048.0).abs() < 1e-4,
                "{adv} {scale}"
            );
        }
    }

    #[test]
    fn a_fallback_glyph_scales_by_its_own_face() {
        // Arial text, "⇒" from Cambria Math (1773 units, xAvgCharWidth 1260):
        // Word draws it 1507 units wide at 85%, by Arial's 904 it would be 1506.
        let entry = FontEntry {
            char_widths_1000: Some(HashMap::from([
                ('n', 1139.0 * 1000.0 / 2048.0),
                ('⇒', 1773.0 * 1000.0 / 2048.0),
            ])),
            avg_char_width: Some((2048, 904)),
            fallback_avg_char_width: HashMap::from([('⇒', Some((2048, 1260)))]),
            ..FontEntry::stub()
        };
        let w = entry.scaled_word_width("n⇒⇒", 11.0, false, 85.0);
        assert!(
            (w - (968.0 + 2.0 * 1507.0) * 11.0 / 2048.0).abs() < 1e-4,
            "{w}"
        );
    }

    #[test]
    fn test_primary_font_name() {
        for (name, primary) in [
            ("Arial", "Arial"),
            ("Times New Roman", "Times New Roman"),
            ("Arial; Helvetica", "Arial"),
            ("Calibri; sans-serif", "Calibri"),
            ("  Arial  ; Helvetica", "Arial"),
            ("Verdana, Geneva, sans-serif", "Verdana"),
            ("\"Calibri\", Arial", "\"Calibri\""),
            ("", ""),
        ] {
            assert_eq!(primary_font_name(name), primary);
        }
    }

    fn make_run(font: &str, bold: bool, italic: bool) -> Run {
        Run {
            font_name: font.to_string(),
            font_size: 12.0,
            bold,
            italic,
            text_scale: 100.0,
            ..Run::default()
        }
    }

    #[test]
    fn test_font_key() {
        for (font, bold, italic, key) in [
            ("Arial", false, false, "Arial"),
            ("Arial", true, false, "Arial/B"),
            ("Arial", false, true, "Arial/I"),
            ("Arial", true, true, "Arial/BI"),
            ("Calibri; sans-serif", false, false, "Calibri"),
        ] {
            let run = make_run(font, bold, italic);
            let mut buf = String::new();
            assert_eq!(font_key_buf(&run, &mut buf), key);
        }
    }
}
