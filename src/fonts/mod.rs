mod cache;
mod discovery;
mod embed;
mod encoding;

use std::collections::{HashMap, HashSet};
use std::path::PathBuf;
use std::time::Instant;

use pdf_writer::{Name, Pdf, Ref};

use crate::model::{FontFamily, FontTable, Run};

pub(crate) use encoding::{encode_as_gids, to_winansi_bytes};

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
    pub(crate) char_to_gid: HashMap<char, u16>,
    pub(crate) char_widths_1000: HashMap<char, f32>,
    pub(crate) kern_pairs: HashMap<(u16, u16), f32>,
}

/// Font metrics bundled with resolution metadata from the font discovery phase.
struct ResolvedFont {
    metrics: FontMetrics,
    synthetic_bold: bool,
    font_path: Option<PathBuf>,
    face_index: u32,
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
    pub(crate) char_to_gid: Option<HashMap<char, u16>>,
    pub(crate) char_widths_1000: Option<HashMap<char, f32>>,
    pub(crate) kern_pairs: Option<HashMap<(u16, u16), f32>>,
    pub(crate) synthetic_bold: bool,
    /// True when the requested font was missing and a metric-changing fallback
    /// (CJK/family/standard-14) was used — altName/alias mappings don't count.
    pub(crate) is_substituted: bool,
    /// CJK chars requested but not present in this font (need fallback rendering).
    pub(crate) missing_cjk_chars: HashSet<char>,
    /// Font file path for glyph outline extraction (text warping).
    pub(crate) font_path: Option<PathBuf>,
    pub(crate) face_index: u32,
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

    /// Encode text for a PDF show operator: glyph IDs when this font carries a
    /// char→gid map (embedded subset), else WinAnsi bytes (standard font).
    pub(crate) fn encode(&self, text: &str) -> Vec<u8> {
        match &self.char_to_gid {
            Some(map) => encoding::encode_as_gids(text, map),
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
) -> Option<ResolvedFont> {
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
    if let Some(metrics) = embedded_fonts.get(&embedded_key).and_then(|d| embed(d, 0)) {
        return Some(ResolvedFont {
            metrics,
            synthetic_bold: false,
            font_path: None,
            face_index: 0,
        });
    }

    let (path, face_index, exact_match) = discovery::find_font_file(candidate, bold, italic)?;
    let data = std::fs::read(&path).ok()?;
    let metrics = embed(&data, face_index)?;
    Some(ResolvedFont {
        metrics,
        synthetic_bold: bold && !exact_match,
        font_path: Some(path),
        face_index,
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
        FontFamily::Script | FontFamily::Decorative => Some("Times New Roman"),
        // `w:family="auto"` (unspecified) — substitute a real vendored sans
        // instead of dropping to the base-14 Helvetica last resort, which matches
        // Word's output poorly. INTERIM choice: Arial. NOTE: for the known case
        // (Bosch Office Sans missing) Word's own PDF export actually substitutes
        // Calibri, not Arial — so Calibri may be the better universal default, or
        // the right rule may be panose/theme-based. Left as Arial pending a survey
        // of what Word substitutes across multiple missing-font fixtures.
        FontFamily::Auto => Some("Arial"),
    }
}

/// True if the face declares itself a script/handwriting design
/// (OS/2 sFamilyClass class 10, or PANOSE family kind 3 "Latin Script").
fn face_is_script_design(path: &std::path::Path, face_index: u32) -> bool {
    discovery::probe_face(path, face_index, |face| {
        // sFamilyClass high byte at offset 30, PANOSE bFamilyType at offset 32
        face.raw_face()
            .table(ttf_parser::Tag::from_bytes(b"OS/2"))
            .is_some_and(|os2| os2.get(30) == Some(&10) || os2.get(32) == Some(&3))
    })
    .unwrap_or(false)
}

fn known_font_alias(name: &str) -> Option<&'static str> {
    match name {
        "Palatino Linotype" => Some("Palatino"),
        "標楷體" | "DFKai-SB" => Some("BiauKai"),
        _ => None,
    }
}

fn has_cjk_chars(chars: &HashSet<char>) -> bool {
    chars.iter().any(|&c| crate::docx::is_east_asian_char(c))
}

#[derive(Copy, Clone, PartialEq, Eq, Debug)]
pub(crate) enum CjkScript {
    Unknown,
    SimplifiedChinese,
    TraditionalChinese,
    Japanese,
    Korean,
}

/// Korean if any Hangul, Japanese if any kana; Han alone is ambiguous → Unknown.
pub(crate) fn script_of_chars(chars: impl Iterator<Item = char>) -> CjkScript {
    let mut script = CjkScript::Unknown;
    for c in chars {
        match c as u32 {
            0x1100..=0x11FF | 0x3130..=0x318F | 0xAC00..=0xD7AF => return CjkScript::Korean,
            0x3040..=0x30FF | 0x31F0..=0x31FF => script = CjkScript::Japanese,
            _ => {}
        }
    }
    script
}

/// Script of a missing CJK font: the fontTable charset first (what Word itself
/// keys substitution on), then the font name, then the text it has to render.
pub(crate) fn classify_cjk_script(
    primary: &str,
    charset: Option<u8>,
    used_chars: &HashSet<char>,
) -> CjkScript {
    match charset {
        Some(0x80) => return CjkScript::Japanese,
        Some(0x81) | Some(0x82) => return CjkScript::Korean,
        Some(0x86) => return CjkScript::SimplifiedChinese,
        Some(0x88) => return CjkScript::TraditionalChinese,
        _ => {}
    }
    // Simplified-Chinese family names (宋体/仿宋/黑体/楷体 + 华文 variants).
    const SC_HINTS: &[&str] = &[
        "宋体", "仿宋", "黑体", "楷体", "华文", "微软雅黑", "方正",
        "SimSun", "SimHei", "FangSong", "KaiTi", "Microsoft YaHei",
        "STSong", "STFangsong", "STKaiti", "STHeiti", "STZhongsong",
    ];
    // Traditional-Chinese hints.
    const TC_HINTS: &[&str] = &[
        "細明體", "新細明體", "標楷體", "微軟正黑體", "華康",
        "PMingLiU", "MingLiU", "DFKai-SB", "Microsoft JhengHei",
    ];
    // Japanese hints (kanji/kana forms + common font names).
    const JA_HINTS: &[&str] = &[
        "明朝", "ゴシック", "メイリオ", "游明朝", "游ゴシック",
        "ＭＳ明朝", "ＭＳ ゴシック", "ＭＳ Ｐ明朝", "ＭＳ Ｐゴシック",
        "MS Mincho", "MS Gothic", "MS PMincho", "MS PGothic",
        "Meiryo", "Yu Mincho", "Yu Gothic", "Hiragino",
    ];
    // Korean hints.
    const KO_HINTS: &[&str] = &[
        "바탕", "돋움", "굴림", "궁서", "맑은 고딕", "나눔",
        "Batang", "Dotum", "Gulim", "Gungsuh", "Malgun Gothic", "Nanum",
    ];

    let by_name = [
        (SC_HINTS, CjkScript::SimplifiedChinese),
        (TC_HINTS, CjkScript::TraditionalChinese),
        (JA_HINTS, CjkScript::Japanese),
        (KO_HINTS, CjkScript::Korean),
    ]
    .into_iter()
    .find(|(hints, _)| hints.iter().any(|h| primary.contains(h)));
    if let Some((_, script)) = by_name {
        return script;
    }
    match script_of_chars(primary.chars()) {
        CjkScript::Unknown => script_of_chars(used_chars.iter().copied()),
        s => s,
    }
}

/// Substitutes for a missing CJK font, best first: one list for every platform,
/// vendored Word fonts leading so local and CI agree, Apple and Noto faces
/// trailing, and the lookup skips what is absent. `serif` (fontTable family
/// roman) picks Batang over Malgun Gothic and so on, as Word does. Evidence per
/// row: roadmap, "CJK Rendering Polish".
pub(crate) fn cjk_fallback_fonts(script: CjkScript, serif: bool) -> &'static [&'static str] {
    use CjkScript::*;
    match (script, serif) {
        (Korean, true) => &[
            "Batang", "Malgun Gothic", "Gulim", "AppleMyungjo", "Apple SD Gothic Neo",
            "Noto Serif CJK KR", "Noto Sans CJK KR", "Arial Unicode MS",
        ],
        (Korean, false) => &[
            "Malgun Gothic", "Gulim", "Batang", "Apple SD Gothic Neo", "AppleGothic",
            "Noto Sans CJK KR", "Arial Unicode MS",
        ],
        (Japanese, true) => &[
            "MS Mincho", "Yu Mincho", "MS Gothic", "Yu Gothic", "Meiryo",
            "Hiragino Mincho ProN W3", "Hiragino Kaku Gothic ProN W3",
            "Noto Serif CJK JP", "Noto Sans CJK JP", "Arial Unicode MS",
        ],
        (Japanese, false) => &[
            "MS Gothic", "Yu Gothic", "Meiryo", "MS Mincho", "Yu Mincho",
            "Hiragino Kaku Gothic ProN W3", "Hiragino Sans W3",
            "Noto Sans CJK JP", "Arial Unicode MS",
        ],
        (SimplifiedChinese, true) => &[
            "SimSun", "Microsoft YaHei", "Songti SC", "PingFang SC", "Hiragino Sans GB W3",
            "Noto Serif CJK SC", "Noto Sans CJK SC", "Arial Unicode MS",
        ],
        (SimplifiedChinese, false) => &[
            "Microsoft YaHei", "SimSun", "PingFang SC", "Hiragino Sans GB W3", "Songti SC",
            "Noto Sans CJK SC", "Arial Unicode MS",
        ],
        (TraditionalChinese, true) => &[
            "PMingLiU", "MingLiU", "Microsoft JhengHei", "Songti TC", "PingFang TC",
            "Noto Serif CJK TC", "Noto Sans CJK TC", "Arial Unicode MS",
        ],
        // Word rendered the missing script-family 標楷體 in Microsoft YaHei
        // (taiwanese_education_fraud_ruling), the same face it uses for missing
        // Simplified fonts, so YaHei leads the sans list here too.
        (TraditionalChinese, false) => &[
            "Microsoft YaHei", "Microsoft JhengHei", "PMingLiU", "MingLiU", "PingFang TC",
            "Songti TC", "Noto Sans CJK TC", "Arial Unicode MS",
        ],
        // Han only. Kanji missing from a Korean face are usually Japanese
        // shinjitai (the reference rescued Batang's gaps with MS Mincho), while
        // SimSun/YaHei cover all 20 902 unified ideographs and catch the rest.
        (Unknown, true) => &[
            "MS Mincho", "SimSun", "PMingLiU", "Batang", "Songti SC",
            "Hiragino Mincho ProN W3", "Noto Serif CJK SC", "Noto Sans CJK SC",
            "Arial Unicode MS",
        ],
        (Unknown, false) => &[
            "Microsoft YaHei", "MS Gothic", "Malgun Gothic", "PMingLiU", "PingFang SC",
            "Hiragino Sans GB W3", "Noto Sans CJK SC", "Arial Unicode MS",
        ],
    }
}

/// How many of `chars` the named font has glyphs for; 0 when it is not installed.
fn glyph_coverage(name: &str, chars: &HashSet<char>) -> usize {
    let Some((path, face_index, _)) = discovery::find_font_file(name, false, false) else {
        return 0;
    };
    discovery::probe_face(&path, face_index, |face| {
        chars.iter().filter(|&&c| face.glyph_index(c).is_some()).count()
    })
    .unwrap_or(0)
}

/// Font for characters the resolved fonts lack, shared by the whole document
/// (Word rescues per character too: kanji missing from Batang came out in
/// MS Mincho). The best-covering candidate leads and the rest follow in list
/// order, semicolon-separated so `register_font` tries each in turn.
pub(crate) fn cjk_rescue_fonts(missing: &HashSet<char>) -> String {
    let script = script_of_chars(missing.iter().copied());
    // Hangul/kana gaps take the sans default (맑은 고딕 / MS Gothic); Han-only
    // gaps lead with MS Mincho, see `cjk_fallback_fonts`.
    let candidates = cjk_fallback_fonts(script, script == CjkScript::Unknown);
    let mut best = (0usize, 0usize);
    for (i, name) in candidates.iter().enumerate() {
        let coverage = glyph_coverage(name, missing);
        if coverage > best.0 {
            best = (coverage, i);
        }
        if coverage == missing.len() {
            break;
        }
    }
    let lead = candidates[best.1];
    std::iter::once(lead)
        .chain(candidates.iter().copied().filter(|n| *n != lead))
        .collect::<Vec<_>>()
        .join(";")
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
    let t0 = Instant::now();
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

    let table_entry = lookup_font_table(font_table, primary);
    let script =
        classify_cjk_script(primary, table_entry.and_then(|e| e.charset), used_chars);
    // The declared script, not the sampled text, decides whether this is a CJK
    // slot: an empty Korean paragraph's mark font still resolves to Batang.
    let needs_cjk = script != CjkScript::Unknown || has_cjk_chars(used_chars);
    let serif = table_entry.is_some_and(|e| e.family == FontFamily::Roman);
    let substituted = std::cell::Cell::new(false);
    // List order, not glyph coverage: Word substitutes the whole run by script and
    // family and rescues single missing glyphs per character (`cjk_rescue_fonts`).
    let try_cjk_fallback = |tc: &mut dyn FnMut(&str) -> Option<ResolvedFont>| {
        cjk_fallback_fonts(script, serif).iter().find_map(|cjk_font| {
            log::debug!("Trying CJK fallback \"{cjk_font}\" for \"{primary}\"");
            let m = tc(cjk_font)?;
            log::info!("Font substitution: {primary} → CJK fallback \"{cjk_font}\"");
            substituted.set(true);
            Some(m)
        })
    };

    // If the fontTable provides an altName, try it first — it's the document's
    // explicit mapping and more reliable than the system font index (which may
    // resolve a localized name like "바탕" to a different font than "Batang").
    // Math fonts are excluded: an altName like "Cambria Math" (seen for
    // "Korinna BT") has enormous win ascent/descent metrics that balloon every
    // line; Word substitutes body text with a normal family fallback instead.
    let result = table_entry
        .and_then(|entry| {
            let alt = entry.alt_name.as_ref()?;
            // "SignPainter-HouseScript": Word-for-Mac writes this cursive face as
            // altName for fonts missing on the authoring machine (e.g. Merriweather);
            // it never reflects what the reference render used.
            if alt.contains("Math") || alt == "SignPainter-HouseScript" {
                return None;
            }
            let m = try_candidate(alt)?;
            // Reject a script/handwriting altName for a non-script family: Word on
            // macOS records whatever it substituted on screen (e.g. Merriweather →
            // SignPainter-HouseScript), but the reference machine had the real font.
            // A cursive body face is always worse than the family fallback.
            if entry.family != crate::model::FontFamily::Script
                && m.font_path
                    .as_deref()
                    .is_some_and(|p| face_is_script_design(p, m.face_index))
            {
                log::info!("Rejecting script-classified altName \"{alt}\" for {primary}");
                return None;
            }
            log::info!("Font substitution: {primary} → altName \"{alt}\"");
            Some(m)
        })
        .or_else(|| {
            font_name
                .split(';')
                .map(|s| s.trim())
                .find_map(|c| try_candidate(c))
        })
        .or_else(|| {
            let alias = known_font_alias(primary)?;
            let m = try_candidate(alias)?;
            log::info!("Font substitution: {primary} → alias \"{alias}\"");
            Some(m)
        })
        .or_else(|| {
            let entry = table_entry?;
            // Try CJK fallback before family fallback — family fonts (TNR, Courier)
            // lack CJK glyphs and would produce squares
            if needs_cjk {
                if let Some(m) = try_cjk_fallback(&mut try_candidate) {
                    return Some(m);
                }
            }
            let fallback = family_fallback(entry.family)?;
            let m = try_candidate(fallback)?;
            log::info!(
                "Font substitution: {primary} → family {:?} fallback \"{fallback}\"",
                entry.family
            );
            substituted.set(true);
            Some(m)
        })
        .or_else(|| {
            if !needs_cjk {
                return None;
            }
            try_cjk_fallback(&mut try_candidate)
        });

    // Compute which CJK chars are missing from the resolved font
    let missing_cjk = if needs_cjk {
        let covered = result.as_ref().map(|r| &r.metrics.char_to_gid);
        used_chars
            .iter()
            .copied()
            .filter(|ch| {
                crate::docx::is_east_asian_char(*ch)
                    && !covered.is_some_and(|map| map.contains_key(ch))
            })
            .collect()
    } else {
        HashSet::new()
    };

    let entry = match result {
        Some(r) => FontEntry {
            pdf_name,
            font_ref,
            widths_1000: r.metrics.widths_1000,
            line_h_ratio: Some(r.metrics.line_h_ratio),
            ascender_ratio: Some(r.metrics.ascender_ratio),
            grid_line_ratio: r.metrics.grid_line_ratio,
            plain_line_h_ratio: Some(r.metrics.plain_line_h_ratio),
            plain_ascender_ratio: Some(r.metrics.plain_ascender_ratio),
            char_to_gid: Some(r.metrics.char_to_gid),
            char_widths_1000: Some(r.metrics.char_widths_1000),
            kern_pairs: if r.metrics.kern_pairs.is_empty() {
                None
            } else {
                Some(r.metrics.kern_pairs)
            },
            synthetic_bold: r.synthetic_bold,
            is_substituted: substituted.get(),
            missing_cjk_chars: missing_cjk,
            font_path: r.font_path,
            face_index: r.face_index,
        },
        None => {
            // Pick the matching standard-14 Helvetica variant so a bold/italic run of an
            // unresolved font still renders bold/italic. Previously this always emitted plain
            // Helvetica, dropping the weight for every unresolved font across the corpus.
            let base_font: &[u8] = match (bold, italic) {
                (true, true) => b"Helvetica-BoldOblique",
                (true, false) => b"Helvetica-Bold",
                (false, true) => b"Helvetica-Oblique",
                (false, false) => b"Helvetica",
            };
            log::warn!(
                "Font not found: {font_name} bold={bold} italic={italic} — using {}",
                String::from_utf8_lossy(base_font)
            );
            pdf.type1_font(font_ref)
                .base_font(Name(base_font))
                .encoding_predefined(Name(b"WinAnsiEncoding"));
            FontEntry {
                pdf_name,
                font_ref,
                widths_1000: encoding::helvetica_widths(),
                line_h_ratio: None,
                ascender_ratio: None,
                grid_line_ratio: None,
                plain_line_h_ratio: None,
                plain_ascender_ratio: None,
                char_to_gid: None,
                char_widths_1000: None,
                kern_pairs: None,
                synthetic_bold: false,
                is_substituted: true,
                missing_cjk_chars: missing_cjk,
                font_path: None,
                face_index: 0,
            }
        }
    };

    log::debug!(
        "register_font: {font_name} bold={bold} italic={italic} → {:.1}ms",
        t0.elapsed().as_secs_f64() * 1000.0,
    );

    entry
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cjk_script_from_charset_then_name_then_text() {
        let none = HashSet::new();
        assert_eq!(classify_cjk_script("Whatever", Some(0x80), &none), CjkScript::Japanese);
        assert_eq!(classify_cjk_script("HY헤드라인M", Some(0x81), &none), CjkScript::Korean);
        assert_eq!(classify_cjk_script("X", Some(0x86), &none), CjkScript::SimplifiedChinese);
        assert_eq!(classify_cjk_script("X", Some(0x88), &none), CjkScript::TraditionalChinese);
        // Hangul in the name is a hint by itself.
        assert_eq!(classify_cjk_script("HY헤드라인M", None, &none), CjkScript::Korean);
        // Otherwise the text decides; Han alone stays Unknown.
        let kana: HashSet<char> = "表タイトル".chars().collect();
        assert_eq!(classify_cjk_script("Mystery", None, &kana), CjkScript::Japanese);
        let han: HashSet<char> = "発表".chars().collect();
        assert_eq!(classify_cjk_script("Mystery", None, &han), CjkScript::Unknown);
    }

    #[test]
    fn cjk_fallback_picks_word_face_by_family() {
        // Word substituted the roman-family HY헤드라인M with Batang in the reference.
        assert_eq!(cjk_fallback_fonts(CjkScript::Korean, true)[0], "Batang");
        assert_eq!(cjk_fallback_fonts(CjkScript::Korean, false)[0], "Malgun Gothic");
        assert_eq!(cjk_fallback_fonts(CjkScript::Japanese, true)[0], "MS Mincho");
        assert_eq!(cjk_fallback_fonts(CjkScript::Unknown, true)[0], "MS Mincho");
    }

    #[test]
    fn test_primary_font_name_simple() {
        assert_eq!(primary_font_name("Arial"), "Arial");
        assert_eq!(primary_font_name("Times New Roman"), "Times New Roman");
    }

    #[test]
    fn test_primary_font_name_with_fallback() {
        assert_eq!(primary_font_name("Arial; Helvetica"), "Arial");
        assert_eq!(primary_font_name("Calibri; sans-serif"), "Calibri");
    }

    #[test]
    fn test_primary_font_name_with_whitespace() {
        assert_eq!(primary_font_name("  Arial  ; Helvetica"), "Arial");
    }

    #[test]
    fn test_primary_font_name_empty() {
        assert_eq!(primary_font_name(""), "");
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
    fn test_font_key_regular() {
        let run = make_run("Arial", false, false);
        let mut buf = String::new();
        assert_eq!(font_key_buf(&run, &mut buf), "Arial");
    }

    #[test]
    fn test_font_key_bold() {
        let run = make_run("Arial", true, false);
        let mut buf = String::new();
        assert_eq!(font_key_buf(&run, &mut buf), "Arial/B");
    }

    #[test]
    fn test_font_key_italic() {
        let run = make_run("Arial", false, true);
        let mut buf = String::new();
        assert_eq!(font_key_buf(&run, &mut buf), "Arial/I");
    }

    #[test]
    fn test_font_key_bold_italic() {
        let run = make_run("Arial", true, true);
        let mut buf = String::new();
        assert_eq!(font_key_buf(&run, &mut buf), "Arial/BI");
    }

    #[test]
    fn test_font_key_with_fallback_font() {
        let run = make_run("Calibri; sans-serif", false, false);
        let mut buf = String::new();
        assert_eq!(font_key_buf(&run, &mut buf), "Calibri");
    }
}
