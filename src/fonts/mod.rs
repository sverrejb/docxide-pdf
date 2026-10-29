mod cache;
mod discovery;
mod embed;
mod encoding;

use std::collections::{BTreeMap, HashMap, HashSet};
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
    /// External leading per em, above the win box (`embed::external_leading`).
    pub(crate) leading_ratio: f32,
    pub(crate) grid_baseline_shift: f32,
    /// OS/2 ySuperscriptYSize / ySubscriptYSize per em (`pdf::layout::effective_font_size`).
    pub(crate) superscript_ratio: Option<f32>,
    pub(crate) subscript_ratio: Option<f32>,
    /// post underlinePosition (below the baseline) and underlineThickness, per em.
    pub(crate) underline: Option<(f32, f32)>,
    /// OS/2 yStrikeoutPosition (above the baseline) and yStrikeoutSize, per em.
    pub(crate) strikeout: Option<(f32, f32)>,
    pub(crate) east_asian: bool,
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

/// Font metrics bundled with resolution metadata from the font discovery phase.
struct ResolvedFont {
    metrics: FontMetrics,
    synthetic_bold: bool,
    synthetic_italic: bool,
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
    pub(crate) leading_ratio: Option<f32>,
    /// See `embed::LineMetrics::grid_baseline_shift`.
    pub(crate) grid_baseline_shift: Option<f32>,
    /// OS/2 ySuperscriptYSize / ySubscriptYSize per em, sizing raised and
    /// lowered runs (`pdf::layout::effective_font_size`).
    pub(crate) superscript_ratio: Option<f32>,
    pub(crate) subscript_ratio: Option<f32>,
    /// post underlinePosition (below the baseline) and underlineThickness, per em.
    pub(crate) underline: Option<(f32, f32)>,
    /// OS/2 yStrikeoutPosition (above the baseline) and yStrikeoutSize, per em.
    pub(crate) strikeout: Option<(f32, f32)>,
    /// See `embed::LineMetrics::east_asian`.
    pub(crate) east_asian: bool,
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
            pdf_name: "F1".to_string(),
            font_ref: pdf_writer::Ref::new(1),
            widths_1000: vec![500.0; 224],
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
            plain_ascender_ratio: None,
            leading_ratio: None,
            char_to_gid: None,
            char_widths_1000: None,
            char_ink_1000: None,
            kern_pairs: None,
            synthetic_bold: false,
            synthetic_italic: false,
            is_substituted: false,
            missing_chars: Default::default(),
            drew_notdef: Default::default(),
            font_path: None,
            face_index: 0,
            gid_remap: Default::default(),
            math_boxes: Default::default(),
        }
    }
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
pub(crate) fn font_file(name: &str, embedded: &EmbeddedFonts) -> Option<(Vec<u8>, u32)> {
    if let Some(data) = embedded.get(&(name.to_lowercase(), false, false)) {
        return Some((data.clone(), 0));
    }
    let (path, index, ..) = discovery::find_font_file(name, false, false)?;
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

/// What a font is resolved against: the document's embedded fonts and
/// fontTable, the characters the font has to cover, and whether it is for Word
/// text (DrawingML text with no fontTable entry keeps the last resort).
pub(crate) struct FontContext<'a> {
    pub(crate) embedded_fonts: &'a EmbeddedFonts,
    pub(crate) font_table: &'a FontTable,
    pub(crate) used_chars: &'a HashSet<char>,
    pub(crate) word_text: bool,
    /// Glyphs to embed by original id, with the character each reads as
    /// (math's script-size alternates and stretched variants have none).
    pub(crate) extra_glyphs: Option<&'a BTreeMap<u16, char>>,
}

/// The object numbers one font occupies, allocated before resolution so every
/// candidate writes to the same slots.
#[derive(Clone, Copy)]
struct FontRefs {
    font: Ref,
    descriptor: Ref,
    data: Ref,
}

fn try_font(
    pdf: &mut Pdf,
    candidate: &str,
    bold: bool,
    italic: bool,
    refs: FontRefs,
    alloc: &mut impl FnMut() -> Ref,
    ctx: &FontContext,
) -> Option<ResolvedFont> {
    let mut embed = |data: &[u8], face_index: u32| {
        embed::embed_truetype(pdf, refs, candidate, data, face_index, ctx, alloc)
    };

    let embedded_key = (candidate.to_lowercase(), bold, italic);
    if let Some(metrics) = ctx
        .embedded_fonts
        .get(&embedded_key)
        .and_then(|d| embed(d, 0))
    {
        return Some(ResolvedFont {
            metrics,
            synthetic_bold: false,
            synthetic_italic: false,
            font_path: None,
            face_index: 0,
        });
    }

    let (path, face_index, face_bold, face_italic) =
        discovery::find_font_file(candidate, bold, italic)?;
    let data = std::fs::read(&path).ok()?;
    let metrics = embed(&data, face_index)?;
    Some(ResolvedFont {
        metrics,
        synthetic_bold: bold && !face_bold,
        synthetic_italic: italic && !face_italic,
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

/// For a font that resolves nowhere else: Arial, then its metric clones, then
/// what a bare Linux or macOS box has.
const LAST_RESORT_FONTS: &[&str] = &[
    "Arial",
    "Liberation Sans",
    "Arimo",
    "Helvetica",
    "DejaVu Sans",
];

/// Faces Windows' font linking uses for glyphs the run's font lacks.
///
/// `cjk_fallback_fonts` only lists CJK faces, and a CJK face can lack symbols
/// that Word still draws: SimSun has U+25A1 (□) but not U+2611 (☑), so a
/// checkbox run in 宋体 came out with a blank where Word shows a checked box.
/// The rescue path (`rescue_fonts`) appends these for every gap and lets glyph
/// coverage pick the leader, so pure-CJK gaps still resolve to the CJK list.
const SYMBOL_FALLBACK_FONTS: &[&str] = &[
    "Segoe UI Symbol",
    "Segoe UI Emoji",
    "Arial Unicode MS",
    "MS Gothic",
    "Noto Sans Symbols 2",
    "DejaVu Sans",
];

/// Word's face for a missing font with no usable altName: Cambria for a roman
/// family or no fontTable entry at all, Calibri for every other family. Panose,
/// pitch and the theme fonts play no part; local and online exports agree
/// (fixture fonts/missing_font_substitution).
fn family_fallback(family: Option<FontFamily>) -> &'static str {
    match family {
        Some(FontFamily::Roman) | None => "Cambria",
        Some(_) => "Calibri",
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

/// A face's hhea ascent and descent per em. Read from the table itself:
/// ttf-parser's `ascender()` switches to the typo values under USE_TYPO_METRICS.
pub(crate) fn hhea_box(path: &std::path::Path, face_index: u32) -> Option<(f32, f32)> {
    discovery::probe_face(path, face_index, |face| {
        let hhea = face.tables().hhea;
        let units = face.units_per_em() as f32;
        (hhea.ascender as f32 / units, -(hhea.descender as f32) / units)
    })
}

/// Names Word draws with its own face even where the OS has one: Mac Word sets
/// "Times" in Times New Roman (no fontTable entry, or altName Times New Roman),
/// and Windows maps Times, Courier and Helvetica to Times New Roman, Courier New
/// and Arial. Word's online export, which made most references, runs on Windows;
/// local Mac Word draws macOS Helvetica. Apple's Times, Courier and Helvetica are
/// indexed only as Mac-only faces (`discovery`).
fn word_substitute(name: &str) -> Option<&'static str> {
    match name.to_ascii_lowercase().as_str() {
        "times" => Some("Times New Roman"),
        "courier" => Some("Courier New"),
        "helvetica" => Some("Arial"),
        _ => None,
    }
}

/// The family of a missing name that is an installed face's full name: Word
/// draws "Arial Bold" (fontTable family roman, no altName) in Arial at the run's
/// own weight, ArialMT for the 37 plain runs of welsh_palliative_care_abstract_form
/// and Arial-BoldMT for its one w:b run. An altName comes first: "Bradley Hand
/// Bold" with altName Courier New is Courier New (pendulum_mechanics_oscillation_lab).
fn family_of_face_name(name: &str) -> Option<&str> {
    ["Bold Italic", "Bold", "Italic", "Regular"]
        .iter()
        .find_map(|style| name.strip_suffix(style)?.strip_suffix(' '))
}

fn known_font_alias(name: &str) -> Option<&'static str> {
    match name {
        // Word's own mapping for LibreOffice's metric clones; Liberation Mono,
        // Arimo and Tinos get none (Cambria, like any unknown name).
        "Liberation Sans" => Some("Arial"),
        "Liberation Serif" => Some("Times New Roman"),
        "Carlito" => Some("Calibri"),
        // The other way round, for machines without Office's fonts (a stock Mac or Linux box):
        // Word's default faces fall to their metric-compatible clones instead of Arial, which
        // keeps line breaks and pagination. Only reached when the real font isn't installed.
        "Calibri" => Some("Carlito"),
        "Cambria" => Some("Caladea"),
        "Palatino Linotype" => Some("Palatino"),
        "標楷體" | "DFKai-SB" => Some("BiauKai"),
        _ => None,
    }
}

fn has_cjk_chars(chars: &HashSet<char>) -> bool {
    chars.iter().any(|&c| crate::docx::is_east_asian_char(c))
}

#[derive(Copy, Clone, PartialEq, Eq, Debug)]
enum CjkScript {
    Unknown,
    SimplifiedChinese,
    TraditionalChinese,
    Japanese,
    Korean,
}

/// Korean if any Hangul, Japanese if any kana; Han alone is ambiguous → Unknown.
fn script_of_chars(chars: impl Iterator<Item = char>) -> CjkScript {
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

/// Whether a font family name is a Chinese, Japanese or Korean face (SimSun,
/// ＭＳ 明朝, Batang…); Times New Roman as a document's East Asian font is not.
pub(crate) fn is_cjk_family(name: &str) -> bool {
    classify_cjk_script(name, None, &HashSet::new()) != CjkScript::Unknown
}

/// Script of a missing CJK font: the fontTable charset first (what Word itself
/// keys substitution on), then the font name, then the text it has to render.
fn classify_cjk_script(
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
        "宋体",
        "仿宋",
        "黑体",
        "楷体",
        "华文",
        "微软雅黑",
        "方正",
        "SimSun",
        "SimHei",
        "FangSong",
        "KaiTi",
        "Microsoft YaHei",
        "STSong",
        "STFangsong",
        "STKaiti",
        "STHeiti",
        "STZhongsong",
    ];
    // Traditional-Chinese hints.
    const TC_HINTS: &[&str] = &[
        "細明體",
        "新細明體",
        "標楷體",
        "微軟正黑體",
        "華康",
        "PMingLiU",
        "MingLiU",
        "DFKai-SB",
        "Microsoft JhengHei",
    ];
    // Japanese hints (kanji/kana forms + common font names).
    const JA_HINTS: &[&str] = &[
        "明朝",
        "ゴシック",
        "メイリオ",
        "游明朝",
        "游ゴシック",
        "ＭＳ明朝",
        "ＭＳ ゴシック",
        "ＭＳ Ｐ明朝",
        "ＭＳ Ｐゴシック",
        "MS Mincho",
        "MS Gothic",
        "MS PMincho",
        "MS PGothic",
        "Meiryo",
        "Yu Mincho",
        "Yu Gothic",
        "Hiragino",
    ];
    // Korean hints.
    const KO_HINTS: &[&str] = &[
        "바탕",
        "돋움",
        "굴림",
        "궁서",
        "맑은 고딕",
        "나눔",
        "Batang",
        "Dotum",
        "Gulim",
        "Gungsuh",
        "Malgun Gothic",
        "Nanum",
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
fn cjk_fallback_fonts(script: CjkScript, serif: bool) -> &'static [&'static str] {
    use CjkScript::*;
    match (script, serif) {
        (Korean, true) => &[
            "Batang",
            "Malgun Gothic",
            "Gulim",
            "AppleMyungjo",
            "Apple SD Gothic Neo",
            "Noto Serif CJK KR",
            "Noto Sans CJK KR",
            "Arial Unicode MS",
        ],
        (Korean, false) => &[
            "Malgun Gothic",
            "Gulim",
            "Batang",
            "Apple SD Gothic Neo",
            "AppleGothic",
            "Noto Sans CJK KR",
            "Arial Unicode MS",
        ],
        (Japanese, true) => &[
            "MS Mincho",
            "Yu Mincho",
            "MS Gothic",
            "Yu Gothic",
            "Meiryo",
            "Hiragino Mincho ProN W3",
            "Hiragino Kaku Gothic ProN W3",
            "Noto Serif CJK JP",
            "Noto Sans CJK JP",
            "Arial Unicode MS",
        ],
        (Japanese, false) => &[
            "MS Gothic",
            "Yu Gothic",
            "Meiryo",
            "MS Mincho",
            "Yu Mincho",
            "Hiragino Kaku Gothic ProN W3",
            "Hiragino Sans W3",
            "Noto Sans CJK JP",
            "Arial Unicode MS",
        ],
        (SimplifiedChinese, true) => &[
            "SimSun",
            "Microsoft YaHei",
            "Songti SC",
            "PingFang SC",
            "Hiragino Sans GB W3",
            "Noto Serif CJK SC",
            "Noto Sans CJK SC",
            "Arial Unicode MS",
        ],
        (SimplifiedChinese, false) => &[
            "Microsoft YaHei",
            "SimSun",
            "PingFang SC",
            "Hiragino Sans GB W3",
            "Songti SC",
            "Noto Sans CJK SC",
            "Arial Unicode MS",
        ],
        (TraditionalChinese, true) => &[
            "PMingLiU",
            "MingLiU",
            "Microsoft JhengHei",
            "Songti TC",
            "PingFang TC",
            "Noto Serif CJK TC",
            "Noto Sans CJK TC",
            "Arial Unicode MS",
        ],
        // Word rendered the missing script-family 標楷體 in Microsoft YaHei
        // (taiwanese_education_fraud_ruling), the same face it uses for missing
        // Simplified fonts, so YaHei leads the sans list here too.
        (TraditionalChinese, false) => &[
            "Microsoft YaHei",
            "Microsoft JhengHei",
            "PMingLiU",
            "MingLiU",
            "PingFang TC",
            "Songti TC",
            "Noto Sans CJK TC",
            "Arial Unicode MS",
        ],
        // Han only. Kanji missing from a Korean face are usually Japanese
        // shinjitai (the reference rescued Batang's gaps with MS Mincho), while
        // SimSun/YaHei cover all 20 902 unified ideographs and catch the rest.
        (Unknown, true) => &[
            "MS Mincho",
            "SimSun",
            "PMingLiU",
            "Batang",
            "Songti SC",
            "Hiragino Mincho ProN W3",
            "Noto Serif CJK SC",
            "Noto Sans CJK SC",
            "Arial Unicode MS",
        ],
        (Unknown, false) => &[
            "Microsoft YaHei",
            "MS Gothic",
            "Malgun Gothic",
            "PMingLiU",
            "PingFang SC",
            "Hiragino Sans GB W3",
            "Noto Sans CJK SC",
            "Arial Unicode MS",
        ],
    }
}

/// How many of `chars` the named font has glyphs for; 0 when it is not installed.
fn glyph_coverage(name: &str, chars: &HashSet<char>) -> usize {
    covered_chars(name, chars).len()
}

/// The chars the named font has glyphs for; none when it is not installed.
fn covered_chars(name: &str, chars: &HashSet<char>) -> HashSet<char> {
    let Some((path, face_index, ..)) = discovery::find_font_file(name, false, false) else {
        return HashSet::new();
    };
    discovery::probe_face(&path, face_index, |face| {
        chars
            .iter()
            .copied()
            .filter(|&c| face.glyph_index(c).is_some())
            .collect()
    })
    .unwrap_or_default()
}

/// Fonts for the characters the resolved fonts lack, each with the characters
/// it draws. Word picks per character: kanji missing from Batang came out in
/// MS Mincho, and in one document ⇒ came from Cambria Math while ☑ came from
/// Segoe UI Symbol. East Asian gaps share the best-covering CJK face; every
/// other character takes the first face in `symbol_candidates` that has it.
pub(crate) fn rescue_fonts(missing: &HashSet<char>) -> Vec<(String, HashSet<char>)> {
    let (cjk, other): (HashSet<char>, HashSet<char>) = missing
        .iter()
        .partition(|&&ch| crate::docx::is_east_asian_char(ch));
    let mut groups: BTreeMap<String, HashSet<char>> = BTreeMap::new();
    if !cjk.is_empty() {
        groups.insert(cjk_rescue_font(&cjk), cjk);
    }
    let mut covered: HashMap<&str, HashSet<char>> = HashMap::new();
    for ch in other {
        let font = symbol_candidates(ch)
            .find(|&name| {
                covered
                    .entry(name)
                    .or_insert_with(|| covered_chars(name, missing))
                    .contains(&ch)
            })
            // Nothing has it: the symbol face draws its .notdef, as before.
            .unwrap_or(SYMBOL_FALLBACK_FONTS[0]);
        // The chosen face leads; the rest stay behind it for `register_font`.
        let chain = std::iter::once(font)
            .chain(symbol_candidates(ch).filter(|&n| n != font))
            .collect::<Vec<_>>()
            .join(";");
        groups.entry(chain).or_default().insert(ch);
    }
    groups.into_iter().collect()
}

/// Word's fallback order for a non-East-Asian character the run font lacks
/// (online export probe 2026-10-29: 151 characters under Arial, Aptos, Calibri
/// and Times New Roman). Misc Symbols and Dingbats (☑ ★ ✓ ♠) come from the
/// symbol face even when Arial has them; everything else from the first of
/// Arial, Cambria Math and the symbol face with the glyph (‒ from Arial, ⇒ ‐ ∀
/// from Cambria Math, ‱ ⁎ from Segoe UI Symbol).
fn symbol_candidates(ch: char) -> impl Iterator<Item = &'static str> {
    let lead: &[&str] = if ('\u{2600}'..='\u{27BF}').contains(&ch) {
        &[]
    } else {
        &["Arial", "Cambria Math"]
    };
    lead.iter().chain(SYMBOL_FALLBACK_FONTS).copied()
}

/// Font for East Asian characters the resolved fonts lack, shared by all of
/// them. The best-covering candidate leads and the rest follow in list order,
/// semicolon-separated so `register_font` tries each in turn.
fn cjk_rescue_font(missing: &HashSet<char>) -> String {
    let script = script_of_chars(missing.iter().copied());
    let has_cjk = missing
        .iter()
        .any(|&ch| crate::docx::is_east_asian_char(ch));
    let mut candidates: Vec<&str> = Vec::new();
    if has_cjk || script != CjkScript::Unknown {
        // Hangul/kana gaps take the sans default (맑은 고딕 / MS Gothic); Han-only
        // gaps lead with MS Mincho, see `cjk_fallback_fonts`.
        candidates.extend(cjk_fallback_fonts(script, script == CjkScript::Unknown));
    }
    // Symbols (☑, □, ✓) are not CJK: Word draws them from the linked symbol
    // face, so give them their own candidates instead of the Han list.
    candidates.extend(SYMBOL_FALLBACK_FONTS);

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
    ctx: &FontContext,
) -> FontEntry {
    let t0 = Instant::now();
    let refs = FontRefs {
        font: alloc(),
        descriptor: alloc(),
        data: alloc(),
    };

    let primary = primary_font_name(font_name);

    let mut try_candidate = |name: &str| try_font(pdf, name, bold, italic, refs, alloc, ctx);

    // Word looks the run's whole name up: "Archivo;sans-serif" has no entry
    // even when "Archivo" has one.
    let table_entry = lookup_font_table(ctx.font_table, font_name.trim());
    let script = classify_cjk_script(primary, table_entry.and_then(|e| e.charset), ctx.used_chars);
    // The declared script, not the sampled text, decides whether this is a CJK
    // slot: an empty Korean paragraph's mark font still resolves to Batang.
    let needs_cjk = script != CjkScript::Unknown || has_cjk_chars(ctx.used_chars);
    let serif = table_entry.is_some_and(|e| e.family == FontFamily::Roman);
    let substituted = std::cell::Cell::new(false);
    // List order, not glyph coverage: Word substitutes the whole run by script and
    // family and rescues single missing glyphs per character (`rescue_fonts`).
    let try_list =
        |label: &str, names: &[&str], tc: &mut dyn FnMut(&str) -> Option<ResolvedFont>| {
            names.iter().find_map(|&name| {
                log::debug!("Trying {label} \"{name}\" for \"{primary}\"");
                let m = tc(name)?;
                log::info!("Font substitution: {primary} → {label} \"{name}\"");
                substituted.set(true);
                Some(m)
            })
        };
    let cjk_fonts = cjk_fallback_fonts(script, serif);

    // The fontTable altName only stands in for a missing font: Word draws an
    // installed Calibri (altName DejaVu Sans, chinese_student_union) and Source
    // Sans Pro (altName Corbel) as requested. A macOS-only face counts as missing
    // when there is an altName, since the reference may come from Windows Word,
    // which lacks it.
    // Math fonts are excluded: an altName like "Cambria Math" (seen for "Korinna
    // BT") has enormous win ascent/descent metrics that balloon every line; Word
    // substitutes body text with a normal family fallback instead.
    let has_alt = table_entry.is_some_and(|e| e.alt_name.is_some());
    let result = font_name
        .split(';')
        .map(first_css_name)
        .map(|s| word_substitute(s).unwrap_or(s))
        .filter(|c| !(has_alt && discovery::is_mac_only_family(c)))
        .find_map(&mut try_candidate)
        .or_else(|| {
            let entry = table_entry?;
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
        .or_else(|| try_candidate(family_of_face_name(primary)?))
        .or_else(|| {
            let alias = known_font_alias(primary)?;
            let m = try_candidate(alias)?;
            log::info!("Font substitution: {primary} → alias \"{alias}\"");
            Some(m)
        })
        .or_else(|| {
            // Word draws a missing name starting with "Helv" as Arial (Helvetica 45
            // Light, Helvetica Foo, HelveticaNeue ThinCond, Helv; probes 2026-10-23;
            // swiss_materials' "Helvetica 35 Thin" with an unused Corbel Light
            // altName) unless the altName is Bahnschrift: Word online finds that
            // variable font but cannot export it and falls to the family fallback
            // (Helvetica 55 Roman -> Calibri). "Zqx Helvetica" is not a match, and
            // "Helvetica Neue" alone is Sylfaen.
            let bahnschrift = table_entry
                .and_then(|e| e.alt_name.as_deref())
                .is_some_and(|a| a.starts_with("Bahnschrift"));
            if bahnschrift || !primary.to_ascii_lowercase().starts_with("helv") {
                return None;
            }
            let m = try_candidate("Arial")?;
            log::info!("Font substitution: {primary} → Helvetica family \"Arial\"");
            Some(m)
        })
        // CJK fallback before Word's Latin defaults, which lack CJK glyphs and
        // would produce squares
        .or_else(|| {
            if !needs_cjk {
                return None;
            }
            try_list("CJK fallback", cjk_fonts, &mut try_candidate)
        })
        .or_else(|| {
            // DrawingML text (SmartArt) with no fontTable entry keeps the last
            // resort: case60's "Futura Medium" is Arial in Word's export.
            if table_entry.is_none() && !ctx.word_text {
                return None;
            }
            let family = table_entry.map(|e| e.family);
            let fallback = family_fallback(family);
            let m = try_candidate(fallback)?;
            log::info!("Font substitution: {primary} → family {family:?} fallback \"{fallback}\"");
            substituted.set(true);
            Some(m)
        })
        // A real font before the standard-14 Helvetica, which is neither
        // embedded nor ToUnicode-mapped (PDF/UA 7.21.4.1). Word gives an
        // unknown font Arial: case60's SmartArt "Futura Medium" embeds ArialMT.
        .or_else(|| try_list("last resort", LAST_RESORT_FONTS, &mut try_candidate));

    // Every char the resolved font lacks, not just CJK: Word's font linking
    // rescues any missing glyph per character, and a CJK face can lack a
    // symbol the document uses (SimSun has □ but not ☑), which used to render
    // as a blank. When no face resolved at all, only CJK is rescued: the
    // standard-14 Helvetica fallback can draw Latin but not Han.
    let missing_chars: HashSet<char> = ctx
        .used_chars
        .iter()
        .copied()
        .filter(|ch| match result.as_ref() {
            Some(r) => !r.metrics.char_to_gid.contains_key(ch) || r.metrics.stand_ins.contains(ch),
            None => crate::docx::is_east_asian_char(*ch),
        })
        .collect();

    let entry = match result {
        Some(r) => FontEntry {
            pdf_name,
            font_ref: refs.font,
            widths_1000: r.metrics.widths_1000,
            line_h_ratio: Some(r.metrics.line_h_ratio),
            ascender_ratio: Some(r.metrics.ascender_ratio),
            grid_line_ratio: r.metrics.grid_line_ratio,
            plain_line_h_ratio: Some(r.metrics.plain_line_h_ratio),
            grid_baseline_shift: Some(r.metrics.grid_baseline_shift),
            superscript_ratio: r.metrics.superscript_ratio,
            subscript_ratio: r.metrics.subscript_ratio,
            underline: r.metrics.underline,
            strikeout: r.metrics.strikeout,
            east_asian: r.metrics.east_asian,
            plain_ascender_ratio: Some(r.metrics.plain_ascender_ratio),
            leading_ratio: Some(r.metrics.leading_ratio),
            char_to_gid: Some(r.metrics.char_to_gid),
            char_widths_1000: Some(r.metrics.char_widths_1000),
            char_ink_1000: Some(r.metrics.char_ink_1000),
            kern_pairs: if r.metrics.kern_pairs.is_empty() {
                None
            } else {
                Some(r.metrics.kern_pairs)
            },
            synthetic_bold: r.synthetic_bold,
            synthetic_italic: r.synthetic_italic,
            is_substituted: substituted.get(),
            missing_chars,
            drew_notdef: Default::default(),
            font_path: r.font_path,
            face_index: r.face_index,
            gid_remap: r.metrics.gid_remap,
            math_boxes: HashMap::new(),
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
            pdf.type1_font(refs.font)
                .base_font(Name(base_font))
                .encoding_predefined(Name(b"WinAnsiEncoding"));
            FontEntry {
                pdf_name,
                font_ref: refs.font,
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
                plain_ascender_ratio: None,
                leading_ratio: None,
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
    };

    log::debug!(
        "register_font: {font_name} bold={bold} italic={italic} → {:.1}ms",
        t0.elapsed().as_secs_f64() * 1000.0,
    );

    entry
}

#[cfg(test)]
mod tests {

    #[test]
    fn word_defaults_fall_back_to_their_metric_clones() {
        assert_eq!(known_font_alias("Calibri"), Some("Carlito"));
        assert_eq!(known_font_alias("Cambria"), Some("Caladea"));
        assert_eq!(known_font_alias("Carlito"), Some("Calibri"));
    }
    use super::*;

    #[test]
    fn cjk_script_from_charset_then_name_then_text() {
        let none = HashSet::new();
        assert_eq!(
            classify_cjk_script("Whatever", Some(0x80), &none),
            CjkScript::Japanese
        );
        assert_eq!(
            classify_cjk_script("HY헤드라인M", Some(0x81), &none),
            CjkScript::Korean
        );
        assert_eq!(
            classify_cjk_script("X", Some(0x86), &none),
            CjkScript::SimplifiedChinese
        );
        assert_eq!(
            classify_cjk_script("X", Some(0x88), &none),
            CjkScript::TraditionalChinese
        );
        // Hangul in the name is a hint by itself.
        assert_eq!(
            classify_cjk_script("HY헤드라인M", None, &none),
            CjkScript::Korean
        );
        // Otherwise the text decides; Han alone stays Unknown.
        let kana: HashSet<char> = "表タイトル".chars().collect();
        assert_eq!(
            classify_cjk_script("Mystery", None, &kana),
            CjkScript::Japanese
        );
        let han: HashSet<char> = "発表".chars().collect();
        assert_eq!(
            classify_cjk_script("Mystery", None, &han),
            CjkScript::Unknown
        );
    }

    #[test]
    fn cjk_fallback_picks_word_face_by_family() {
        // Word substituted the roman-family HY헤드라인M with Batang in the reference.
        assert_eq!(cjk_fallback_fonts(CjkScript::Korean, true)[0], "Batang");
        assert_eq!(
            cjk_fallback_fonts(CjkScript::Korean, false)[0],
            "Malgun Gothic"
        );
        assert_eq!(
            cjk_fallback_fonts(CjkScript::Japanese, true)[0],
            "MS Mincho"
        );
        assert_eq!(cjk_fallback_fonts(CjkScript::Unknown, true)[0], "MS Mincho");
    }

    #[test]
    fn missing_glyphs_take_words_fallback_per_character() {
        let lead = |missing: &str, ch: char| {
            rescue_fonts(&missing.chars().collect())
                .into_iter()
                .find(|(_, chars)| chars.contains(&ch))
                .map(|(chain, _)| chain.split(';').next().unwrap().to_string())
        };
        // One document, three faces, as in Word's export.
        let missing = "\u{2010}\u{21D2}\u{2611}\u{2605}\u{2012}\u{2031}";
        assert_eq!(lead(missing, '\u{2010}').as_deref(), Some("Cambria Math"));
        assert_eq!(lead(missing, '\u{21D2}').as_deref(), Some("Cambria Math"));
        assert_eq!(
            lead(missing, '\u{2611}').as_deref(),
            Some("Segoe UI Symbol")
        );
        assert_eq!(
            lead(missing, '\u{2605}').as_deref(),
            Some("Segoe UI Symbol")
        );
        assert_eq!(lead(missing, '\u{2012}').as_deref(), Some("Arial"));
        assert_eq!(
            lead(missing, '\u{2031}').as_deref(),
            Some("Segoe UI Symbol")
        );
        // Misc Symbols skip Arial even where it has the glyph.
        assert_eq!(
            lead("\u{2660}", '\u{2660}').as_deref(),
            Some("Segoe UI Symbol")
        );
    }

    #[test]
    fn missing_fonts_resolve_as_word_does() {
        // Rows of the fixture fonts/missing_font_substitution.
        let resolve = |run: &str, entry: Option<(&str, FontFamily, Option<&str>)>| {
            let table: FontTable = entry
                .map(|(name, family, alt)| {
                    let alt_name = alt.map(Into::into);
                    let e = crate::model::FontTableEntry {
                        alt_name,
                        family,
                        charset: None,
                    };
                    (name.to_string(), e)
                })
                .into_iter()
                .collect();
            let mut next = 0;
            let mut alloc = || {
                next += 1;
                Ref::new(next)
            };
            let chars: HashSet<char> = "Ab".chars().collect();
            let ctx = FontContext {
                embedded_fonts: &EmbeddedFonts::new(),
                font_table: &table,
                used_chars: &chars,
                word_text: true,
                extra_glyphs: None,
            };
            let entry = register_font(
                &mut Pdf::new(),
                run,
                false,
                false,
                "F1".into(),
                &mut alloc,
                &ctx,
            );
            let path = entry.font_path.expect("resolved to a file");
            path.file_name().unwrap().to_string_lossy().to_lowercase()
        };
        use FontFamily::*;
        let aptos = Some("Aptos");
        // The altName only stands in for a missing font.
        assert_eq!(
            resolve("Calibri", Some(("Calibri", Swiss, aptos))),
            "calibri.ttf"
        );
        assert_eq!(
            resolve("Zqx Dalt", Some(("Zqx Dalt", Swiss, aptos))),
            "aptos.ttf"
        );
        // No usable altName: Cambria for roman or no entry, Calibri otherwise.
        assert_eq!(
            resolve("Zqx Broman", Some(("Zqx Broman", Roman, None))),
            "cambria.ttc"
        );
        assert_eq!(
            resolve("Zqx Bauto", Some(("Zqx Bauto", Auto, Some("sans-serif")))),
            "calibri.ttf"
        );
        assert_eq!(resolve("Zqx Alpha", None), "cambria.ttc");
        // A "X;Y" run has no entry even when "X" does.
        assert_eq!(
            resolve("Zqx Elist;sans-serif", Some(("Zqx Elist", Auto, None))),
            "cambria.ttc"
        );
        // Windows' and Word's own name mappings beat the altName.
        assert_eq!(
            resolve("Helvetica", Some(("Helvetica", Swiss, aptos))),
            "arial.ttf"
        );
        assert_eq!(resolve("Liberation Sans", None), "arial.ttf");
        // A face's full name resolves to its family, not to the missing-font fallback.
        assert_eq!(
            resolve("Arial Bold", Some(("Arial Bold", Roman, None))),
            "arial.ttf"
        );
        assert_eq!(
            resolve(
                "Arial Bold",
                Some(("Arial Bold", Auto, Some("Courier New")))
            ),
            "courier new.ttf"
        );
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
        assert_eq!(primary_font_name("Verdana, Geneva, sans-serif"), "Verdana");
        assert_eq!(primary_font_name("\"Calibri\", Arial"), "\"Calibri\"");
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
