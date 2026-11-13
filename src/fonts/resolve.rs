//! Resolving a run's font to a face: embedded fonts, the fontTable, Word's substitution and CJK rules, and per-character rescue fonts.

use super::*;

/// Font metrics bundled with resolution metadata from the font discovery phase.
struct ResolvedFont {
    metrics: FontMetrics,
    synthetic_bold: bool,
    synthetic_italic: bool,
    font_path: Option<PathBuf>,
    face_index: u32,
}

/// What a font is resolved against: the document's embedded fonts and
/// fontTable, the characters the font has to cover, and whether it is for Word
/// text (DrawingML text with no fontTable entry keeps the last resort).
#[derive(Clone, Copy)]
pub(crate) struct FontContext<'a> {
    pub(crate) embedded_fonts: &'a EmbeddedFonts,
    pub(crate) font_table: &'a FontTable,
    pub(crate) used_chars: &'a HashSet<char>,
    pub(crate) word_text: bool,
    /// Glyphs to embed by original id, with the character each reads as
    /// (math's script-size alternates and stretched variants have none).
    pub(crate) extra_glyphs: Option<&'a BTreeMap<u16, char>>,
    /// The installed fonts, scanned for this conversion's directories.
    pub(crate) font_index: &'a FontIndex,
    pub(crate) font_files: &'a FontFiles,
}

/// The object numbers one font occupies, allocated before resolution so every
/// candidate writes to the same slots.
#[derive(Clone, Copy)]
pub(crate) struct FontRefs {
    pub(crate) font: Ref,
    pub(crate) descriptor: Ref,
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

    // An installed family wins over the document's embedded copy; the copy
    // only stands in for a font the machine lacks (online and Mac probe
    // 2026-11-09: Courier New embedded as "Times New Roman" prints in the
    // installed Times New Roman, embedded as an unknown name in Courier).
    let installed = ctx
        .font_index
        .find_font_file(candidate, bold, italic)
        .and_then(|(path, face_index, face_bold, face_italic)| {
            let data = std::fs::read(&path).ok()?;
            Some(ResolvedFont {
                metrics: embed(&data, face_index)?,
                synthetic_bold: bold && !face_bold,
                synthetic_italic: italic && !face_italic,
                font_path: Some(path),
                face_index,
            })
        });
    installed.or_else(|| {
        let metrics = ctx
            .embedded_fonts
            .get(&(candidate.to_lowercase(), bold, italic))
            .and_then(|d| embed(d, 0))?;
        Some(ResolvedFont {
            metrics,
            synthetic_bold: false,
            synthetic_italic: false,
            font_path: None,
            face_index: 0,
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

const BAHNSCHRIFT_ASCENT: f32 = 0.9965;

const BAHNSCHRIFT_LINE_H: f32 = 1.2004;

fn is_bahnschrift(name: &str) -> bool {
    name.starts_with("Bahnschrift")
}

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
        (
            hhea.ascender as f32 / units,
            -(hhea.descender as f32) / units,
        )
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
// The online export takes the face's weight and slant too: since the
// welsh_palliative reference was re-exported online, its plain "Arial Bold"
// runs draw in Arial-BoldMT (the desktop export had ArialMT).
fn family_of_face_name(name: &str) -> Option<(&str, bool, bool)> {
    [
        ("Bold Italic", true, true),
        ("Bold", true, false),
        ("Italic", false, true),
        ("Regular", false, false),
    ]
    .iter()
    .find_map(|&(style, b, i)| Some((name.strip_suffix(style)?.strip_suffix(' ')?, b, i)))
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

/// The chars the named font has glyphs for; none when it is not installed.
fn covered_chars(name: &str, chars: &HashSet<char>, font_index: &FontIndex) -> HashSet<char> {
    let Some((path, face_index, ..)) = font_index.find_font_file(name, false, false) else {
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
pub(crate) fn rescue_fonts(
    missing: &HashSet<char>,
    font_index: &FontIndex,
) -> Vec<(String, HashSet<char>)> {
    let (cjk, other): (HashSet<char>, HashSet<char>) = missing
        .iter()
        .partition(|&&ch| crate::docx::is_east_asian_char(ch));
    let mut groups: BTreeMap<String, HashSet<char>> = BTreeMap::new();
    if !cjk.is_empty() {
        groups.insert(cjk_rescue_font(&cjk, font_index), cjk);
    }
    let mut covered: HashMap<&str, HashSet<char>> = HashMap::new();
    for ch in other {
        let font = symbol_candidates(ch)
            .find(|&name| {
                covered
                    .entry(name)
                    .or_insert_with(|| covered_chars(name, missing, font_index))
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
fn cjk_rescue_font(missing: &HashSet<char>, font_index: &FontIndex) -> String {
    let script = script_of_chars(missing.iter().copied());
    // Hangul/kana gaps take the sans default (맑은 고딕 / MS Gothic); Han-only
    // gaps lead with MS Mincho, see `cjk_fallback_fonts`.
    let mut candidates: Vec<&str> =
        cjk_fallback_fonts(script, script == CjkScript::Unknown).to_vec();
    // Symbols (☑, □, ✓) are not CJK: Word draws them from the linked symbol
    // face, so give them their own candidates instead of the Han list.
    candidates.extend(SYMBOL_FALLBACK_FONTS);

    let mut best = (0usize, 0usize);
    for (i, name) in candidates.iter().enumerate() {
        let coverage = covered_chars(name, missing, font_index).len();
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
    };

    let primary = primary_font_name(font_name);

    let style = std::cell::Cell::new((bold, italic));
    let mut try_candidate = |name: &str| {
        let (bold, italic) = style.get();
        try_font(pdf, name, bold, italic, refs, alloc, ctx)
    };

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
    let bahnschrift = is_bahnschrift(primary)
        || table_entry
            .and_then(|e| e.alt_name.as_deref())
            .is_some_and(is_bahnschrift);
    let result = font_name
        .split(';')
        .map(first_css_name)
        .map(|s| word_substitute(s).unwrap_or(s))
        .filter(|c| !(has_alt && ctx.font_index.is_mac_only_family(c)))
        .filter(|c| !is_bahnschrift(c))
        .find_map(&mut try_candidate)
        .or_else(|| {
            let entry = table_entry?;
            let alt = entry.alt_name.as_ref()?;
            // "SignPainter-HouseScript": Word-for-Mac writes this cursive face as
            // altName for fonts missing on the authoring machine (e.g. Merriweather);
            // it never reflects what the reference render used.
            if alt.contains("Math") || alt == "SignPainter-HouseScript" || is_bahnschrift(alt) {
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
            let (family, b, i) = family_of_face_name(primary)?;
            style.set((bold || b, italic || i));
            let m = try_candidate(family);
            style.set((bold, italic));
            m
        })
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
            if bahnschrift || !primary.to_ascii_lowercase().starts_with("helv") {
                return None;
            }
            let m = try_candidate("Arial")?;
            log::info!("Font substitution: {primary} → Helvetica family \"Arial\"");
            Some(m)
        })
        .or_else(|| {
            // Likewise a missing one-word name starting with "Times" is Times
            // New Roman (Word for Mac probe panose: "TimesLT" and "TimesXY"
            // with any panose; "Zorblax" with Times New Roman's panose is
            // still Cambria). A Lithuanian decree in "TimesLT" wrapped a word
            // early on every line in Cambria and ran a page over. A spaced
            // variant is not: "Times New Roman CYR" is Cambria in Word
            // (a Tatar health document, 71 → 12 when it was drawn in Times).
            let lower = primary.to_ascii_lowercase();
            if !lower.starts_with("times") || lower.contains(' ') {
                return None;
            }
            let m = try_candidate("Times New Roman")?;
            log::info!("Font substitution: {primary} → Times family \"Times New Roman\"");
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
            let fallback = if bahnschrift {
                "Calibri"
            } else {
                family_fallback(family)
            };
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

    let mut entry = match result {
        Some(r) => FontEntry {
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
            avg_char_width: r.metrics.avg_char_width,
            plain_ascender_ratio: Some(r.metrics.plain_ascender_ratio),
            leading_ratio: Some(r.metrics.leading_ratio),
            drop_slack_ratio: Some(r.metrics.drop_slack_ratio),
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
            font_path: r.font_path,
            face_index: r.face_index,
            gid_remap: r.metrics.gid_remap,
            ..FontEntry::unresolved(pdf_name, refs.font, missing_chars)
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
            FontEntry::unresolved(pdf_name, refs.font, missing_chars)
        }
    };

    // Word online measures Bahnschrift lines with Bahnschrift's own metrics but
    // cannot export the variable font, so it draws Calibri glyphs on them (100pt
    // probes 2026-11-07: baseline 0.9965 em below the line top, 1.2004 em lines, for
    // "Bahnschrift" and for swiss_materials' Helvetica 55 Roman with altName
    // Bahnschrift Light; real Calibri lines are 0.954/1.221).
    if bahnschrift && entry.font_path.is_some() {
        entry.set_line_box(BAHNSCHRIFT_ASCENT, BAHNSCHRIFT_LINE_H);
    }

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
    fn word_defaults_fall_back_to_their_metric_clones() {
        assert_eq!(known_font_alias("Calibri"), Some("Carlito"));
        assert_eq!(known_font_alias("Cambria"), Some("Caladea"));
        assert_eq!(known_font_alias("Carlito"), Some("Calibri"));
    }

    #[test]
    fn a_face_used_under_two_styles_embeds_one_font_file() {
        let face = include_bytes!("testdata/nb-hyphen-native.ttf").to_vec();
        let embedded: EmbeddedFonts = [false, true]
            .map(|bold| (("synthetic hyphen".to_string(), bold, false), face.clone()))
            .into();
        let chars: HashSet<char> = "-\u{2011}".chars().collect();
        let ctx = FontContext {
            embedded_fonts: &embedded,
            font_table: &FontTable::new(),
            used_chars: &chars,
            word_text: true,
            extra_glyphs: None,
            font_index: discovery::test_index(),
            font_files: &FontFiles::default(),
        };
        let mut pdf = Pdf::new();
        let mut next = 0;
        let mut alloc = || {
            next += 1;
            Ref::new(next)
        };
        for (bold, name) in [(false, "F1"), (true, "F2")] {
            let entry = register_font(
                &mut pdf,
                "Synthetic Hyphen",
                bold,
                false,
                name.into(),
                &mut alloc,
                &ctx,
            );
            assert!(entry.font_path.is_none(), "{name} uses the embedded face");
        }
        let pdf = String::from_utf8_lossy(&pdf.finish()).into_owned();
        let font_files: Vec<&str> = pdf
            .split("/FontFile2 ")
            .skip(1)
            .map(|rest| rest.split_once(" R").expect("indirect reference").0)
            .collect();
        assert_eq!(font_files.len(), 2);
        assert_eq!(font_files[0], font_files[1]);
        let stream = format!("{} obj", font_files[0]);
        assert!(pdf.contains(&stream), "{stream} is written");
        assert_eq!(pdf.matches("/Length1").count(), 1);
        assert!(!pdf.contains(" 00000 f"), "no object number is left unused");
    }

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
            rescue_fonts(&missing.chars().collect(), discovery::test_index())
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
                font_index: discovery::test_index(),
                font_files: &FontFiles::default(),
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
        // A face's full name resolves to that face of its family, not to the
        // missing-font fallback.
        assert_eq!(
            resolve("Arial Bold", Some(("Arial Bold", Roman, None))),
            "arialbd.ttf"
        );
        assert_eq!(
            resolve(
                "Arial Bold",
                Some(("Arial Bold", Auto, Some("Courier New")))
            ),
            "courier new.ttf"
        );
    }
}
