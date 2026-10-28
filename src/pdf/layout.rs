use std::borrow::Cow;
use std::collections::HashMap;
use std::sync::{Arc, LazyLock};

use pdf_writer::types::TextRenderingMode;
use pdf_writer::{Content, Name, Rect, Str};

use crate::fonts::{FontEntry, encode_as_gids, font_key_buf, to_winansi_bytes};
use crate::model::{
    Alignment, FormCheckbox, LineSpacing, Paragraph, ParagraphBorder, Run, TabAlignment, TabStop,
    TextFill, TextOutline, TextShadow, VertAlign,
};

use super::RenderContext;
use super::color::{fill_color_or_black, stroke_color_or_black};
use super::images::EffectXObjs;

/// Placeholders for a paragraph without inline pictures.
pub(super) static EMPTY_INLINE_IMAGES: LazyLock<HashMap<usize, String>> =
    LazyLock::new(HashMap::new);
pub(super) static EMPTY_EFFECTS: LazyLock<HashMap<usize, EffectXObjs>> =
    LazyLock::new(HashMap::new);

/// Run index → registered value (XObject name, effects) for each run holding
/// an inline picture that `names` knows.
pub(super) fn run_image_map<V: Clone>(
    runs: &[Run],
    names: &HashMap<usize, V>,
) -> HashMap<usize, V> {
    runs.iter()
        .enumerate()
        .filter_map(|(ri, run)| Some((ri, names.get(&run.inline_image.as_ref()?.key())?.clone())))
        .collect()
}

/// How many gaps a stretched line's slack spreads across when the slack goes
/// between *characters* (applied as PDF `Tc`) rather than between word gaps.
/// `None` means the line falls back to word-gap justification.
///
/// Two callers want char-level spreading: CJK (Word treats inter-character gaps
/// like word gaps) and `w:jc="distribute"` ("Distribute All Characters Equally",
/// §17.18.44). They differ at the right edge — CJK justify keeps the trailing
/// cell gap of the grid, so its slack divides by the character count, while
/// `distribute` ends flush at both margins (Japanese 均等割り付け behaves the
/// same way), so it divides by one gap fewer.
///
/// ponytail: the count excludes inter-word spaces, so a distributed Latin line
/// with spaces spreads across one gap too few — case77 sample 3 sits up to 34pt
/// off Word's interior letter positions, though both margins still land flush.
/// Word counts spaces as distributable characters. Fixing it means carrying a
/// per-chunk space count out of the line-breaking loop (`pending_space_w` keeps
/// only the width, across 5 push sites and ~8 reset points), which is a poor
/// trade while no corpus fixture distributes text containing spaces.
fn char_justify_gaps(
    alignment: Alignment,
    can_justify: bool,
    has_cjk: bool,
    char_count: usize,
) -> Option<usize> {
    if !can_justify || char_count < 2 {
        return None;
    }
    match (alignment, has_cjk) {
        (Alignment::Distribute, _) => Some(char_count - 1),
        (_, true) => Some(char_count),
        _ => None,
    }
}

/// Word's justification of a line mixing Latin and East Asian text: the slack
/// spreads over every space and the gap after every East Asian character
/// alike, each space getting a fixed 1/9 em on top, and Latin letters never
/// spread (Word online probes 2026-10-28, issue #142: a 10pt line with 6
/// spaces and 17 CJK gaps widens the spaces by 5.17pt and the gaps by
/// 4.04pt; 12pt lines of 2–11 spaces and 2–24 gaps fit the same model within
/// Word's quarter-point rounding). The gap before an East Asian character
/// counts only after a Latin letter, and the line's last character has none.
struct MixedJustify {
    /// Spaces and East Asian gaps before each chunk.
    spaces_before: Vec<usize>,
    points_before: Vec<usize>,
    /// The last chunk holding text: its final character gets no gap.
    last_text_chunk: usize,
    extra: f32,
    space_bonus: f32,
}

impl MixedJustify {
    fn new(chunks: &[WordChunk], spaces_before: &[usize], slack: f32) -> Self {
        let last_text_chunk = chunks
            .iter()
            .rposition(|c| c.text.chars().any(|ch| !is_break_space(ch)))
            .unwrap_or(0);
        let mut points_before = Vec::with_capacity(chunks.len());
        let mut points = 0usize;
        for (i, c) in chunks.iter().enumerate() {
            points_before.push(points);
            points += cjk_gap_points(&c.text, i == last_text_chunk)
                .iter()
                .filter(|p| **p)
                .count();
        }
        let spaces = spaces_before.last().copied().unwrap_or(0) as f32;
        let total = spaces + points as f32;
        let bonus = chunks.first().map_or(0.0, |c| c.font_size / 9.0);
        let (extra, space_bonus) = if total == 0.0 {
            (0.0, 0.0)
        } else if slack >= spaces * bonus {
            ((slack - spaces * bonus) / total, bonus)
        } else {
            (0.0, slack / spaces.max(1.0))
        };
        Self {
            spaces_before: spaces_before.to_vec(),
            points_before,
            last_text_chunk,
            extra,
            space_bonus,
        }
    }

    /// How far chunk `idx` moves right.
    fn offset_before(&self, idx: usize) -> f32 {
        let at = |v: &[usize]| v.get(idx).copied().unwrap_or(0) as f32;
        at(&self.spaces_before) * (self.extra + self.space_bonus)
            + at(&self.points_before) * self.extra
    }

    /// Extra advance after each character of chunk `idx`, if any widens.
    fn char_extras(&self, idx: usize, text: &str) -> Option<Vec<f32>> {
        let points = cjk_gap_points(text, idx == self.last_text_chunk);
        points.iter().any(|p| *p).then(|| {
            points
                .into_iter()
                .map(|p| if p { self.extra } else { 0.0 })
                .collect()
        })
    }
}

/// Per character of `text`, whether the gap after it is a justification
/// point (see `MixedJustify`); `last` marks the line's last text chunk.
fn cjk_gap_points(text: &str, last: bool) -> Vec<bool> {
    let chars: Vec<char> = text.chars().collect();
    let n = chars.len();
    (0..n)
        .map(|i| {
            let c = chars[i];
            if is_break_space(c) {
                false
            } else if crate::docx::is_east_asian_char(c) {
                !(last && i + 1 == n)
            } else {
                chars
                    .get(i + 1)
                    .is_some_and(|next| crate::docx::is_east_asian_char(*next))
            }
        })
        .collect()
}

/// Resolve aligned segment start position given a tab stop, segment runs, and minimum x.
fn resolve_tab_aligned_start(
    stop: &TabStop,
    tab_target: f32,
    seg_runs: &[&Run],
    seen_fonts: &HashMap<String, FontEntry>,
    min_x: f32,
) -> f32 {
    // The segment's leading spaces sit before the aligned text, not inside
    // it (Word probes on a right tab: 0 or 4 leading spaces end the text at
    // the stop alike; 24 push it right by what no longer fits before it).
    let lead = || leading_space_width(seg_runs, seen_fonts);
    match stop.alignment {
        TabAlignment::Left => tab_target.max(min_x),
        TabAlignment::Center => {
            let sw = segment_width(seg_runs, seen_fonts);
            (tab_target - sw / 2.0 - lead()).max(min_x)
        }
        TabAlignment::Right => {
            let sw = segment_width(seg_runs, seen_fonts);
            (tab_target - sw - lead()).max(min_x)
        }
        TabAlignment::Decimal => {
            let bw = decimal_before_width(seg_runs, seen_fonts);
            (tab_target - bw).max(min_x)
        }
    }
}

/// U+00A0 (non-breaking space) renders as a space but must not allow line breaks.
fn is_break_space(c: char) -> bool {
    c.is_whitespace() && c != '\u{00a0}' && c != '\u{3000}'
}

/// Word wraps a URL only after a hyphen, or at the margin once it is wider
/// than the line: over 102 URL line ends in Word PDFs, 2 fall after a `/`.
/// UAX #14 also allows breaks after `/`, `?` and the like, so those go.
fn drop_url_breaks(text: &str, breaks: &mut Vec<usize>) {
    let mut urls = Vec::new();
    let mut from = 0;
    while let Some(rel) = [text[from..].find("://"), text[from..].find("www.")]
        .into_iter()
        .flatten()
        .min()
    {
        let at = from + rel;
        let start = text[..at]
            .rfind(|c: char| !c.is_ascii_alphanumeric())
            .map_or(0, |i| i + 1);
        let end = text[at..]
            .find(char::is_whitespace)
            .map_or(text.len(), |i| at + i);
        urls.push(start..end);
        from = end;
    }
    breaks.retain(|&b| !urls.iter().any(|u| u.start < b && b < u.end) || text[..b].ends_with('-'));
}

/// Word's departures from UAX #14 that depend only on the two adjacent
/// characters: `Some(true)` adds a break, `Some(false)` removes one, `None`
/// keeps UAX #14's answer. The URL rule needs more context (`drop_url_breaks`).
fn word_pair_rule(a: char, b: char) -> Option<bool> {
    // A period immediately followed by an opening parenthesis remains joined.
    if a == '.' && b == '(' {
        return Some(false);
    }
    // Word keeps an abbreviation joined to an adjacent numero sign.
    if a == '.' && b == '№' {
        return Some(false);
    }
    // Word breaks after a hyphen-minus where UAX #14 does not: before another
    // hyphen (LB21 would make a dash run one unbreakable word; family_kinship's
    // 87-dash rules wrap into two columns of dashes) and before a digit (LB25
    // keeps "2019-2024" whole; 13 reference line ends such as "Sindh 2019-",
    // "1(4): 108-", "about 3-"). Opening guillemets also permit a break:
    // Word wraps `Office «North Ural»-«Branch»` after the hyphen.
    if a == '-' && (b == '-' || b.is_ascii_digit() || b == '«') {
        return Some(true);
    }
    // Word breaks after any breaking space, also before what LB13 glues to it
    // (`.`, `,`, `/`, `)`, `!`): czech_wastewater_discharge_permit wraps
    // "Telefon" + 3 spaces + a dot leader, czech_works_contract before
    // ",,zákon“", stem_partnerships before "/ events".
    if is_break_space(a) && !is_break_space(b) {
        return Some(true);
    }
    // UAX #14 breaks after a solidus, Word doesn't: croatian_grant keeps
    // "troškova/izdataka" whole and wraps before it.
    if a == '/' && b.is_alphanumeric() {
        return Some(false);
    }
    // Class IN allows a break after an ellipsis before digits, but Word keeps
    // tokens like TOC dot-leaders typed as "…………45" unbreakable.
    if matches!(a, '\u{2024}' | '\u{2025}' | '\u{2026}') && !b.is_whitespace() {
        return Some(false);
    }
    None
}

/// `word_pair_rule` for the characters either side of byte offset `i`.
fn word_pair_rule_at(text: &str, i: usize) -> Option<bool> {
    word_pair_rule(text[..i].chars().next_back()?, text[i..].chars().next()?)
}

/// True when `split_preserving_spaces` would break between `a` and `b` inside one
/// run: UAX #14 plus Word's pair rules. Run boundaries call this for every
/// glued word, so it stays off the heap.
pub(super) fn breaks_between(a: char, b: char) -> bool {
    if let Some(rule) = word_pair_rule(a, b) {
        return rule;
    }
    let mut buf = [0u8; 8];
    let la = a.encode_utf8(&mut buf).len();
    let lb = b.encode_utf8(&mut buf[la..]).len();
    let Ok(pair) = std::str::from_utf8(&buf[..la + lb]) else {
        return false;
    };
    unicode_linebreak::linebreaks(pair).any(|(pos, _)| pos == la)
}

/// Split text into (preceding_space_count, word) segments using UAX #14 line break rules.
/// Handles CJK character boundaries, hyphens, and punctuation break opportunities
/// in addition to whitespace. Non-breaking spaces (U+00A0) and ideographic spaces
/// (U+3000) are kept within words. Trailing spaces after the last word are handled
/// separately by the caller.
pub(super) fn split_preserving_spaces(text: &str) -> Vec<(usize, &str)> {
    if text.is_empty() {
        return Vec::new();
    }

    let mut result = Vec::new();
    let mut pending_spaces: usize = 0;

    // Collect UAX #14 break positions (byte offsets where breaks are allowed/mandatory)
    let mut breaks: Vec<usize> = unicode_linebreak::linebreaks(text)
        .map(|(pos, _)| pos)
        .collect();
    drop_url_breaks(text, &mut breaks);
    breaks.retain(|&b| word_pair_rule_at(text, b) != Some(false));
    breaks.extend(
        text.char_indices()
            .map(|(i, _)| i)
            .filter(|&i| word_pair_rule_at(text, i) == Some(true)),
    );
    breaks.sort_unstable();

    let mut prev = 0;
    for &brk in &breaks {
        let segment = &text[prev..brk];
        prev = brk;

        if segment.is_empty() {
            continue;
        }

        // Count leading break-spaces in this segment
        let leading_bytes: usize = segment
            .chars()
            .take_while(|c| is_break_space(*c))
            .map(|c| c.len_utf8())
            .sum();
        let leading_count = segment[..leading_bytes].chars().count();

        // Count trailing break-spaces
        let trailing_bytes: usize = segment
            .chars()
            .rev()
            .take_while(|c| is_break_space(*c))
            .map(|c| c.len_utf8())
            .sum();
        let trailing_count = if trailing_bytes < segment.len() - leading_bytes {
            segment[segment.len() - trailing_bytes..].chars().count()
        } else {
            0
        };

        pending_spaces += leading_count;

        let word_start = leading_bytes;
        let word_end = segment.len() - trailing_bytes;

        if word_start < word_end {
            result.push((pending_spaces, &segment[word_start..word_end]));
            pending_spaces = trailing_count;
        } else {
            // Segment is all spaces
            pending_spaces += trailing_count;
        }
    }

    result
}

#[derive(Clone)]
pub(super) struct WordChunk {
    pub(super) pdf_font: String,
    pub(super) text: String,
    pub(super) font_size: f32,
    /// The run's own size before small caps or super/subscript shrink it; a
    /// line is as tall as its runs at this size (`size_lines_by_own_runs`).
    pub(super) line_font_size: f32,
    pub(super) color: Option<[u8; 3]>,
    pub(super) highlight: Option<[u8; 3]>,
    pub(super) shading: Option<[u8; 3]>,
    pub(super) border: Option<ParagraphBorder>,
    pub(super) x_offset: f32, // x relative to line start
    pub(super) width: f32,
    pub(super) underline: bool,
    pub(super) double_underline: bool,
    pub(super) strikethrough: bool,
    pub(super) dstrike: bool,
    pub(super) char_spacing: f32,
    pub(super) text_scale: f32, // percentage, 100.0 = normal
    pub(super) y_offset: f32,   // superscript/subscript, or inline picture's bottom extent
    /// `w:position` share of `y_offset`: it stretches the line box, unlike
    /// super/subscript (see `size_lines_by_own_runs`).
    pub(super) raise: f32,
    pub(super) hyperlink_url: Option<String>,
    pub(super) inline_image_name: Option<String>,
    pub(super) inline_image_height: f32,
    /// Running-head wrapped pictures can reserve their effect/wrap extent
    /// separately from the rotated visible image box.
    pub(super) inline_image_extra_height: f32,
    pub(super) inline_image_stroke_color: Option<[u8; 3]>,
    pub(super) inline_image_stroke_width: f32,
    pub(super) inline_image_shadow: Option<crate::model::ImageShadow>,
    pub(super) inline_image_glow: Option<crate::model::ImageGlow>,
    pub(super) inline_image_effect_xobjs: Option<super::images::EffectXObjs>,
    pub(super) inline_image_clip: Option<crate::model::ShapeGeometry>,
    /// Natural (unrotated) size; `width`/`inline_image_height` hold the rotated box.
    pub(super) inline_image_size: (f32, f32),
    /// Clockwise degrees (OOXML).
    pub(super) inline_image_rotation_deg: f32,
    /// The picture's `docPr@descr` and adec:decorative flag, for tagging.
    pub(super) inline_image_alt: Option<String>,
    pub(super) inline_image_decorative: bool,
    pub(super) synthetic_bold: bool,
    pub(super) synthetic_italic: bool,
    pub(super) text_outline: Option<TextOutline>,
    pub(super) text_fill: Option<TextFill>,
    /// Legacy w:shadow/emboss/imprint drop-shadow: drawn as an offset gray copy
    /// of the glyphs behind the main text.
    pub(super) text_shadow: Option<TextShadow>,
    pub(super) comment_ids: Vec<u32>,
    /// Footnote this chunk is the reference mark of. Word puts a footnote on
    /// the page where its reference mark lands, so pagination needs to know
    /// which line carries which reference.
    pub(super) footnote_id: Option<u32>,
    /// Endnote this chunk is the reference mark of (tagging only).
    pub(super) endnote_id: Option<u32>,
    /// The source letters when caps/small caps draw different ones, given to
    /// text extraction and screen readers as `/ActualText`.
    pub(super) actual_text: Option<String>,
    /// The language of this text (`w:lang`), for a `/Lang` Span when it isn't
    /// the document's.
    pub(super) lang: Option<Arc<str>>,
    /// Points already trimmed from a trailing full-width punctuation mark by
    /// `compress_punctuation`; caps further squeezing at half an em.
    pub(super) punct_compressed: f32,
    /// A break-space followed this word in the source. Words are positioned
    /// individually, so the space is drawn as an invisible glyph after the word
    /// purely so text extraction and screen readers see the word boundary.
    pub(super) space_after: bool,
    /// Literal break spaces before this chunk (more than one when typed so).
    pub(super) preceding_spaces: usize,
    /// Ordinary word-space decoration; excludes tab-stop decorations.
    pub(super) word_space_bridge: bool,
    /// The math zone's spoken form, for its Formula (see `Run::formula`).
    pub(super) formula: Option<Arc<str>>,
    /// Pair kerning is on for this run at this size (`w:kern`); `width` already
    /// includes it, so the glyphs are drawn kerned too.
    pub(super) kern: bool,
    /// Drawn as a square in place of the text (`Run::checkbox`).
    pub(super) checkbox: Option<FormCheckbox>,
    /// From a `w:rtl` run (bidi embedding, mirrored brackets).
    pub(super) rtl: bool,
    /// From a `w:hint="eastAsia"` run: its curly quotes are full-width
    /// punctuation (see `is_hint_ambiguous`).
    pub(super) east_asia_hint: bool,
    /// At a right-to-left level: glyphs are placed from the right
    /// (`show_rtl`), in logical order.
    pub(super) rtl_glyphs: bool,
    /// The space before this chunk came from a `w:rtl` run: bidi ordering
    /// gives it that run's embedding.
    pub(super) gap_rtl: bool,
    /// An Office Math zone drawn from its box; `text` is its linear form.
    pub(super) math: Option<Arc<crate::math::MathBox>>,
}

/// A run's Office Math box and the math font it draws with, once laid out
/// (`fonts::register_math_font`); None leaves the zone's linear text.
fn math_box<'f>(
    run: &'f Run,
    fonts: &'f HashMap<String, FontEntry>,
) -> Option<(&'f FontEntry, &'f Arc<crate::math::MathBox>)> {
    let entry = fonts.get(super::fonts::MATH_FONT_KEY)?;
    let mb = entry
        .math_boxes
        .get(&Arc::as_ptr(run.math.as_ref()?).addr())?;
    Some((entry, mb))
}

/// Pale-pink highlight color Word uses for comment-anchored text spans.
pub(super) const COMMENT_HIGHLIGHT_RGB: [u8; 3] = [251, 220, 217];

/// A decoration rectangle (underline, strikethrough, shading): x, y, width, height, colour.
type Decoration = (f32, f32, f32, f32, Option<[u8; 3]>);

fn push_decoration(
    decorations: &mut Vec<Decoration>,
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    color: Option<[u8; 3]>,
) {
    // Look back past the other line of a double underline (or a strike on the
    // same run), which interleave with this one chunk by chunk.
    // Only a touching span joins: an underlined word after a gap of plain text
    // or tabs gets its own line.
    let merged = decorations
        .iter_mut()
        .rev()
        .take(3)
        .find(|(dx, dy, dw, dh, dc)| {
            (*dy - y).abs() < 0.01
                && (*dh - height).abs() < 0.01
                && *dc == color
                && (x - (*dx + *dw)).abs() < 0.5
        });
    if let Some(prev) = merged {
        prev.2 = (x + width) - prev.0;
    } else {
        decorations.push((x, y, width, height, color));
    }
}

impl WordChunk {
    /// The note this chunk is the reference mark of: (endnote, id).
    fn note(&self) -> Option<(bool, u32)> {
        self.footnote_id
            .map(|id| (false, id))
            .or(self.endnote_id.map(|id| (true, id)))
    }

    /// A space glyph goes after the word, marking the word boundary for text
    /// extraction (a font without one leaves it out).
    pub(super) fn boundary_space(&self, entry: Option<&FontEntry>) -> bool {
        self.space_after && entry.is_none_or(|e| e.has_char(' '))
    }

    fn text(
        entry: &FontEntry,
        run: &Run,
        word: &str,
        eff_fs: f32,
        char_spacing: f32,
        y_offset: f32,
        x_offset: f32,
        width: f32,
    ) -> Self {
        Self {
            pdf_font: entry.pdf_name.clone(),
            text: word.to_string(),
            font_size: eff_fs,
            line_font_size: run.font_size,
            color: run.color,
            highlight: run.highlight,
            shading: if run.shading.is_none() && !run.comment_ids.is_empty() {
                Some(COMMENT_HIGHLIGHT_RGB)
            } else {
                run.shading
            },
            border: run.border.clone(),
            x_offset,
            width,
            underline: run.underline,
            double_underline: run.double_underline,
            strikethrough: run.strikethrough,
            dstrike: run.dstrike,
            char_spacing,
            text_scale: run.text_scale,
            y_offset,
            raise: run.position,
            hyperlink_url: run.hyperlink_url.clone(),
            inline_image_name: None,
            inline_image_height: 0.0,
            inline_image_extra_height: 0.0,
            inline_image_stroke_color: None,
            inline_image_stroke_width: 0.0,
            inline_image_shadow: None,
            inline_image_glow: None,
            inline_image_effect_xobjs: None,
            inline_image_clip: None,
            inline_image_size: (0.0, 0.0),
            inline_image_rotation_deg: 0.0,
            inline_image_alt: None,
            inline_image_decorative: false,
            punct_compressed: 0.0,
            synthetic_bold: entry.synthetic_bold,
            synthetic_italic: entry.synthetic_italic,
            text_outline: run.text_outline.clone(),
            text_fill: run.text_fill.clone(),
            text_shadow: run.text_shadow.clone(),
            comment_ids: [&run.comment_ids[..], &run.comment_anchor_ids[..]].concat(),
            footnote_id: run.footnote_id,
            endnote_id: run.endnote_id,
            actual_text: None,
            lang: chunk_lang(run, word),
            space_after: false,
            preceding_spaces: 1,
            word_space_bridge: false,
            formula: run.formula.clone(),
            kern: run.kerns_at(eff_fs),
            checkbox: run.checkbox,
            rtl: run.rtl,
            east_asia_hint: run.east_asia_hint,
            rtl_glyphs: false,
            gap_rtl: false,
            math: None,
        }
    }

    /// A laid-out Office Math zone at `x_offset`, drawn in the math font.
    fn math(entry: &FontEntry, run: &Run, mb: &Arc<crate::math::MathBox>, x_offset: f32) -> Self {
        let mut chunk = Self::text(
            entry,
            run,
            &run.text,
            run.font_size,
            0.0,
            0.0,
            x_offset,
            mb.width,
        );
        chunk.math = Some(mb.clone());
        chunk
    }

    fn image(
        pdf_name: &str,
        font_size: f32,
        x_offset: f32,
        img: &crate::model::EmbeddedImage,
        effect_xobjs: Option<super::images::EffectXObjs>,
    ) -> Self {
        let (width, height) = img.layout_size();
        Self {
            pdf_font: String::new(),
            text: String::new(),
            font_size,
            line_font_size: font_size,
            color: None,
            highlight: None,
            shading: None,
            border: None,
            x_offset,
            width,
            underline: false,
            double_underline: false,
            strikethrough: false,
            dstrike: false,
            char_spacing: 0.0,
            text_scale: 100.0,
            y_offset: 0.0,
            raise: 0.0,
            hyperlink_url: None,
            inline_image_name: Some(pdf_name.to_string()),
            inline_image_height: height,
            inline_image_extra_height: 0.0,
            inline_image_stroke_color: img.stroke_color,
            inline_image_stroke_width: img.stroke_width,
            inline_image_shadow: img.shadow.clone(),
            inline_image_glow: img.glow.clone(),
            inline_image_effect_xobjs: effect_xobjs,
            inline_image_clip: img.clip_geometry.clone(),
            inline_image_size: (img.display_width, img.display_height),
            inline_image_rotation_deg: img.rotation_deg,
            inline_image_alt: img.alt.clone(),
            inline_image_decorative: img.decorative,
            punct_compressed: 0.0,
            formula: None,
            kern: false,
            checkbox: None,
            rtl: false,
            east_asia_hint: false,
            rtl_glyphs: false,
            gap_rtl: false,
            math: None,
            synthetic_bold: false,
            synthetic_italic: false,
            text_outline: None,
            text_fill: None,
            text_shadow: None,
            comment_ids: Vec::new(),
            footnote_id: None,
            endnote_id: None,
            actual_text: None,
            lang: None,
            space_after: false,
            preceding_spaces: 1,
            word_space_bridge: false,
        }
    }

    fn leader(
        entry: &FontEntry,
        text: String,
        font_size: f32,
        color: Option<[u8; 3]>,
        x_offset: f32,
        width: f32,
    ) -> Self {
        Self {
            pdf_font: entry.pdf_name.clone(),
            text,
            font_size,
            line_font_size: font_size,
            color,
            highlight: None,
            shading: None,
            border: None,
            x_offset,
            width,
            underline: false,
            double_underline: false,
            strikethrough: false,
            dstrike: false,
            char_spacing: 0.0,
            text_scale: 100.0,
            y_offset: 0.0,
            raise: 0.0,
            hyperlink_url: None,
            inline_image_name: None,
            inline_image_height: 0.0,
            inline_image_extra_height: 0.0,
            inline_image_stroke_color: None,
            inline_image_stroke_width: 0.0,
            inline_image_shadow: None,
            inline_image_glow: None,
            inline_image_effect_xobjs: None,
            inline_image_clip: None,
            inline_image_size: (0.0, 0.0),
            inline_image_rotation_deg: 0.0,
            inline_image_alt: None,
            inline_image_decorative: false,
            punct_compressed: 0.0,
            formula: None,
            kern: false,
            checkbox: None,
            rtl: false,
            east_asia_hint: false,
            rtl_glyphs: false,
            gap_rtl: false,
            math: None,
            synthetic_bold: false,
            synthetic_italic: false,
            text_outline: None,
            text_fill: None,
            text_shadow: None,
            comment_ids: Vec::new(),
            footnote_id: None,
            endnote_id: None,
            actual_text: None,
            lang: None,
            space_after: false,
            preceding_spaces: 1,
            word_space_bridge: false,
        }
    }

    fn word_space_underline(
        entry: &FontEntry,
        font_size: f32,
        color: Option<[u8; 3]>,
        double_underline: bool,
        border: Option<ParagraphBorder>,
        x_offset: f32,
        width: f32,
    ) -> Self {
        let mut chunk = Self::tab_underline(
            entry,
            font_size,
            color,
            double_underline,
            border,
            x_offset,
            width,
        );
        chunk.word_space_bridge = true;
        chunk
    }

    /// Decoration-only chunk: no glyphs, but the underline decoration pass
    /// draws a line spanning the chunk's width. Used for underlined tab gaps.
    fn tab_underline(
        entry: &FontEntry,
        font_size: f32,
        color: Option<[u8; 3]>,
        double_underline: bool,
        border: Option<ParagraphBorder>,
        x_offset: f32,
        width: f32,
    ) -> Self {
        Self {
            pdf_font: entry.pdf_name.clone(),
            text: String::new(),
            font_size,
            line_font_size: font_size,
            color,
            highlight: None,
            shading: None,
            // Carry the run's border so a bridged space keeps the border span
            // contiguous (otherwise a None-border space chunk splits the box).
            border,
            x_offset,
            width,
            underline: true,
            double_underline,
            strikethrough: false,
            dstrike: false,
            char_spacing: 0.0,
            text_scale: 100.0,
            y_offset: 0.0,
            raise: 0.0,
            hyperlink_url: None,
            inline_image_name: None,
            inline_image_height: 0.0,
            inline_image_extra_height: 0.0,
            inline_image_stroke_color: None,
            inline_image_stroke_width: 0.0,
            inline_image_shadow: None,
            inline_image_glow: None,
            inline_image_effect_xobjs: None,
            inline_image_clip: None,
            inline_image_size: (0.0, 0.0),
            inline_image_rotation_deg: 0.0,
            inline_image_alt: None,
            inline_image_decorative: false,
            punct_compressed: 0.0,
            formula: None,
            kern: false,
            checkbox: None,
            rtl: false,
            east_asia_hint: false,
            rtl_glyphs: false,
            gap_rtl: false,
            math: None,
            synthetic_bold: false,
            synthetic_italic: false,
            text_outline: None,
            text_fill: None,
            text_shadow: None,
            comment_ids: Vec::new(),
            footnote_id: None,
            endnote_id: None,
            actual_text: None,
            lang: None,
            space_after: false,
            preceding_spaces: 1,
            word_space_bridge: false,
        }
    }
}

pub(crate) struct LinkAnnotation {
    pub(super) rect: Rect,
    pub(super) url: String,
    /// Link structure element the annotation belongs to (tagged body text).
    pub(super) node: Option<usize>,
    /// Link text, for the annotation's /Contents description.
    pub(super) text: String,
}

/// Tags a body paragraph's hyperlinks as Link elements with their own marked
/// content while its lines are drawn (see `tagging`).
pub(super) struct LinkTagger<'a> {
    pub(super) tags: &'a mut super::tagging::Tags,
    pub(super) page: usize,
    pub(super) para: usize,
    link: Option<(String, usize)>,
    /// The open Span or Formula (see `inline`) and its element.
    inline: Option<(Inline, usize)>,
}

/// A stretch of chunks tagged inside the open Link or paragraph.
enum Inline {
    /// Its `/Lang`, and whether it carries `/ActualText`.
    Span { lang: Option<String>, actual: bool },
    /// An Office Math zone, by its spoken form.
    Formula(Arc<str>),
}

impl<'a> LinkTagger<'a> {
    pub(super) fn new(tags: &'a mut super::tagging::Tags, page: usize, para: usize) -> Self {
        Self {
            tags,
            page,
            para,
            link: None,
            inline: None,
        }
    }

    /// Text in another language than the document's goes in a Span with
    /// `/Lang`, so a screen reader switches voice; caps and small caps, which
    /// draw other letters than the source has, in one whose `/ActualText`
    /// carries the source. One Span per stretch of chunks that need the same,
    /// inside the open Link or paragraph: a structure element, not nested
    /// marked content, which Poppler's structure reader loses text after.
    /// An Office Math zone's text goes in a Formula instead, whose `/Alt`
    /// speaks all of it (Word: "cap T equals 2 pi"); its glyphs stay the
    /// content for text extraction. Returns true when the marked content
    /// switched (the text matrix is reset).
    fn inline(
        &mut self,
        content: &mut Content,
        formula: Option<&Arc<str>>,
        lang: Option<&str>,
        actual: Option<&str>,
    ) -> bool {
        let lang = lang.filter(|l| !self.tags.is_document_lang(l));
        let same = match (&self.inline, formula) {
            (Some((Inline::Formula(open), _)), Some(f)) => Arc::ptr_eq(open, f),
            (Some((Inline::Span { lang: l, actual: a }, span)), None) => {
                let same = l.as_deref() == lang && *a == actual.is_some();
                if let (true, Some(text)) = (same, actual) {
                    self.tags.push_actual(*span, text);
                }
                same
            }
            (None, None) => lang.is_none() && actual.is_none(),
            _ => false,
        };
        if same {
            return false;
        }
        content.end_text();
        let parent = self.open();
        self.inline = match formula {
            Some(f) => Some((Inline::Formula(f.clone()), self.tags.add_formula(parent, f))),
            None if lang.is_some() || actual.is_some() => {
                let span = self.tags.add_span(parent, lang, actual);
                let state = Inline::Span {
                    lang: lang.map(str::to_string),
                    actual: actual.is_some(),
                };
                Some((state, span))
            }
            None => None,
        };
        let node = self.inline.as_ref().map_or(parent, |i| i.1);
        self.tags.begin(content, self.page, node);
        content.begin_text();
        true
    }

    /// Before drawing `chunk` in a text object: switch to its Link (a note's
    /// reference mark links to the note), Span or Formula, and nest its Note.
    /// Returns true when the marked content switched (the text matrix is
    /// reset).
    pub(super) fn chunk(
        &mut self,
        content: &mut Content,
        chunk: &WordChunk,
        link_url: Option<&str>,
        boundary_space: bool,
    ) -> bool {
        let mut switched = false;
        // Glyph-less chunks (underline bridges over spaces) don't break a link.
        if !chunk.text.is_empty() {
            switched |= self.enter(content, link_url);
            // A Span's /ActualText covers the boundary space the Tj carries too.
            let actual = chunk.actual_text.as_ref().map(|t| {
                if boundary_space {
                    format!("{t} ")
                } else {
                    t.clone()
                }
            });
            let formula = chunk.formula.as_ref();
            // Punctuation and digits have no language: they stay in the open
            // language Span (the commas between Arabic words aren't English).
            let neutral = formula.is_none()
                && actual.is_none()
                && !chunk.text.chars().any(char::is_alphabetic);
            let lang = match &self.inline {
                None if neutral => None,
                Some((
                    Inline::Span {
                        lang,
                        actual: false,
                    },
                    _,
                )) if neutral => lang.clone(),
                _ => chunk.lang.as_deref().map(str::to_string),
            };
            switched |= self.inline(content, formula, lang.as_deref(), actual.as_deref());
        }
        // The Note goes inside the link on its reference mark (Word nests it there).
        if let Some((endnote, id)) = chunk.note() {
            let parent = self.open();
            self.tags.note(endnote, id, parent);
        }
        switched
    }

    /// Switch the open marked content to the chunk's Link (or back to the
    /// paragraph). Like Word, marked content never changes inside a text
    /// object, so it is ended and reopened around the switch; returns true
    /// when it did (the text matrix is reset).
    fn enter(&mut self, content: &mut Content, url: Option<&str>) -> bool {
        if self.link.as_ref().map(|(u, _)| u.as_str()) == url {
            return false;
        }
        content.end_text();
        self.inline = None;
        let node = match url {
            Some(u) => {
                let n = self.tags.add(self.para, "Link");
                self.link = Some((u.to_string(), n));
                n
            }
            None => {
                self.link = None;
                self.para
            }
        };
        self.tags.begin(content, self.page, node);
        content.begin_text();
        true
    }

    /// The element the current text belongs to.
    fn open(&self) -> usize {
        self.link.as_ref().map_or(self.para, |&(_, n)| n)
    }

    /// Before drawing an inline picture (outside a text object): a Figure
    /// inside the open element, read where the picture sits in the text, or
    /// an artifact when the picture is decorative.
    fn begin_picture(&mut self, content: &mut Content, alt: Option<&str>, decorative: bool) {
        if decorative {
            super::tagging::Tags::end(content);
        } else {
            let figure = self.tags.add_figure(self.open(), alt);
            self.tags.begin(content, self.page, figure);
        }
    }

    /// Back to the open element after a picture or a Span.
    fn resume(&mut self, content: &mut Content) {
        self.inline = None;
        self.tags.begin(content, self.page, self.open());
    }

    /// Draw something that isn't the paragraph's content (outside a text
    /// object) as an artifact, then go on in the open Span, Link or paragraph.
    fn artifact(&mut self, content: &mut Content, draw: impl FnOnce(&mut Content)) {
        super::tagging::Tags::end(content);
        draw(content);
        let node = self.inline.as_ref().map_or_else(|| self.open(), |i| i.1);
        self.tags.begin(content, self.page, node);
    }

    /// `artifact` inside a text object (a text shadow's gray copy): the
    /// object is closed around it, so the text matrix starts afresh.
    fn text_artifact(&mut self, content: &mut Content, draw: impl FnOnce(&mut Content)) {
        content.end_text();
        self.artifact(content, |content| {
            content.begin_text();
            draw(content);
            content.end_text();
        });
        content.begin_text();
    }

    pub(super) fn finish(mut self, content: &mut Content) {
        if self.link.take().is_some() | self.inline.take().is_some() {
            self.tags.begin(content, self.page, self.para);
        }
    }
}

/// When a line spans two text regions (bothSides wrapping around a float),
/// this stores where the right region begins and its geometry.
pub(super) struct RightRegion {
    pub(super) first_chunk_idx: usize,
    pub(super) region_x: f32,
    pub(super) region_width: f32,
    pub(super) content_width: f32,
}

#[derive(Default)]
pub(super) struct TextLine {
    pub(super) chunks: Vec<WordChunk>,
    pub(super) total_width: f32,
    pub(super) ends_with_break: bool,
    pub(super) right_region: Option<RightRegion>,
    /// Font size of the line break run that created this empty line.
    /// Used to compute the correct line height for break-created lines
    /// (Word uses the break run's font metrics, not the paragraph's).
    pub(super) break_font_size: Option<f32>,
    /// Line-height ratio of the break run that ended this otherwise empty line.
    pub(super) break_lhr: Option<f32>,
    /// An empty line made by leading spaces wrapping an inline picture.
    pub(super) leading_space_wrap: bool,
    /// An empty line retained by a wrapping tab; its consumed tabs size the line.
    pub(super) leading_tab_wrap: bool,
    /// First chunk after the line's last tab: justification stretches only the
    /// gaps from here on (Word starts the text at the tab stop).
    pub(super) justify_from: usize,
    /// The breaker kept this line's last word by narrowing its spaces
    /// (`SPACE_SQUEEZE`), so it is wider than the measure until justified.
    pub(super) squeezed: bool,
    /// The squeeze narrows every character, not just the spaces
    /// (`CjkLayout::squeeze_line`).
    pub(super) squeeze_chars: bool,
    /// Left at natural width although justified (`doNotExpandShiftReturn`).
    pub(super) natural_width: bool,
    /// An empty line beside a float that could not hold the next word: Word
    /// moves on to the next line (and takes the first-line indent there).
    pub(super) float_skip: bool,
    /// This line's own advance when its tallest face differs from the
    /// paragraph's (see `size_lines_by_own_runs`); None uses the paragraph pitch.
    pub(super) pitch: Option<f32>,
    /// Chunk indices in logical order when bidi reordering stored the chunks
    /// visually (`bidi::reorder_lines`); empty: the stored order.
    pub(super) draw_order: Vec<usize>,
    /// Width of the space the line wrapped at (dropped from the line); bidi
    /// ordering places it (`bidi::reorder_lines`).
    pub(super) trailing_space: f32,
    /// That space came from a `w:rtl` run.
    pub(super) trailing_space_rtl: bool,
    /// In a `w:bidi` paragraph: the first-line indent is on the right and a
    /// justified paragraph's last line ends at the right.
    pub(super) rtl_paragraph: bool,
    /// How far this line's baseline sits below where the paragraph ascent
    /// would put it (negative: above), for lines with their own `pitch`.
    pub(super) ascent_shift: f32,
}

/// Word sizes each line by the runs on that line, not the paragraph's tallest
/// run: russian_sports_ranking_decree's 20pt "ГЛАВА" followed by `w:br` lines
/// of 14pt steps 17.25 then 16.0 in Word, not 23.0 twice. Lines that come out
/// at the paragraph pitch keep it; whitespace and picture chunks do not size a
/// line, and text-less lines (breaks) keep their own handling.
pub(super) fn size_lines_by_own_runs(
    lines: &mut [TextLine],
    fonts: &HashMap<String, FontEntry>,
    line_spacing: LineSpacing,
    para_pitch: f32,
    para_ascent: f32,
    top_aligned: bool,
) {
    let by_pdf_name: HashMap<&str, &FontEntry> =
        fonts.values().map(|e| (e.pdf_name.as_str(), e)).collect();
    for line in lines.iter_mut() {
        if line.leading_space_wrap || line.leading_tab_wrap {
            if let Some(fs) = line.break_font_size {
                line.pitch = Some(super::resolve_line_h(line_spacing, fs, line.break_lhr));
            }
            continue;
        }
        if line.chunks.iter().any(|c| c.inline_image_name.is_some()) {
            continue;
        }
        // The line spans the highest top and the lowest bottom of its runs,
        // which can come from different runs: case8's 16pt pixel font has
        // almost no descent, so the 12pt Arial beside it sets the bottom.
        // Math faces never size a line (their operator metrics are huge).
        let (mut ascent, mut below, mut sized) = (0.0f32, 0.0f32, false);
        for mb in line.chunks.iter().filter_map(|c| c.math.as_ref()) {
            ascent = ascent.max(mb.line_ascent);
            below = below.max(mb.line_descent);
            sized = true;
        }
        for c in line
            .chunks
            .iter()
            .filter(|c| c.math.is_none() && !c.text.trim().is_empty())
        {
            let Some(entry) = by_pdf_name.get(c.pdf_font.as_str()) else {
                continue;
            };
            let (lhr, ar) = run_line_metrics(entry, &c.text);
            let (lhr, ar) = (lhr.unwrap_or(1.2), ar.unwrap_or(0.75));
            // Word makes room for a run border's box around the glyphs
            // (run-borders: 12pt Calibri in a 1.44pt border steps
            // (14.65 + 2 × 1.44) × 1.079 = 18.96).
            let pad = c.border.as_ref().map_or(0.0, |b| b.width_pt + b.space_pt);
            // A raised or lowered run keeps the line's box: lithuanian_excise's
            // subscript "CO2" lines step like the lines around them.
            // Small caps' 80% letters and raised/lowered runs keep the run's
            // own size: italian_project_proposal's small-caps cell lines step
            // 14.64 like full-size ones.
            // Position shifts both edges. A wholly raised line can lose its
            // descent before growing taller; an unshifted run alongside it
            // retains that descent. Do not stretch both sides independently.
            if top_aligned {
                ascent = ascent.max(c.line_font_size * ar + pad);
                below = below.max(c.line_font_size * (lhr - ar).max(0.0) + pad + top_aligned_drop(entry, c.line_font_size, c.raise));
            } else {
                ascent = ascent.max((c.line_font_size * ar + pad + c.raise).max(0.0));
                below = below.max((c.line_font_size * (lhr - ar).max(0.0) + pad - c.raise).max(0.0));
            }
            sized = true;
        }
        if !sized {
            continue;
        }
        let (above_extra, below_extra) = math_line_stretch(line, ascent, below, line_spacing);
        let pitch = super::helpers::resolve_line_h(line_spacing, 1.0, Some(ascent + below))
            + above_extra
            + below_extra;
        if (pitch - para_pitch).abs() > 0.01
            || (line.chunks.iter().any(|c| c.raise != 0.0) && (ascent - para_ascent).abs() > 0.01)
        {
            line.pitch = Some(pitch);
            line.ascent_shift = ascent + above_extra - para_ascent;
        }
    }
}

/// How far a run moved by `raise` stretches a `w:textAlignment="top"` line
/// below. Word hangs the glyphs from the line top there, so the external
/// leading sits under them and absorbs a drop first; raising never grows the
/// line. Probes (2026-10-22, online export) of Calibri, Arial and Times New
/// Roman at 11 and 24pt lowered 0.5-7pt: growth = drop − leading, ≥ 0.
pub(super) fn top_aligned_drop(entry: &FontEntry, size: f32, raise: f32) -> f32 {
    (-raise - size * entry.leading_ratio.unwrap_or(0.0)).max(0.0)
}

/// How far a line's math boxes reach past its text box (`ascent` above,
/// `below` under the baseline, before line spacing), added to the line
/// unscaled: ink above the ascent less the font's MathLeading, and ink below
/// the descent, stretch it (OpenType MATH `MathLeading`). Word probes at 12
/// and 24pt in Calibri and Arial, single and 1.5 spacing: an inline a/b line
/// grows 1.75-2.0pt above and 2.5-3.5pt below whatever the multiple.
fn math_line_stretch(line: &TextLine, ascent: f32, below: f32, ls: LineSpacing) -> (f32, f32) {
    if matches!(ls, LineSpacing::Exact(_)) {
        return (0.0, 0.0);
    }
    line.chunks
        .iter()
        .filter_map(|c| c.math.as_ref())
        .fold((0.0f32, 0.0f32), |(a, b), mb| {
            (
                a.max(mb.ascent + mb.leading - ascent),
                b.max(mb.descent - below),
            )
        })
}

/// What `size_lines_by_own_runs` adds to a one-line paragraph whose runs are
/// raised or lowered (the highest raise above, the deepest drop below), for a
/// height estimate that builds no lines. ponytail: assumes the positioned runs
/// set the line's ascent and descent; build the lines if mixed sizes matter.
pub(super) fn position_stretch(runs: &[Run], ls: LineSpacing) -> f32 {
    let (up, down) = runs
        .iter()
        .filter(|r| sizes_line(r) && !r.text.is_empty())
        .fold((0.0f32, 0.0f32), |(u, d), r| {
            (u.max(r.position), d.max(-r.position))
        });
    match ls {
        LineSpacing::Auto(m) => (up + down) * m,
        LineSpacing::AtLeast(_) => up + down,
        LineSpacing::Exact(_) => 0.0,
    }
}

/// True when a paragraph has no visible text (may still have phantom font-info runs).
pub(super) fn is_text_empty(runs: &[Run]) -> bool {
    runs.iter().all(|r| {
        r.vanish || (r.text.is_empty() && !r.is_tab && !r.is_line_break && r.inline_image.is_none())
    })
}

/// Word sizes a superscript or subscript by the face's OS/2 script size,
/// rounded to the nearest half point: Aptos (0.600) 12pt → 7.0, Palatino
/// (0.601) 10pt → 6.0, Times/Arial/Calibri (0.650) 12pt → 8.0, 11pt → 7.0,
/// 9.5pt → 6.0 (census over 30 fixtures' references).
fn effective_font_size(run: &Run, entry: &FontEntry) -> f32 {
    let ratio = match run.vertical_align {
        VertAlign::Superscript => entry.superscript_ratio,
        VertAlign::Subscript => entry.subscript_ratio,
        VertAlign::Baseline => return run.font_size,
    };
    // ponytail: 0.65 is the common OS/2 value, for faces without one (Type1 fallback).
    round_half_point(run.font_size * ratio.unwrap_or(0.65))
    // Note: smallCaps sizing is handled per-segment via smallcaps_segments()
}

fn effective_text(run: &Run) -> Cow<'_, str> {
    if run.caps {
        Cow::Owned(run.text.to_uppercase())
    } else {
        Cow::Borrowed(&run.text)
    }
    // Note: smallCaps uppercasing is handled per-segment via smallcaps_segments()
}

/// A word as `w:caps` draws it (capitals); words are uppercased one by one so
/// the word as written stays at hand for `/ActualText`.
fn caps_word<'a>(run: &Run, word: &'a str) -> Cow<'a, str> {
    if run.caps {
        Cow::Owned(word.to_uppercase())
    } else {
        Cow::Borrowed(word)
    }
}

fn round_half_point(pt: f32) -> f32 {
    (pt * 2.0).round() / 2.0
}

/// Word sets small capitals at 80% of the size, to the nearest half point
/// (probe 2026-10-05: 8→6.5, 11→9, 12→9.5, 13→10.5, 20→16).
fn small_caps_size(base_fs: f32) -> f32 {
    round_half_point(base_fs * 0.8)
}

/// In small caps a space beside a lowercase letter is small too (probe:
/// "def X", "X ghi", "12 jk" small; ". 12" full).
fn small_caps_space(prev: Option<char>, next: Option<char>) -> bool {
    prev.is_some_and(char::is_lowercase) || next.is_some_and(char::is_lowercase)
}

/// The width of the word spaces before `word` in `run`, per space; `prev`
/// carries the run's previous word along for the small-caps rule.
fn word_space_width(
    run: &Run,
    entry: &FontEntry,
    eff_fs: f32,
    word: &str,
    prev: &mut Option<char>,
) -> f32 {
    let small = run.small_caps && small_caps_space(*prev, word.chars().next());
    *prev = word.chars().next_back();
    let fs = if small {
        small_caps_size(eff_fs)
    } else {
        eff_fs
    };
    entry.space_width(fs) * run.text_scale / 100.0 + run.char_spacing
}

/// Split a word into (text, font_size, source) segments for smallCaps rendering.
/// Lowercase chars are uppercased and rendered at `small_caps_size`;
/// uppercase chars and non-letters stay at base_fs. `source` is the
/// segment's letters as written.
pub(super) fn smallcaps_segments(word: &str, base_fs: f32) -> Vec<(String, f32, &str)> {
    let reduced = small_caps_size(base_fs);
    let mut segments: Vec<(String, f32, &str)> = Vec::new();
    let mut start = 0;
    let mut prev = None;
    let mut chars = word.char_indices().peekable();
    while let Some((i, ch)) = chars.next() {
        let next = chars.peek().map(|&(_, c)| c);
        let is_lower = ch.is_lowercase() || (ch == ' ' && small_caps_space(prev, next));
        prev = Some(ch);
        let fs = if is_lower { reduced } else { base_fs };
        let display: String = if is_lower {
            ch.to_uppercase().collect()
        } else {
            ch.to_string()
        };
        let end = i + ch.len_utf8();
        if let Some(last) = segments.last_mut()
            && (last.1 - fs).abs() < 0.001
        {
            last.0.push_str(&display);
            last.2 = &word[start..end];
            continue;
        }
        start = i;
        segments.push((display, fs, &word[i..end]));
    }
    segments
}

/// Compute the width of a word, handling smallCaps per-segment sizing.
/// Byte length of the longest prefix of `word` (at least one character)
/// that fits in `room`; None when the whole word fits or is one character.
fn fitting_prefix_len(word: &str, room: f32, width: impl Fn(&str) -> f32) -> Option<usize> {
    let ends: Vec<usize> = word.char_indices().map(|(i, c)| i + c.len_utf8()).collect();
    if ends.len() < 2 {
        return None;
    }
    // Widths grow with the prefix, so the last fitting end is a partition point.
    let fits = ends[..ends.len() - 1].partition_point(|&e| width(&word[..e]) <= room);
    Some(ends[fits.saturating_sub(1)])
}

pub(super) fn word_width_for_run(
    entry: &FontEntry,
    run: &Run,
    word: &str,
    eff_fs: f32,
    kern: bool,
    cs: f32,
    ts: f32,
) -> f32 {
    if let Some(cb) = run.checkbox {
        return checkbox_cell(entry, cb).0;
    }
    if run.small_caps {
        smallcaps_segments(word, eff_fs)
            .iter()
            .map(|(seg, fs, _)| {
                let seg_kern = run.kerns_at(*fs);
                entry.word_width(seg, *fs, seg_kern) * ts + cs * seg.chars().count() as f32
            })
            .sum()
    } else {
        let char_count = word.chars().count();
        entry.word_width(word, eff_fs, kern) * ts + cs * char_count as f32
    }
}

/// Word sets a legacy check box in a square cell as wide and tall as one line
/// of the run's font at the box's size, down to that line's descent, and
/// strokes the box 1pt inside its top and left edges, 1.5pt inside the others
/// (Word probes: Arial, Calibri, Courier New). Returns (side, descent).
fn checkbox_cell(entry: &FontEntry, cb: FormCheckbox) -> (f32, f32) {
    let (line, ascent) = run_line_metrics(entry, FormCheckbox::TEXT);
    let line = line.unwrap_or(1.15);
    (cb.size * line, cb.size * (line - ascent.unwrap_or(0.9)))
}

fn draw_checkbox(
    content: &mut Content,
    entry: &FontEntry,
    cb: FormCheckbox,
    x: f32,
    baseline: f32,
    color: Option<[u8; 3]>,
) {
    let (cell, descent) = checkbox_cell(entry, cb);
    let side = (cell - 2.5).max(0.5);
    let (left, bottom) = (x + 1.0, baseline - descent + 1.5);
    content.save_state();
    stroke_color_or_black(content, color);
    content.set_line_width(0.75);
    content.rect(left, bottom, side, side);
    content.stroke();
    if cb.checked {
        content.set_line_width(0.5);
        content.move_to(left, bottom);
        content.line_to(left + side, bottom + side);
        content.move_to(left, bottom + side);
        content.line_to(left + side, bottom);
        content.stroke();
    }
    content.restore_state();
}

/// Record a break-space before the next word on the last glyph chunk laid out
/// so far (see `WordChunk::space_after`).
fn mark_space_after(chunks: &mut [WordChunk]) {
    if let Some(c) = chunks.iter_mut().rev().find(|c| !c.text.is_empty()) {
        c.space_after = true;
    }
}

#[derive(Clone, Copy, PartialEq)]
pub(super) enum Script {
    Latin,
    EastAsian,
    Complex,
}

impl Script {
    pub(super) fn of(ch: char) -> Script {
        if crate::docx::is_complex_script_char(ch) {
            Script::Complex
        } else if crate::docx::is_east_asian_char(ch) {
            Script::EastAsian
        } else {
            Script::Latin
        }
    }
}

/// The language of a run's text: East Asian text's from `w:lang/@eastAsia`,
/// complex-script text's (Arabic, Hebrew, …) from `@bidi`, the rest from `@val`.
pub(super) fn run_lang(run: &Run, script: Script) -> Option<&Arc<str>> {
    match script {
        Script::Latin => None,
        Script::EastAsian => run.text_lang_east_asia.as_ref(),
        Script::Complex => run.text_lang_bidi.as_ref(),
    }
    .or(run.text_lang.as_ref())
}

/// The language of `text` in `run`; only scans the text when the run gives
/// East Asian or complex-script letters a language of their own.
fn chunk_lang(run: &Run, text: &str) -> Option<Arc<str>> {
    let own = |l: &Option<Arc<str>>| l.is_some() && *l != run.text_lang;
    let script = if own(&run.text_lang_bidi) || own(&run.text_lang_east_asia) {
        text.chars()
            .map(Script::of)
            .find(|&s| s != Script::Latin)
            .unwrap_or(Script::Latin)
    } else {
        Script::Latin
    };
    run_lang(run, script).cloned()
}

/// Push WordChunks for a word, splitting into per-segment chunks for smallCaps.
/// `original` is the word before `w:caps` uppercased it.
fn push_word_chunks(
    chunks: &mut Vec<WordChunk>,
    entry: &FontEntry,
    run: &Run,
    word: &str,
    original: Option<&str>,
    eff_fs: f32,
    cs: f32,
    y_off: f32,
    x_start: f32,
    total_ww: f32,
    gap_rtl: bool,
    split_punct: bool,
) {
    let first = chunks.len();
    let actual = |shown: &str, source: &str| (shown != source).then(|| source.to_string());
    if run.small_caps {
        let segs = smallcaps_segments(word, eff_fs);
        let mut seg_x = x_start;
        for (seg_text, seg_fs, source) in &segs {
            let seg_kern = run.kerns_at(*seg_fs);
            let ts = run.text_scale / 100.0;
            let seg_w = entry.word_width(seg_text, *seg_fs, seg_kern) * ts
                + cs * seg_text.chars().count() as f32;
            let mut chunk = WordChunk::text(entry, run, seg_text, *seg_fs, cs, y_off, seg_x, seg_w);
            // With caps on too the word is already all capitals, one segment:
            // its source is the caps original.
            let source = match original {
                Some(o) if segs.len() == 1 => o,
                _ => *source,
            };
            chunk.actual_text = actual(seg_text, source);
            chunks.push(chunk);
            seg_x += seg_w;
        }
    } else if split_punct
        && original.is_none()
        && punct_segments(word, run.east_asia_hint).len() > 1
    {
        // UAX #14 glues quotes to their neighbours and brackets to a number
        // (LB30: 2021《), but `compress_punctuation` squeezes a mark only at a
        // chunk's edge: give each its own.
        let mut seg_x = x_start;
        let segs = punct_segments(word, run.east_asia_hint);
        for (i, seg_text) in segs.iter().enumerate() {
            let mut seg_w =
                entry.word_width(seg_text, eff_fs, run.kerns_at(eff_fs)) * run.text_scale / 100.0
                    + cs * seg_text.chars().count() as f32;
            let mut chunk = WordChunk::text(entry, run, seg_text, eff_fs, cs, y_off, seg_x, seg_w);
            // Japanese: a closing mark another mark follows keeps its ink
            // half only (see `ja_adjacent_marks`).
            if run_is_japanese(run)
                && seg_text.chars().last().is_some_and(is_compressible_punct)
                && segs
                    .get(i + 1)
                    .and_then(|next| next.chars().next())
                    .is_some_and(is_fullwidth_mark)
            {
                let cut = 0.5 * eff_fs;
                seg_w -= cut;
                chunk.width = seg_w;
                chunk.punct_compressed = cut;
            }
            chunks.push(chunk);
            seg_x += seg_w;
        }
    } else {
        let mut chunk = WordChunk::text(entry, run, word, eff_fs, cs, y_off, x_start, total_ww);
        chunk.actual_text = original.and_then(|o| actual(word, o));
        chunks.push(chunk);
    }
    if let Some(c) = chunks.get_mut(first) {
        c.gap_rtl = gap_rtl;
    }
}

fn run_is_japanese(run: &Run) -> bool {
    run.text_lang_east_asia
        .as_deref()
        .is_some_and(|l| l.starts_with("ja"))
}

fn is_fullwidth_mark(c: char) -> bool {
    is_compressible_punct(c) || is_compressible_opening(c)
}

/// Closing marks inside `word` that another full-width mark follows: in
/// Japanese text each keeps only its ink half, tight line or not (JIS X 4051;
/// japanese_medical's "院）・専門" draws the ） 5.25pt wide in 10.5pt text on
/// a line 13pt short of the measure).
fn ja_adjacent_marks(word: &str) -> usize {
    word.chars()
        .zip(word.chars().skip(1))
        .filter(|&(a, b)| is_compressible_punct(a) && is_fullwidth_mark(b))
        .count()
}

/// `word` cut before each compressible opening mark and after each closing
/// one (the curly quotes only count in a `hint="eastAsia"` run).
fn punct_segments(word: &str, hint: bool) -> Vec<&str> {
    let mut segs = Vec::new();
    let mut start = 0;
    for (i, c) in word.char_indices() {
        if (is_compressible_opening(c) || (hint && hinted_opening(c))) && i > start {
            segs.push(&word[start..i]);
            start = i;
        }
        if is_compressible_punct(c) || (hint && hinted_closing(c)) {
            let end = i + c.len_utf8();
            segs.push(&word[start..end]);
            start = end;
        }
    }
    if start < word.len() {
        segs.push(&word[start..]);
    }
    segs
}

fn vert_y_offset(run: &Run) -> f32 {
    run.position
        + match run.vertical_align {
            VertAlign::Superscript => run.font_size * 0.35,
            VertAlign::Subscript => -run.font_size * 0.14,
            VertAlign::Baseline => 0.0,
        }
}

const DEFAULT_TAB_INTERVAL: f32 = 36.0; // 0.5 inches

fn is_cjk_punctuation(ch: char) -> bool {
    matches!(ch as u32,
        0x3000..=0x303F  // CJK Symbols and Punctuation (includes ，。、)
        | 0xFF00..=0xFFEF // Halfwidth and Fullwidth Forms
        | 0xFE30..=0xFE4F // CJK Compatibility Forms
    )
}

fn finish_line(chunks: &mut Vec<WordChunk>) -> TextLine {
    let total_width = chunks.last().map(|c| c.x_offset + c.width).unwrap_or(0.0);
    TextLine {
        chunks: std::mem::take(chunks),
        total_width,
        ..TextLine::default()
    }
}

fn finish_line_with_break(chunks: &mut Vec<WordChunk>) -> TextLine {
    let mut line = finish_line(chunks);
    line.ends_with_break = true;
    line
}

/// Paragraph line-breaking switches for `build_paragraph_lines`.
#[derive(Clone, Copy)]
pub(super) struct CjkLayout {
    /// `w:autoSpaceDE`/`DN`: quarter-em gap where Latin text meets ideographs.
    pub(super) auto_space: bool,
    /// `w:characterSpacingControl compressPunctuation`: squeeze full-width
    /// punctuation to keep one more character on the line.
    pub(super) compress_punct: bool,
    /// Whether this paragraph pulls a word up by squeezing (see `CjkLayout`
    /// construction: the 2013+ layout only when justified).
    pub(super) punct_pull: bool,
    /// The most each mark gives, as a share of its em (`PUNCT_SQUEEZE`).
    pub(super) punct_cap: f32,
    /// A word is pulled up only while this share of the measure is free
    /// before it, and this many ems of its font free before it including
    /// its autoSpaceDE gap (one of the two is 0 per compatibility mode, see
    /// `PUNCT_PULL_FREE`).
    pub(super) punct_pull_free_ratio: f32,
    pub(super) punct_pull_free_em: f32,
    /// Justified paragraph under Word 2013+ layout (compatibilityMode 15): a
    /// word that overflows stays on the line when `squeeze_keeps` allows it
    /// and the line's spaces can shrink to make room (see `SPACE_SQUEEZE`).
    pub(super) squeeze_spaces: bool,
    /// Chinese, Japanese or Korean document: a word ending in a Latin mark
    /// may overflow by half the mark on top of what `squeeze_keeps` allows
    /// (Word online probes 2026-10-28: "barn." stays 1.5pt longer with
    /// eastAsia zh-CN than en-US, "barnx" the same in both; issue #142).
    /// The 25% space cap is not widened: slovak_misdemeanor_amendment (zh-CN
    /// defaults) wraps "predstaviteľov." at 28% of its spaces in Word.
    pub(super) punct_half_out: bool,
    /// Justified paragraph under `wpJustification`: a word that overflows by
    /// at most two spaces' width stays, and the whole line is squeezed
    /// (`SQUEEZE_LINE_SPACES`).
    pub(super) squeeze_line: bool,
    /// Justify a line that ends in a manual break, as Word does unless
    /// `doNotExpandShiftReturn` is set.
    pub(super) expand_shift_return: bool,
    /// `w:bidi`: the paragraph's base direction is right to left.
    pub(super) bidi: bool,
}

/// How far Word 2013+ narrows the spaces of a justified line to keep its
/// last word: to 75% of their width. Measured, not from the spec: over 4,216
/// squeezed lines in 207 Word PDFs the per-line space ratio piles up at 0.75
/// with almost nothing below it, and only compat-15 documents squeeze. Of
/// 14,122 line-end decisions in those PDFs, "within this cap and the word's
/// midpoint inside the margin" reproduces 94.6%; either condition alone ~90%.
const SPACE_SQUEEZE: f32 = 0.25;

/// Whether Word 2013+ keeps an overflowing word on a squeezed line. After a
/// hyphen at most a third of the word may stick out; after a space, 0.35 of
/// it plus a third of the space. Word probes at 12 and 14 pt in 0.1 pt steps
/// (2026-10-26): 12 and 24 pt words after a hyphen stay to 4.0 and 8.0 pt
/// over, after a 3 pt space to 5.2 and 9.4 pt; "half the word" kept short
/// words 1-3 pt too long.
fn squeeze_keeps(overflow: f32, word_w: f32, lead_space: f32) -> bool {
    let share = if lead_space > 0.0 { 0.35 } else { 1.0 / 3.0 };
    overflow <= share * word_w + lead_space / 3.0
}

/// How far a `wpJustification` line may overflow and still keep its last
/// word, in spaces of the word's font. Measured on three compat-14 Courier
/// New statute documents: 58 squeezed lines all overflow by two characters
/// or less (letters and spaces narrowed alike, to 0.965-0.985), none by
/// three; the cap sits between the two. ponytail: a few lines one character
/// over still wrap in Word, which no width cap explains; find that rule if
/// a proportional-font document turns up.
const SQUEEZE_LINE_SPACES: f32 = 2.5;

/// Per chunk, how many word spaces precede it on the line. Justification
/// stretches or squeezes those, not the joins between runs inside a word: a
/// space is a gap between chunks or a chunk that is itself a space (spaces
/// carrying underline or shading are text-less chunks of their own).
fn spaces_before_each(chunks: &[WordChunk]) -> Vec<usize> {
    chunks
        .iter()
        .scan((0usize, None::<(f32, bool)>), |(n, prev), c| {
            if prev.is_some_and(|(end, was_space)| was_space || c.x_offset > end + 0.01) {
                *n += c.preceding_spaces;
            }
            let is_space = c.inline_image_name.is_none() && c.text.chars().all(is_break_space);
            *prev = Some((c.x_offset + c.width, is_space));
            Some(*n)
        })
        .collect()
}

/// Total inter-word space already on a line: the gaps between its chunks.
fn line_space_width(chunks: &[WordChunk]) -> f32 {
    let gaps: f32 = chunks
        .windows(2)
        .map(|w| (w[1].x_offset - (w[0].x_offset + w[0].width)).max(0.0))
        .sum();
    gaps + chunks
        .iter()
        .filter(|chunk| chunk.word_space_bridge)
        .map(|chunk| chunk.width)
        .sum::<f32>()
}

/// Full-width East Asian closing punctuation whose right half is blank, which
/// Word's `compressPunctuation` may squeeze (§17.15.1.15). The middle dot,
/// blank on both sides, squeezes like them.
fn is_compressible_punct(c: char) -> bool {
    matches!(
        c,
        '、' | '。'
            | '，'
            | '．'
            | '：'
            | '；'
            | '！'
            | '？'
            | '）'
            | '］'
            | '｝'
            | '」'
            | '』'
            | '】'
            | '〕'
            | '〉'
            | '》'
            | '〙'
            | '〗'
            | '・'
    )
}

/// Full-width opening brackets, blank on their left, which `compressPunctuation`
/// squeezes from that side: japanese_medical's "（" sits 2pt into the gap
/// before it on a tight line.
fn is_compressible_opening(c: char) -> bool {
    matches!(
        c,
        '（' | '［' | '｛' | '「' | '『' | '【' | '〔' | '〈' | '《' | '〘' | '〖'
    )
}

/// Hinted curly quotes squeeze like the brackets: ‘“ from the left, ’” from
/// the right.
fn hinted_closing(c: char) -> bool {
    matches!(c, '’' | '”')
}
fn hinted_opening(c: char) -> bool {
    matches!(c, '‘' | '“')
}

/// How much of its em a full-width mark gives up to `compressPunctuation` at
/// most under the Word 2013+ layout: half, whatever its glyph (a KaiTi
/// comma's blank right half is 0.7 em and never comes into it). Every mark on
/// the line gives the same amount. Word online probes, 2026-10-26 (issue
/// #139): 12pt KaiTi lines with 2 to 8 commas or curly quotes before a Latin
/// word, 10.5 to 22pt, word widths of 1 to 4.5 em; 3 marks take 6.0pt each
/// and refuse 6.65pt. The 2010 layout (compatibilityMode 14) squeezes any
/// alignment but stops at 0.4 em (16pt YaHei takes 6.26pt and refuses 7.26).
pub(super) const PUNCT_SQUEEZE: f32 = 0.5;

/// Word 2013+ pulls a word up by squeezing marks only when the line before it
/// still has this share of the measure free (same probes: 12pt lines of 20
/// to 40 glyphs, measures of 240 to 510pt, pull at 1.22% free and wrap at
/// 1.19% whatever the word's width). A nearly full line wraps the word
/// instead, however little the marks would have to give. The 2010 layout
/// asks for half an em instead, whatever the measure, and counts the word's
/// autoSpaceDE gap as used (12pt pulls "SN/T" at 6.7pt free past the gap and
/// wraps it at 5.8 on 250 and 490pt lines alike; taiwanese_education_fraud_
/// ruling pulls 業 at 8.1pt free, 16pt).
pub(super) const PUNCT_PULL_FREE: f32 = 0.012;

/// Recover `needed` points on the current line by trimming the chunks that end
/// in compressible punctuation, the same amount off every mark (counting what
/// earlier words on the line already took) and at most `cap` of the em.
/// Returns false and touches nothing when that is not enough.
fn compress_punctuation(chunks: &mut [WordChunk], needed: f32, cap: f32) -> bool {
    // (chunk, squeezes at its end, squeezes at its start)
    let marks: Vec<(usize, f32, bool)> = chunks
        .iter()
        .enumerate()
        .filter(|(_, c)| c.inline_image_name.is_none())
        .map(|(i, c)| {
            let end = c.text.chars().last().is_some_and(|ch| {
                is_compressible_punct(ch) || (c.east_asia_hint && hinted_closing(ch))
            });
            let start = c.text.chars().next().is_some_and(|ch| {
                is_compressible_opening(ch) || (c.east_asia_hint && hinted_opening(ch))
            });
            (i, end as u8 as f32 + start as u8 as f32, start)
        })
        .filter(|&(_, sides, _)| sides > 0.0)
        .collect();
    let sides_total: f32 = marks.iter().map(|&(_, sides, _)| sides).sum();
    if needed <= 0.0 || sides_total == 0.0 {
        return false;
    }
    let taken: f32 = marks
        .iter()
        .map(|&(i, _, _)| chunks[i].punct_compressed)
        .sum();
    let level = (taken + needed) / sides_total;
    if marks
        .iter()
        .any(|&(i, _, _)| level > chunks[i].font_size * cap + 0.01)
    {
        return false;
    }
    let mut shift = 0.0;
    let mut k = 0;
    for (i, chunk) in chunks.iter_mut().enumerate() {
        chunk.x_offset -= shift;
        if k < marks.len() && marks[k].0 == i {
            let (_, sides, start) = marks[k];
            let cut = (level * sides - chunk.punct_compressed).max(0.0);
            // An opening bracket loses its blank left: the chunk slides into
            // the gap before it by its share of the cut, and only the rest
            // comes off its right end.
            let lead = if start { cut / sides } else { 0.0 };
            chunk.x_offset -= lead;
            chunk.width -= cut - lead;
            chunk.punct_compressed += cut;
            shift += cut;
            k += 1;
        }
    }
    true
}

/// Dual-region geometry for bothSides wrapping: (left_x, left_w, right_x, right_w).
/// For lines outside the float zone, right_w is 0.0 (single region).
pub(super) type DualRegion = (f32, f32, f32, f32);

/// What line breaking takes besides the runs and the measure of a paragraph.
/// The default is a plain paragraph: no pictures, tab stops, indents or floats.
#[derive(Default, Clone, Copy)]
pub(super) struct LineOpts<'a> {
    /// Inline pictures and their effect XObjects, by run index.
    pub(super) inline_images: Option<&'a HashMap<usize, String>>,
    pub(super) effects: Option<&'a HashMap<usize, EffectXObjs>>,
    pub(super) tab_stops: &'a [TabStop],
    pub(super) indent_left: f32,
    pub(super) indent_right: f32,
    pub(super) hanging: f32,
    /// Spans a left tab skips (see `build_tabbed_line`).
    pub(super) tab_exclusions: &'a [(f32, f32)],
    /// The measure of each line beside a float (see `build_paragraph_lines`).
    pub(super) per_line_widths: Option<&'a [f32]>,
    pub(super) dual: Option<&'a [DualRegion]>,
}

/// Lay a paragraph out as lines: against its tab stops when it holds a tab,
/// by plain word wrapping otherwise.
pub(super) fn build_lines(
    runs: &[Run],
    ctx: &RenderContext,
    width: f32,
    cjk: CjkLayout,
    opts: &LineOpts<'_>,
) -> Vec<TextLine> {
    let inline_images = opts.inline_images.unwrap_or(&EMPTY_INLINE_IMAGES);
    let effects = opts.effects.unwrap_or(&EMPTY_EFFECTS);
    if runs.iter().any(|r| r.is_tab) {
        build_tabbed_line(
            runs,
            ctx.fonts,
            opts.tab_stops,
            opts.indent_left,
            width,
            opts.indent_right,
            opts.hanging,
            inline_images,
            effects,
            ctx.default_tab_stop,
            opts.tab_exclusions,
            ctx.compat_mode,
            cjk,
        )
    } else {
        build_paragraph_lines(
            runs,
            ctx.fonts,
            width,
            opts.hanging,
            inline_images,
            effects,
            None,
            opts.per_line_widths,
            opts.dual,
            cjk,
        )
    }
}

/// Layout runs into wrapped lines.
/// Handles cross-run contiguous text correctly: no space is inserted between
/// runs unless the preceding text ended with whitespace or the new run starts
/// with whitespace (e.g., "bold" + ", " → "bold," not "bold ,").
/// `width_after_line`: after building this many lines, switch to a different
/// max_width (used for text wrapping around floating tables where lines beside
/// the table are narrow, then lines below it expand to full column width).
pub(super) fn build_paragraph_lines(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    max_width: f32,
    first_line_hanging: f32,
    inline_image_names: &HashMap<usize, String>,
    effect_inline_names: &HashMap<usize, super::images::EffectXObjs>,
    width_after_line: Option<(usize, f32)>,
    per_line_widths: Option<&[f32]>,
    per_line_dual: Option<&[DualRegion]>,
    cjk: CjkLayout,
) -> Vec<TextLine> {
    let mut lines: Vec<TextLine> = Vec::new();
    // Indices of lines that kept a word by squeezing their spaces.
    let mut squeezed_lines: Vec<usize> = Vec::new();
    let mut current_chunks: Vec<WordChunk> = Vec::new();
    // A display equation broke over lines (see the group alignment below).
    let mut display_wrapped = false;
    let mut current_x: f32 = 0.0;
    let mut pending_space_w: f32 = 0.0;
    let mut pending_space_count: usize = 0;
    // pending_space_w minus CJK auto-spacing: a space character really occurred.
    let mut pending_real_space = false;
    // Underline state of the run that emitted the pending whitespace. A space
    // is underlined when its *own* run is underlined (Word draws underline
    // continuously across spaces inside an underlined run, but not across a
    // space that belongs to a neighbouring non-underlined run).
    let mut pending_space_underline = false;
    // Whether the pending spaces came from a `w:rtl` run (`WordChunk::gap_rtl`).
    let mut space_rtl = false;
    let mut pending_space_double = false;
    let mut pending_space_color: Option<[u8; 3]> = None;
    let mut pending_space_border: Option<ParagraphBorder> = None;
    let mut key_buf = String::new();
    // Track whether we're filling the right region on the current line
    let mut in_right_region = false;
    // Info for the right region of the current line being built
    let mut cur_right_info: Option<(usize, f32, f32)> = None; // (first_chunk_idx, region_x, region_w)
    // Track last non-space character for CJK auto-spacing (autoSpaceDE/DN)
    let mut prev_last_char: Option<char> = None;
    // Index in `current_chunks` where the word being placed began; a word can
    // span several runs.
    let mut word_start = 0usize;

    // Lines beside a float that cannot hold the first word are left out: the
    // paragraph's first line is the first that can (`float_skip`).
    let skip = std::cell::Cell::new(0usize);
    let left_max = |line_count: usize| -> f32 {
        let line_count = line_count + skip.get();
        if let Some(dual) = per_line_dual
            && let Some(&(_, lw, _, _)) = dual.get(line_count)
        {
            return lw;
        }
        if let Some(widths) = per_line_widths {
            if let Some(&w) = widths.get(line_count) {
                return w;
            }
            return max_width;
        }
        match width_after_line {
            Some((n, w)) if line_count >= n => w,
            _ => max_width,
        }
    };

    let right_region_for = |line_count: usize| -> Option<(f32, f32, f32)> {
        per_line_dual.and_then(|dual| {
            dual.get(line_count + skip.get()).and_then(|&(_, _, rx, rw)| {
                // The first line gets its first-line/hanging shift like the
                // left region does (`eff_margin` in render_paragraph_lines).
                let shift = if line_count == 0 {
                    first_line_hanging
                } else {
                    0.0
                };
                let rw = rw + shift;
                (rw > 0.0).then_some((rx - shift, rw, rw))
            })
        })
    };

    // Finish current line, recording right-region info if applicable
    let finish_dual_line = |chunks: &mut Vec<WordChunk>,
                            in_right: &mut bool,
                            right_info: &mut Option<(usize, f32, f32)>|
     -> TextLine {
        let mut line = finish_line(chunks);
        if let Some((first_idx, rx, rw)) = right_info.take() {
            let content_w = line.chunks[first_idx..]
                .last()
                .map(|c| c.x_offset + c.width)
                .unwrap_or(0.0);
            line.right_region = Some(RightRegion {
                first_chunk_idx: first_idx,
                region_x: rx,
                region_width: rw,
                content_width: content_w,
            });
        }
        *in_right = false;
        line
    };

    for (run_idx, run) in runs.iter().enumerate() {
        if run.vanish {
            continue;
        }
        // Tabs and breaks move the pen without a glyph; extraction still needs
        // the word boundary, as Word's text has one there.
        if run.is_tab {
            pending_real_space = true;
            continue;
        }

        if run.is_line_break {
            mark_space_after(&mut current_chunks);
            let holds_only_break = current_chunks
                .iter()
                .all(|c| c.text.trim().is_empty() && c.inline_image_name.is_none());
            let line = finish_dual_line(
                &mut current_chunks,
                &mut in_right_region,
                &mut cur_right_info,
            );
            let line = TextLine {
                ends_with_break: true,
                break_font_size: holds_only_break.then_some(run.font_size),
                break_lhr: holds_only_break
                    .then(|| seen_fonts.get(&crate::fonts::font_key(run)))
                    .flatten()
                    .and_then(|e| e.line_h_ratio),
                ..line
            };
            lines.push(line);
            current_x = 0.0;
            pending_space_w = 0.0;
            pending_space_count = 0;
            pending_real_space = false;
            prev_last_char = None;
            continue;
        }

        // Handle inline images as single block elements in the line
        if let Some(img) = &run.inline_image {
            if std::mem::take(&mut pending_real_space) {
                mark_space_after(&mut current_chunks);
            }
            if let Some(pdf_name) = inline_image_names.get(&run_idx) {
                let img_w = img.layout_size().0;
                // Leading spaces indent a picture as they do a word
                // (russian_chess: 19 spaces push the board 57pt right).
                let proposed_x = current_x + pending_space_w;

                let eff_w = left_max(lines.len());
                let line_max = if lines.is_empty() && !in_right_region {
                    eff_w + first_line_hanging
                } else {
                    eff_w
                };
                let cur_max = if in_right_region {
                    cur_right_info.map(|(_, _, rw)| rw).unwrap_or(eff_w)
                } else {
                    line_max
                };
                if (!current_chunks.is_empty() || pending_space_w > 0.0)
                    && proposed_x + img_w > cur_max
                {
                    if !in_right_region && let Some((rx, rw, _)) = right_region_for(lines.len()) {
                        cur_right_info = Some((current_chunks.len(), rx, rw));
                        in_right_region = true;
                        pending_space_w = 0.0;
                        pending_space_count = 0;
                        // Retry placement in right region
                        let proposed_x2 = 0.0;
                        if proposed_x2 + img_w <= rw {
                            current_chunks.push(WordChunk::image(
                                pdf_name,
                                run.font_size,
                                proposed_x2,
                                img,
                                effect_inline_names.get(&run_idx).cloned(),
                            ));
                            current_x = img_w;
                            continue;
                        }
                    }
                    let leading_space_run = current_chunks
                        .is_empty()
                        .then(|| runs[..run_idx].iter().rev().find(|r| !r.text.is_empty()))
                        .flatten();
                    let mut line = finish_dual_line(
                        &mut current_chunks,
                        &mut in_right_region,
                        &mut cur_right_info,
                    );
                    if let Some(space_run) = leading_space_run {
                        let entry = seen_fonts
                            .get(font_key_buf(space_run, &mut key_buf))
                            .expect("font registered");
                        line.break_font_size = Some(effective_font_size(space_run, entry));
                        line.break_lhr = entry.line_h_ratio;
                        line.leading_space_wrap = true;
                    }
                    lines.push(line);
                    current_x = 0.0;
                } else {
                    current_x = proposed_x;
                }
                pending_space_w = 0.0;
                pending_space_count = 0;

                current_chunks.push(WordChunk::image(
                    pdf_name,
                    run.font_size,
                    current_x,
                    img,
                    effect_inline_names.get(&run_idx).cloned(),
                ));
                current_x += img_w;
                word_start = current_chunks.len();
                // A picture is an object character: a line may break after
                // it (UAX #14 LB20), so the text glued to it can wrap (Word
                // moves "Step" below a full-width banner picture).
                prev_last_char = Some('\u{FFFC}');
            }
            continue;
        }

        if let Some((entry, mb)) = math_box(run, seen_fonts) {
            // A math zone is one box, glued to text it touches without a
            // space: Word wraps pendulum's "cambiando la lunghezza 𝑙 del
            // filo" before the zone "𝑙 " since nothing may break before "del".
            let glued = pending_space_w == 0.0
                && !current_chunks.is_empty()
                && !prev_last_char.is_some_and(|a| breaks_between(a, 'x'));
            let mut start = if glued {
                word_start
            } else {
                current_chunks.len()
            };
            let proposed_x = current_x + pending_space_w;
            let cur_max = if in_right_region {
                cur_right_info
                    .map(|(_, _, rw)| rw)
                    .unwrap_or(left_max(lines.len()))
            } else if lines.is_empty() {
                left_max(0) + first_line_hanging
            } else {
                left_max(lines.len())
            };
            current_x = proposed_x;
            // A zone too long for the line breaks after one of its operators,
            // as many times as it takes; what is left places like a whole zone.
            let mut mb = mb.clone();
            let free = |lines: &[TextLine], x: f32| {
                if lines.is_empty() {
                    left_max(0) + first_line_hanging - x
                } else {
                    left_max(lines.len()) - x
                }
            };
            if !in_right_region && right_region_for(lines.len()).is_none() {
                while let Some((head, tail)) = Some(free(&lines, current_x) + 0.05)
                    .filter(|&room| mb.width > room)
                    .and_then(|room| mb.fit(room))
                {
                    if pending_space_w > 0.0 {
                        mark_space_after(&mut current_chunks);
                        pending_space_w = 0.0;
                    }
                    current_chunks.push(WordChunk::math(entry, run, &Arc::new(head), current_x));
                    let line = finish_dual_line(
                        &mut current_chunks,
                        &mut in_right_region,
                        &mut cur_right_info,
                    );
                    lines.push(line);
                    // A display equation's continuation lines start at
                    // m:wrapIndent (72pt by default: probe `mp17_wrap`).
                    (current_x, start) = (if mb.display { mb.wrap_indent } else { 0.0 }, 0);
                    display_wrapped |= mb.display;
                    mb = Arc::new(tail);
                }
            }
            let mb = &mb;
            if start > 0
                && current_x + mb.width > cur_max + 0.05
                && !in_right_region
                && right_region_for(lines.len()).is_none()
            {
                let carried: Vec<WordChunk> = current_chunks.drain(start..).collect();
                let dx = carried.first().map_or(current_x, |c| c.x_offset);
                let line = finish_dual_line(
                    &mut current_chunks,
                    &mut in_right_region,
                    &mut cur_right_info,
                );
                lines.push(line);
                current_chunks.extend(carried.into_iter().map(|mut c| {
                    c.x_offset -= dx;
                    c
                }));
                current_x -= dx;
                start = 0;
            }
            if pending_space_w > 0.0 {
                mark_space_after(&mut current_chunks);
            }
            pending_space_w = 0.0;
            current_chunks.push(WordChunk::math(entry, run, mb, current_x));
            current_x += mb.width;
            word_start = start;
            // Text right after the zone continues its word.
            prev_last_char = Some('x');
            continue;
        }

        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key).expect("font registered");
        let eff_fs = effective_font_size(run, entry);
        let text = &run.text;
        let y_off = vert_y_offset(run);

        let cs = run.char_spacing;
        let ts = run.text_scale / 100.0;

        let mut is_first_word_in_run = true;
        let mut prev_char = None;
        let mut words: std::collections::VecDeque<_> = split_preserving_spaces(text).into();
        while let Some((space_count, mut source)) = words.pop_front() {
            pending_space_count += space_count;
            let literal_spaces = pending_space_count;
            let mut shown = caps_word(run, source);
            pending_space_w +=
                space_count as f32 * word_space_width(run, entry, eff_fs, source, &mut prev_char);
            if space_count > 0 {
                pending_real_space = true;
                space_rtl = run.rtl;
                pending_space_underline = run.underline;
                pending_space_double = run.double_underline;
                pending_space_color = run.color;
                pending_space_border = run.border.clone();
            }
            // current_chunks still holds what precedes this word on its line,
            // so a wrap below keeps the space at the end of the finished line.
            if std::mem::take(&mut pending_real_space) {
                mark_space_after(&mut current_chunks);
            }

            // CJK auto-spacing (autoSpaceDE/DN): add ~0.25em gap at
            // script boundaries between East Asian and Latin/digit text
            // when there is no explicit whitespace.
            if cjk.auto_space
                && let Some(prev_ch) = prev_last_char
                && let Some(first_ch) = shown.chars().next()
                && pending_space_w == 0.0
                && space_count == 0
            {
                // Only a letter or digit meets an ideograph: Word sets no gap
                // after japanese_interlibrary's "〒" before "2", nor between
                // chinese_costume's ")" and "格". An ideographic space is a
                // space: japanese_medical's "　kg　" has no gap around "kg".
                let ideograph =
                    |c: char| crate::docx::is_east_asian_char(c) && !is_cjk_punctuation(c);
                let latin = |c: char| c.is_alphanumeric() && !crate::docx::is_east_asian_char(c);
                if (ideograph(prev_ch) && latin(first_ch)) || (latin(prev_ch) && ideograph(first_ch)) {
                    pending_space_w += eff_fs * 0.25;
                }
            }

            // A word continues the previous word across a run boundary when
            // there is no whitespace between runs, no leading spaces, and the
            // two characters could not break inside one run either.
            let nbsp_bridge = pending_space_w > 0.0
                && (prev_last_char == Some('\u{00a0}') || shown.starts_with('\u{00a0}'));
            let is_continuation = !current_chunks.is_empty()
                && (nbsp_bridge
                    || (is_first_word_in_run
                        && space_count == 0
                        && pending_space_w == 0.0
                        && !current_chunks.is_empty()
                        && !prev_last_char
                            .zip(shown.chars().next())
                            .is_some_and(|(a, b)| breaks_between(a, b))));
            is_first_word_in_run = false;

            let kern = run.kerns_at(eff_fs);
            let width = |w: &str| word_width_for_run(entry, run, w, eff_fs, kern, cs, ts);
            let mut ww = width(&shown);
            // A word wider than a whole line breaks at the margin, character
            // by character (a 280-dot leader fills two lines).
            let line_room = left_max(lines.len())
                + if lines.is_empty() {
                    first_line_hanging
                } else {
                    0.0
                };
            // A strip left of a wrapping float is no line: a word that fits
            // the strip right of it moves there whole (Word puts "Extended"
            // past a logo, not "Extend" before it).
            let fits_right = right_region_for(lines.len()).is_some_and(|(_, rw, _)| ww <= rw);
            // A word too wide for both strips beside a float leaves the line
            // empty and goes on to the next; before any line, the paragraph
            // itself starts lower (greensboro's 143 no-break spaces take four
            // lines beside a logo, Word probe 2026-10-23).
            if current_chunks.is_empty()
                && !in_right_region
                && !is_continuation
                && ww > line_room
                && !fits_right
                && right_region_for(lines.len()).is_some()
            {
                if lines.is_empty() {
                    skip.set(skip.get() + 1);
                } else {
                    lines.push(TextLine {
                        float_skip: true,
                        ..finish_dual_line(
                            &mut current_chunks,
                            &mut in_right_region,
                            &mut cur_right_info,
                        )
                    });
                }
                pending_space_w = 0.0;
                pending_space_count = 0;
                words.push_front((0, source));
                continue;
            }
            if current_chunks.is_empty() && !in_right_region && ww > line_room && !fits_right {
                let room = line_room - pending_space_w;
                // Cut the letters as written, not the capitals drawn (ß → SS
                // changes length), so each half keeps its own /ActualText.
                if let Some(cut) = fitting_prefix_len(source, room, |w| width(&caps_word(run, w))) {
                    words.push_front((0, &source[cut..]));
                    source = &source[..cut];
                    shown = caps_word(run, source);
                    ww = width(&shown);
                }
            }
            let mut word: &str = &shown;
            let mut original = run.caps.then_some(source);
            prev_last_char = word.chars().last();

            // Japanese: a closing mark directly followed by another full-width
            // mark keeps only its ink half, whether the line is tight or not
            // (japanese_medical's "院）・専門" draws the ） 5.25pt wide in
            // 10.5pt text on a line 13pt short of the measure: JIS X 4051's
            // adjacent-mark squeeze).
            let japanese = cjk.compress_punct && run_is_japanese(run);
            if japanese {
                ww -= ja_adjacent_marks(word) as f32 * 0.5 * eff_fs;
            }
            if japanese
                && pending_space_w == 0.0
                && word.chars().next().is_some_and(is_fullwidth_mark)
                && let Some(last) = current_chunks.last_mut()
                && last.punct_compressed == 0.0
                && last.text.chars().last().is_some_and(is_compressible_punct)
            {
                let cut = last.font_size * 0.5;
                last.width -= cut;
                last.punct_compressed += cut;
                current_x -= cut;
            }

            let need_space = !current_chunks.is_empty() && pending_space_w > 0.0;

            let proposed_x = if need_space {
                current_x + pending_space_w
            } else if current_chunks.is_empty() && pending_space_w > 0.0 {
                // Leading spaces at the start of a line act as a visual indent
                pending_space_w
            } else {
                current_x
            };

            let eff_w = left_max(lines.len());
            let line_max = if lines.is_empty() && !in_right_region {
                eff_w + first_line_hanging
            } else {
                eff_w
            };
            let cur_max = if in_right_region {
                cur_right_info.map(|(_, _, rw)| rw).unwrap_or(eff_w)
            } else {
                line_max
            };

            // Small tolerance for floating-point width accumulation over many glyphs
            let mut overflows = proposed_x + ww > cur_max + 0.05;
            // The segment after a hyphen break may stay by narrowing the
            // line's spaces, judged like a word after a space: Word 2013+ keeps
            // "x", "12-", "based", "and/or" and "12345/2026" alike, within the
            // same midpoint and budget (Word probe 2026-10-26).
            let after_hyphen = current_chunks.last().is_some_and(|c| c.text.ends_with('-'));
            // A word glued to the previous run's word is judged as one word.
            let (word_x, word_w) = if is_continuation && word_start < current_chunks.len() {
                let x = current_chunks[word_start].x_offset;
                (x, proposed_x + ww - x)
            } else {
                (proposed_x, ww)
            };
            if overflows && (need_space || is_continuation || after_hyphen) && !in_right_region {
                // How far the word may overflow and stay: Word 2013+ narrows
                // the spaces, wpJustification the whole line.
                let lead_space = if need_space { pending_space_w } else { 0.0 };
                let half_mark = if cjk.punct_half_out {
                    word.chars()
                        .last()
                        .filter(|c| ".,;:!?".contains(*c))
                        .map_or(0.0, |c| entry.char_width_1000(c) * eff_fs / 2000.0)
                } else {
                    0.0
                };
                let keeps = squeeze_keeps(word_x + word_w - cur_max - half_mark, word_w, lead_space);
                let space_room = (cjk.squeeze_spaces && keeps)
                    .then(|| SPACE_SQUEEZE * (line_space_width(&current_chunks) + pending_space_w));
                let line_room = (cjk.squeeze_line && (need_space || is_continuation))
                    .then(|| SQUEEZE_LINE_SPACES * entry.space_width(eff_fs));
                if [space_room, line_room]
                    .into_iter()
                    .flatten()
                    .any(|room| proposed_x + ww - cur_max <= room)
                {
                    overflows = false;
                    squeezed_lines.push(lines.len());
                }
            }

            // compressPunctuation: before wrapping, make room for this word by
            // squeezing the full-width punctuation already on the line.
            let (proposed_x, overflows) = if overflows
                && cjk.compress_punct
                && cjk.punct_pull
                && !current_chunks.is_empty()
                && !in_right_region
                && right_region_for(lines.len()).is_none()
                && cur_max - current_x >= cjk.punct_pull_free_ratio * cur_max
                && cur_max - proposed_x >= cjk.punct_pull_free_em * eff_fs
                && compress_punctuation(
                    &mut current_chunks,
                    proposed_x + ww - cur_max,
                    cjk.punct_cap,
                ) {
                current_x = current_chunks.last().map_or(0.0, |c| c.x_offset + c.width);
                (
                    if need_space {
                        current_x + pending_space_w
                    } else {
                        current_x
                    },
                    false,
                )
            } else {
                (proposed_x, overflows)
            };

            // When leading spaces push the first word past the line width,
            // emit a blank line for the spaces and start the word at x=0.
            // Word absorbs the spaces onto the current line and wraps the
            // word to the next line. (Skipped when a right region exists —
            // the spill below carries the space indent past the float.)
            if current_chunks.is_empty()
                && overflows
                && pending_space_w > 0.0
                && !in_right_region
                && right_region_for(lines.len()).is_none()
            {
                lines.push(finish_dual_line(
                    &mut current_chunks,
                    &mut in_right_region,
                    &mut cur_right_info,
                ));
                pending_space_w = 0.0;
                pending_space_count = 0;
                word_start = 0;
                push_word_chunks(
                    &mut current_chunks,
                    entry,
                    run,
                    word,
                    original,
                    eff_fs,
                    cs,
                    y_off,
                    0.0,
                    ww,
                    space_rtl,
                    cjk.compress_punct,
                );
                current_x = ww;
                continue;
            }

            // A word split over several runs wraps as a whole: carry the part
            // already placed to the next line. One wider than that line breaks
            // at its margin, character by character, like a single-run word.
            if is_continuation
                && overflows
                && word_start < current_chunks.len()
                && !in_right_region
                && right_region_for(lines.len()).is_none()
            {
                let dx = current_chunks[word_start].x_offset;
                if word_start > 0 {
                    let carried: Vec<WordChunk> = current_chunks.drain(word_start..).collect();
                    lines.push(finish_dual_line(
                        &mut current_chunks,
                        &mut in_right_region,
                        &mut cur_right_info,
                    ));
                    current_chunks.extend(carried.into_iter().map(|mut c| {
                        c.x_offset -= dx;
                        c
                    }));
                    current_x -= dx;
                    word_start = 0;
                }
                if pending_space_underline && pending_space_w > 0.0 {
                    current_chunks.push(WordChunk::word_space_underline(
                        entry,
                        eff_fs,
                        pending_space_color,
                        pending_space_double,
                        pending_space_border.clone(),
                        current_x,
                        pending_space_w,
                    ));
                }
                current_x += pending_space_w;
                pending_space_w = 0.0;
                pending_space_count = 0;
                let room = left_max(lines.len()) - current_x;
                if ww > room + 0.05 {
                    let cut = fitting_prefix_len(source, room, |w| width(&caps_word(run, w)))
                        .filter(|&c| width(&caps_word(run, &source[..c])) <= room + 0.05);
                    if let Some(cut) = cut {
                        words.push_front((0, &source[cut..]));
                        source = &source[..cut];
                        shown = caps_word(run, source);
                        word = &shown;
                        original = run.caps.then_some(source);
                        ww = width(word);
                        prev_last_char = word.chars().last();
                    } else {
                        // Not one more character fits: the rest goes on as the
                        // next word and wraps there.
                        words.push_front((0, source));
                        continue;
                    }
                }
                let chunk_start = current_chunks.len();
                push_word_chunks(
                    &mut current_chunks,
                    entry,
                    run,
                    word,
                    original,
                    eff_fs,
                    cs,
                    y_off,
                    current_x,
                    ww,
                    space_rtl,
                    cjk.compress_punct,
                );
                if literal_spaces > 1 && !pending_space_underline && chunk_start > 0 {
                    current_chunks[chunk_start].preceding_spaces = literal_spaces;
                }
                current_x += ww;
                continue;
            }

            // For the first word on a line, also overflow if the
            // left region is zero-width (so words go straight to
            // the right region for e.g. left-aligned images).
            let first_word_overflow = current_chunks.is_empty()
                && !in_right_region
                && (cur_max <= 0.0 || overflows)
                && right_region_for(lines.len()).is_some();

            if (!current_chunks.is_empty() && overflows && !is_continuation) || first_word_overflow
            {
                // Word doesn't fit in current region
                if !in_right_region {
                    // Try spilling to the right region on the same line
                    if let Some((rx, rw, _)) = right_region_for(lines.len()) {
                        cur_right_info = Some((current_chunks.len(), rx, rw));
                        in_right_region = true;
                        // Leading spaces on a line that starts in the right
                        // region indent the first word (Word keeps them).
                        let start_x = if current_chunks.is_empty()
                            && pending_space_w > 0.0
                            && pending_space_w + ww <= rw
                        {
                            pending_space_w
                        } else {
                            0.0
                        };
                        pending_space_w = 0.0;
                        pending_space_count = 0;
                        word_start = current_chunks.len();
                        push_word_chunks(
                            &mut current_chunks,
                            entry,
                            run,
                            word,
                            original,
                            eff_fs,
                            cs,
                            y_off,
                            start_x,
                            ww,
                            space_rtl,
                            cjk.compress_punct,
                        );
                        current_x = start_x + ww;
                        continue;
                    }
                }
                // Try hyphenation before wrapping whole word
                // No right region or right region also full — wrap to next line
                lines.push(TextLine {
                    trailing_space: pending_space_w,
                    trailing_space_rtl: space_rtl,
                    ..finish_dual_line(
                        &mut current_chunks,
                        &mut in_right_region,
                        &mut cur_right_info,
                    )
                });
                current_x = 0.0;
                // A word wider than the new line breaks at its margin too.
                let room = left_max(lines.len());
                let right = right_region_for(lines.len());
                if ww > room
                    && right.is_none()
                    && let Some(cut) =
                        fitting_prefix_len(source, room, |w| width(&caps_word(run, w)))
                {
                    words.push_front((0, &source[cut..]));
                    source = &source[..cut];
                    shown = caps_word(run, source);
                    word = &shown;
                    original = run.caps.then_some(source);
                    ww = width(word);
                    prev_last_char = word.chars().last();
                }
                // If the new line's left region is zero-width, go
                // straight to the right region for this word.
                if let Some((rx, rw, _)) = right {
                    // So does a word the left region is too narrow for: Word
                    // leaves that gap empty rather than split the word around
                    // the float (case42's "ullamcorper." beside the arm).
                    // One too wide for the right region as well leaves the whole
                    // line empty and tries the next one, down to the first line
                    // clear of the float, where it breaks at the margin
                    // (greensboro's 143 no-break spaces take 4 lines, Word probe
                    // 2026-10-23).
                    if (room <= 0.0 || ww > room) && ww > rw + 0.05 {
                        lines.push(TextLine {
                            float_skip: true,
                            ..finish_dual_line(
                                &mut current_chunks,
                                &mut in_right_region,
                                &mut cur_right_info,
                            )
                        });
                        words.push_front((0, source));
                        continue;
                    }
                    if room <= 0.0 || ww > room {
                        cur_right_info = Some((0, rx, rw));
                        in_right_region = true;
                        pending_space_w = 0.0;
                        pending_space_count = 0;
                        word_start = 0;
                        push_word_chunks(
                            &mut current_chunks,
                            entry,
                            run,
                            word,
                            original,
                            eff_fs,
                            cs,
                            y_off,
                            0.0,
                            ww,
                            space_rtl,
                            cjk.compress_punct,
                        );
                        current_x = ww;
                        continue;
                    }
                }
            } else {
                current_x = proposed_x;
            }
            // Bridge the underline across the space preceding this word when the
            // space's own run is underlined, so a continuously-underlined run
            // renders as one unbroken line rather than per-word segments.
            if pending_space_underline && !current_chunks.is_empty() && pending_space_w > 0.0 {
                current_chunks.push(WordChunk::word_space_underline(
                    entry,
                    eff_fs,
                    pending_space_color,
                    pending_space_double,
                    pending_space_border.clone(),
                    current_x - pending_space_w,
                    pending_space_w,
                ));
            }
            pending_space_w = 0.0;
            pending_space_count = 0;

            if !is_continuation {
                word_start = current_chunks.len();
            }
            let chunk_start = current_chunks.len();
            push_word_chunks(
                &mut current_chunks,
                entry,
                run,
                word,
                original,
                eff_fs,
                cs,
                y_off,
                current_x,
                ww,
                space_rtl,
                cjk.compress_punct,
            );
            if literal_spaces > 1
                && !pending_space_underline
                && chunk_start > 0
                && let Some(c) = current_chunks.get_mut(chunk_start)
            {
                c.preceding_spaces = literal_spaces;
            }
            current_x += ww;
        }

        // Accumulate trailing whitespace for the next run
        let trailing_spaces = text
            .chars()
            .rev()
            .take_while(|c| is_break_space(*c))
            .count();
        if trailing_spaces > 0 {
            pending_space_count += trailing_spaces;
            pending_space_w +=
                trailing_spaces as f32 * word_space_width(run, entry, eff_fs, "", &mut prev_char);
            pending_real_space = true;
            space_rtl = run.rtl;
            pending_space_underline = run.underline;
            pending_space_double = run.double_underline;
            pending_space_color = run.color;
            pending_space_border = run.border.clone();
        }
    }

    if !current_chunks.is_empty() {
        lines.push(finish_dual_line(
            &mut current_chunks,
            &mut in_right_region,
            &mut cur_right_info,
        ));
    }

    // A trailing line break creates an empty line after it (Word adds a blank
    // line for each w:br at the end of a paragraph).  Store the break run's
    // font size so the caller can compute the correct height for this line.
    if lines.last().is_some_and(|l| l.ends_with_break) {
        let break_fs = runs
            .iter()
            .rev()
            .find(|r| r.is_line_break)
            .map(|r| r.font_size);
        lines.push(TextLine {
            break_font_size: break_fs,
            ..TextLine::default()
        });
    }

    if lines.is_empty() {
        lines.push(TextLine::default());
    }
    // A display equation broken over lines is aligned as a group: each line
    // as wide as the widest, so centring puts every line's start (the
    // continuation lines' indent included) on one edge: Word starts a
    // full-width first line 1.7pt in, the rest 72pt further (`mp18_after`).
    if display_wrapped {
        let display = |l: &TextLine| {
            l.chunks
                .iter()
                .any(|c| c.math.as_ref().is_some_and(|m| m.display))
        };
        let group_w = lines
            .iter()
            .filter(|l| display(l))
            .fold(0.0f32, |w, l| w.max(l.total_width));
        for line in lines.iter_mut().filter(|l| display(l)) {
            line.total_width = group_w;
        }
    }
    for i in squeezed_lines {
        if let Some(line) = lines.get_mut(i) {
            line.squeezed = true;
            line.squeeze_chars = cjk.squeeze_line;
        }
    }
    if !cjk.expand_shift_return {
        for line in lines.iter_mut().filter(|l| l.ends_with_break) {
            line.natural_width = true;
        }
    }
    if skip.get() > 0 {
        let mut skipped: Vec<TextLine> = (0..skip.get())
            .map(|_| TextLine {
                float_skip: true,
                ..TextLine::default()
            })
            .collect();
        skipped.append(&mut lines);
        lines = skipped;
    }
    super::bidi::reorder_lines(&mut lines, cjk.bidi, seen_fonts);
    for line in lines.iter_mut() {
        line.rtl_paragraph = cjk.bidi;
    }
    lines
}

fn find_next_tab_stop(
    current_x: f32,
    tab_stops: &[TabStop],
    indent_left: f32,
    default_tab_interval: f32,
) -> (TabStop, bool) {
    let abs_x = current_x + indent_left;
    if let Some(s) = tab_stops.iter().find(|s| s.position > abs_x + 0.5) {
        return (s.clone(), true);
    }
    let interval = if default_tab_interval > 0.0 {
        default_tab_interval
    } else {
        DEFAULT_TAB_INTERVAL
    };
    let next = ((abs_x / interval).floor() + 1.0) * interval;
    let default = TabStop {
        position: next,
        alignment: TabAlignment::Left,
        leader: None,
    };
    (default, false)
}

/// The width of the whitespace a tab segment opens with, before its first
/// word (`segment_width` leaves it out).
fn leading_space_width(runs: &[&Run], seen_fonts: &HashMap<String, FontEntry>) -> f32 {
    let mut w = 0.0;
    let mut key_buf = String::new();
    for run in runs {
        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key).expect("font registered");
        let eff_fs = effective_font_size(run, entry);
        let text = effective_text(run);
        let spaces = text.chars().take_while(|c| c.is_whitespace()).count();
        let mut prev_char = None;
        w += spaces as f32 * word_space_width(run, entry, eff_fs, " ", &mut prev_char);
        if spaces < text.chars().count() {
            break;
        }
    }
    w
}

fn segment_width(runs: &[&Run], seen_fonts: &HashMap<String, FontEntry>) -> f32 {
    let mut w: f32 = 0.0;
    let mut first = true;
    // Whitespace since the last word, across run boundaries: every space
    // counts ("Office:  317-477-4380" after a right tab has two), and runs
    // that meet without one join their words.
    let mut spaces = 0usize;
    let mut key_buf = String::new();
    for run in runs {
        if run.is_line_break {
            break;
        }
        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key).expect("font registered");
        let eff_fs = effective_font_size(run, entry);
        let ts = run.text_scale / 100.0;
        let cs = run.char_spacing;
        let text = effective_text(run);
        let mut prev_char = None;
        for piece in text.split(char::is_whitespace) {
            if piece.is_empty() {
                spaces += 1;
                continue;
            }
            let space_w = word_space_width(run, entry, eff_fs, piece, &mut prev_char);
            if !first {
                w += space_w * spaces as f32;
            }
            w += word_width_for_run(entry, run, piece, eff_fs, run.kerns_at(eff_fs), cs, ts);
            first = false;
            spaces = 1;
        }
        // The last piece is followed by no separator of its own.
        spaces -= 1;
    }
    w
}

/// The text a tab takes along when it wraps: the segment's words up to its
/// first space. Word never breaks between a tab and the word after it.
fn first_word_width(runs: &[&Run], seen_fonts: &HashMap<String, FontEntry>) -> f32 {
    let mut w = 0.0;
    let mut key_buf = String::new();
    for run in runs {
        if run.is_line_break || run.inline_image.is_some() {
            break;
        }
        let text = effective_text(run);
        let word = text.split(is_break_space).next().unwrap_or("");
        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key).expect("font registered");
        let eff_fs = effective_font_size(run, entry);
        let (cs, ts) = (run.char_spacing, run.text_scale / 100.0);
        w += word_width_for_run(entry, run, word, eff_fs, run.kerns_at(eff_fs), cs, ts);
        if word.len() < text.len() {
            break;
        }
    }
    w
}

fn decimal_before_width(runs: &[&Run], seen_fonts: &HashMap<String, FontEntry>) -> f32 {
    let texts: Vec<Cow<'_, str>> = runs.iter().map(|r| effective_text(r)).collect();
    let full_text: String = texts.iter().map(|t| t.as_ref()).collect();
    let before = if let Some(dot_pos) = full_text.find('.') {
        &full_text[..dot_pos]
    } else {
        &full_text
    };
    let mut w: f32 = 0.0;
    let mut chars_remaining = before.len();
    let mut key_buf = String::new();
    for (run, text) in runs.iter().zip(texts.iter()) {
        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key).expect("font registered");
        let eff_fs = effective_font_size(run, entry);
        let ts = run.text_scale / 100.0;
        let cs = run.char_spacing;
        let text_to_measure = if text.len() <= chars_remaining {
            chars_remaining -= text.len();
            text.as_ref()
        } else {
            let s = &text[..chars_remaining];
            chars_remaining = 0;
            s
        };
        let kern = run.kerns_at(eff_fs);
        w += entry.word_width(text_to_measure, eff_fs, kern) * ts
            + cs * text_to_measure.chars().count() as f32;
        if chars_remaining == 0 {
            break;
        }
    }
    w
}

/// Build TextLines for a paragraph that contains tab characters.
/// Wraps to new lines when content exceeds `max_width`.
pub(super) fn build_tabbed_line(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    tab_stops: &[TabStop],
    indent_left: f32,
    max_width: f32,
    indent_right: f32,
    first_line_hanging: f32,
    inline_image_names: &HashMap<usize, String>,
    effect_inline_names: &HashMap<usize, super::images::EffectXObjs>,
    default_tab_stop: f32,
    tab_exclusions: &[(f32, f32)],
    compat_mode: u32,
    cjk: CjkLayout,
) -> Vec<TextLine> {
    // Split runs into segments at tab markers, tracking original run indices.
    // The fourth tuple element is the `<w:tab/>` run itself (when present) so the
    // segment can see its underline/color formatting and render a decoration line
    // across the tab gap — the Word pattern for underlined signature/HR lines.
    let mut segments: Vec<(Vec<&Run>, Vec<usize>, Option<TabStop>, Option<&Run>)> = Vec::new();
    let mut current_seg: Vec<&Run> = Vec::new();
    let mut current_indices: Vec<usize> = Vec::new();
    let mut pending_tab: Option<TabStop> = None;
    let mut pending_tab_run: Option<&Run> = None;

    for (global_idx, run) in runs.iter().enumerate() {
        if run.vanish {
            continue;
        }
        if run.is_tab {
            segments.push((
                std::mem::take(&mut current_seg),
                std::mem::take(&mut current_indices),
                pending_tab.take(),
                pending_tab_run.take(),
            ));
            pending_tab = Some(TabStop {
                position: 0.0, // placeholder, resolved below
                alignment: TabAlignment::Left,
                leader: None,
            });
            pending_tab_run = Some(run);
        } else {
            current_seg.push(run);
            current_indices.push(global_idx);
        }
    }
    segments.push((
        std::mem::take(&mut current_seg),
        std::mem::take(&mut current_indices),
        pending_tab.take(),
        pending_tab_run.take(),
    ));

    let mut result_lines: Vec<TextLine> = Vec::new();
    let mut squeezed_lines: Vec<usize> = Vec::new();
    let mut justify_from = 0usize;
    let mut all_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;
    let mut pending_space_w: f32 = 0.0;
    // Underline state of the run that emitted the pending whitespace, so an
    // underline bridges spaces inside an underlined run (see build_paragraph_lines).
    let mut pending_space_underline = false;
    let mut pending_space_double = false;
    let mut pending_space_color: Option<[u8; 3]> = None;
    let mut pending_space_border: Option<ParagraphBorder> = None;
    let mut key_buf = String::new();
    let mut is_first_line = true;
    // Set when tabs wrap onto a new line that has nothing drawn yet.
    let mut tab_wrapped_line = false;
    // How far this line's words may run past its end: to a right, centre or
    // decimal stop in the right indent (western_australia's "34(1), 48(3)"),
    // or anywhere after an explicit stop past the margin before compat 15.
    let mut line_reach = 0.0f32;

    let mut tab_line_metrics: Option<(f32, Option<f32>)> = None;
    for (seg_idx, (seg_runs, seg_indices, tab_before, tab_run_before)) in
        segments.iter().enumerate()
    {
        let line_max = if is_first_line {
            max_width + first_line_hanging
        } else {
            max_width
        };
        let line_indent = if is_first_line {
            indent_left - first_line_hanging
        } else {
            indent_left
        };

        // Track the tab stop position so we can ensure current_x advances past it
        let mut tab_stop_pos: Option<f32> = None;

        if seg_idx > 0 {
            // A tab is pure positioning here, but extraction must still see
            // the word boundary (Word's text reads "(2) If", not "(2)If").
            mark_space_after(&mut all_chunks);
            // Trailing-space handling before a tab depends on whether an
            // explicit tab stop applies:
            // - An explicit stop AFTER current_x consumes the trailing spaces
            //   (header/footer center+right pattern — the tab must snap to that
            //   stop regardless of how wide the preceding whitespace is).
            // - Otherwise we fall through to default tab stops; trailing spaces
            //   advance the cursor so the tab can snap past them. This matters
            //   for list paragraphs with long dot-leader trailing-space runs.
            let abs_x_no_spaces = current_x + line_indent;
            let has_explicit_after = tab_stops.iter().any(|s| s.position > abs_x_no_spaces + 0.5);
            if !has_explicit_after {
                current_x += pending_space_w;
            }
            pending_space_w = 0.0;
            // A positional tab (w:ptab) resolves against the margin box with its own
            // alignment, bypassing the paragraph's tab stops. Left→box left, Center→box
            // center, Right→box right edge. This is what produces Word's left/center/right
            // footer layout that ordinary tab stops can't express here.
            let ptab_align = tab_run_before.and_then(|r| r.ptab_alignment);
            let (stop, mut effective_tab_target, explicit) = if let Some(palign) = ptab_align {
                let target = match palign {
                    TabAlignment::Center => max_width / 2.0,
                    TabAlignment::Right => max_width,
                    _ => 0.0,
                };
                (
                    TabStop {
                        position: target + line_indent,
                        alignment: palign,
                        leader: None,
                    },
                    target,
                    false,
                )
            } else {
                let (mut s, mut explicit) =
                    find_next_tab_stop(current_x, tab_stops, line_indent, default_tab_stop);
                // Word advances a left tab past any floating image whose body
                // occludes the stop, snapping to the first stop clear of the
                // image's right edge. `tab_exclusions` are (left, right) spans
                // in the same from-text-margin space as `s.position`.
                loop {
                    let bumped = tab_exclusions
                        .iter()
                        .find(|&&(ex_l, ex_r)| s.position > ex_l + 0.5 && s.position < ex_r - 0.5);
                    match bumped {
                        Some(&(_, ex_r)) => {
                            (s, explicit) = find_next_tab_stop(
                                ex_r - line_indent,
                                tab_stops,
                                line_indent,
                                default_tab_stop,
                            );
                        }
                        None => break,
                    }
                }
                let t = s.position - line_indent;
                (s, t, explicit)
            };
            let mut seg_start = resolve_tab_aligned_start(
                &stop,
                effective_tab_target,
                seg_runs,
                seen_fonts,
                current_x,
            );
            let mut resolved_leader = stop.leader;

            // Explicit tab stops may legitimately target positions beyond the
            // paragraph's right indent — TOC entries are a common case where a
            // right-aligned page-number tab sits near the content edge. Allow
            // segments to extend up to the physical content edge (indent_right
            // beyond max_width) before forcing a wrap.
            let wrap_limit = line_max + indent_right;
            // Word probes on an explicit stop past the margin: before Word
            // 2013 layout it keeps the tab and all that follows on this line,
            // off the page if need be (polish_building's 1216pt stop); from
            // 2013 on a right, centre or decimal stop clamps to the margin
            // (carbon_farming's TOC page numbers) and a left one wraps.
            // Otherwise a tab wraps when it lands past the margin or the word
            // after it doesn't fit (a default stop on the margin stays:
            // massachusetts' signature line keeps its fifth tab there).
            let stop_past = effective_tab_target > wrap_limit;
            if explicit && stop_past && compat_mode < 15 {
                line_reach = f32::INFINITY;
            } else if explicit && stop_past && stop.alignment != TabAlignment::Left {
                effective_tab_target = wrap_limit;
                seg_start =
                    resolve_tab_aligned_start(&stop, wrap_limit, seg_runs, seen_fonts, current_x);
            } else if (seg_start > wrap_limit
                || (stop.alignment == TabAlignment::Left
                    && seg_start + first_word_width(seg_runs, seen_fonts) > line_max))
                // A line of nothing but tabs wraps too (chiseldon's 27 default
                // tabs take two lines in Word); one tab on an empty line stays.
                && (!all_chunks.is_empty() || current_x > 0.0)
            {
                let mut line = TextLine {
                    justify_from: std::mem::take(&mut justify_from),
                    ..finish_line(&mut all_chunks)
                };
                // Leading tabs occupy a line even without printable chunks.
                // The overflowing tab moves with its following word; the
                // tabs already consumed size the otherwise blank line.
                if line.chunks.is_empty()
                    && let Some((font_size, lhr)) = tab_line_metrics
                {
                    line.break_font_size = Some(font_size);
                    line.break_lhr = lhr;
                    line.leading_tab_wrap = true;
                }
                tab_line_metrics = None;
                result_lines.push(line);
                tab_wrapped_line = true;
                line_reach = 0.0;
                current_x = 0.0;
                is_first_line = false;
                let (new_stop, _) =
                    find_next_tab_stop(0.0, tab_stops, indent_left, default_tab_stop);
                let new_target = new_stop.position - indent_left;
                seg_start =
                    resolve_tab_aligned_start(&new_stop, new_target, seg_runs, seen_fonts, 0.0);
                resolved_leader = new_stop.leader;
                effective_tab_target = new_target;
            }
            if stop.alignment != TabAlignment::Left {
                line_reach = line_reach.max(effective_tab_target);
            }
            if let Some(tab_run) = tab_run_before {
                let entry = seen_fonts
                    .get(font_key_buf(tab_run, &mut key_buf))
                    .expect("font registered");
                let fs = effective_font_size(tab_run, entry);
                let lhr = entry.line_h_ratio;
                if tab_line_metrics.is_none_or(|(old_fs, old_lhr)| {
                    fs * lhr.unwrap_or(1.2) > old_fs * old_lhr.unwrap_or(1.2)
                }) {
                    tab_line_metrics = Some((fs, lhr));
                }
            }

            // Draw leader fill between end of previous text and start of aligned text
            if tab_before.is_some() {
                let leader = resolved_leader;

                // Tab runs with `<w:u>` render the tab gap as an underlined line
                // (Word pattern for signature/HR lines). Emit a decoration-only
                // chunk so the underline decoration pass picks it up.
                if let Some(tab_run) = tab_run_before
                    && tab_run.underline
                    && seg_start > current_x + 0.01
                {
                    let font_run: &Run = seg_runs
                        .first()
                        .copied()
                        .or_else(|| runs.iter().find(|r| !r.font_name.is_empty()))
                        .unwrap_or(tab_run);
                    let key = font_key_buf(font_run, &mut key_buf);
                    let entry = seen_fonts.get(key).expect("font registered");
                    let eff_fs = effective_font_size(tab_run, entry).max(font_run.font_size);
                    all_chunks.push(WordChunk::tab_underline(
                        entry,
                        eff_fs,
                        tab_run.color,
                        tab_run.double_underline,
                        tab_run.border.clone(),
                        current_x,
                        seg_start - current_x,
                    ));
                }

                if let Some(leader_char) = leader {
                    let font_run: Option<&Run> = seg_runs
                        .first()
                        .copied()
                        .or_else(|| {
                            segments[..seg_idx]
                                .iter()
                                .rev()
                                .flat_map(|(r, _, _, _)| r.last().copied())
                                .next()
                        })
                        .or_else(|| {
                            // Tab-only paragraphs: fall back to any run (including tab runs)
                            runs.iter().find(|r| !r.font_name.is_empty())
                        });
                    if let Some(run) = font_run {
                        let key = font_key_buf(run, &mut key_buf);
                        let entry = seen_fonts.get(key).expect("font registered");
                        let eff_fs = effective_font_size(run, entry);
                        let char_w = entry.char_width_1000(leader_char) * eff_fs / 1000.0;
                        let leader_gap = seg_start - current_x;
                        if char_w > 0.0 && leader_gap > char_w * 2.0 {
                            let count = ((leader_gap - char_w) / char_w).floor() as usize;
                            if count > 0 {
                                let leader_text: String =
                                    std::iter::repeat_n(leader_char, count).collect();
                                let leader_w = count as f32 * char_w;
                                let leader_start = seg_start - leader_w;
                                all_chunks.push(WordChunk::leader(
                                    entry,
                                    leader_text,
                                    eff_fs,
                                    run.color,
                                    leader_start,
                                    leader_w,
                                ));
                            }
                        }
                    }
                }
            }

            current_x = seg_start;
            tab_stop_pos = Some(effective_tab_target);
            justify_from = all_chunks.len();
        }

        // Layout text in this segment from current_x
        for (local_idx, run) in seg_runs.iter().enumerate() {
            if run.is_line_break {
                mark_space_after(&mut all_chunks);
                result_lines.push(TextLine {
                    justify_from: std::mem::take(&mut justify_from),
                    ..finish_line_with_break(&mut all_chunks)
                });
                tab_wrapped_line = false;
                tab_line_metrics = None;
                line_reach = 0.0;
                current_x = 0.0;
                is_first_line = false;
                pending_space_w = 0.0;
                continue;
            }

            // Handle inline images (same pattern as build_paragraph_lines)
            if let Some(img) = &run.inline_image {
                if let Some(pdf_name) = inline_image_names.get(&seg_indices[local_idx]) {
                    all_chunks.push(WordChunk::image(
                        pdf_name,
                        run.font_size,
                        current_x,
                        img,
                        effect_inline_names.get(&seg_indices[local_idx]).cloned(),
                    ));
                    current_x += img.layout_size().0;
                }
                continue;
            }

            if let Some((entry, mb)) = math_box(run, seen_fonts) {
                current_x += pending_space_w;
                pending_space_w = 0.0;
                all_chunks.push(WordChunk::math(entry, run, mb, current_x));
                current_x += mb.width;
                continue;
            }

            let key = font_key_buf(run, &mut key_buf);
            let entry = seen_fonts.get(key).expect("font registered");
            let eff_fs = effective_font_size(run, entry);
            let y_off = vert_y_offset(run);
            let text = &run.text;

            let cs = run.char_spacing;
            let ts = run.text_scale / 100.0;
            let segments = split_preserving_spaces(text);
            let mut prev_char = None;
            for (seg_idx, &(space_count, source)) in segments.iter().enumerate() {
                let shown = caps_word(run, source);
                let word: &str = &shown;
                let original = run.caps.then_some(source);
                let kern = run.kerns_at(eff_fs);
                let ww = word_width_for_run(entry, run, word, eff_fs, kern, cs, ts);
                pending_space_w += space_count as f32
                    * word_space_width(run, entry, eff_fs, source, &mut prev_char);
                if space_count > 0 {
                    pending_space_underline = run.underline;
                    pending_space_double = run.double_underline;
                    pending_space_color = run.color;
                    pending_space_border = run.border.clone();
                }
                // Word advances over explicit spaces wherever they occur, including
                // as the first content of a line or right after a tab. Whitespace
                // that is alone in its run (e.g. "<w:tab/>   " at one size, then
                // text at another) only reaches here via pending_space_w with
                // space_count == 0 and nothing emitted yet, so it must still be
                // applied. pending_space_w is already zeroed at wraps, breaks and
                // tabs, so this never re-applies absorbed line-end spaces.
                let applied_space = pending_space_w > 0.0;
                if applied_space {
                    mark_space_after(&mut all_chunks);
                    if pending_space_underline && !all_chunks.is_empty() {
                        all_chunks.push(WordChunk::word_space_underline(
                            entry,
                            eff_fs,
                            pending_space_color,
                            pending_space_double,
                            pending_space_border.clone(),
                            current_x,
                            pending_space_w,
                        ));
                    }
                    current_x += pending_space_w;
                    pending_space_w = 0.0;
                }
                // Word continues previous word across run boundary (no whitespace between)
                let is_continuation = seg_idx == 0 && !applied_space && !all_chunks.is_empty();
                let cur_line_max = if is_first_line {
                    max_width + first_line_hanging
                } else {
                    max_width
                };
                let mut overflows = current_x + ww > cur_line_max.max(line_reach)
                    && !all_chunks.is_empty()
                    && !is_continuation;
                // Justified compat-15 lines keep the word by narrowing the
                // spaces after the last tab, as untabbed lines do
                // (`SPACE_SQUEEZE`); ukrainian_municipal's "1.<tab>Надати …"
                // line ends in "опалення" with its spaces at 93%.
                if overflows && cjk.squeeze_spaces && cur_line_max >= line_reach {
                    let after_tab = all_chunks.get(justify_from..).unwrap_or(&[]);
                    let lead_space = after_tab.last().map_or(0.0, |last| {
                        (current_x - last.x_offset - last.width).max(0.0)
                    });
                    let spaces = line_space_width(after_tab) + lead_space;
                    if squeeze_keeps(current_x + ww - cur_line_max, ww, lead_space)
                        && current_x + ww - cur_line_max <= SPACE_SQUEEZE * spaces
                    {
                        overflows = false;
                        squeezed_lines.push(result_lines.len());
                    }
                }
                if overflows {
                    result_lines.push(TextLine {
                        justify_from: std::mem::take(&mut justify_from),
                        ..finish_line(&mut all_chunks)
                    });
                    tab_wrapped_line = false;
                    tab_line_metrics = None;
                    line_reach = 0.0;
                    current_x = 0.0;
                    is_first_line = false;
                }
                push_word_chunks(
                    &mut all_chunks,
                    entry,
                    run,
                    word,
                    original,
                    eff_fs,
                    cs,
                    y_off,
                    current_x,
                    ww,
                    false,
                    cjk.compress_punct,
                );
                current_x += ww;
            }
            // Accumulate trailing whitespace for the next run (or the next tab stop)
            let trailing_spaces = text
                .chars()
                .rev()
                .take_while(|c| is_break_space(*c))
                .count();
            if trailing_spaces > 0 {
                pending_space_w += trailing_spaces as f32
                    * word_space_width(run, entry, eff_fs, "", &mut prev_char);
                pending_space_underline = run.underline;
                pending_space_double = run.double_underline;
                pending_space_color = run.color;
                pending_space_border = run.border.clone();
            }
        }

        // For decimal/right/center tabs, text may end before the tab stop position.
        // Advance current_x past the tab stop so the next tab finds the correct stop.
        if let Some(ts_pos) = tab_stop_pos {
            current_x = current_x.max(ts_pos);
        }
    }

    // Finalize remaining chunks into the last line. Tabs that wrapped keep
    // their line though nothing is drawn on it: bulgarian_road_safety's
    // trailing tabs after "/Зл. Атанасова/" take a second line in Word.
    if !all_chunks.is_empty() {
        result_lines.push(TextLine {
            justify_from: std::mem::take(&mut justify_from),
            ..finish_line(&mut all_chunks)
        });
    } else if result_lines.is_empty() || tab_wrapped_line {
        result_lines.push(TextLine::default());
    }
    for i in squeezed_lines {
        if let Some(line) = result_lines.get_mut(i) {
            line.squeezed = true;
        }
    }

    // Trailing break creates an empty line (same as build_paragraph_lines)
    if result_lines.last().is_some_and(|l| l.ends_with_break) {
        let break_fs = runs
            .iter()
            .rev()
            .find(|r| r.is_line_break)
            .map(|r| r.font_size);
        result_lines.push(TextLine {
            break_font_size: break_fs,
            ..TextLine::default()
        });
    }

    for line in result_lines.iter_mut() {
        line.rtl_paragraph = cjk.bidi;
    }
    result_lines
}

/// A font's OS/2 strikeout as Word draws it: its top edge above the baseline
/// and its thickness, each on a 0.25pt grid (online export probe, the
/// underline's 12 cases). The footnote separator's rule is one too.
pub(super) fn os2_strike(entry: &FontEntry, font_size: f32) -> Option<(f32, f32)> {
    let (pos, th) = entry.strikeout?;
    let q = |v: f32| (v * 4.0).round() / 4.0;
    Some((q(pos * font_size), q(th * font_size).max(0.25)))
}

/// Draws `text` with the font's pair kerning as TJ adjustments, the same pairs
/// the word's width was measured with: Word kerns inside words too (Aptos
/// Display "Te" 1.6pt tighter at 20pt, case3), so plain Tj left every glyph
/// after a kerned pair out of place.
fn show_kerned(content: &mut Content, entry: &FontEntry, text: &str, boundary_space: bool) {
    let mut tj = content.show_positioned();
    let mut items = tj.items();
    let mut run = String::new();
    let mut prev: Option<char> = None;
    for ch in text.chars() {
        let k = prev.map_or(0.0, |p| entry.kern_1000(p, ch));
        if k != 0.0 {
            items.show(Str(&entry.encode(&run)));
            run.clear();
            // TJ subtracts: a negative kern (tighter) is a positive adjustment.
            items.adjust(-k);
        }
        run.push(ch);
        prev = Some(ch);
    }
    if boundary_space {
        run.push(' ');
    }
    items.show(Str(&entry.encode(&run)));
}

/// Draws `glyphs` (text, x, advance) in the order given, each at its own x
/// from the pen's start, with TJ moves between them. Moves and advances
/// scale alike with Tz, so all of this is in thousandths of the font size.
/// Whole thousandths keep the numbers short; the pen follows the rounded
/// moves, so the error never adds up. `trailing` is shown last, unplaced.
fn show_placed(
    content: &mut Content,
    entry: &FontEntry,
    glyphs: &[(&str, f32, f32)],
    trailing: Option<&str>,
) {
    let mut tj = content.show_positioned();
    let mut items = tj.items();
    let mut pen = 0.0f32;
    for &(text, x, advance) in glyphs {
        // TJ subtracts: moving right is a negative adjustment.
        let adjust = (pen - x).round();
        if adjust != 0.0 {
            items.adjust(adjust);
        }
        items.show(Str(&entry.encode(text)));
        pen += advance - adjust;
    }
    if let Some(t) = trailing {
        items.show(Str(&entry.encode(t)));
    }
}

/// `text` split into characters, each with its advance in thousandths;
/// `marks_attach` keeps combining marks with the letter before them.
fn glyph_clusters<'a>(
    entry: &FontEntry,
    text: &'a str,
    extra: f32,
    marks_attach: bool,
) -> Vec<(&'a str, f32)> {
    let mut out: Vec<(&'a str, f32)> = Vec::new();
    let mut start = 0;
    for (i, ch) in text.char_indices() {
        let w = entry.char_width_1000(ch) + extra;
        let end = i + ch.len_utf8();
        match out.last_mut() {
            Some((s, aw)) if marks_attach && super::bidi::class(ch) == super::bidi::Class::NSM => {
                *s = &text[start..end];
                *aw += w;
            }
            _ => {
                start = i;
                out.push((&text[i..end], w));
            }
        }
    }
    out
}

/// A right-to-left chunk's glyphs in logical order, each placed left of the
/// one before (the pen starts at the chunk's left edge), so text extraction
/// and tags read the word as written. A letter's combining marks follow it
/// with no move: they hang back over it from its end, as in left-to-right
/// text.
fn show_rtl(
    content: &mut Content,
    entry: &FontEntry,
    text: &str,
    font_size: f32,
    char_spacing: f32,
    boundary_space: bool,
) {
    let tc = char_spacing * 1000.0 / font_size.max(0.01);
    let clusters = glyph_clusters(entry, text, tc, true);
    let mut right: f32 = clusters.iter().map(|&(_, w)| w).sum();
    let placed: Vec<(&str, f32, f32)> = clusters
        .into_iter()
        .map(|(s, w)| {
            right -= w;
            (s, right, w)
        })
        .collect();
    show_placed(content, entry, &placed, boundary_space.then_some(" "));
}

/// `text`'s glyphs in logical order, each at the place its right-to-left
/// layout (bidi, paragraph level 1) gives it, the pen starting at the left
/// end: a list label that reads "3." draws ".3".
pub(super) fn show_visual(content: &mut Content, entry: &FontEntry, text: &str) {
    let glyphs = glyph_clusters(entry, text, 0.0, false);
    let cls: Vec<_> = text.chars().map(super::bidi::class).collect();
    let (levels, cluster) = super::bidi::resolve(&cls, &vec![1; cls.len()], 1);
    let mut x = vec![0.0f32; glyphs.len()];
    let mut left = 0.0;
    for i in super::bidi::visual_order(&levels, &cluster) {
        x[i] = left;
        left += glyphs[i].1;
    }
    let placed: Vec<(&str, f32, f32)> =
        glyphs.iter().zip(x).map(|(&(s, w), x)| (s, x, w)).collect();
    show_placed(content, entry, &placed, None);
}

pub(super) fn encode_text_for_pdf(
    text: &str,
    pdf_font: &str,
    pdf_name_to_entry: &HashMap<&str, &FontEntry>,
) -> Vec<u8> {
    match pdf_name_to_entry.get(pdf_font) {
        Some(e) => e.encode(text),
        None => to_winansi_bytes(text),
    }
}

/// §17.6.8 line numbering for a body paragraph's rendered lines. The counter is
/// shared (continuous) across paragraphs so it survives between calls.
pub(super) struct LineNumberArg<'a> {
    /// Body lines counted so far = 0-based index of the next line to number.
    pub counter: &'a mut u32,
    pub start: i32,
    pub count_by: u32,
    /// Page-coordinate x of the right edge the (right-aligned) numbers end at.
    pub right_x: f32,
    /// The registered face (PDF name) and size of the numbers; None falls
    /// back to the line's own font.
    pub font: Option<(String, f32)>,
}

pub(super) fn line_max_image_h(line: &TextLine) -> f32 {
    line.chunks
        .iter()
        .map(|c| c.inline_image_height + c.inline_image_extra_height)
        .fold(0.0f32, f32::max)
}

/// The tallest inline picture among `runs`, wrap distances included.
pub(super) fn runs_max_image_h(runs: &[Run]) -> f32 {
    runs.iter()
        .filter_map(|r| r.inline_image.as_ref())
        .map(|img| img.display_height + img.layout_extra_height)
        .fold(0.0f32, f32::max)
}

/// How far an inline picture lowers its line's baseline. Word sits the picture on
/// the baseline, so a picture taller than the paragraph's text ascent pushes the
/// baseline down by the difference. `ascent` is the paragraph's baseline offset
/// (tallest run's ascent), the same figure the caller places the first baseline
/// with, so the picture top lands exactly on the line top.
pub(super) fn inline_image_line_extra(line: &TextLine, ascent: f32) -> f32 {
    (line_max_image_h(line) - ascent).max(0.0)
}

/// Height of one laid-out line: the paragraph pitch, or for a picture line the
/// picture plus the text descent. Word gives the picture line no line gap and no
/// spacing-multiplier leading (italian_evaluation_minutes p7: a 36pt signature in
/// 10pt Arial makes a 38.2pt line; english_town_council p1: 149pt logo with a 48pt
/// run makes 159pt, annotation #230).
pub(super) fn inline_line_advance(
    line: &TextLine,
    line_pitch: f32,
    (ascent, descent): (f32, f32),
) -> f32 {
    let line_pitch = line.pitch.unwrap_or(line_pitch);
    let img_h = line_max_image_h(line);
    if img_h > ascent {
        line_pitch.max(img_h + descent)
    } else {
        line_pitch
    }
}

/// Height of a laid-out paragraph: the sum of its line advances, one pitch when
/// it has no lines.
pub(super) fn lines_height(lines: &[TextLine], line_pitch: f32, metrics: (f32, f32)) -> f32 {
    if lines.is_empty() {
        line_pitch
    } else {
        lines
            .iter()
            .map(|l| inline_line_advance(l, line_pitch, metrics))
            .sum()
    }
}

/// Fake a bold face by stroking the glyph outlines along with the fill.
fn begin_synthetic_bold(content: &mut Content, chunk: &WordChunk) {
    content.set_line_width(chunk.font_size * 0.02);
    stroke_color_or_black(content, chunk.color);
    content.set_text_rendering_mode(TextRenderingMode::FillStroke);
}

/// winDescent as a fraction of the font size: the line-height ratio less the
/// ascender ratio (identity used throughout), 0.25 when the font is unknown.
pub(super) fn descender_ratio(lhr: Option<f32>, ar: Option<f32>) -> f32 {
    lhr.zip(ar).map_or(0.25, |(l, a)| (l - a).max(0.0))
}

/// Render pre-built lines applying the paragraph alignment.
/// `first_baseline_y` is line 0's text baseline as if it held no picture; a
/// picture line drops its own baseline here (see `inline_image_line_extra`).
/// `total_line_count` is the full paragraph line count (for justify: last line stays left-aligned).
pub(super) fn render_paragraph_lines(
    content: &mut Content,
    lines: &[TextLine],
    alignment: &Alignment,
    margin_left: f32,
    text_width: f32,
    first_baseline_y: f32,
    line_pitch: f32,
    // Paragraph (ascent, descent) in points; only picture lines use them (see
    // `inline_line_advance`).
    text_metrics: (f32, f32),
    total_line_count: usize,
    first_line_index: usize,
    links: &mut Vec<LinkAnnotation>,
    first_line_hanging: f32,
    seen_fonts: &HashMap<String, FontEntry>,
    line_geometry: Option<&[(f32, f32)]>,
    gradient_specs: &mut Vec<super::GradientSpec>,
    mut comment_anchors: Option<&mut Vec<(u32, f32, f32, f32)>>,
    mut line_numbering: Option<LineNumberArg<'_>>,
    mut link_tags: Option<LinkTagger<'_>>,
) {
    let mut current_color: Option<[u8; 3]> = None;
    let mut pattern_fill_active = false;
    let mut cur_font_name = String::new();
    let mut cur_font_size: f32 = -1.0;
    let mut cur_char_spacing: f32 = 0.0;
    let mut cur_text_scale: f32 = 100.0;
    let mut cur_synthetic_bold = false;
    let mut has_text_outline = false;

    let pdf_name_to_entry: HashMap<&str, &FontEntry> = seen_fonts
        .values()
        .map(|e| (e.pdf_name.as_str(), e))
        .collect();

    // Per-line baseline offsets below `first_baseline_y`: each line's top is the
    // sum of the previous lines' advances, and a picture line drops its baseline
    // by the picture's surplus over the ascent; a line sized by its own smaller
    // face raises its baseline by `ascent_shift`. Normal lines reduce to `line_pitch`.
    let mut line_y_offsets: Vec<f32> = Vec::with_capacity(lines.len());
    let mut line_top = 0.0f32;
    for line in lines {
        line_y_offsets
            .push(line_top + line.ascent_shift + inline_image_line_extra(line, text_metrics.0));
        line_top += inline_line_advance(line, line_pitch, text_metrics);
    }

    let last_line_idx = total_line_count.saturating_sub(1);
    // The first-line indent goes to the first line with content, past any
    // `float_skip` lines.
    let first_content_idx = lines.iter().take_while(|l| l.float_skip).count();
    for (line_num, line) in lines.iter().enumerate() {
        let y = first_baseline_y - line_y_offsets[line_num];
        let global_line_idx = first_line_index + line_num;

        let (base_margin, base_width) = line_geometry
            .and_then(|g| g.get(global_line_idx))
            .copied()
            .unwrap_or((margin_left, text_width));
        let (eff_margin, eff_width) = if global_line_idx == first_content_idx
            && first_line_hanging.abs() > 0.001
        {
            let shift = if line.rtl_paragraph {
                0.0
            } else {
                first_line_hanging
            };
            (base_margin - shift, base_width + first_line_hanging)
        } else {
            (base_margin, base_width)
        };

        // §17.6.8 margin line number: every counted body line advances the
        // shared counter; numbers are right-aligned in the page margin at the
        // line's own baseline, in that line's font/size (matching Word).
        if let Some(ln) = line_numbering.as_mut() {
            let idx = *ln.counter;
            *ln.counter = idx + 1;
            let value = ln.start + idx as i32;
            let show = value >= 1 && (ln.count_by <= 1 || value % ln.count_by as i32 == 0);
            if show
                && let Some((font, fs)) = ln.font.clone().or_else(|| {
                    line.chunks
                        .iter()
                        .find(|c| c.inline_image_name.is_none() && !c.text.is_empty())
                        .map(|c| (c.pdf_font.clone(), c.font_size))
                })
            {
                let s = value.to_string();
                let w = pdf_name_to_entry
                    .get(font.as_str())
                    .map(|e| e.word_width(&s, fs, false))
                    .unwrap_or(fs * 0.5 * s.chars().count() as f32);
                let bytes = encode_text_for_pdf(&s, &font, &pdf_name_to_entry);
                let x = ln.right_x - w;
                let draw = |content: &mut Content| {
                    content.save_state();
                    content.set_char_spacing(0.0);
                    content.set_horizontal_scaling(100.0);
                    fill_color_or_black(content, None);
                    content.begin_text();
                    content.set_font(Name(font.as_bytes()), fs);
                    content.next_line(x, y);
                    content.show(Str(&bytes));
                    content.end_text();
                    content.restore_state();
                };
                // Margin numbering isn't the paragraph's text: a screen
                // reader would read "2Numbered line".
                match link_tags.as_mut() {
                    Some(lt) => lt.artifact(content, draw),
                    None => draw(content),
                }
            }
        }

        // Compute left-region content width (for alignment/justify)
        let left_content_width = if let Some(ref rr) = line.right_region {
            line.chunks[..rr.first_chunk_idx]
                .last()
                .map(|c| c.x_offset + c.width)
                .unwrap_or(0.0)
        } else {
            line.total_width
        };

        let left_chunk_count = line
            .right_region
            .as_ref()
            .map(|rr| rr.first_chunk_idx)
            .unwrap_or(line.chunks.len());
        let gaps_before = spaces_before_each(&line.chunks[..left_chunk_count]);
        let jf = line.justify_from.min(left_chunk_count.saturating_sub(1));
        let tab_gaps = gaps_before.get(jf).copied().unwrap_or(0);
        let gaps_before: Vec<usize> = gaps_before
            .iter()
            .enumerate()
            .map(|(i, &g)| if i < jf { 0 } else { g - tab_gaps })
            .collect();
        let left_gaps = gaps_before.last().copied().unwrap_or(0);

        // CJK justification: distribute space between every character, not just chunks.
        // Word treats CJK inter-character gaps the same as word gaps for justify.
        let left_char_count: usize = line.chunks[..left_chunk_count]
            .iter()
            .map(|c| c.text.chars().count())
            .sum();
        let has_cjk_content = line.chunks[..left_chunk_count]
            .iter()
            .any(|c| c.text.chars().any(crate::docx::is_east_asian_char));

        // Soft line breaks (w:br) should still be justified — only the
        // paragraph's true last line suppresses justification. `distribute`
        // stretches every line, last one included (§17.18.44).
        // A line wider than the measure only exists where the breaker kept a
        // word by squeezing spaces (`SPACE_SQUEEZE`); Word paints those spaces
        // narrower even on the paragraph's last line (mongolian_human_rights).
        let squeezed = line.squeezed && left_gaps > 0;
        let can_justify = match *alignment {
            Alignment::Justify => {
                (global_line_idx != last_line_idx || squeezed) && !line.natural_width
            }
            Alignment::Distribute => true,
            _ => false,
        };

        let mixed = (can_justify
            && has_cjk_content
            && matches!(*alignment, Alignment::Justify)
            && !(squeezed && line.squeeze_chars))
        .then(|| {
            MixedJustify::new(
                &line.chunks[..left_chunk_count],
                &gaps_before,
                eff_width - left_content_width,
            )
        });
        let char_justify_gaps = if mixed.is_some() {
            None
        } else if squeezed && line.squeeze_chars {
            Some(left_char_count.saturating_sub(1).max(1))
        } else {
            char_justify_gaps(*alignment, can_justify, has_cjk_content, left_char_count)
        };
        let is_char_justified = char_justify_gaps.is_some();
        let is_justified =
            mixed.is_some() || is_char_justified || (can_justify && left_gaps > 0);

        let line_start_x = match alignment {
            Alignment::Center => eff_margin + (eff_width - left_content_width) / 2.0,
            Alignment::Right => eff_margin + eff_width - left_content_width,
            // Word probe: a justified right-to-left paragraph's last line
            // ends at the right margin.
            Alignment::Justify | Alignment::Distribute if line.rtl_paragraph && !is_justified => {
                eff_margin + eff_width - left_content_width
            }
            Alignment::Left | Alignment::Justify | Alignment::Distribute => eff_margin,
        };

        // Char-level distribution via Tc (char spacing): see char_justify_gaps
        // for how many gaps the slack divides across.
        let justify_tc = match char_justify_gaps {
            Some(gaps) => (eff_width - left_content_width) / gaps as f32,
            None => 0.0,
        };

        let extra_per_gap = if is_justified && !is_char_justified && mixed.is_none() {
            let per_gap = (eff_width - left_content_width) / left_gaps.max(1) as f32;
            if squeezed { per_gap } else { per_gap.max(0.0) }
        } else {
            0.0
        };

        // Right-region start x and justify gap (if dual-region line)
        let (right_start_x, right_extra_per_gap) = if let Some(ref rr) = line.right_region {
            let right_chunks = line.chunks.len() - rr.first_chunk_idx;
            let rx = match alignment {
                Alignment::Center => rr.region_x + (rr.region_width - rr.content_width) / 2.0,
                Alignment::Right => rr.region_x + rr.region_width - rr.content_width,
                Alignment::Left | Alignment::Justify | Alignment::Distribute => rr.region_x,
            };
            let rgap = if is_justified && right_chunks > 1 {
                ((rr.region_width - rr.content_width) / (right_chunks - 1) as f32).max(0.0)
            } else {
                0.0
            };
            (rx, rgap)
        } else {
            (0.0, 0.0)
        };

        // Helper: compute absolute x for a chunk, accounting for dual regions
        let chunk_abs_x = |chunk_idx: usize, chunk: &WordChunk| -> f32 {
            if let Some(ref rr) = line.right_region
                && chunk_idx >= rr.first_chunk_idx
            {
                let local_idx = chunk_idx - rr.first_chunk_idx;
                return right_start_x + chunk.x_offset + local_idx as f32 * right_extra_per_gap;
            }
            if let Some(m) = &mixed {
                line_start_x + chunk.x_offset + m.offset_before(chunk_idx)
            } else if is_char_justified {
                // Tc adds extra space after each character; shift chunk start
                // by the cumulative Tc of all chars in preceding chunks.
                let chars_before: usize = line.chunks[..chunk_idx]
                    .iter()
                    .map(|c| c.text.chars().count())
                    .sum();
                line_start_x + chunk.x_offset + chars_before as f32 * justify_tc
            } else {
                line_start_x + chunk.x_offset + gaps_before[chunk_idx] as f32 * extra_per_gap
            }
        };

        let mut decorations: Vec<Decoration> = Vec::new();

        // Draw run shading first (bottom layer), then run highlights on top.
        // Both use the same rectangle geometry; when a run sets both, highlight
        // visually occludes shading — matching Word.
        let draw_run_backgrounds =
            |content: &mut Content, accessor: fn(&WordChunk) -> Option<[u8; 3]>| {
                let mut bg_start_x = 0.0f32;
                let mut bg_color: Option<[u8; 3]> = None;
                let mut bg_end_x = 0.0f32;
                let mut bg_fs = 0.0f32;

                let flush =
                    |content: &mut Content, color: [u8; 3], sx: f32, ex: f32, fs: f32, y: f32| {
                        let bg_bottom = y - fs * 0.2;
                        let bg_height = fs * 1.15;
                        content.save_state();
                        fill_color_or_black(content, Some(color));
                        content.rect(sx, bg_bottom, ex - sx, bg_height);
                        content.fill_nonzero();
                        content.restore_state();
                    };

                for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
                    let x = chunk_abs_x(chunk_idx, chunk);
                    let chunk_color = accessor(chunk);
                    if chunk_color == bg_color && bg_color.is_some() {
                        bg_end_x = x + chunk.width;
                        bg_fs = bg_fs.max(chunk.font_size);
                    } else {
                        if let Some(c) = bg_color {
                            flush(content, c, bg_start_x, bg_end_x, bg_fs, y);
                        }
                        if let Some(c) = chunk_color {
                            bg_start_x = x;
                            bg_end_x = x + chunk.width;
                            bg_fs = chunk.font_size;
                            bg_color = Some(c);
                        } else {
                            bg_color = None;
                        }
                    }
                }
                if let Some(c) = bg_color {
                    flush(content, c, bg_start_x, bg_end_x, bg_fs, y);
                }
            };

        draw_run_backgrounds(content, |c| c.shading);
        draw_run_backgrounds(content, |c| c.highlight);

        let mut border_start_x = 0.0f32;
        let mut border_end_x = 0.0f32;
        let mut border_fs = 0.0f32;
        let mut active_border: Option<ParagraphBorder> = None;
        let flush_border =
            |content: &mut Content, border: &ParagraphBorder, sx: f32, ex: f32, fs: f32, y: f32| {
                let pad = border.space_pt;
                let bottom = y - fs * 0.2 - pad;
                let height = fs * 1.15 + pad * 2.0;
                content.save_state();
                content.set_line_width(border.width_pt.max(0.1));
                stroke_color_or_black(content, Some(border.color));
                content.rect(sx - pad, bottom, (ex - sx) + pad * 2.0, height);
                content.stroke();
                content.restore_state();
            };
        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            let x = chunk_abs_x(chunk_idx, chunk);
            match (&active_border, &chunk.border) {
                (Some(active), Some(next)) if active == next => {
                    border_end_x = x + chunk.width;
                    border_fs = border_fs.max(chunk.font_size);
                }
                (Some(active), next) => {
                    flush_border(content, active, border_start_x, border_end_x, border_fs, y);
                    active_border = next.clone();
                    if chunk.border.is_some() {
                        border_start_x = x;
                        border_end_x = x + chunk.width;
                        border_fs = chunk.font_size;
                    }
                }
                (None, Some(next)) => {
                    active_border = Some(next.clone());
                    border_start_x = x;
                    border_end_x = x + chunk.width;
                    border_fs = chunk.font_size;
                }
                (None, None) => {}
            }
        }
        if let Some(active) = &active_border {
            flush_border(content, active, border_start_x, border_end_x, border_fs, y);
        }

        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            if let Some(cb) = chunk.checkbox
                && let Some(entry) = pdf_name_to_entry.get(chunk.pdf_font.as_str())
            {
                let x = chunk_abs_x(chunk_idx, chunk);
                draw_checkbox(content, entry, cb, x, y + chunk.y_offset, chunk.color);
            }
        }

        if let Some(ref mut anchors) = comment_anchors {
            for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
                if chunk.comment_ids.is_empty() {
                    continue;
                }
                let end_x = chunk_abs_x(chunk_idx, chunk) + chunk.width;
                // Anchor at the TOP of the highlight rectangle so the matching
                // callout's top edge can align with the highlight's top edge
                // (matches Word's pane layout). The pane renderer computes the
                // connector origin (highlight bottom) using the font_size we
                // store alongside, matching `bg_bottom = y - fs * 0.2`.
                let anchor_y = y - chunk.font_size * 0.2 + chunk.font_size * 1.15;
                for &cid in &chunk.comment_ids {
                    if let Some(entry) = anchors.iter_mut().find(|(id, _, _, _)| *id == cid) {
                        entry.1 = end_x;
                        entry.2 = anchor_y;
                        entry.3 = chunk.font_size;
                    } else {
                        anchors.push((cid, end_x, anchor_y, chunk.font_size));
                    }
                }
            }
        }

        // Decoration-only chunks (empty text with underline) also need the text-block
        // pass so their underline geometry is collected into `decorations`.
        let has_text_chunks = line
            .chunks
            .iter()
            .any(|c| c.inline_image_name.is_none() && (!c.text.is_empty() || c.underline));

        if has_text_chunks {
            content.begin_text();
            let mut td_x = 0.0_f32;
            let mut td_y = 0.0_f32;
            let mut cur_shear = 0.0_f32;

            // Logical order, so the content stream (and the tags) read the
            // text as written even where bidi stored the chunks visually.
            for k in 0..line.chunks.len() {
                let chunk_idx = line.draw_order.get(k).copied().unwrap_or(k);
                let chunk = &line.chunks[chunk_idx];
                if chunk.inline_image_name.is_some() || chunk.checkbox.is_some() {
                    continue;
                }
                // A note's reference mark links to the note text (keyboard and
                // screen-reader navigation), like Word's. Kept apart from
                // hyperlink_url, which also moves the underline.
                let note = chunk.note();
                let note_url = note.map(|(endnote, id)| {
                    format!("#{}", super::footnotes::note_anchor(endnote, id))
                });
                let link_url = chunk.hyperlink_url.as_deref().or(note_url.as_deref());
                let primary_entry = pdf_name_to_entry.get(chunk.pdf_font.as_str());
                // The next chunk is positioned from the line start, so the
                // space's advance moves nothing; it only marks the word boundary.
                let boundary_space = chunk.boundary_space(primary_entry.copied());
                if let Some(lt) = link_tags.as_mut()
                    && lt.chunk(content, chunk, link_url, boundary_space)
                {
                    td_x = 0.0;
                    td_y = 0.0;
                }

                let x = chunk_abs_x(chunk_idx, chunk);
                let cy = y + chunk.y_offset;

                // w14:textFill gradient → register an axial-shading pattern spanning
                // the glyph extent and use it as the fill color for this chunk.
                // Solid fill overrides w:color; NoFill falls through to w:color (the
                // outline path zeroes the fill via text rendering mode = Stroke).
                let mut chunk_uses_gradient = false;
                if let Some(TextFill::Gradient {
                    ref stops,
                    angle_deg,
                }) = chunk.text_fill
                {
                    let pat_name = format!("Grd{}", gradient_specs.len());
                    let y_bottom = cy - chunk.font_size * 0.2;
                    gradient_specs.push(super::GradientSpec {
                        pattern_name: pat_name.clone(),
                        stops: stops.clone(),
                        angle_deg,
                        x,
                        y: y_bottom,
                        w: chunk.width.max(1.0),
                        h: chunk.font_size,
                    });
                    content.set_fill_color_space(pdf_writer::types::ColorSpaceOperand::Pattern);
                    content.set_fill_pattern([], Name(pat_name.as_bytes()));
                    pattern_fill_active = true;
                    current_color = None;
                    chunk_uses_gradient = true;
                }

                let effective_color = match chunk.text_fill {
                    Some(TextFill::Solid(c)) => Some(c),
                    _ => chunk.color,
                };
                if !chunk_uses_gradient && (pattern_fill_active || effective_color != current_color)
                {
                    fill_color_or_black(content, effective_color);
                    current_color = effective_color;
                    pattern_fill_active = false;
                }

                // Text outline takes priority over synthetic bold
                if let Some(ref outline) = chunk.text_outline {
                    if !has_text_outline {
                        super::wordart::apply_text_outline(
                            content,
                            outline,
                            chunk.text_fill.as_ref(),
                        );
                        has_text_outline = true;
                    }
                } else if has_text_outline {
                    super::wordart::reset_text_outline(content);
                    has_text_outline = false;
                    if cur_synthetic_bold {
                        begin_synthetic_bold(content, chunk);
                    }
                }

                if !has_text_outline && chunk.synthetic_bold != cur_synthetic_bold {
                    if chunk.synthetic_bold {
                        begin_synthetic_bold(content, chunk);
                    } else {
                        content.set_text_rendering_mode(TextRenderingMode::Fill);
                    }
                    cur_synthetic_bold = chunk.synthetic_bold;
                }

                let effective_cs = chunk.char_spacing + justify_tc;
                if effective_cs != cur_char_spacing {
                    content.set_char_spacing(effective_cs);
                    cur_char_spacing = effective_cs;
                }
                if chunk.text_scale != cur_text_scale {
                    content.set_horizontal_scaling(chunk.text_scale);
                    cur_text_scale = chunk.text_scale;
                }

                if cur_font_name != chunk.pdf_font || cur_font_size != chunk.font_size {
                    content.set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size);
                    cur_font_name.clone_from(&chunk.pdf_font);
                    cur_font_size = chunk.font_size;
                }

                if let Some(mb) = &chunk.math {
                    // Each glyph by its subset id at its own place; the font is
                    // set at 1pt and each glyph's size goes in its matrix. The
                    // rules join the line's decorations.
                    if cur_font_size != 1.0 {
                        content.set_font(Name(chunk.pdf_font.as_bytes()), 1.0);
                        cur_font_size = 1.0;
                    }
                    // m:nor text switches to its own font, glyphs back to the
                    // math font.
                    let mut font = chunk.pdf_font.as_str();
                    for item in &mb.items {
                        let (gx, gy, size, color, item_font, bytes) = match item {
                            crate::math::MathItem::Rule { x: rx, y: ry, w, h } => {
                                decorations.push((x + rx, cy + ry, *w, *h, effective_color));
                                continue;
                            }
                            crate::math::MathItem::Glyph {
                                gid,
                                x,
                                y,
                                size,
                                color,
                                ..
                            } => {
                                let bytes = gid.to_be_bytes().to_vec();
                                (*x, *y, *size, *color, chunk.pdf_font.as_str(), bytes)
                            }
                            crate::math::MathItem::Text {
                                text,
                                font,
                                x,
                                y,
                                size,
                                color,
                            } => {
                                let bytes = encode_text_for_pdf(text, font, &pdf_name_to_entry);
                                (*x, *y, *size, *color, font.as_str(), bytes)
                            }
                        };
                        let color = color.or(effective_color);
                        if color != current_color {
                            fill_color_or_black(content, color);
                            current_color = color;
                        }
                        if item_font != font {
                            content.set_font(Name(item_font.as_bytes()), 1.0);
                            font = item_font;
                        }
                        (td_x, td_y) = (x + gx, cy + gy);
                        content.set_text_matrix([size, 0.0, 0.0, size, td_x, td_y]);
                        content.show(Str(&bytes));
                    }
                    if font != chunk.pdf_font {
                        content.set_font(Name(chunk.pdf_font.as_bytes()), 1.0);
                    }
                    if let Some(entry) = primary_entry.filter(|_| boundary_space) {
                        let size = chunk.font_size;
                        (td_x, td_y) = (x + mb.width, cy);
                        content.set_text_matrix([size, 0.0, 0.0, size, td_x, td_y]);
                        content.show(Str(&entry.encode(" ")));
                    }
                    // Unscaled again, so the next chunk's relative move is in points.
                    content.set_text_matrix([1.0, 0.0, 0.0, 1.0, td_x, td_y]);
                    cur_shear = 0.0;
                    continue;
                }

                // Legacy w:shadow/emboss/imprint: draw an offset gray copy of the
                // glyphs behind the main text, then fall through to draw the real
                // glyphs on top. ponytail: single offset, no highlight pass, and
                // CJK-fallback chars use the primary font in the shadow copy.
                // Word draws italic in a face that has none (macOS Comic Sans
                // MS, Tahoma) sheared 87/256 about the baseline:
                // greek_history_lecture_press_release, door_air_cooling_unit_spec.
                // A Td under a sheared line matrix would shift x by shear × dy,
                // so every move in or out of a sheared chunk is an absolute Tm.
                let shear = if chunk.synthetic_italic {
                    87.0 / 256.0
                } else {
                    0.0
                };
                // Emboss/imprint move the face off the pen position, which
                // keeps the gray copy (see TextShadow::relief).
                let (fx, fy) = match chunk.text_shadow {
                    Some(ref sh) if sh.relief => (x - sh.offset_x, cy - sh.offset_y),
                    _ => (x, cy),
                };
                if let Some(ref sh) = chunk.text_shadow {
                    let (sx, sy) = (fx + sh.offset_x, fy + sh.offset_y);
                    let bytes =
                        encode_text_for_pdf(&chunk.text, &chunk.pdf_font, &pdf_name_to_entry);
                    let (ex, ey) = (fx - sh.offset_x, fy - sh.offset_y);
                    let draw_shadow = |content: &mut Content| {
                        content.set_text_matrix([1.0, 0.0, shear, 1.0, sx, sy]);
                        fill_color_or_black(content, Some(sh.color));
                        content.show(Str(&bytes));
                        if sh.relief {
                            content.set_text_matrix([1.0, 0.0, shear, 1.0, ex, ey]);
                            fill_color_or_black(content, None);
                            content.show(Str(&bytes));
                        }
                        fill_color_or_black(content, current_color);
                    };
                    // The gray copy is an artifact, so the word is read,
                    // searched and copied once.
                    (td_x, td_y, cur_shear) = match link_tags.as_mut() {
                        Some(lt) => {
                            lt.text_artifact(content, draw_shadow);
                            (0.0, 0.0, 0.0)
                        }
                        None => {
                            draw_shadow(content);
                            if sh.relief {
                                (ex, ey, shear)
                            } else {
                                (sx, sy, shear)
                            }
                        }
                    };
                }
                let mut move_to = |content: &mut Content, mx: f32, my: f32| {
                    if shear != 0.0 || cur_shear != 0.0 {
                        content.set_text_matrix([1.0, 0.0, shear, 1.0, mx, my]);
                        cur_shear = shear;
                    } else {
                        content.next_line(mx - td_x, my - td_y);
                    }
                    td_x = mx;
                    td_y = my;
                };

                move_to(content, fx, fy);

                // Per-character font fallback: if some chars are missing
                // from the primary font, split into segments and render
                // missing chars with the CJK fallback font.
                let has_missing = primary_entry.is_some_and(|e| !e.missing_chars.is_empty());
                let fallback_entry = has_missing
                    .then(|| seen_fonts.get("__glyph_fallback"))
                    .flatten();

                if let Some(extras) = mixed
                    .as_ref()
                    .and_then(|m| m.char_extras(chunk_idx, &chunk.text))
                {
                    // Only the East Asian gaps widen: one Tc per segment, the
                    // fallback font where the primary lacks the glyphs.
                    let missing = |c: char| primary_entry.is_some_and(|p| p.missing_chars.contains(&c));
                    let fallback_gids = fallback_entry.and_then(|f| f.char_to_gid.as_ref());
                    let chars: Vec<char> = chunk.text.chars().collect();
                    let mut i = 0;
                    while i < chars.len() {
                        let in_fallback = missing(chars[i]);
                        let mut j = i + 1;
                        while j < chars.len()
                            && extras[j] == extras[i]
                            && missing(chars[j]) == in_fallback
                        {
                            j += 1;
                        }
                        let cs = chunk.char_spacing + extras[i];
                        if cs != cur_char_spacing {
                            content.set_char_spacing(cs);
                            cur_char_spacing = cs;
                        }
                        let seg: String = chars[i..j].iter().collect();
                        match (in_fallback, fallback_entry, fallback_gids) {
                            (true, Some(fb), Some(map)) => {
                                content.set_font(Name(fb.pdf_name.as_bytes()), chunk.font_size);
                                content.show(Str(&encode_as_gids(&seg, map)));
                                content.set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size);
                            }
                            _ => {
                                content.show(Str(&encode_text_for_pdf(
                                    &seg,
                                    &chunk.pdf_font,
                                    &pdf_name_to_entry,
                                )));
                            }
                        }
                        i = j;
                    }
                    if boundary_space {
                        if effective_cs != cur_char_spacing {
                            content.set_char_spacing(effective_cs);
                            cur_char_spacing = effective_cs;
                        }
                        content.show(Str(&encode_text_for_pdf(
                            " ",
                            &chunk.pdf_font,
                            &pdf_name_to_entry,
                        )));
                    }
                } else if let Some(entry) = primary_entry.filter(|_| chunk.rtl_glyphs) {
                    // ponytail: no CJK-fallback glyphs or kerning in right-to-left chunks
                    show_rtl(
                        content,
                        entry,
                        &chunk.text,
                        chunk.font_size,
                        effective_cs,
                        boundary_space,
                    );
                } else if let (Some(primary), Some(fallback)) = (primary_entry, fallback_entry) {
                    let fallback_gids = fallback.char_to_gid.as_ref();
                    // A segment of the chunk in its own font, or in the fallback
                    // font when its characters are missing from the primary one.
                    let show_seg = |content: &mut Content, seg: &[char], in_fallback: bool| {
                        let seg: String = seg.iter().collect();
                        if in_fallback {
                            if let Some(map) = fallback_gids {
                                content
                                    .set_font(Name(fallback.pdf_name.as_bytes()), chunk.font_size);
                                content.show(Str(&encode_as_gids(&seg, map)));
                                content.set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size);
                            }
                        } else {
                            let bytes =
                                encode_text_for_pdf(&seg, &chunk.pdf_font, &pdf_name_to_entry);
                            content.show(Str(&bytes));
                        }
                    };
                    // Split text into runs of primary vs fallback chars
                    let mut seg_start = 0;
                    let mut in_fallback = false;
                    let chars: Vec<char> = chunk.text.chars().collect();
                    for (i, &ch) in chars.iter().enumerate() {
                        let needs_fb = primary.missing_chars.contains(&ch);
                        if i == 0 {
                            in_fallback = needs_fb;
                        } else if needs_fb != in_fallback {
                            show_seg(content, &chars[seg_start..i], in_fallback);
                            seg_start = i;
                            in_fallback = needs_fb;
                        }
                    }
                    show_seg(content, &chars[seg_start..], in_fallback);
                    if boundary_space {
                        content.show(Str(&encode_text_for_pdf(
                            " ",
                            &chunk.pdf_font,
                            &pdf_name_to_entry,
                        )));
                    }
                } else if let Some(entry) = pdf_name_to_entry
                    .get(chunk.pdf_font.as_str())
                    .filter(|e| chunk.kern && e.kern_pairs.is_some())
                {
                    show_kerned(content, entry, &chunk.text, boundary_space);
                } else {
                    let mut text_bytes =
                        encode_text_for_pdf(&chunk.text, &chunk.pdf_font, &pdf_name_to_entry);
                    if boundary_space {
                        text_bytes.extend(encode_text_for_pdf(
                            " ",
                            &chunk.pdf_font,
                            &pdf_name_to_entry,
                        ));
                    }
                    content.show(Str(&text_bytes));
                };

                let thick = (chunk.font_size * 0.05).max(0.5);
                // Word draws a single underline from the font's post table, its
                // offset below the baseline and its thickness each on a 0.25pt
                // grid (online export probe: Calibri, Times New Roman, Arial,
                // Aptos at 10/12/22pt). Double underlines keep the old geometry.
                let post_ul = (!chunk.double_underline)
                    .then(|| pdf_name_to_entry.get(chunk.pdf_font.as_str()))
                    .flatten()
                    .and_then(|e| e.underline)
                    .map(|(pos, th)| {
                        let q = |v: f32| (v * 4.0).round() / 4.0;
                        (q(pos * chunk.font_size), q(th * chunk.font_size).max(0.25))
                    });
                if let (true, Some((offset, ul_thick))) = (chunk.underline, post_ul) {
                    let bottom = y - offset - ul_thick;
                    push_decoration(
                        &mut decorations,
                        x,
                        bottom,
                        chunk.width,
                        ul_thick,
                        chunk.color,
                    );
                } else if chunk.underline {
                    let ul_y = if chunk.hyperlink_url.is_some() {
                        y - chunk.font_size * 0.08
                    } else {
                        y - chunk.font_size * 0.12
                    };
                    let ul_top = ul_y - thick;
                    push_decoration(&mut decorations, x, ul_top, chunk.width, thick, chunk.color);
                    if chunk.double_underline {
                        let gap = (thick * 1.5).max(1.0);
                        push_decoration(
                            &mut decorations,
                            x,
                            ul_top - thick - gap,
                            chunk.width,
                            thick,
                            chunk.color,
                        );
                    }
                }
                if chunk.strikethrough {
                    let strike = pdf_name_to_entry
                        .get(chunk.pdf_font.as_str())
                        .and_then(|e| os2_strike(e, chunk.font_size))
                        .map(|(top, th)| (y + top - th, th));
                    let (st_y, st_thick) = strike.unwrap_or((y + chunk.font_size * 0.3, thick));
                    // One line across the spaces of a struck run, like Word's.
                    push_decoration(
                        &mut decorations,
                        x,
                        st_y,
                        chunk.width,
                        st_thick,
                        chunk.color,
                    );
                }
                if chunk.dstrike {
                    let gap = thick * 1.5;
                    let mid_y = y + chunk.font_size * 0.3;
                    decorations.push((x, mid_y - gap / 2.0, chunk.width, thick, chunk.color));
                    decorations.push((x, mid_y + gap / 2.0, chunk.width, thick, chunk.color));
                }

                if let Some(url) = link_url {
                    let bottom = y - chunk.font_size * 0.2;
                    let top = y + chunk.font_size * 0.8;
                    let node = link_tags
                        .as_ref()
                        .and_then(|lt| lt.link.as_ref().map(|&(_, n)| n));
                    let merged = links
                        .last_mut()
                        .filter(|prev| prev.url == url && (prev.rect.y1 - bottom).abs() < 1.0);
                    let link = match merged {
                        Some(prev) => {
                            prev.rect.x2 = x + chunk.width;
                            prev
                        }
                        None => {
                            // A bare "1" says little as the link's description.
                            let text = match note.filter(|_| chunk.hyperlink_url.is_none()) {
                                Some((true, _)) => "Endnote ",
                                Some((false, _)) => "Footnote ",
                                None => "",
                            };
                            links.push(LinkAnnotation {
                                rect: Rect::new(x, bottom, x + chunk.width, top),
                                url: url.to_string(),
                                node,
                                text: text.to_string(),
                            });
                            links.last_mut().unwrap()
                        }
                    };
                    link.text.push_str(&chunk.text);
                    if chunk.space_after {
                        link.text.push(' ');
                    }
                }
            }
            if cur_synthetic_bold {
                content.set_text_rendering_mode(TextRenderingMode::Fill);
                cur_synthetic_bold = false;
            }
            // Text rendering mode persists across BT/ET, so a paragraph that
            // ends in an outlined chunk must reset to Fill before ET — otherwise
            // subsequent paragraphs inherit the stroke and look outlined too.
            if has_text_outline {
                super::wordart::reset_text_outline(content);
                has_text_outline = false;
            }
            if cur_char_spacing != 0.0 {
                content.set_char_spacing(0.0);
                cur_char_spacing = 0.0;
            }
            if cur_text_scale != 100.0 {
                content.set_horizontal_scaling(100.0);
                cur_text_scale = 100.0;
            }
            content.end_text();
            // Pattern fill from a gradient chunk leaves the fill colorspace
            // set to Pattern; switch back to DeviceRGB before drawing
            // decorations or proceeding to the next line.
            if pattern_fill_active {
                fill_color_or_black(content, None);
                current_color = Some([0, 0, 0]);
                pattern_fill_active = false;
            }
        }

        // Draw inline images outside text block. Every inline picture sits on the
        // baseline (see inline_image_line_extra), whatever its height.
        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            if let Some(ref img_name) = chunk.inline_image_name {
                if let Some(lt) = link_tags.as_mut() {
                    lt.begin_picture(
                        content,
                        chunk.inline_image_alt.as_deref(),
                        chunk.inline_image_decorative,
                    );
                }
                let box_x = chunk_abs_x(chunk_idx, chunk);
                let box_bottom = y + chunk.y_offset;

                // The chunk box is the rotated frame's bounding box; draw the picture at
                // its natural size centred in it, turned about that centre like Word.
                let (w, h) = chunk.inline_image_size;
                let x = box_x + (chunk.width - w) / 2.0;
                let img_bottom = box_bottom + (chunk.inline_image_height - h) / 2.0;
                let turned = chunk.inline_image_rotation_deg.abs() > 0.01;
                if turned {
                    content.save_state();
                    super::positioning::push_center_rotation(
                        content,
                        box_x,
                        box_bottom,
                        chunk.width,
                        chunk.inline_image_height,
                        chunk.inline_image_rotation_deg,
                    );
                }

                // Pre-image effects: shadow, glow (rendered before image so they appear behind)
                let chunk_fx = chunk.inline_image_effect_xobjs.as_ref();
                if let Some(ref shadow) = chunk.inline_image_shadow {
                    super::color::draw_image_shadow(
                        content,
                        shadow,
                        x,
                        img_bottom,
                        w,
                        h,
                        chunk_fx.and_then(|fx| fx.shadow.as_deref()),
                    );
                }
                if let Some(ref glow) = chunk.inline_image_glow {
                    super::color::draw_image_glow(
                        content,
                        glow,
                        x,
                        img_bottom,
                        w,
                        h,
                        chunk_fx.and_then(|fx| fx.glow.as_deref()),
                    );
                }

                super::smartart::render_image_with_clip(
                    content,
                    img_name,
                    x,
                    img_bottom,
                    w,
                    h,
                    chunk.inline_image_clip.as_ref(),
                );

                if let Some(sc) = chunk.inline_image_stroke_color {
                    super::smartart::stroke_image_border(
                        content,
                        x,
                        img_bottom,
                        w,
                        h,
                        sc,
                        chunk.inline_image_stroke_width,
                        chunk.inline_image_clip.as_ref(),
                    );
                }
                if turned {
                    content.restore_state();
                }
                if let Some(lt) = link_tags.as_mut() {
                    lt.resume(content);
                }
            }
        }

        for &(dx, dy, dw, dh, dcolor) in &decorations {
            if dcolor != current_color {
                fill_color_or_black(content, dcolor);
                current_color = dcolor;
            }
            content.rect(dx, dy, dw, dh).fill_nonzero();
        }
    }
    if current_color.is_some() {
        content.set_fill_gray(0.0);
    }
    if let Some(lt) = link_tags {
        lt.finish(content);
    }
}

/// Compute the effective font_size, line_h_ratio, and ascender_ratio for a set of runs
/// by picking the run that produces the tallest visual ascent (font_size * ascender_ratio).
/// A tab never raises its line (mandated_reporter: 12pt tabs between 11pt
/// footer text leave Word's footer baseline at the 11pt line), but sizes a
/// line of nothing else.
pub(super) fn tallest_run_metrics(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> (f32, Option<f32>, Option<f32>) {
    tallest_by_ascent(runs.iter().filter(|r| !r.is_tab), seen_fonts)
        .or_else(|| tallest_by_ascent(runs.iter(), seen_fonts))
        .unwrap_or((runs.first().map_or(12.0, |r| r.font_size), None, None))
}

/// `tallest_run_metrics` for a body paragraph, where a run of only spaces or
/// tabs sizes nothing; with nothing left the caller falls back to the mark.
pub(super) fn tallest_glyph_run_metrics(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> (f32, Option<f32>, Option<f32>) {
    tallest_by_ascent(runs.iter().filter(|r| sizes_line(r)), seen_fonts).unwrap_or((
        runs.first().map_or(12.0, |r| r.font_size),
        None,
        None,
    ))
}

/// Distance from the top of an exact or at-least line box to its baseline, or
/// None for an ordinary line. Word bottom-aligns an exact-height box, and an at-least
/// one whose minimum wins (online export: czech_census's 10pt lines under
/// atLeast 12.05 start 0.55pt lower), at the glyphs' descent: line_h_ratio −
/// ascender_ratio, less the East Asian leading Word puts below normal lines.
/// A Latin exact line instead puts its baseline 80% down the box whatever the
/// font (Word probes: Calibri, Times New Roman, Arial and Cambria at 8-24pt).
pub(super) fn boxed_line_ascent(
    ls: LineSpacing,
    line_h: f32,
    font_size: f32,
    lhr: Option<f32>,
    ar: Option<f32>,
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> Option<f32> {
    if matches!(ls, LineSpacing::Exact(_))
        && tallest_glyph_run_half_leading(runs, seen_fonts) == 0.0
    {
        return Some(line_h * 0.8);
    }
    let (lhr, ar) = (lhr?, ar?);
    let bottom_aligned = match ls {
        // East Asian: a Latin exact line returned above.
        LineSpacing::Exact(_) => true,
        LineSpacing::AtLeast(min) => min > font_size * lhr,
        LineSpacing::Auto(_) => false,
    };
    (lhr > ar && bottom_aligned).then(|| {
        let half_lead = tallest_glyph_run_half_leading(runs, seen_fonts);
        line_h - font_size * (lhr - ar - half_lead)
    })
}

/// Half the East Asian 1.3× leading, per em, of the tallest glyph run: the part
/// an exact or at-least box keeps below the glyphs (0 for other fonts).
pub(super) fn tallest_glyph_run_half_leading(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> f32 {
    let mut best = (0.0f32, 0.0f32);
    let mut key_buf = String::new();
    for run in runs.iter().filter(|r| sizes_line(r) && !r.is_line_break) {
        let Some(entry) = seen_fonts.get(font_key_buf(run, &mut key_buf)) else {
            continue;
        };
        let (lhr, ar) = run_line_metrics(entry, &run.text);
        let ascent = run.font_size * ar.unwrap_or(0.75);
        if ascent > best.0 {
            let cjk_box = entry.east_asian && ar == entry.ascender_ratio;
            best = (
                ascent,
                if cjk_box {
                    lhr.unwrap_or(0.0) * 0.3 / 2.6
                } else {
                    0.0
                },
            );
        }
    }
    best.1
}

/// (font_size, line_h_ratio, ascender_ratio) of the run with the tallest ascent
/// (font_size × ascender ratio); None when no run has one.
fn tallest_by_ascent<'a>(
    runs: impl Iterator<Item = &'a Run>,
    seen_fonts: &HashMap<String, FontEntry>,
) -> Option<(f32, Option<f32>, Option<f32>)> {
    let mut best: Option<(f32, Option<f32>, Option<f32>)> = None;
    let mut best_ascent = 0.0f32;
    let mut key_buf = String::new();
    for run in runs {
        // Line-break runs only affect the empty line they create, not
        // the paragraph's overall line height.
        if run.is_line_break {
            continue;
        }
        let entry = seen_fonts.get(font_key_buf(run, &mut key_buf));
        let (lhr, ascender_ratio) = if let Some((_, mb)) = math_box(run, seen_fonts) {
            // A laid-out zone counts as text in the math font's hhea box
            // (`MathBox::line_ascent`); its taller ink is `math_line_stretch`'s.
            let fs = run.font_size.max(f32::EPSILON);
            (
                Some((mb.line_ascent + mb.line_descent) / fs),
                Some(mb.line_ascent / fs),
            )
        } else {
            entry.map_or((None, None), |e| run_line_metrics(e, &run.text))
        };
        let ascent = run.font_size * ascender_ratio.unwrap_or(0.75);
        if ascent > best_ascent {
            best_ascent = ascent;
            best = Some((run.font_size, lhr, ascender_ratio));
        }
    }
    best
}

/// Height below the picture on a picture line, for `inline_line_advance`: the
/// descent of the tallest run with visible glyphs, plus the extra leading that
/// multiple line spacing adds to the tallest non-picture run's font (Word puts
/// that leading above the *next* line; we carry it here). The picture run's own
/// font size never counts. Word measured: 36pt signature beside 10pt Arial,
/// single spacing → +2.2 (italian_evaluation_minutes p7); 149pt logo after a
/// 14pt tab at 1.15 → +2.6 (english_town_council p1); 177pt picture after a run
/// of spaces at 1.5 → +7.6 (family_kinship p4); four 48pt chord diagrams alone
/// → +0 (old_blue_truck p1).
pub(super) fn picture_line_bottom(
    runs: &[Run],
    para: &Paragraph,
    seen_fonts: &HashMap<String, FontEntry>,
    ls: LineSpacing,
) -> f32 {
    let text_runs = || {
        runs.iter()
            .filter(|r| r.inline_image.is_none() && !r.vanish)
    };
    // With no text run the mark's font sets the leading (czech_village's
    // header logo under 1.5 lines of 12pt Times New Roman).
    let mark = || {
        para.paragraph_mark_font_size.map(|fs| {
            let lhr = para
                .paragraph_mark_font_name
                .as_deref()
                .and_then(|n| seen_fonts.get(n))
                .and_then(|e| run_line_metrics(e, "").0);
            (fs, lhr, None)
        })
    };
    let leading = tallest_by_ascent(text_runs(), seen_fonts)
        .or_else(mark)
        .map_or(0.0, |(fs, lhr, _)| {
            super::helpers::spacing_leading(super::helpers::resolve_line_h(ls, fs, lhr), fs, lhr)
        });
    let glyph_runs = text_runs().filter(|r| !r.text.trim().is_empty());
    let descent = tallest_by_ascent(glyph_runs, seen_fonts)
        .map_or(0.0, |(fs, lhr, ar)| fs * descender_ratio(lhr, ar));
    descent + leading
}

/// A run's (line_h_ratio, ascender_ratio). A run of nothing but spaces in an
/// East Asian font gets the plain metrics rather than Word's 1.3× leading, so
/// it cannot raise a Latin line (see `fonts::embed::compute_line_metrics`).
/// Empty text — a tab, a paragraph mark, the blank line after a break — keeps
/// the font's real metrics: those lines are sized by the East Asian font itself.
pub(super) fn run_line_metrics(entry: &FontEntry, text: &str) -> (Option<f32>, Option<f32>) {
    if east_asian_leading(entry, text) {
        (entry.line_h_ratio, entry.ascender_ratio)
    } else {
        (entry.plain_line_h_ratio, entry.plain_ascender_ratio)
    }
}

/// Whether a run of `text` in `entry` gets Word's East Asian 1.3× leading.
pub(super) fn east_asian_leading(entry: &FontEntry, text: &str) -> bool {
    entry.east_asian && (text.is_empty() || !text.chars().all(is_break_space))
}

/// A run whose glyphs can size its line: not a tab, and empty (a paragraph
/// mark) or holding more than whitespace.
fn sizes_line(r: &Run) -> bool {
    !r.is_tab && (r.text.is_empty() || !r.text.trim().is_empty())
}

/// Grid-snapped line height: Word counts docGrid cells with each font's
/// `grid_line_ratio` (sTypo for Latin fonts, the 1.3× height for East Asian
/// ones). Falls back to `line_h` when no run provides one.
pub(super) fn grid_snapped_line_h(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    effective_ls: crate::model::LineSpacing,
    line_h: f32,
    pitch: f32,
) -> f32 {
    let mut grid_h = 0.0f32;
    let mut key_buf = String::new();
    for run in runs {
        if run.is_line_break {
            continue;
        }
        if let Some(e) = seen_fonts.get(font_key_buf(run, &mut key_buf))
            && let Some(t) = e.grid_line_ratio
        {
            grid_h = grid_h.max(effective_font_size(run, e) * t);
        }
    }
    // Tolerance so an exact fit stays one cell despite f32 error.
    let cells = |h: f32| ((h / pitch) - 0.02).ceil().max(1.0) * pitch;
    // A line-spacing multiple is a floor of m unsnapped pitches under the cells
    // the glyphs need: 1.5 lines of one 18pt cell is 27pt (case79), while text
    // needing two 15.6pt cells stays 31.2pt at 1.25, 1.5 or 2 lines (Word probes).
    match effective_ls {
        crate::model::LineSpacing::Auto(m) if grid_h > 0.0 => cells(grid_h).max(m * pitch),
        _ => cells(line_h),
    }
}

/// Baseline offset of a grid-snapped line `cell_h` tall: the cell's centre plus
/// the largest run's `grid_baseline_shift`.
pub(super) fn grid_baseline_offset(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    cell_h: f32,
) -> Option<f32> {
    let mut key_buf = String::new();
    runs.iter()
        .filter(|r| r.inline_image.is_none() && !r.vanish && !r.is_line_break && sizes_line(r))
        .filter_map(|r| {
            let e = seen_fonts.get(font_key_buf(r, &mut key_buf))?;
            Some(e.grid_baseline_shift? * effective_font_size(r, e))
        })
        .reduce(f32::max)
        .map(|shift| cell_h / 2.0 + shift)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn top_aligned_drop_spends_external_leading_first() {
        // Times New Roman's leading is 87/2048 em; Calibri has none.
        let tnr = FontEntry { leading_ratio: Some(87.0 / 2048.0), ..FontEntry::stub() };
        let calibri = FontEntry { leading_ratio: Some(0.0), ..FontEntry::stub() };
        assert!(top_aligned_drop(&tnr, 11.0, -0.5) < 0.05);
        assert!((top_aligned_drop(&tnr, 24.0, -7.0) - 5.98).abs() < 0.01);
        assert_eq!(top_aligned_drop(&calibri, 11.0, -0.5), 0.5);
        assert_eq!(top_aligned_drop(&calibri, 11.0, 7.0), 0.0);
    }

    #[test]
    fn leading_space_wrap_uses_space_font_pitch() {
        let mut lines = vec![
            TextLine {
                leading_space_wrap: true,
                break_font_size: Some(7.5),
                break_lhr: Some(1.2),
                ..Default::default()
            },
            TextLine {
                break_font_size: Some(7.5),
                break_lhr: Some(1.2),
                ..Default::default()
            },
        ];
        size_lines_by_own_runs(
            &mut lines,
            &HashMap::new(),
            LineSpacing::Auto(1.0),
            14.0,
            10.0,
            false,
        );
        assert!((lines[0].pitch.unwrap() - 9.0).abs() < 0.001);
        assert!(
            lines[1].pitch.is_none(),
            "ordinary empty lines keep paragraph pitch"
        );
    }

    fn plain_layout(squeeze_spaces: bool) -> CjkLayout {
        CjkLayout {
            punct_half_out: false,
            auto_space: false,
            compress_punct: false,
            punct_pull: false,
            punct_cap: PUNCT_SQUEEZE,
            punct_pull_free_ratio: 0.0,
            punct_pull_free_em: 0.0,
            squeeze_spaces,
            squeeze_line: false,
            expand_shift_return: true,
            bidi: false,
        }
    }
    use crate::model::VertAlign;

    #[test]
    fn overwide_word_cuts_at_the_last_fitting_char() {
        let width = |w: &str| w.chars().count() as f32 * 3.0;
        assert_eq!(fitting_prefix_len("..........", 10.0, width), Some(3));
        assert_eq!(fitting_prefix_len("....", 1.0, width), Some(1));
        assert_eq!(fitting_prefix_len("é…é", 6.0, width), Some("é…".len()));
        assert_eq!(fitting_prefix_len(".", 1.0, width), None);
    }

    #[test]
    fn underlining_word_spaces_preserves_squeeze_decisions() {
        let fonts = HashMap::from([("Arial".to_string(), FontEntry::stub())]);
        let build = |underline, width| {
            let run = Run {
                text: "aa aa aa aa aa".into(),
                underline,
                ..make_run(10.0, VertAlign::Baseline, false)
            };
            build_paragraph_lines(
                &[run],
                &fonts,
                width,
                0.0,
                &HashMap::new(),
                &HashMap::new(),
                None,
                None,
                None,
                CjkLayout {
                    auto_space: false,
                    compress_punct: false,
                    punct_pull: false,
                    punct_cap: PUNCT_SQUEEZE,
                    punct_pull_free_ratio: 0.0,
                    punct_pull_free_em: 0.0,
                    squeeze_spaces: true,
                    punct_half_out: false,
                    expand_shift_return: true,
                    squeeze_line: false,
                    bidi: false,
                },
            )
            .len()
        };
        for width in [45.0, 50.0, 55.0, 60.0, 65.0, 68.0, 70.0] {
            assert_eq!(build(false, width), build(true, width), "width={width}");
        }
        let entry = FontEntry::stub();
        let tab = WordChunk::tab_underline(&entry, 10.0, None, false, None, 0.0, 30.0);
        let space = WordChunk::word_space_underline(&entry, 10.0, None, false, None, 30.0, 3.0);
        assert_eq!(line_space_width(&[tab, space]), 3.0);
    }

    #[test]
    fn justify_counts_word_spaces_not_run_joins() {
        let entry = FontEntry::stub();
        let chunk = |text: &str, x: f32, w: f32| {
            let mut c = WordChunk::tab_underline(&entry, 12.0, None, false, None, x, w);
            c.text = text.to_string();
            c
        };
        // "ab" split across two runs, a gap, then "cd", an underlined space chunk, "ef"
        let chunks = vec![
            chunk("a", 0.0, 6.0),
            chunk("b", 6.0, 6.0),
            chunk("cd", 15.0, 12.0),
            chunk("", 27.0, 3.0),
            chunk("ef", 30.0, 12.0),
        ];
        assert_eq!(spaces_before_each(&chunks), vec![0, 0, 1, 1, 2]);
    }

    #[test]
    fn japanese_adjacent_marks_are_counted() {
        assert_eq!(ja_adjacent_marks("大学院）・専門"), 1);
        assert_eq!(ja_adjacent_marks("高卒・高校"), 0);
        assert_eq!(ja_adjacent_marks("）（"), 1);
    }

    #[test]
    fn hinted_quotes_get_their_own_segments() {
        assert_eq!(punct_segments("在“测”船", true), vec!["在", "“测”", "船"]);
        assert_eq!(punct_segments("在“测”船", false), vec!["在“测”船"]);
        assert_eq!(punct_segments("“测试", true), vec!["“测试"]);
        assert_eq!(punct_segments("试”", true), vec!["试”"]);
        assert_eq!(punct_segments("2021《进出", false), vec!["2021", "《进出"]);
    }

    #[test]
    fn compress_punctuation_trims_marks_evenly_and_shifts_followers() {
        let entry = FontEntry::stub();
        let chunk = |text: &str, x: f32| {
            let mut c = WordChunk::tab_underline(&entry, 16.0, None, false, None, x, 32.0);
            c.text = text.to_string();
            c
        };
        let mut chunks = vec![
            chunk("任，", 0.0),
            chunk("負", 32.0),
            chunk("事、", 64.0),
            chunk("公", 96.0),
        ];
        assert!(compress_punctuation(&mut chunks, 4.0, PUNCT_SQUEEZE));
        // 2pt off each mark; everything after a mark slides left
        assert!((chunks[0].width - 30.0).abs() < 1e-4 && (chunks[2].width - 30.0).abs() < 1e-4);
        assert!((chunks[1].x_offset - 30.0).abs() < 1e-4);
        assert!((chunks[2].x_offset - 62.0).abs() < 1e-4);
        assert!((chunks[3].x_offset - 92.0).abs() < 1e-4);
        // The next word raises the same level on every mark: 2.5pt more each.
        assert!(compress_punctuation(&mut chunks, 5.0, PUNCT_SQUEEZE));
        assert!((chunks[0].width - 27.5).abs() < 1e-4 && (chunks[2].width - 27.5).abs() < 1e-4);
        assert!((chunks[3].x_offset - 87.0).abs() < 1e-4);
        // Past PUNCT_SQUEEZE of the em (8pt at 16pt) the word wraps instead:
        // 8 more would need 8.5 per mark, so nothing moves.
        assert!(!compress_punctuation(&mut chunks, 8.0, PUNCT_SQUEEZE));
        assert!((chunks[0].width - 27.5).abs() < 1e-4);
        let mut plain = vec![chunk("任負", 0.0)];
        assert!(!compress_punctuation(&mut plain, 1.0, PUNCT_SQUEEZE));
    }

    #[test]
    fn test_char_justify_gaps() {
        // Latin justify spreads word gaps, not characters.
        assert_eq!(char_justify_gaps(Alignment::Justify, true, false, 10), None);
        // CJK justify keeps the grid's trailing cell gap: divide by char count.
        assert_eq!(
            char_justify_gaps(Alignment::Justify, true, true, 10),
            Some(10)
        );
        // distribute ends flush at the right margin: one gap fewer.
        assert_eq!(
            char_justify_gaps(Alignment::Distribute, true, false, 10),
            Some(9)
        );
        // distribute ends flush at both margins whatever the script.
        assert_eq!(
            char_justify_gaps(Alignment::Distribute, true, true, 10),
            Some(9)
        );
        // A single character has no gap to spread into (guards a 0 divisor).
        assert_eq!(
            char_justify_gaps(Alignment::Distribute, true, false, 1),
            None
        );
        // Last line of a justified paragraph, and non-stretching alignments.
        assert_eq!(char_justify_gaps(Alignment::Justify, false, true, 10), None);
        assert_eq!(char_justify_gaps(Alignment::Left, false, true, 10), None);
    }

    #[test]
    fn url_wraps_only_after_hyphens() {
        let words: Vec<&str> =
            split_preserving_spaces("see https://example.com/foo/bar?x=1&y=2#frag after")
                .into_iter()
                .map(|(_, w)| w)
                .collect();
        assert_eq!(
            words,
            vec!["see", "https://example.com/foo/bar?x=1&y=2#frag", "after"]
        );
        let words: Vec<&str> = split_preserving_spaces("www.gov.hr/pristup-informacijama/ x")
            .into_iter()
            .map(|(_, w)| w)
            .collect();
        assert_eq!(words, vec!["www.gov.hr/pristup-", "informacijama/", "x"]);
    }

    #[test]
    fn test_no_break_after_ellipsis_inside_token() {
        // TOC dot-leaders typed as ellipses: "Preparation………45" is one
        // unbreakable token in Word, but UAX #14 allows IN→NU breaks.
        let words: Vec<&str> = split_preserving_spaces("Preparation………45 done")
            .into_iter()
            .map(|(_, w)| w)
            .collect();
        assert_eq!(words, vec!["Preparation………45", "done"]);
        // A break after ellipsis before whitespace is still fine.
        let words: Vec<&str> = split_preserving_spaces("wait… go")
            .into_iter()
            .map(|(_, w)| w)
            .collect();
        assert_eq!(words, vec!["wait…", "go"]);
    }

    #[test]
    fn test_is_break_space() {
        assert!(is_break_space(' '));
        assert!(is_break_space('\t'));
        assert!(is_break_space('\n'));
        // Non-breaking space is NOT a break space
        assert!(!is_break_space('\u{00a0}'));
        // Ideographic space is NOT a break space
        assert!(!is_break_space('\u{3000}'));
        // Regular chars
        assert!(!is_break_space('a'));
        assert!(!is_break_space('1'));
    }

    fn make_run(font_size: f32, valign: VertAlign, small_caps: bool) -> Run {
        Run {
            font_name: "Arial".to_string(),
            font_size,
            vertical_align: valign,
            small_caps,
            text_scale: 100.0,
            ..Run::default()
        }
    }

    #[test]
    fn breaks_between_matches_splitting_inside_a_run() {
        let chars = [
            'a', 'Z', '-', '/', ',', '.', '(', ')', '…', '⁞', '中', '文', '。', '「', 'é', '1',
            '%', '$',
        ];
        for a in chars {
            for b in chars {
                let split = split_preserving_spaces(&format!("{a}{b}")).len() > 1;
                assert_eq!(breaks_between(a, b), split, "{a:?}{b:?}");
            }
        }
    }

    #[test]
    fn period_and_opening_parenthesis_stay_in_one_token() {
        assert_eq!(
            split_preserving_spaces("2023г.(note)"),
            vec![(0, "2023г.(note)")]
        );
        assert_eq!(
            split_preserving_spaces("2023г. (note)"),
            vec![(0, "2023г."), (1, "(note)")]
        );
        assert!(!breaks_between('.', '('));
    }

    #[test]
    fn hyphen_runs_break_between_hyphens() {
        assert_eq!(
            split_preserving_spaces("a---b"),
            vec![(0, "a-"), (0, "-"), (0, "-"), (0, "b")]
        );
    }

    #[test]
    fn hyphen_can_break_before_opening_guillemet() {
        assert_eq!(
            split_preserving_spaces("«North Ural»-«Branch»"),
            vec![(0, "«North"), (1, "Ural»-"), (0, "«Branch»")]
        );
        assert!(breaks_between('-', '«'));
        assert!(!breaks_between('\u{2011}', '«'));
    }

    #[test]
    fn hyphen_breaks_before_a_digit() {
        assert_eq!(
            split_preserving_spaces("2019-2024"),
            vec![(0, "2019-"), (0, "2024")]
        );
    }

    #[test]
    fn nbsp_space_sequences_keep_words_together_across_runs() {
        let fonts = HashMap::from([("Arial".to_string(), FontEntry::stub())]);
        let run = |s: &str| Run {
            text: s.into(),
            ..make_run(10.0, VertAlign::Baseline, false)
        };
        for pieces in [
            vec!["aaaa bb\u{00a0} cc dd"],
            vec!["aaaa ", "bb\u{00a0}", " ", "cc dd"],
            vec!["aaaa ", "bb\u{00a0}", "  ", "cc dd"],
            vec!["aaaa ", "bb ", "\u{00a0}", "cc dd"],
        ] {
            let runs: Vec<_> = pieces.iter().map(|s| run(s)).collect();
            let lines = build_paragraph_lines(
                &runs,
                &fonts,
                50.0,
                0.0,
                &HashMap::new(),
                &HashMap::new(),
                None,
                None,
                None,
                plain_layout(false),
            );
            assert_eq!(lines.len(), 2, "{pieces:?}");
            assert_eq!(
                lines[0]
                    .chunks
                    .iter()
                    .map(|c| c.text.as_str())
                    .collect::<String>(),
                "aaaa",
                "{pieces:?}"
            );
            let second = lines[1]
                .chunks
                .iter()
                .map(|c| c.text.as_str())
                .collect::<String>();
            assert!(
                second.contains("bb") && second.contains("cc"),
                "{pieces:?}: {second}"
            );
            let expected_spaces = if pieces.iter().any(|s| *s == "  ") {
                3
            } else {
                2
            };
            assert_eq!(
                spaces_before_each(&lines[1].chunks).last().copied(),
                Some(expected_spaces)
            );
        }
        assert_eq!(
            split_preserving_spaces("bb  cc"),
            vec![(0, "bb"), (2, "cc")]
        );
    }

    #[test]
    fn abbreviation_before_numero_sign_stays_together() {
        assert_eq!(split_preserving_spaces("ст.№4"), vec![(0, "ст.№4")]);
        assert_eq!(
            split_preserving_spaces("st.№ 4"),
            vec![(0, "st.№"), (1, "4")]
        );
        assert_eq!(
            split_preserving_spaces("st. №4"),
            vec![(0, "st."), (1, "№4")]
        );
        assert!(!breaks_between('.', '№'));
        assert!(breaks_between(' ', '№'));
    }

    #[test]
    fn squeeze_keeps_words_to_word_probe_thresholds() {
        for (word_w, lead, keeps_to, wraps_at) in [
            (24.0, 3.0, 9.39, 9.5),
            (12.0, 3.0, 5.19, 5.3),
            (24.0, 0.0, 7.99, 8.1),
            (12.0, 0.0, 3.99, 4.1),
        ] {
            assert!(squeeze_keeps(keeps_to, word_w, lead), "{word_w} {lead}");
            assert!(!squeeze_keeps(wraps_at, word_w, lead), "{word_w} {lead}");
        }
    }

    #[test]
    fn segment_after_hyphen_squeezes_like_a_word_after_a_space() {
        let fonts = HashMap::from([("Arial".to_string(), FontEntry::stub())]);
        let run = |s: &str| Run {
            text: s.into(),
            ..make_run(10.0, VertAlign::Baseline, false)
        };
        for suffix in ["PP.", "x", "12-", "05215", "AB2C2", "12345/2026", "and/or"] {
            let text = format!("aa aa aa aa 11-{suffix}");
            for runs in [vec![run(&text)], vec![run("aa aa aa aa 11-"), run(suffix)]] {
                let build = |overflow, squeeze| {
                    build_paragraph_lines(
                        &runs,
                        &fonts,
                        text.len() as f32 * 5.0 - overflow,
                        0.0,
                        &HashMap::new(),
                        &HashMap::new(),
                        None,
                        None,
                        None,
                        plain_layout(squeeze),
                    )
                };
                let lines = build(1.0, true);
                assert_eq!(lines.len(), 1, "{suffix}");
                assert!(lines[0].squeezed);
                assert_eq!(build(1.0, false).len(), 2, "{suffix}");
                assert_eq!(build(6.0, true).len(), 2, "{suffix}");
            }
        }
    }

    #[test]
    fn repeated_spaces_keep_their_justification_weight_across_runs() {
        let fonts = HashMap::from([("Arial".to_string(), FontEntry::stub())]);
        let run = |s: &str| Run {
            text: s.into(),
            ..make_run(10.0, VertAlign::Baseline, false)
        };
        for runs in [
            vec![run("Label  MARK tail")],
            vec![run("Label "), run(" "), run("MARK tail")],
        ] {
            let lines = build_paragraph_lines(
                &runs,
                &fonts,
                200.0,
                0.0,
                &HashMap::new(),
                &HashMap::new(),
                None,
                None,
                None,
                plain_layout(false),
            );
            assert_eq!(spaces_before_each(&lines[0].chunks), vec![0, 2, 3]);
        }
        let lines = build_paragraph_lines(
            &[run("Label  MARK tail")],
            &fonts,
            35.0,
            0.0,
            &HashMap::new(),
            &HashMap::new(),
            None,
            None,
            None,
            plain_layout(false),
        );
        for line in lines {
            assert_eq!(spaces_before_each(&line.chunks)[0], 0);
        }
    }

    #[test]
    fn word_split_over_runs_wraps_as_a_whole() {
        let mut fonts = HashMap::new();
        fonts.insert("Arial".to_string(), FontEntry::stub());
        let text_run = |text: &str| Run {
            text: text.to_string(),
            ..make_run(10.0, VertAlign::Baseline, false)
        };
        let runs = [text_run("aaaa bbb"), text_run("ccc")];
        let cjk = plain_layout(false);
        let lines = build_paragraph_lines(
            &runs,
            &fonts,
            40.0,
            0.0,
            &HashMap::new(),
            &HashMap::new(),
            None,
            None,
            None,
            cjk,
        );
        let texts: Vec<Vec<&str>> = lines
            .iter()
            .map(|l| l.chunks.iter().map(|c| c.text.as_str()).collect())
            .collect();
        assert_eq!(texts, vec![vec!["aaaa"], vec!["bbb", "ccc"]]);
        assert_eq!(lines[1].chunks[0].x_offset, 0.0);
    }

    #[test]
    fn overwide_word_after_text_breaks_at_the_new_lines_margin() {
        let mut fonts = HashMap::new();
        fonts.insert("Arial".to_string(), FontEntry::stub());
        let runs = [Run {
            text: "aa bbbbbbbbbb".to_string(),
            ..make_run(10.0, VertAlign::Baseline, false)
        }];
        let cjk = plain_layout(false);
        let lines = build_paragraph_lines(
            &runs,
            &fonts,
            40.0,
            0.0,
            &HashMap::new(),
            &HashMap::new(),
            None,
            None,
            None,
            cjk,
        );
        let texts: Vec<String> = lines
            .iter()
            .map(|l| l.chunks.iter().map(|c| c.text.as_str()).collect())
            .collect();
        assert_eq!(texts, vec!["aa", "bbbbbbbb", "bb"]);
    }

    #[test]
    fn grid_line_spacing_multiple_scales_the_cells() {
        // 12pt Times New Roman (sTypo 1.06 em) needs one 18pt cell; at 1.5
        // lines Word makes the line 27pt (case79), not two cells.
        let mut fonts = HashMap::new();
        fonts.insert(
            "Arial".to_string(),
            FontEntry {
                grid_line_ratio: Some(1.06),
                ..FontEntry::stub()
            },
        );
        let runs = [Run {
            text: "Hxgp".to_string(),
            ..make_run(12.0, VertAlign::Baseline, false)
        }];
        let h = |ls| grid_snapped_line_h(&runs, &fonts, ls, 13.8, 18.0);
        assert_eq!(h(crate::model::LineSpacing::Auto(1.0)), 18.0);
        assert_eq!(h(crate::model::LineSpacing::Auto(1.5)), 27.0);
        let big = [Run {
            text: "Hxgp".to_string(),
            ..make_run(22.0, VertAlign::Baseline, false)
        }];
        let h2 = |ls| grid_snapped_line_h(&big, &fonts, ls, 25.3, 15.6);
        assert_eq!(h2(crate::model::LineSpacing::Auto(1.5)), 31.2);
        assert_eq!(h2(crate::model::LineSpacing::Auto(3.0)), 15.6 * 3.0);
    }

    #[test]
    fn grid_baseline_centres_the_largest_run_in_its_cell() {
        // MS Gothic (win 0.859 / 0.141) 16pt on two 18pt cells: Word's baseline
        // sits 23.67pt into the line (japanese_interlibrary_loan).
        let mut fonts = HashMap::new();
        fonts.insert(
            "Arial".to_string(),
            FontEntry {
                grid_baseline_shift: Some(0.3594),
                ..FontEntry::stub()
            },
        );
        let text_run = |text: &str, font_size: f32| Run {
            text: text.to_string(),
            ..make_run(font_size, VertAlign::Baseline, false)
        };
        let runs = [
            text_run("a", 10.0),
            text_run("b", 16.0),
            text_run("  ", 30.0),
        ];
        let offset = grid_baseline_offset(&runs, &fonts, 36.0).unwrap();
        assert!((offset - 23.75).abs() < 0.01, "{offset}");
    }

    #[test]
    fn segment_width_counts_each_space_once_across_runs() {
        let mut fonts = HashMap::new();
        fonts.insert("Arial".to_string(), FontEntry::stub());
        let width = |texts: &[&str]| {
            let runs: Vec<Run> = texts
                .iter()
                .map(|t| Run {
                    text: t.to_string(),
                    ..make_run(10.0, VertAlign::Baseline, false)
                })
                .collect();
            let refs: Vec<&Run> = runs.iter().collect();
            segment_width(&refs, &fonts)
        };
        // Every glyph is 5pt: "a b" is three of them however the runs split.
        assert_eq!(width(&["a b"]), 15.0);
        assert_eq!(width(&["a ", "b"]), 15.0);
        assert_eq!(width(&["a", " b"]), 15.0);
        assert_eq!(width(&["a", " ", "b"]), 15.0);
        assert_eq!(width(&["a ", " b"]), 20.0);
        assert_eq!(width(&["a", "b"]), 10.0);
        // Leading and trailing spaces are not the segment's.
        assert_eq!(width(&["  a", "b  "]), 10.0);
    }

    /// Standard-14-style entry with every WinAnsi glyph 500/1000 wide.
    #[test]
    fn list_text_tabs_past_label_and_first_line_indent() {
        let mut fonts = HashMap::new();
        fonts.insert("Arial".into(), FontEntry::stub());
        let mut para = crate::model::Paragraph {
            runs: vec![make_run(10.0, VertAlign::Baseline, false)],
            list_label: "1.".into(),
            indent_first_line: 28.35,
            tab_stops: vec![
                TabStop {
                    position: 21.3,
                    alignment: TabAlignment::Left,
                    leader: None,
                },
                TabStop {
                    position: 42.55,
                    alignment: TabAlignment::Left,
                    leader: None,
                },
            ],
            ..Default::default()
        };
        let hanging =
            |p: &crate::model::Paragraph| super::super::list_label::text_hanging(p, 36.0, &fonts);
        assert!((hanging(&para) + 42.55).abs() < 0.01);
        para.indent_first_line = 0.0;
        para.indent_left = -5.0;
        para.indent_hanging = 0.25;
        para.tab_stops[0].position = 8.55;
        assert!((hanging(&para) + 13.55).abs() < 0.01);
        para.indent_left = 0.0;
        para.indent_hanging = 0.0;
        para.tab_stops.clear();
        assert!((hanging(&para) + 36.0).abs() < 0.01);
        para.indent_left = 36.0;
        para.indent_hanging = 18.0;
        para.tab_stops.push(TabStop {
            position: 20.1,
            alignment: TabAlignment::Left,
            leader: None,
        });
        assert!(hanging(&para).abs() < 0.01);
        para.indent_left = 0.0;
        para.indent_hanging = 0.0;
        para.tab_stops.clear();
        para.list_label = "12345678.".into();
        assert!((hanging(&para) + 72.0).abs() < 0.01);
    }

    #[test]
    fn test_tabbed_line_applies_whitespace_only_run_after_tab() {
        // "<w:tab/>   " at 10pt followed by "x" at 9pt: the size difference keeps
        // the spaces in their own run, and Word still advances over them, so the
        // word starts at tab stop + 3 space widths rather than at the stop itself.
        let text_run = |text: &str, font_size: f32| Run {
            text: text.to_string(),
            ..make_run(font_size, VertAlign::Baseline, false)
        };
        let tab = Run {
            is_tab: true,
            ..make_run(10.0, VertAlign::Baseline, false)
        };
        let runs = [tab, text_run("   ", 10.0), text_run("x", 9.0)];
        let mut fonts = HashMap::new();
        fonts.insert("Arial".to_string(), FontEntry::stub());
        let stops = [TabStop {
            position: 100.0,
            alignment: TabAlignment::Left,
            leader: None,
        }];
        let lines = build_tabbed_line(
            &runs,
            &fonts,
            &stops,
            0.0,
            400.0,
            0.0,
            0.0,
            &HashMap::new(),
            &HashMap::new(),
            36.0,
            &[],
            15,
            plain_layout(false),
        );
        assert_eq!(lines.len(), 1);
        let word = lines[0]
            .chunks
            .iter()
            .find(|c| c.text == "x")
            .expect("word chunk");
        let expected = 100.0 + 3.0 * fonts["Arial"].space_width(10.0);
        assert!(
            (word.x_offset - expected).abs() < 0.01,
            "x_offset {} != {expected}",
            word.x_offset
        );
    }

    #[test]
    fn tabbed_justified_line_squeezes_spaces_after_the_tab() {
        // Stub glyphs are 5pt at 10pt: after the stop at 10, "aa aa aa aa aa"
        // ends at 80, 3pt past a 77pt measure; four 5pt spaces may give 5pt.
        let tab = Run {
            is_tab: true,
            ..make_run(10.0, VertAlign::Baseline, false)
        };
        let text = Run {
            text: "aa aa aa aa aa".to_string(),
            ..make_run(10.0, VertAlign::Baseline, false)
        };
        let runs = [tab, text];
        let mut fonts = HashMap::new();
        fonts.insert("Arial".to_string(), FontEntry::stub());
        let stops = [TabStop {
            position: 10.0,
            alignment: TabAlignment::Left,
            leader: None,
        }];
        let lines = |squeeze| {
            build_tabbed_line(
                &runs,
                &fonts,
                &stops,
                0.0,
                77.0,
                0.0,
                0.0,
                &HashMap::new(),
                &HashMap::new(),
                36.0,
                &[],
                15,
                plain_layout(squeeze),
            )
        };
        let squeezed = lines(true);
        assert_eq!(squeezed.len(), 1);
        assert!(squeezed[0].squeezed);
        assert_eq!(lines(false).len(), 2);
    }

    #[test]
    fn test_effective_font_size_baseline() {
        let run = make_run(12.0, VertAlign::Baseline, false);
        assert_eq!(effective_font_size(&run, &FontEntry::stub()), 12.0);
    }

    #[test]
    fn test_effective_font_size_superscript() {
        // Aptos (OS/2 0.600) 12pt superscripts are 7pt in Word, Times (0.650) 8pt.
        let mut entry = FontEntry::stub();
        entry.superscript_ratio = Some(0.6);
        let run = make_run(12.0, VertAlign::Superscript, false);
        assert_eq!(effective_font_size(&run, &entry), 7.0);
        entry.superscript_ratio = Some(0.65);
        assert_eq!(effective_font_size(&run, &entry), 8.0);
    }

    #[test]
    fn test_effective_font_size_subscript() {
        // 9.5pt at 0.650 is 6.175, Word draws 6.0.
        let mut entry = FontEntry::stub();
        entry.subscript_ratio = Some(0.65);
        let run = make_run(9.5, VertAlign::Subscript, false);
        assert_eq!(effective_font_size(&run, &entry), 6.0);
    }

    #[test]
    fn test_effective_font_size_ignores_small_caps() {
        // smallCaps sizing is per-segment, not per-run — effective_font_size returns base size
        let run = make_run(12.0, VertAlign::Baseline, true);
        assert_eq!(effective_font_size(&run, &FontEntry::stub()), 12.0);
    }

    #[test]
    fn test_smallcaps_segments_mixed() {
        let segs = smallcaps_segments("Hello", 12.0);
        assert_eq!(segs.len(), 2);
        assert_eq!(segs[0], ("H".to_string(), 12.0, "H")); // uppercase stays at 12pt
        assert_eq!(segs[1], ("ELLO".to_string(), 9.5, "ello")); // lowercase → uppercase at 9.5pt
    }

    #[test]
    fn test_smallcaps_segments_all_upper() {
        let segs = smallcaps_segments("ABC", 12.0);
        assert_eq!(segs.len(), 1);
        assert_eq!(segs[0], ("ABC".to_string(), 12.0, "ABC"));
    }

    #[test]
    fn test_smallcaps_segments_all_lower() {
        let segs = smallcaps_segments("abc", 12.0);
        assert_eq!(segs.len(), 1);
        assert_eq!(segs[0], ("ABC".to_string(), 9.5, "abc"));
    }

    #[test]
    fn small_caps_spaces_follow_their_neighbours() {
        // Word probe: a space beside a lowercase letter is small, ". 1" is not.
        let segs = smallcaps_segments("def X. 12", 12.0);
        assert_eq!(segs[0], ("DEF ".to_string(), 9.5, "def "));
        assert_eq!(segs[1], ("X. 12".to_string(), 12.0, "X. 12"));
    }

    #[test]
    fn test_smallcaps_segments_with_nonletters() {
        // Non-letter chars (digits, punctuation) stay at base size, grouped with adjacent same-size
        let segs = smallcaps_segments("A1b", 12.0);
        assert_eq!(segs.len(), 2);
        assert_eq!(segs[0], ("A1".to_string(), 12.0, "A1")); // uppercase + digit both at base size
        assert_eq!(segs[1], ("B".to_string(), 9.5, "b")); // lowercase → uppercase at reduced size
    }

    #[test]
    fn caps_and_small_caps_keep_their_source_letters() {
        let entry = FontEntry::stub();
        let chunks_for = |run: &Run, word: &str, original: Option<&str>| {
            let mut chunks = Vec::new();
            push_word_chunks(
                &mut chunks,
                &entry,
                run,
                word,
                original,
                12.0,
                0.0,
                0.0,
                0.0,
                30.0,
                false,
                false,
            );
            chunks
                .into_iter()
                .map(|c| (c.text, c.actual_text))
                .collect::<Vec<_>>()
        };
        let small_caps = Run {
            small_caps: true,
            ..Run::default()
        };
        assert_eq!(
            chunks_for(&small_caps, "Hello", None),
            [("H".into(), None), ("ELLO".into(), Some("ello".into()))]
        );
        let caps = Run {
            caps: true,
            ..Run::default()
        };
        assert_eq!(caps_word(&caps, "Pirmasis"), "PIRMASIS");
        assert_eq!(
            chunks_for(&caps, "PIRMASIS", Some("Pirmasis")),
            [("PIRMASIS".into(), Some("Pirmasis".into()))]
        );
        // Both on: the word is already all capitals, one segment, the caps original.
        let both = Run {
            caps: true,
            small_caps: true,
            ..Run::default()
        };
        assert_eq!(
            chunks_for(&both, "PIRMASIS", Some("Pirmasis")),
            [("PIRMASIS".into(), Some("Pirmasis".into()))]
        );
        assert_eq!(
            chunks_for(&Run::default(), "plain", None),
            [("plain".into(), None)]
        );
    }

    #[test]
    fn caps_span_groups_chunks_and_returns_to_the_paragraph() {
        use super::super::tagging::{self, ROOT, Tags};
        let mut tags = Tags::new();
        let p = tags.add(ROOT, "P");
        let mut content = tagging::artifact_content();
        tags.begin(&mut content, 0, p);
        content.begin_text();
        let mut lt = LinkTagger::new(&mut tags, 0, p);
        assert!(
            lt.inline(&mut content, None, None, Some("Pirmasis ")),
            "opens a Span"
        );
        assert!(
            !lt.inline(&mut content, None, None, Some("skirsnis")),
            "extends it"
        );
        assert!(
            lt.inline(&mut content, None, None, None),
            "back to the paragraph"
        );
        content.end_text();
        lt.finish(&mut content);
        let stream = String::from_utf8_lossy(&content.finish()).into_owned();
        assert_eq!(stream.matches("/Span").count(), 1);
        assert_eq!(stream.matches("BDC").count(), 3, "P, Span, P again");

        let mut pdf = pdf_writer::Pdf::new();
        let mut next = 1;
        let mut alloc = || {
            next += 1;
            pdf_writer::Ref::new(next)
        };
        tags.write(&mut pdf, &mut alloc, &[pdf_writer::Ref::new(1)], &[]);
        let bytes = pdf.finish();
        assert!(bytes.windows(19).any(|w| w == b"(Pirmasis skirsnis)"));
    }

    #[test]
    fn complex_script_words_take_the_bidi_language() {
        let run = Run {
            text_lang: Some("en-US".into()),
            text_lang_east_asia: Some("ja-JP".into()),
            text_lang_bidi: Some("ar-SA".into()),
            ..Run::default()
        };
        let lang = |text| chunk_lang(&run, text).map(|l| l.to_string());
        assert_eq!(lang("الأرز").as_deref(), Some("ar-SA"));
        assert_eq!(lang("日本").as_deref(), Some("ja-JP"));
        assert_eq!(lang("rice").as_deref(), Some("en-US"));
        let no_bidi = Run {
            text_lang_bidi: None,
            ..run.clone()
        };
        assert_eq!(
            chunk_lang(&no_bidi, "الأرز").as_deref(),
            Some("en-US"),
            "without @bidi the run's language stands"
        );
    }

    #[test]
    fn another_language_gets_a_lang_span() {
        use super::super::tagging::{self, ROOT, Tags};
        let mut tags = Tags::new();
        tags.lang = "en-US".into();
        let p = tags.add(ROOT, "P");
        let mut content = tagging::artifact_content();
        tags.begin(&mut content, 0, p);
        content.begin_text();
        let mut lt = LinkTagger::new(&mut tags, 0, p);
        assert!(
            !lt.inline(&mut content, None, Some("en-GB"), None),
            "same language, no Span"
        );
        assert!(
            lt.inline(&mut content, None, Some("fr-FR"), None),
            "French opens one"
        );
        assert!(
            !lt.inline(&mut content, None, Some("fr-FR"), None),
            "and keeps it"
        );
        assert!(
            lt.inline(&mut content, None, Some("fr-FR"), Some("Bonjour")),
            "caps need their own"
        );
        assert!(
            lt.inline(&mut content, None, None, None),
            "back to the paragraph"
        );
        content.end_text();
        lt.finish(&mut content);

        let mut pdf = pdf_writer::Pdf::new();
        let mut next = 1;
        let mut alloc = || {
            next += 1;
            pdf_writer::Ref::new(next)
        };
        tags.write(&mut pdf, &mut alloc, &[pdf_writer::Ref::new(1)], &[]);
        let bytes = pdf.finish();
        let count = |pat: &[u8]| bytes.windows(pat.len()).filter(|w| *w == pat).count();
        assert_eq!(count(b"/Lang (fr-FR)"), 2);
        assert_eq!(count(b"/ActualText (Bonjour)"), 1);
    }

    #[test]
    fn test_vert_y_offset_baseline() {
        let run = make_run(12.0, VertAlign::Baseline, false);
        assert_eq!(vert_y_offset(&run), 0.0);
    }

    #[test]
    fn test_vert_y_offset_superscript() {
        let run = make_run(12.0, VertAlign::Superscript, false);
        let expected = 12.0 * 0.35; // 4.2
        assert!((vert_y_offset(&run) - expected).abs() < 0.01);
    }

    #[test]
    fn test_vert_y_offset_subscript() {
        let run = make_run(12.0, VertAlign::Subscript, false);
        let expected = -12.0 * 0.14; // -1.68
        assert!((vert_y_offset(&run) - expected).abs() < 0.01);
    }

    #[test]
    fn test_vert_y_offset_adds_position() {
        let mut run = make_run(12.0, VertAlign::Superscript, false);
        run.position = -3.0;
        assert!((vert_y_offset(&run) - (12.0 * 0.35 - 3.0)).abs() < 0.01);
    }
}
