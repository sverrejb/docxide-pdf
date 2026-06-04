use std::borrow::Cow;
use std::collections::HashMap;

use pdf_writer::types::TextRenderingMode;
use pdf_writer::{Content, Name, Rect, Str};

use crate::fonts::{FontEntry, encode_as_gids, font_key, font_key_buf, to_winansi_bytes};
use crate::model::{
    Alignment, ParagraphBorder, Run, TabAlignment, TabStop, TextFill, TextOutline, VertAlign,
};

use super::color::{fill_color_or_black, stroke_color_or_black};

/// Resolve aligned segment start position given a tab stop, segment runs, and minimum x.
fn resolve_tab_aligned_start(
    stop: &TabStop,
    tab_target: f32,
    seg_runs: &[&Run],
    seen_fonts: &HashMap<String, FontEntry>,
    min_x: f32,
) -> f32 {
    match stop.alignment {
        TabAlignment::Left => tab_target.max(min_x),
        TabAlignment::Center => {
            let sw = segment_width(seg_runs, seen_fonts);
            (tab_target - sw / 2.0).max(min_x)
        }
        TabAlignment::Right => {
            let sw = segment_width(seg_runs, seen_fonts);
            (tab_target - sw).max(min_x)
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

/// Insert extra UAX #14 break positions inside URL-like tokens. UAX #14 is
/// conservative around `://` and percent-encoded runs, so a long URL ends up
/// as a single unbreakable word that overflows the right margin. Break
/// opportunities after `/`, `?`, `#`, `&`, `=`, `;` (skipping the `://`) let
/// the line breaker wrap URLs the way Word does.
fn augment_url_breaks(text: &str, breaks: &mut Vec<usize>) {
    let bytes = text.as_bytes();
    let mut search_from = 0;
    while let Some(rel) = text[search_from..].find("://") {
        let scheme_end = search_from + rel + 3;
        // Break after each URL-structure char, but stop at the next whitespace
        // — that terminates the URL token.
        let mut i = scheme_end;
        while i < bytes.len() {
            let b = bytes[i];
            if b.is_ascii_whitespace() {
                break;
            }
            if matches!(b, b'/' | b'?' | b'#' | b'&' | b'=' | b';') {
                breaks.push(i + 1);
            }
            i += 1;
        }
        search_from = i;
    }
    breaks.sort_unstable();
    breaks.dedup();
}

/// Split text into (preceding_space_count, word) segments using UAX #14 line break rules.
/// Handles CJK character boundaries, hyphens, and punctuation break opportunities
/// in addition to whitespace. Non-breaking spaces (U+00A0) and ideographic spaces
/// (U+3000) are kept within words. Trailing spaces after the last word are handled
/// separately by the caller.
fn split_preserving_spaces(text: &str) -> Vec<(usize, &str)> {
    if text.is_empty() {
        return Vec::new();
    }

    let mut result = Vec::new();
    let mut pending_spaces: usize = 0;

    // Collect UAX #14 break positions (byte offsets where breaks are allowed/mandatory)
    let mut breaks: Vec<usize> = unicode_linebreak::linebreaks(text)
        .map(|(pos, _)| pos)
        .collect();
    augment_url_breaks(text, &mut breaks);

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

pub(super) struct WordChunk {
    pub(super) pdf_font: String,
    pub(super) text: String,
    pub(super) font_size: f32,
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
    pub(super) y_offset: f32,   // vertical offset for superscript/subscript
    pub(super) hyperlink_url: Option<String>,
    pub(super) inline_image_name: Option<String>,
    pub(super) inline_image_height: f32,
    pub(super) inline_image_stroke_color: Option<[u8; 3]>,
    pub(super) inline_image_stroke_width: f32,
    pub(super) inline_image_shadow: Option<crate::model::ImageShadow>,
    pub(super) inline_image_glow: Option<crate::model::ImageGlow>,
    pub(super) inline_image_effect_xobjs: Option<super::images::EffectXObjs>,
    pub(super) inline_image_clip: Option<crate::model::ShapeGeometry>,
    pub(super) synthetic_bold: bool,
    pub(super) text_outline: Option<TextOutline>,
    pub(super) text_fill: Option<TextFill>,
    pub(super) comment_ids: Vec<u32>,
}

/// Pale-pink highlight color Word uses for comment-anchored text spans.
pub(super) const COMMENT_HIGHLIGHT_RGB: [u8; 3] = [251, 220, 217];

fn push_decoration(
    decorations: &mut Vec<(f32, f32, f32, f32, Option<[u8; 3]>)>,
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    color: Option<[u8; 3]>,
) {
    let merged = decorations.last_mut().filter(|(_, dy, _, dh, dc)| {
        (*dy - y).abs() < 0.01 && (*dh - height).abs() < 0.01 && *dc == color
    });
    if let Some(prev) = merged {
        prev.2 = (x + width) - prev.0;
    } else {
        decorations.push((x, y, width, height, color));
    }
}

impl WordChunk {
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
            hyperlink_url: run.hyperlink_url.clone(),
            inline_image_name: None,
            inline_image_height: 0.0,
            inline_image_stroke_color: None,
            inline_image_stroke_width: 0.0,
            inline_image_shadow: None,
            inline_image_glow: None,
            inline_image_effect_xobjs: None,
            inline_image_clip: None,
            synthetic_bold: entry.synthetic_bold,
            text_outline: run.text_outline.clone(),
            text_fill: run.text_fill.clone(),
            comment_ids: run.comment_ids.clone(),
        }
    }

    fn image(
        pdf_name: &str,
        font_size: f32,
        x_offset: f32,
        display_width: f32,
        display_height: f32,
        stroke_color: Option<[u8; 3]>,
        stroke_width: f32,
        shadow: Option<crate::model::ImageShadow>,
        glow: Option<crate::model::ImageGlow>,
        effect_xobjs: Option<super::images::EffectXObjs>,
        clip: Option<crate::model::ShapeGeometry>,
    ) -> Self {
        Self {
            pdf_font: String::new(),
            text: String::new(),
            font_size,
            color: None,
            highlight: None,
            shading: None,
            border: None,
            x_offset,
            width: display_width,
            underline: false,
            double_underline: false,
            strikethrough: false,
            dstrike: false,
            char_spacing: 0.0,
            text_scale: 100.0,
            y_offset: 0.0,
            hyperlink_url: None,
            inline_image_name: Some(pdf_name.to_string()),
            inline_image_height: display_height,
            inline_image_stroke_color: stroke_color,
            inline_image_stroke_width: stroke_width,
            inline_image_shadow: shadow,
            inline_image_glow: glow,
            inline_image_effect_xobjs: effect_xobjs,
            inline_image_clip: clip,
            synthetic_bold: false,
            text_outline: None,
            text_fill: None,
            comment_ids: Vec::new(),
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
            hyperlink_url: None,
            inline_image_name: None,
            inline_image_height: 0.0,
            inline_image_stroke_color: None,
            inline_image_stroke_width: 0.0,
            inline_image_shadow: None,
            inline_image_glow: None,
            inline_image_effect_xobjs: None,
            inline_image_clip: None,
            synthetic_bold: false,
            text_outline: None,
            text_fill: None,
            comment_ids: Vec::new(),
        }
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
            hyperlink_url: None,
            inline_image_name: None,
            inline_image_height: 0.0,
            inline_image_stroke_color: None,
            inline_image_stroke_width: 0.0,
            inline_image_shadow: None,
            inline_image_glow: None,
            inline_image_effect_xobjs: None,
            inline_image_clip: None,
            synthetic_bold: false,
            text_outline: None,
            text_fill: None,
            comment_ids: Vec::new(),
        }
    }
}

pub(crate) struct LinkAnnotation {
    pub(super) rect: Rect,
    pub(super) url: String,
}

/// When a line spans two text regions (bothSides wrapping around a float),
/// this stores where the right region begins and its geometry.
pub(super) struct RightRegion {
    pub(super) first_chunk_idx: usize,
    pub(super) region_x: f32,
    pub(super) region_width: f32,
    pub(super) content_width: f32,
}

pub(super) struct TextLine {
    pub(super) chunks: Vec<WordChunk>,
    pub(super) total_width: f32,
    pub(super) ends_with_break: bool,
    pub(super) right_region: Option<RightRegion>,
    /// Font size of the line break run that created this empty line.
    /// Used to compute the correct line height for break-created lines
    /// (Word uses the break run's font metrics, not the paragraph's).
    pub(super) break_font_size: Option<f32>,
}

/// True when a paragraph has no visible text (may still have phantom font-info runs).
pub(super) fn is_text_empty(runs: &[Run]) -> bool {
    runs.iter()
        .all(|r| r.vanish || (r.text.is_empty() && !r.is_tab && !r.is_line_break && r.inline_image.is_none()))
}

fn effective_font_size(run: &Run) -> f32 {
    match run.vertical_align {
        VertAlign::Superscript | VertAlign::Subscript => run.font_size * 0.58,
        VertAlign::Baseline => run.font_size,
    }
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

/// Split a word into (text, font_size) segments for smallCaps rendering.
/// Lowercase chars are uppercased and rendered at base_fs - 2pt;
/// uppercase chars and non-letters stay at base_fs.
pub(super) fn smallcaps_segments(word: &str, base_fs: f32) -> Vec<(String, f32)> {
    let reduced = (base_fs - 2.0).max(1.0);
    let mut segments: Vec<(String, f32)> = Vec::new();
    for ch in word.chars() {
        let is_lower = ch.is_lowercase();
        let fs = if is_lower { reduced } else { base_fs };
        let display: String = if is_lower {
            ch.to_uppercase().collect()
        } else {
            ch.to_string()
        };
        if let Some(last) = segments.last_mut() {
            if (last.1 - fs).abs() < 0.001 {
                last.0.push_str(&display);
                continue;
            }
        }
        segments.push((display, fs));
    }
    segments
}

/// Compute the width of a word, handling smallCaps per-segment sizing.
fn word_width_for_run(
    entry: &FontEntry, run: &Run, word: &str,
    eff_fs: f32, kern: bool, cs: f32, ts: f32,
) -> f32 {
    if run.small_caps {
        smallcaps_segments(word, eff_fs).iter().map(|(seg, fs)| {
            let seg_kern = run.kern_threshold.is_some_and(|t| *fs >= t);
            entry.word_width(seg, *fs, seg_kern) * ts + cs * seg.chars().count() as f32
        }).sum()
    } else {
        let char_count = word.chars().count();
        entry.word_width(word, eff_fs, kern) * ts + cs * char_count as f32
    }
}

/// Push WordChunks for a word, splitting into per-segment chunks for smallCaps.
fn push_word_chunks(
    chunks: &mut Vec<WordChunk>,
    entry: &FontEntry,
    run: &Run,
    word: &str,
    eff_fs: f32,
    cs: f32,
    y_off: f32,
    x_start: f32,
    total_ww: f32,
) {
    if run.small_caps {
        let segs = smallcaps_segments(word, eff_fs);
        let mut seg_x = x_start;
        for (seg_text, seg_fs) in &segs {
            let seg_kern = run.kern_threshold.is_some_and(|t| *seg_fs >= t);
            let ts = run.text_scale / 100.0;
            let seg_w = entry.word_width(seg_text, *seg_fs, seg_kern) * ts
                + cs * seg_text.chars().count() as f32;
            chunks.push(WordChunk::text(entry, run, seg_text, *seg_fs, cs, y_off, seg_x, seg_w));
            seg_x += seg_w;
        }
    } else {
        chunks.push(WordChunk::text(entry, run, word, eff_fs, cs, y_off, x_start, total_ww));
    }
}

fn vert_y_offset(run: &Run) -> f32 {
    match run.vertical_align {
        VertAlign::Superscript => run.font_size * 0.35,
        VertAlign::Subscript => -run.font_size * 0.14,
        VertAlign::Baseline => 0.0,
    }
}

const DEFAULT_TAB_INTERVAL: f32 = 36.0; // 0.5 inches

/// Count CJK↔Latin/Digit boundaries in text for autoSpaceDE/DN.
/// Word adds ~1/4 em spacing at each boundary by default.
#[allow(dead_code)]
fn count_script_boundaries(text: &str) -> usize {
    let mut count = 0;
    let mut prev_cjk: Option<bool> = None;
    for ch in text.chars() {
        if ch.is_whitespace() {
            prev_cjk = None;
            continue;
        }
        let is_cjk = crate::docx::is_east_asian_char(ch)
            || is_cjk_punctuation(ch);
        if let Some(was_cjk) = prev_cjk {
            if was_cjk != is_cjk {
                count += 1;
            }
        }
        prev_cjk = Some(is_cjk);
    }
    count
}

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
        ends_with_break: false,
        right_region: None,
        break_font_size: None,
    }
}

fn finish_line_with_break(chunks: &mut Vec<WordChunk>) -> TextLine {
    let mut line = finish_line(chunks);
    line.ends_with_break = true;
    line
}

/// Layout runs into wrapped lines.
/// Handles cross-run contiguous text correctly: no space is inserted between
/// runs unless the preceding text ended with whitespace or the new run starts
/// with whitespace (e.g., "bold" + ", " → "bold," not "bold ,").
/// `width_after_line`: after building this many lines, switch to a different
/// max_width (used for text wrapping around floating tables where lines beside
/// the table are narrow, then lines below it expand to full column width).
/// Dual-region geometry for bothSides wrapping: (left_x, left_w, right_x, right_w).
/// For lines outside the float zone, right_w is 0.0 (single region).
pub(super) type DualRegion = (f32, f32, f32, f32);

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
    auto_space: bool,
) -> Vec<TextLine> {
    let mut lines: Vec<TextLine> = Vec::new();
    let mut current_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;
    let mut pending_space_w: f32 = 0.0;
    // Underline state of the run that emitted the pending whitespace. A space
    // is underlined when its *own* run is underlined (Word draws underline
    // continuously across spaces inside an underlined run, but not across a
    // space that belongs to a neighbouring non-underlined run).
    let mut pending_space_underline = false;
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

    let left_max = |line_count: usize| -> f32 {
        if let Some(dual) = per_line_dual {
            if let Some(&(_, lw, _, _)) = dual.get(line_count) {
                return lw;
            }
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
            dual.get(line_count).and_then(|&(_, _, rx, rw)| {
                if rw > 0.0 { Some((rx, rw, rw)) } else { None }
            })
        })
    };

    // Finish current line, recording right-region info if applicable
    let finish_dual_line = |chunks: &mut Vec<WordChunk>,
                            in_right: &mut bool,
                            right_info: &mut Option<(usize, f32, f32)>| -> TextLine {
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
        if run.vanish || run.is_tab {
            continue;
        }

        if run.is_line_break {
            let line = finish_dual_line(&mut current_chunks, &mut in_right_region, &mut cur_right_info);
            let line = TextLine { ends_with_break: true, ..line };
            lines.push(line);
            current_x = 0.0;
            pending_space_w = 0.0;
            prev_last_char = None;
            continue;
        }

        // Handle inline images as single block elements in the line
        if let Some(img) = &run.inline_image {
            if let Some(pdf_name) = inline_image_names.get(&run_idx) {
                let img_w = img.display_width;
                let need_space = !current_chunks.is_empty() && pending_space_w > 0.0;
                let proposed_x = if need_space {
                    current_x + pending_space_w
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
                if !current_chunks.is_empty() && proposed_x + img_w > cur_max {
                    if !in_right_region {
                        if let Some((rx, rw, _)) = right_region_for(lines.len()) {
                            cur_right_info = Some((current_chunks.len(), rx, rw));
                            in_right_region = true;
                            pending_space_w = 0.0;
                            // Retry placement in right region
                            let proposed_x2 = 0.0;
                            if proposed_x2 + img_w <= rw {
                                current_chunks.push(WordChunk::image(
                                    pdf_name, run.font_size, proposed_x2, img_w, img.display_height,
                                    img.stroke_color, img.stroke_width, img.shadow.clone(),
                                    img.glow.clone(), effect_inline_names.get(&run_idx).cloned(),
                                    img.clip_geometry.clone(),
                                ));
                                current_x = img_w;
                                continue;
                            }
                        }
                    }
                    lines.push(finish_dual_line(&mut current_chunks, &mut in_right_region, &mut cur_right_info));
                    current_x = 0.0;
                } else {
                    current_x = proposed_x;
                }
                pending_space_w = 0.0;

                current_chunks.push(WordChunk::image(
                    pdf_name, run.font_size, current_x, img_w, img.display_height,
                    img.stroke_color, img.stroke_width, img.shadow.clone(),
                    img.glow.clone(), effect_inline_names.get(&run_idx).cloned(),
                    img.clip_geometry.clone(),
                ));
                current_x += img_w;
            }
            continue;
        }

        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key).expect("font registered");
        let eff_fs = effective_font_size(run);
        let space_w = entry.space_width(eff_fs);
        let text = effective_text(run);
        let y_off = vert_y_offset(run);

        let cs = run.char_spacing;
        let ts = run.text_scale / 100.0;
        let space_w_cs = space_w * ts + cs;

        let mut is_first_word_in_run = true;
        for (space_count, word) in split_preserving_spaces(&text) {
            pending_space_w += space_count as f32 * space_w_cs;
            if space_count > 0 {
                pending_space_underline = run.underline;
                pending_space_double = run.double_underline;
                pending_space_color = run.color;
                pending_space_border = run.border.clone();
            }

            // CJK auto-spacing (autoSpaceDE/DN): add ~0.25em gap at
            // script boundaries between East Asian and Latin/digit text
            // when there is no explicit whitespace.
            if auto_space {
                if let Some(prev_ch) = prev_last_char {
                    if let Some(first_ch) = word.chars().next() {
                        if pending_space_w == 0.0 && space_count == 0 {
                            let prev_ea = crate::docx::is_east_asian_char(prev_ch)
                                || is_cjk_punctuation(prev_ch);
                            let cur_ea = crate::docx::is_east_asian_char(first_ch)
                                || is_cjk_punctuation(first_ch);
                            if prev_ea != cur_ea {
                                pending_space_w += eff_fs * 0.25;
                            }
                        }
                    }
                }
            }

            // A word continues the previous word across a run boundary when
            // there is no whitespace between runs and no leading spaces.
            let is_continuation = is_first_word_in_run
                && space_count == 0
                && pending_space_w == 0.0
                && !current_chunks.is_empty();
            is_first_word_in_run = false;

            let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
            let ww = word_width_for_run(entry, run, word, eff_fs, kern, cs, ts);
            prev_last_char = word.chars().last();

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
            let overflows = proposed_x + ww > cur_max + 0.05;

            // When leading spaces push the first word past the line width,
            // emit a blank line for the spaces and start the word at x=0.
            // Word absorbs the spaces onto the current line and wraps the
            // word to the next line.
            if current_chunks.is_empty() && overflows && pending_space_w > 0.0 && !in_right_region {
                lines.push(finish_dual_line(&mut current_chunks, &mut in_right_region, &mut cur_right_info));
                pending_space_w = 0.0;
                push_word_chunks(&mut current_chunks, entry, run, word, eff_fs, cs, y_off, 0.0, ww);
                current_x = ww;
                continue;
            }

            // For the first word on a line, also overflow if the
            // left region is zero-width (so words go straight to
            // the right region for e.g. left-aligned images).
            let first_word_overflow = current_chunks.is_empty()
                && !in_right_region
                && cur_max <= 0.0
                && right_region_for(lines.len()).is_some();

            if (!current_chunks.is_empty() && overflows && !is_continuation) || first_word_overflow {
                // Word doesn't fit in current region
                if !in_right_region {
                    // Try spilling to the right region on the same line
                    if let Some((rx, rw, _)) = right_region_for(lines.len()) {
                        cur_right_info = Some((current_chunks.len(), rx, rw));
                        in_right_region = true;
                        pending_space_w = 0.0;
                        // Place this word at the start of the right region
                        push_word_chunks(&mut current_chunks, entry, run, word, eff_fs, cs, y_off, 0.0, ww);
                        current_x = ww;
                        continue;
                    }
                }
                // Try hyphenation before wrapping whole word
                // No right region or right region also full — wrap to next line
                lines.push(finish_dual_line(&mut current_chunks, &mut in_right_region, &mut cur_right_info));
                current_x = 0.0;
                // If the new line's left region is zero-width, go
                // straight to the right region for this word.
                if let Some((rx, rw, _)) = right_region_for(lines.len()) {
                    let new_left_max = left_max(lines.len());
                    if new_left_max <= 0.0 {
                        cur_right_info = Some((0, rx, rw));
                        in_right_region = true;
                        pending_space_w = 0.0;
                        push_word_chunks(&mut current_chunks, entry, run, word, eff_fs, cs, y_off, 0.0, ww);
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
                current_chunks.push(WordChunk::tab_underline(
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

            push_word_chunks(&mut current_chunks, entry, run, word, eff_fs, cs, y_off, current_x, ww);
            current_x += ww;
        }

        // Accumulate trailing whitespace for the next run
        let trailing_spaces = text.chars().rev().take_while(|c| is_break_space(*c)).count();
        if trailing_spaces > 0 {
            pending_space_w += trailing_spaces as f32 * space_w_cs;
            pending_space_underline = run.underline;
            pending_space_double = run.double_underline;
            pending_space_color = run.color;
            pending_space_border = run.border.clone();
        }
    }

    if !current_chunks.is_empty() {
        lines.push(finish_dual_line(&mut current_chunks, &mut in_right_region, &mut cur_right_info));
    }

    // A trailing line break creates an empty line after it (Word adds a blank
    // line for each w:br at the end of a paragraph).  Store the break run's
    // font size so the caller can compute the correct height for this line.
    if lines.last().is_some_and(|l| l.ends_with_break) {
        let break_fs = runs.iter().rev()
            .find(|r| r.is_line_break)
            .map(|r| r.font_size);
        lines.push(TextLine {
            chunks: vec![],
            total_width: 0.0,
            ends_with_break: false,
            right_region: None,
            break_font_size: break_fs,
        });
    }

    if lines.is_empty() {
        lines.push(TextLine {
            chunks: vec![],
            total_width: 0.0,
            ends_with_break: false,
            right_region: None,
            break_font_size: None,
        });
    }
    lines
}

fn find_next_tab_stop(
    current_x: f32,
    tab_stops: &[TabStop],
    indent_left: f32,
    default_tab_interval: f32,
) -> TabStop {
    let abs_x = current_x + indent_left;
    tab_stops
        .iter()
        .find(|s| s.position > abs_x + 0.5)
        .cloned()
        .unwrap_or_else(|| {
            let interval = if default_tab_interval > 0.0 {
                default_tab_interval
            } else {
                DEFAULT_TAB_INTERVAL
            };
            let next = ((abs_x / interval).floor() + 1.0) * interval;
            TabStop {
                position: next,
                alignment: TabAlignment::Left,
                leader: None,
            }
        })
}

fn segment_width(runs: &[&Run], seen_fonts: &HashMap<String, FontEntry>) -> f32 {
    let mut w: f32 = 0.0;
    let mut first = true;
    let mut key_buf = String::new();
    for run in runs {
        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key).expect("font registered");
        let eff_fs = effective_font_size(run);
        let ts = run.text_scale / 100.0;
        let cs = run.char_spacing;
        let space_w = entry.space_width(eff_fs) * ts + cs;
        let text = effective_text(run);
        for (i, word) in text.split_whitespace().enumerate() {
            if !first || i > 0 {
                w += space_w;
            }
            let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
            w += entry.word_width(word, eff_fs, kern) * ts + cs * word.chars().count() as f32;
            first = false;
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
        let eff_fs = effective_font_size(run);
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
        let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
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

    for (seg_idx, (seg_runs, seg_indices, tab_before, tab_run_before)) in segments.iter().enumerate() {
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
            // Trailing-space handling before a tab depends on whether an
            // explicit tab stop applies:
            // - An explicit stop AFTER current_x consumes the trailing spaces
            //   (header/footer center+right pattern — the tab must snap to that
            //   stop regardless of how wide the preceding whitespace is).
            // - Otherwise we fall through to default tab stops; trailing spaces
            //   advance the cursor so the tab can snap past them. This matters
            //   for list paragraphs with long dot-leader trailing-space runs.
            let abs_x_no_spaces = current_x + line_indent;
            let has_explicit_after =
                tab_stops.iter().any(|s| s.position > abs_x_no_spaces + 0.5);
            if !has_explicit_after {
                current_x += pending_space_w;
            }
            pending_space_w = 0.0;
            let stop = find_next_tab_stop(current_x, tab_stops, line_indent, default_tab_stop);
            let mut effective_tab_target = stop.position - line_indent;
            let mut seg_start =
                resolve_tab_aligned_start(&stop, effective_tab_target, seg_runs, seen_fonts, current_x);
            let mut resolved_leader = stop.leader;

            // Explicit tab stops may legitimately target positions beyond the
            // paragraph's right indent — TOC entries are a common case where a
            // right-aligned page-number tab sits near the content edge. Allow
            // segments to extend up to the physical content edge (indent_right
            // beyond max_width) before forcing a wrap.
            let wrap_limit = line_max + indent_right;
            if seg_start > wrap_limit && !all_chunks.is_empty() {
                result_lines.push(finish_line(&mut all_chunks));
                current_x = 0.0;
                is_first_line = false;
                let new_stop = find_next_tab_stop(0.0, tab_stops, indent_left, default_tab_stop);
                let new_target = new_stop.position - indent_left;
                seg_start =
                    resolve_tab_aligned_start(&new_stop, new_target, seg_runs, seen_fonts, 0.0);
                resolved_leader = new_stop.leader;
                effective_tab_target = new_target;
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
                    let eff_fs = effective_font_size(tab_run).max(font_run.font_size);
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
                    let font_run: Option<&Run> = seg_runs.first().copied().or_else(|| {
                        segments[..seg_idx]
                            .iter()
                            .rev()
                            .flat_map(|(r, _, _, _)| r.last().copied())
                            .next()
                    }).or_else(|| {
                        // Tab-only paragraphs: fall back to any run (including tab runs)
                        runs.iter().find(|r| !r.font_name.is_empty())
                    });
                    if let Some(run) = font_run {
                        let key = font_key_buf(run, &mut key_buf);
                        let entry = seen_fonts.get(key).expect("font registered");
                        let eff_fs = effective_font_size(run);
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
        }

        // Layout text in this segment from current_x
        for (local_idx, run) in seg_runs.iter().enumerate() {
            if run.is_line_break {
                result_lines.push(finish_line_with_break(&mut all_chunks));
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
                        img.display_width,
                        img.display_height,
                        img.stroke_color,
                        img.stroke_width,
                        img.shadow.clone(),
                        img.glow.clone(),
                        effect_inline_names.get(&seg_indices[local_idx]).cloned(),
                        img.clip_geometry.clone(),
                    ));
                    current_x += img.display_width;
                }
                continue;
            }

            let key = font_key_buf(run, &mut key_buf);
            let entry = seen_fonts.get(key).expect("font registered");
            let eff_fs = effective_font_size(run);
            let space_w = entry.space_width(eff_fs);
            let y_off = vert_y_offset(run);
            let text = effective_text(run);

            let cs = run.char_spacing;
            let ts = run.text_scale / 100.0;
            let space_w_cs = space_w * ts + cs;
            let segments = split_preserving_spaces(&text);
            for (seg_idx, &(space_count, word)) in segments.iter().enumerate() {
                let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
                let ww = word_width_for_run(entry, run, word, eff_fs, kern, cs, ts);
                pending_space_w += space_count as f32 * space_w_cs;
                if space_count > 0 {
                    pending_space_underline = run.underline;
                    pending_space_double = run.double_underline;
                    pending_space_color = run.color;
                    pending_space_border = run.border.clone();
                }
                let applied_space = pending_space_w > 0.0
                    && (!all_chunks.is_empty() || space_count > 0);
                if applied_space {
                    if pending_space_underline && !all_chunks.is_empty() {
                        all_chunks.push(WordChunk::tab_underline(
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
                // Wrap word to new line if it exceeds max_width
                if current_x + ww > cur_line_max && !all_chunks.is_empty() && !is_continuation {
                    result_lines.push(finish_line(&mut all_chunks));
                    current_x = 0.0;
                    is_first_line = false;
                }
                push_word_chunks(&mut all_chunks, entry, run, word, eff_fs, cs, y_off, current_x, ww);
                current_x += ww;
            }
            // Accumulate trailing whitespace for the next run (or the next tab stop)
            let trailing_spaces = text.chars().rev().take_while(|c| is_break_space(*c)).count();
            if trailing_spaces > 0 {
                pending_space_w += trailing_spaces as f32 * space_w_cs;
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

    // Finalize remaining chunks into the last line
    if !all_chunks.is_empty() {
        result_lines.push(finish_line(&mut all_chunks));
    } else if result_lines.is_empty() {
        result_lines.push(TextLine {
            chunks: vec![],
            total_width: 0.0,
            ends_with_break: false,
            right_region: None,
            break_font_size: None,
        });
    }

    // Trailing break creates an empty line (same as build_paragraph_lines)
    if result_lines.last().is_some_and(|l| l.ends_with_break) {
        let break_fs = runs.iter().rev()
            .find(|r| r.is_line_break)
            .map(|r| r.font_size);
        result_lines.push(TextLine {
            chunks: vec![],
            total_width: 0.0,
            ends_with_break: false,
            right_region: None,
            break_font_size: break_fs,
        });
    }

    result_lines
}

pub(super) fn encode_text_for_pdf(
    text: &str,
    pdf_font: &str,
    pdf_name_to_entry: &HashMap<&str, &FontEntry>,
) -> Vec<u8> {
    match pdf_name_to_entry
        .get(pdf_font)
        .and_then(|e| e.char_to_gid.as_ref())
    {
        Some(map) => encode_as_gids(text, map),
        None => to_winansi_bytes(text),
    }
}

/// Render pre-built lines applying the paragraph alignment.
/// `total_line_count` is the full paragraph line count (for justify: last line stays left-aligned).
pub(super) fn render_paragraph_lines(
    content: &mut Content,
    lines: &[TextLine],
    alignment: &Alignment,
    margin_left: f32,
    text_width: f32,
    first_baseline_y: f32,
    line_pitch: f32,
    total_line_count: usize,
    first_line_index: usize,
    links: &mut Vec<LinkAnnotation>,
    first_line_hanging: f32,
    seen_fonts: &HashMap<String, FontEntry>,
    line_geometry: Option<&[(f32, f32)]>,
    gradient_specs: &mut Vec<super::GradientSpec>,
    mut comment_anchors: Option<&mut Vec<(u32, f32, f32, f32)>>,
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

    // Pre-compute per-line y offsets accounting for inline images making lines taller
    let mut line_y_offsets: Vec<f32> = Vec::with_capacity(lines.len());
    let mut cumulative_y = 0.0f32;
    for (i, line) in lines.iter().enumerate() {
        line_y_offsets.push(cumulative_y);
        let img_h = line
            .chunks
            .iter()
            .map(|c| c.inline_image_height)
            .fold(0.0f32, f32::max);
        cumulative_y += if img_h > line_pitch {
            img_h
        } else {
            line_pitch
        };
        // First line offset is always 0
        if i == 0 {
            cumulative_y = line_pitch.max(img_h);
            line_y_offsets[0] = 0.0;
        }
    }

    let last_line_idx = total_line_count.saturating_sub(1);
    for (line_num, line) in lines.iter().enumerate() {
        let y = first_baseline_y - line_y_offsets[line_num];
        let global_line_idx = first_line_index + line_num;

        let (base_margin, base_width) = line_geometry
            .and_then(|g| g.get(global_line_idx))
            .copied()
            .unwrap_or((margin_left, text_width));
        let (eff_margin, eff_width) = if global_line_idx == 0 && first_line_hanging.abs() > 0.001 {
            (
                base_margin - first_line_hanging,
                base_width + first_line_hanging,
            )
        } else {
            (base_margin, base_width)
        };

        // Compute left-region content width (for alignment/justify)
        let left_content_width = if let Some(ref rr) = line.right_region {
            line.chunks[..rr.first_chunk_idx]
                .last()
                .map(|c| c.x_offset + c.width)
                .unwrap_or(0.0)
        } else {
            line.total_width
        };

        let left_chunk_count = line.right_region.as_ref()
            .map(|rr| rr.first_chunk_idx)
            .unwrap_or(line.chunks.len());

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
        // paragraph's true last line suppresses justification.
        let can_justify = *alignment == Alignment::Justify
            && global_line_idx != last_line_idx;

        let is_cjk_justified = can_justify && has_cjk_content && left_char_count > 1;
        let is_justified = is_cjk_justified
            || (can_justify && left_chunk_count > 1);

        let line_start_x = match alignment {
            Alignment::Center => eff_margin + (eff_width - left_content_width) / 2.0,
            Alignment::Right => eff_margin + eff_width - left_content_width,
            Alignment::Left | Alignment::Justify => eff_margin,
        };

        // For CJK justify, distribute via Tc (char spacing) across all characters.
        // Tc adds space after each char including the last, so:
        //   content_width + total_chars * Tc = eff_width
        //   Tc = (eff_width - content_width) / total_chars
        let justify_tc = if is_cjk_justified {
            (eff_width - left_content_width) / left_char_count as f32
        } else {
            0.0
        };

        let extra_per_gap = if is_justified && !is_cjk_justified {
            ((eff_width - left_content_width) / (left_chunk_count - 1).max(1) as f32).max(0.0)
        } else {
            0.0
        };

        // Right-region start x and justify gap (if dual-region line)
        let (right_start_x, right_extra_per_gap) = if let Some(ref rr) = line.right_region {
            let right_chunks = line.chunks.len() - rr.first_chunk_idx;
            let rx = match alignment {
                Alignment::Center => rr.region_x + (rr.region_width - rr.content_width) / 2.0,
                Alignment::Right => rr.region_x + rr.region_width - rr.content_width,
                Alignment::Left | Alignment::Justify => rr.region_x,
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
            if let Some(ref rr) = line.right_region {
                if chunk_idx >= rr.first_chunk_idx {
                    let local_idx = chunk_idx - rr.first_chunk_idx;
                    return right_start_x + chunk.x_offset + local_idx as f32 * right_extra_per_gap;
                }
            }
            if is_cjk_justified {
                // Tc adds extra space after each character; shift chunk start
                // by the cumulative Tc of all chars in preceding chunks.
                let chars_before: usize = line.chunks[..chunk_idx]
                    .iter()
                    .map(|c| c.text.chars().count())
                    .sum();
                line_start_x + chunk.x_offset + chars_before as f32 * justify_tc
            } else {
                line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap
            }
        };

        let mut decorations: Vec<(f32, f32, f32, f32, Option<[u8; 3]>)> = Vec::new();

        // Draw run shading first (bottom layer), then run highlights on top.
        // Both use the same rectangle geometry; when a run sets both, highlight
        // visually occludes shading — matching Word.
        let draw_run_backgrounds =
            |content: &mut Content, accessor: fn(&WordChunk) -> Option<[u8; 3]>| {
                let mut bg_start_x = 0.0f32;
                let mut bg_color: Option<[u8; 3]> = None;
                let mut bg_end_x = 0.0f32;
                let mut bg_fs = 0.0f32;

                let flush = |content: &mut Content,
                             color: [u8; 3],
                             sx: f32,
                             ex: f32,
                             fs: f32,
                             y: f32| {
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
        let flush_border = |content: &mut Content,
                            border: &ParagraphBorder,
                            sx: f32,
                            ex: f32,
                            fs: f32,
                            y: f32| {
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
        let has_text_chunks = line.chunks.iter().any(|c| {
            c.inline_image_name.is_none() && (!c.text.is_empty() || c.underline)
        });

        if has_text_chunks {
            content.begin_text();
            let mut td_x = 0.0_f32;
            let mut td_y = 0.0_f32;

            for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
                if chunk.inline_image_name.is_some() {
                    continue;
                }

                let x = chunk_abs_x(chunk_idx, chunk);
                let cy = y + chunk.y_offset;

                // w14:textFill gradient → register an axial-shading pattern spanning
                // the glyph extent and use it as the fill color for this chunk.
                // Solid fill overrides w:color; NoFill falls through to w:color (the
                // outline path zeroes the fill via text rendering mode = Stroke).
                let mut chunk_uses_gradient = false;
                if let Some(TextFill::Gradient { ref stops, angle_deg }) = chunk.text_fill {
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
                    content.set_fill_color_space(
                        pdf_writer::types::ColorSpaceOperand::Pattern,
                    );
                    content.set_fill_pattern([], Name(pat_name.as_bytes()));
                    pattern_fill_active = true;
                    current_color = None;
                    chunk_uses_gradient = true;
                }

                let effective_color = match chunk.text_fill {
                    Some(TextFill::Solid(c)) => Some(c),
                    _ => chunk.color,
                };
                if !chunk_uses_gradient
                    && (pattern_fill_active || effective_color != current_color)
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
                        content.set_line_width(chunk.font_size * 0.02);
                        stroke_color_or_black(content, chunk.color);
                        content.set_text_rendering_mode(TextRenderingMode::FillStroke);
                    }
                }

                if !has_text_outline && chunk.synthetic_bold != cur_synthetic_bold {
                    if chunk.synthetic_bold {
                        content.set_line_width(chunk.font_size * 0.02);
                        stroke_color_or_black(content, chunk.color);
                        content.set_text_rendering_mode(TextRenderingMode::FillStroke);
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

                content.next_line(x - td_x, cy - td_y);
                td_x = x;
                td_y = cy;

                // Per-character font fallback: if some chars are missing
                // from the primary font, split into segments and render
                // missing chars with the CJK fallback font.
                let primary_entry = pdf_name_to_entry.get(chunk.pdf_font.as_str());
                let has_missing = primary_entry
                    .is_some_and(|e| !e.missing_cjk_chars.is_empty());
                let fallback_entry = has_missing
                    .then(|| seen_fonts.get("__cjk_fallback"))
                    .flatten();

                if let (Some(primary), Some(fallback)) = (primary_entry, fallback_entry) {
                    let _primary_gids = primary.char_to_gid.as_ref();
                    let fallback_gids = fallback.char_to_gid.as_ref();
                    // Split text into runs of primary vs fallback chars
                    let mut seg_start = 0;
                    let mut in_fallback = false;
                    let chars: Vec<char> = chunk.text.chars().collect();
                    for (i, &ch) in chars.iter().enumerate() {
                        let needs_fb = primary.missing_cjk_chars.contains(&ch);
                        if i == 0 {
                            in_fallback = needs_fb;
                        } else if needs_fb != in_fallback {
                            let seg: String = chars[seg_start..i].iter().collect();
                            if in_fallback {
                                if let Some(map) = fallback_gids {
                                    let fb_name = &fallback.pdf_name;
                                    content.set_font(
                                        Name(fb_name.as_bytes()),
                                        chunk.font_size,
                                    );
                                    content.show(Str(&encode_as_gids(&seg, map)));
                                    content.set_font(
                                        Name(chunk.pdf_font.as_bytes()),
                                        chunk.font_size,
                                    );
                                }
                            } else {
                                let bytes = encode_text_for_pdf(
                                    &seg,
                                    &chunk.pdf_font,
                                    &pdf_name_to_entry,
                                );
                                content.show(Str(&bytes));
                            }
                            seg_start = i;
                            in_fallback = needs_fb;
                        }
                    }
                    // Flush last segment
                    let seg: String = chars[seg_start..].iter().collect();
                    if in_fallback {
                        if let Some(map) = fallback_gids {
                            let fb_name = &fallback.pdf_name;
                            content
                                .set_font(Name(fb_name.as_bytes()), chunk.font_size);
                            content.show(Str(&encode_as_gids(&seg, map)));
                            content.set_font(
                                Name(chunk.pdf_font.as_bytes()),
                                chunk.font_size,
                            );
                        }
                    } else {
                        let bytes = encode_text_for_pdf(
                            &seg,
                            &chunk.pdf_font,
                            &pdf_name_to_entry,
                        );
                        content.show(Str(&bytes));
                    }
                } else {
                    let text_bytes = encode_text_for_pdf(
                        &chunk.text,
                        &chunk.pdf_font,
                        &pdf_name_to_entry,
                    );
                    content.show(Str(&text_bytes));
                };

                if chunk.underline {
                    let thick = (chunk.font_size * 0.05).max(0.5);
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
                    let thick = (chunk.font_size * 0.05).max(0.5);
                    let st_y = y + chunk.font_size * 0.3;
                    decorations.push((x, st_y, chunk.width, thick, chunk.color));
                }
                if chunk.dstrike {
                    let thick = (chunk.font_size * 0.05).max(0.5);
                    let gap = thick * 1.5;
                    let mid_y = y + chunk.font_size * 0.3;
                    decorations.push((x, mid_y - gap / 2.0, chunk.width, thick, chunk.color));
                    decorations.push((x, mid_y + gap / 2.0, chunk.width, thick, chunk.color));
                }

                if let Some(ref url) = chunk.hyperlink_url {
                    let bottom = y - chunk.font_size * 0.2;
                    let top = y + chunk.font_size * 0.8;
                    let merged = links
                        .last_mut()
                        .filter(|prev| prev.url == *url && (prev.rect.y1 - bottom).abs() < 1.0);
                    if let Some(prev) = merged {
                        prev.rect.x2 = x + chunk.width;
                    } else {
                        links.push(LinkAnnotation {
                            rect: Rect::new(x, bottom, x + chunk.width, top),
                            url: url.clone(),
                        });
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

        // Draw inline images outside text block.
        // Word bottom-aligns inline images of differing heights on the same line
        // (anchored to the common baseline slot). Compute the tallest image so
        // shorter ones sit on the same bottom rather than top-aligning.
        let line_max_img_h = line
            .chunks
            .iter()
            .map(|c| c.inline_image_height)
            .fold(0.0f32, f32::max);
        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            if let Some(ref img_name) = chunk.inline_image_name {
                let x = chunk_abs_x(chunk_idx, chunk);
                let img_bottom = y + chunk.font_size - line_max_img_h;

                // Pre-image effects: shadow, glow (rendered before image so they appear behind)
                let chunk_fx = chunk.inline_image_effect_xobjs.as_ref();
                if let Some(ref shadow) = chunk.inline_image_shadow {
                    super::color::draw_image_shadow(
                        content, shadow, x, img_bottom,
                        chunk.width, chunk.inline_image_height,
                        chunk_fx.and_then(|fx| fx.shadow.as_deref()),
                    );
                }
                if let Some(ref glow) = chunk.inline_image_glow {
                    super::color::draw_image_glow(
                        content, glow, x, img_bottom,
                        chunk.width, chunk.inline_image_height,
                        chunk_fx.and_then(|fx| fx.glow.as_deref()),
                    );
                }

                super::smartart::render_image_with_clip(
                    content, img_name, x, img_bottom,
                    chunk.width, chunk.inline_image_height,
                    chunk.inline_image_clip.as_ref(),
                );

                if let Some(sc) = chunk.inline_image_stroke_color {
                    super::smartart::stroke_image_border(
                        content, x, img_bottom,
                        chunk.width, chunk.inline_image_height,
                        sc, chunk.inline_image_stroke_width,
                        chunk.inline_image_clip.as_ref(),
                    );
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
}

pub(super) fn font_metric(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    get: impl Fn(&FontEntry) -> Option<f32>,
) -> Option<f32> {
    runs.first()
        .map(font_key)
        .and_then(|k| seen_fonts.get(&k))
        .and_then(get)
}

/// Compute the effective font_size, line_h_ratio, and ascender_ratio for a set of runs
/// by picking the run that produces the tallest visual ascent (font_size * ascender_ratio).
pub(super) fn tallest_run_metrics(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> (f32, Option<f32>, Option<f32>) {
    let mut best_font_size = runs.first().map_or(12.0, |r| r.font_size);
    let mut best_ascent = 0.0f32;
    let mut best_line_h_ratio: Option<f32> = None;
    let mut best_ascender_ratio: Option<f32> = None;
    let mut key_buf = String::new();

    for run in runs {
        // Line-break runs only affect the empty line they create, not
        // the paragraph's overall line height.
        if run.is_line_break {
            continue;
        }
        let key = font_key_buf(run, &mut key_buf);
        let entry = seen_fonts.get(key);
        let ar = entry.and_then(|e| e.ascender_ratio).unwrap_or(0.75);
        let ascent = run.font_size * ar;
        if ascent > best_ascent {
            best_ascent = ascent;
            best_font_size = run.font_size;
            best_ascender_ratio = entry.and_then(|e| e.ascender_ratio);
            best_line_h_ratio = entry.and_then(|e| e.line_h_ratio);
        }
    }
    (best_font_size, best_line_h_ratio, best_ascender_ratio)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::VertAlign;

    #[test]
    fn test_url_breaks_split_into_chunks() {
        let words: Vec<&str> = split_preserving_spaces(
            "see https://example.com/foo/bar?x=1&y=2#frag after",
        )
        .into_iter()
        .map(|(_, w)| w)
        .collect();
        // The URL must be broken at /, ?, #, &, = so the line wrapper
        // can split it across lines if needed.
        assert!(words.iter().any(|w| w.ends_with('/')), "no break after /: {:?}", words);
        assert!(words.iter().any(|w| w.ends_with('?')), "no break after ?: {:?}", words);
        assert!(words.iter().any(|w| w.ends_with('#')), "no break after #: {:?}", words);
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
            text: String::new(),
            font_name: "Arial".to_string(),
            font_size,
            bold: false,
            italic: false,
            underline: false,
            double_underline: false,
            strikethrough: false,
            dstrike: false,
            color: None,
            highlight: None,
            shading: None,
            border: None,
            vertical_align: valign,
            text_shadow: None,
            text_glow: None,
            text_outline: None,
            text_fill: None,
            vanish: false,
            small_caps,
            caps: false,
            char_style_id: None,
            lang: None,
            char_spacing: 0.0,
            text_scale: 100.0,
            east_asia_font_name: None,
            is_tab: false,
            is_line_break: false,
            field_code: None,
            hyperlink_url: None,
            inline_image: None,
            footnote_id: None,
            is_footnote_ref_mark: false,
            endnote_id: None,
            is_endnote_ref_mark: false,
            kern_threshold: None,
            font_size_from_default: false,
            font_name_from_default: false,
            comment_ids: Vec::new(),
        }
    }

    #[test]
    fn test_effective_font_size_baseline() {
        let run = make_run(12.0, VertAlign::Baseline, false);
        assert_eq!(effective_font_size(&run), 12.0);
    }

    #[test]
    fn test_effective_font_size_superscript() {
        let run = make_run(12.0, VertAlign::Superscript, false);
        let expected = 12.0 * 0.58; // 6.96
        assert!((effective_font_size(&run) - expected).abs() < 0.01);
    }

    #[test]
    fn test_effective_font_size_subscript() {
        let run = make_run(12.0, VertAlign::Subscript, false);
        let expected = 12.0 * 0.58;
        assert!((effective_font_size(&run) - expected).abs() < 0.01);
    }

    #[test]
    fn test_effective_font_size_ignores_small_caps() {
        // smallCaps sizing is per-segment, not per-run — effective_font_size returns base size
        let run = make_run(12.0, VertAlign::Baseline, true);
        assert_eq!(effective_font_size(&run), 12.0);
    }

    #[test]
    fn test_smallcaps_segments_mixed() {
        let segs = smallcaps_segments("Hello", 12.0);
        assert_eq!(segs.len(), 2);
        assert_eq!(segs[0], ("H".to_string(), 12.0));     // uppercase stays at 12pt
        assert_eq!(segs[1], ("ELLO".to_string(), 10.0));   // lowercase → uppercase at 10pt
    }

    #[test]
    fn test_smallcaps_segments_all_upper() {
        let segs = smallcaps_segments("ABC", 12.0);
        assert_eq!(segs.len(), 1);
        assert_eq!(segs[0], ("ABC".to_string(), 12.0));
    }

    #[test]
    fn test_smallcaps_segments_all_lower() {
        let segs = smallcaps_segments("abc", 12.0);
        assert_eq!(segs.len(), 1);
        assert_eq!(segs[0], ("ABC".to_string(), 10.0));
    }

    #[test]
    fn test_smallcaps_segments_with_nonletters() {
        // Non-letter chars (digits, punctuation) stay at base size, grouped with adjacent same-size
        let segs = smallcaps_segments("A1b", 12.0);
        assert_eq!(segs.len(), 2);
        assert_eq!(segs[0], ("A1".to_string(), 12.0));  // uppercase + digit both at base size
        assert_eq!(segs[1], ("B".to_string(), 10.0));    // lowercase → uppercase at reduced size
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
}
