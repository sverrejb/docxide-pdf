use std::borrow::Cow;
use std::collections::HashMap;

use pdf_writer::types::TextRenderingMode;
use pdf_writer::{Content, Name, Rect, Str};

use crate::fonts::{FontEntry, encode_as_gids, font_key, font_key_buf, to_winansi_bytes};
use crate::model::{Alignment, Run, TabAlignment, TabStop, TextFill, TextOutline, VertAlign};

use super::color::{fill_color_or_black, stroke_color_or_black, stroke_rgb};

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
    let breaks: Vec<usize> = unicode_linebreak::linebreaks(text)
        .map(|(pos, _)| pos)
        .collect();

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
    pub(super) x_offset: f32, // x relative to line start
    pub(super) width: f32,
    pub(super) underline: bool,
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
    pub(super) synthetic_bold: bool,
    pub(super) text_outline: Option<TextOutline>,
    pub(super) text_fill: Option<TextFill>,
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
            x_offset,
            width,
            underline: run.underline,
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
            synthetic_bold: entry.synthetic_bold,
            text_outline: run.text_outline.clone(),
            text_fill: run.text_fill.clone(),
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
    ) -> Self {
        Self {
            pdf_font: String::new(),
            text: String::new(),
            font_size,
            color: None,
            highlight: None,
            x_offset,
            width: display_width,
            underline: false,
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
            synthetic_bold: false,
            text_outline: None,
            text_fill: None,
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
            x_offset,
            width,
            underline: false,
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
            synthetic_bold: false,
            text_outline: None,
            text_fill: None,
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
    let base = match run.vertical_align {
        VertAlign::Superscript | VertAlign::Subscript => run.font_size * 0.58,
        VertAlign::Baseline => run.font_size,
    };
    if run.small_caps {
        (base - 2.0).max(1.0)
    } else {
        base
    }
}

/// Returns the text to use for layout/rendering, applying caps/smallCaps transforms.
fn effective_text(run: &Run) -> Cow<'_, str> {
    if run.caps || run.small_caps {
        Cow::Owned(run.text.to_uppercase())
    } else {
        Cow::Borrowed(&run.text)
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
    width_after_line: Option<(usize, f32)>,
    per_line_widths: Option<&[f32]>,
    per_line_dual: Option<&[DualRegion]>,
) -> Vec<TextLine> {
    let mut lines: Vec<TextLine> = Vec::new();
    let mut current_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;
    let mut pending_space_w: f32 = 0.0;
    let mut key_buf = String::new();
    // Track whether we're filling the right region on the current line
    let mut in_right_region = false;
    // Info for the right region of the current line being built
    let mut cur_right_info: Option<(usize, f32, f32)> = None; // (first_chunk_idx, region_x, region_w)

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
                            current_x = 0.0;
                            pending_space_w = 0.0;
                            // Retry placement in right region
                            let proposed_x2 = 0.0;
                            if proposed_x2 + img_w <= rw {
                                current_chunks.push(WordChunk::image(
                                    pdf_name, run.font_size, proposed_x2, img_w, img.display_height,
                                    img.stroke_color, img.stroke_width, img.shadow.clone(),
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

            // A word continues the previous word across a run boundary when
            // there is no whitespace between runs and no leading spaces.
            let is_continuation = is_first_word_in_run
                && space_count == 0
                && pending_space_w == 0.0
                && !current_chunks.is_empty();
            is_first_word_in_run = false;

            let char_count = word.chars().count();
            let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
            let ww = entry.word_width(word, eff_fs, kern) * ts + cs * char_count as f32;

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

            let overflows = proposed_x + ww > cur_max;
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
                        current_x = 0.0;
                        pending_space_w = 0.0;
                        // Place this word at the start of the right region
                        current_chunks.push(WordChunk::text(
                            entry, run, word, eff_fs, cs, y_off, 0.0, ww,
                        ));
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
                        current_chunks.push(WordChunk::text(
                            entry, run, word, eff_fs, cs, y_off, 0.0, ww,
                        ));
                        current_x = ww;
                        continue;
                    }
                }
            } else {
                current_x = proposed_x;
            }
            pending_space_w = 0.0;

            current_chunks.push(WordChunk::text(
                entry, run, word, eff_fs, cs, y_off, current_x, ww,
            ));
            current_x += ww;
        }

        // Accumulate trailing whitespace for the next run
        let trailing_spaces = text.chars().rev().take_while(|c| is_break_space(*c)).count();
        pending_space_w += trailing_spaces as f32 * space_w_cs;
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
        let space_w = entry.space_width(eff_fs) * ts;
        let text = effective_text(run);
        for (i, word) in text.split_whitespace().enumerate() {
            if !first || i > 0 {
                w += space_w;
            }
            let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
            w += entry.word_width(word, eff_fs, kern) * ts;
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
        let text_to_measure = if text.len() <= chars_remaining {
            chars_remaining -= text.len();
            text.as_ref()
        } else {
            let s = &text[..chars_remaining];
            chars_remaining = 0;
            s
        };
        let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
        w += entry.word_width(text_to_measure, eff_fs, kern);
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
    first_line_hanging: f32,
    inline_image_names: &HashMap<usize, String>,
    default_tab_stop: f32,
) -> Vec<TextLine> {
    // Split runs into segments at tab markers, tracking original run indices
    let mut segments: Vec<(Vec<&Run>, Vec<usize>, Option<TabStop>)> = Vec::new();
    let mut current_seg: Vec<&Run> = Vec::new();
    let mut current_indices: Vec<usize> = Vec::new();
    let mut pending_tab: Option<TabStop> = None;

    for (global_idx, run) in runs.iter().enumerate() {
        if run.vanish {
            continue;
        }
        if run.is_tab {
            segments.push((
                std::mem::take(&mut current_seg),
                std::mem::take(&mut current_indices),
                pending_tab.take(),
            ));
            pending_tab = Some(TabStop {
                position: 0.0, // placeholder, resolved below
                alignment: TabAlignment::Left,
                leader: None,
            });
        } else {
            current_seg.push(run);
            current_indices.push(global_idx);
        }
    }
    segments.push((
        std::mem::take(&mut current_seg),
        std::mem::take(&mut current_indices),
        pending_tab.take(),
    ));

    let mut result_lines: Vec<TextLine> = Vec::new();
    let mut all_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;
    let mut key_buf = String::new();
    let mut is_first_line = true;

    for (seg_idx, (seg_runs, seg_indices, tab_before)) in segments.iter().enumerate() {
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
            let stop = find_next_tab_stop(current_x, tab_stops, line_indent, default_tab_stop);
            let mut effective_tab_target = stop.position - line_indent;
            let mut seg_start =
                resolve_tab_aligned_start(&stop, effective_tab_target, seg_runs, seen_fonts, current_x);
            let mut resolved_leader = stop.leader;

            if seg_start > line_max && !all_chunks.is_empty() {
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

                if let Some(leader_char) = leader {
                    let font_run: Option<&Run> = seg_runs.first().copied().or_else(|| {
                        segments[..seg_idx]
                            .iter()
                            .rev()
                            .flat_map(|(r, _, _)| r.last().copied())
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
        let mut prev_ws = false;
        for (local_idx, run) in seg_runs.iter().enumerate() {
            if run.is_line_break {
                result_lines.push(finish_line_with_break(&mut all_chunks));
                current_x = 0.0;
                is_first_line = false;
                prev_ws = false;
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
            let segments = split_preserving_spaces(&text);
            for (seg_idx, &(space_count, word)) in segments.iter().enumerate() {
                let char_count = word.chars().count();
                let kern = run.kern_threshold.is_some_and(|t| eff_fs >= t);
                let ww = entry.word_width(word, eff_fs, kern) * ts + cs * char_count as f32;
                let has_space_before =
                    space_count > 0 || (seg_idx == 0 && (prev_ws || text.starts_with(is_break_space)));
                if !all_chunks.is_empty() && has_space_before {
                    current_x += space_w * ts + cs;
                }
                // Word continues previous word across run boundary (no whitespace between)
                let is_continuation = seg_idx == 0 && !has_space_before && !all_chunks.is_empty();
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
                all_chunks.push(WordChunk::text(
                    entry, run, word, eff_fs, cs, y_off, current_x, ww,
                ));
                current_x += ww;
            }
            prev_ws = text.ends_with(is_break_space);
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
) {
    let mut current_color: Option<[u8; 3]> = None;
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

        let is_justified = *alignment == Alignment::Justify
            && global_line_idx != last_line_idx
            && !line.ends_with_break
            && left_chunk_count > 1;

        let line_start_x = match alignment {
            Alignment::Center => eff_margin + (eff_width - left_content_width) / 2.0,
            Alignment::Right => eff_margin + eff_width - left_content_width,
            Alignment::Left | Alignment::Justify => eff_margin,
        };

        let extra_per_gap = if is_justified {
            (eff_width - left_content_width) / (left_chunk_count - 1).max(1) as f32
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
                (rr.region_width - rr.content_width) / (right_chunks - 1) as f32
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
            line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap
        };

        let mut decorations: Vec<(f32, f32, f32, f32, Option<[u8; 3]>)> = Vec::new();

        // Draw run highlights as merged spans (contiguous same-color highlights)
        {
            let mut hl_start_x = 0.0f32;
            let mut hl_color: Option<[u8; 3]> = None;
            let mut hl_end_x = 0.0f32;
            let mut hl_fs = 0.0f32;

            let flush_hl =
                |content: &mut Content, color: [u8; 3], sx: f32, ex: f32, fs: f32, y: f32| {
                    let hl_bottom = y - fs * 0.2;
                    let hl_height = fs * 1.15;
                    content.save_state();
                    fill_color_or_black(content, Some(color));
                    content.rect(sx, hl_bottom, ex - sx, hl_height);
                    content.fill_nonzero();
                    content.restore_state();
                };

            for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
                let x = chunk_abs_x(chunk_idx, chunk);
                if chunk.highlight == hl_color && hl_color.is_some() {
                    hl_end_x = x + chunk.width;
                    hl_fs = hl_fs.max(chunk.font_size);
                } else {
                    if let Some(c) = hl_color {
                        flush_hl(content, c, hl_start_x, hl_end_x, hl_fs, y);
                    }
                    if let Some(c) = chunk.highlight {
                        hl_start_x = x;
                        hl_end_x = x + chunk.width;
                        hl_fs = chunk.font_size;
                        hl_color = Some(c);
                    } else {
                        hl_color = None;
                    }
                }
            }
            if let Some(c) = hl_color {
                flush_hl(content, c, hl_start_x, hl_end_x, hl_fs, y);
            }
        }

        let has_text_chunks = line
            .chunks
            .iter()
            .any(|c| c.inline_image_name.is_none() && !c.text.is_empty());

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

                // w14:textFill solid color overrides w:color
                let effective_color = match chunk.text_fill {
                    Some(TextFill::Solid(c)) => Some(c),
                    _ => chunk.color,
                };
                if effective_color != current_color {
                    fill_color_or_black(content, effective_color);
                    current_color = effective_color;
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

                if chunk.char_spacing != cur_char_spacing {
                    content.set_char_spacing(chunk.char_spacing);
                    cur_char_spacing = chunk.char_spacing;
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
                    let primary_gids = primary.char_to_gid.as_ref();
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
                    // Merge with previous underline decoration if same y and color
                    let merged = decorations.last_mut().filter(|(_, dy, _, dh, dc)| {
                        (*dy - ul_top).abs() < 0.01
                            && (*dh - thick).abs() < 0.01
                            && *dc == chunk.color
                    });
                    if let Some(prev) = merged {
                        prev.2 = (x + chunk.width) - prev.0;
                    } else {
                        decorations.push((x, ul_top, chunk.width, thick, chunk.color));
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
            if cur_char_spacing != 0.0 {
                content.set_char_spacing(0.0);
                cur_char_spacing = 0.0;
            }
            if cur_text_scale != 100.0 {
                content.set_horizontal_scaling(100.0);
                cur_text_scale = 100.0;
            }
            content.end_text();
        }

        // Draw inline images outside text block
        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            if let Some(ref img_name) = chunk.inline_image_name {
                let x = chunk_abs_x(chunk_idx, chunk);
                let img_bottom = y - (chunk.inline_image_height - chunk.font_size);

                // Drop shadow (rendered before image so it appears behind)
                if let Some(ref shadow) = chunk.inline_image_shadow {
                    super::color::draw_image_shadow(
                        content, shadow, x, img_bottom,
                        chunk.width, chunk.inline_image_height,
                    );
                }

                content.save_state();
                content.transform([
                    chunk.width,
                    0.0,
                    0.0,
                    chunk.inline_image_height,
                    x,
                    img_bottom,
                ]);
                content.x_object(Name(img_name.as_bytes()));
                content.restore_state();

                if let Some(sc) = chunk.inline_image_stroke_color {
                    content.save_state();
                    stroke_rgb(content, sc);
                    content.set_line_width(chunk.inline_image_stroke_width);
                    content.rect(x, img_bottom, chunk.width, chunk.inline_image_height);
                    content.stroke();
                    content.restore_state();
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
