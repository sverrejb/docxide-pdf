use std::collections::HashMap;

use pdf_writer::{Content, Name, Rect, Str};

use crate::fonts::{FontEntry, encode_as_gids, font_key, to_winansi_bytes};
use crate::model::{Alignment, Run, TabAlignment, TabStop, VertAlign};

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
    pub(super) y_offset: f32, // vertical offset for superscript/subscript
    pub(super) hyperlink_url: Option<String>,
    pub(super) inline_image_name: Option<String>,
    pub(super) inline_image_height: f32,
}

pub(super) struct LinkAnnotation {
    pub(super) rect: Rect,
    pub(super) url: String,
}

pub(super) struct TextLine {
    pub(super) chunks: Vec<WordChunk>,
    pub(super) total_width: f32,
}

/// True when a paragraph has no visible text (may still have phantom font-info runs).
pub(super) fn is_text_empty(runs: &[Run]) -> bool {
    runs.iter()
        .all(|r| r.text.is_empty() && !r.is_tab && r.inline_image.is_none())
}

fn effective_font_size(run: &Run) -> f32 {
    match run.vertical_align {
        VertAlign::Superscript | VertAlign::Subscript => run.font_size * 0.58,
        VertAlign::Baseline => run.font_size,
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
    }
}

/// Layout runs into wrapped lines.
/// Handles cross-run contiguous text correctly: no space is inserted between
/// runs unless the preceding text ended with whitespace or the new run starts
/// with whitespace (e.g., "bold" + ", " → "bold," not "bold ,").
pub(super) fn build_paragraph_lines(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    max_width: f32,
    first_line_hanging: f32,
    inline_image_names: &HashMap<usize, String>,
) -> Vec<TextLine> {
    let mut lines: Vec<TextLine> = Vec::new();
    let mut current_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;
    let mut prev_ended_with_ws = false;
    let mut prev_space_w: f32 = 0.0;

    for (run_idx, run) in runs.iter().enumerate() {
        if run.is_tab {
            continue; // tabs handled in build_tabbed_line
        }

        // Handle inline images as single block elements in the line
        if let Some(img) = &run.inline_image {
            if let Some(pdf_name) = inline_image_names.get(&run_idx) {
                let img_w = img.display_width;
                let need_space = !current_chunks.is_empty() && prev_ended_with_ws;
                let proposed_x = if need_space {
                    current_x + prev_space_w
                } else {
                    current_x
                };

                let line_max = if lines.is_empty() {
                    max_width + first_line_hanging
                } else {
                    max_width
                };
                if !current_chunks.is_empty() && proposed_x + img_w > line_max {
                    lines.push(finish_line(&mut current_chunks));
                    current_x = 0.0;
                } else {
                    current_x = proposed_x;
                }

                current_chunks.push(WordChunk {
                    pdf_font: String::new(),
                    text: String::new(),
                    font_size: run.font_size,
                    color: None,
                    highlight: None,
                    x_offset: current_x,
                    width: img_w,
                    underline: false,
                    strikethrough: false,
                    y_offset: 0.0,
                    hyperlink_url: None,
                    inline_image_name: Some(pdf_name.clone()),
                    inline_image_height: img.display_height,
                });
                current_x += img_w;
                prev_ended_with_ws = false;
            }
            continue;
        }

        let key = font_key(run);
        let entry = seen_fonts.get(&key).expect("font registered");
        let eff_fs = effective_font_size(run);
        let space_w = entry.widths_1000[0] * eff_fs / 1000.0;
        let starts_with_ws = run.text.starts_with(char::is_whitespace);
        let y_off = vert_y_offset(run);

        for (i, word) in run.text.split_whitespace().enumerate() {
            let ww: f32 = to_winansi_bytes(word)
                .iter()
                .filter(|&&b| b >= 32)
                .map(|&b| entry.widths_1000[(b - 32) as usize] * eff_fs / 1000.0)
                .sum();

            let need_space =
                !current_chunks.is_empty() && (i > 0 || starts_with_ws || prev_ended_with_ws);

            // Use the space width from the run that owns the space character:
            // within a run (i > 0) or leading ws → this run's space_w;
            // trailing ws from previous run → previous run's space_w
            let effective_space_w = if i > 0 || starts_with_ws {
                space_w
            } else {
                prev_space_w
            };

            let proposed_x = if need_space {
                current_x + effective_space_w
            } else {
                current_x
            };

            let line_max = if lines.is_empty() {
                max_width + first_line_hanging
            } else {
                max_width
            };
            if !current_chunks.is_empty() && proposed_x + ww > line_max {
                lines.push(finish_line(&mut current_chunks));
                current_x = 0.0;
            } else {
                current_x = proposed_x;
            }

            current_chunks.push(WordChunk {
                pdf_font: entry.pdf_name.clone(),
                text: word.to_string(),
                font_size: eff_fs,
                color: run.color,
                highlight: run.highlight,
                x_offset: current_x,
                width: ww,
                underline: run.underline,
                strikethrough: run.strikethrough,
                y_offset: y_off,
                hyperlink_url: run.hyperlink_url.clone(),
                inline_image_name: None,
                inline_image_height: 0.0,
            });
            current_x += ww;
        }

        prev_ended_with_ws = run.text.ends_with(char::is_whitespace);
        prev_space_w = space_w;
    }

    if !current_chunks.is_empty() {
        lines.push(finish_line(&mut current_chunks));
    }

    if lines.is_empty() {
        lines.push(TextLine {
            chunks: vec![],
            total_width: 0.0,
        });
    }
    lines
}

fn find_next_tab_stop(current_x: f32, tab_stops: &[TabStop], indent_left: f32) -> TabStop {
    let abs_x = current_x + indent_left;
    for stop in tab_stops {
        if stop.position > abs_x + 0.5 {
            return stop.clone();
        }
    }
    let next_default = ((abs_x / DEFAULT_TAB_INTERVAL).floor() + 1.0) * DEFAULT_TAB_INTERVAL;
    TabStop {
        position: next_default,
        alignment: TabAlignment::Left,
        leader: None,
    }
}

fn segment_width(runs: &[&Run], seen_fonts: &HashMap<String, FontEntry>) -> f32 {
    let mut w: f32 = 0.0;
    let mut first = true;
    for run in runs {
        let key = font_key(run);
        let entry = seen_fonts.get(&key).expect("font registered");
        let eff_fs = effective_font_size(run);
        let space_w = entry.widths_1000[0] * eff_fs / 1000.0;
        for (i, word) in run.text.split_whitespace().enumerate() {
            if !first || i > 0 {
                w += space_w;
            }
            w += to_winansi_bytes(word)
                .iter()
                .filter(|&&b| b >= 32)
                .map(|&b| entry.widths_1000[(b - 32) as usize] * eff_fs / 1000.0)
                .sum::<f32>();
            first = false;
        }
    }
    w
}

fn decimal_before_width(runs: &[&Run], seen_fonts: &HashMap<String, FontEntry>) -> f32 {
    let full_text: String = runs.iter().map(|r| r.text.as_str()).collect();
    let before = if let Some(dot_pos) = full_text.find('.') {
        &full_text[..dot_pos]
    } else {
        &full_text
    };
    let mut w: f32 = 0.0;
    let mut chars_remaining = before.len();
    for run in runs {
        let key = font_key(run);
        let entry = seen_fonts.get(&key).expect("font registered");
        let eff_fs = effective_font_size(run);
        let text_to_measure = if run.text.len() <= chars_remaining {
            chars_remaining -= run.text.len();
            &run.text
        } else {
            let s = &run.text[..chars_remaining];
            chars_remaining = 0;
            s
        };
        for &b in to_winansi_bytes(text_to_measure)
            .iter()
            .filter(|&&b| b >= 32)
        {
            w += entry.widths_1000[(b - 32) as usize] * eff_fs / 1000.0;
        }
        if chars_remaining == 0 {
            break;
        }
    }
    w
}

/// Build a single TextLine for a paragraph that contains tab characters.
pub(super) fn build_tabbed_line(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    tab_stops: &[TabStop],
    indent_left: f32,
) -> Vec<TextLine> {
    // Split runs into segments at tab markers
    let mut segments: Vec<(Vec<&Run>, Option<TabStop>)> = Vec::new();
    let mut current_seg: Vec<&Run> = Vec::new();
    let mut pending_tab: Option<TabStop> = None;

    for run in runs {
        if run.is_tab {
            segments.push((std::mem::take(&mut current_seg), pending_tab.take()));
            // Find which tab stop this tab activates — we'll resolve position during layout
            pending_tab = Some(TabStop {
                position: 0.0, // placeholder, resolved below
                alignment: TabAlignment::Left,
                leader: None,
            });
        } else {
            current_seg.push(run);
        }
    }
    segments.push((std::mem::take(&mut current_seg), pending_tab.take()));

    let mut all_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;

    for (seg_idx, (seg_runs, tab_before)) in segments.iter().enumerate() {
        if seg_idx > 0 {
            let stop = find_next_tab_stop(current_x, tab_stops, indent_left);
            let tab_target = stop.position - indent_left;

            // Calculate where segment text will start based on alignment
            let seg_start = match stop.alignment {
                TabAlignment::Left => tab_target.max(current_x),
                TabAlignment::Center => {
                    let sw = segment_width(seg_runs, seen_fonts);
                    (tab_target - sw / 2.0).max(current_x)
                }
                TabAlignment::Right => {
                    let sw = segment_width(seg_runs, seen_fonts);
                    (tab_target - sw).max(current_x)
                }
                TabAlignment::Decimal => {
                    let bw = decimal_before_width(seg_runs, seen_fonts);
                    (tab_target - bw).max(current_x)
                }
            };

            // Draw leader fill between end of previous text and start of aligned text
            if tab_before.is_some() {
                let abs_x = current_x + indent_left;
                let leader = tab_stops
                    .iter()
                    .find(|s| s.position > abs_x + 0.5)
                    .and_then(|s| s.leader);

                if let Some(leader_char) = leader {
                    let font_run = seg_runs.first().or_else(|| {
                        segments[..seg_idx]
                            .iter()
                            .rev()
                            .flat_map(|(r, _)| r.last())
                            .next()
                    });
                    if let Some(run) = font_run {
                        let key = font_key(run);
                        let entry = seen_fonts.get(&key).expect("font registered");
                        let eff_fs = effective_font_size(run);
                        let leader_bytes = to_winansi_bytes(&leader_char.to_string());
                        if let Some(&byte) = leader_bytes.first()
                            && byte >= 32
                        {
                            let char_w = entry.widths_1000[(byte - 32) as usize] * eff_fs / 1000.0;
                            let leader_gap = seg_start - current_x;
                            if char_w > 0.0 && leader_gap > char_w * 2.0 {
                                let count = ((leader_gap - char_w) / char_w).floor() as usize;
                                if count > 0 {
                                    let leader_text: String =
                                        std::iter::repeat_n(leader_char, count).collect();
                                    let leader_w = count as f32 * char_w;
                                    let leader_start = seg_start - leader_w;
                                    all_chunks.push(WordChunk {
                                        pdf_font: entry.pdf_name.clone(),
                                        text: leader_text,
                                        font_size: eff_fs,
                                        color: run.color,
                                        highlight: None,
                                        x_offset: leader_start,
                                        width: leader_w,
                                        underline: false,
                                        strikethrough: false,
                                        y_offset: 0.0,
                                        hyperlink_url: None,
                                        inline_image_name: None,
                                        inline_image_height: 0.0,
                                    });
                                }
                            }
                        }
                    }
                }
            }

            current_x = seg_start;
        }

        // Layout text in this segment from current_x
        let mut prev_ws = false;
        for run in seg_runs {
            let key = font_key(run);
            let entry = seen_fonts.get(&key).expect("font registered");
            let eff_fs = effective_font_size(run);
            let space_w = entry.widths_1000[0] * eff_fs / 1000.0;
            let y_off = vert_y_offset(run);

            for (i, word) in run.text.split_whitespace().enumerate() {
                let ww: f32 = to_winansi_bytes(word)
                    .iter()
                    .filter(|&&b| b >= 32)
                    .map(|&b| entry.widths_1000[(b - 32) as usize] * eff_fs / 1000.0)
                    .sum();
                if !all_chunks.is_empty()
                    && (i > 0 || prev_ws || run.text.starts_with(char::is_whitespace))
                {
                    current_x += space_w;
                }
                all_chunks.push(WordChunk {
                    pdf_font: entry.pdf_name.clone(),
                    text: word.to_string(),
                    font_size: eff_fs,
                    color: run.color,
                    highlight: run.highlight,
                    x_offset: current_x,
                    width: ww,
                    underline: run.underline,
                    strikethrough: run.strikethrough,
                    y_offset: y_off,
                    hyperlink_url: run.hyperlink_url.clone(),
                    inline_image_name: None,
                    inline_image_height: 0.0,
                });
                current_x += ww;
            }
            prev_ws = run.text.ends_with(char::is_whitespace);
        }
    }

    let total_width = all_chunks
        .last()
        .map(|c| c.x_offset + c.width)
        .unwrap_or(0.0);
    vec![TextLine {
        chunks: all_chunks,
        total_width,
    }]
}

fn encode_text_for_pdf(text: &str, pdf_font: &str, seen_fonts: &HashMap<String, FontEntry>) -> Vec<u8> {
    let entry = seen_fonts.values().find(|e| e.pdf_name == pdf_font);
    match entry.and_then(|e| e.char_to_gid.as_ref()) {
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
) {
    let mut current_color: Option<[u8; 3]> = None;
    let mut cur_font_name = String::new();
    let mut cur_font_size: f32 = -1.0;

    // Pre-compute per-line y offsets accounting for inline images making lines taller
    let mut line_y_offsets: Vec<f32> = Vec::with_capacity(lines.len());
    let mut cumulative_y = 0.0f32;
    for (i, line) in lines.iter().enumerate() {
        line_y_offsets.push(cumulative_y);
        let img_h = line.chunks.iter()
            .map(|c| c.inline_image_height)
            .fold(0.0f32, f32::max);
        cumulative_y += if img_h > line_pitch { img_h } else { line_pitch };
        // First line offset is always 0
        if i == 0 { cumulative_y = line_pitch.max(img_h); line_y_offsets[0] = 0.0; }
    }

    let last_line_idx = total_line_count.saturating_sub(1);
    for (line_num, line) in lines.iter().enumerate() {
        let y = first_baseline_y - line_y_offsets[line_num];
        let global_line_idx = first_line_index + line_num;

        let is_justified = *alignment == Alignment::Justify
            && global_line_idx != last_line_idx
            && line.chunks.len() > 1;

        let (eff_margin, eff_width) = if global_line_idx == 0 && first_line_hanging > 0.0 {
            (margin_left - first_line_hanging, text_width + first_line_hanging)
        } else {
            (margin_left, text_width)
        };

        let line_start_x = match alignment {
            Alignment::Center => eff_margin + (eff_width - line.total_width) / 2.0,
            Alignment::Right => eff_margin + eff_width - line.total_width,
            Alignment::Left | Alignment::Justify => eff_margin,
        };

        let extra_per_gap = if is_justified {
            (eff_width - line.total_width) / (line.chunks.len() - 1) as f32
        } else {
            0.0
        };

        let mut decorations: Vec<(f32, f32, f32, f32, Option<[u8; 3]>)> = Vec::new();
        let mut image_draws: Vec<(f32, f32, f32, f32, &str)> = Vec::new();

        // Draw run highlights as merged spans (contiguous same-color highlights)
        {
            let mut hl_start_x = 0.0f32;
            let mut hl_color: Option<[u8; 3]> = None;
            let mut hl_end_x = 0.0f32;
            let mut hl_fs = 0.0f32;

            let flush_hl = |content: &mut Content,
                            color: [u8; 3],
                            sx: f32,
                            ex: f32,
                            fs: f32,
                            y: f32| {
                let hl_bottom = y - fs * 0.2;
                let hl_height = fs * 1.15;
                content.save_state();
                content.set_fill_rgb(
                    color[0] as f32 / 255.0,
                    color[1] as f32 / 255.0,
                    color[2] as f32 / 255.0,
                );
                content.rect(sx, hl_bottom, ex - sx, hl_height);
                content.fill_nonzero();
                content.restore_state();
            };

            for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
                let x = line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap;
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

        let has_text_chunks = line.chunks.iter().any(|c| c.inline_image_name.is_none() && !c.text.is_empty());

        if has_text_chunks {
            content.begin_text();
            let mut td_x = 0.0_f32;
            let mut td_y = 0.0_f32;

            for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
                if chunk.inline_image_name.is_some() {
                    continue;
                }

                let x = line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap;
                let cy = y + chunk.y_offset;

                if chunk.color != current_color {
                    if let Some([r, g, b]) = chunk.color {
                        content.set_fill_rgb(
                            r as f32 / 255.0,
                            g as f32 / 255.0,
                            b as f32 / 255.0,
                        );
                    } else {
                        content.set_fill_gray(0.0);
                    }
                    current_color = chunk.color;
                }

                if cur_font_name != chunk.pdf_font || cur_font_size != chunk.font_size {
                    content.set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size);
                    cur_font_name.clear();
                    cur_font_name.push_str(&chunk.pdf_font);
                    cur_font_size = chunk.font_size;
                }

                content.next_line(x - td_x, cy - td_y);
                td_x = x;
                td_y = cy;

                let text_bytes =
                    encode_text_for_pdf(&chunk.text, &chunk.pdf_font, seen_fonts);
                content.show(Str(&text_bytes));

                if chunk.underline {
                    let thick = (chunk.font_size * 0.05).max(0.5);
                    let ul_y = if chunk.hyperlink_url.is_some() {
                        y - chunk.font_size * 0.08
                    } else {
                        y - chunk.font_size * 0.12
                    };
                    decorations.push((x, ul_y - thick, chunk.width, thick, chunk.color));
                }
                if chunk.strikethrough {
                    let thick = (chunk.font_size * 0.05).max(0.5);
                    let st_y = y + chunk.font_size * 0.3;
                    decorations.push((x, st_y, chunk.width, thick, chunk.color));
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
            content.end_text();
        }

        // Collect inline image draws
        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            if let Some(ref img_name) = chunk.inline_image_name {
                let x = line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap;
                let img_bottom = y - (chunk.inline_image_height - chunk.font_size);
                image_draws.push((x, img_bottom, chunk.width, chunk.inline_image_height, img_name));
            }
        }

        // Draw inline images outside text block
        for &(ix, iy, iw, ih, ref img_name) in &image_draws {
            content.save_state();
            content.transform([iw, 0.0, 0.0, ih, ix, iy]);
            content.x_object(Name(img_name.as_bytes()));
            content.restore_state();
        }

        for &(dx, dy, dw, dh, dcolor) in &decorations {
            if dcolor != current_color {
                if let Some([r, g, b]) = dcolor {
                    content.set_fill_rgb(
                        r as f32 / 255.0,
                        g as f32 / 255.0,
                        b as f32 / 255.0,
                    );
                } else {
                    content.set_fill_gray(0.0);
                }
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

    for run in runs {
        let key = font_key(run);
        let entry = seen_fonts.get(&key);
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
