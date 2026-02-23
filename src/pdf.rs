use std::collections::HashMap;

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str};

use crate::error::Error;
use crate::fonts::{FontEntry, font_key, register_font, to_winansi_bytes};
use crate::model::{
    Alignment, Block, CellVAlign, Document, FieldCode, HeaderFooter, ImageFormat, Run,
    TabAlignment, TabStop, Table, VMerge, VertAlign,
};

struct WordChunk {
    pdf_font: String,
    text: String,
    font_size: f32,
    color: Option<[u8; 3]>,
    x_offset: f32, // x relative to line start
    width: f32,
    underline: bool,
    strikethrough: bool,
    y_offset: f32, // vertical offset for superscript/subscript
    hyperlink_url: Option<String>,
}

struct LinkAnnotation {
    rect: Rect,
    url: String,
}

/// True when a paragraph has no visible text (may still have phantom font-info runs).
fn is_text_empty(runs: &[Run]) -> bool {
    runs.iter().all(|r| r.text.is_empty() && !r.is_tab)
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

struct TextLine {
    chunks: Vec<WordChunk>,
    total_width: f32,
}

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
fn build_paragraph_lines(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    max_width: f32,
    first_line_hanging: f32,
) -> Vec<TextLine> {
    let mut lines: Vec<TextLine> = Vec::new();
    let mut current_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;
    let mut prev_ended_with_ws = false;
    let mut prev_space_w: f32 = 0.0;

    for run in runs {
        if run.is_tab {
            continue; // tabs handled in build_tabbed_line
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
                x_offset: current_x,
                width: ww,
                underline: run.underline,
                strikethrough: run.strikethrough,
                y_offset: y_off,
                hyperlink_url: run.hyperlink_url.clone(),
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
fn build_tabbed_line(
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
                                        x_offset: leader_start,
                                        width: leader_w,
                                        underline: false,
                                        strikethrough: false,
                                        y_offset: 0.0,
                                        hyperlink_url: None,
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
                    x_offset: current_x,
                    width: ww,
                    underline: run.underline,
                    strikethrough: run.strikethrough,
                    y_offset: y_off,
                    hyperlink_url: run.hyperlink_url.clone(),
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

/// Render pre-built lines applying the paragraph alignment.
/// `total_line_count` is the full paragraph line count (for justify: last line stays left-aligned).
fn render_paragraph_lines(
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
) {
    let mut current_color: Option<[u8; 3]> = None;

    let last_line_idx = total_line_count.saturating_sub(1);
    for (line_num, line) in lines.iter().enumerate() {
        let y = first_baseline_y - line_num as f32 * line_pitch;
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

        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            let x = line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap;
            if chunk.color != current_color {
                if let Some([r, g, b]) = chunk.color {
                    content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                } else {
                    content.set_fill_gray(0.0);
                }
                current_color = chunk.color;
            }
            let text_bytes = to_winansi_bytes(&chunk.text);
            content
                .begin_text()
                .set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size)
                .next_line(x, y + chunk.y_offset)
                .show(Str(&text_bytes))
                .end_text();

            if chunk.underline {
                let thick = (chunk.font_size * 0.05).max(0.5);
                let ul_y = if chunk.hyperlink_url.is_some() {
                    y - chunk.font_size * 0.08
                } else {
                    y - chunk.font_size * 0.12
                };
                content
                    .rect(x, ul_y - thick, chunk.width, thick)
                    .fill_nonzero();
            }
            if chunk.strikethrough {
                let thick = (chunk.font_size * 0.05).max(0.5);
                let st_y = y + chunk.font_size * 0.3;
                content.rect(x, st_y, chunk.width, thick).fill_nonzero();
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
    }
    if current_color.is_some() {
        content.set_fill_gray(0.0);
    }
}

fn font_metric(
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
fn tallest_run_metrics(
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

/// Auto-fit column widths so that the longest non-breakable word in each column
/// fits within the cell (including padding). Columns that need more space grow;
/// other columns shrink proportionally. Total width is preserved.
fn auto_fit_columns(table: &Table, seen_fonts: &HashMap<String, FontEntry>) -> Vec<f32> {
    let ncols = table.col_widths.len();
    if ncols == 0 {
        return table.col_widths.clone();
    }

    let mut min_widths = vec![0.0f32; ncols];

    for row in &table.rows {
        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            if grid_col >= ncols || span > 1 {
                grid_col += span;
                continue;
            }
            for para in &cell.paragraphs {
                for run in &para.runs {
                    let key = font_key(run);
                    let Some(entry) = seen_fonts.get(&key) else {
                        continue;
                    };
                    for word in run.text.split_whitespace() {
                        let ww: f32 = to_winansi_bytes(word)
                            .iter()
                            .filter(|&&b| b >= 32)
                            .map(|&b| entry.widths_1000[(b - 32) as usize] * run.font_size / 1000.0)
                            .sum();
                        min_widths[grid_col] = min_widths[grid_col].max(ww);
                    }
                }
            }
            grid_col += span;
        }
    }

    let total: f32 = table.col_widths.iter().sum();
    let mut widths = table.col_widths.clone();

    // Expand columns that need it, track how much extra space is needed
    let mut extra_needed: f32 = 0.0;
    let mut shrinkable: f32 = 0.0;
    for i in 0..ncols {
        if min_widths[i] > widths[i] {
            extra_needed += min_widths[i] - widths[i];
            widths[i] = min_widths[i];
        } else {
            shrinkable += widths[i] - min_widths[i];
        }
    }

    if extra_needed > 0.0 && shrinkable > 0.0 {
        let factor = extra_needed.min(shrinkable) / shrinkable;
        for i in 0..ncols {
            if widths[i] > min_widths[i] {
                let available = widths[i] - min_widths[i];
                widths[i] -= available * factor;
            }
        }
        // Normalize to preserve total
        let new_total: f32 = widths.iter().sum();
        if (new_total - total).abs() > 0.01 {
            let scale = total / new_total;
            for w in &mut widths {
                *w *= scale;
            }
        }
    }

    widths
}

struct RowLayout {
    height: f32,
    cell_lines: Vec<(Vec<TextLine>, f32, f32)>, // (lines, line_h, font_size) per cell
}

fn compute_row_layouts(
    table: &Table,
    col_widths: &[f32],
    doc: &Document,
    seen_fonts: &HashMap<String, FontEntry>,
) -> Vec<RowLayout> {
    let cm = &table.cell_margins;
    table
        .rows
        .iter()
        .map(|row| {
            let mut max_h: f32 = 0.0;
            let mut grid_col = 0usize;
            let cell_lines: Vec<(Vec<TextLine>, f32, f32)> = row
                .cells
                .iter()
                .map(|cell| {
                    let span = cell.grid_span.max(1) as usize;
                    let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
                        .iter()
                        .sum::<f32>()
                        .max(cell.width);
                    grid_col += span;

                    if cell.v_merge == VMerge::Continue {
                        return (vec![], 14.4, 12.0);
                    }

                    let cell_text_w = (col_w - cm.left - cm.right).max(0.0);
                    let mut total_h: f32 = cm.top + cm.bottom;
                    let mut all_lines = Vec::new();
                    let mut first_font_size = 12.0f32;
                    let mut first_line_h = 14.4f32;

                    for para in &cell.paragraphs {
                        let font_size = para.runs.first().map_or(12.0, |r| r.font_size);
                        let effective_ls = para.line_spacing.unwrap_or(doc.line_spacing);
                        let line_h = font_metric(&para.runs, seen_fonts, |e| e.line_h_ratio)
                            .map(|ratio| font_size * ratio * effective_ls)
                            .unwrap_or(font_size * 1.2 * effective_ls);

                        if all_lines.is_empty() {
                            first_font_size = font_size;
                            first_line_h = line_h;
                        }

                        if !is_text_empty(&para.runs) {
                            let lines = build_paragraph_lines(&para.runs, seen_fonts, cell_text_w, 0.0);
                            total_h += lines.len() as f32 * line_h;
                            all_lines.extend(lines);
                        }
                    }

                    max_h = max_h.max(total_h);
                    (all_lines, first_line_h, first_font_size)
                })
                .collect();

            // Word's row height includes the end-of-cell paragraph mark glyph,
            // adding roughly 0.5pt beyond the content metrics.
            let content_h = max_h + 0.5;
            let height = match (row.height, row.height_exact) {
                (Some(h), true) => h,
                (Some(h), false) => content_h.max(h),
                _ => content_h,
            };

            RowLayout {
                height,
                cell_lines,
            }
        })
        .collect()
}

fn render_table(
    table: &Table,
    doc: &Document,
    seen_fonts: &HashMap<String, FontEntry>,
    content: &mut Content,
    all_contents: &mut Vec<Content>,
    all_page_links: &mut Vec<Vec<LinkAnnotation>>,
    current_page_links: &mut Vec<LinkAnnotation>,
    slot_top: &mut f32,
    prev_space_after: f32,
) {
    let col_widths = auto_fit_columns(table, seen_fonts);
    let row_layouts = compute_row_layouts(table, &col_widths, doc, seen_fonts);
    let cm = &table.cell_margins;
    // Word positions tables so cell text aligns with the paragraph margin
    let table_left = doc.margin_left + table.table_indent - cm.left;

    *slot_top -= prev_space_after;

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        log::debug!(
            "TABLE row={} row_h={:.2} cells={} slot_top={:.2}",
            ri,
            row_h,
            layout.cell_lines.len(),
            *slot_top
        );
        let at_page_top = (*slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;

        if !at_page_top && *slot_top - row_h < doc.margin_bottom {
            all_contents.push(std::mem::replace(content, Content::new()));
            all_page_links.push(std::mem::take(current_page_links));
            *slot_top = doc.page_height - doc.margin_top;
        }

        let row_top = *slot_top;
        let row_bottom = row_top - row_h;

        let mut grid_col = 0usize;
        for (cell, (lines, line_h, font_size)) in
            row.cells.iter().zip(layout.cell_lines.iter())
        {
            let span = cell.grid_span.max(1) as usize;
            let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
                .iter()
                .sum();
            let cell_x = table_left
                + col_widths[..grid_col.min(col_widths.len())]
                    .iter()
                    .sum::<f32>();
            grid_col += span;

            if cell.v_merge == VMerge::Continue {
                continue;
            }

            if let Some([r, g, b]) = cell.shading {
                let b_borders = &cell.borders;
                let inset = (b_borders.top.width + b_borders.bottom.width
                    + b_borders.left.width + b_borders.right.width)
                    / 8.0;
                content.save_state();
                content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                content.rect(
                    cell_x + inset,
                    row_bottom + inset,
                    col_w - 2.0 * inset,
                    row_h - 2.0 * inset,
                );
                content.fill_nonzero();
                content.restore_state();
            }

            if !lines.is_empty() && !lines.iter().all(|l| l.chunks.is_empty()) {
                let text_x = cell_x + cm.left;
                let text_w = (col_w - cm.left - cm.right).max(0.0);
                let first_run = cell.paragraphs.first().and_then(|p| p.runs.first());
                let ascender_ratio = first_run
                    .map(font_key)
                    .and_then(|k| seen_fonts.get(&k))
                    .and_then(|e| e.ascender_ratio)
                    .unwrap_or(0.75);

                let content_h = lines.len() as f32 * line_h;
                let baseline_y = match cell.v_align {
                    CellVAlign::Top => row_top - cm.top - font_size * ascender_ratio,
                    CellVAlign::Center => {
                        let avail = row_h - cm.top - cm.bottom;
                        let offset = (avail - content_h) / 2.0;
                        row_top - cm.top - offset.max(0.0) - font_size * ascender_ratio
                    }
                    CellVAlign::Bottom => {
                        let avail = row_h - cm.top - cm.bottom;
                        let offset = avail - content_h;
                        row_top - cm.top - offset.max(0.0) - font_size * ascender_ratio
                    }
                };

                let alignment = cell
                    .paragraphs
                    .first()
                    .map(|p| p.alignment)
                    .unwrap_or(Alignment::Left);

                render_paragraph_lines(
                    content,
                    lines,
                    &alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    *line_h,
                    lines.len(),
                    0,
                    &mut Vec::new(),
                    0.0,
                );
            }
        }

        // Draw per-cell borders with color/width
        let mut grid_col = 0usize;
        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let col_w: f32 = col_widths[grid_col..col_widths.len().min(grid_col + span)]
                .iter()
                .sum();
            let bx = table_left
                + col_widths[..grid_col.min(col_widths.len())]
                    .iter()
                    .sum::<f32>();
            grid_col += span;

            let b = &cell.borders;
            let draw_border =
                |content: &mut Content, border: &crate::model::CellBorder, x1, y1, x2, y2| {
                    if !border.present {
                        return;
                    }
                    content.save_state();
                    content.set_line_width(border.width);
                    if let Some([r, g, b]) = border.color {
                        content
                            .set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                    }
                    content.move_to(x1, y1);
                    content.line_to(x2, y2);
                    content.stroke();
                    content.restore_state();
                };

            if cell.v_merge != VMerge::Continue {
                draw_border(content, &b.top, bx, row_top, bx + col_w, row_top);
            }
            draw_border(content, &b.bottom, bx, row_bottom, bx + col_w, row_bottom);
            draw_border(content, &b.left, bx, row_top, bx, row_bottom);
            draw_border(content, &b.right, bx + col_w, row_top, bx + col_w, row_bottom);
        }

        *slot_top = row_bottom;
    }
}

fn render_header_footer(
    content: &mut Content,
    hf: &HeaderFooter,
    seen_fonts: &HashMap<String, FontEntry>,
    doc: &Document,
    is_header: bool,
    page_num: usize,
    total_pages: usize,
) {
    let text_width = doc.page_width - doc.margin_left - doc.margin_right;

    for para in &hf.paragraphs {
        if is_text_empty(&para.runs) {
            continue;
        }

        let substituted_runs: Vec<Run> = para
            .runs
            .iter()
            .map(|run| {
                let mut r = run.clone();
                r.field_code = None;
                if let Some(ref fc) = run.field_code {
                    r.text = match fc {
                        FieldCode::Page => page_num.to_string(),
                        FieldCode::NumPages => total_pages.to_string(),
                    };
                }
                r
            })
            .collect();

        let lines = build_paragraph_lines(&substituted_runs, seen_fonts, text_width, 0.0);

        let (font_size, _, tallest_ar) = tallest_run_metrics(&substituted_runs, seen_fonts);
        let ascender_ratio = tallest_ar.unwrap_or(0.75);

        let baseline_y = if is_header {
            doc.page_height - doc.header_margin - font_size * ascender_ratio
        } else {
            doc.footer_margin + font_size * (1.0 - ascender_ratio)
        };

        let effective_ls = para.line_spacing.unwrap_or(doc.line_spacing);
        let line_h = font_metric(&substituted_runs, seen_fonts, |e| e.line_h_ratio)
            .map(|ratio| font_size * ratio * effective_ls)
            .unwrap_or(font_size * 1.2 * effective_ls);

        render_paragraph_lines(
            content,
            &lines,
            &para.alignment,
            doc.margin_left,
            text_width,
            baseline_y,
            line_h,
            lines.len(),
            0,
            &mut Vec::new(),
            0.0,
        );
    }
}

pub fn render(doc: &Document) -> Result<Vec<u8>, Error> {
    let t0 = std::time::Instant::now();
    let mut pdf = Pdf::new();
    let mut next_id = 1i32;
    let mut alloc = || {
        let r = Ref::new(next_id);
        next_id += 1;
        r
    };

    let catalog_id = alloc();
    let pages_id = alloc();

    // Phase 1: collect unique font names (with variant) and embed them
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();

    // Collect all runs from all blocks (paragraphs, table cells, headers/footers)
    let hf_options = [
        &doc.header_default,
        &doc.header_first,
        &doc.footer_default,
        &doc.footer_first,
    ];
    let hf_runs = hf_options
        .iter()
        .filter_map(|hf| hf.as_ref())
        .flat_map(|hf| hf.paragraphs.iter())
        .flat_map(|p| p.runs.iter());

    let all_runs: Vec<&Run> = doc
        .blocks
        .iter()
        .flat_map(|block| -> Box<dyn Iterator<Item = &Run> + '_> {
            match block {
                Block::Paragraph(para) => Box::new(para.runs.iter()),
                Block::Table(table) => Box::new(
                    table
                        .rows
                        .iter()
                        .flat_map(|row| row.cells.iter())
                        .flat_map(|cell| cell.paragraphs.iter())
                        .flat_map(|para| para.runs.iter()),
                ),
            }
        })
        .chain(hf_runs)
        .collect();

    let t_collect = t0.elapsed();

    for run in &all_runs {
        let key = font_key(run);
        if !seen_fonts.contains_key(&key) {
            let pdf_name = format!("F{}", font_order.len() + 1);
            let entry = register_font(
                &mut pdf,
                &run.font_name,
                run.bold,
                run.italic,
                pdf_name,
                &mut alloc,
                &doc.embedded_fonts,
            );
            seen_fonts.insert(key.clone(), entry);
            font_order.push(key);
        }
    }

    if seen_fonts.is_empty() {
        let pdf_name = "F1".to_string();
        let entry = register_font(
            &mut pdf,
            "Helvetica",
            false,
            false,
            pdf_name,
            &mut alloc,
            &doc.embedded_fonts,
        );
        seen_fonts.insert("Helvetica".to_string(), entry);
        font_order.push("Helvetica".to_string());
    }

    let t_fonts = t0.elapsed();

    let text_width = doc.page_width - doc.margin_left - doc.margin_right;

    // Phase 1b: embed images
    let mut image_pdf_names: HashMap<usize, String> = HashMap::new();
    let mut image_xobjects: Vec<(String, Ref)> = Vec::new();
    for (block_idx, block) in doc.blocks.iter().enumerate() {
        if let Block::Paragraph(para) = block
            && let Some(img) = &para.image
        {
            let xobj_ref = alloc();
            let pdf_name = format!("Im{}", image_xobjects.len() + 1);

            match img.format {
                ImageFormat::Jpeg => {
                    let mut xobj = pdf.image_xobject(xobj_ref, &img.data);
                    xobj.filter(Filter::DctDecode);
                    xobj.width(img.pixel_width as i32);
                    xobj.height(img.pixel_height as i32);
                    xobj.color_space().device_rgb();
                    xobj.bits_per_component(8);
                }
                ImageFormat::Png => {
                    let cursor = std::io::Cursor::new(&img.data);
                    let reader = image::ImageReader::with_format(
                        std::io::BufReader::new(cursor),
                        image::ImageFormat::Png,
                    );
                    if let Ok(decoded) = reader.decode() {
                        let rgba = decoded.to_rgba8();
                        let (w, h) = (rgba.width(), rgba.height());
                        let has_alpha = rgba.pixels().any(|p| p.0[3] < 255);

                        let rgb_data: Vec<u8> = rgba
                            .pixels()
                            .flat_map(|p| [p.0[0], p.0[1], p.0[2]])
                            .collect();
                        let compressed_rgb =
                            miniz_oxide::deflate::compress_to_vec_zlib(&rgb_data, 6);

                        let smask_ref = if has_alpha {
                            let alpha_data: Vec<u8> = rgba.pixels().map(|p| p.0[3]).collect();
                            let compressed_alpha =
                                miniz_oxide::deflate::compress_to_vec_zlib(&alpha_data, 6);
                            let mask_ref = alloc();
                            let mut mask = pdf.image_xobject(mask_ref, &compressed_alpha);
                            mask.filter(Filter::FlateDecode);
                            mask.width(w as i32);
                            mask.height(h as i32);
                            mask.color_space().device_gray();
                            mask.bits_per_component(8);
                            Some(mask_ref)
                        } else {
                            None
                        };

                        let mut xobj = pdf.image_xobject(xobj_ref, &compressed_rgb);
                        xobj.filter(Filter::FlateDecode);
                        xobj.width(w as i32);
                        xobj.height(h as i32);
                        xobj.color_space().device_rgb();
                        xobj.bits_per_component(8);
                        if let Some(mask_ref) = smask_ref {
                            xobj.s_mask(mask_ref);
                        }
                    } else {
                        continue;
                    }
                }
            }

            image_xobjects.push((pdf_name.clone(), xobj_ref));
            image_pdf_names.insert(block_idx, pdf_name);
        }
    }

    let t_images = t0.elapsed();

    // Phase 2: build multi-page content streams
    let mut all_contents: Vec<Content> = Vec::new();
    let mut current_content = Content::new();
    let mut slot_top = doc.page_height - doc.margin_top;
    let mut prev_space_after: f32 = 0.0;
    let mut all_page_links: Vec<Vec<LinkAnnotation>> = Vec::new();
    let mut current_page_links: Vec<LinkAnnotation> = Vec::new();

    let adjacent_para = |idx: usize| -> Option<&crate::model::Paragraph> {
        match doc.blocks.get(idx)? {
            Block::Paragraph(p) => Some(p),
            Block::Table(_) => None,
        }
    };

    for (block_idx, block) in doc.blocks.iter().enumerate() {
        match block {
            Block::Paragraph(para) => {
                // Handle explicit page breaks
                if para.page_break_before {
                    let at_top = (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;
                    if !at_top {
                        all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                        all_page_links.push(std::mem::take(&mut current_page_links));
                        slot_top = doc.page_height - doc.margin_top;
                    }
                    prev_space_after = 0.0;
                    // If the paragraph only contains the break (no text), skip rendering
                    if is_text_empty(&para.runs) {
                        continue;
                    }
                }

                let next_para = adjacent_para(block_idx + 1);
                let prev_para = if block_idx > 0 {
                    adjacent_para(block_idx - 1)
                } else {
                    None
                };

                let effective_space_before =
                    if para.contextual_spacing && prev_para.is_some_and(|p| p.contextual_spacing) {
                        0.0
                    } else {
                        para.space_before
                    };
                let effective_space_after =
                    if para.contextual_spacing && next_para.is_some_and(|p| p.contextual_spacing) {
                        0.0
                    } else {
                        para.space_after
                    };

                let mut inter_gap = f32::max(prev_space_after, effective_space_before);

                let (font_size, tallest_lhr, tallest_ar) =
                    tallest_run_metrics(&para.runs, &seen_fonts);
                let effective_line_spacing = para.line_spacing.unwrap_or(doc.line_spacing);
                let line_h = tallest_lhr
                    .map(|ratio| font_size * ratio * effective_line_spacing)
                    .unwrap_or(font_size * 1.2 * effective_line_spacing);

                let para_text_x = doc.margin_left + para.indent_left;
                let para_text_width = (text_width - para.indent_left).max(1.0);
                let label_x = doc.margin_left + para.indent_left - para.indent_hanging;
                // Only apply hanging first-line shift when there's no visible label;
                // with a visible label, the hanging area is for the label only.
                let text_hanging = if para.list_label.is_empty() {
                    para.indent_hanging
                } else {
                    0.0
                };

                let text_empty = is_text_empty(&para.runs);
                let has_tabs = para.runs.iter().any(|r| r.is_tab);
                let lines = if para.image.is_some() || text_empty {
                    vec![]
                } else if has_tabs {
                    build_tabbed_line(&para.runs, &seen_fonts, &para.tab_stops, para.indent_left)
                } else {
                    build_paragraph_lines(&para.runs, &seen_fonts, para_text_width, text_hanging)
                };

                let content_h = if para.image.is_some() {
                    para.content_height.max(doc.line_pitch)
                } else if text_empty {
                    line_h
                } else {
                    let min_lines = 1 + para.extra_line_breaks as usize;
                    lines.len().max(min_lines) as f32 * line_h
                };

                let needed = inter_gap + content_h;
                let at_page_top = (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;

                let keep_next_extra = if para.keep_next {
                    let mut extra = 0.0;
                    let mut prev_sa = effective_space_after;
                    let mut i = block_idx + 1;
                    while let Some(next) = adjacent_para(i) {
                        let (nfs, nlhr, _) = tallest_run_metrics(&next.runs, &seen_fonts);
                        let next_inter = f32::max(prev_sa, next.space_before);
                        let next_first_line_h = nlhr.map(|ratio| nfs * ratio).unwrap_or(nfs * 1.2);
                        if !next.keep_next {
                            let next_ls = next.line_spacing.unwrap_or(doc.line_spacing);
                            let next_line_h = nlhr
                                .map(|ratio| nfs * ratio * next_ls)
                                .unwrap_or(nfs * 1.2 * next_ls);
                            extra += next_inter + next_first_line_h + next_line_h;
                            break;
                        }
                        extra += next_inter + next_first_line_h;
                        prev_sa = next.space_after;
                        i += 1;
                    }
                    extra
                } else {
                    0.0
                };

                if !at_page_top && slot_top - needed - keep_next_extra < doc.margin_bottom {
                    let available = slot_top - inter_gap - doc.margin_bottom;
                    let first_line_h = tallest_lhr
                        .map(|ratio| font_size * ratio)
                        .unwrap_or(font_size);
                    let mut lines_that_fit = if line_h > 0.0 && available >= first_line_h {
                        1 + ((available - first_line_h) / line_h).floor() as usize
                    } else {
                        0
                    };

                    // Reduce to ensure at least 2 lines remain on next page (orphan control)
                    if lines_that_fit > 0 && lines.len().saturating_sub(lines_that_fit) < 2 {
                        lines_that_fit = lines.len().saturating_sub(2);
                    }

                    if lines_that_fit >= 2 && lines_that_fit < lines.len() {
                        let first_part = &lines[..lines_that_fit];
                        slot_top -= inter_gap;
                        let ascender_ratio = tallest_ar.unwrap_or(0.75);
                        let baseline_y = slot_top - font_size * ascender_ratio;

                        if !para.list_label.is_empty() {
                            let (label_font_name, label_bytes) =
                                label_for_run(&para.runs[0], &seen_fonts, &para.list_label);
                            current_content
                                .begin_text()
                                .set_font(Name(label_font_name.as_bytes()), font_size)
                                .next_line(label_x, baseline_y)
                                .show(Str(&label_bytes))
                                .end_text();
                        }

                        render_paragraph_lines(
                            &mut current_content,
                            first_part,
                            &para.alignment,
                            para_text_x,
                            para_text_width,
                            baseline_y,
                            line_h,
                            lines.len(),
                            0,
                            &mut current_page_links,
                            text_hanging,
                        );

                        all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                        all_page_links.push(std::mem::take(&mut current_page_links));
                        slot_top = doc.page_height - doc.margin_top;

                        let rest = &lines[lines_that_fit..];
                        let rest_content_h = rest.len() as f32 * line_h;
                        let baseline_y2 = slot_top - font_size * ascender_ratio;

                        render_paragraph_lines(
                            &mut current_content,
                            rest,
                            &para.alignment,
                            para_text_x,
                            para_text_width,
                            baseline_y2,
                            line_h,
                            lines.len(),
                            lines_that_fit,
                            &mut current_page_links,
                            text_hanging,
                        );

                        slot_top -= rest_content_h;
                        prev_space_after = effective_space_after;
                        continue;
                    }

                    all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                    all_page_links.push(std::mem::take(&mut current_page_links));
                    slot_top = doc.page_height - doc.margin_top;
                    inter_gap = 0.0;
                }

                // Suppress space_before at the top of a page (after a page break, not first page)
                let at_new_page_top = !all_contents.is_empty()
                    && (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;
                if at_new_page_top {
                    inter_gap = 0.0;
                }

                slot_top -= inter_gap;

                if (para.image.is_some() || text_empty) && para.content_height > 0.0 {
                    if let Some(pdf_name) = image_pdf_names.get(&block_idx) {
                        let img = para.image.as_ref().unwrap();
                        let y_bottom = slot_top - img.display_height;
                        let x = doc.margin_left
                            + match para.alignment {
                                Alignment::Center => {
                                    (text_width - img.display_width).max(0.0) / 2.0
                                }
                                Alignment::Right => (text_width - img.display_width).max(0.0),
                                _ => 0.0,
                            };
                        current_content.save_state();
                        current_content.transform([
                            img.display_width,
                            0.0,
                            0.0,
                            img.display_height,
                            x,
                            y_bottom,
                        ]);
                        current_content.x_object(Name(pdf_name.as_bytes()));
                        current_content.restore_state();
                    } else {
                        current_content
                            .set_fill_gray(0.5)
                            .rect(doc.margin_left, slot_top - content_h, text_width, content_h)
                            .fill_nonzero()
                            .set_fill_gray(0.0);
                    }
                } else if !lines.is_empty() {
                    let ascender_ratio = tallest_ar.unwrap_or(0.75);
                    let baseline_y = slot_top - font_size * ascender_ratio;

                    if !para.list_label.is_empty() {
                        let (label_font_name, label_bytes) =
                            label_for_run(&para.runs[0], &seen_fonts, &para.list_label);
                        current_content
                            .begin_text()
                            .set_font(Name(label_font_name.as_bytes()), font_size)
                            .next_line(label_x, baseline_y)
                            .show(Str(&label_bytes))
                            .end_text();
                    }

                    render_paragraph_lines(
                        &mut current_content,
                        &lines,
                        &para.alignment,
                        para_text_x,
                        para_text_width,
                        baseline_y,
                        line_h,
                        lines.len(),
                        0,
                        &mut current_page_links,
                        text_hanging,
                    );
                }

                // Draw bottom border if present
                if let Some(bdr) = &para.border_bottom {
                    let line_y = slot_top - content_h - bdr.space_pt;
                    let [r, g, b] = bdr.color;
                    current_content
                        .set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0)
                        .rect(
                            doc.margin_left,
                            line_y - bdr.width_pt,
                            text_width,
                            bdr.width_pt,
                        )
                        .fill_nonzero()
                        .set_fill_rgb(0.0, 0.0, 0.0);
                }

                slot_top -= content_h;
                prev_space_after = effective_space_after;
            }

            Block::Table(table) => {
                render_table(
                    table,
                    doc,
                    &seen_fonts,
                    &mut current_content,
                    &mut all_contents,
                    &mut all_page_links,
                    &mut current_page_links,
                    &mut slot_top,
                    prev_space_after,
                );
                prev_space_after = 0.0;
            }
        }
    }
    all_contents.push(current_content);
    all_page_links.push(current_page_links);

    let t_layout = t0.elapsed();

    // Phase 2b: render headers and footers on each page
    let total_pages = all_contents.len();
    let has_hf = doc.header_default.is_some()
        || doc.header_first.is_some()
        || doc.footer_default.is_some()
        || doc.footer_first.is_some();

    if has_hf {
        for (page_idx, content) in all_contents.iter_mut().enumerate() {
            let is_first = page_idx == 0;
            let page_num = page_idx + 1;

            // Header
            let header = if is_first && doc.different_first_page {
                doc.header_first.as_ref()
            } else {
                doc.header_default.as_ref()
            };
            if let Some(hf) = header {
                render_header_footer(content, hf, &seen_fonts, doc, true, page_num, total_pages);
            }

            // Footer
            let footer = if is_first && doc.different_first_page {
                doc.footer_first.as_ref()
            } else {
                doc.footer_default.as_ref()
            };
            if let Some(hf) = footer {
                render_header_footer(content, hf, &seen_fonts, doc, false, page_num, total_pages);
            }
        }
    }

    let t_headers = t0.elapsed();

    // Phase 3: allocate page and content IDs now that page count is known
    let n = all_contents.len();
    let page_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();

    // Allocate annotation refs and write annotation objects
    let page_annot_refs: Vec<Vec<Ref>> = all_page_links
        .iter()
        .map(|links| {
            links
                .iter()
                .map(|link| {
                    let annot_ref = alloc();
                    let mut annot = pdf.annotation(annot_ref);
                    annot
                        .subtype(pdf_writer::types::AnnotationType::Link)
                        .rect(link.rect)
                        .border(0.0, 0.0, 0.0, None);
                    annot
                        .action()
                        .action_type(pdf_writer::types::ActionType::Uri)
                        .uri(Str(link.url.as_bytes()));
                    annot_ref
                })
                .collect()
        })
        .collect();

    for (i, c) in all_contents.into_iter().enumerate() {
        pdf.stream(content_ids[i], &c.finish());
    }

    pdf.catalog(catalog_id).pages(pages_id);
    pdf.pages(pages_id)
        .kids(page_ids.iter().copied())
        .count(n as i32);

    let font_pairs: Vec<(String, Ref)> = font_order
        .iter()
        .map(|name| (seen_fonts[name].pdf_name.clone(), seen_fonts[name].font_ref))
        .collect();

    for i in 0..n {
        let mut page = pdf.page(page_ids[i]);
        page.media_box(Rect::new(0.0, 0.0, doc.page_width, doc.page_height))
            .parent(pages_id)
            .contents(content_ids[i]);
        if !page_annot_refs[i].is_empty() {
            page.annotations(page_annot_refs[i].iter().copied());
        }
        {
            let mut resources = page.resources();
            {
                let mut fonts = resources.fonts();
                for (name, font_ref) in &font_pairs {
                    fonts.pair(Name(name.as_bytes()), *font_ref);
                }
            }
            if !image_xobjects.is_empty() {
                let mut xobjects = resources.x_objects();
                for (name, xobj_ref) in &image_xobjects {
                    xobjects.pair(Name(name.as_bytes()), *xobj_ref);
                }
            }
        }
    }

    let t_assembly = t0.elapsed();

    log::info!(
        "Render phases: collect_runs={:.1}ms, font_embed={:.1}ms, images={:.1}ms, layout={:.1}ms, headers={:.1}ms, assembly={:.1}ms",
        t_collect.as_secs_f64() * 1000.0,
        (t_fonts - t_collect).as_secs_f64() * 1000.0,
        (t_images - t_fonts).as_secs_f64() * 1000.0,
        (t_layout - t_images).as_secs_f64() * 1000.0,
        (t_headers - t_layout).as_secs_f64() * 1000.0,
        (t_assembly - t_headers).as_secs_f64() * 1000.0,
    );

    Ok(pdf.finish())
}

fn label_for_run<'a>(
    run: &Run,
    seen_fonts: &'a HashMap<String, FontEntry>,
    label: &str,
) -> (&'a str, Vec<u8>) {
    let key = font_key(run);
    let entry = seen_fonts.get(&key).expect("font registered");
    (entry.pdf_name.as_str(), to_winansi_bytes(label))
}
