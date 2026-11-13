use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use super::color::{fill_rgb, stroke_rgb};
use crate::fonts::FontEntry;
use crate::model::{Comment, SectionProperties};

/// Width of the balloon area Word adds beside the text column.
const BALLOON_AREA: f32 = 279.7;
/// Unzoomed gaps from the text column to the gray pane, and from the pane to
/// the balloon area's end.
const PANE_LEFT_GAP: f32 = 9.1;
const PANE_RIGHT_GAP: f32 = 13.3;
/// The balloon area width Word's zoom rounding works with (`page_zoom`).
const ZOOM_BALLOON_AREA: f32 = 279.15;

/// Word's zoom `(scale, tx, ty)` for a page of a document with comments: the
/// text column plus a 279.7pt balloon area fills the page width, so 12pt text
/// prints at 9.16pt with a 90pt right margin (case63, case64) and at 8.96pt with
/// 72pt (door_air_cooling_unit_spec). The zoomed page sits 0.96pt from the left
/// edge and near the vertical centre (Word's ty 94.32 and 100.80).
///
/// Word picks the zoom in 1/300 steps: rounding with a 279.15pt area
/// reproduces all 148 zooms measured on print exports (Letter and A4, 54-90pt
/// right margins), where the unrounded 279.7 missed one in six.
pub(super) fn page_zoom(page_width: f32, page_height: f32, margin_right: f32) -> (f32, f32, f32) {
    let scale =
        (page_width / (page_width - margin_right + ZOOM_BALLOON_AREA) * 300.0).round() / 300.0;
    // The zoomed page's top: the centred position floored to the 1/300in
    // grid, then two units up on Letter paper and two down on A4 (121 print
    // references; the print path's paper margins, presumably).
    let unit = 0.24;
    let centred = (page_height * (1.0 - scale) / 2.0 / unit).floor() * unit;
    let paper_shift = if (page_height - 792.0).abs() < 1.0 {
        -2.0 * unit
    } else if (page_height - 841.9).abs() < 1.0 {
        2.0 * unit
    } else {
        0.0
    };
    let top = centred + paper_shift;
    (scale, 0.96, page_height * (1.0 - scale) - top)
}

const PANE_BG: [u8; 3] = [240, 240, 240];
const CALLOUT_FILL: [u8; 3] = [251, 220, 217];
const CALLOUT_STROKE: [u8; 3] = [200, 90, 90];
const CONNECTOR_RGB: [u8; 3] = [209, 52, 56];
/// Unzoomed stroke widths in Word's print export: the connector is one
/// 1/300in unit (dashed 3 on, 3 off), the balloon border two.
const CONNECTOR_WIDTH: f32 = 0.24;
const BALLOON_BORDER: f32 = 0.48;
const TEXT_RGB: [u8; 3] = [0, 0, 0];

/// Unzoomed: a balloon's gaps to the pane's left and right edges, and its
/// text's inset (case63's balloons span 423.8-599.5, text from 427.0).
const BALLOON_LEFT_GAP: f32 = 22.8;
const BALLOON_RIGHT_GAP: f32 = 4.3;
const BALLOON_TEXT_PAD: f32 = 4.3;
const CALLOUT_RADIUS: f32 = 6.0;
/// Unzoomed balloon box, measured on Word's print export (Tahoma 8pt and
/// Times New Roman 9pt labels over Calibri text): the first baseline sits
/// this far plus the text's ascent below the top, the box is its lines plus
/// this much, and balloons stack this far apart.
const BALLOON_TOP_TO_ASCENT: f32 = 2.35;
const BALLOON_EXTRA_H: f32 = 4.7;
const BALLOON_GAP: f32 = 0.97;

/// Anchor tuple: (comment_id, end_x, highlight_top_y, font_size). The highlight
/// top is what callouts align to; the connector originates from
/// `highlight_top_y - 1.15 * font_size` (i.e. the highlight bottom).
pub(super) fn render_comment_pane(
    content: &mut Content,
    comments: &HashMap<u32, Comment>,
    anchors: &[(u32, f32, f32, f32)],
    sp: &SectionProperties,
    // The page's `page_zoom`, which `anchors` are already zoomed by.
    (zoom, tx, ty): (f32, f32, f32),
    seen_fonts: &HashMap<String, FontEntry>,
    // The balloon text: body and label font keys (`fonts::comment_pane_keys`)
    // and its size before the zoom (`Document::comment_font_size`).
    (body_key, label_key, font_size): (&str, &str, f32),
) {
    let (Some(body_entry), Some(label_entry)) =
        (seen_fonts.get(body_key), seen_fonts.get(label_key))
    else {
        return;
    };

    // The pane is the zoomed page's balloon area, as tall as the zoomed page:
    // 406.4-602.8 at case63's 90pt right margin, 411.0-603.1 at door_air's 72pt.
    let pane_x = tx + (sp.page_width - sp.margin_right + PANE_LEFT_GAP) * zoom;
    let pane_width = (BALLOON_AREA - PANE_LEFT_GAP - PANE_RIGHT_GAP) * zoom;
    let pane_bottom = ty;
    let pane_h = sp.page_height * zoom;
    let pane_top = pane_bottom + pane_h;

    content.save_state();
    fill_rgb(content, PANE_BG);
    content.rect(pane_x, pane_bottom, pane_width, pane_h);
    content.fill_nonzero();
    content.restore_state();

    if anchors.is_empty() {
        return;
    }

    let inner_x = pane_x + BALLOON_LEFT_GAP * zoom;
    let inner_w = pane_width - (BALLOON_LEFT_GAP + BALLOON_RIGHT_GAP) * zoom;
    let text_x = inner_x + BALLOON_TEXT_PAD * zoom;
    let text_w = inner_w - 2.0 * BALLOON_TEXT_PAD * zoom;
    // Glyphs land on Word's 1/300in grid; lines step by the comment font's
    // own line height at the unrounded size.
    let fs = (font_size / 0.24).round() * 0.24 * zoom;
    let line_h = body_entry.line_h_ratio.unwrap_or(1.2) * font_size * zoom;
    let first_baseline =
        (BALLOON_TOP_TO_ASCENT + body_entry.ascender_ratio.unwrap_or(0.8) * font_size) * zoom;

    // Balloons stack in reading order of where they hang (the reference
    // marks), not in the order their ranges open.
    let mut anchors = anchors.to_vec();
    anchors.sort_by(|a, b| {
        (b.2.round(), a.1)
            .partial_cmp(&(a.2.round(), b.1))
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    let anchors = &anchors[..];
    let mut placements: Vec<(u32, f32, f32, Vec<(bool, String)>)> = Vec::new();
    // Stack callouts flush top-to-bottom: only the first uses its anchor_y to
    // pick a starting position; subsequent ones sit directly under the previous
    // box so there's no whitespace between them (Word's behavior).
    let mut next_top: f32 = pane_top;
    let mut first = true;
    for &(cid, _anchor_x, anchor_y, _fs) in anchors {
        let Some(comment) = comments.get(&cid) else {
            continue;
        };
        let label = format_label(comment);
        let mut lines: Vec<(bool, String)> = Vec::new();
        let label_w = label_entry.word_width(&label, fs, false);
        let first_line_space = (text_w - label_w).max(0.0);
        let mut first_body = String::new();
        let mut remainder = comment.text.as_str();
        if first_line_space > 0.0 {
            let (head, tail) = take_words_for_width(remainder, first_line_space, body_entry, fs);
            first_body = head;
            remainder = tail;
        }
        lines.push((true, format!("{label}{first_body}")));
        wrap_into_lines(remainder, text_w, body_entry, fs, &mut lines);

        let h = lines.len() as f32 * line_h + BALLOON_EXTRA_H * zoom;
        let mut top = if first {
            anchor_y.min(next_top)
        } else {
            next_top
        };
        first = false;
        if top - h < pane_bottom {
            top = h + pane_bottom;
        }
        placements.push((cid, top, h, lines));
        next_top = top - h - BALLOON_GAP * zoom;
    }

    for (_cid, top, h, lines) in &placements {
        draw_rounded_rect(
            content,
            inner_x,
            top - h,
            inner_w,
            *h,
            CALLOUT_RADIUS * zoom,
        );
        content.save_state();
        fill_rgb(content, CALLOUT_FILL);
        stroke_rgb(content, CALLOUT_STROKE);
        content.set_line_width(BALLOON_BORDER * zoom);
        content.fill_even_odd_and_stroke();
        content.restore_state();

        let show = |content: &mut Content, entry: &FontEntry, x: f32, y: f32, text: &str| {
            content.begin_text();
            fill_rgb(content, TEXT_RGB);
            content.set_font(Name(entry.pdf_name.as_bytes()), fs);
            content.set_text_matrix([1.0, 0.0, 0.0, 1.0, x, y]);
            content.show(Str(&entry.encode(text)));
            content.end_text();
        };
        let mut text_y = top - first_baseline;
        for (is_first, line_text) in lines {
            if *is_first {
                let label = format_label_from_text(line_text);
                let label_only = &line_text[..label.len()];
                let body_part = &line_text[label.len()..];
                show(content, label_entry, text_x, text_y, label_only);
                if !body_part.is_empty() {
                    let label_w = label_entry.word_width(label_only, fs, false);
                    show(content, body_entry, text_x + label_w, text_y, body_part);
                }
            } else {
                show(content, body_entry, text_x, text_y, line_text);
            }
            text_y -= line_h;
        }
    }

    content.save_state();
    stroke_rgb(content, CONNECTOR_RGB);
    content.set_line_width(CONNECTOR_WIDTH * zoom);
    content.set_dash_pattern([3.0 * CONNECTOR_WIDTH * zoom; 2], 0.0);
    for (i, &(_cid, anchor_x, anchor_y, fs)) in anchors.iter().enumerate() {
        if let Some((_, top, _, _)) = placements.get(i) {
            // Two-segment connector: horizontal from the BOTTOM of the
            // highlighted phrase to the gray-pane edge, then a short diagonal
            // leg into the callout's left side. anchor_y is the highlight top;
            // subtract the full highlight height (1.15 * fs) to get the
            // bottom edge where the connector originates.
            let connector_origin_y = anchor_y - 1.15 * fs;
            let callout_attach_y = *top - 0.4 * line_h;
            content.move_to(anchor_x, connector_origin_y);
            content.line_to(pane_x, connector_origin_y);
            content.line_to(inner_x, callout_attach_y);
            content.stroke();
        }
    }
    content.restore_state();
}

pub(super) fn format_label(c: &Comment) -> String {
    let initials = if c.initials.is_empty() {
        "?".to_string()
    } else {
        c.initials.clone()
    };
    format!("Commented [{initials}{}]: ", c.display_index)
}

fn format_label_from_text(s: &str) -> &str {
    if let Some(idx) = s.find("]: ") {
        &s[..idx + 3]
    } else {
        s
    }
}

fn take_words_for_width<'a>(
    text: &'a str,
    max_w: f32,
    entry: &FontEntry,
    fs: f32,
) -> (String, &'a str) {
    let mut head = String::new();
    let mut last_split = 0usize;
    let mut cursor = 0usize;
    for (i, ch) in text.char_indices() {
        let candidate = &text[..i + ch.len_utf8()];
        let w = entry.word_width(candidate, fs, false);
        if w > max_w {
            break;
        }
        cursor = i + ch.len_utf8();
        if ch == ' ' {
            last_split = cursor;
        }
    }
    if cursor == text.len() {
        return (text.to_string(), "");
    }
    if last_split == 0 {
        return (head, text);
    }
    head.push_str(&text[..last_split]);
    (head, &text[last_split..])
}

fn wrap_into_lines(
    text: &str,
    max_w: f32,
    entry: &FontEntry,
    fs: f32,
    out: &mut Vec<(bool, String)>,
) {
    let mut remainder = text.trim_start();
    while !remainder.is_empty() {
        let (head, tail) = take_words_for_width(remainder, max_w, entry, fs);
        if head.is_empty() {
            let mut take_n = 1;
            for (n, _) in remainder.char_indices().take(60) {
                if entry.word_width(&remainder[..n + 1], fs, false) > max_w {
                    break;
                }
                take_n = n + 1;
            }
            out.push((false, remainder[..take_n].to_string()));
            remainder = &remainder[take_n..];
        } else {
            out.push((false, head.trim_end().to_string()));
            remainder = tail.trim_start();
        }
    }
}

fn draw_rounded_rect(content: &mut Content, x: f32, y: f32, w: f32, h: f32, r: f32) {
    let r = r.min(w / 2.0).min(h / 2.0);
    let k = r * 0.5523;
    content.move_to(x + r, y);
    content.line_to(x + w - r, y);
    content.cubic_to(x + w - r + k, y, x + w, y + r - k, x + w, y + r);
    content.line_to(x + w, y + h - r);
    content.cubic_to(x + w, y + h - r + k, x + w - r + k, y + h, x + w - r, y + h);
    content.line_to(x + r, y + h);
    content.cubic_to(x + r - k, y + h, x, y + h - r + k, x, y + h - r);
    content.line_to(x, y + r);
    content.cubic_to(x, y + r - k, x + r - k, y, x + r, y);
    content.close_path();
}
