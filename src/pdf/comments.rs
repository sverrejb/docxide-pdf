use std::collections::HashMap;

use pdf_writer::{Content, Name, Str};

use crate::fonts::{FontEntry, encode_as_gids, to_winansi_bytes};
use crate::model::Comment;

/// Width of the right-side pane drawn inside the page when the document has comments.
const PANE_WIDTH: f32 = 197.0;
/// Distance from the right page edge to the right edge of the pane.
const PANE_RIGHT_MARGIN: f32 = 10.0;
/// Whitespace above and below the gray pane band.
const PANE_VPAD: f32 = 94.0;

/// Body content scale factor Word applies when comments are present (9.16pt /
/// 12pt). Combined with the BODY_TX/BODY_TY translations below in a single
/// `q s 0 0 s tx ty cm` operator that wraps the body content stream.
pub(super) const BODY_SCALE: f32 = 0.7633;
/// Horizontal translation paired with BODY_SCALE; keeps the scaled left margin
/// aligned with the docx margin_left.
pub(super) const BODY_TX: f32 = 14.7;
/// Vertical translation paired with BODY_SCALE; shifts the scaled first baseline
/// down to roughly the pane top (matches the y origin Word uses in its cm).
pub(super) const BODY_TY: f32 = 94.4;

const PANE_BG: [u8; 3] = [240, 240, 240];
const CALLOUT_FILL: [u8; 3] = [251, 220, 217];
const CALLOUT_STROKE: [u8; 3] = [200, 90, 90];
const CONNECTOR_RGB: [u8; 3] = [209, 52, 56];
const LABEL_RGB: [u8; 3] = [80, 30, 30];
const BODY_RGB: [u8; 3] = [50, 30, 30];

const CALLOUT_PAD_X: f32 = 6.0;
const CALLOUT_PAD_Y: f32 = 4.0;
const CALLOUT_GAP: f32 = 0.0;
const CALLOUT_RADIUS: f32 = 4.0;
const CALLOUT_FONT_SIZE: f32 = 8.5;
const CALLOUT_LINE_H: f32 = 11.0;
const PANE_LEFT_PAD: f32 = 8.0;

pub(super) fn render_comment_pane(
    content: &mut Content,
    comments: &HashMap<u32, Comment>,
    anchors: &[(u32, f32, f32)],
    page_width: f32,
    page_height: f32,
    seen_fonts: &HashMap<String, FontEntry>,
) {
    let Some((body_font_key, body_entry)) = pick_font(seen_fonts, &["Aptos", "Calibri"]) else {
        return;
    };
    let label_entry = pick_font(seen_fonts, &["Aptos/B", "Aptos Bold", "Calibri/B"])
        .map(|(k, e)| (k, e))
        .unwrap_or((body_font_key.clone(), body_entry));

    let pane_x = page_width - PANE_WIDTH - PANE_RIGHT_MARGIN;
    let pane_top = page_height - PANE_VPAD;
    let pane_bottom = PANE_VPAD;
    let pane_h = pane_top - pane_bottom;

    content.save_state();
    fill_rgb(content, PANE_BG);
    content.rect(pane_x, pane_bottom, PANE_WIDTH, pane_h);
    content.fill_nonzero();
    content.restore_state();

    if anchors.is_empty() {
        return;
    }

    let inner_x = pane_x + PANE_LEFT_PAD;
    let inner_w = PANE_WIDTH - 2.0 * PANE_LEFT_PAD;
    let text_w = inner_w - 2.0 * CALLOUT_PAD_X;

    let mut placements: Vec<(u32, f32, f32, Vec<(bool, String)>)> = Vec::new();
    // Stack callouts flush top-to-bottom: only the first uses its anchor_y to
    // pick a starting position; subsequent ones sit directly under the previous
    // box so there's no whitespace between them (Word's behavior).
    let mut next_top: f32 = pane_top - CALLOUT_PAD_Y;
    let mut first = true;
    for &(cid, _anchor_x, anchor_y) in anchors {
        let Some(comment) = comments.get(&cid) else {
            continue;
        };
        let label = format_label(comment);
        let mut lines: Vec<(bool, String)> = Vec::new();
        let label_w = label_entry.1.word_width(&label, CALLOUT_FONT_SIZE, false);
        let first_line_space = (text_w - label_w).max(0.0);
        let mut first_body = String::new();
        let mut remainder = comment.text.as_str();
        if first_line_space > 0.0 {
            let (head, tail) = take_words_for_width(remainder, first_line_space, body_entry, CALLOUT_FONT_SIZE);
            first_body = head;
            remainder = tail;
        }
        lines.push((true, format!("{label}{first_body}")));
        wrap_into_lines(remainder, text_w, body_entry, CALLOUT_FONT_SIZE, &mut lines);

        let h = lines.len() as f32 * CALLOUT_LINE_H + 2.0 * CALLOUT_PAD_Y;
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
        next_top = top - h - CALLOUT_GAP;
    }

    for (_cid, top, h, lines) in &placements {
        draw_rounded_rect(content, inner_x, top - h, inner_w, *h, CALLOUT_RADIUS);
        content.save_state();
        fill_rgb(content, CALLOUT_FILL);
        stroke_rgb(content, CALLOUT_STROKE);
        content.set_line_width(0.6);
        content.fill_even_odd_and_stroke();
        content.restore_state();

        let pdf_name_map: HashMap<&str, &FontEntry> = seen_fonts
            .values()
            .map(|e| (e.pdf_name.as_str(), e))
            .collect();
        let mut text_y = top - CALLOUT_PAD_Y - CALLOUT_FONT_SIZE;
        for (is_first, line_text) in lines {
            if *is_first {
                let label = format_label_from_text(line_text);
                let label_only = &line_text[..label.len()];
                let body_part = &line_text[label.len()..];
                let label_bytes = encode_for(label_entry.1, label_only);
                content.begin_text();
                fill_rgb_in_text(content, LABEL_RGB);
                content.set_font(Name(label_entry.1.pdf_name.as_bytes()), CALLOUT_FONT_SIZE);
                content.set_text_matrix([1.0, 0.0, 0.0, 1.0, inner_x + CALLOUT_PAD_X, text_y]);
                content.show(Str(&label_bytes));
                content.end_text();
                if !body_part.is_empty() {
                    let label_w = label_entry.1.word_width(label_only, CALLOUT_FONT_SIZE, false);
                    let body_bytes = encode_for(body_entry, body_part);
                    content.begin_text();
                    fill_rgb_in_text(content, BODY_RGB);
                    content.set_font(Name(body_entry.pdf_name.as_bytes()), CALLOUT_FONT_SIZE);
                    content.set_text_matrix([
                        1.0, 0.0, 0.0, 1.0,
                        inner_x + CALLOUT_PAD_X + label_w,
                        text_y,
                    ]);
                    content.show(Str(&body_bytes));
                    content.end_text();
                }
            } else {
                let body_bytes = encode_for(body_entry, line_text);
                content.begin_text();
                fill_rgb_in_text(content, BODY_RGB);
                content.set_font(Name(body_entry.pdf_name.as_bytes()), CALLOUT_FONT_SIZE);
                content.set_text_matrix([1.0, 0.0, 0.0, 1.0, inner_x + CALLOUT_PAD_X, text_y]);
                content.show(Str(&body_bytes));
                content.end_text();
            }
            text_y -= CALLOUT_LINE_H;
        }
        let _ = pdf_name_map;
    }

    content.save_state();
    stroke_rgb(content, CONNECTOR_RGB);
    content.set_line_width(0.55);
    content.set_dash_pattern([0.55, 0.55], 0.0);
    for (i, &(_cid, anchor_x, anchor_y)) in anchors.iter().enumerate() {
        if let Some((_, top, _, _)) = placements.get(i) {
            // Two-segment connector: horizontal from highlighted phrase to the
            // gray-pane edge, then a short diagonal/vertical leg into the
            // callout's left side. Matches Word's PDF export style.
            let callout_attach_y = *top - CALLOUT_LINE_H * 0.4;
            content.move_to(anchor_x, anchor_y);
            content.line_to(pane_x, anchor_y);
            content.line_to(inner_x, callout_attach_y);
            content.stroke();
        }
    }
    content.restore_state();
}

fn format_label(c: &Comment) -> String {
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

fn pick_font<'a>(
    seen: &'a HashMap<String, FontEntry>,
    candidates: &[&str],
) -> Option<(String, &'a FontEntry)> {
    for c in candidates {
        if let Some(e) = seen.get(*c) {
            return Some((c.to_string(), e));
        }
    }
    seen.iter().next().map(|(k, v)| (k.clone(), v))
}

fn encode_for(entry: &FontEntry, text: &str) -> Vec<u8> {
    match entry.char_to_gid.as_ref() {
        Some(map) => encode_as_gids(text, map),
        None => to_winansi_bytes(text),
    }
}

fn fill_rgb(content: &mut Content, rgb: [u8; 3]) {
    content.set_fill_rgb(
        rgb[0] as f32 / 255.0,
        rgb[1] as f32 / 255.0,
        rgb[2] as f32 / 255.0,
    );
}

fn fill_rgb_in_text(content: &mut Content, rgb: [u8; 3]) {
    content.set_fill_rgb(
        rgb[0] as f32 / 255.0,
        rgb[1] as f32 / 255.0,
        rgb[2] as f32 / 255.0,
    );
}

fn stroke_rgb(content: &mut Content, rgb: [u8; 3]) {
    content.set_stroke_rgb(
        rgb[0] as f32 / 255.0,
        rgb[1] as f32 / 255.0,
        rgb[2] as f32 / 255.0,
    );
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
