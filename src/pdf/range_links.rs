//! Link annotations over the text between two bookmarks
//! (`ConvertOptions::range_links`).
//!
//! The parser marks the runs of each range (`model::RangeSpan`) and layout
//! the chunks cut from them (`layout::ChunkRange`). Each drawn line leaves
//! one placeholder link per range it holds: the box of that range's glyphs
//! and pictures on the line. Placeholders ride with the page's links through
//! placement and trial layouts, like `bookmarks`, and are turned into the
//! range's annotations before the PDF is assembled.
use pdf_writer::Rect;

use super::LinkAnnotation;
use super::layout::WordChunk;
use crate::RangeLink;
use crate::fonts::FontEntry;
use std::collections::HashMap;

const MARKER: &str = "\0docxide-range:";
/// Per range and page: the first line, the lines in between, the last line.
const MAX_RECTS_PER_PAGE: usize = 3;
/// A heading is at most this many leading lines.
const MAX_HEADING_ROWS: usize = 4;
/// Lines whose tops differ by less than this (points) share a row: table
/// cells side by side.
const ROW_TOLERANCE: f32 = 2.0;
/// Ascender and descender per em of a face without metrics (unresolved
/// fonts drawn in a standard face).
const DEFAULT_GLYPH_BOX: (f32, f32) = (0.8, -0.2);

/// One line's share of a range.
#[derive(Clone, Copy, Debug)]
pub(super) struct LineBox {
    pub(super) range: u32,
    pub(super) rect: Rect,
    /// Every run with text on the line (the whole line, not only the
    /// range's part) is bold.
    pub(super) bold: bool,
}

fn union(a: Rect, b: Rect) -> Rect {
    Rect::new(
        a.x1.min(b.x1),
        a.y1.min(b.y1),
        a.x2.max(b.x2),
        a.y2.max(b.y2),
    )
}

/// Visible ink of a chunk for a range: text that is not only spaces, or a
/// picture. Leaders, tab gaps and decorations are not part of the text.
fn is_visible_text(chunk: &WordChunk) -> bool {
    chunk.inline_image_name.is_none()
        && chunk.leader.is_none()
        && chunk.checkbox.is_none()
        && chunk.text.chars().any(|c| !c.is_whitespace())
}

/// The box a chunk draws: its glyphs' advance and the face's ascender and
/// descender around its baseline, or the picture drawn.
fn chunk_box(
    chunk: &WordChunk,
    x: f32,
    baseline: f32,
    picture_baseline: f32,
    fonts: &HashMap<&str, &FontEntry>,
) -> Rect {
    if chunk.inline_image_name.is_some() {
        let box_bottom = picture_baseline + chunk.y_offset;
        if chunk.inline_image_rotation_deg.abs() > 0.01 {
            return Rect::new(
                x,
                box_bottom,
                x + chunk.width,
                box_bottom + chunk.inline_image_height,
            );
        }
        // As `render_paragraph_lines` draws it: natural size, centred.
        let (w, h) = chunk.inline_image_size;
        let left =
            x + (chunk.width - w - chunk.inline_image_trim) / 2.0 + chunk.inline_image_effect_left;
        let bottom = box_bottom + (chunk.inline_image_height - h) / 2.0;
        return Rect::new(left, bottom, left + w, bottom + h);
    }
    let (ascender, descender) = fonts
        .get(chunk.pdf_font.as_str())
        .and_then(|entry| entry.glyph_box)
        .unwrap_or(DEFAULT_GLYPH_BOX);
    let base = baseline + chunk.y_offset;
    Rect::new(
        x,
        base + descender * chunk.font_size,
        x + chunk.width,
        base + ascender * chunk.font_size,
    )
}

/// Leave one placeholder per range on a drawn line. `x_of` places a chunk
/// on the page; `baseline` is the text's and `picture_baseline` the one the
/// line's pictures stand on.
pub(super) fn mark_line(
    links: &mut Vec<LinkAnnotation>,
    chunks: &[WordChunk],
    x_of: impl Fn(usize, &WordChunk) -> f32,
    baseline: f32,
    picture_baseline: f32,
    fonts: &HashMap<&str, &FontEntry>,
) {
    let mut boxes: Vec<(u32, Rect)> = Vec::new();
    for (index, chunk) in chunks.iter().enumerate() {
        if chunk.ranges.is_empty() || !(chunk.inline_image_name.is_some() || is_visible_text(chunk))
        {
            continue;
        }
        let x = x_of(index, chunk);
        let full = chunk_box(chunk, x, baseline, picture_baseline, fonts);
        let width = full.x2 - full.x1;
        for part in &chunk.ranges {
            let rect = if chunk.inline_image_name.is_some() {
                full
            } else {
                Rect::new(
                    full.x1 + part.from * width,
                    full.y1,
                    full.x1 + part.to * width,
                    full.y2,
                )
            };
            match boxes.iter_mut().find(|(range, _)| *range == part.range) {
                Some((_, line)) => *line = union(*line, rect),
                None => boxes.push((part.range, rect)),
            }
        }
    }
    if boxes.is_empty() {
        return;
    }
    let mut text = chunks.iter().filter(|c| is_visible_text(c)).peekable();
    let bold = text.peek().is_some() && text.all(|c| c.bold);
    for (range, rect) in boxes {
        links.push(LinkAnnotation {
            rect,
            url: format!("{MARKER}{range}:{}", u8::from(bold)),
            node: None,
            text: String::new(),
        });
    }
}

/// Leave a placeholder per range for a picture drawn as its paragraph's only
/// content (`Paragraph::image`): `(x, y, width, height)`, PDF space.
pub(super) fn mark_picture(
    links: &mut Vec<LinkAnnotation>,
    ranges: &[u32],
    (x, y, width, height): (f32, f32, f32, f32),
) {
    for range in ranges {
        links.push(LinkAnnotation {
            rect: Rect::new(x, y, x + width, y + height),
            url: format!("{MARKER}{range}:0"),
            node: None,
            text: String::new(),
        });
    }
}

/// Take the placeholders out of every page's links: per page, the line boxes
/// in the order they were drawn.
pub(super) fn collect(pages: &mut [Vec<LinkAnnotation>]) -> Vec<Vec<LineBox>> {
    pages
        .iter_mut()
        .map(|links| {
            let mut lines = Vec::new();
            links.retain(|link| {
                let Some(marker) = link.url.strip_prefix(MARKER) else {
                    return true;
                };
                let (range, bold) = marker.split_once(':').unwrap_or((marker, "0"));
                if let Ok(range) = range.parse() {
                    lines.push(LineBox {
                        range,
                        rect: link.rect,
                        bold: bold == "1",
                    });
                }
                false
            });
            lines
        })
        .collect()
}

/// Top to bottom, then left to right.
fn by_position(a: &Rect, b: &Rect) -> std::cmp::Ordering {
    b.y2.total_cmp(&a.y2).then(a.x1.total_cmp(&b.x1))
}

/// A page's rectangles for one range: the first and last lines as they are,
/// the lines in between merged.
fn collapse(lines: &[LineBox]) -> Vec<Rect> {
    let mut rects: Vec<Rect> = lines.iter().map(|line| line.rect).collect();
    rects.sort_by(by_position);
    if rects.len() <= MAX_RECTS_PER_PAGE {
        return rects;
    }
    let last = rects[rects.len() - 1];
    let middle = rects[2..rects.len() - 1]
        .iter()
        .fold(rects[1], |acc, rect| union(acc, *rect));
    vec![rects[0], middle, last]
}

/// The range's heading on a page: its leading rows while they are bold and
/// follow each other without a gap taller than a row.
fn heading(lines: &[LineBox]) -> Option<Rect> {
    let mut sorted: Vec<LineBox> = lines.to_vec();
    sorted.sort_by(|a, b| by_position(&a.rect, &b.rect));
    let mut rows: Vec<(Rect, bool)> = Vec::new();
    for line in sorted {
        match rows.last_mut() {
            Some((row, bold)) if (line.rect.y2 - row.y2).abs() < ROW_TOLERANCE => {
                *row = union(*row, line.rect);
                *bold &= line.bold;
            }
            _ => rows.push((line.rect, line.bold)),
        }
    }
    let mut title: Option<Rect> = None;
    let mut previous: Option<Rect> = None;
    for (row, bold) in rows.into_iter().take(MAX_HEADING_ROWS) {
        if previous.is_some_and(|p| p.y1 - row.y2 > row.y2 - row.y1) || !bold {
            break;
        }
        title = Some(title.map_or(row, |t| union(t, row)));
        previous = Some(row);
    }
    title
}

/// The range links' annotations, added to each page's links. `lines` are the
/// placeholders by page; `offsets` how far each page's body is moved down
/// (section vertical alignment). Returns how many annotations were added to
/// each page.
pub(super) fn annotate(
    pages: &mut [Vec<LinkAnnotation>],
    lines: &[Vec<LineBox>],
    links: &[RangeLink],
    offsets: &[f32],
) -> Vec<usize> {
    let mut added = vec![0; pages.len()];
    for (index, link) in links.iter().enumerate() {
        let mut first_page = true;
        for (page, page_lines) in lines.iter().enumerate() {
            let offset = offsets.get(page).copied().unwrap_or(0.0);
            let mine: Vec<LineBox> = page_lines
                .iter()
                .filter(|line| line.range as usize == index)
                .map(|line| LineBox {
                    rect: Rect::new(
                        line.rect.x1,
                        line.rect.y1 - offset,
                        line.rect.x2,
                        line.rect.y2 - offset,
                    ),
                    ..*line
                })
                .collect();
            if mine.is_empty() || page >= pages.len() {
                continue;
            }
            let mut push = |rect: Rect, url: &str| {
                pages[page].push(LinkAnnotation {
                    rect,
                    url: url.to_string(),
                    node: None,
                    text: String::new(),
                });
                added[page] += 1;
            };
            for rect in collapse(&mine) {
                push(rect, &link.uri);
            }
            if first_page
                && let Some(uri) = &link.heading_uri
                && let Some(rect) = heading(&mine)
            {
                push(rect, uri);
            }
            first_page = false;
        }
    }
    added
}

#[cfg(test)]
mod tests {
    use super::*;

    fn line(top: f32, bottom: f32, x1: f32, x2: f32, bold: bool) -> LineBox {
        LineBox {
            range: 0,
            rect: Rect::new(x1, bottom, x2, top),
            bold,
        }
    }

    #[test]
    fn keeps_first_and_last_lines_and_merges_the_rest() {
        let lines: Vec<LineBox> = (0..5)
            .map(|i| {
                line(
                    700.0 - 14.0 * i as f32,
                    688.0 - 14.0 * i as f32,
                    72.0 + i as f32,
                    300.0,
                    false,
                )
            })
            .rev()
            .collect();
        let rects = collapse(&lines);
        assert_eq!(rects.len(), 3);
        assert_eq!((rects[0].y2, rects[0].x1), (700.0, 72.0));
        assert_eq!(
            (rects[1].y2, rects[1].y1, rects[1].x1),
            (686.0, 646.0, 73.0)
        );
        assert_eq!((rects[2].y2, rects[2].x1), (644.0, 76.0));
        assert_eq!(collapse(&lines[..2]).len(), 2);
    }

    #[test]
    fn a_heading_is_the_leading_bold_rows_without_a_gap() {
        let bold = [
            line(700.0, 688.0, 72.0, 200.0, true),
            line(700.5, 688.0, 300.0, 400.0, true),
            line(686.0, 674.0, 72.0, 150.0, true),
            line(672.0, 660.0, 72.0, 500.0, false),
        ];
        let title = heading(&bold).unwrap();
        assert_eq!(
            (title.x1, title.y1, title.x2, title.y2),
            (72.0, 674.0, 400.0, 700.5)
        );
        // A gap taller than a row ends the heading.
        let apart = [
            line(700.0, 688.0, 72.0, 200.0, true),
            line(650.0, 638.0, 72.0, 200.0, true),
        ];
        assert_eq!(heading(&apart).unwrap().y1, 688.0);
        assert!(heading(&[line(700.0, 688.0, 72.0, 200.0, false)]).is_none());
        let five: Vec<LineBox> = (0..5)
            .map(|i| {
                line(
                    700.0 - 12.0 * i as f32,
                    688.0 - 12.0 * i as f32,
                    72.0,
                    90.0,
                    true,
                )
            })
            .collect();
        assert_eq!(heading(&five).unwrap().y1, 688.0 - 36.0);
    }

    #[test]
    fn placeholders_become_annotations_on_their_pages() {
        let mut pages = vec![Vec::new(), Vec::new()];
        let first = Rect::new(72.0, 688.0, 200.0, 700.0);
        let mut chunk_links = Vec::new();
        for (page, rect) in [(0, first), (1, Rect::new(72.0, 600.0, 90.0, 612.0))] {
            chunk_links.clear();
            chunk_links.push(LinkAnnotation {
                rect,
                url: format!("{MARKER}0:1"),
                node: None,
                text: String::new(),
            });
            pages[page].append(&mut chunk_links);
        }
        pages[0].push(LinkAnnotation {
            rect: first,
            url: "https://example.com".into(),
            node: None,
            text: String::new(),
        });
        let lines = collect(&mut pages);
        assert_eq!(pages[0].len(), 1);
        assert_eq!(lines[1].len(), 1);
        let links = [RangeLink::new("a", "b", "x:range").heading_uri("x:heading")];
        let added = annotate(&mut pages, &lines, &links, &[0.0, 10.0]);
        assert_eq!(added, [2, 1]);
        assert_eq!(pages[0][1].url, "x:range");
        assert_eq!(pages[0][2].url, "x:heading");
        // The heading is only looked for on the range's first page; the
        // second page's body is moved 10pt down.
        assert_eq!(
            (pages[1][0].url.as_str(), pages[1][0].rect.y2),
            ("x:range", 602.0)
        );
    }
}
