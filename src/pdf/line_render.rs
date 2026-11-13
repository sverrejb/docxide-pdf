//! Drawing laid-out lines: glyphs, decorations, line numbers and change bars.

use std::collections::HashMap;
use std::sync::Arc;

use pdf_writer::types::TextRenderingMode;
use pdf_writer::{Buf, Content, Name, Primitive, Rect, Str};

use crate::fonts::{FontEntry, encode_as_gids, to_winansi_bytes};
use crate::model::{Alignment, FormCheckbox, ParagraphBorder, TextFill};

use super::color::{fill_color_or_black, stroke_color_or_black};
use super::layout::{
    LineGrid, LinkAnnotation, LinkTagger, MixedJustify, TextLine, WordChunk, baseline_drop,
    char_justify_gaps, glyph_fallbacks, inline_line_advance, run_line_metrics, spaces_before_each,
};

/// A decoration rectangle (underline, strikethrough, shading): x, y, width, height, colour.
type Decoration = (f32, f32, f32, f32, Option<[u8; 3]>);

/// The lines (top, thickness) of an underline at `font_size` under the
/// baseline `y`. Word draws a single underline from the font's post table, its
/// offset below the baseline and its thickness each on a 0.25pt grid (online
/// export probe: Calibri, Times New Roman, Arial, Aptos at 10/12/22pt). A
/// double underline is two lines half as thick, see `double_underline`.
pub(super) fn underline_lines(
    entry: Option<&FontEntry>,
    font_size: f32,
    double: bool,
    link: bool,
    y: f32,
) -> Vec<(f32, f32)> {
    if let Some(e) = entry
        && let Some((offset, th)) = post_underline(e, font_size)
    {
        if !double {
            return vec![(y - offset - th, th)];
        }
        let descent = e
            .underline
            .map_or(f32::MAX, |(_, _, d)| quarter_pt(d * font_size));
        return double_underline(offset, th, descent)
            .into_iter()
            .map(|(top, th)| (y - top - th, th))
            .collect();
    }
    let thick = (font_size * 0.05).max(0.5);
    let ul_y = if link {
        y - font_size * 0.08
    } else {
        y - font_size * 0.12
    };
    let top = ul_y - thick;
    let mut lines = vec![(top, thick)];
    if double {
        lines.push((top - thick - (thick * 1.5).max(1.0), thick));
    }
    lines
}

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

/// Word sets a legacy check box in a square cell as wide and tall as one line
/// of the run's font at the box's size, down to that line's descent, and
/// strokes the box 1pt inside its top and left edges, 1.5pt inside the others
/// (Word probes: Arial, Calibri, Courier New). Returns (side, descent).
pub(super) fn checkbox_cell(entry: &FontEntry, cb: FormCheckbox) -> (f32, f32) {
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

/// A font's OS/2 strikeout as Word draws it: its top edge above the baseline
/// and its thickness, each on a 0.25pt grid (online export probe, the
/// underline's 12 cases). The footnote separator's rule is one too.
pub(super) fn os2_strike(entry: &FontEntry, font_size: f32) -> Option<(f32, f32)> {
    let (pos, th) = entry.strikeout?;
    Some(on_quarter_point_grid(pos * font_size, th * font_size))
}

/// The strike through text on the baseline `y`: its bottom edge and
/// thickness (see `os2_strike`).
pub(super) fn strike_line(entry: &FontEntry, font_size: f32, y: f32) -> Option<(f32, f32)> {
    os2_strike(entry, font_size).map(|(top, th)| (y + top - th, th))
}

/// A font's post-table underline as Word draws it: its offset below the
/// baseline and its thickness, on the same 0.25pt grid as `os2_strike`.
pub(super) fn post_underline(entry: &FontEntry, font_size: f32) -> Option<(f32, f32)> {
    let (pos, th, descent) = entry.underline?;
    let (offset, thickness) = on_quarter_point_grid(pos * font_size, th * font_size);
    let descent = quarter_pt(descent * font_size);
    Some((offset.min(descent - thickness), thickness))
}

/// Word's double underline from the single one (offset below the baseline and
/// thickness, both on the 0.25pt grid): two lines half as thick, the upper
/// starting at half the single's offset, the lower ending where the single
/// does, at least 1pt between them, the pair pushed up (upper line no higher
/// than the baseline) to stay inside the font's descent. Returns (offset,
/// thickness) per line. Online export probe, Arial / Calibri / Times New Roman
/// / Courier New at every point size 6–48: all 172 pairs exact.
fn double_underline(offset: f32, thickness: f32, descent: f32) -> [(f32, f32); 2] {
    let th = quarter_pt(thickness / 2.0).max(0.25);
    let mut top = quarter_pt(offset / 2.0);
    let mut bottom = (offset + thickness).max(top + 2.0 * th + 1.0);
    if bottom > descent {
        bottom = descent;
        top = (bottom - 2.0 * th - 1.0).max(0.0);
    }
    [(top, th), (bottom - th, th)]
}

fn on_quarter_point_grid(offset: f32, thickness: f32) -> (f32, f32) {
    (quarter_pt(offset), quarter_pt(thickness).max(0.25))
}

pub(super) fn quarter_pt(v: f32) -> f32 {
    (v * 4.0).round() / 4.0
}

/// How a chunk's glyphs are drawn, chosen once for the face and its shadow
/// and relief copies.
enum GlyphDraw<'a> {
    Placed(&'a FontEntry, Vec<(&'a FontEntry, &'a str, f32, f32)>),
    Extras(Vec<f32>),
    Rtl(&'a FontEntry),
    Fallbacks,
    Kerned(&'a FontEntry),
    Plain,
}

/// Draws `text` with the font's pair kerning as TJ adjustments, the same pairs
/// the word's width was measured with: Word kerns inside words too (Aptos
/// Display "Te" 1.6pt tighter at 20pt, case3), so plain Tj left every glyph
/// after a kerned pair out of place.
fn show_kerned(content: &mut Content, entry: &FontEntry, text: &str, boundary_space: bool) {
    let mut items = Vec::new();
    let mut run = String::new();
    for (ch, adjust) in kern_moves(entry, text) {
        if adjust != 0.0 {
            items.push(TjItem::Text(entry.encode(&run)));
            run.clear();
            items.push(TjItem::Adjust(adjust));
        }
        run.push_str(ch);
    }
    if boundary_space {
        run.push(' ');
    }
    items.push(TjItem::Text(entry.encode(&run)));
    write_tj(content, &items);
}

/// Each character of `text` with the TJ adjustment its pair kerning puts
/// before it. TJ subtracts: a negative kern (tighter) is a positive
/// adjustment.
fn kern_moves<'a>(
    entry: &'a FontEntry,
    text: &'a str,
) -> impl Iterator<Item = (&'a str, f32)> + 'a {
    let mut prev: Option<char> = None;
    text.char_indices().map(move |(i, ch)| {
        let kern = prev.map_or(0.0, |p| entry.kern_1000(p, ch));
        prev = Some(ch);
        (&text[i..i + ch.len_utf8()], -kern)
    })
}

/// The TJ adjustment that takes `pen` to a glyph at `x`, and the pen after
/// the glyph: whole thousandths keep the numbers short, and the pen follows
/// the rounded move, so the error never adds up. TJ subtracts: moving right
/// is a negative adjustment.
fn placed_move(pen: &mut f32, x: f32, advance: f32) -> f32 {
    let adjust = (*pen - x).round();
    *pen += advance - adjust;
    adjust
}

enum TjItem {
    Text(Vec<u8>),
    Adjust(f32),
}

/// A TJ array without pdf-writer's spaces between strings and numbers,
/// `[(ab)9(cd)]` as Word writes it: about 3% of a text page.
struct TjArray<'a>(&'a [TjItem]);

impl Primitive for TjArray<'_> {
    fn write(self, buf: &mut Buf) {
        buf.extend([b'[']);
        for item in self.0 {
            match item {
                TjItem::Text(text) => Str(text).write(buf),
                TjItem::Adjust(adjust) => adjust.write(buf),
            }
        }
        buf.extend([b']']);
    }
}

fn write_tj(content: &mut Content, items: &[TjItem]) {
    content.op("TJ").operand(TjArray(items));
}

/// Draws `glyphs` (text, x, advance) in the order given, each at its own x
/// from the pen's start, with TJ moves between them (`placed_move`). Moves
/// and advances scale alike with Tz, so all of this is in thousandths of the
/// font size. `trailing` is shown last, unplaced.
/// Starts with the pen at `pen` and returns where it ends.
fn show_placed(
    content: &mut Content,
    entry: &FontEntry,
    glyphs: &[(&str, f32, f32)],
    trailing: Option<&str>,
    mut pen: f32,
) -> f32 {
    let mut items = Vec::new();
    for &(text, x, advance) in glyphs {
        let adjust = placed_move(&mut pen, x, advance);
        if adjust != 0.0 {
            items.push(TjItem::Adjust(adjust));
        }
        items.push(TjItem::Text(entry.encode(text)));
    }
    if let Some(t) = trailing {
        items.push(TjItem::Text(entry.encode(t)));
    }
    write_tj(content, &items);
    pen
}

/// The chunks of a line that share one text state, drawn as a single TJ as
/// Word draws a line: a `Td` and a TJ per word cost about a quarter of a
/// dense page's compressed size. Each chunk starts where a `Td` would put
/// it; inside it the glyphs move as they would in a TJ of their own.
#[derive(Default)]
struct TjRun {
    items: Vec<TjItem>,
    /// Where the run starts and its baseline.
    origin: (f32, f32),
    /// The pen after `items`, in thousandths of the font size from the
    /// start, as a viewer moves it: by the widths the font dictionary holds
    /// (`FontEntry::pen_width_1000`).
    pen: f32,
}

impl TjRun {
    /// Whether a chunk on baseline `y` goes on in this run.
    fn continues(&self, y: f32) -> bool {
        !self.items.is_empty() && self.origin.1 == y
    }

    fn start(&mut self, x: f32, y: f32) {
        self.origin = (x, y);
        self.pen = 0.0;
    }

    /// Moves the pen to `x` thousandths from the run's start, within 0.001pt:
    /// the pen follows the rounded move, so the error never adds up. Whole
    /// thousandths (0.006pt at 12pt) are enough to tip a glyph over a
    /// rasteriser's subpixel boundary.
    fn advance_to(&mut self, x: f32, font_size: f32) {
        let step = if font_size > 20.0 { 100.0 } else { 10.0 };
        self.adjust(((self.pen - x) * step).round() / step);
    }

    /// TJ subtracts: moving right is a negative adjustment.
    fn adjust(&mut self, adjust: f32) {
        if adjust == 0.0 {
            return;
        }
        match self.items.last_mut() {
            Some(TjItem::Adjust(open)) => *open += adjust,
            _ => self.items.push(TjItem::Adjust(adjust)),
        }
        self.pen -= adjust;
    }

    /// Shows `text` where the pen is; `advance` is how far it moves the pen.
    fn show(&mut self, text: Vec<u8>, advance: f32) {
        match self.items.last_mut() {
            Some(TjItem::Text(open)) => open.extend(text),
            _ => self.items.push(TjItem::Text(text)),
        }
        self.pen += advance;
    }

    /// Writes the run. Called before anything else is written to the
    /// content: what follows may change how text is drawn.
    fn flush(&mut self, content: &mut Content) {
        if self.items.is_empty() {
            return;
        }
        write_tj(content, &self.items);
        self.items.clear();
    }
}

/// `show_placed` with each glyph in its own face (`base`, the face set, or
/// a fallback one): one TJ per face, the pen carried across them.
fn show_placed_faces(
    content: &mut Content,
    base: &FontEntry,
    glyphs: &[(&FontEntry, &str, f32, f32)],
    font_size: f32,
    trailing: Option<&str>,
) {
    let mut pen = 0.0;
    for run in glyphs.chunk_by(|a, b| std::ptr::eq(a.0, b.0)) {
        let face = run[0].0;
        let other = !std::ptr::eq(face, base);
        if other {
            content.set_font(Name(face.pdf_name.as_bytes()), font_size);
        }
        let seg: Vec<(&str, f32, f32)> = run.iter().map(|&(_, t, x, a)| (t, x, a)).collect();
        pen = show_placed(content, face, &seg, None, pen);
        if other {
            content.set_font(Name(base.pdf_name.as_bytes()), font_size);
        }
    }
    if let Some(t) = trailing {
        content.show(Str(&base.encode(t)));
    }
}

/// Each chunk's glyph x positions as Word's online export draws them
/// (`glyph_grid`), for the chunks the rule is known for. Chunks set edge to
/// edge or a plain space apart share one grid, the line's exact positions;
/// a justified line's spaces are plain ones plus `extra_per_gap` shared out.
/// `line_start` is the line's exact start and where Word lays it out: a
/// sequence starting there moves by the whole pixels between the two.
fn device_grid(
    line: &TextLine,
    entries: &HashMap<&str, &FontEntry>,
    chunk_abs_x: &dyn Fn(usize, &WordChunk) -> f32,
    extra_per_gap: f32,
    line_start: (f32, f32),
) -> Vec<Option<Vec<f32>>> {
    use super::glyph_grid::{GridGlyph, drawn_positions};
    let flush = |line: &TextLine,
                 x0: f32,
                 seq: &mut Vec<GridGlyph>,
                 members: &mut Vec<(usize, usize)>,
                 out: &mut [Option<Vec<f32>>]| {
        let px = quarter_pt;
        let shift = if (x0 - line_start.0).abs() < 1e-3 {
            px(line_start.1) - px(x0)
        } else {
            0.0
        };
        let drawn = drawn_positions(x0, seq);
        for &(ci, first) in members.iter() {
            let n = line.chunks[ci].text.chars().count();
            out[ci] = Some(drawn[first..first + n].iter().map(|x| x + shift).collect());
        }
        seq.clear();
        members.clear();
    };
    let mut out = vec![None; line.chunks.len()];
    let mut seq: Vec<GridGlyph> = Vec::new();
    // (chunk, its first glyph in `seq`)
    let mut members: Vec<(usize, usize)> = Vec::new();
    let mut x0 = 0.0;
    // The previous chunk when it is in `seq`, and where it ends.
    let mut prev: Option<(usize, f32)> = None;
    for (i, chunk) in line.chunks.iter().enumerate() {
        let Some(entry) = entries
            .get(chunk.pdf_font.as_str())
            .copied()
            .filter(|e| grid_eligible(chunk, e))
        else {
            flush(line, x0, &mut seq, &mut members, &mut out);
            prev = None;
            continue;
        };
        let x = chunk_abs_x(i, chunk);
        let size = chunk.font_size;
        // The spaces before this chunk (count, width, and the part of it Word
        // rounds apart: letter spacing and justification stretch), when they
        // are plain ones.
        let joined = prev.and_then(|(p, end)| {
            let gap = x - end;
            let n = chunk.preceding_spaces;
            if n == 0 {
                return (gap.abs() < 0.01).then_some((0, 0.0, 0.0));
            }
            let pc = &line.chunks[p];
            let sw = gap / n as f32;
            let plain = [entry.space_width(size)].into_iter().chain(
                entries
                    .get(pc.pdf_font.as_str())
                    .map(|e| e.space_width(pc.font_size)),
            );
            let spacing = [0.0, chunk.char_spacing, pc.char_spacing];
            let stretch = [0.0, extra_per_gap / n as f32];
            plain
                .flat_map(|w| spacing.map(|cs| (w, cs)))
                .flat_map(|(w, cs)| stretch.map(|st| (w, cs + st)))
                .find(|&(w, apart)| (w + apart - sw).abs() < 0.01)
                .map(|(_, apart)| (n, sw, apart))
        });
        let new_run = match (joined, prev) {
            (Some((n, sw, stretch)), Some((p, _))) => {
                let pc = &line.chunks[p];
                let shaped = line.shaped_kerning
                    && pc.kern
                    && entries
                        .get(pc.pdf_font.as_str())
                        .is_some_and(|e| e.kern_pairs.is_some());
                seq.extend((0..n).map(|_| GridGlyph {
                    advance: sw,
                    apart: stretch,
                    shaped,
                    font_size: pc.font_size,
                    run_start: false,
                }));
                pc.pdf_font != chunk.pdf_font
                    || pc.font_size != size
                    || pc.color != chunk.color
                    || pc.y_offset != chunk.y_offset
            }
            _ => {
                flush(line, x0, &mut seq, &mut members, &mut out);
                x0 = x;
                true
            }
        };
        members.push((i, seq.len()));
        let shaped = line.shaped_kerning && chunk.kern && entry.kern_pairs.is_some();
        seq.extend(
            glyph_advances(chunk, entry)
                .enumerate()
                .map(|(j, (advance, kern))| GridGlyph {
                    advance,
                    apart: chunk.char_spacing + if shaped { 0.0 } else { kern },
                    shaped,
                    font_size: size,
                    run_start: (j == 0 && new_run) || chunk.source_starts.contains(&j),
                }),
        );
        prev = Some((i, x + chunk.width));
    }
    flush(line, x0, &mut seq, &mut members, &mut out);
    out
}

/// Each leader chunk's glyph x positions (`glyph_grid::leader_positions`).
/// Both ends of the tab are where Word's device pen stands: the line's
/// device start plus the rounded exact distance from the line start (example:
/// 552 leaders, none off by a slot; a stop 5.4pt from a 71.5pt margin ends
/// at 308 px and holds one 22 px underscore).
fn leader_grid(
    line: &TextLine,
    entries: &HashMap<&str, &FontEntry>,
    chunk_abs_x: &dyn Fn(usize, &WordChunk) -> f32,
    line_start: (f32, f32),
) -> Vec<Option<Vec<f32>>> {
    if line.chunks.iter().all(|c| c.leader.is_none()) {
        return Vec::new();
    }
    line.chunks
        .iter()
        .enumerate()
        .map(|(i, chunk)| {
            let ch = chunk.leader?;
            let entry = entries.get(chunk.pdf_font.as_str())?;
            let x = chunk_abs_x(i, chunk);
            let px =
                |v: f32| ((line_start.1 * 4.0).round() + ((v - line_start.0) * 4.0).round()) as i64;
            Some(super::glyph_grid::leader_positions(
                entry.char_width_1000(ch) * chunk.font_size / 1000.0,
                px(x),
                px(x + chunk.width),
            ))
        })
        .collect()
}

/// Plain text whose drawn width is its glyphs' exact advances.
fn grid_eligible(chunk: &WordChunk, entry: &FontEntry) -> bool {
    let exact: f32 = glyph_advances(chunk, entry).map(|(a, _)| a).sum();
    !chunk.text.is_empty()
        && chunk.leader.is_none()
        && chunk.math.is_none()
        && chunk.text_shadow.is_none()
        && chunk.checkbox.is_none()
        && chunk.inline_image_name.is_none()
        && !chunk.rtl
        && chunk.text_scale == 100.0
        && chunk
            .text
            .chars()
            .all(|c| c < '\u{2E80}' && !entry.missing_chars.contains(&c))
        && (exact - chunk.width).abs() < 0.01
}

/// Each glyph's exact advance in points, letter spacing included, and the
/// kerning in it: a kerned pair's adjustment on its first glyph. A scaled
/// chunk's glyphs advance as layout measured them (`FontEntry::scaled_advances`).
fn glyph_advances<'a>(
    chunk: &'a WordChunk,
    entry: &'a FontEntry,
) -> impl Iterator<Item = (f32, f32)> + 'a {
    entry
        .scaled_advances(&chunk.text, chunk.font_size, chunk.kern, chunk.text_scale)
        .map(|(advance, kerning)| (advance + chunk.char_spacing, kerning))
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

/// A right-to-left chunk's glyphs left to right as they sit on the page, the
/// pen starting at the chunk's left edge, as Word writes them: MuPDF reads a
/// line in content order and applies the bidi algorithm to it as visual text,
/// so logical order with leftward moves came out reversed ("صنلا" for "النص").
/// The boundary space goes first, left of the chunk, where Word puts it. A
/// letter's combining marks follow it with no move: they hang back over it
/// from its end, as in left-to-right text.
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
    let mut placed: Vec<(&str, f32, f32)> = clusters
        .into_iter()
        .map(|(s, w)| {
            right -= w;
            (s, right, w)
        })
        .collect();
    placed.reverse();
    if boundary_space {
        let w = entry.char_width_1000(' ') + tc;
        placed.insert(0, (" ", -w, w));
    }
    show_placed(content, entry, &placed, None, 0.0);
}

/// `text`'s glyphs left to right (see `show_rtl`), each at the place its
/// right-to-left layout (bidi, paragraph level 1) gives it, the pen starting
/// at the left end: a list label that reads "3." draws ".3".
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
    let mut placed: Vec<(&str, f32, f32)> =
        glyphs.iter().zip(x).map(|(&(s, w), x)| (s, x, w)).collect();
    placed.sort_by(|a, b| a.1.total_cmp(&b.1));
    show_placed(content, entry, &placed, None, 0.0);
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

/// §17.6.8 line numbering for a body paragraph's rendered lines. The counter
/// lives on the PageBuilder, which resets it per page or section as
/// `lnNumType restart` asks.
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

impl LineNumberArg<'_> {
    /// Count one line and draw its number, if shown, at baseline `y`,
    /// `x_shift` right of the usual place, in the numbers' font or else
    /// `line_font`.
    pub(super) fn number(
        &mut self,
        content: &mut Content,
        x_shift: f32,
        y: f32,
        line_font: impl FnOnce() -> Option<(String, f32)>,
        pdf_name_to_entry: &HashMap<&str, &FontEntry>,
        link_tags: Option<&mut LinkTagger<'_>>,
    ) {
        let idx = *self.counter;
        *self.counter = idx + 1;
        let value = self.start + idx as i32;
        let show = value >= 1 && (self.count_by <= 1 || value % self.count_by as i32 == 0);
        let Some((font, fs)) = show.then(|| self.font.clone().or_else(line_font)).flatten() else {
            return;
        };
        let s = value.to_string();
        let w = pdf_name_to_entry
            .get(font.as_str())
            .map(|e| e.word_width(&s, fs, false))
            .unwrap_or(fs * 0.5 * s.chars().count() as f32);
        let bytes = encode_text_for_pdf(&s, &font, pdf_name_to_entry);
        let x = self.right_x + x_shift - w;
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
        match link_tags {
            Some(lt) => lt.artifact(content, draw),
            None => draw(content),
        }
    }
}

/// Fake a bold face by stroking the glyph outlines along with the fill.
fn begin_synthetic_bold(content: &mut Content, chunk: &WordChunk) {
    content.set_line_width(chunk.font_size * 0.02);
    stroke_color_or_black(content, chunk.color);
    content.set_text_rendering_mode(TextRenderingMode::FillStroke);
}

// ponytail: change bars are collected here and drawn by the page pass in
// `render_pass` (margin side, mirror-aware); an out-parameter through the ten
// `render_paragraph_lines` callers can wait until headers, footers and text
// boxes want bars too.
thread_local! {
    static CHANGE_BARS: std::cell::RefCell<Vec<(f32, f32)>> =
        const { std::cell::RefCell::new(Vec::new()) };
    /// (above, below): what the paragraph's first revised line adds above
    /// its box and its last below (Word for Mac probe p16: the paragraph's
    /// own share of the gap above it, max(0, before - previous after), and
    /// its space after).
    static BAR_EXTENSION: std::cell::Cell<(f32, f32)> = const { std::cell::Cell::new((0.0, 0.0)) };
}

/// Set before a paragraph's lines render, cleared after.
pub(super) fn set_change_bar_extension(above: f32, below: f32) {
    BAR_EXTENSION.with(|e| e.set((above, below)));
}

/// A revised line box that drew no chunk (an empty paragraph with a revised
/// mark: Word's bar runs unbroken over a decree's empty inserted paragraphs).
pub(super) fn record_change_bar(bottom: f32, top: f32) {
    let (above, below) = BAR_EXTENSION.with(|e| e.get());
    CHANGE_BARS.with(|b| b.borrow_mut().push((bottom - below, top + above)));
}

pub(super) fn change_bars_len() -> usize {
    CHANGE_BARS.with(|b| b.borrow().len())
}

/// Word bars a revised table row over the row's whole box, borders included
/// (probe p06: one bar 638.88-680.16 over three revised rows where the cells'
/// line boxes left 0.48pt gaps at the borders): the bars recorded since
/// `from` become one.
pub(super) fn collapse_change_bars_from(from: usize, bottom: f32, top: f32) {
    CHANGE_BARS.with(|b| {
        let mut b = b.borrow_mut();
        if b.len() > from {
            b.truncate(from);
            b.push((bottom, top));
        }
    });
}

/// (bottom, top) of every revised line drawn since the last call.
pub(super) fn take_change_bars() -> Vec<(f32, f32)> {
    CHANGE_BARS.with(|b| std::mem::take(&mut *b.borrow_mut()))
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
    // The part of `margin_left` that is the paragraph's left indent.
    indent_left: f32,
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
    // (x, y, w, h) the inline pictures are clipped to.
    picture_clip: Option<(f32, f32, f32, f32)>,
    line_grid: Option<LineGrid>,
) {
    let mut current_color: Option<[u8; 3]> = None;
    let mut pattern_fill_active = false;
    let mut cur_font_name = String::new();
    let mut cur_font_size: f32 = -1.0;
    let mut cur_char_spacing: f32 = 0.0;
    let mut cur_text_scale: f32 = 100.0;
    let mut cur_synthetic_bold = false;
    let mut has_text_outline = false;

    let pdf_name_to_entry = super::fonts::by_pdf_name(seen_fonts);
    let fallback_faces = glyph_fallbacks(seen_fonts);

    // Per-line baseline offsets below `first_baseline_y`: each line's top is the
    // sum of the previous lines' advances, and a picture line drops its baseline
    // by the picture's surplus over the ascent; a line sized by its own smaller
    // face raises its baseline by `ascent_shift`. Normal lines reduce to `line_pitch`.
    let mut line_y_offsets: Vec<f32> = Vec::with_capacity(lines.len());
    let mut line_top = 0.0f32;
    for line in lines {
        line_y_offsets.push(line_top + baseline_drop(line, text_metrics.0));
        line_top += inline_line_advance(line, line_pitch, text_metrics);
    }

    let last_line_idx = total_line_count.saturating_sub(1);
    // The first-line indent goes to the first line with content, past any
    // `float_skip` lines.
    let first_content_idx = lines.iter().take_while(|l| l.float_skip).count();
    for (line_num, line) in lines.iter().enumerate() {
        let raw_y = first_baseline_y - line_y_offsets[line_num];
        let y = line_grid.map_or(raw_y, |g| {
            g.drawn_baseline(
                raw_y,
                Some(line),
                line_num == 0,
                first_line_index + line_num == last_line_idx,
            )
        });
        let global_line_idx = first_line_index + line_num;
        // Markup view: a revised line's change bar spans its line box, the
        // ascent above the baseline and the rest of the advance below (Word
        // for Mac probes 2026-11-05, single, 1.15, double and exact spacing).
        if line.chunks.iter().any(|c| c.revised) {
            let (above, below) = BAR_EXTENSION.with(|e| e.get());
            // A picture standing above the ascent raises the line box (p10:
            // the bar starts at the deleted picture's top).
            let visual_top = y + line
                .chunks
                .iter()
                .map(|c| c.inline_image_height + c.y_offset)
                .fold(text_metrics.0, f32::max);
            let top = visual_top + if global_line_idx == 0 { above } else { 0.0 };
            let advance = inline_line_advance(line, line_pitch, text_metrics)
                + if global_line_idx + 1 == total_line_count {
                    below
                } else {
                    0.0
                };
            CHANGE_BARS.with(|b| b.borrow_mut().push((visual_top - advance, top)));
        }

        let (base_margin, base_width) = line_geometry
            .and_then(|g| g.get(global_line_idx))
            .copied()
            .unwrap_or((margin_left, text_width));
        let shift = if global_line_idx == first_content_idx
            && first_line_hanging.abs() > 0.001
            && !line.rtl_paragraph
        {
            first_line_hanging
        } else {
            0.0
        };
        let (eff_margin, eff_width) =
            if global_line_idx == first_content_idx && first_line_hanging.abs() > 0.001 {
                (base_margin - shift, base_width + first_line_hanging)
            } else {
                (base_margin, base_width)
            };
        // Word lays a line out from its frame or margin, left indent and
        // first-line indent each rounded to device pixels on their own
        // (paragraph-offset-frame-*: 261 − 284 + 312 px, not 289.6 → 290).
        let px = quarter_pt;
        let laid_out_margin = px(base_margin - indent_left) + px(indent_left) + px(-shift);

        // §17.6.8 margin line number: every counted body line advances the
        // page's counter; numbers are right-aligned in the page margin at the
        // line's own baseline, in that line's font/size (matching Word).
        if let Some(ln) = line_numbering.as_mut() {
            // Word keeps the distance from a line pushed right by a float
            // (torture_test's drop cap: the numbers land on the cap).
            let float_shift = match &line.right_region {
                Some(rr) if rr.first_chunk_idx == 0 => (rr.region_x - base_margin).max(0.0),
                _ => 0.0,
            };
            let line_font = || {
                line.chunks
                    .iter()
                    .find(|c| c.inline_image_name.is_none() && !c.text.is_empty())
                    .map(|c| (c.pdf_font.clone(), c.font_size))
            };
            ln.number(
                content,
                float_shift,
                y,
                line_font,
                &pdf_name_to_entry,
                link_tags.as_mut(),
            );
        }

        let left_content_width = line.left_content_width();

        let left_chunk_count = line
            .right_region
            .as_ref()
            .map(|rr| rr.first_chunk_idx)
            .unwrap_or(line.chunks.len());
        let gaps_before = spaces_before_each(&line.chunks[..left_chunk_count]);
        let jf = line.justify_from.min(left_chunk_count);
        let tab_gaps = gaps_before.get(jf).copied().unwrap_or(0);
        let gaps_before: Vec<usize> = gaps_before
            .iter()
            .enumerate()
            .map(|(i, &g)| if i < jf { 0 } else { g - tab_gaps })
            .collect();
        let left_gaps = gaps_before.last().copied().unwrap_or(0);

        // CJK justification: distribute space between every character, not just chunks.
        // Word treats CJK inter-character gaps the same as word gaps for justify.
        // Like the word gaps, only the text after the last tab spreads; a
        // leader keeps its pitch (Word example, distributed lines with tabs).
        let spread_chars = |chunks: &[WordChunk]| -> usize {
            chunks
                .iter()
                .filter(|c| c.leader.is_none())
                .map(|c| c.text.chars().count())
                .sum()
        };
        let left_char_count = spread_chars(&line.chunks[jf..left_chunk_count]);
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
        } else if line.rtl_paragraph && matches!(*alignment, Alignment::Distribute) {
            // Word example: a distributed right-to-left paragraph spreads only
            // its spaces, on every line, Latin and Hebrew alike.
            None
        } else {
            char_justify_gaps(*alignment, can_justify, has_cjk_content, left_char_count)
        };
        let is_char_justified = char_justify_gaps.is_some();
        let is_justified = mixed.is_some() || is_char_justified || (can_justify && left_gaps > 0);

        // A line wider than the measure (a picture too wide for it) starts at
        // the left edge and overflows to the right, whatever the alignment.
        let slack = if line.rtl_paragraph {
            eff_width - left_content_width
        } else {
            (eff_width - left_content_width).max(0.0)
        };
        let line_start_x = match alignment {
            Alignment::Center => eff_margin + slack / 2.0,
            Alignment::Right => eff_margin + slack,
            // A distributed line with nothing to spread after its last tab
            // (one character, a tab at the end) is centred, tabs and all, in
            // right-to-left paragraphs too.
            Alignment::Distribute if !is_justified => eff_margin + slack / 2.0,
            // Word probe: a justified right-to-left paragraph's last line
            // ends at the right margin.
            Alignment::Justify if line.rtl_paragraph && !is_justified => eff_margin + slack,
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
        // (past the last chunk: the line's end spaces)
        let x_at = |chunk_idx: usize, x_offset: f32| -> f32 {
            if let Some(ref rr) = line.right_region
                && chunk_idx >= rr.first_chunk_idx
            {
                let local_idx = chunk_idx - rr.first_chunk_idx;
                return right_start_x + x_offset + local_idx as f32 * right_extra_per_gap;
            }
            if let Some(m) = &mixed {
                line_start_x + x_offset + m.offset_before(chunk_idx)
            } else if is_char_justified {
                // Tc adds extra space after each character; shift chunk start
                // by the cumulative Tc of all chars in preceding chunks.
                let before = &line.chunks[jf.min(chunk_idx)..chunk_idx.min(line.chunks.len())];
                line_start_x + x_offset + spread_chars(before) as f32 * justify_tc
            } else {
                let gaps = gaps_before.get(chunk_idx).map_or(left_gaps, |&g| g);
                line_start_x + x_offset + gaps as f32 * extra_per_gap
            }
        };
        let chunk_abs_x = |chunk_idx: usize, chunk: &WordChunk| x_at(chunk_idx, chunk.x_offset);

        let line_right = match &line.right_region {
            Some(rr) => rr.region_x + rr.region_width,
            None => eff_margin + eff_width,
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
                let (left, right) = (sx - pad, ex + pad);
                let (bottom, top) = (y - fs * 0.2 - pad, y + fs * 0.95 + pad);
                // Each side in its style, the box's inside to its left.
                for (p0, p1, inward) in [
                    ((left, top), (right, top), -1.0),
                    ((left, bottom), (right, bottom), 1.0),
                    ((left, top), (left, bottom), 1.0),
                    ((right, top), (right, bottom), -1.0),
                ] {
                    super::color::stroke_border(content, border, p0, p1, inward);
                }
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

        // Only lines like the probes': left to right, no East Asian text,
        // justified by widening the spaces only.
        // The rest keep exact positions until each gets its own probe.
        let plain_line = line.draw_order.is_empty()
            && !line.rtl_paragraph
            && line.right_region.is_none()
            && mixed.is_none()
            && !is_char_justified
            && extra_per_gap >= 0.0;
        let line_start = match alignment {
            Alignment::Left | Alignment::Justify | Alignment::Distribute
                if line_start_x == eff_margin =>
            {
                (line_start_x, laid_out_margin)
            }
            _ => (line_start_x, line_start_x),
        };
        let grid = if plain_line {
            device_grid(
                line,
                &pdf_name_to_entry,
                &chunk_abs_x,
                extra_per_gap,
                line_start,
            )
        } else {
            Vec::new()
        };
        let leaders = leader_grid(line, &pdf_name_to_entry, &chunk_abs_x, line_start);

        if has_text_chunks {
            content.begin_text();
            let mut td_x = 0.0_f32;
            let mut td_y = 0.0_f32;
            let mut cur_shear = 0.0_f32;
            let mut run = TjRun::default();

            // Logical order, so the content stream (and the tags) read the
            // text as written even where bidi stored the chunks visually.
            for k in 0..line.chunks.len() {
                let chunk_idx = line.draw_order.get(k).copied().unwrap_or(k);
                let chunk = &line.chunks[chunk_idx];
                let leader_xs = leaders.get(chunk_idx).and_then(|l| l.as_deref());
                if chunk.inline_image_name.is_some()
                    || chunk.checkbox.is_some()
                    || leader_xs.is_some_and(|xs| xs.is_empty())
                {
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
                    && lt.chunk(content, chunk, link_url, boundary_space, &mut |c| {
                        run.flush(c)
                    })
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
                    run.flush(content);
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
                    run.flush(content);
                    fill_color_or_black(content, effective_color);
                    current_color = effective_color;
                    pattern_fill_active = false;
                }

                // Text outline takes priority over synthetic bold
                if let Some(ref outline) = chunk.text_outline {
                    if !has_text_outline {
                        run.flush(content);
                        super::wordart::apply_text_outline(
                            content,
                            outline,
                            chunk.text_fill.as_ref(),
                        );
                        has_text_outline = true;
                    }
                } else if has_text_outline {
                    run.flush(content);
                    super::wordart::reset_text_outline(content);
                    has_text_outline = false;
                    if cur_synthetic_bold {
                        begin_synthetic_bold(content, chunk);
                    }
                }

                if !has_text_outline && chunk.synthetic_bold != cur_synthetic_bold {
                    run.flush(content);
                    if chunk.synthetic_bold {
                        begin_synthetic_bold(content, chunk);
                    } else {
                        content.set_text_rendering_mode(TextRenderingMode::Fill);
                    }
                    cur_synthetic_bold = chunk.synthetic_bold;
                }

                let spreads = chunk_idx >= jf && chunk.leader.is_none();
                let effective_cs = chunk.char_spacing + if spreads { justify_tc } else { 0.0 };
                if effective_cs != cur_char_spacing {
                    run.flush(content);
                    content.set_char_spacing(effective_cs);
                    cur_char_spacing = effective_cs;
                }
                if chunk.text_scale != cur_text_scale {
                    run.flush(content);
                    content.set_horizontal_scaling(chunk.text_scale);
                    cur_text_scale = chunk.text_scale;
                }

                if cur_font_name != chunk.pdf_font || cur_font_size != chunk.font_size {
                    run.flush(content);
                    content.set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size);
                    cur_font_name.clone_from(&chunk.pdf_font);
                    cur_font_size = chunk.font_size;
                }

                if let Some(mb) = &chunk.math {
                    run.flush(content);
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

                let char_extras = mixed
                    .as_ref()
                    .and_then(|m| m.char_extras(chunk_idx, &chunk.text));
                let grid_xs = leader_xs.or_else(|| grid.get(chunk_idx).and_then(|g| g.as_deref()));
                // Scaled glyphs at their measured advances: Word's land 0.15pt
                // from these on average in char-scale-widths and 0.20pt in a
                // justified Calibri 85% document, 0.17 and 0.22pt from Tz's nominal
                // ones; the device grid model fits the first, not the second.
                let scaled_xs: Option<Vec<f32>> = primary_entry
                    .filter(|_| grid_xs.is_none() && chunk.text_scale != 100.0 && !chunk.rtl_glyphs)
                    .map(|e| {
                        let mut gx = fx;
                        glyph_advances(chunk, e)
                            .enumerate()
                            .map(|(i, (advance, _))| {
                                let x = gx;
                                let extra = char_extras.as_ref().map_or(0.0, |ex| ex[i]);
                                gx += advance + if spreads { justify_tc } else { 0.0 } + extra;
                                x
                            })
                            .collect()
                    });
                let grid_xs = grid_xs.or(scaled_xs.as_deref());
                let start_x = grid_xs.map_or(fx, |g| g[0]);
                // Per-character font fallback: if some chars are missing
                // from the primary font, split into segments and render
                // missing chars with the fallback font that has them.
                let has_missing = primary_entry.is_some_and(|e| !e.missing_chars.is_empty());
                let fallbacks: &[&FontEntry] = if has_missing { &fallback_faces } else { &[] };
                // Which of `fallbacks` draws `c`; None when the primary font has it.
                let fallback_idx = |c: char| {
                    primary_entry.filter(|p| p.missing_chars.contains(&c))?;
                    fallbacks
                        .iter()
                        .position(|f| f.char_to_gid.as_ref().is_some_and(|m| m.contains_key(&c)))
                };
                let leader_char = leader_xs.and(chunk.leader);
                let leader_face = leader_char.and_then(fallback_idx).map(|k| fallbacks[k]);
                // The pen also moves by Tc, the letter spacing.
                let tc = effective_cs * 1000.0 / chunk.font_size;
                let placed = grid_xs.zip(primary_entry).map(|(xs, &primary)| {
                    let entry = leader_face.unwrap_or(primary);
                    let chars: Vec<(usize, char)> = match leader_char {
                        Some(c) => vec![(0, c); xs.len()],
                        None => chunk.text.char_indices().collect(),
                    };
                    let glyphs: Vec<(&FontEntry, &str, f32, f32)> = chars
                        .into_iter()
                        .zip(xs)
                        .map(|((i, ch), gx)| {
                            let x = (gx - xs[0]) * 1000.0
                                / chunk.font_size
                                / (chunk.text_scale / 100.0);
                            let face = leader_face
                                .or_else(|| fallback_idx(ch).map(|k| fallbacks[k]))
                                .unwrap_or(primary);
                            (
                                face,
                                &chunk.text[i..i + ch.len_utf8()],
                                x,
                                face.char_width_1000(ch) + tc,
                            )
                        })
                        .collect();
                    (entry, glyphs)
                });
                let draw = if let Some((entry, glyphs)) = placed {
                    GlyphDraw::Placed(entry, glyphs)
                } else if let Some(extras) = char_extras {
                    GlyphDraw::Extras(extras)
                } else if let Some(entry) = primary_entry.filter(|_| chunk.rtl_glyphs) {
                    GlyphDraw::Rtl(entry)
                } else if !fallbacks.is_empty() {
                    GlyphDraw::Fallbacks
                } else if let Some(entry) =
                    primary_entry.filter(|e| chunk.kern && e.kern_pairs.is_some())
                {
                    GlyphDraw::Kerned(entry)
                } else {
                    GlyphDraw::Plain
                };
                if let Some(face) = leader_face {
                    run.flush(content);
                    content.set_font(Name(face.pdf_name.as_bytes()), chunk.font_size);
                    cur_font_name.clone_from(&face.pdf_name);
                }

                if let Some(ref sh) = chunk.text_shadow {
                    run.flush(content);
                    let (sx, sy) = (start_x + sh.offset_x, fy + sh.offset_y);
                    let (ex, ey) = (start_x - sh.offset_x, fy - sh.offset_y);
                    // The copies are drawn as the face is, only moved.
                    let draw_copy = |content: &mut Content, ox: f32, oy: f32| {
                        content.set_text_matrix([1.0, 0.0, shear, 1.0, ox, oy]);
                        let text = &chunk.text;
                        match &draw {
                            GlyphDraw::Placed(e, glyphs) => {
                                show_placed_faces(content, e, glyphs, chunk.font_size, None);
                            }
                            GlyphDraw::Rtl(e) => {
                                show_rtl(content, e, text, chunk.font_size, effective_cs, false);
                            }
                            GlyphDraw::Kerned(e) => show_kerned(content, e, text, false),
                            GlyphDraw::Extras(_) | GlyphDraw::Fallbacks | GlyphDraw::Plain => {
                                let bytes =
                                    encode_text_for_pdf(text, &chunk.pdf_font, &pdf_name_to_entry);
                                content.show(Str(&bytes));
                            }
                        }
                    };
                    let draw_shadow = |content: &mut Content| {
                        fill_color_or_black(content, Some(sh.color));
                        draw_copy(content, sx, sy);
                        if sh.relief {
                            fill_color_or_black(content, None);
                            draw_copy(content, ex, ey);
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
                let sheared = shear != 0.0 || cur_shear != 0.0;
                let mut move_to = |content: &mut Content, mx: f32, my: f32| {
                    if sheared {
                        content.set_text_matrix([1.0, 0.0, shear, 1.0, mx, my]);
                        cur_shear = shear;
                    } else {
                        content.next_line(mx - td_x, my - td_y);
                    }
                    td_x = mx;
                    td_y = my;
                };

                // What a run can take: one face the font holds every character
                // of, so the pen is where the widths say, and nothing drawn
                // between the chunks.
                let held = |e: &FontEntry, text: &str| text.chars().all(|c| e.has_char(c));
                let run_entry = match &draw {
                    _ if sheared || chunk.text_shadow.is_some() => None,
                    GlyphDraw::Placed(e, glyphs) => Some(*e).filter(|e| {
                        glyphs.iter().all(|g| std::ptr::eq(g.0, *e) && held(e, g.1))
                    }),
                    GlyphDraw::Kerned(e) => Some(*e).filter(|e| held(e, &chunk.text)),
                    GlyphDraw::Plain => primary_entry.copied().filter(|e| held(e, &chunk.text)),
                    _ => None,
                }
                .filter(|e| e.char_to_gid.is_some());
                if let Some(entry) = run_entry {
                    // Nothing to show otherwise: its decorations are drawn below.
                    if !chunk.text.is_empty() || boundary_space {
                        if !run.continues(fy) {
                            run.flush(content);
                            move_to(content, start_x, fy);
                            run.start(start_x, fy);
                        }
                        let unit = chunk.font_size * chunk.text_scale / 100_000.0;
                        run.advance_to((start_x - run.origin.0) / unit, chunk.font_size);
                        let show = |run: &mut TjRun, adjust: f32, text: &str| {
                            let advance = text.chars().map(|c| entry.pen_width_1000(c) + tc).sum();
                            run.adjust(adjust);
                            run.show(entry.encode(text), advance);
                        };
                        // Inside the chunk, the moves `show_placed` and
                        // `show_kerned` make.
                        match &draw {
                            GlyphDraw::Placed(_, glyphs) => {
                                let mut pen = 0.0;
                                for &(_, text, x, advance) in glyphs {
                                    show(&mut run, placed_move(&mut pen, x, advance), text);
                                }
                            }
                            GlyphDraw::Kerned(_) => {
                                for (ch, adjust) in kern_moves(entry, &chunk.text) {
                                    show(&mut run, adjust, ch);
                                }
                            }
                            _ => show(&mut run, 0.0, &chunk.text),
                        }
                        if boundary_space {
                            show(&mut run, 0.0, " ");
                        }
                    }
                } else {
                    run.flush(content);
                    move_to(content, start_x, fy);

                    // A segment of the chunk in its own font, or in the fallback
                    // font when its characters are missing from the primary one.
                    let show_seg = |content: &mut Content, seg: &str, fb: Option<usize>| {
                        let fb = fb.map(|k| fallbacks[k]);
                        match fb.and_then(|f| Some((f, f.char_to_gid.as_ref()?))) {
                            Some((fb, map)) => {
                                content.set_font(Name(fb.pdf_name.as_bytes()), chunk.font_size);
                                content.show(Str(&encode_as_gids(seg, map)));
                                content.set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size);
                            }
                            _ => {
                                content.show(Str(&encode_text_for_pdf(
                                    seg,
                                    &chunk.pdf_font,
                                    &pdf_name_to_entry,
                                )));
                            }
                        }
                    };
                    match draw {
                        GlyphDraw::Placed(entry, glyphs) => {
                            let trailing = boundary_space.then_some(" ");
                            show_placed_faces(content, entry, &glyphs, chunk.font_size, trailing);
                        }
                        GlyphDraw::Extras(extras) => {
                            // Only the East Asian gaps widen: one Tc per segment, the
                            // fallback font where the primary lacks the glyphs.
                            let keyed: Vec<(char, f32, Option<usize>)> = chunk
                                .text
                                .chars()
                                .zip(extras)
                                .map(|(c, x)| (c, x, fallback_idx(c)))
                                .collect();
                            for seg in keyed.chunk_by(|a, b| a.1 == b.1 && a.2 == b.2) {
                                let cs = chunk.char_spacing + seg[0].1;
                                if cs != cur_char_spacing {
                                    content.set_char_spacing(cs);
                                    cur_char_spacing = cs;
                                }
                                let text: String = seg.iter().map(|k| k.0).collect();
                                show_seg(content, &text, seg[0].2);
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
                        }
                        GlyphDraw::Rtl(entry) => {
                            // ponytail: no CJK-fallback glyphs or kerning in right-to-left chunks
                            show_rtl(
                                content,
                                entry,
                                &chunk.text,
                                chunk.font_size,
                                effective_cs,
                                boundary_space,
                            );
                        }
                        GlyphDraw::Fallbacks => {
                            let keyed: Vec<(char, Option<usize>)> =
                                chunk.text.chars().map(|c| (c, fallback_idx(c))).collect();
                            // An empty chunk (an underlined tab gap) still shows its
                            // empty string, as the plain path does.
                            if keyed.is_empty() {
                                show_seg(content, "", None);
                            }
                            for seg in keyed.chunk_by(|a, b| a.1 == b.1) {
                                let text: String = seg.iter().map(|k| k.0).collect();
                                show_seg(content, &text, seg[0].1);
                            }
                            if boundary_space {
                                content.show(Str(&encode_text_for_pdf(
                                    " ",
                                    &chunk.pdf_font,
                                    &pdf_name_to_entry,
                                )));
                            }
                        }
                        GlyphDraw::Kerned(entry) => {
                            show_kerned(content, entry, &chunk.text, boundary_space);
                        }
                        GlyphDraw::Plain => {
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
                        }
                }
                }

                let thick = (chunk.font_size * 0.05).max(0.5);
                // An underlined space spans its whole justified gap: Word's
                // line runs unbroken under underlined-space-compat14-*.
                let deco_w = if chunk.word_space_bridge {
                    line.chunks.get(chunk_idx + 1).map_or(chunk.width, |next| {
                        (chunk_abs_x(chunk_idx + 1, next) - x).max(chunk.width)
                    })
                } else {
                    chunk.width
                };
                if chunk.underline && deco_w > 0.0 {
                    let entry = pdf_name_to_entry.get(chunk.pdf_font.as_str()).copied();
                    let link = chunk.hyperlink_url.is_some();
                    for (top, th) in
                        underline_lines(entry, chunk.font_size, chunk.double_underline, link, y)
                    {
                        push_decoration(&mut decorations, x, top, deco_w, th, chunk.color);
                    }
                }
                if chunk.strikethrough || chunk.dstrike {
                    let (st_y, st_thick) = pdf_name_to_entry
                        .get(chunk.pdf_font.as_str())
                        .and_then(|e| strike_line(e, chunk.font_size, y))
                        .unwrap_or((y + chunk.font_size * 0.3, thick));
                    // One line across the spaces of a struck run, like Word's.
                    // A double strike is two such lines one thickness apart,
                    // centred on the single strike (17 corpus redlines, TNR,
                    // Arial, Calibri, Aptos, Gill Sans at 9.6–12pt).
                    let offsets: &[f32] = if chunk.dstrike { &[-1.0, 1.0] } else { &[0.0] };
                    for k in offsets {
                        push_decoration(
                            &mut decorations,
                            x,
                            st_y + k * st_thick,
                            deco_w,
                            st_thick,
                            chunk.color,
                        );
                    }
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
            run.flush(content);
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

        // A right-to-left line's end spaces sit left of its start.
        for sp in &line.end_spaces {
            let start = if sp.x_offset < 0.0 {
                line_start_x + sp.x_offset
            } else {
                x_at(line.chunks.len(), sp.x_offset)
            };
            let left = start.max(eff_margin);
            let w = (start + sp.width).min(line_right) - left;
            if w > 0.0 {
                let entry = pdf_name_to_entry.get(sp.pdf_font.as_str()).copied();
                for (top, th) in underline_lines(entry, sp.font_size, sp.double, false, y) {
                    push_decoration(&mut decorations, left, top, w, th, sp.color);
                }
            }
        }

        // Draw inline images outside text block. Every inline picture sits on the
        // baseline (see inline_image_line_extra), whatever its height, as laid
        // out: Word draws pictures off the text's grid and unstretched (Word for
        // Mac probes: 1-45pt pictures down an A4 page).
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
                let box_bottom = raw_y + chunk.y_offset;

                // The chunk box is the rotated frame's bounding box; draw the picture at
                // its natural size centred in it, turned about that centre like Word.
                let (w, h) = chunk.inline_image_size;
                let x = box_x + (chunk.width - w - chunk.inline_image_trim) / 2.0;
                let img_bottom = box_bottom + (chunk.inline_image_height - h) / 2.0;
                let thick = (chunk.font_size * 0.05).max(0.5);
                let entry = pdf_name_to_entry.get(chunk.pdf_font.as_str());
                if chunk.strikethrough {
                    let mid = box_bottom + chunk.inline_image_height / 2.0;
                    push_decoration(
                        &mut decorations,
                        box_x,
                        mid - 0.72,
                        chunk.width,
                        1.44,
                        chunk.color,
                    );
                    let (st_y, st_thick) = entry
                        .and_then(|e| strike_line(e, chunk.font_size, y))
                        .unwrap_or((y + chunk.font_size * 0.3, thick));
                    push_decoration(
                        &mut decorations,
                        box_x,
                        st_y,
                        chunk.width,
                        st_thick,
                        chunk.color,
                    );
                }
                if chunk.underline {
                    for (bottom, th) in
                        underline_lines(entry.copied(), chunk.font_size, false, false, y)
                    {
                        push_decoration(
                            &mut decorations,
                            box_x,
                            bottom,
                            chunk.width,
                            th,
                            chunk.color,
                        );
                    }
                }
                let clipped = push_clip(content, picture_clip);
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
                if clipped {
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

        let same_ruby = |a: &WordChunk, b: &WordChunk| matches!((&a.ruby, &b.ruby), (Some(x), Some(y)) if Arc::ptr_eq(x, y));
        let mut ci = 0;
        for group in line.chunks.chunk_by(same_ruby) {
            let (first, last) = (&group[0], &group[group.len() - 1]);
            if let Some(ruby) = &first.ruby {
                let x0 = chunk_abs_x(ci, first);
                let span = chunk_abs_x(ci + group.len() - 1, last) + last.width - x0;
                draw_ruby_guide(
                    content,
                    ruby,
                    pdf_name_to_entry.get(first.pdf_font.as_str()).copied(),
                    (x0, span),
                    y + first.y_offset,
                    first.color,
                );
            }
            ci += group.len();
        }
    }
    if current_color.is_some() {
        content.set_fill_gray(0.0);
    }
    if let Some(lt) = link_tags {
        lt.finish(content);
    }
}

/// A ruby's guide text `raise` above its base's baseline, spread over the base's
/// `span` (Word's distributeSpace: equal gaps, half a gap at each end).
fn draw_ruby_guide(
    content: &mut Content,
    ruby: &crate::model::Ruby,
    entry: Option<&FontEntry>,
    (x0, span): (f32, f32),
    baseline: f32,
    color: Option<[u8; 3]>,
) {
    let Some(entry) = entry else { return };
    let n = ruby.text.chars().count().max(1) as f32;
    let gap = ((span - entry.word_width(&ruby.text, ruby.font_size, false)) / n).max(0.0);
    content.save_state();
    fill_color_or_black(content, color);
    content.set_char_spacing(gap);
    super::charts::show_text(
        content,
        &entry.pdf_name,
        ruby.font_size,
        x0 + gap / 2.0,
        baseline + ruby.raise,
        &ruby.text,
        Some(entry),
    );
    content.restore_state();
}

/// Clip to `rect` (x, y, w, h) in a saved state; returns whether one was saved.
pub(super) fn push_clip(content: &mut Content, rect: Option<(f32, f32, f32, f32)>) -> bool {
    let Some((x, y, w, h)) = rect else {
        return false;
    };
    content.save_state();
    content.rect(x, y, w, h);
    content.clip_nonzero();
    content.end_path();
    true
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tj_run_moves_each_chunk_to_its_start_and_joins_text() {
        let mut run = TjRun::default();
        run.start(72.0, 700.0);
        run.advance_to(0.0, 12.0);
        run.show(b"ab".to_vec(), 900.0);
        run.show(b" ".to_vec(), 250.0);
        // The next word starts 1500.44 thousandths in: 350.4 right of the pen.
        run.advance_to(1500.44, 12.0);
        run.show(b"c".to_vec(), 500.0);
        assert!((run.pen - 2000.4).abs() < 1e-3);
        assert!(run.continues(700.0) && !run.continues(690.0));
        let mut content = Content::new();
        run.flush(&mut content);
        assert_eq!(content.finish().as_slice(), b"[(ab )-350.4(c)] TJ");
        assert!(run.items.is_empty());
    }
}
