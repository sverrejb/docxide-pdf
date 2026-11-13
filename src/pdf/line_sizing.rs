//! Line boxes: heights, ascents and baselines from a line's runs, document grid snapping.

use std::collections::HashMap;

use crate::fonts::{FontEntry, font_key_buf};
use crate::model::{LineSpacing, Paragraph, Run};

use super::layout::{TextLine, effective_font_size, is_break_space, math_box};
use super::line_render::quarter_pt;

/// The faces registered for glyphs the document's own fonts lack
/// (`fonts::GLYPH_FALLBACK`), in the font map's order.
pub(super) fn glyph_fallbacks(fonts: &HashMap<String, FontEntry>) -> Vec<&FontEntry> {
    fonts
        .iter()
        .filter(|(key, _)| key.starts_with(super::fonts::GLYPH_FALLBACK))
        .map(|(_, e)| e)
        .collect()
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
    ascent_scale: f32,
    top_aligned: bool,
) {
    let by_pdf_name = super::fonts::by_pdf_name(fonts);
    let fallbacks = glyph_fallbacks(fonts);
    for line in lines.iter_mut() {
        if top_aligned {
            for c in line
                .chunks
                .iter_mut()
                .filter(|c| c.math.is_none() && c.raise != 0.0)
            {
                let drop = by_pdf_name
                    .get(c.pdf_font.as_str())
                    .map_or((-c.raise).max(0.0), |e| {
                        top_aligned_drop(e, c.line_font_size, c.raise)
                    });
                c.y_offset -= c.raise + (quarter_pt(drop) - 0.25).max(0.0);
                c.raise = -drop;
            }
        }
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
        let mut unraised_below = 0.0f32;
        let mut single = 0.0f32;
        for mb in line.chunks.iter().filter_map(|c| c.math.as_ref()) {
            ascent = ascent.max(mb.line_ascent);
            below = below.max(mb.line_descent);
            unraised_below = unraised_below.max(mb.line_descent);
            single = single.max(mb.line_ascent + mb.line_descent);
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
            let own = (lhr.unwrap_or(1.2), ar.unwrap_or(0.75));
            // A form check box's ☐ is a placeholder: it is drawn as a square.
            let (lhr, ar) = if c.checkbox.is_some() {
                own
            } else {
                with_fallback_faces(own, entry, &c.text, &fallbacks)
            };
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
            let ruby_top = c.ruby.as_ref().map_or(0.0, |r| r.raise + r.font_size * ar);
            ascent = ascent.max(ruby_top);
            single = single.max(
                c.line_font_size * lhr + 2.0 * pad + (ruby_top - c.line_font_size * ar).max(0.0),
            );
            if top_aligned {
                ascent = ascent.max(c.line_font_size * ar + pad);
                below = below.max(c.line_font_size * (lhr - ar).max(0.0) + pad - c.raise);
            } else {
                ascent = ascent.max((c.line_font_size * ar + pad + c.raise).max(0.0));
                below =
                    below.max((c.line_font_size * (lhr - ar).max(0.0) + pad - c.raise).max(0.0));
                unraised_below = unraised_below.max(c.line_font_size * (lhr - ar).max(0.0) + pad);
            }
            sized = true;
        }
        if !sized {
            continue;
        }
        let (above_extra, below_extra) = math_line_stretch(line, ascent, below, line_spacing);
        // A multiple adds (or takes) its leading once, as much as the tallest
        // run's own single line would get: the highest ascent and the deepest
        // descent can come from different runs (Word for Mac probes: 12pt
        // Verdana beside 12pt Courier New steps 15.67 at single, 14.21 at 0.9,
        // 17.86 at 1.15 and 30.25 at 2 lines; 15.67 + (m − 1) × 14.58).
        let natural = match line_spacing {
            LineSpacing::Auto(m) => ascent + below + (m - 1.0) * single,
            _ => super::helpers::resolve_line_h(line_spacing, 1.0, Some(ascent + below)),
        };
        let pitch = natural + above_extra + below_extra;
        // A line an at-least minimum sets stands on its foot: raised or
        // lowered text moves its baseline by the descent it changes, so a
        // wholly raised line's glyphs stay put and a raised run beside plain
        // text rises (Word examples, 10pt Arial under 40pt, ±1pt).
        let min_sets_line = matches!(line_spacing, LineSpacing::AtLeast(_))
            && !top_aligned
            && (pitch - para_pitch).abs() <= 0.01
            && ascent + below + above_extra + below_extra < pitch - 0.01;
        if min_sets_line {
            if (unraised_below - below).abs() > 0.01 {
                line.pitch = Some(pitch);
                line.ascent_shift = unraised_below - below;
            }
        } else if (pitch - para_pitch).abs() > 0.01
            || (line.chunks.iter().any(|c| c.raise != 0.0) && (ascent - para_ascent).abs() > 0.01)
        {
            line.pitch = Some(pitch);
            line.ascent_shift = (ascent - para_ascent) * ascent_scale + above_extra;
            line.text_box = matches!(line_spacing, LineSpacing::Auto(m) if m > 1.0)
                .then_some(ascent + below + above_extra + below_extra);
        }
    }
}

/// A run's (line_h_ratio, ascender_ratio) grown to the box of every fallback
/// face drawing one of its characters: Word sizes the line by the highest
/// ascent and the deepest descent of all faces on it. A 14pt Arial line
/// holding ☑ (Segoe UI Symbol) steps 21.5pt at 1.15, not 18.5.
fn with_fallback_faces(
    (lhr, ar): (f32, f32),
    entry: &FontEntry,
    text: &str,
    fallbacks: &[&FontEntry],
) -> (f32, f32) {
    if entry.missing_chars.is_empty() {
        return (lhr, ar);
    }
    let (asc, desc) = text
        .chars()
        .filter(|c| entry.missing_chars.contains(c) && !crate::docx::is_east_asian_char(*c))
        .filter_map(|c| {
            fallbacks
                .iter()
                .find(|f| f.char_to_gid.as_ref().is_some_and(|m| m.contains_key(&c)))
        })
        .filter_map(|f| Some((f.plain_ascender_ratio?, f.plain_line_h_ratio?)))
        .fold((ar, lhr - ar), |(a, d), (fa, fl)| {
            (a.max(fa), d.max(fl - fa))
        });
    (asc + desc, asc)
}

/// How far a run moved by `raise` sits, and stretches a
/// `w:textAlignment="top"` line, below. Word hangs the glyphs from the line
/// top there, so the external leading and one font unit sit under them and
/// absorb a drop first; a raised run stays put (Word for Mac probes: Arial,
/// Calibri and Times New Roman 11pt at position -8 to +2 draw 2pt lowered
/// Arial 1.64pt down, raised text where unraised text is). Probes (2026-10-22, online export) of Calibri, Arial
/// and Times New Roman at 11 and 24pt lowered 0.5-7pt: growth = drop −
/// leading, ≥ 0; Word for Mac probes 2026-11-06: one unit less again.
pub(super) fn top_aligned_drop(entry: &FontEntry, size: f32, raise: f32) -> f32 {
    (-raise - size * entry.drop_slack_ratio.unwrap_or(0.0)).max(0.0)
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
        LineSpacing::Auto(_) | LineSpacing::AtLeast(_) => up + down,
        LineSpacing::Exact(_) => 0.0,
    }
}

/// True when a paragraph has no visible text (may still have phantom font-info runs).
pub(super) fn is_text_empty(runs: &[Run]) -> bool {
    runs.iter().all(|r| {
        r.vanish || (r.text.is_empty() && !r.is_tab && !r.is_line_break && r.inline_image.is_none())
    })
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
        .filter_map(Run::shown_image)
        .map(|img| img.display_height + img.layout_extra_height)
        .fold(0.0f32, f32::max)
}

/// A paragraph's text line box, for drawing its baselines on Word's grid.
#[derive(Clone, Copy)]
pub(super) struct LineGrid {
    /// PDF y of the page top, where the grid starts.
    pub page_height: f32,
    /// Line top to first baseline.
    pub ascent: f32,
    /// The text box: one line of the paragraph's font before line spacing
    /// adds leading to it (the whole line for exact and shrunk spacing).
    pub natural: f32,
    pub line_pitch: f32,
    pub text_ascent: f32,
    /// At-least spacing puts its leading above the text, multiples below.
    pub leading_above: bool,
    /// Lines snapped to the document grid: only the writer's stretch.
    pub snapped: bool,
    /// How far the section's vertical alignment moves the page body down.
    pub shift: f32,
    pub label_boost: f32,
    pub space_after: f32,
}

impl LineGrid {
    /// `text_h` is one line of the paragraph's font at single spacing.
    pub(super) fn new(
        page_height: f32,
        ascent: f32,
        ls: LineSpacing,
        line_h: f32,
        text_h: f32,
        text_ascent: f32,
        space_after: f32,
    ) -> Self {
        let natural = match ls {
            LineSpacing::Exact(_) => line_h,
            LineSpacing::Auto(m) if m < 1.0 => line_h,
            _ => text_h.min(line_h),
        };
        LineGrid {
            page_height,
            ascent,
            natural,
            line_pitch: line_h,
            text_ascent,
            leading_above: matches!(ls, LineSpacing::AtLeast(_)),
            snapped: false,
            shift: 0.0,
            label_boost: 0.0,
            space_after,
        }
    }

    /// The PDF y a line laid out with its baseline at `y` is drawn at. A
    /// picture taller than the text ascent rounds the baseline itself, a
    /// smaller one leaves the line to the text (Word for Mac probes: 1-45pt
    /// pictures in Arial 11 and Times New Roman 12 lines). Vertical alignment
    /// moves the line before Word rounds and stretches it; `y` and the result
    /// are before that move.
    pub(super) fn drawn_baseline(
        self,
        y: f32,
        line: Option<&TextLine>,
        first: bool,
        last: bool,
    ) -> f32 {
        let below_top = self.page_height + self.shift - y;
        let baseline = match line {
            _ if self.snapped => below_top,
            Some(l) if inline_image_line_extra(l, self.text_ascent) > 0.0 => {
                to_grid(below_top as f64) as f32
            }
            _ => {
                let boost = if first { self.label_boost } else { 0.0 };
                let ascent =
                    self.ascent - self.label_boost + boost + line.map_or(0.0, |l| l.ascent_shift);
                let advance = line.and_then(|l| l.pitch).unwrap_or(self.line_pitch);
                let natural = if let Some(text_box) = line.and_then(|l| l.text_box) {
                    text_box
                } else if !self.leading_above && self.line_pitch > 0.0 {
                    self.natural * advance / self.line_pitch
                } else if advance > self.line_pitch {
                    advance
                } else {
                    self.natural
                };
                let top = below_top - ascent;
                let after = if last && !self.leading_above && advance - natural <= 0.001 {
                    self.space_after.max(0.0)
                } else {
                    0.0
                };
                grid_baseline(
                    top,
                    ascent,
                    natural + boost,
                    advance + boost + after,
                    self.leading_above,
                )
            }
        };
        self.page_height + self.shift - writer_stretched(baseline, self.page_height)
    }
}

pub(super) fn baseline_drop(line: &TextLine, ascent: f32) -> f32 {
    line.ascent_shift + inline_image_line_extra(line, ascent)
}

/// Where the first line's baseline is drawn, which its list label shares.
pub(super) fn first_drawn_baseline(
    first_baseline_y: f32,
    lines: &[TextLine],
    text_ascent: f32,
    grid: Option<LineGrid>,
    last: bool,
) -> f32 {
    let first = lines.first();
    let y = first_baseline_y - first.map_or(0.0, |l| baseline_drop(l, text_ascent));
    grid.map_or(y, |g| g.drawn_baseline(y, first, true, last))
}

fn to_grid(v: f64) -> f64 {
    (v * 4.0 + 0.5).floor() / 4.0
}

/// Where Word draws a baseline, in points below the page top. Word keeps
/// line tops exact but draws on a 288 dpi grid (0.25pt) from the page top:
/// the line's box runs from its rounded top to its rounded bottom (`advance`
/// below the top), the rounded leading comes off it and the text box gets
/// the rest. A text box rounded up keeps the rounded descent above its foot,
/// one rounded down loses half the difference above the baseline. A line box
/// a cell longer than its rounding keeps the baseline at its top, one a cell
/// shorter at its bottom. Leading below the text (multiples above single)
/// only ever lowers the baseline from its rounded ascent: the text box may
/// spill into it (Word for Mac probes, fixture line-grid: Arial, Times New
/// Roman, Calibri and Verdana 6-24pt at single, 0.9 to 2 lines, at-least and
/// exact spacing, top margins in twip steps).
fn grid_baseline(top: f32, ascent: f32, natural: f32, advance: f32, leading_above: bool) -> f32 {
    let r = to_grid;
    let (top, a, nat, adv) = (top as f64, ascent as f64, natural as f64, advance as f64);
    let lead = (adv - nat).max(0.0);
    let text_box = r(adv) - r(lead);
    if !leading_above && lead > 0.001 {
        return (r(top) + r(a + (text_box - nat).max(0.0))) as f32;
    }
    let a = if leading_above { a - lead } else { a };
    let o = r(lead)
        + if text_box >= nat {
            r(text_box - (nat - a))
        } else {
            r(a - (nat - text_box) / 2.0)
        };
    (r(top) + o).min(r(top + adv) - (r(adv) - o)) as f32
}

/// Where Word draws a rule whose `band` is centred on PDF `y`: the band's
/// top edge on the 0.25pt grid from `page_top`, then stretched like a
/// baseline, so a 0.75pt band's centre sits between grid lines
/// (alpharetta_school_governance_council's 0.5pt and paragraph-border-indents'
/// 0.75pt table and paragraph rules within 0.02pt). A zero band snaps the point.
pub(super) fn rule_on_grid(y: f32, page_top: f32, band: f32) -> f32 {
    let top_edge = to_grid((page_top - y - band / 2.0) as f64) as f32;
    page_top - writer_stretched(top_edge + band / 2.0, page_top)
}

/// How far below the page top Word's PDF writer puts a baseline or rule the
/// layout drew `y` below it: it stretches the page by one grid cell (0.25pt)
/// over its height, the page first rounded to the grid (A4's 841.9pt box is
/// 842), so a line 700pt down an A4 page lands 0.21pt lower (Word for Mac
/// probes, fixture line-grid: 1pt lines down pages 500-1000pt tall). Pictures
/// keep their place (case181's header pictures).
fn writer_stretched(y: f32, page_height: f32) -> f32 {
    let page = quarter_pt(page_height);
    if page <= 0.25 {
        return y;
    }
    y * page / (page - 0.25)
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
pub(super) fn inline_line_advance(line: &TextLine, line_pitch: f32, metrics: (f32, f32)) -> f32 {
    line_own_advance(line, line_pitch, metrics) + line.gap_after
}

/// `inline_line_advance` without the line's `gap_after`.
pub(super) fn line_own_advance(
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

/// winDescent as a fraction of the font size: the line-height ratio less the
/// ascender ratio (identity used throughout), 0.25 when the font is unknown.
pub(super) fn descender_ratio(lhr: Option<f32>, ar: Option<f32>) -> f32 {
    lhr.zip(ar).map_or(0.25, |(l, a)| (l - a).max(0.0))
}

/// Compute the effective font_size, line_h_ratio, and ascender_ratio for a set of runs
/// by picking the run that produces the tallest visual ascent (font_size * ascender_ratio).
/// A tab never raises its line (mandated_reporter: 12pt tabs between 11pt
/// footer text leave Word's footer baseline at the 11pt line), but sizes a
/// line of nothing else. A hidden run sizes nothing beside a shown one.
pub(super) fn tallest_run_metrics(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> (f32, Option<f32>, Option<f32>) {
    let shown = || runs.iter().filter(|r| !r.vanish);
    tallest_by_ascent(shown().filter(|r| !r.is_tab), seen_fonts)
        .or_else(|| tallest_by_ascent(shown(), seen_fonts))
        .or_else(|| tallest_by_ascent(runs.iter(), seen_fonts))
        .unwrap_or((runs.first().map_or(12.0, |r| r.font_size), None, None))
}

/// (font_size, line_h_ratio, ascender_ratio) of the paragraph mark, None
/// when it has no size of its own.
pub(super) fn mark_metrics(
    para: &Paragraph,
    seen_fonts: &HashMap<String, FontEntry>,
) -> Option<(f32, Option<f32>, Option<f32>)> {
    let (fs, (lhr, ar)) = super::mark_font(para, para.paragraph_mark_font_size?, seen_fonts);
    Some((fs, lhr, ar))
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
/// None for an ordinary line. An exact line puts its baseline 80% down the box
/// whatever the font (Word probes: Calibri, Times New Roman, Arial and Cambria
/// at 8-24pt; online export 2026-10-30, issue #139: KaiTi, SimSun, SimHei and
/// Microsoft YaHei at 12-24pt in 14-40pt boxes). An at-least box whose minimum
/// wins (online export: czech_census's 10pt lines under atLeast 12.05 start
/// 0.55pt lower) bottom-aligns at the glyphs' descent: line_h_ratio −
/// ascender_ratio, less the East Asian leading Word puts below normal lines.
pub(super) fn boxed_line_ascent(
    ls: LineSpacing,
    line_h: f32,
    font_size: f32,
    lhr: Option<f32>,
    ar: Option<f32>,
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> Option<f32> {
    let min = match ls {
        LineSpacing::Exact(_) => return Some(line_h * 0.8),
        LineSpacing::AtLeast(min) => min,
        LineSpacing::Auto(_) => return None,
    };
    let (lhr, ar) = (lhr?, ar?);
    (lhr > ar && min > font_size * lhr).then(|| {
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
/// → +0 (old_blue_truck p1). A zero width space draws nothing and adds no
/// descent (textbox-oversized-inline-image: 50pt picture after U+200B at
/// 1.13 lines steps the same 51.6 as the picture alone).
pub(super) fn picture_line_bottom(
    runs: &[Run],
    para: &Paragraph,
    seen_fonts: &HashMap<String, FontEntry>,
    ls: LineSpacing,
) -> f32 {
    let shown = || runs.iter().filter(|r| !r.vanish);
    let text_runs = || shown().filter(|r| r.inline_image.is_none());
    // With no text run the mark's font sets the leading (czech_village's
    // header logo under 1.5 lines of 12pt Times New Roman).
    let leading = tallest_by_ascent(text_runs(), seen_fonts)
        .or_else(|| mark_metrics(para, seen_fonts))
        .map_or(0.0, |(fs, lhr, _)| {
            super::helpers::picture_leading(ls, fs, lhr)
        });
    let glyph_runs = text_runs().filter(|r| {
        r.text
            .chars()
            .any(|c| !c.is_whitespace() && c != '\u{200B}')
    });
    // An underlined picture keeps its run font's descent: the underline goes there.
    let underlined_pictures = shown().filter(|r| r.underlined_picture());
    let descent = tallest_by_ascent(glyph_runs.chain(underlined_pictures), seen_fonts)
        .map_or(0.0, |(fs, lhr, ar)| fs * descender_ratio(lhr, ar));
    descent.max(inserted_picture_descent(runs, seen_fonts)) + leading
}

/// Markup view: an inserted picture is underlined, and Word keeps its run's
/// descent below the picture as under underlined text (Word for Mac probe
/// p22d: 50pt pictures alone in 12pt Times New Roman paragraphs push the next
/// baseline 2.6pt further than plain ones at single, 1.5 and triple spacing;
/// deleted pictures, p22e, add nothing). The run is textless: a lone picture's
/// run keeps its flags after the picture moves to `Paragraph::image`.
pub(super) fn inserted_picture_descent(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> f32 {
    let inserted = runs.iter().filter(|r| {
        r.revised && r.underline && !r.vanish && (r.inline_image.is_some() || r.text.is_empty())
    });
    tallest_by_ascent(inserted, seen_fonts)
        .map_or(0.0, |(fs, lhr, ar)| fs * descender_ratio(lhr, ar))
}

/// The line a picture `picture_h` tall makes: the picture plus the text
/// descent, never less than the pitch `line_h`, which a line without one
/// keeps (the laid-out form is `inline_line_advance`).
pub(super) fn line_h_with_picture(
    picture_h: f32,
    runs: &[Run],
    para: &Paragraph,
    seen_fonts: &HashMap<String, FontEntry>,
    ls: LineSpacing,
    line_h: f32,
) -> f32 {
    if picture_h > 0.0 {
        line_h.max(picture_h + picture_line_bottom(runs, para, seen_fonts, ls))
    } else {
        line_h
    }
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

/// A run whose glyphs can size its line: shown, not a tab, and empty (a
/// paragraph mark) or holding more than whitespace. A hidden run sizes
/// nothing (Word example, fixture underlined-picture-places: a paragraph of
/// only a hidden 16pt picture run is a line of its 10pt mark).
fn sizes_line(r: &Run) -> bool {
    !r.vanish && !r.is_tab && (r.text.is_empty() || !r.text.trim().is_empty())
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
    // A line-spacing multiple is a floor of m unsnapped pitches under the cells
    // the glyphs need: 1.5 lines of one 18pt cell is 27pt (case79), while text
    // needing two 15.6pt cells stays 31.2pt at 1.25, 1.5 or 2 lines (Word probes).
    match effective_ls {
        crate::model::LineSpacing::Auto(m) if grid_h > 0.0 => {
            grid_cells(grid_h, pitch).max(m * pitch)
        }
        _ => grid_cells(line_h, pitch),
    }
}

/// Whole docGrid cells for a line `h` tall, at least one. Tolerance so an
/// exact fit stays one cell despite f32 error.
pub(super) fn grid_cells(h: f32, pitch: f32) -> f32 {
    ((h / pitch) - 0.02).ceil().max(1.0) * pitch
}

/// Word counts a grid-snapped paragraph's cells line by line, from the runs on
/// each line: an English paper in a Chinese template (Times New Roman 12 under
/// a 312-twip grid, DengXian for its one CJK word) takes two cells for the
/// line holding that word and one cell for every other, where the paragraph's
/// tallest run (`grid_snapped_line_h`) would make them all two (a corpus
/// document: Word's lines step 15.6pt, ours 31.2). A line whose cells differ
/// from the paragraph's gets its own `pitch` and sits in its own cell
/// (`grid_baseline_offset` of its runs, relative to the paragraph's).
/// ponytail: body paragraphs under a line-spacing multiple only; tables,
/// headers and footers keep one box per paragraph.
pub(super) fn size_grid_lines(
    lines: &mut [TextLine],
    fonts: &HashMap<String, FontEntry>,
    effective_ls: crate::model::LineSpacing,
    para_cells: f32,
    pitch: f32,
    para_baseline: f32,
) {
    let crate::model::LineSpacing::Auto(m) = effective_ls else {
        return;
    };
    let by_pdf_name = super::fonts::by_pdf_name(fonts);
    for line in lines.iter_mut() {
        if line.chunks.iter().any(|c| c.inline_image_name.is_some()) {
            continue;
        }
        let (mut grid_h, mut shift) = (0.0f32, None::<f32>);
        for c in line
            .chunks
            .iter()
            .filter(|c| c.math.is_none() && !c.text.trim().is_empty())
        {
            let Some(e) = by_pdf_name.get(c.pdf_font.as_str()) else {
                continue;
            };
            if let Some(t) = e.grid_line_ratio {
                grid_h = grid_h.max(c.line_font_size * t);
            }
            if let Some(s) = e.grid_baseline_shift {
                let s = s * c.line_font_size;
                shift = Some(shift.map_or(s, |v: f32| v.max(s)));
            }
        }
        if grid_h <= 0.0 {
            continue;
        }
        let own = grid_cells(grid_h, pitch).max(m * pitch);
        if (own - para_cells).abs() > 0.01 {
            line.pitch = Some(own);
            line.ascent_shift = own / 2.0 + shift.unwrap_or(0.0) - para_baseline;
        }
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
    fn fallback_faces_grow_the_line_box() {
        // Arial (0.938 / 0.212) and Segoe UI Symbol's hhea box (1.079 / 0.251).
        let arial = FontEntry {
            missing_chars: ['☑', '中'].into_iter().collect(),
            ..FontEntry::stub()
        };
        let segoe = FontEntry {
            plain_ascender_ratio: Some(1.079),
            plain_line_h_ratio: Some(1.33),
            char_to_gid: Some([('☑', 1), ('中', 2)].into_iter().collect()),
            ..FontEntry::stub()
        };
        let own = (1.15, 0.938);
        let (lhr, ar) = with_fallback_faces(own, &arial, "A☑B", &[&segoe]);
        assert!((ar - 1.079).abs() < 1e-4 && (lhr - 1.33).abs() < 1e-4);
        // East Asian rescues and characters the font has leave the line alone.
        assert_eq!(with_fallback_faces(own, &arial, "A中B", &[&segoe]), own);
        assert_eq!(with_fallback_faces(own, &arial, "AB", &[&segoe]), own);
    }

    #[test]
    fn inserted_picture_keeps_its_runs_descent() {
        // Times New Roman's plain box: 1.149 line, 0.891 ascent → 0.258 descent.
        let tnr = FontEntry {
            plain_line_h_ratio: Some(1.149),
            plain_ascender_ratio: Some(0.891),
            ..FontEntry::stub()
        };
        let run = |revised: bool, underline: bool| Run {
            font_name: "Times New Roman".to_string(),
            font_size: 12.0,
            revised,
            underline,
            ..Run::default()
        };
        let fonts: HashMap<String, FontEntry> = [(crate::fonts::font_key(&run(true, true)), tnr)]
            .into_iter()
            .collect();
        let inserted = [run(true, true)];
        assert!((inserted_picture_descent(&inserted, &fonts) - 12.0 * 0.258).abs() < 0.01);
        // A deleted (struck) picture and a plain one add nothing.
        assert_eq!(inserted_picture_descent(&[run(true, false)], &fonts), 0.0);
        assert_eq!(inserted_picture_descent(&[run(false, false)], &fonts), 0.0);
    }

    #[test]
    fn top_aligned_drop_spends_external_leading_first() {
        // Times New Roman's leading is 87/2048 em; Calibri has none. Both
        // keep one more font unit.
        let tnr = FontEntry {
            drop_slack_ratio: Some(88.0 / 2048.0),
            ..FontEntry::stub()
        };
        let calibri = FontEntry {
            drop_slack_ratio: Some(1.0 / 2048.0),
            ..FontEntry::stub()
        };
        assert!(top_aligned_drop(&tnr, 11.0, -0.5) < 0.05);
        assert!((top_aligned_drop(&tnr, 24.0, -7.0) - 5.97).abs() < 0.01);
        assert!((top_aligned_drop(&calibri, 11.0, -0.5) - 0.4946).abs() < 0.001);
        assert_eq!(top_aligned_drop(&calibri, 11.0, 7.0), 0.0);
    }

    /// Word's rules in paragraph-border-indents (Letter) and
    /// alpharetta_school_governance_council: layout centre below the page top
    /// → drawn centre.
    #[test]
    fn rules_snap_their_top_edge_then_stretch() {
        let drawn = |below: f32, band: f32| 792.0 - rule_on_grid(792.0 - below, 792.0, band);
        assert!((drawn(84.87, 0.75) - 84.90).abs() < 0.01);
        assert!((drawn(388.01, 0.75) - 388.25).abs() < 0.01);
        assert!((drawn(174.88, 0.5) - 175.05).abs() < 0.01);
    }
}
