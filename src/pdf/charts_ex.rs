//! Office 2016 charts (the chartEx part) as Word lays them out: waterfall,
//! histogram, pareto, box & whisker, funnel, treemap and sunburst. All their
//! text takes the theme minor font at the chartStyle part's sizes and
//! colours, and every type fills one plot box inside the chart's frame.
//!
//! Measured on Word's exports of `tests/output/torture_test/probe3d/inx*`
//! (12pt and 9pt text, with and without title and legend; Cambria).

use std::f32::consts::{FRAC_PI_2, TAU};

use pdf_writer::Content;
use pdf_writer::types::{LineCapStyle, LineJoinStyle};

use crate::fonts::FontEntry;
use crate::model::{
    CHARTEX_GRID as GRID, Chart, ChartExText, ChartText, ChartType, InlineChart, LegendPosition,
    WATERFALL_LEGEND,
};

use super::chart_legend::{SwatchStyle, render_swatch};
use super::charts::{polyline, resolve_accent_colors, show_text, show_text_matrix, text_width};
use super::charts_radial::{arc_segment_count, emit_arc};
use super::color::{fill_rgb, stroke_rgb};
use super::helpers::draw_circle;

const WHITE: [u8; 3] = [255, 255, 255];
/// A waterfall's connecting lines (the style's `seriesLine`, tx1 at 35%).
const CONNECTOR: [u8; 3] = [0xA6, 0xA6, 0xA6];

/// The plot box's insets from the frame: left, right, top without a title,
/// bottom without a legend.
const PAD_L: f32 = 5.25;
const PAD_R: f32 = 5.4;
const PAD_T: f32 = 5.2;
const PAD_B: f32 = 5.5;
/// A title's baseline sits `TITLE_TOP` + its size below the frame (Cambria
/// at 14 and 18.5pt both fit to 0.02pt); the plot box starts `TITLE_GAP`
/// under the title's descent.
const TITLE_TOP: f32 = 4.67;
const TITLE_GAP: f32 = 3.75;
/// A bottom legend's band, its keys' centre above the frame's bottom, and
/// the row's centre right of the frame's middle.
const LEGEND_BAND: f32 = 28.75;
const LEGEND_KEY_MID: f32 = 18.12;
const LEGEND_SHIFT: f32 = 1.6;
const KEY: f32 = 7.0;
const KEY_GAP: f32 = 3.9;
const ENTRY_GAP: f32 = 7.37;
/// A right legend: the plot box ends `RIGHT_LEGEND_GAP` before it, its
/// entries padded `RIGHT_LEGEND_PAD` each side and `RIGHT_LEGEND_ROW` +
/// the line height apart (tm_rlegend, 9pt: rows 18pt apart).
const RIGHT_LEGEND_GAP: f32 = 7.05;
const RIGHT_LEGEND_PAD: f32 = 5.45;
const RIGHT_LEGEND_ROW: f32 = 7.45;
/// Value labels end this far left of the plot; category labels hang this
/// far below it; the plot ends this far inside the box's right edge.
const VAL_LABEL_GAP: f32 = 6.4;
const CAT_LABEL_GAP: f32 = 6.7;
const AXIS_INSET_R: f32 = 1.35;
/// A data label's box stands this far off the end of its bar.
const DATA_LABEL_GAP: f32 = 0.2;

/// Top-down box in PDF space: `top` > `bottom`.
#[derive(Clone, Copy, Debug, PartialEq)]
struct Rect {
    l: f32,
    r: f32,
    top: f32,
    bottom: f32,
}

impl Rect {
    fn w(&self) -> f32 {
        self.r - self.l
    }
    fn h(&self) -> f32 {
        self.top - self.bottom
    }
}

/// The label font and its line box (ascent, descent per em).
struct Pen<'a> {
    key: &'a str,
    font: Option<&'a FontEntry>,
    asc: f32,
    desc: f32,
}

impl Pen<'_> {
    fn width(&self, text: &str, t: ChartText) -> f32 {
        text_width(text, t.size, self.font)
    }
    fn line_h(&self, t: ChartText) -> f32 {
        (self.asc + self.desc) * t.size
    }
    /// The baseline that centres a line on `y`.
    fn centre_baseline(&self, y: f32, t: ChartText) -> f32 {
        y - (self.asc - self.desc) / 2.0 * t.size
    }
    fn text(&self, content: &mut Content, x: f32, baseline: f32, text: &str, t: ChartText) {
        if self.font.is_none() {
            return;
        }
        fill_rgb(content, t.color);
        show_text(content, self.key, t.size, x, baseline, text, self.font);
    }
    fn centred(&self, content: &mut Content, cx: f32, baseline: f32, text: &str, t: ChartText) {
        self.text(content, cx - self.width(text, t) / 2.0, baseline, text, t);
    }
    /// Text centred on (`cx`, `cy`) and turned `angle` radians anticlockwise.
    fn rotated(
        &self,
        content: &mut Content,
        cx: f32,
        cy: f32,
        angle: f32,
        text: &str,
        t: ChartText,
    ) {
        if self.font.is_none() {
            return;
        }
        let (sin, cos) = angle.sin_cos();
        let (dx, dy) = (-self.width(text, t) / 2.0, self.centre_baseline(0.0, t));
        let (x, y) = (cx + dx * cos - dy * sin, cy + dx * sin + dy * cos);
        fill_rgb(content, t.color);
        show_text_matrix(
            content,
            self.key,
            t.size,
            [cos, sin, -sin, cos, x, y],
            text,
            self.font,
        );
    }
}

/// A number as Excel's General format shows it.
fn general(v: f32) -> String {
    format!("{}", (v * 1e6).round() / 1e6)
}

pub(super) fn render(
    chart: &InlineChart,
    content: &mut Content,
    x: f32,
    y: f32,
    key: &str,
    font: Option<&FontEntry>,
) {
    let c = &chart.chart;
    let (Some(text), Some(series)) = (c.ex_text, c.series.first()) else {
        return;
    };
    let (w, h) = (chart.display_width, chart.display_height);
    let asc = font.and_then(|f| f.plain_ascender_ratio).unwrap_or(0.9);
    let line = font.and_then(|f| f.plain_line_h_ratio).unwrap_or(1.15);
    let pen = Pen {
        key,
        font,
        asc,
        desc: line - asc,
    };
    let accents = resolve_accent_colors(&c.accent_colors);
    let labels = c
        .cat_axis
        .as_ref()
        .map(|a| a.labels.as_slice())
        .unwrap_or(&[]);
    let groups = parent_groups(&c.parent_labels, series.values.len());

    content.save_state();
    let mut plot = Rect {
        l: x + PAD_L,
        r: x + w - PAD_R,
        top: y - PAD_T,
        bottom: y - h + PAD_B,
    };
    if let Some(title) = &c.title {
        let baseline = y - TITLE_TOP - text.title.size;
        pen.centred(content, x + w / 2.0, baseline, title, text.title);
        plot.top = baseline - pen.desc * text.title.size - TITLE_GAP;
    }

    let legend: Vec<(&str, [u8; 3])> = match c.chart_type {
        ChartType::Waterfall { .. } => WATERFALL_LEGEND
            .iter()
            .enumerate()
            .map(|(i, n)| (*n, accents[i % accents.len()]))
            .collect(),
        ChartType::Treemap | ChartType::Sunburst => groups
            .iter()
            .enumerate()
            .map(|(i, (n, _))| (*n, accents[i % accents.len()]))
            .collect(),
        _ => vec![(series.label.as_str(), accents[0])],
    };
    // ponytail: top and left legends are laid out as bottom and right ones;
    // no probe has them.
    let right_legend = c
        .legend
        .as_ref()
        .map(|l| matches!(l.position, LegendPosition::Right | LegendPosition::Left));
    let widths: Vec<f32> = legend
        .iter()
        .map(|(n, _)| pen.width(n, text.legend))
        .collect();
    match right_legend {
        Some(false) => plot.bottom += LEGEND_BAND,
        Some(true) => {
            let widest = widths.iter().copied().fold(0.0, f32::max);
            plot.r -= KEY + KEY_GAP + widest + 2.0 * RIGHT_LEGEND_PAD + RIGHT_LEGEND_GAP;
        }
        None => {}
    }

    let ctx = Ctx {
        c,
        pen: &pen,
        text: &text,
        accents,
        labels,
    };
    match c.chart_type {
        ChartType::Funnel => draw_funnel(content, &ctx, plot),
        ChartType::Treemap => draw_treemap(content, &ctx, plot, &groups),
        ChartType::Sunburst => draw_sunburst(content, &ctx, plot, &groups),
        _ => draw_axis_chart(content, &ctx, plot),
    }

    // Bottom legends run as one centred row, right ones as a column centred
    // on the plot box.
    let keys: Vec<(f32, f32)> = match right_legend {
        Some(false) => {
            let total = widths
                .iter()
                .map(|tw| KEY + KEY_GAP + tw + ENTRY_GAP)
                .sum::<f32>()
                - ENTRY_GAP;
            let mut lx = x + w / 2.0 + LEGEND_SHIFT - total / 2.0;
            widths
                .iter()
                .map(|tw| {
                    let at = (lx, y - h + LEGEND_KEY_MID);
                    lx += KEY + KEY_GAP + tw + ENTRY_GAP;
                    at
                })
                .collect()
        }
        Some(true) => {
            let row = pen.line_h(text.legend) + RIGHT_LEGEND_ROW;
            let lx = plot.r + RIGHT_LEGEND_GAP + RIGHT_LEGEND_PAD;
            let first = (plot.top + plot.bottom) / 2.0 + (legend.len() as f32 - 1.0) * row / 2.0;
            (0..legend.len())
                .map(|i| (lx, first - i as f32 * row))
                .collect()
        }
        None => Vec::new(),
    };
    for ((name, color), (lx, mid)) in legend.iter().zip(keys) {
        render_swatch(
            content,
            &SwatchStyle::Rect,
            *color,
            lx,
            mid - KEY / 2.0,
            KEY,
        );
        let baseline = pen.centre_baseline(mid, text.legend);
        pen.text(content, lx + KEY + KEY_GAP, baseline, name, text.legend);
    }
    content.restore_state();
}

/// What every chart type's drawing reads.
struct Ctx<'a> {
    c: &'a Chart,
    pen: &'a Pen<'a>,
    text: &'a ChartExText,
    accents: &'a [[u8; 3]],
    labels: &'a [String],
}

impl Ctx<'_> {
    fn values(&self) -> &[f32] {
        &self.c.series[0].values
    }
    fn accent(&self, i: usize) -> [u8; 3] {
        self.accents[i % self.accents.len()]
    }
}

/// The outer categories in order of first appearance, each with its points;
/// without outer categories every point is its own group.
fn parent_groups(parents: &[String], n: usize) -> Vec<(&str, Vec<usize>)> {
    let mut groups: Vec<(&str, Vec<usize>)> = Vec::new();
    for i in 0..n {
        let name = parents.get(i).map_or("", String::as_str);
        match groups
            .iter_mut()
            .find(|(g, _)| *g == name && !name.is_empty())
        {
            Some((_, pts)) => pts.push(i),
            None => groups.push((name, vec![i])),
        }
    }
    groups
}

/// The value axis's major unit and top: 5% headroom over the largest value,
/// then the smallest 1-2-5 unit that needs at most ten steps (waterfall 145
/// → 160 by 20, histogram 4 → 4.5 by 0.5, box & whisker 40 → 45 by 5,
/// pareto 30 → 35 by 5 and its share axis 0–100% by 10%).
fn value_axis(max: f32) -> (f32, f32) {
    let top = max.max(0.0) * 1.05;
    if top <= 0.0 {
        return (1.0, 1.0);
    }
    let mut unit = 10f32.powf((top / 10.0).log10().floor());
    for step in [1.0, 2.0, 5.0, 10.0] {
        if (top / (unit * step) - 1e-4).ceil() <= 10.0 {
            unit *= step;
            break;
        }
    }
    (unit, (top / unit - 1e-4).ceil() * unit)
}

/// A waterfall bar's kind, which is also its accent colour's index.
#[derive(Clone, Copy, PartialEq)]
enum Step {
    Increase = 0,
    Decrease = 1,
    Total = 2,
}

/// A waterfall's bars: each point's span on the value axis. A total stands
/// from zero at the running sum it names; any other point steps from the
/// running sum.
fn waterfall_steps(values: &[f32], totals: &[usize]) -> Vec<(f32, f32, Step)> {
    let mut run = 0.0f32;
    values
        .iter()
        .enumerate()
        .map(|(i, &v)| {
            if totals.contains(&i) {
                run = v;
                (v.min(0.0), v.max(0.0), Step::Total)
            } else {
                let from = run;
                run += v;
                let step = if v >= 0.0 {
                    Step::Increase
                } else {
                    Step::Decrease
                };
                (from.min(run), from.max(run), step)
            }
        })
        .collect()
}

/// A box & whisker's statistics; whiskers reach the furthest points within
/// 1.5 IQR of the box.
struct BoxStats {
    lo: f32,
    q1: f32,
    median: f32,
    q3: f32,
    hi: f32,
    mean: f32,
    outliers: Vec<f32>,
}

fn box_stats(mut pts: Vec<f32>, exclusive: bool) -> Option<BoxStats> {
    if pts.is_empty() {
        return None;
    }
    pts.sort_by(f32::total_cmp);
    let n = pts.len() as f32;
    let q = |p: f32| {
        let rank = if exclusive {
            p * (n + 1.0)
        } else {
            p * (n - 1.0) + 1.0
        };
        let r = rank.clamp(1.0, n) - 1.0;
        let (i, f) = (r.floor() as usize, r.fract());
        pts[i] + f * (pts[(i + 1).min(pts.len() - 1)] - pts[i])
    };
    let (q1, median, q3) = (q(0.25), q(0.5), q(0.75));
    let fence = 1.5 * (q3 - q1);
    let inside = |v: &&f32| **v >= q1 - fence && **v <= q3 + fence;
    Some(BoxStats {
        lo: pts.iter().find(inside).copied().unwrap_or(q1),
        q1,
        median,
        q3,
        hi: pts.iter().rfind(inside).copied().unwrap_or(q3),
        mean: pts.iter().sum::<f32>() / n,
        outliers: pts.iter().filter(|v| !inside(v)).copied().collect(),
    })
}

/// Waterfall, histogram, pareto and box & whisker: a value axis on the left
/// (and a 0–100% one on the right for a pareto), gridlines at every step,
/// category names under their slots.
fn draw_axis_chart(content: &mut Content, ctx: &Ctx, bx: Rect) {
    let Ctx {
        c,
        pen,
        text,
        labels,
        ..
    } = *ctx;
    let series = &c.series[0];
    let values = &series.values;
    let waterfall = matches!(c.chart_type, ChartType::Waterfall { .. });
    let bars = match &c.chart_type {
        ChartType::Waterfall { totals } => waterfall_steps(values, totals),
        _ => values
            .iter()
            .map(|&v| (v.min(0.0), v.max(0.0), Step::Increase))
            .collect(),
    };
    let boxes: Vec<BoxStats> = match c.chart_type {
        ChartType::BoxWhisker { exclusive } => {
            let mut pts = vec![Vec::new(); labels.len()];
            let index = series.x_values.as_deref().unwrap_or(&[]);
            for (&v, &b) in values.iter().zip(index) {
                if let Some(p) = pts.get_mut(b as usize) {
                    p.push(v);
                }
            }
            pts.into_iter()
                .filter_map(|p| box_stats(p, exclusive))
                .collect()
        }
        _ => Vec::new(),
    };
    let n = if boxes.is_empty() {
        values.len()
    } else {
        boxes.len()
    };
    if n == 0 {
        return;
    }
    let max = if boxes.is_empty() {
        bars.iter().map(|s| s.1).fold(0.0, f32::max)
    } else {
        values.iter().copied().fold(0.0, f32::max)
    };
    let (unit, top) = value_axis(max);
    let ticks = (top / unit).round() as usize;
    let tick_labels: Vec<String> = (0..=ticks).map(|i| general(i as f32 * unit)).collect();
    let share_labels: Vec<String> = (0..=10).map(|i| format!("{}%", i * 10)).collect();
    let tick_widths: Vec<f32> = tick_labels
        .iter()
        .map(|l| pen.width(l, text.axis))
        .collect();
    let widest = |ws: &[f32]| ws.iter().copied().fold(0.0, f32::max);
    let share_w = share_labels
        .iter()
        .map(|l| pen.width(l, text.axis))
        .fold(0.0, f32::max);

    let pareto = matches!(c.chart_type, ChartType::Pareto);
    // A pareto's two label columns each keep 6.6pt inside the box, and its
    // share labels start 6.75pt right of the plot (9pt probe).
    // ponytail: Word's pareto plot also starts 1.2pt lower than the other
    // axis charts' for the same text; not derived, not applied.
    let inset = if pareto { 6.6 } else { 0.0 };
    let plot = Rect {
        l: bx.l + inset + widest(&tick_widths) + VAL_LABEL_GAP,
        r: if pareto {
            bx.r - inset - share_w - 6.75
        } else {
            bx.r - AXIS_INSET_R
        },
        top: bx.top - pen.line_h(text.axis) / 2.0,
        bottom: bx.bottom + pen.line_h(text.axis) + CAT_LABEL_GAP,
    };
    if plot.w() <= 0.0 || plot.h() <= 0.0 {
        return;
    }
    let y_of = |v: f32| plot.bottom + v / top * plot.h();

    content.set_line_width(0.75);
    stroke_rgb(content, GRID);
    for i in 1..=ticks {
        let gy = y_of(i as f32 * unit);
        content.move_to(plot.l, gy);
        content.line_to(plot.r, gy);
    }
    content.move_to(plot.l, plot.top);
    content.line_to(plot.l, plot.bottom);
    content.line_to(plot.r, plot.bottom);
    if pareto {
        content.line_to(plot.r, plot.top);
    }
    content.stroke();
    for (i, (l, lw)) in tick_labels.iter().zip(&tick_widths).enumerate() {
        let baseline = pen.centre_baseline(y_of(i as f32 * unit), text.axis);
        pen.text(content, plot.l - VAL_LABEL_GAP - lw, baseline, l, text.axis);
    }
    if pareto {
        for (i, l) in share_labels.iter().enumerate() {
            let baseline = pen.centre_baseline(plot.bottom + i as f32 / 10.0 * plot.h(), text.axis);
            pen.text(content, plot.r + 6.75, baseline, l, text.axis);
        }
    }
    let slot = plot.w() / n as f32;
    let bar_w = slot / (1.0 + c.gap_width_pct / 100.0);
    let mid = |i: usize| plot.l + (i as f32 + 0.5) * slot;
    for (i, l) in labels.iter().enumerate().take(n) {
        pen.centred(
            content,
            mid(i),
            bx.bottom + pen.desc * text.axis.size,
            l,
            text.axis,
        );
    }

    let t = text.data_label;
    let above = |v: f32| y_of(v) + pen.desc * t.size + DATA_LABEL_GAP;
    let below = |v: f32| y_of(v) - pen.asc * t.size - DATA_LABEL_GAP;
    if !boxes.is_empty() {
        draw_boxes(content, ctx, &boxes, mid, bar_w, &y_of);
        return;
    }
    for (i, (&v, &(lo, hi, step))) in values.iter().zip(&bars).enumerate() {
        let color = ctx.accent(step as usize);
        fill_rgb(content, color);
        content.rect(mid(i) - bar_w / 2.0, y_of(lo), bar_w, y_of(hi) - y_of(lo));
        content.fill_nonzero();
        if c.show_values {
            let baseline = if step == Step::Decrease {
                below(lo)
            } else {
                above(hi)
            };
            pen.centred(content, mid(i), baseline, &general(v), t);
        }
    }
    // A waterfall's running total carries from each bar to the next.
    if waterfall {
        stroke_rgb(content, CONNECTOR);
        for (i, &(lo, hi, step)) in bars.iter().enumerate().take(n - 1) {
            let level = y_of(if step == Step::Decrease { lo } else { hi });
            content.move_to(mid(i) + bar_w / 2.0, level);
            content.line_to(mid(i + 1) - bar_w / 2.0, level);
        }
        content.stroke();
    }
    if pareto {
        let total: f32 = values.iter().sum();
        if total > 0.0 {
            content.set_line_width(2.25);
            content.set_line_cap(LineCapStyle::RoundCap);
            content.set_line_join(LineJoinStyle::RoundJoin);
            stroke_rgb(content, ctx.accent(1));
            let cumulative = values.iter().scan(0.0, |run, v| {
                *run += v;
                Some(*run)
            });
            polyline(
                content,
                cumulative
                    .enumerate()
                    .map(|(i, run)| (mid(i), plot.bottom + run / total * plot.h())),
            );
            content.stroke();
        }
    }
}

/// Boxes from Q1 to Q3 with their edges, median and whiskers in 2.25pt
/// lines of the box's colour, caps 0.3 of the box wide, the mean an X 6pt
/// across, outliers 3pt dots; every statistic labelled on the box's centre
/// line (boxwhisker probes).
fn draw_boxes(
    content: &mut Content,
    ctx: &Ctx,
    boxes: &[BoxStats],
    mid: impl Fn(usize) -> f32,
    bar_w: f32,
    y_of: &impl Fn(f32) -> f32,
) {
    let Ctx { c, pen, text, .. } = *ctx;
    let t = text.data_label;
    let color = ctx.accent(0);
    content.set_line_width(2.25);
    content.set_line_cap(LineCapStyle::RoundCap);
    for (i, b) in boxes.iter().enumerate() {
        let (cx, l, r) = (mid(i), mid(i) - bar_w / 2.0, mid(i) + bar_w / 2.0);
        let cap = 0.3 * bar_w / 2.0;
        fill_rgb(content, color);
        stroke_rgb(content, color);
        content.rect(l, y_of(b.q1), bar_w, y_of(b.q3) - y_of(b.q1));
        content.fill_nonzero();
        for v in [b.q1, b.median, b.q3] {
            content.move_to(l, y_of(v));
            content.line_to(r, y_of(v));
        }
        for (from, to) in [(b.q1, b.lo), (b.q3, b.hi)] {
            content.move_to(cx, y_of(from));
            content.line_to(cx, y_of(to));
            content.move_to(cx - cap, y_of(to));
            content.line_to(cx + cap, y_of(to));
        }
        let my = y_of(b.mean);
        content.move_to(cx - 3.0, my - 3.0);
        content.line_to(cx + 3.0, my + 3.0);
        content.move_to(cx - 3.0, my + 3.0);
        content.line_to(cx + 3.0, my - 3.0);
        content.stroke();
        for &o in &b.outliers {
            draw_circle(content, cx, y_of(o), 1.5);
            content.fill_nonzero();
        }
        if c.show_values {
            for v in [b.lo, b.q1, b.median, b.q3, b.hi, b.mean]
                .iter()
                .chain(&b.outliers)
            {
                pen.centred(
                    content,
                    cx,
                    pen.centre_baseline(y_of(*v), t),
                    &general(*v),
                    t,
                );
            }
        }
    }
}

/// Bars centred down the plot, the largest full width, category names right
/// of the box's left edge and `CAT` before the axis line, values white
/// inside (funnel probes).
fn draw_funnel(content: &mut Content, ctx: &Ctx, bx: Rect) {
    let Ctx {
        c,
        pen,
        text,
        labels: names,
        ..
    } = *ctx;
    let values = ctx.values();
    /// The axis line stands 1pt inside the box; names end 6.5pt before it.
    const INSET: f32 = 1.0;
    const CAT: f32 = 6.5;
    let max = values.iter().copied().fold(0.0f32, f32::max);
    if values.is_empty() || max <= 0.0 {
        return;
    }
    let name_ws: Vec<f32> = names.iter().map(|l| pen.width(l, text.axis)).collect();
    let plot = Rect {
        l: bx.l + name_ws.iter().copied().fold(0.0, f32::max) + CAT,
        r: bx.r - AXIS_INSET_R,
        top: bx.top - INSET,
        bottom: bx.bottom + INSET,
    };
    if plot.w() <= 0.0 {
        return;
    }
    content.set_line_width(0.75);
    stroke_rgb(content, GRID);
    content.move_to(plot.l, plot.top);
    content.line_to(plot.l, plot.bottom);
    content.stroke();
    let row = plot.h() / values.len() as f32;
    let thick = row / (1.0 + c.gap_width_pct / 100.0);
    let centre = (plot.l + plot.r) / 2.0;
    for (i, &v) in values.iter().enumerate() {
        let cy = plot.top - (i as f32 + 0.5) * row;
        let bw = v / max * plot.w();
        fill_rgb(content, ctx.accent(0));
        content.rect(centre - bw / 2.0, cy - thick / 2.0, bw, thick);
        content.fill_nonzero();
        if let (Some(name), Some(nw)) = (names.get(i), name_ws.get(i)) {
            let x = plot.l - CAT - nw;
            pen.text(
                content,
                x,
                pen.centre_baseline(cy, text.axis),
                name,
                text.axis,
            );
        }
        if c.show_values {
            let t = ChartText {
                color: WHITE,
                ..text.data_label
            };
            pen.centred(content, centre, pen.centre_baseline(cy, t), &general(v), t);
        }
    }
}

/// Tiles for `values` within `r`, largest first, returned in input order.
/// Word's rows run along the shorter side and take the next value while the
/// row's mean tile aspect falls (treemap probes: plain squarifying, which
/// compares the worst tile, splits South 30 | 20 where Word keeps them in
/// one row).
fn tile(values: &[f32], r: Rect) -> Vec<Rect> {
    let mut out = vec![r; values.len()];
    let total: f32 = values.iter().sum();
    if total <= 0.0 {
        return out;
    }
    let mut order: Vec<usize> = (0..values.len()).collect();
    order.sort_by(|&a, &b| values[b].total_cmp(&values[a]));
    let scale = r.w() * r.h() / total;
    let mut rest = r;
    let mut i = 0;
    while i < order.len() {
        let short = rest.w().min(rest.h());
        let mean_aspect = |row: &[usize]| {
            let depth = row.iter().map(|&k| values[k] * scale).sum::<f32>() / short;
            row.iter()
                .map(|&k| {
                    let len = values[k] * scale / depth;
                    (len / depth).max(depth / len)
                })
                .sum::<f32>()
                / row.len() as f32
        };
        let mut j = i + 1;
        let mut aspect = mean_aspect(&order[i..j]);
        while j < order.len() {
            let next = mean_aspect(&order[i..=j]);
            if next >= aspect {
                break;
            }
            (aspect, j) = (next, j + 1);
        }
        let row = &order[i..j];
        let depth = row.iter().map(|&k| values[k] * scale).sum::<f32>() / short;
        // A tall box takes the row across its top, a wide one down its left.
        let tall = rest.w() < rest.h();
        let mut offset = 0.0;
        for &k in row {
            let len = values[k] * scale / depth;
            out[k] = if tall {
                Rect {
                    l: rest.l + offset,
                    r: rest.l + offset + len,
                    top: rest.top,
                    bottom: rest.top - depth,
                }
            } else {
                Rect {
                    l: rest.l,
                    r: rest.l + depth,
                    top: rest.top - offset,
                    bottom: rest.top - offset - len,
                }
            };
            offset += len;
        }
        if tall {
            rest.top -= depth;
        } else {
            rest.l += depth;
        }
        i = j;
    }
    out
}

/// Outer categories tiled first, then their points within them, inset 0.75pt
/// from the box with 0.25pt gaps, coloured by outer category; values at each
/// tile's bottom left, 3.6pt in and 3.5pt up (treemap probes).
fn draw_treemap(content: &mut Content, ctx: &Ctx, bx: Rect, groups: &[(&str, Vec<usize>)]) {
    let (pen, text, values) = (ctx.pen, ctx.text, ctx.values());
    let inner = Rect {
        l: bx.l + 0.75,
        r: bx.r - 0.75,
        top: bx.top - 0.5,
        bottom: bx.bottom + 0.75,
    };
    let sums: Vec<f32> = groups
        .iter()
        .map(|(_, pts)| pts.iter().map(|&i| values[i]).sum())
        .collect();
    for ((gi, (_, pts)), outer) in groups.iter().enumerate().zip(tile(&sums, inner)) {
        let leaves: Vec<f32> = pts.iter().map(|&i| values[i]).collect();
        for (&i, t) in pts.iter().zip(tile(&leaves, outer)) {
            fill_rgb(content, ctx.accent(gi));
            content.rect(t.l + 0.125, t.bottom + 0.125, t.w() - 0.25, t.h() - 0.25);
            content.fill_nonzero();
            if ctx.c.show_values {
                let d = text.data_label;
                pen.text(
                    content,
                    t.l + 3.6,
                    t.bottom + 3.5 + pen.desc * d.size,
                    &general(values[i]),
                    d,
                );
            }
        }
    }
}

/// Rings out from a hole, each as wide as the hole's radius: the outer
/// categories, then their points, clockwise from 12 o'clock and coloured by
/// outer category, filling the box's height less 0.5pt. Values sit white at
/// the middle of their segment, turned along the radius and never upside
/// down (sunburst probes).
fn draw_sunburst(content: &mut Content, ctx: &Ctx, bx: Rect, groups: &[(&str, Vec<usize>)]) {
    let (pen, text, values) = (ctx.pen, ctx.text, ctx.values());
    let total: f32 = values.iter().sum();
    if total <= 0.0 {
        return;
    }
    let (cx, cy) = ((bx.l + bx.r) / 2.0, (bx.top + bx.bottom) / 2.0);
    let has_rings = groups.iter().any(|(name, _)| !name.is_empty());
    let rings = if has_rings { 2.0 } else { 1.0 };
    let r_out = bx.w().min(bx.h()) / 2.0 - 0.5;
    let band = r_out / (rings + 1.0);
    // Angles run clockwise from 12 o'clock.
    let point = |r: f32, a: f32| (cx + r * a.sin(), cy + r * a.cos());
    let sector = |content: &mut Content, r0: f32, r1: f32, a0: f32, a1: f32| {
        let n = arc_segment_count(a1 - a0);
        let (sx, sy) = point(r1, a0);
        content.move_to(sx, sy);
        emit_arc(content, cx, cy, r1, FRAC_PI_2 - a0, a1 - a0, n);
        emit_arc(content, cx, cy, r0, FRAC_PI_2 - a1, a0 - a1, n);
        content.close_path();
        content.fill_nonzero();
    };
    let leaf_r0 = r_out - band;
    let label = ChartText {
        color: WHITE,
        ..text.data_label
    };
    let mut a = 0.0f32;
    for (gi, (_, pts)) in groups.iter().enumerate() {
        let color = ctx.accent(gi);
        fill_rgb(content, color);
        let span = pts.iter().map(|&i| values[i]).sum::<f32>() / total * TAU;
        if has_rings {
            sector(content, band, leaf_r0, a, a + span);
        }
        let mut b = a;
        for &i in pts {
            let leaf = values[i] / total * TAU;
            fill_rgb(content, color);
            sector(content, leaf_r0, r_out, b, b + leaf);
            if ctx.c.show_values {
                let m = b + leaf / 2.0;
                let (px, py) = point(leaf_r0 + band / 2.0, m);
                // Along the radius, flipped where it would read right to left.
                let radial = FRAC_PI_2 - m;
                let angle = if radial.cos() < 0.0 {
                    radial + std::f32::consts::PI
                } else {
                    radial
                };
                pen.rotated(content, px, py, angle, &general(values[i]), label);
            }
            b += leaf;
        }
        a += span;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn value_axis_matches_word() {
        assert_eq!(value_axis(145.0), (20.0, 160.0));
        assert_eq!(value_axis(4.0), (0.5, 4.5));
        assert_eq!(value_axis(22.0), (5.0, 25.0));
        assert_eq!(value_axis(40.0), (5.0, 45.0));
        assert_eq!(value_axis(30.0), (5.0, 35.0));
    }

    #[test]
    fn treemap_rows_follow_word() {
        let r = Rect {
            l: 0.0,
            r: 180.0,
            top: 185.75,
            bottom: 0.0,
        };
        let t = tile(&[30.0, 20.0, 10.0], r);
        // 30 and 20 share the top row, 10 runs along the bottom.
        assert!((t[0].bottom - t[1].bottom).abs() < 0.01);
        assert!((t[2].w() - 180.0).abs() < 0.01);
    }
}
