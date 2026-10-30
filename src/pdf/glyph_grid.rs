//! Where Word's online export draws the glyphs of a line (probe 1 of
//! `current_focus/fundamentals.md`, kit `tests/output/fundamentals/glyph_grid`).
//!
//! Layout keeps the exact font advances; only drawing moves. Word places each
//! glyph on a 288 dpi device grid (1 px = 0.25pt) following the exact advances
//! accumulated over the line, then its PDF writer rounds every advance to
//! whole thousandths of an em with a bias that loses part of a unit per glyph,
//! restarting at each run's text matrix. Together they leave a run's last
//! words up to ~0.3pt left of their exact places.

/// One glyph of a sequence of glyphs set edge to edge on a line.
pub(super) struct GridGlyph {
    /// Exact advance in points.
    pub advance: f32,
    pub font_size: f32,
    /// Word's writer starts a new text matrix here (a new run).
    pub run_start: bool,
}

const PX_PER_PT: f64 = 4.0;
/// A device chunk restarts the pixel position at the rounded exact one; the
/// probe pins the length at 1024 px (1000 or 1040 quadruple the misses).
const CHUNK_PX: f64 = 1023.5;
/// Word's text matrices sit 0.025pt right of the 0.25pt grid (36.025 for a
/// 36pt margin, in every reference checked).
const PHASE_PT: f64 = 0.025;

/// A glyph's own device advance rounds up from 0.46 px, not 0.5: over the
/// fixture references 236/512 px goes up and 235/512 down in every face and
/// size (bar a few hinted glyphs), which halves the wrong advances.
const ROUND_UP_FROM: f64 = 0.46;

/// Device advances (px) of glyphs whose exact origins are `ideal` (absolute
/// px) and exact advances `exact` (px): each glyph keeps its device advance
/// unless the glyph after next would then land 2 px or more from its rounded
/// exact place, measured from the chunk start; then it moves 1 px toward it.
fn device_steps(ideal: &[f64], exact: &[f64]) -> Vec<i64> {
    let n = exact.len();
    let ideal_end = ideal[n - 1] + exact[n - 1];
    let base: Vec<i64> = exact
        .iter()
        .map(|w| w.floor() as i64 + i64::from(w.fract() >= ROUND_UP_FROM))
        .collect();
    let mut starts = vec![false; n];
    starts[0] = true;
    let mut chunk0 = 0;
    for j in 1..n {
        if ideal[j] + exact[j] - ideal[chunk0] > CHUNK_PX {
            starts[j] = true;
            chunk0 = j;
        }
    }
    let (mut pos, mut p0, mut i0) = (0i64, 0i64, 0.0f64);
    let mut steps = Vec::with_capacity(n.saturating_sub(1));
    for k in 0..n.saturating_sub(1) {
        if starts[k] {
            pos = ideal[k].round() as i64;
            p0 = pos;
            i0 = ideal[k];
        }
        let adv = if starts[k + 1] {
            ideal[k + 1].round() as i64 - pos
        } else {
            let next2 = if k + 2 < n { ideal[k + 2] } else { ideal_end };
            let placed = (pos - p0) + base[k] + base[k + 1];
            let wanted = (next2 - i0).round() as i64;
            if placed - wanted >= 2 {
                base[k] - 1
            } else if wanted - placed >= 2 {
                base[k] + 1
            } else {
                base[k]
            }
        };
        steps.push(adv);
        pos += adv;
    }
    steps
}

/// The drawn x (points) of each glyph of a sequence whose exact origin is `x0`.
pub(super) fn drawn_positions(x0: f32, glyphs: &[GridGlyph]) -> Vec<f32> {
    if glyphs.is_empty() {
        return Vec::new();
    }
    let exact: Vec<f64> = glyphs
        .iter()
        .map(|g| g.advance as f64 * PX_PER_PT)
        .collect();
    let mut ideal = Vec::with_capacity(exact.len());
    let mut x = x0 as f64 * PX_PER_PT;
    for w in &exact {
        ideal.push(x);
        x += w;
    }
    let steps = device_steps(&ideal, &exact);
    let mut device = ideal[0].round() as i64;
    let mut drawn = Vec::with_capacity(glyphs.len());
    let mut pen = 0.0f64;
    for (k, g) in glyphs.iter().enumerate() {
        if k == 0 || g.run_start {
            pen = device as f64 / PX_PER_PT + PHASE_PT;
        } else {
            // The writer's TJ moves are whole units; for the advance before
            // this glyph it writes W − adj with adj rounded half away from
            // zero, one more when the device advance exceeds W.
            let prev = &glyphs[k - 1];
            let unit = prev.font_size as f64 / 1000.0;
            let w = (prev.advance as f64 / unit).round();
            let d = steps[k - 1] as f64 / PX_PER_PT / unit;
            let gap = w - d;
            let adj = if gap.abs() < 0.5 {
                0.0
            } else if gap > 0.0 {
                gap.round()
            } else {
                gap.round() + 1.0
            };
            pen += (w - adj) * unit;
        }
        drawn.push(pen as f32);
        if k + 1 < glyphs.len() {
            device += steps[k];
        }
    }
    drawn
}

#[cfg(test)]
mod tests {
    use super::*;

    fn run(advances: &[f32], size: f32) -> Vec<GridGlyph> {
        advances
            .iter()
            .map(|&advance| GridGlyph {
                advance,
                font_size: size,
                run_start: false,
            })
            .collect()
    }

    #[test]
    fn steps_stay_within_a_pixel_of_the_exact_positions() {
        // Times New Roman 12pt "e" (443.848 thousandths): 21.3 px each.
        let w = 443.848 * 12.0 / 1000.0;
        let g = run(&[w; 70], 12.0);
        let exact: Vec<f64> = g.iter().map(|g| g.advance as f64 * 4.0).collect();
        let ideal: Vec<f64> = (0..70).map(|k| 288.0 + k as f64 * exact[0]).collect();
        let steps = device_steps(&ideal, &exact);
        let mut pos = 288i64;
        for (k, s) in steps.iter().enumerate() {
            assert!((21..=22).contains(s));
            pos += s;
            assert!((pos as f64 - ideal[k + 1]).abs() < 1.5);
        }
    }

    #[test]
    fn writer_loses_a_fraction_of_a_unit_per_glyph() {
        // 21 px = 437.5 units against W = 444: the writer moves 7 back and
        // draws 437, half a unit short of the device advance.
        let w = 443.848 * 12.0 / 1000.0;
        let drawn = drawn_positions(72.0, &run(&[w; 3], 12.0));
        assert!((drawn[0] - 72.025).abs() < 1e-4);
        assert!((drawn[1] - (72.025 + 0.437 * 12.0)).abs() < 1e-4);
    }

    #[test]
    fn a_run_start_returns_to_the_device_grid() {
        let w = 443.848 * 12.0 / 1000.0;
        let mut g = run(&[w; 40], 12.0);
        g[20].run_start = true;
        let drawn = drawn_positions(72.0, &g);
        let on_grid = (drawn[20] as f64 - PHASE_PT) * 4.0;
        assert!((on_grid - on_grid.round()).abs() < 1e-3);
        assert!(drawn[19] < drawn[20]);
    }
}
