pub mod definitions;
pub mod formulas;
pub mod path;

pub use definitions::PresetDef;
pub use formulas::{FormulaOp, GuideDef, GuideEnv};
pub use path::{PathCommandDef, PathDef, PathFill, ResolvedCommand, resolve_custom_path};

pub struct EvaluatedPath {
    pub commands: Vec<ResolvedCommand>,
    pub fill: PathFill,
    pub stroke: bool,
}

pub struct EvaluatedShape {
    pub paths: Vec<EvaluatedPath>,
    /// Text rectangle in PDF coordinates (x, y_bottom, width, height).
    pub text_rect: Option<(f64, f64, f64, f64)>,
}

/// Evaluate a preset shape geometry, returning resolved paths in PDF coordinate space.
///
/// `w` and `h` are the shape dimensions in points.
/// Coordinates are y-flipped (OOXML y-down -> PDF y-up).
pub fn evaluate_preset(
    name: &str,
    w: f64,
    h: f64,
    adj_overrides: &[(String, i64)],
) -> Option<EvaluatedShape> {
    let def = definitions::lookup(name)?;
    Some(evaluate_def(def, w, h, adj_overrides))
}

/// Evaluate a preset definition with given dimensions and adjustment overrides.
pub fn evaluate_def(
    def: &PresetDef,
    w: f64,
    h: f64,
    adj_overrides: &[(String, i64)],
) -> EvaluatedShape {
    // OOXML coordinates are integers; scale points to a large integer space
    // for guide computation, then scale back to f64 points for the resolved paths.
    //
    // We use EMU-like units: 1 point = 12700 EMU. This preserves precision
    // for the integer formula evaluator while keeping values within i64 range.
    let scale = 12700.0;
    let wi = (w * scale) as i64;
    let hi = (h * scale) as i64;

    let mut env = GuideEnv::new(wi, hi);
    env.set_adjustments(def.adjust_defaults, adj_overrides);
    env.evaluate_guides(def.guides);

    let paths = def
        .paths
        .iter()
        .map(|p| EvaluatedPath {
            commands: path::resolve_path(p, &env, w, h, wi, hi),
            fill: p.fill,
            stroke: p.stroke,
        })
        .collect();

    let text_rect = def.text_rect.as_ref().map(|tr| {
        let l = env.resolve(tr.l) as f64 / scale;
        let t = env.resolve(tr.t) as f64 / scale;
        let r = env.resolve(tr.r) as f64 / scale;
        let b = env.resolve(tr.b) as f64 / scale;
        // Convert to PDF coords: (x, y_bottom, width, height)
        (l, h - b, r - l, b - t)
    });

    EvaluatedShape { paths, text_rect }
}

/// Evaluate a custom geometry (a:custGeom), returning resolved paths in PDF coordinate space.
pub fn evaluate_custom(
    custom: &crate::model::CustomGeometry,
    w: f64,
    h: f64,
    adj_overrides: &[(String, i64)],
) -> EvaluatedShape {
    let scale = 12700.0;
    let wi = (w * scale) as i64;
    let hi = (h * scale) as i64;

    let mut env = GuideEnv::new(wi, hi);
    // Set custom adjustment defaults, then apply overrides
    for (name, val) in &custom.adjust_defaults {
        env.set_adjustments(&[], &[(name.clone(), *val)]);
    }
    if !adj_overrides.is_empty() {
        env.set_adjustments(&[], adj_overrides);
    }
    env.evaluate_custom_guides(&custom.guides);

    let paths = custom
        .paths
        .iter()
        .map(|p| EvaluatedPath {
            commands: resolve_custom_path(p, &env, w, h, wi, hi),
            fill: p.fill,
            stroke: p.stroke,
        })
        .collect();

    EvaluatedShape { paths, text_rect: None }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_evaluate_preset_unknown() {
        assert!(evaluate_preset("nonexistent_shape", 100.0, 50.0, &[]).is_none());
    }

    fn approx_eq(a: f64, b: f64, tol: f64) -> bool {
        (a - b).abs() < tol
    }

    /// Verify geometry engine rect matches the hand-coded `content.rect(x, y, w, h)`.
    ///
    /// PDF rect(0, 0, w, h) covers corners (0,0)-(w,0)-(w,h)-(0,h).
    /// Geometry engine produces MoveTo(0,h) LineTo(w,h) LineTo(w,0) LineTo(0,0) Close.
    /// Both describe the same rectangle.
    #[test]
    fn test_rect_matches_handcoded() {
        let w = 100.0;
        let h = 50.0;
        let shape = evaluate_preset("rect", w, h, &[]).unwrap();
        let cmds = &shape.paths[0].commands;

        assert_eq!(cmds.len(), 5);
        // Top-left in PDF coords (OOXML origin 0,0 -> PDF 0,h)
        assert!(matches!(&cmds[0], ResolvedCommand::MoveTo(x, y) if approx_eq(*x, 0.0, 0.01) && approx_eq(*y, h, 0.01)));
        // Top-right
        assert!(matches!(&cmds[1], ResolvedCommand::LineTo(x, y) if approx_eq(*x, w, 0.01) && approx_eq(*y, h, 0.01)));
        // Bottom-right
        assert!(matches!(&cmds[2], ResolvedCommand::LineTo(x, y) if approx_eq(*x, w, 0.01) && approx_eq(*y, 0.0, 0.01)));
        // Bottom-left
        assert!(matches!(&cmds[3], ResolvedCommand::LineTo(x, y) if approx_eq(*x, 0.0, 0.01) && approx_eq(*y, 0.0, 0.01)));
        assert!(matches!(&cmds[4], ResolvedCommand::Close));
    }

    /// Verify geometry engine ellipse approximates the same curve as the hand-coded
    /// 4-cubic-bezier approach using K=0.5522847498.
    ///
    /// The hand-coded ellipse starts at (cx+rx, cy) and goes counterclockwise.
    /// The geometry engine starts at (l, vc) = (0, h/2) and goes clockwise in OOXML
    /// (counterclockwise after y-flip).
    ///
    /// We verify the ellipse passes through the 4 cardinal points.
    #[test]
    fn test_ellipse_matches_handcoded() {
        let w = 100.0;
        let h = 80.0;
        let shape = evaluate_preset("ellipse", w, h, &[]).unwrap();
        let cmds = &shape.paths[0].commands;

        // Starts at left center: (0, h/2) in PDF coords
        assert!(matches!(&cmds[0], ResolvedCommand::MoveTo(x, y) if approx_eq(*x, 0.0, 0.1) && approx_eq(*y, h / 2.0, 0.1)));

        // Collect all endpoints from the cubic segments
        let endpoints: Vec<(f64, f64)> = cmds.iter().filter_map(|c| match c {
            ResolvedCommand::CubicTo { x, y, .. } => Some((*x, *y)),
            _ => None,
        }).collect();

        // 4 arcs of 90 degrees each produce 4 cubic segments
        assert_eq!(endpoints.len(), 4, "Expected 4 cubic bezier segments for 4 quarter-arcs");

        // After each 90° arc, should hit: top center, right center, bottom center, left center
        let cardinal_points = [
            (w / 2.0, h),       // top center
            (w, h / 2.0),       // right center
            (w / 2.0, 0.0),     // bottom center
            (0.0, h / 2.0),     // left center (back to start)
        ];

        for (i, (ex, ey)) in cardinal_points.iter().enumerate() {
            let (gx, gy) = endpoints[i];
            assert!(
                approx_eq(gx, *ex, 0.5) && approx_eq(gy, *ey, 0.5),
                "Arc {i} endpoint ({gx:.2}, {gy:.2}) != expected ({ex:.2}, {ey:.2})"
            );
        }

        // Verify the cubic control points approximate the K=0.5522847498 formula
        // For the first arc (left center -> top center), the hand-coded version would use:
        //   K = 0.5522847498, rx = w/2, ry = h/2
        // Starting at (0, h/2), ending at (w/2, h):
        //   cp1 = (0, h/2 + K*ry) = (0, 40 + 0.5523*40) = (0, 62.09)
        //   cp2 = (w/2 - K*rx, h) = (50 - 0.5523*50, 80) = (22.39, 80)
        let k = 0.5522847498_f64;
        let rx = w / 2.0;
        let ry = h / 2.0;
        if let ResolvedCommand::CubicTo { x1, y1, x2, y2, .. } = &cmds[1] {
            let expected_cp1x = 0.0;
            let expected_cp1y = h / 2.0 + k * ry;
            let expected_cp2x = w / 2.0 - k * rx;
            let expected_cp2y = h;
            assert!(
                approx_eq(*x1, expected_cp1x, 0.5) && approx_eq(*y1, expected_cp1y, 0.5),
                "CP1 ({x1:.2}, {y1:.2}) != expected ({expected_cp1x:.2}, {expected_cp1y:.2})"
            );
            assert!(
                approx_eq(*x2, expected_cp2x, 0.5) && approx_eq(*y2, expected_cp2y, 0.5),
                "CP2 ({x2:.2}, {y2:.2}) != expected ({expected_cp2x:.2}, {expected_cp2y:.2})"
            );
        } else {
            panic!("Expected CubicTo as first arc segment");
        }
    }

    /// Verify geometry engine notchedRightArrow matches the hand-coded version exactly.
    ///
    /// Hand-coded (w=200, h=100, at origin):
    ///   ss=100, arrow_dx=50, arrow_start=150, shaft_inset=25, notch_depth=25
    ///   Vertices: (0,75) (150,75) (150,100) (200,50) (150,0) (150,25) (0,25) (25,50)
    #[test]
    fn test_notched_right_arrow_matches_handcoded() {
        let w = 200.0;
        let h = 100.0;
        let shape = evaluate_preset("notchedRightArrow", w, h, &[]).unwrap();
        let cmds = &shape.paths[0].commands;

        // Hand-coded vertices in PDF coords (y=0 is bottom)
        let expected = [
            (0.0, 75.0),    // (l, y1): shaft top-left
            (150.0, 75.0),  // (x2, y1): shaft top-right
            (150.0, 100.0), // (x2, t): arrowhead top
            (200.0, 50.0),  // (r, vc): arrowhead tip
            (150.0, 0.0),   // (x2, b): arrowhead bottom
            (150.0, 25.0),  // (x2, y2): shaft bottom-right
            (0.0, 25.0),    // (l, y2): shaft bottom-left
            (25.0, 50.0),   // (x1, vc): notch point
        ];

        assert_eq!(cmds.len(), 9, "Expected 9 commands: MoveTo + 7 LineTo + Close");
        assert!(matches!(&cmds[0], ResolvedCommand::MoveTo(x, y) if approx_eq(*x, expected[0].0, 0.01) && approx_eq(*y, expected[0].1, 0.01)),
            "MoveTo mismatch: got {:?}, expected {:?}", cmds[0], expected[0]);
        for i in 1..8 {
            assert!(matches!(&cmds[i], ResolvedCommand::LineTo(x, y) if approx_eq(*x, expected[i].0, 0.01) && approx_eq(*y, expected[i].1, 0.01)),
                "LineTo {i} mismatch: got {:?}, expected {:?}", cmds[i], expected[i]);
        }
        assert!(matches!(&cmds[8], ResolvedCommand::Close));
    }
}
