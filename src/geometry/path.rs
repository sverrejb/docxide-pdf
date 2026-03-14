use super::formulas::GuideEnv;

#[derive(Clone, Debug)]
pub enum PathCommandDef {
    MoveTo {
        x: &'static str,
        y: &'static str,
    },
    LineTo {
        x: &'static str,
        y: &'static str,
    },
    ArcTo {
        wr: &'static str,
        hr: &'static str,
        st_ang: &'static str,
        sw_ang: &'static str,
    },
    CubicBezTo {
        x1: &'static str,
        y1: &'static str,
        x2: &'static str,
        y2: &'static str,
        x3: &'static str,
        y3: &'static str,
    },
    QuadBezTo {
        x1: &'static str,
        y1: &'static str,
        x2: &'static str,
        y2: &'static str,
    },
    Close,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum PathFill {
    Norm,
    None,
}

impl PathFill {
    pub fn is_filled(self) -> bool {
        self != PathFill::None
    }
}

#[derive(Clone, Debug)]
pub struct PathDef {
    pub commands: &'static [PathCommandDef],
    pub w: Option<i64>,
    pub h: Option<i64>,
    pub fill: PathFill,
    pub stroke: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedCommand {
    MoveTo(f64, f64),
    LineTo(f64, f64),
    CubicTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    Close,
}

enum Cmd<'a> {
    MoveTo {
        x: &'a str,
        y: &'a str,
    },
    LineTo {
        x: &'a str,
        y: &'a str,
    },
    ArcTo {
        wr: &'a str,
        hr: &'a str,
        st_ang: &'a str,
        sw_ang: &'a str,
    },
    CubicBezTo {
        x1: &'a str,
        y1: &'a str,
        x2: &'a str,
        y2: &'a str,
        x3: &'a str,
        y3: &'a str,
    },
    QuadBezTo {
        x1: &'a str,
        y1: &'a str,
        x2: &'a str,
        y2: &'a str,
    },
    Close,
}

impl PathCommandDef {
    fn as_cmd(&self) -> Cmd<'_> {
        match self {
            Self::MoveTo { x, y } => Cmd::MoveTo { x, y },
            Self::LineTo { x, y } => Cmd::LineTo { x, y },
            Self::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => Cmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            },
            Self::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            } => Cmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            },
            Self::QuadBezTo { x1, y1, x2, y2 } => Cmd::QuadBezTo { x1, y1, x2, y2 },
            Self::Close => Cmd::Close,
        }
    }
}

impl crate::model::CustomPathCommand {
    fn as_cmd(&self) -> Cmd<'_> {
        match self {
            Self::MoveTo { x, y } => Cmd::MoveTo { x, y },
            Self::LineTo { x, y } => Cmd::LineTo { x, y },
            Self::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => Cmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            },
            Self::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            } => Cmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            },
            Self::QuadBezTo { x1, y1, x2, y2 } => Cmd::QuadBezTo { x1, y1, x2, y2 },
            Self::Close => Cmd::Close,
        }
    }
}

/// Resolve a path definition into concrete commands.
///
/// Coordinates are transformed: OOXML y-down -> PDF y-up via `pdf_y = shape_h - ooxml_y`.
/// Guide values are in the `coord_w`/`coord_h` coordinate space (typically EMU-scaled),
/// which gets mapped to the `shape_w`/`shape_h` output space (points).
/// If the path defines its own w/h, those override the coordinate space dimensions.
pub fn resolve_path(
    def: &PathDef,
    env: &GuideEnv,
    shape_w: f64,
    shape_h: f64,
    coord_w: i64,
    coord_h: i64,
) -> Vec<ResolvedCommand> {
    let path_w = def.w.unwrap_or(coord_w) as f64;
    let path_h = def.h.unwrap_or(coord_h) as f64;
    resolve_commands(
        def.commands.iter().map(|c| c.as_cmd()),
        env,
        shape_w,
        shape_h,
        path_w,
        path_h,
    )
}

/// Resolve a custom geometry path (owned String fields) into concrete commands.
pub fn resolve_custom_path(
    path: &crate::model::CustomPathDef,
    env: &GuideEnv,
    shape_w: f64,
    shape_h: f64,
    coord_w: i64,
    coord_h: i64,
) -> Vec<ResolvedCommand> {
    let path_w = path.w.unwrap_or(coord_w) as f64;
    let path_h = path.h.unwrap_or(coord_h) as f64;
    resolve_commands(
        path.commands.iter().map(|c| c.as_cmd()),
        env,
        shape_w,
        shape_h,
        path_w,
        path_h,
    )
}

fn resolve_commands<'a>(
    commands: impl Iterator<Item = Cmd<'a>>,
    env: &GuideEnv,
    shape_w: f64,
    shape_h: f64,
    path_w: f64,
    path_h: f64,
) -> Vec<ResolvedCommand> {
    let scale_x = |v: f64| -> f64 {
        if path_w > 0.0 {
            v / path_w * shape_w
        } else {
            v
        }
    };
    let scale_y = |v: f64| -> f64 {
        let scaled = if path_h > 0.0 {
            v / path_h * shape_h
        } else {
            v
        };
        shape_h - scaled
    };
    let rx = |arg: &str| -> f64 { scale_x(env.resolve(arg) as f64) };
    let ry = |arg: &str| -> f64 { scale_y(env.resolve(arg) as f64) };

    let mut result = Vec::new();
    let mut cur_x = 0.0_f64;
    let mut cur_y = 0.0_f64;

    for cmd in commands {
        match cmd {
            Cmd::MoveTo { x, y } => {
                cur_x = rx(x);
                cur_y = ry(y);
                result.push(ResolvedCommand::MoveTo(cur_x, cur_y));
            }
            Cmd::LineTo { x, y } => {
                cur_x = rx(x);
                cur_y = ry(y);
                result.push(ResolvedCommand::LineTo(cur_x, cur_y));
            }
            Cmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            } => {
                let cmd = ResolvedCommand::CubicTo {
                    x1: rx(x1),
                    y1: ry(y1),
                    x2: rx(x2),
                    y2: ry(y2),
                    x: rx(x3),
                    y: ry(y3),
                };
                if let ResolvedCommand::CubicTo { x, y, .. } = &cmd {
                    cur_x = *x;
                    cur_y = *y;
                }
                result.push(cmd);
            }
            Cmd::QuadBezTo { x1, y1, x2, y2 } => {
                let (qx, qy) = (rx(x1), ry(y1));
                let (ex, ey) = (rx(x2), ry(y2));
                result.push(quad_to_cubic(cur_x, cur_y, qx, qy, ex, ey));
                cur_x = ex;
                cur_y = ey;
            }
            Cmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => {
                let wr_val = scale_x(env.resolve(wr) as f64).abs();
                let hr_raw = env.resolve(hr) as f64;
                let hr_val = if path_h > 0.0 {
                    (hr_raw / path_h * shape_h).abs()
                } else {
                    hr_raw.abs()
                };
                let st = env.resolve(st_ang) as f64;
                let sw = env.resolve(sw_ang) as f64;
                arc_to_cubics(&mut result, &mut cur_x, &mut cur_y, wr_val, hr_val, st, sw);
            }
            Cmd::Close => {
                result.push(ResolvedCommand::Close);
            }
        }
    }

    result
}

fn quad_to_cubic(cx: f64, cy: f64, qx: f64, qy: f64, ex: f64, ey: f64) -> ResolvedCommand {
    ResolvedCommand::CubicTo {
        x1: cx + 2.0 / 3.0 * (qx - cx),
        y1: cy + 2.0 / 3.0 * (qy - cy),
        x2: ex + 2.0 / 3.0 * (qx - ex),
        y2: ey + 2.0 / 3.0 * (qy - ey),
        x: ex,
        y: ey,
    }
}

/// Convert an OOXML arcTo into cubic bezier segments appended to `out`.
///
/// OOXML arcTo is relative to the current point:
/// - (wr, hr): x and y radii
/// - st_ang: start angle in 60000ths of a degree (OOXML clockwise in y-down)
/// - sw_ang: sweep angle in 60000ths of a degree (positive = clockwise in y-down)
///
/// Since we've already flipped y (shape_h - y), angles need adjustment:
/// In PDF y-up space, OOXML clockwise becomes counterclockwise, so we negate angles.
fn arc_to_cubics(
    out: &mut Vec<ResolvedCommand>,
    cur_x: &mut f64,
    cur_y: &mut f64,
    wr: f64,
    hr: f64,
    st_ang_60k: f64,
    sw_ang_60k: f64,
) {
    use std::f64::consts::{FRAC_PI_2, PI};

    if wr < 0.001 || hr < 0.001 || sw_ang_60k.abs() < 1.0 {
        return;
    }

    let ang_to_rad = |a: f64| a / (60000.0 * 180.0) * PI;

    let st_rad = ang_to_rad(-st_ang_60k);
    let sw_rad = ang_to_rad(-sw_ang_60k);

    let cx = *cur_x - wr * st_rad.cos();
    let cy = *cur_y - hr * st_rad.sin();

    let n_segs = ((sw_rad.abs() / FRAC_PI_2).ceil() as usize).max(1);
    let step = sw_rad / n_segs as f64;

    let mut angle = st_rad;
    for _ in 0..n_segs {
        let a0 = angle;
        let a1 = angle + step;

        let half = step / 2.0;
        let alpha = (4.0 / 3.0) * (1.0 - half.cos()) / half.sin();

        let x0 = cx + wr * a0.cos();
        let y0 = cy + hr * a0.sin();
        let x3 = cx + wr * a1.cos();
        let y3 = cy + hr * a1.sin();

        out.push(ResolvedCommand::CubicTo {
            x1: x0 - alpha * wr * a0.sin(),
            y1: y0 + alpha * hr * a0.cos(),
            x2: x3 + alpha * wr * a1.sin(),
            y2: y3 - alpha * hr * a1.cos(),
            x: x3,
            y: y3,
        });

        angle = a1;
    }

    let end_rad = st_rad + sw_rad;
    *cur_x = cx + wr * end_rad.cos();
    *cur_y = cy + hr * end_rad.sin();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_quad_to_cubic_conversion() {
        let env = GuideEnv::new(100, 100);
        static CMDS: &[PathCommandDef] = &[
            PathCommandDef::MoveTo { x: "0", y: "0" },
            PathCommandDef::QuadBezTo {
                x1: "50",
                y1: "100",
                x2: "100",
                y2: "0",
            },
        ];
        let def = PathDef {
            commands: CMDS,
            w: None,
            h: None,
            fill: PathFill::Norm,
            stroke: false,
        };
        let resolved = resolve_path(&def, &env, 100.0, 100.0, 100, 100);
        assert_eq!(resolved.len(), 2);
        // MoveTo(0, 100) — y-flipped from (0,0)
        assert!(
            matches!(resolved[0], ResolvedCommand::MoveTo(x, y) if (x - 0.0).abs() < 0.01 && (y - 100.0).abs() < 0.01)
        );
        // CubicTo — endpoint should be (100, 100) — y-flipped from (100, 0)
        if let ResolvedCommand::CubicTo { x, y, .. } = &resolved[1] {
            assert!((*x - 100.0).abs() < 0.01);
            assert!((*y - 100.0).abs() < 0.01);
        } else {
            panic!("Expected CubicTo");
        }
    }

    #[test]
    fn test_simple_rect_path() {
        let env = GuideEnv::new(200, 100);
        static CMDS: &[PathCommandDef] = &[
            PathCommandDef::MoveTo { x: "l", y: "t" },
            PathCommandDef::LineTo { x: "r", y: "t" },
            PathCommandDef::LineTo { x: "r", y: "b" },
            PathCommandDef::LineTo { x: "l", y: "b" },
            PathCommandDef::Close,
        ];
        let def = PathDef {
            commands: CMDS,
            w: None,
            h: None,
            fill: PathFill::Norm,
            stroke: false,
        };
        let resolved = resolve_path(&def, &env, 200.0, 100.0, 200, 100);
        assert_eq!(resolved.len(), 5);
        // (l,t) = (0,0) -> PDF (0, 100)
        assert!(matches!(&resolved[0], ResolvedCommand::MoveTo(x, y) if *x == 0.0 && *y == 100.0));
        // (r,t) = (200,0) -> PDF (200, 100)
        assert!(
            matches!(&resolved[1], ResolvedCommand::LineTo(x, y) if *x == 200.0 && *y == 100.0)
        );
        // (r,b) = (200,100) -> PDF (200, 0)
        assert!(matches!(&resolved[2], ResolvedCommand::LineTo(x, y) if *x == 200.0 && *y == 0.0));
        // (l,b) = (0,100) -> PDF (0, 0)
        assert!(matches!(&resolved[3], ResolvedCommand::LineTo(x, y) if *x == 0.0 && *y == 0.0));
        assert!(matches!(&resolved[4], ResolvedCommand::Close));
    }

    #[test]
    fn test_path_coordinate_scaling() {
        // Path defines its own coordinate system (w=2, h=2), shape is 100x50
        let env = GuideEnv::new(2, 2);
        static CMDS: &[PathCommandDef] = &[
            PathCommandDef::MoveTo { x: "0", y: "0" },
            PathCommandDef::LineTo { x: "2", y: "0" },
            PathCommandDef::LineTo { x: "1", y: "2" },
            PathCommandDef::Close,
        ];
        let def = PathDef {
            commands: CMDS,
            w: Some(2),
            h: Some(2),
            fill: PathFill::Norm,
            stroke: false,
        };
        let resolved = resolve_path(&def, &env, 100.0, 50.0, 2, 2);
        // (0,0) -> scaled (0,0) -> y-flip (0, 50)
        assert!(matches!(&resolved[0], ResolvedCommand::MoveTo(x, y) if *x == 0.0 && *y == 50.0));
        // (2,0) -> scaled (100,0) -> y-flip (100, 50)
        assert!(matches!(&resolved[1], ResolvedCommand::LineTo(x, y) if *x == 100.0 && *y == 50.0));
        // (1,2) -> scaled (50,50) -> y-flip (50, 0)
        assert!(matches!(&resolved[2], ResolvedCommand::LineTo(x, y) if *x == 50.0 && *y == 0.0));
    }

    #[test]
    fn test_arc_90_degree() {
        // Test a 90-degree arc (quarter circle) top-right corner of a roundRect
        let mut env = GuideEnv::new(1_000_000, 1_000_000);
        env.set_adjustments(&[("r1", 100_000)], &[]);
        env.evaluate_guides(&[]);

        // Arc starting from top of corner: stAng=3cd4 (270 deg = top), swAng=cd4 (90 deg CW)
        // This should sweep from top to right of the corner
        static CMDS: &[PathCommandDef] = &[
            // Start at top of arc: the point where the corner radius begins on the top edge
            PathCommandDef::MoveTo {
                x: "900000",
                y: "0",
            },
            PathCommandDef::ArcTo {
                wr: "r1",
                hr: "r1",
                st_ang: "3cd4",
                sw_ang: "cd4",
            },
        ];
        let def = PathDef {
            commands: CMDS,
            w: None,
            h: None,
            fill: PathFill::Norm,
            stroke: false,
        };
        let resolved = resolve_path(&def, &env, 100.0, 100.0, 1_000_000, 1_000_000);
        // Should have MoveTo + at least one CubicTo
        assert!(resolved.len() >= 2);
        assert!(matches!(resolved[0], ResolvedCommand::MoveTo(..)));
        assert!(matches!(resolved[1], ResolvedCommand::CubicTo { .. }));
    }
}
