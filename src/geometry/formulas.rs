use std::collections::HashMap;

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum FormulaOp {
    Val,
    MulDiv,
    AddSub,
    AddDiv,
    IfElse,
    Abs,
    Sqrt,
    Min,
    Max,
    Pin,
    Sin,
    Cos,
    Tan,
    Atan2,
    CosAtan2,
    SinAtan2,
    Mod,
}

impl FormulaOp {
    pub fn from_str(s: &str) -> Option<Self> {
        Some(match s {
            "val" => Self::Val,
            "*/" => Self::MulDiv,
            "+-" => Self::AddSub,
            "+/" => Self::AddDiv,
            "?:" => Self::IfElse,
            "abs" => Self::Abs,
            "sqrt" => Self::Sqrt,
            "min" => Self::Min,
            "max" => Self::Max,
            "pin" => Self::Pin,
            "sin" => Self::Sin,
            "cos" => Self::Cos,
            "tan" => Self::Tan,
            "at2" => Self::Atan2,
            "cat2" => Self::CosAtan2,
            "sat2" => Self::SinAtan2,
            "mod" => Self::Mod,
            _ => return None,
        })
    }
}

#[derive(Clone, Debug)]
pub struct GuideDef {
    pub name: &'static str,
    pub op: FormulaOp,
    pub x: &'static str,
    pub y: &'static str,
    pub z: &'static str,
}

const ANG_UNIT: f64 = 60000.0; // 60000ths of a degree per degree

pub struct GuideEnv {
    values: HashMap<String, i64>,
}

impl GuideEnv {
    pub fn new(w: i64, h: i64) -> Self {
        let mut values = HashMap::new();
        let ss = w.min(h);
        let ls = w.max(h);

        values.insert("w".into(), w);
        values.insert("h".into(), h);
        values.insert("l".into(), 0);
        values.insert("t".into(), 0);
        values.insert("r".into(), w);
        values.insert("b".into(), h);
        values.insert("ss".into(), ss);
        values.insert("ls".into(), ls);
        values.insert("hc".into(), w / 2);
        values.insert("vc".into(), h / 2);

        // wd{n} = w / n
        for d in [2, 3, 4, 5, 6, 8, 10, 12] {
            values.insert(format!("wd{d}"), w / d);
        }
        // hd{n} = h / n
        for d in [2, 3, 4, 5, 6, 8, 10] {
            values.insert(format!("hd{d}"), h / d);
        }
        // ssd{n} = ss / n
        for d in [2, 4, 6, 8, 16, 32] {
            values.insert(format!("ssd{d}"), ss / d);
        }

        // Angle constants (in 60000ths of a degree)
        values.insert("cd2".into(), 10_800_000);  // 180 deg
        values.insert("cd4".into(), 5_400_000);   // 90 deg
        values.insert("cd8".into(), 2_700_000);   // 45 deg
        values.insert("3cd4".into(), 16_200_000); // 270 deg
        values.insert("3cd8".into(), 8_100_000);  // 135 deg
        values.insert("5cd8".into(), 13_500_000); // 225 deg
        values.insert("7cd8".into(), 18_900_000); // 315 deg

        Self { values }
    }

    pub fn set_adjustments(
        &mut self,
        defaults: &[(&str, i64)],
        overrides: &[(String, i64)],
    ) {
        for &(name, val) in defaults {
            self.values.insert(name.to_string(), val);
        }
        for (name, val) in overrides {
            self.values.insert(name.clone(), *val);
        }
    }

    pub fn resolve(&self, arg: &str) -> i64 {
        if arg.is_empty() {
            return 0;
        }
        if let Some(&val) = self.values.get(arg) {
            return val;
        }
        arg.parse::<i64>().unwrap_or(0)
    }

    pub fn evaluate_guides(&mut self, guides: &[GuideDef]) {
        for g in guides {
            let val = self.eval_op(g.op, g.x, g.y, g.z);
            self.values.insert(g.name.to_string(), val);
        }
    }

    pub fn evaluate_custom_guides(&mut self, guides: &[crate::model::CustomGuideDef]) {
        for g in guides {
            let val = self.eval_op(g.op, &g.x, &g.y, &g.z);
            self.values.insert(g.name.clone(), val);
        }
    }

    fn eval_op(&self, op: FormulaOp, xa: &str, ya: &str, za: &str) -> i64 {
        let x = || self.resolve(xa);
        let y = || self.resolve(ya);
        let z = || self.resolve(za);

        match op {
            FormulaOp::Val => x(),
            FormulaOp::MulDiv => {
                let denom = z();
                if denom == 0 { 0 } else { (x() as i128 * y() as i128 / denom as i128) as i64 }
            }
            FormulaOp::AddSub => x() + y() - z(),
            FormulaOp::AddDiv => {
                let denom = z();
                if denom == 0 { 0 } else { (x() + y()) / denom }
            }
            FormulaOp::IfElse => {
                if x() > 0 { y() } else { z() }
            }
            FormulaOp::Abs => x().abs(),
            FormulaOp::Sqrt => {
                let v = x();
                if v <= 0 { 0 } else { (v as f64).sqrt() as i64 }
            }
            FormulaOp::Min => x().min(y()),
            FormulaOp::Max => x().max(y()),
            FormulaOp::Pin => {
                let lo = x();
                let val = y();
                let hi = z();
                val.clamp(lo, hi)
            }
            FormulaOp::Sin => {
                let mag = x() as f64;
                let ang = y() as f64 / (ANG_UNIT * 180.0) * std::f64::consts::PI;
                (mag * ang.sin()) as i64
            }
            FormulaOp::Cos => {
                let mag = x() as f64;
                let ang = y() as f64 / (ANG_UNIT * 180.0) * std::f64::consts::PI;
                (mag * ang.cos()) as i64
            }
            FormulaOp::Tan => {
                let mag = x() as f64;
                let ang = y() as f64 / (ANG_UNIT * 180.0) * std::f64::consts::PI;
                (mag * ang.tan()) as i64
            }
            FormulaOp::Atan2 => {
                let xv = x() as f64;
                let yv = y() as f64;
                let rad = yv.atan2(xv);
                (rad / std::f64::consts::PI * 180.0 * ANG_UNIT) as i64
            }
            FormulaOp::CosAtan2 => {
                let mag = x() as f64;
                let yv = y() as f64;
                let zv = z() as f64;
                let ang = zv.atan2(yv);
                (mag * ang.cos()) as i64
            }
            FormulaOp::SinAtan2 => {
                let mag = x() as f64;
                let yv = y() as f64;
                let zv = z() as f64;
                let ang = zv.atan2(yv);
                (mag * ang.sin()) as i64
            }
            FormulaOp::Mod => {
                let xv = x() as f64;
                let yv = y() as f64;
                let zv = z() as f64;
                (xv * xv + yv * yv + zv * zv).sqrt() as i64
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn env_1m() -> GuideEnv {
        GuideEnv::new(1_000_000, 1_000_000)
    }

    #[test]
    fn test_val() {
        let env = env_1m();
        assert_eq!(env.resolve("w"), 1_000_000);
        assert_eq!(env.resolve("h"), 1_000_000);
        assert_eq!(env.resolve("l"), 0);
        assert_eq!(env.resolve("r"), 1_000_000);
        assert_eq!(env.resolve("ss"), 1_000_000);
        assert_eq!(env.resolve("hc"), 500_000);
        assert_eq!(env.resolve("vc"), 500_000);
    }

    #[test]
    fn test_builtins_non_square() {
        let env = GuideEnv::new(2_000_000, 1_000_000);
        assert_eq!(env.resolve("ss"), 1_000_000);
        assert_eq!(env.resolve("ls"), 2_000_000);
        assert_eq!(env.resolve("hc"), 1_000_000);
        assert_eq!(env.resolve("vc"), 500_000);
        assert_eq!(env.resolve("wd2"), 1_000_000);
        assert_eq!(env.resolve("hd2"), 500_000);
    }

    #[test]
    fn test_angle_constants() {
        let env = env_1m();
        assert_eq!(env.resolve("cd2"), 10_800_000);
        assert_eq!(env.resolve("cd4"), 5_400_000);
        assert_eq!(env.resolve("cd8"), 2_700_000);
        assert_eq!(env.resolve("3cd4"), 16_200_000);
    }

    #[test]
    fn test_mul_div() {
        let mut env = env_1m();
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::MulDiv,
            x: "w", y: "50000", z: "100000",
        }]);
        assert_eq!(env.resolve("g1"), 500_000);
    }

    #[test]
    fn test_mul_div_zero_denom() {
        let mut env = env_1m();
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::MulDiv,
            x: "w", y: "50000", z: "0",
        }]);
        assert_eq!(env.resolve("g1"), 0);
    }

    #[test]
    fn test_add_sub() {
        let mut env = env_1m();
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::AddSub,
            x: "r", y: "0", z: "100000",
        }]);
        assert_eq!(env.resolve("g1"), 900_000);
    }

    #[test]
    fn test_add_div() {
        let mut env = env_1m();
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::AddDiv,
            x: "w", y: "h", z: "2",
        }]);
        assert_eq!(env.resolve("g1"), 1_000_000);
    }

    #[test]
    fn test_if_else() {
        let mut env = env_1m();
        env.evaluate_guides(&[
            GuideDef { name: "pos", op: FormulaOp::IfElse, x: "w", y: "100", z: "200" },
            GuideDef { name: "neg", op: FormulaOp::IfElse, x: "l", y: "100", z: "200" },
        ]);
        assert_eq!(env.resolve("pos"), 100);
        assert_eq!(env.resolve("neg"), 200);
    }

    #[test]
    fn test_abs() {
        let mut env = env_1m();
        env.values.insert("neg".into(), -500);
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::Abs, x: "neg", y: "", z: "",
        }]);
        assert_eq!(env.resolve("g1"), 500);
    }

    #[test]
    fn test_sqrt() {
        let mut env = env_1m();
        env.values.insert("v".into(), 1_000_000);
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::Sqrt, x: "v", y: "", z: "",
        }]);
        assert_eq!(env.resolve("g1"), 1000);
    }

    #[test]
    fn test_min_max() {
        let mut env = env_1m();
        env.evaluate_guides(&[
            GuideDef { name: "lo", op: FormulaOp::Min, x: "300", y: "500", z: "" },
            GuideDef { name: "hi", op: FormulaOp::Max, x: "300", y: "500", z: "" },
        ]);
        assert_eq!(env.resolve("lo"), 300);
        assert_eq!(env.resolve("hi"), 500);
    }

    #[test]
    fn test_pin() {
        let mut env = env_1m();
        env.evaluate_guides(&[
            GuideDef { name: "clamped", op: FormulaOp::Pin, x: "100", y: "50", z: "200" },
            GuideDef { name: "above", op: FormulaOp::Pin, x: "100", y: "300", z: "200" },
        ]);
        assert_eq!(env.resolve("clamped"), 100); // 50 clamped to [100,200]
        assert_eq!(env.resolve("above"), 200);   // 300 clamped to [100,200]
    }

    #[test]
    fn test_sin_cos() {
        let mut env = env_1m();
        // sin(1000000, 5400000) = 1000000 * sin(90 deg) = 1000000
        env.evaluate_guides(&[
            GuideDef { name: "s", op: FormulaOp::Sin, x: "1000000", y: "cd4", z: "" },
            GuideDef { name: "c", op: FormulaOp::Cos, x: "1000000", y: "cd4", z: "" },
        ]);
        assert_eq!(env.resolve("s"), 1_000_000);
        assert!(env.resolve("c").abs() < 100); // cos(90) ~ 0
    }

    #[test]
    fn test_atan2() {
        let mut env = env_1m();
        // at2(1000000, 1000000) = atan2(1000000, 1000000) = 45 deg = 2700000
        env.evaluate_guides(&[GuideDef {
            name: "a", op: FormulaOp::Atan2, x: "1000000", y: "1000000", z: "",
        }]);
        assert!((env.resolve("a") - 2_700_000).abs() < 100);
    }

    #[test]
    fn test_mod_vector_magnitude() {
        let mut env = env_1m();
        // mod(3, 4, 0) = sqrt(9+16+0) = 5
        env.evaluate_guides(&[GuideDef {
            name: "m", op: FormulaOp::Mod, x: "3", y: "4", z: "0",
        }]);
        assert_eq!(env.resolve("m"), 5);
    }

    #[test]
    fn test_cat2_sat2() {
        let mut env = env_1m();
        // cat2(1000, 1000000, 0) = 1000 * cos(atan2(0, 1000000)) = 1000 * cos(0) = 1000
        // sat2(1000, 1000000, 0) = 1000 * sin(atan2(0, 1000000)) = 1000 * sin(0) = 0
        env.evaluate_guides(&[
            GuideDef { name: "ct", op: FormulaOp::CosAtan2, x: "1000", y: "1000000", z: "0" },
            GuideDef { name: "st", op: FormulaOp::SinAtan2, x: "1000", y: "1000000", z: "0" },
        ]);
        assert_eq!(env.resolve("ct"), 1000);
        assert!(env.resolve("st").abs() < 2);
    }

    #[test]
    fn test_guide_chaining() {
        let mut env = GuideEnv::new(1_000_000, 500_000);
        env.set_adjustments(&[("adj", 25000)], &[]);
        env.evaluate_guides(&[
            GuideDef { name: "a", op: FormulaOp::Pin, x: "0", y: "adj", z: "50000" },
            GuideDef { name: "x1", op: FormulaOp::MulDiv, x: "ss", y: "a", z: "100000" },
            GuideDef { name: "x2", op: FormulaOp::AddSub, x: "r", y: "0", z: "x1" },
        ]);
        // ss = min(1M, 500K) = 500000
        // a = pin(0, 25000, 50000) = 25000
        // x1 = 500000 * 25000 / 100000 = 125000
        // x2 = 1000000 - 125000 = 875000
        assert_eq!(env.resolve("a"), 25000);
        assert_eq!(env.resolve("x1"), 125000);
        assert_eq!(env.resolve("x2"), 875000);
    }

    #[test]
    fn test_adjustment_override() {
        let mut env = env_1m();
        env.set_adjustments(&[("adj", 50000)], &[("adj".to_string(), 75000)]);
        assert_eq!(env.resolve("adj"), 75000);
    }

    #[test]
    fn test_tan() {
        let mut env = env_1m();
        // tan(1000, cd8) = 1000 * tan(45 deg) = 1000
        env.evaluate_guides(&[GuideDef {
            name: "t", op: FormulaOp::Tan, x: "1000", y: "cd8", z: "",
        }]);
        assert!((env.resolve("t") - 1000).abs() < 2);
    }

    #[test]
    fn test_sqrt_negative() {
        let mut env = env_1m();
        env.values.insert("neg".into(), -100);
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::Sqrt, x: "neg", y: "", z: "",
        }]);
        assert_eq!(env.resolve("g1"), 0);
    }

    #[test]
    fn test_add_div_zero_denom() {
        let mut env = env_1m();
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::AddDiv, x: "w", y: "h", z: "0",
        }]);
        assert_eq!(env.resolve("g1"), 0);
    }

    #[test]
    fn test_if_else_zero() {
        let mut env = env_1m();
        // x == 0 should go to else branch
        env.evaluate_guides(&[GuideDef {
            name: "g1", op: FormulaOp::IfElse, x: "0", y: "100", z: "200",
        }]);
        assert_eq!(env.resolve("g1"), 200);
    }

    #[test]
    fn test_builtin_wd_divisors() {
        let env = GuideEnv::new(1_200_000, 600_000);
        assert_eq!(env.resolve("wd2"), 600_000);
        assert_eq!(env.resolve("wd3"), 400_000);
        assert_eq!(env.resolve("wd4"), 300_000);
        assert_eq!(env.resolve("wd5"), 240_000);
        assert_eq!(env.resolve("wd6"), 200_000);
        assert_eq!(env.resolve("wd8"), 150_000);
        assert_eq!(env.resolve("wd10"), 120_000);
        assert_eq!(env.resolve("wd12"), 100_000);
    }

    #[test]
    fn test_builtin_hd_divisors() {
        let env = GuideEnv::new(1_200_000, 700_000);
        assert_eq!(env.resolve("hd2"), 350_000);
        assert_eq!(env.resolve("hd3"), 233_333);
        assert_eq!(env.resolve("hd4"), 175_000);
        assert_eq!(env.resolve("hd5"), 140_000);
        assert_eq!(env.resolve("hd6"), 116_666);
        assert_eq!(env.resolve("hd8"), 87_500);
        assert_eq!(env.resolve("hd10"), 70_000);
    }

    #[test]
    fn test_builtin_ssd_divisors() {
        let env = GuideEnv::new(2_000_000, 640_000);
        // ss = min(2M, 640K) = 640_000
        assert_eq!(env.resolve("ssd2"), 320_000);
        assert_eq!(env.resolve("ssd4"), 160_000);
        assert_eq!(env.resolve("ssd6"), 106_666);
        assert_eq!(env.resolve("ssd8"), 80_000);
        assert_eq!(env.resolve("ssd16"), 40_000);
        assert_eq!(env.resolve("ssd32"), 20_000);
    }

    #[test]
    fn test_builtin_all_angle_constants() {
        let env = env_1m();
        assert_eq!(env.resolve("cd2"), 10_800_000);   // 180 deg
        assert_eq!(env.resolve("cd4"), 5_400_000);     // 90 deg
        assert_eq!(env.resolve("cd8"), 2_700_000);     // 45 deg
        assert_eq!(env.resolve("3cd4"), 16_200_000);   // 270 deg
        assert_eq!(env.resolve("3cd8"), 8_100_000);    // 135 deg
        assert_eq!(env.resolve("5cd8"), 13_500_000);   // 225 deg
        assert_eq!(env.resolve("7cd8"), 18_900_000);   // 315 deg
    }

    #[test]
    fn test_literal_resolution() {
        let env = env_1m();
        assert_eq!(env.resolve("12345"), 12345);
        assert_eq!(env.resolve("-5000"), -5000);
        assert_eq!(env.resolve(""), 0);
        assert_eq!(env.resolve("unknown_name"), 0);
    }
}
