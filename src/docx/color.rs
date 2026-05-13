use super::styles::ThemeFonts;
use super::{DML_NS, W14_NS, dml as find_dml, parse_hex_color, resolve_theme_color_key};

#[derive(Clone, Default)]
pub(super) struct ColorTransforms {
    pub(super) lum_mod: Option<f32>,
    pub(super) lum_off: Option<f32>,
    pub(super) sat_mod: Option<f32>,
    pub(super) tint: Option<f32>,
    pub(super) shade: Option<f32>,
    /// Hue offset in 60,000ths of a degree (SmartArt per-shape color variation)
    pub(super) hue_off: Option<f32>,
    /// Saturation offset in 1/1000ths of a percent
    pub(super) sat_off: Option<f32>,
}

pub(super) fn parse_color_transforms(scheme_clr: roxmltree::Node) -> ColorTransforms {
    let mut t = ColorTransforms::default();
    for child in scheme_clr.children().filter(|n| {
        matches!(n.tag_name().namespace(), Some(DML_NS) | Some(W14_NS))
    }) {
        let val = child
            .attribute("val")
            .and_then(|v| v.parse::<f32>().ok())
            .map(|v| v / 100_000.0);
        // hueOff/satOff/lumOff use raw integer units, not the /100_000 scaling
        let raw_val = child.attribute("val").and_then(|v| v.parse::<f32>().ok());
        match child.tag_name().name() {
            "lumMod" => t.lum_mod = val,
            "lumOff" => t.lum_off = val,
            "satMod" => t.sat_mod = val,
            "tint" => t.tint = val,
            "shade" => t.shade = val,
            "hueOff" => t.hue_off = raw_val,
            "satOff" => t.sat_off = raw_val,
            _ => {}
        }
    }
    t
}

pub(super) fn apply_color_transforms(base: [u8; 3], t: &ColorTransforms) -> [u8; 3] {
    let mut color = base;
    if let Some(tint) = t.tint {
        color = [
            (255.0 - tint * (255.0 - color[0] as f32)).clamp(0.0, 255.0) as u8,
            (255.0 - tint * (255.0 - color[1] as f32)).clamp(0.0, 255.0) as u8,
            (255.0 - tint * (255.0 - color[2] as f32)).clamp(0.0, 255.0) as u8,
        ];
    }
    if let Some(shade) = t.shade {
        color = [
            (color[0] as f32 * shade).clamp(0.0, 255.0) as u8,
            (color[1] as f32 * shade).clamp(0.0, 255.0) as u8,
            (color[2] as f32 * shade).clamp(0.0, 255.0) as u8,
        ];
    }
    if let Some(sat_mod) = t.sat_mod {
        if (sat_mod - 1.0).abs() > 0.001 {
            let (h, s, l) = rgb_to_hsl(color);
            color = hsl_to_rgb(h, (s * sat_mod).clamp(0.0, 1.0), l);
        }
    }
    if t.lum_mod.is_some() || t.lum_off.is_some() {
        let m = t.lum_mod.unwrap_or(1.0);
        let o = t.lum_off.unwrap_or(0.0);
        color = [
            ((color[0] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
            ((color[1] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
            ((color[2] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
        ];
    }
    // HSL offsets (used by SmartArt for per-shape color variation)
    if t.hue_off.is_some() || t.sat_off.is_some() {
        let (h, s, l) = rgb_to_hsl(color);
        // hueOff is in 60,000ths of a degree; convert to 0..1 range
        let h2 = if let Some(ho) = t.hue_off {
            (h + ho / (360.0 * 60_000.0)).rem_euclid(1.0)
        } else {
            h
        };
        // satOff is in 1/1000ths of a percent; convert to 0..1 range
        let s2 = if let Some(so) = t.sat_off {
            (s + so / 100_000.0).clamp(0.0, 1.0)
        } else {
            s
        };
        color = hsl_to_rgb(h2, s2, l);
    }
    color
}

/// Resolve a DrawingML color child (srgbClr or schemeClr with transforms).
pub(super) fn resolve_dml_color(parent: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    if let Some(srgb) = find_dml(parent, "srgbClr") {
        return srgb.attribute("val").and_then(parse_hex_color);
    }
    if let Some(scheme) = find_dml(parent, "schemeClr") {
        let val = scheme.attribute("val")?;
        let theme_key = resolve_theme_color_key(val);
        let base = *theme.colors.get(theme_key)?;
        let transforms = parse_color_transforms(scheme);
        return Some(apply_color_transforms(base, &transforms));
    }
    None
}

/// Extract solid fill color from a shape/text properties node.
pub(super) fn parse_solid_fill(sp_pr: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    let fill = find_dml(sp_pr, "solidFill")?;
    resolve_dml_color(fill, theme)
}

fn rgb_to_hsl(c: [u8; 3]) -> (f32, f32, f32) {
    let r = c[0] as f32 / 255.0;
    let g = c[1] as f32 / 255.0;
    let b = c[2] as f32 / 255.0;
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < f32::EPSILON {
        return (0.0, 0.0, l);
    }
    let d = max - min;
    let s = if l > 0.5 {
        d / (2.0 - max - min)
    } else {
        d / (max + min)
    };
    let h = if (max - r).abs() < f32::EPSILON {
        ((g - b) / d + if g < b { 6.0 } else { 0.0 }) / 6.0
    } else if (max - g).abs() < f32::EPSILON {
        ((b - r) / d + 2.0) / 6.0
    } else {
        ((r - g) / d + 4.0) / 6.0
    };
    (h, s, l)
}

fn hsl_to_rgb(h: f32, s: f32, l: f32) -> [u8; 3] {
    if s.abs() < f32::EPSILON {
        let v = (l * 255.0).clamp(0.0, 255.0) as u8;
        return [v, v, v];
    }
    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;
    let hue_to_rgb = |t: f32| -> f32 {
        let t = ((t % 1.0) + 1.0) % 1.0;
        if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 1.0 / 2.0 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        }
    };
    [
        (hue_to_rgb(h + 1.0 / 3.0) * 255.0).clamp(0.0, 255.0) as u8,
        (hue_to_rgb(h) * 255.0).clamp(0.0, 255.0) as u8,
        (hue_to_rgb(h - 1.0 / 3.0) * 255.0).clamp(0.0, 255.0) as u8,
    ]
}
