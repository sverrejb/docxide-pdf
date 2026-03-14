use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    Alignment, ConnectorShape, ConnectorType, CustomGeometry, CustomGuideDef, CustomPathCommand,
    CustomPathDef, HRelativeFrom, HorizontalPosition, LineSpacing, Paragraph, ShapeFill,
    ShapeGeometry, TextAnchor, Textbox, VRelativeFrom, WrapType,
};

use super::images::{extent_dimensions, parse_anchor_position};
use super::numbering::{ListLabelInfo, NumberingInfo};
use super::runs::parse_runs;
use super::styles::{
    ColorTransforms, StylesInfo, ThemeFillStyle, ThemeFonts, parse_alignment,
    parse_color_transforms,
};
use super::{
    DML_NS, MC_NS_TOP, WML_NS, WPD_NS, WPS_NS, extract_indents, parse_paragraph_spacing,
    resolve_theme_color_key, wml, wml_attr,
};

fn find_dml<'a>(parent: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    parent
        .children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

fn find_dml_all<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    parent
        .children()
        .filter(move |n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

fn collect_dml_points(parent: roxmltree::Node) -> Vec<(String, String)> {
    find_dml_all(parent, "pt")
        .map(|pt| {
            (
                pt.attribute("x").unwrap_or("0").to_string(),
                pt.attribute("y").unwrap_or("0").to_string(),
            )
        })
        .collect()
}

fn find_wps<'a>(parent: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    parent
        .children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(WPS_NS))
}

fn find_sp_pr<'a>(wsp: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    wsp.children().find(|n| {
        n.tag_name().name() == "spPr"
            && (n.tag_name().namespace() == Some(WPS_NS)
                || n.tag_name().namespace() == Some(DML_NS))
    })
}

fn resolve_color_child(parent: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    if let Some(srgb) = find_dml(parent, "srgbClr") {
        return srgb.attribute("val").and_then(super::parse_hex_color);
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

fn find_wps_style_ref<'a>(
    wsp: roxmltree::Node<'a, 'a>,
    ref_name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    let style = find_wps(wsp, "style")?;
    find_dml(style, ref_name)
}

fn emu_attr(node: roxmltree::Node, attr: &str) -> f32 {
    node.attribute(attr)
        .and_then(|v| v.parse::<f32>().ok())
        .unwrap_or(0.0)
        / 12700.0
}

pub(super) fn parse_txbx_content_paragraphs<R: Read + std::io::Seek>(
    txbx_content: roxmltree::Node,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    numbering: &NumberingInfo,
) -> Vec<Paragraph> {
    let mut paragraphs = Vec::new();
    let mut counters: HashMap<(String, u8), u32> = HashMap::new();
    let mut last_seen_level: HashMap<String, u8> = HashMap::new();
    for p in txbx_content
        .children()
        .filter(|n| n.tag_name().name() == "p" && n.tag_name().namespace() == Some(WML_NS))
    {
        let parsed = parse_runs(p, styles, theme, rels, zip, numbering);
        let ppr = wml(p, "pPr");
        let para_style_id = ppr
            .and_then(|ppr| wml_attr(ppr, "pStyle"))
            .unwrap_or(&styles.default_paragraph_style_id);
        let para_style = styles.paragraph_styles.get(para_style_id);
        let alignment = ppr
            .and_then(|ppr| wml_attr(ppr, "jc"))
            .map(parse_alignment)
            .or_else(|| para_style.and_then(|s| s.alignment))
            .unwrap_or(Alignment::Left);
        let (sp_before, sp_after, ls) = parse_paragraph_spacing(ppr, para_style);
        let space_before = sp_before.unwrap_or(0.0);
        let space_after = sp_after.unwrap_or(0.0);
        let line_spacing = Some(ls.unwrap_or(LineSpacing::Auto(1.0)));
        let tab_stops = ppr.map(super::parse_tab_stops).unwrap_or_default();
        let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
        let ListLabelInfo {
            mut indent_left,
            mut indent_hanging,
            label: list_label,
            font: list_label_font,
            font_size: list_label_font_size,
            bold: list_label_bold,
            color: list_label_color,
        } = super::numbering::parse_list_info(
            num_pr,
            None,
            None,
            numbering,
            &mut counters,
            &mut last_seen_level,
        );
        let mut indent_first_line = 0.0f32;
        let mut indent_right = 0.0f32;
        if let Some(ind) = ppr.and_then(|ppr| wml(ppr, "ind")) {
            let (left, right, hanging, first) = extract_indents(ind);
            if let Some(v) = left {
                indent_left = v;
            }
            if let Some(v) = right {
                indent_right = v;
            }
            if let Some(v) = hanging {
                indent_hanging = v;
            }
            if let Some(v) = first {
                indent_first_line = v;
            }
        } else if list_label.is_empty() {
            if let Some(s) = para_style {
                if let Some(v) = s.indent_left {
                    indent_left = v;
                }
                if let Some(v) = s.indent_right {
                    indent_right = v;
                }
                if let Some(v) = s.indent_hanging {
                    indent_hanging = v;
                }
                if let Some(v) = s.indent_first_line {
                    indent_first_line = v;
                }
            }
        }
        paragraphs.push(Paragraph {
            runs: parsed.runs,
            space_before,
            space_after,
            alignment,
            indent_left,
            indent_right,
            indent_hanging,
            indent_first_line,
            list_label,
            list_label_font,
            list_label_font_size,
            list_label_bold,
            list_label_color,
            line_spacing,
            tab_stops,
            floating_images: parsed.floating_images,
            textboxes: parsed.textboxes,
            ..Paragraph::default()
        });
    }
    paragraphs
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

fn apply_color_transforms(base: [u8; 3], t: &ColorTransforms) -> [u8; 3] {
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
    color
}

pub(super) fn parse_solid_fill(sp_pr: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    let fill = find_dml(sp_pr, "solidFill")?;
    resolve_color_child(fill, theme)
}

fn parse_gradient_fill(sp_pr: roxmltree::Node, theme: &ThemeFonts) -> Option<ShapeFill> {
    let grad_fill = find_dml(sp_pr, "gradFill")?;
    let gs_lst = find_dml(grad_fill, "gsLst")?;

    let stops: Vec<([u8; 3], f32)> = find_dml_all(gs_lst, "gs")
        .filter_map(|gs| {
            let pos = gs
                .attribute("pos")
                .and_then(|v| v.parse::<f32>().ok())
                .map(|v| v / 100_000.0)
                .unwrap_or(0.0);
            resolve_color_child(gs, theme).map(|color| (color, pos))
        })
        .collect();
    if stops.is_empty() {
        return None;
    }

    // OOXML a:lin @ang is in 60,000ths of a degree
    let angle_deg = find_dml(grad_fill, "lin")
        .and_then(|lin| lin.attribute("ang"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 60_000.0)
        .unwrap_or(0.0);

    Some(ShapeFill::LinearGradient { stops, angle_deg })
}

fn parse_style_fill(wsp: roxmltree::Node, theme: &ThemeFonts) -> Option<ShapeFill> {
    let fill_ref = find_wps_style_ref(wsp, "fillRef")?;

    let idx = fill_ref
        .attribute("idx")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0);
    if idx == 0 {
        return None;
    }

    let base_color = resolve_color_child(fill_ref, theme)?;

    let fill_style_idx = (idx as usize).saturating_sub(1);
    match theme.fill_styles.get(fill_style_idx) {
        Some(ThemeFillStyle::Gradient { stops, angle_deg }) if !stops.is_empty() => {
            let resolved_stops: Vec<([u8; 3], f32)> = stops
                .iter()
                .map(|stop| {
                    let color = apply_color_transforms(base_color, &stop.transforms);
                    (color, stop.position)
                })
                .collect();
            Some(ShapeFill::LinearGradient {
                stops: resolved_stops,
                angle_deg: *angle_deg,
            })
        }
        _ => Some(ShapeFill::Solid(base_color)),
    }
}

/// Parse `a:prstGeom` or `a:custGeom` from an spPr node into `ShapeGeometry`.
pub(super) fn parse_shape_geometry(sp_pr: roxmltree::Node) -> ShapeGeometry {
    if let Some(prst_geom) = find_dml(sp_pr, "prstGeom") {
        let preset = prst_geom.attribute("prst").unwrap_or("rect").to_string();
        let adjustments = parse_avlst(prst_geom);
        return ShapeGeometry {
            preset: Some(preset),
            adjustments,
            custom: None,
        };
    }

    if let Some(cust_geom) = find_dml(sp_pr, "custGeom") {
        if let Some(custom) = parse_custom_geometry(cust_geom) {
            return ShapeGeometry {
                preset: None,
                adjustments: Vec::new(),
                custom: Some(custom),
            };
        }
    }

    ShapeGeometry::default()
}

fn parse_avlst(parent: roxmltree::Node) -> Vec<(String, i64)> {
    let Some(avlst) = find_dml(parent, "avLst") else {
        return Vec::new();
    };
    find_dml_all(avlst, "gd")
        .filter_map(|gd| {
            let name = gd.attribute("name")?.to_string();
            let fmla = gd.attribute("fmla")?;
            let val = fmla.strip_prefix("val ")?.trim().parse::<i64>().ok()?;
            Some((name, val))
        })
        .collect()
}

fn parse_custom_geometry(cust_geom: roxmltree::Node) -> Option<CustomGeometry> {
    use crate::geometry::FormulaOp;
    use crate::geometry::PathFill;

    let adjust_defaults = parse_avlst(cust_geom);

    let guides = find_dml(cust_geom, "gdLst")
        .map(|gdlst| {
            find_dml_all(gdlst, "gd")
                .filter_map(|gd| {
                    let name = gd.attribute("name")?.to_string();
                    let fmla = gd.attribute("fmla")?;
                    let parts: Vec<&str> = fmla.split_whitespace().collect();
                    let op = FormulaOp::from_str(parts.first()?)?;
                    let x = parts.get(1).unwrap_or(&"").to_string();
                    let y = parts.get(2).unwrap_or(&"").to_string();
                    let z = parts.get(3).unwrap_or(&"").to_string();
                    Some(CustomGuideDef { name, op, x, y, z })
                })
                .collect()
        })
        .unwrap_or_default();

    let paths = find_dml(cust_geom, "pathLst")
        .map(|path_lst| {
            find_dml_all(path_lst, "path")
                .map(|path| {
                    let w = path.attribute("w").and_then(|v| v.parse::<i64>().ok());
                    let h = path.attribute("h").and_then(|v| v.parse::<i64>().ok());
                    let fill = match path.attribute("fill") {
                        Some("none") => PathFill::None,
                        _ => PathFill::Norm,
                    };
                    let stroke = path.attribute("stroke") != Some("0");
                    let commands = parse_path_commands(path);
                    CustomPathDef {
                        commands,
                        w,
                        h,
                        fill,
                        stroke,
                    }
                })
                .collect()
        })
        .unwrap_or_default();

    Some(CustomGeometry {
        adjust_defaults,
        guides,
        paths,
    })
}

fn parse_path_commands(path: roxmltree::Node) -> Vec<CustomPathCommand> {
    let mut commands = Vec::new();
    for child in path
        .children()
        .filter(|n| n.tag_name().namespace() == Some(DML_NS))
    {
        match child.tag_name().name() {
            "moveTo" => {
                if let Some((x, y)) = dml_pt(child) {
                    commands.push(CustomPathCommand::MoveTo { x, y });
                }
            }
            "lnTo" => {
                if let Some((x, y)) = dml_pt(child) {
                    commands.push(CustomPathCommand::LineTo { x, y });
                }
            }
            "arcTo" => {
                commands.push(CustomPathCommand::ArcTo {
                    wr: child.attribute("wR").unwrap_or("0").to_string(),
                    hr: child.attribute("hR").unwrap_or("0").to_string(),
                    st_ang: child.attribute("stAng").unwrap_or("0").to_string(),
                    sw_ang: child.attribute("swAng").unwrap_or("0").to_string(),
                });
            }
            "cubicBezTo" => {
                let pts = collect_dml_points(child);
                if pts.len() == 3 {
                    commands.push(CustomPathCommand::CubicBezTo {
                        x1: pts[0].0.clone(),
                        y1: pts[0].1.clone(),
                        x2: pts[1].0.clone(),
                        y2: pts[1].1.clone(),
                        x3: pts[2].0.clone(),
                        y3: pts[2].1.clone(),
                    });
                }
            }
            "quadBezTo" => {
                let pts = collect_dml_points(child);
                if pts.len() == 2 {
                    commands.push(CustomPathCommand::QuadBezTo {
                        x1: pts[0].0.clone(),
                        y1: pts[0].1.clone(),
                        x2: pts[1].0.clone(),
                        y2: pts[1].1.clone(),
                    });
                }
            }
            "close" => {
                commands.push(CustomPathCommand::Close);
            }
            _ => {}
        }
    }
    commands
}

fn dml_pt(parent: roxmltree::Node) -> Option<(String, String)> {
    let pt = find_dml(parent, "pt")?;
    Some((
        pt.attribute("x").unwrap_or("0").to_string(),
        pt.attribute("y").unwrap_or("0").to_string(),
    ))
}

fn parse_body_margins(wsp: roxmltree::Node) -> (f32, f32, f32, f32) {
    let Some(bp) = find_wps(wsp, "bodyPr") else {
        return (3.6, 7.2, 3.6, 7.2); // Word defaults: 0.05" top/bottom, 0.1" left/right
    };
    let emu_to_pt = |attr: &str, default: f32| -> f32 {
        bp.attribute(attr)
            .and_then(|v| v.parse::<f32>().ok())
            .map(|emu| emu / 12700.0)
            .unwrap_or(default)
    };
    (
        emu_to_pt("tIns", 3.6),
        emu_to_pt("lIns", 7.2),
        emu_to_pt("bIns", 3.6),
        emu_to_pt("rIns", 7.2),
    )
}

pub(super) struct WspResult {
    pub(super) paragraphs: Vec<Paragraph>,
    pub(super) fill: Option<ShapeFill>,
    pub(super) shape_type: ShapeGeometry,
    pub(super) stroke_color: Option<[u8; 3]>,
    pub(super) stroke_width: f32,
    pub(super) text_anchor: TextAnchor,
    pub(super) margin_top: f32,
    pub(super) margin_left: f32,
    pub(super) margin_bottom: f32,
    pub(super) margin_right: f32,
    pub(super) no_text_wrap: bool,
}

pub(super) fn parse_textbox_from_wsp<R: Read + std::io::Seek>(
    anchor: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> Option<WspResult> {
    let wsp = anchor
        .descendants()
        .find(|n| n.tag_name().name() == "wsp" && n.tag_name().namespace() == Some(WPS_NS))?;

    let sp_pr = find_sp_pr(wsp);
    let fill: Option<ShapeFill> = sp_pr
        .and_then(|sp| {
            parse_solid_fill(sp, theme)
                .map(ShapeFill::Solid)
                .or_else(|| parse_gradient_fill(sp, theme))
        })
        .or_else(|| parse_style_fill(wsp, theme));
    let has_no_fill = sp_pr.is_some_and(|sp| find_dml(sp, "noFill").is_some());

    let (stroke_color, stroke_width) = sp_pr
        .and_then(|sp| find_dml(sp, "ln"))
        .and_then(|ln| {
            if find_dml(ln, "noFill").is_some() {
                return None;
            }
            let color = parse_solid_fill(ln, theme)?;
            let width = ln
                .attribute("w")
                .and_then(|v| v.parse::<f32>().ok())
                .map(|emu| emu / 12700.0)
                .unwrap_or(0.75);
            Some((color, width))
        })
        .map_or((None, 0.0), |(c, w)| (Some(c), w));

    let shape_type = sp_pr.map(parse_shape_geometry).unwrap_or_default();

    let (margin_top, margin_left, margin_bottom, margin_right) = parse_body_margins(wsp);

    let body_pr = find_wps(wsp, "bodyPr");
    let no_text_wrap = body_pr
        .and_then(|bp| bp.attribute("wrap"))
        .is_some_and(|w| w == "none");

    let text_anchor = match body_pr.and_then(|bp| bp.attribute("anchor")) {
        Some("ctr") => TextAnchor::Middle,
        Some("b") => TextAnchor::Bottom,
        _ => TextAnchor::Top,
    };

    let paragraphs = find_wps(wsp, "txbx")
        .and_then(|txbx| {
            txbx.children().find(|n| {
                n.tag_name().name() == "txbxContent" && n.tag_name().namespace() == Some(WML_NS)
            })
        })
        .map(|tc| parse_txbx_content_paragraphs(tc, styles, theme, rels, zip, numbering))
        .unwrap_or_default();

    if paragraphs.is_empty() && (has_no_fill || fill.is_none()) {
        return None;
    }

    Some(WspResult {
        paragraphs,
        fill,
        shape_type,
        stroke_color,
        stroke_width,
        text_anchor,
        margin_top,
        margin_left,
        margin_bottom,
        margin_right,
        no_text_wrap,
    })
}

pub(super) fn parse_connector_from_wsp(
    anchor: roxmltree::Node,
    theme: &ThemeFonts,
) -> Option<ConnectorShape> {
    let wsp = anchor
        .descendants()
        .find(|n| n.tag_name().name() == "wsp" && n.tag_name().namespace() == Some(WPS_NS))?;

    let sp_pr = find_sp_pr(wsp)?;
    let prst_geom = find_dml(sp_pr, "prstGeom")?;
    let prst = prst_geom.attribute("prst")?;

    let xfrm = find_dml(sp_pr, "xfrm");

    let connector_type = match prst {
        "line" => {
            let flip_h = xfrm
                .and_then(|x| x.attribute("flipH"))
                .is_some_and(|v| v == "1");
            let flip_v = xfrm
                .and_then(|x| x.attribute("flipV"))
                .is_some_and(|v| v == "1");
            ConnectorType::Line { flip_h, flip_v }
        }
        "arc" => {
            let rotation = xfrm
                .and_then(|x| x.attribute("rot"))
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0)
                / 60000.0;

            let mut adj1 = 0.0_f32;
            let mut adj2 = 0.0_f32;
            for gd in prst_geom
                .descendants()
                .filter(|n| n.tag_name().name() == "gd" && n.tag_name().namespace() == Some(DML_NS))
            {
                let name = gd.attribute("name").unwrap_or("");
                let val = gd
                    .attribute("fmla")
                    .and_then(|f| f.strip_prefix("val "))
                    .and_then(|v| v.parse::<f32>().ok())
                    .unwrap_or(0.0)
                    / 60000.0;
                match name {
                    "adj1" => adj1 = val,
                    "adj2" => adj2 = val,
                    _ => {}
                }
            }
            ConnectorType::Arc {
                start_angle: adj1,
                end_angle: adj2,
                rotation,
            }
        }
        _ => return None,
    };

    let stroke_color = parse_style_stroke(wsp, theme).unwrap_or([0, 0, 0]);
    let stroke_width = parse_style_stroke_width(wsp);

    let (h_position, _, v_pos, _) = parse_anchor_position(anchor);
    let (display_w, display_h) = extent_dimensions(anchor);
    let v_offset = match v_pos {
        crate::model::VerticalPosition::Offset(o) => o,
        _ => 0.0,
    };

    let x = match h_position {
        HorizontalPosition::Offset(v) => v,
        _ => 0.0,
    };

    Some(ConnectorShape {
        x,
        y: v_offset,
        width: display_w,
        height: display_h,
        stroke_color,
        stroke_width,
        connector_type,
    })
}

fn parse_style_stroke(wsp: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    let ln_ref = find_wps_style_ref(wsp, "lnRef")?;
    let idx = ln_ref
        .attribute("idx")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0);
    if idx == 0 {
        return None;
    }
    resolve_color_child(ln_ref, theme)
}

fn parse_style_stroke_width(wsp: roxmltree::Node) -> f32 {
    let idx = find_wps_style_ref(wsp, "lnRef")
        .and_then(|lr| lr.attribute("idx"))
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0);
    match idx {
        0 => 0.0,
        1 => 0.75,
        2 => 1.5,
        3 => 2.25,
        _ => 1.0,
    }
}

const VML_NS: &str = "urn:schemas-microsoft-com:vml";

pub(super) fn parse_textbox_from_vml<R: Read + std::io::Seek>(
    pict_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> Option<Textbox> {
    let shape = pict_node.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS) && matches!(n.tag_name().name(), "shape" | "rect")
    })?;
    let textbox = shape
        .children()
        .find(|n| n.tag_name().name() == "textbox" && n.tag_name().namespace() == Some(VML_NS))?;
    let txbx_content = textbox.children().find(|n| {
        n.tag_name().name() == "txbxContent" && n.tag_name().namespace() == Some(WML_NS)
    })?;

    let style_str = shape.attribute("style").unwrap_or("");
    let mut width = 0.0_f32;
    let mut height = 0.0_f32;
    let mut margin_left = 0.0_f32;
    let mut margin_top = 0.0_f32;
    let mut h_relative = HRelativeFrom::Column;
    let mut v_relative = VRelativeFrom::Paragraph;

    let parse_pt = |s: &str| -> f32 { s.trim_end_matches("pt").parse::<f32>().unwrap_or(0.0) };
    for part in style_str.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            let val = val.trim();
            match key.trim() {
                "width" => width = parse_pt(val),
                "height" => height = parse_pt(val),
                "margin-left" => margin_left = parse_pt(val),
                "margin-top" => margin_top = parse_pt(val),
                "mso-position-horizontal-relative" => {
                    h_relative = match val {
                        "page" => HRelativeFrom::Page,
                        "margin" => HRelativeFrom::Margin,
                        _ => HRelativeFrom::Column,
                    };
                }
                "mso-position-vertical-relative" => {
                    v_relative = match val {
                        "page" => VRelativeFrom::Page,
                        "margin" => VRelativeFrom::Margin,
                        _ => VRelativeFrom::Paragraph,
                    };
                }
                _ => {}
            }
        }
    }

    let paragraphs =
        parse_txbx_content_paragraphs(txbx_content, styles, theme, rels, zip, numbering);
    if paragraphs.is_empty() {
        return None;
    }
    Some(Textbox {
        paragraphs,
        width_pt: width,
        height_pt: height,
        h_position: HorizontalPosition::Offset(margin_left),
        h_relative_from: h_relative,
        v_offset_pt: margin_top,
        v_relative_from: v_relative,
        fill: None,
        shape_type: ShapeGeometry::default(),
        stroke_color: None,
        stroke_width: 0.0,
        text_anchor: TextAnchor::Top,
        margin_left: 7.2,
        margin_right: 7.2,
        margin_top: 3.6,
        margin_bottom: 3.6,
        wrap_type: WrapType::None,
        dist_top: 0.0,
        dist_bottom: 0.0,
        behind_doc: false,
        no_text_wrap: false,
    })
}

pub(super) fn collect_textboxes_from_paragraph<R: Read + std::io::Seek>(
    para_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> Vec<Textbox> {
    let mut textboxes = Vec::new();

    for child in para_node.children() {
        let ns = child.tag_name().namespace();
        let name = child.tag_name().name();
        if ns == Some(MC_NS_TOP) && name == "AlternateContent" {
            let choice = child.children().find(|n| {
                n.tag_name().namespace() == Some(MC_NS_TOP) && n.tag_name().name() == "Choice"
            });
            let fallback = child.children().find(|n| {
                n.tag_name().namespace() == Some(MC_NS_TOP) && n.tag_name().name() == "Fallback"
            });

            if let Some(branch) = choice {
                // DrawingML path: mc:Choice -> w:drawing -> wp:anchor -> wps:wsp -> wps:txbx
                for drawing in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "drawing"
                }) {
                    for container in drawing.children().filter(|n| {
                        n.tag_name().namespace() == Some(WPD_NS) && n.tag_name().name() == "anchor"
                    }) {
                        let (display_w, display_h) = extent_dimensions(container);

                        if let Some(wsp) =
                            parse_textbox_from_wsp(container, rels, zip, styles, theme, numbering)
                        {
                            let (h_position, h_relative, v_pos, v_relative) =
                                parse_anchor_position(container);
                            let v_offset = match v_pos {
                                crate::model::VerticalPosition::Offset(o) => o,
                                _ => 0.0,
                            };
                            let wrap_type = super::images::parse_wrap_type(container);
                            let behind_doc = container.attribute("behindDoc") == Some("1");
                            let dist_top = emu_attr(container, "distT");
                            let dist_bottom = emu_attr(container, "distB");
                            textboxes.push(Textbox {
                                paragraphs: wsp.paragraphs,
                                width_pt: display_w,
                                height_pt: display_h,
                                h_position,
                                h_relative_from: h_relative,
                                v_offset_pt: v_offset,
                                v_relative_from: v_relative,
                                fill: wsp.fill,
                                shape_type: wsp.shape_type,
                                stroke_color: wsp.stroke_color,
                                stroke_width: wsp.stroke_width,
                                text_anchor: wsp.text_anchor,
                                margin_left: wsp.margin_left,
                                margin_right: wsp.margin_right,
                                margin_top: wsp.margin_top,
                                margin_bottom: wsp.margin_bottom,
                                wrap_type,
                                dist_top,
                                dist_bottom,
                                behind_doc,
                                no_text_wrap: wsp.no_text_wrap,
                            });
                        }
                    }
                }
            } else if let Some(branch) = fallback {
                // VML fallback: mc:Fallback -> w:pict -> v:shape -> v:textbox
                for pict in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "pict"
                }) {
                    if let Some(tb) =
                        parse_textbox_from_vml(pict, rels, zip, styles, theme, numbering)
                    {
                        textboxes.push(tb);
                    }
                }
                for r in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "r"
                }) {
                    for pict in r.children().filter(|n| {
                        n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "pict"
                    }) {
                        if let Some(tb) =
                            parse_textbox_from_vml(pict, rels, zip, styles, theme, numbering)
                        {
                            textboxes.push(tb);
                        }
                    }
                }
            }
        }
    }
    textboxes
}
