use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    Alignment, ConnectorShape, ConnectorType, CustomGeometry, CustomGuideDef, CustomPathCommand,
    CustomPathDef, HorizontalPosition, LineSpacing, Paragraph, ShapeFill, ShapeGeometry, Textbox,
    WrapType,
};

use super::images::{extent_dimensions, parse_anchor_position};
use super::numbering::NumberingInfo;
use super::runs::parse_runs;
use super::styles::{StylesInfo, ThemeFonts, parse_alignment, parse_line_spacing};
use super::{DML_NS, MC_NS_TOP, WML_NS, WPD_NS, WPS_NS, twips_attr, wml, wml_attr};

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
        let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
        let space_before = inline_spacing
            .and_then(|n| twips_attr(n, "before"))
            .or_else(|| para_style.and_then(|s| s.space_before))
            .unwrap_or(0.0);
        let space_after = inline_spacing
            .and_then(|n| twips_attr(n, "after"))
            .or_else(|| para_style.and_then(|s| s.space_after))
            .unwrap_or(0.0);
        let line_spacing = Some(
            inline_spacing
                .and_then(|n| {
                    n.attribute((WML_NS, "line"))
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|line_val| parse_line_spacing(n, line_val))
                })
                .or_else(|| para_style.and_then(|s| s.line_spacing))
                .unwrap_or(LineSpacing::Auto(1.0)),
        );
        let tab_stops = ppr.map(super::parse_tab_stops).unwrap_or_default();
        let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
        let (mut indent_left, mut indent_hanging, list_label, list_label_font) =
            super::numbering::parse_list_info(
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
            if let Some(v) = twips_attr(ind, "left") {
                indent_left = v;
            }
            if let Some(v) = twips_attr(ind, "right") {
                indent_right = v;
            }
            if let Some(v) = twips_attr(ind, "hanging") {
                indent_hanging = v;
            }
            if let Some(v) = twips_attr(ind, "firstLine") {
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
            line_spacing,
            tab_stops,
            extra_line_breaks: parsed.line_break_count,
            floating_images: parsed.floating_images,
            textboxes: parsed.textboxes,
            ..Paragraph::default()
        });
    }
    paragraphs
}

pub(super) fn resolve_scheme_color(base: [u8; 3], fill_node: roxmltree::Node) -> [u8; 3] {
    let mut lum_mod: Option<f32> = None;
    let mut lum_off: Option<f32> = None;
    let mut tint: Option<f32> = None;
    for child in fill_node.children() {
        if child.tag_name().namespace() != Some(DML_NS) {
            continue;
        }
        match child.tag_name().name() {
            "lumMod" => {
                lum_mod = child
                    .attribute("val")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(|v| v / 100_000.0);
            }
            "lumOff" => {
                lum_off = child
                    .attribute("val")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(|v| v / 100_000.0);
            }
            "tint" => {
                tint = child
                    .attribute("val")
                    .and_then(|v| v.parse::<f32>().ok())
                    .map(|v| v / 100_000.0);
            }
            _ => {}
        }
    }
    let mut color = base;
    if let Some(t) = tint {
        color = [
            (255.0 - t * (255.0 - color[0] as f32)).clamp(0.0, 255.0) as u8,
            (255.0 - t * (255.0 - color[1] as f32)).clamp(0.0, 255.0) as u8,
            (255.0 - t * (255.0 - color[2] as f32)).clamp(0.0, 255.0) as u8,
        ];
    }
    if lum_mod.is_some() || lum_off.is_some() {
        let m = lum_mod.unwrap_or(1.0);
        let o = lum_off.unwrap_or(0.0);
        color = [
            ((color[0] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
            ((color[1] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
            ((color[2] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
        ];
    }
    color
}

pub(super) fn parse_solid_fill(sp_pr: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    let fill = sp_pr
        .children()
        .find(|n| n.tag_name().name() == "solidFill" && n.tag_name().namespace() == Some(DML_NS))?;
    // Direct sRGB color
    if let Some(srgb) = fill
        .children()
        .find(|n| n.tag_name().name() == "srgbClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        return srgb.attribute("val").and_then(super::parse_hex_color);
    }
    // Theme color reference
    if let Some(scheme) = fill
        .children()
        .find(|n| n.tag_name().name() == "schemeClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        let val = scheme.attribute("val")?;
        // Map OOXML scheme names to theme element names
        let theme_key = match val {
            "dk1" => "dk1",
            "lt1" => "lt1",
            "dk2" => "dk2",
            "lt2" => "lt2",
            "tx1" => "dk1",
            "tx2" => "dk2",
            "bg1" => "lt1",
            "bg2" => "lt2",
            other => other,
        };
        let base = *theme.colors.get(theme_key)?;
        return Some(resolve_scheme_color(base, scheme));
    }
    None
}

fn resolve_stop_color(stop: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    if let Some(srgb) = stop
        .children()
        .find(|n| n.tag_name().name() == "srgbClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        return srgb.attribute("val").and_then(super::parse_hex_color);
    }
    if let Some(scheme) = stop
        .children()
        .find(|n| n.tag_name().name() == "schemeClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        let val = scheme.attribute("val")?;
        let theme_key = match val {
            "dk1" => "dk1",
            "lt1" => "lt1",
            "dk2" => "dk2",
            "lt2" => "lt2",
            "tx1" => "dk1",
            "tx2" => "dk2",
            "bg1" => "lt1",
            "bg2" => "lt2",
            other => other,
        };
        let base = *theme.colors.get(theme_key)?;
        return Some(resolve_scheme_color(base, scheme));
    }
    None
}

fn parse_gradient_fill(sp_pr: roxmltree::Node, theme: &ThemeFonts) -> Option<ShapeFill> {
    let grad_fill = sp_pr
        .children()
        .find(|n| n.tag_name().name() == "gradFill" && n.tag_name().namespace() == Some(DML_NS))?;
    let gs_lst = grad_fill
        .children()
        .find(|n| n.tag_name().name() == "gsLst" && n.tag_name().namespace() == Some(DML_NS))?;

    let mut stops: Vec<([u8; 3], f32)> = Vec::new();
    for gs in gs_lst
        .children()
        .filter(|n| n.tag_name().name() == "gs" && n.tag_name().namespace() == Some(DML_NS))
    {
        let pos = gs
            .attribute("pos")
            .and_then(|v| v.parse::<f32>().ok())
            .map(|v| v / 100_000.0)
            .unwrap_or(0.0);
        if let Some(color) = resolve_stop_color(gs, theme) {
            stops.push((color, pos));
        }
    }
    if stops.is_empty() {
        return None;
    }

    // OOXML a:lin @ang is in 60,000ths of a degree
    let angle_deg = grad_fill
        .children()
        .find(|n| n.tag_name().name() == "lin" && n.tag_name().namespace() == Some(DML_NS))
        .and_then(|lin| lin.attribute("ang"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(|v| v / 60_000.0)
        .unwrap_or(0.0);

    Some(ShapeFill::LinearGradient { stops, angle_deg })
}

fn parse_style_fill(wsp: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    let style = wsp.children().find(|n| {
        n.tag_name().name() == "style" && n.tag_name().namespace() == Some(WPS_NS)
    })?;
    let fill_ref = style.children().find(|n| {
        n.tag_name().name() == "fillRef" && n.tag_name().namespace() == Some(DML_NS)
    })?;

    let idx = fill_ref
        .attribute("idx")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0);
    if idx == 0 {
        return None;
    }

    if let Some(srgb) = fill_ref
        .children()
        .find(|n| n.tag_name().name() == "srgbClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        return srgb.attribute("val").and_then(super::parse_hex_color);
    }
    if let Some(scheme) = fill_ref
        .children()
        .find(|n| n.tag_name().name() == "schemeClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        let val = scheme.attribute("val")?;
        let theme_key = match val {
            "dk1" => "dk1",
            "lt1" => "lt1",
            "dk2" => "dk2",
            "lt2" => "lt2",
            "tx1" => "dk1",
            "tx2" => "dk2",
            "bg1" => "lt1",
            "bg2" => "lt2",
            other => other,
        };
        let base = *theme.colors.get(theme_key)?;
        return Some(resolve_scheme_color(base, scheme));
    }
    None
}

/// Parse `a:prstGeom` or `a:custGeom` from an spPr node into `ShapeGeometry`.
pub(super) fn parse_shape_geometry(sp_pr: roxmltree::Node) -> ShapeGeometry {
    // Try preset geometry first
    if let Some(prst_geom) = sp_pr.children().find(|n| {
        n.tag_name().name() == "prstGeom" && n.tag_name().namespace() == Some(DML_NS)
    }) {
        let preset = prst_geom.attribute("prst").unwrap_or("rect").to_string();
        let adjustments = parse_avlst(prst_geom);
        return ShapeGeometry {
            preset: Some(preset),
            adjustments,
            custom: None,
        };
    }

    // Try custom geometry
    if let Some(cust_geom) = sp_pr.children().find(|n| {
        n.tag_name().name() == "custGeom" && n.tag_name().namespace() == Some(DML_NS)
    }) {
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
    let Some(avlst) = parent.children().find(|n| {
        n.tag_name().name() == "avLst" && n.tag_name().namespace() == Some(DML_NS)
    }) else {
        return Vec::new();
    };
    avlst
        .children()
        .filter(|n| n.tag_name().name() == "gd" && n.tag_name().namespace() == Some(DML_NS))
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

    let guides = cust_geom
        .children()
        .find(|n| n.tag_name().name() == "gdLst" && n.tag_name().namespace() == Some(DML_NS))
        .map(|gdlst| {
            gdlst
                .children()
                .filter(|n| {
                    n.tag_name().name() == "gd" && n.tag_name().namespace() == Some(DML_NS)
                })
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

    let paths = cust_geom
        .children()
        .find(|n| n.tag_name().name() == "pathLst" && n.tag_name().namespace() == Some(DML_NS))
        .map(|path_lst| {
            path_lst
                .children()
                .filter(|n| {
                    n.tag_name().name() == "path" && n.tag_name().namespace() == Some(DML_NS)
                })
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
                if let Some(pt) = dml_pt(child) {
                    commands.push(CustomPathCommand::MoveTo {
                        x: pt.0,
                        y: pt.1,
                    });
                }
            }
            "lnTo" => {
                if let Some(pt) = dml_pt(child) {
                    commands.push(CustomPathCommand::LineTo {
                        x: pt.0,
                        y: pt.1,
                    });
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
                let pts: Vec<(String, String)> = child
                    .children()
                    .filter(|n| {
                        n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(DML_NS)
                    })
                    .map(|pt| {
                        (
                            pt.attribute("x").unwrap_or("0").to_string(),
                            pt.attribute("y").unwrap_or("0").to_string(),
                        )
                    })
                    .collect();
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
                let pts: Vec<(String, String)> = child
                    .children()
                    .filter(|n| {
                        n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(DML_NS)
                    })
                    .map(|pt| {
                        (
                            pt.attribute("x").unwrap_or("0").to_string(),
                            pt.attribute("y").unwrap_or("0").to_string(),
                        )
                    })
                    .collect();
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
    let pt = parent.children().find(|n| {
        n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(DML_NS)
    })?;
    Some((
        pt.attribute("x").unwrap_or("0").to_string(),
        pt.attribute("y").unwrap_or("0").to_string(),
    ))
}

fn parse_body_margins(wsp: roxmltree::Node) -> (f32, f32, f32, f32) {
    let body_pr = wsp
        .children()
        .find(|n| n.tag_name().name() == "bodyPr" && n.tag_name().namespace() == Some(WPS_NS));
    let Some(bp) = body_pr else {
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

    // Extract fill color from spPr
    let sp_pr = wsp.children().find(|n| {
        n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(WPS_NS)
            || n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(DML_NS)
    });
    let fill: Option<ShapeFill> = sp_pr
        .and_then(|sp| {
            parse_solid_fill(sp, theme)
                .map(ShapeFill::Solid)
                .or_else(|| parse_gradient_fill(sp, theme))
        })
        .or_else(|| parse_style_fill(wsp, theme).map(ShapeFill::Solid));
    let has_no_fill = sp_pr.is_some_and(|sp| {
        sp.children()
            .any(|n| n.tag_name().name() == "noFill" && n.tag_name().namespace() == Some(DML_NS))
    });

    let shape_type = sp_pr
        .map(parse_shape_geometry)
        .unwrap_or_default();

    let (margin_top, margin_left, margin_bottom, margin_right) = parse_body_margins(wsp);

    let no_text_wrap = wsp
        .children()
        .find(|n| n.tag_name().name() == "bodyPr" && n.tag_name().namespace() == Some(WPS_NS))
        .and_then(|bp| bp.attribute("wrap"))
        .is_some_and(|w| w == "none");

    // Try to get textbox content
    let paragraphs = wsp
        .children()
        .find(|n| n.tag_name().name() == "txbx" && n.tag_name().namespace() == Some(WPS_NS))
        .and_then(|txbx| {
            txbx.children().find(|n| {
                n.tag_name().name() == "txbxContent" && n.tag_name().namespace() == Some(WML_NS)
            })
        })
        .map(|tc| parse_txbx_content_paragraphs(tc, styles, theme, rels, zip, numbering))
        .unwrap_or_default();

    // Return if there's text content OR a visible fill
    if paragraphs.is_empty() && (has_no_fill || fill.is_none()) {
        return None;
    }

    Some(WspResult {
        paragraphs,
        fill,
        shape_type,
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

    let sp_pr = wsp.children().find(|n| {
        (n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(WPS_NS))
            || (n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(DML_NS))
    })?;

    let prst_geom = sp_pr.children().find(|n| {
        n.tag_name().name() == "prstGeom" && n.tag_name().namespace() == Some(DML_NS)
    })?;
    let prst = prst_geom.attribute("prst")?;

    let connector_type = match prst {
        "line" => {
            let xfrm = sp_pr.children().find(|n| {
                n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(DML_NS)
            });
            let flip_h = xfrm.and_then(|x| x.attribute("flipH")).is_some_and(|v| v == "1");
            let flip_v = xfrm.and_then(|x| x.attribute("flipV")).is_some_and(|v| v == "1");
            ConnectorType::Line { flip_h, flip_v }
        }
        "arc" => {
            let xfrm = sp_pr.children().find(|n| {
                n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(DML_NS)
            });
            let rotation = xfrm
                .and_then(|x| x.attribute("rot"))
                .and_then(|v| v.parse::<f32>().ok())
                .unwrap_or(0.0)
                / 60000.0;

            let mut adj1 = 0.0_f32;
            let mut adj2 = 0.0_f32;
            for gd in prst_geom.descendants().filter(|n| {
                n.tag_name().name() == "gd" && n.tag_name().namespace() == Some(DML_NS)
            }) {
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

    let (h_position, _, v_pos, _) = super::images::parse_anchor_position(anchor);
    let (display_w, display_h) = super::images::extent_dimensions(anchor);
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
    let style = wsp.children().find(|n| {
        n.tag_name().name() == "style" && n.tag_name().namespace() == Some(WPS_NS)
    })?;
    let ln_ref = style.children().find(|n| {
        n.tag_name().name() == "lnRef" && n.tag_name().namespace() == Some(DML_NS)
    })?;
    let idx = ln_ref
        .attribute("idx")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0);
    if idx == 0 {
        return None;
    }
    resolve_scheme_clr_child(ln_ref, theme)
}

fn parse_style_stroke_width(wsp: roxmltree::Node) -> f32 {
    let style = wsp.children().find(|n| {
        n.tag_name().name() == "style" && n.tag_name().namespace() == Some(WPS_NS)
    });
    let idx = style
        .and_then(|s| {
            s.children().find(|n| {
                n.tag_name().name() == "lnRef" && n.tag_name().namespace() == Some(DML_NS)
            })
        })
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

fn resolve_scheme_clr_child(parent: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    if let Some(srgb) = parent
        .children()
        .find(|n| n.tag_name().name() == "srgbClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        return srgb.attribute("val").and_then(super::parse_hex_color);
    }
    if let Some(scheme) = parent
        .children()
        .find(|n| n.tag_name().name() == "schemeClr" && n.tag_name().namespace() == Some(DML_NS))
    {
        let val = scheme.attribute("val")?;
        let theme_key = match val {
            "dk1" => "dk1",
            "lt1" => "lt1",
            "dk2" => "dk2",
            "lt2" => "lt2",
            "tx1" => "dk1",
            "tx2" => "dk2",
            "bg1" => "lt1",
            "bg2" => "lt2",
            other => other,
        };
        let base = *theme.colors.get(theme_key)?;
        return Some(resolve_scheme_color(base, scheme));
    }
    None
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
        n.tag_name().namespace() == Some(VML_NS)
            && (n.tag_name().name() == "shape" || n.tag_name().name() == "rect")
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
    let mut h_relative = "column";
    let mut v_relative = "paragraph";

    for part in style_str.split(';') {
        let part = part.trim();
        if let Some((key, val)) = part.split_once(':') {
            let key = key.trim();
            let val = val.trim();
            let parse_pt =
                |s: &str| -> f32 { s.trim_end_matches("pt").parse::<f32>().unwrap_or(0.0) };
            match key {
                "width" => width = parse_pt(val),
                "height" => height = parse_pt(val),
                "margin-left" => margin_left = parse_pt(val),
                "margin-top" => margin_top = parse_pt(val),
                "mso-position-horizontal-relative" => {
                    h_relative = match val {
                        "page" => "page",
                        "margin" => "margin",
                        _ => "column",
                    };
                }
                "mso-position-vertical-relative" => {
                    v_relative = match val {
                        "page" => "page",
                        "margin" => "margin",
                        _ => "paragraph",
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
                // DrawingML path: mc:Choice → w:drawing → wp:anchor → wps:wsp → wps:txbx
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
                            let dist_top = container
                                .attribute("distT")
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(0.0)
                                / 12700.0;
                            let dist_bottom = container
                                .attribute("distB")
                                .and_then(|v| v.parse::<f32>().ok())
                                .unwrap_or(0.0)
                                / 12700.0;
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
                // VML fallback: mc:Fallback → w:pict → v:shape → v:textbox
                for pict in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "pict"
                }) {
                    if let Some(tb) =
                        parse_textbox_from_vml(pict, rels, zip, styles, theme, numbering)
                    {
                        textboxes.push(tb);
                    }
                }
                // Also check for w:r/w:pict inside fallback
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
