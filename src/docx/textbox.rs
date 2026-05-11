use std::collections::{HashMap, HashSet};
use std::io::Read;

use crate::geometry::{FormulaOp, PathFill};
use crate::model::{
    ArrowEnd, AutoFit, ConnectorShape, ConnectorType, CustomGeometry, CustomGuideDef,
    CustomPathCommand, CustomPathDef, HRelativeFrom, HorizontalPosition, Paragraph, ShapeFill,
    ShapeGeometry, TextAnchor, TextWarp, Textbox, VRelativeFrom, VerticalPosition, WrapType,
};

use super::images::{extent_dimensions, parse_anchor_position};
use super::color::{apply_color_transforms, parse_solid_fill, resolve_dml_color};
use super::styles::{
    ThemeFillStyle, ThemeFonts,
};
use super::{
    DML_NS, MC_NS_TOP, ParseContext, VML_NS, WML_NS, WPD_NS, WPS_NS,
    dml as find_dml, dml_children as find_dml_all, wps as find_wps,
    emu_attr,
};

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

fn find_sp_pr<'a>(wsp: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    wsp.children().find(|n| {
        n.tag_name().name() == "spPr"
            && (n.tag_name().namespace() == Some(WPS_NS)
                || n.tag_name().namespace() == Some(DML_NS))
    })
}

fn find_wps_style_ref<'a>(
    wsp: roxmltree::Node<'a, 'a>,
    ref_name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    let style = find_wps(wsp, "style")?;
    find_dml(style, ref_name)
}

pub(super) fn parse_txbx_content_paragraphs<R: Read + std::io::Seek>(
    txbx_content: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Vec<Paragraph> {
    let mut paragraphs = Vec::new();
    let mut counters: HashMap<(u32, u8), u32> = HashMap::new();
    let mut last_seen_level: HashMap<u32, u8> = HashMap::new();
    let mut applied_overrides: HashSet<(u32, u8)> = HashSet::new();
    let opts = super::paragraph::ParagraphOptions::default();
    for p in txbx_content
        .children()
        .filter(|n| n.tag_name().name() == "p" && n.tag_name().namespace() == Some(WML_NS))
    {
        paragraphs.push(super::paragraph::build_paragraph(
            p, ctx, &mut counters, &mut last_seen_level, &mut applied_overrides, &opts,
        ));
    }
    paragraphs
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
            resolve_dml_color(gs, theme).map(|color| (color, pos))
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

    let base_color = resolve_dml_color(fill_ref, theme)?;

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

pub(super) fn parse_avlst(parent: roxmltree::Node) -> Vec<(String, i64)> {
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
            .map(super::emu_to_pts)
            .unwrap_or(default)
    };
    (
        emu_to_pt("tIns", 3.6),
        emu_to_pt("lIns", 7.2),
        emu_to_pt("bIns", 3.6),
        emu_to_pt("rIns", 7.2),
    )
}

/// WordArt defaults to 0 margins unless explicitly set.
fn parse_wordart_body_margins(wsp: roxmltree::Node) -> (f32, f32, f32, f32) {
    let Some(bp) = find_wps(wsp, "bodyPr") else {
        return (0.0, 0.0, 0.0, 0.0);
    };
    let emu_to_pt = |attr: &str| -> f32 {
        bp.attribute(attr)
            .and_then(|v| v.parse::<f32>().ok())
            .map(super::emu_to_pts)
            .unwrap_or(0.0)
    };
    (
        emu_to_pt("tIns"),
        emu_to_pt("lIns"),
        emu_to_pt("bIns"),
        emu_to_pt("rIns"),
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
    pub(super) is_wordart: bool,
    pub(super) text_warp: Option<TextWarp>,
    pub(super) auto_fit: AutoFit,
}

pub(super) fn parse_textbox_from_wsp<R: Read + std::io::Seek>(
    anchor: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<WspResult> {
    let wsp = anchor
        .descendants()
        .find(|n| n.tag_name().name() == "wsp" && n.tag_name().namespace() == Some(WPS_NS))?;

    let sp_pr = find_sp_pr(wsp);
    let fill: Option<ShapeFill> = sp_pr
        .and_then(|sp| {
            parse_solid_fill(sp, ctx.theme)
                .map(ShapeFill::Solid)
                .or_else(|| parse_gradient_fill(sp, ctx.theme))
        })
        .or_else(|| parse_style_fill(wsp, ctx.theme));
    let has_no_fill = sp_pr.is_some_and(|sp| find_dml(sp, "noFill").is_some());

    let (stroke_color, stroke_width) = sp_pr
        .and_then(|sp| find_dml(sp, "ln"))
        .and_then(|ln| {
            if find_dml(ln, "noFill").is_some() {
                return None;
            }
            let color = parse_solid_fill(ln, ctx.theme)?;
            let width = ln
                .attribute("w")
                .and_then(|v| v.parse::<f32>().ok())
                .map(super::emu_to_pts)
                .unwrap_or(0.75);
            Some((color, width))
        })
        .map_or((None, 0.0), |(c, w)| (Some(c), w));

    let shape_type = sp_pr.map(parse_shape_geometry).unwrap_or_default();

    let (mut margin_top, mut margin_left, mut margin_bottom, mut margin_right) =
        parse_body_margins(wsp);

    let body_pr = find_wps(wsp, "bodyPr");
    let no_text_wrap = body_pr
        .and_then(|bp| bp.attribute("wrap"))
        .is_some_and(|w| w == "none");

    let text_anchor = match body_pr.and_then(|bp| bp.attribute("anchor")) {
        Some("ctr") => TextAnchor::Middle,
        Some("b") => TextAnchor::Bottom,
        _ => TextAnchor::Top,
    };

    let wa_props = body_pr
        .map(super::wordart::parse_wordart_body_pr)
        .unwrap_or(super::wordart::WordArtBodyProps {
            is_wordart: false,
            text_warp: None,
            auto_fit: AutoFit::None,
        });

    // WordArt defaults to zero body margins unless explicitly set
    if wa_props.is_wordart {
        let (wt, wl, wb, wr) = parse_wordart_body_margins(wsp);
        margin_top = wt;
        margin_left = wl;
        margin_bottom = wb;
        margin_right = wr;
    }

    let paragraphs = find_wps(wsp, "txbx")
        .and_then(|txbx| {
            txbx.children().find(|n| {
                n.tag_name().name() == "txbxContent" && n.tag_name().namespace() == Some(WML_NS)
            })
        })
        .map(|tc| parse_txbx_content_paragraphs(tc, ctx))
        .unwrap_or_default();

    if paragraphs.is_empty() && (has_no_fill || fill.is_none()) && stroke_color.is_none() {
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
        is_wordart: wa_props.is_wordart,
        text_warp: wa_props.text_warp,
        auto_fit: wa_props.auto_fit,
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
        "line" | "straightConnector1" => {
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

    let ln_node = find_dml(sp_pr, "ln");
    let stroke_color = parse_style_stroke(wsp, theme)
        .or_else(|| ln_node.and_then(|ln| parse_solid_fill(ln, theme)))
        .unwrap_or([0, 0, 0]);
    let stroke_width = ln_node
        .and_then(|ln| ln.attribute("w"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(super::emu_to_pts)
        .unwrap_or_else(|| parse_style_stroke_width(wsp));

    let (head_end, tail_end) = ln_node
        .map(|ln| {
            let head = find_dml(ln, "headEnd")
                .and_then(|n| n.attribute("type"))
                .map(ArrowEnd::from_attr)
                .unwrap_or_default();
            let tail = find_dml(ln, "tailEnd")
                .and_then(|n| n.attribute("type"))
                .map(ArrowEnd::from_attr)
                .unwrap_or_default();
            (head, tail)
        })
        .unwrap_or_default();

    let (h_position, _, v_pos, _) = parse_anchor_position(anchor);
    let (display_w, display_h) = extent_dimensions(anchor);
    let v_offset = match v_pos {
        VerticalPosition::Offset(o) => o,
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
        head_end,
        tail_end,
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
    resolve_dml_color(ln_ref, theme)
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

pub(super) fn parse_textbox_from_vml<R: Read + std::io::Seek>(
    pict_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<Textbox> {
    let shape = pict_node.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS) && matches!(n.tag_name().name(), "shape" | "rect")
    })?;

    // VML WordArt uses v:textpath instead of v:textbox
    let Some(textbox_node) = shape
        .children()
        .find(|n| n.tag_name().name() == "textbox" && n.tag_name().namespace() == Some(VML_NS))
    else {
        let tp = shape.children().find(|n| {
            n.tag_name().name() == "textpath" && n.tag_name().namespace() == Some(VML_NS)
        });
        return tp.and_then(|tp| super::wordart::parse_vml_wordart(shape, tp, ctx.styles, ctx.theme));
    };
    let txbx_content = textbox_node.children().find(|n| {
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
        parse_txbx_content_paragraphs(txbx_content, ctx);
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
        v_position: VerticalPosition::Offset(margin_top),
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
        is_wordart: false,
        text_warp: None,
        auto_fit: AutoFit::None,
    })
}

pub(super) fn collect_textboxes_from_paragraph<R: Read + std::io::Seek>(
    para_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
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
                for drawing in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "drawing"
                }) {
                    for container in drawing.children().filter(|n| {
                        n.tag_name().namespace() == Some(WPD_NS) && n.tag_name().name() == "anchor"
                    }) {
                        let (display_w, display_h) = extent_dimensions(container);

                        if let Some(wsp) =
                            parse_textbox_from_wsp(container, ctx)
                        {
                            let (h_position, h_relative, v_pos, v_relative) =
                                parse_anchor_position(container);
                            let v_offset = match v_pos {
                                VerticalPosition::Offset(o) => o,
                                _ => 0.0,
                            };
                            let (wrap_type, _, _) = super::images::parse_wrap_type(container);
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
                                v_position: v_pos,
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
                                is_wordart: wsp.is_wordart,
                                text_warp: wsp.text_warp,
                                auto_fit: wsp.auto_fit,
                            });
                        }
                    }
                }
            } else if let Some(branch) = fallback {
                for pict in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS) && n.tag_name().name() == "pict"
                }) {
                    if let Some(tb) =
                        parse_textbox_from_vml(pict, ctx)
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
                            parse_textbox_from_vml(pict, ctx)
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
