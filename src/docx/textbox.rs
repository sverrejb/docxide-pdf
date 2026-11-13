use std::io::Read;

use crate::geometry::{FormulaOp, PathFill};
use crate::model::{
    ArrowEnd, AutoFit, ConnectorShape, ConnectorType, CustomGeometry, CustomGuideDef,
    CustomPathCommand, CustomPathDef, HRelativeFrom, HorizontalPosition, Paragraph, ShapeFill,
    ShapeGeometry, Table, TextAnchor, Textbox, TextboxLink, VRelativeFrom, VerticalPosition,
};

use super::color::{apply_color_transforms, parse_solid_fill, resolve_dml_color};
use super::images::{extent_dimensions, parse_anchor_position};
use super::styles::{ThemeFillStyle, ThemeFonts};
use super::{
    DML_NS, ParseContext, WML_NS, WPG_NS, WPS_NS, angle_attr, dml as find_dml,
    dml_children as find_dml_all, emu_attr_opt, frac_attr, wps as find_wps,
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

pub(super) fn find_sp_pr<'a>(wsp: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    wsp.children()
        .find(|n| n.has_tag_name((WPS_NS, "spPr")) || n.has_tag_name((DML_NS, "spPr")))
}

fn find_wps_style_ref<'a>(
    wsp: roxmltree::Node<'a, 'a>,
    ref_name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    let style = find_wps(wsp, "style")?;
    find_dml(style, ref_name)
}

#[derive(Default)]
pub(super) struct TxbxContent {
    pub(super) paragraphs: Vec<Paragraph>,
    pub(super) tables: Vec<(usize, Table)>,
}

pub(super) fn parse_txbx_content<R: Read + std::io::Seek>(
    txbx_content: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> TxbxContent {
    let outer = std::mem::replace(&mut ctx.in_text_box, true);
    let mut paragraphs = Vec::new();
    let mut tables = Vec::new();
    let mut lists = super::numbering::ListCounters::default();
    // Outline levels only tag headings here; the outline itself comes from the body.
    let opts = super::paragraph::ParagraphOptions {
        resolve_outline_level: true,
        ..Default::default()
    };
    for node in txbx_content.children() {
        if node.has_tag_name((WML_NS, "tbl")) {
            let table = super::parse_table_node(node, ctx, &mut lists);
            tables.push((paragraphs.len(), table));
            continue;
        }
        if !node.has_tag_name((WML_NS, "p")) {
            continue;
        }
        let mut para = super::paragraph::build_paragraph(node, ctx, &mut lists, &opts);
        // Auto spacing opens no story, as at the top of the body (air_pollution's
        // Normal (Web) text box starts at its inset, not 14pt below it).
        if paragraphs.is_empty() && tables.is_empty() && para.space_before_auto {
            para.space_before = 0.0;
        }
        paragraphs.push(para);
    }
    // A story ending in a table keeps a final paragraph mark below it, in the
    // default paragraph style: it counts for the box's anchor and autofit.
    if tables.last().is_some_and(|&(at, _)| at == paragraphs.len()) {
        let mark = format!(r#"<w:p xmlns:w="{WML_NS}"/>"#);
        if let Ok(mark) = roxmltree::Document::parse(&mark) {
            let para =
                super::paragraph::build_paragraph(mark.root_element(), ctx, &mut lists, &opts);
            paragraphs.push(para);
        }
    }
    ctx.in_text_box = outer;
    TxbxContent { paragraphs, tables }
}

fn parse_gradient_fill(sp_pr: roxmltree::Node, theme: &ThemeFonts) -> Option<ShapeFill> {
    let grad_fill = find_dml(sp_pr, "gradFill")?;
    let gs_lst = find_dml(grad_fill, "gsLst")?;

    let stops: Vec<([u8; 3], f32)> = find_dml_all(gs_lst, "gs")
        .filter_map(|gs| {
            let pos = frac_attr(gs, "pos").unwrap_or(0.0);
            resolve_dml_color(gs, theme).map(|color| (color, pos))
        })
        .collect();
    if stops.is_empty() {
        return None;
    }

    // OOXML a:lin @ang is in 60,000ths of a degree
    let angle_deg = find_dml(grad_fill, "lin")
        .and_then(|lin| angle_attr(lin, "ang"))
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

    if let Some(cust_geom) = find_dml(sp_pr, "custGeom")
        && let Some(custom) = parse_custom_geometry(cust_geom)
    {
        return ShapeGeometry {
            preset: None,
            adjustments: Vec::new(),
            custom: Some(custom),
        };
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
                if let Some((x, y)) = collect_dml_points(child).into_iter().next() {
                    commands.push(CustomPathCommand::MoveTo { x, y });
                }
            }
            "lnTo" => {
                if let Some((x, y)) = collect_dml_points(child).into_iter().next() {
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

/// Line and arc presets without text are connectors.
pub(super) fn is_connector_wsp(wsp: roxmltree::Node, sp_pr: roxmltree::Node) -> bool {
    let prst = find_dml(sp_pr, "prstGeom").and_then(|g| g.attribute("prst"));
    matches!(prst, Some("line" | "straightConnector1" | "arc")) && find_wps(wsp, "txbx").is_none()
}

pub(super) fn parse_textbox_from_wsp<R: Read + std::io::Seek>(
    anchor: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<Textbox> {
    let wsp = anchor
        .descendants()
        .find(|n| n.has_tag_name((WPS_NS, "wsp")))?;
    parse_wsp_shape(wsp, ctx, extent_dimensions(anchor))
}

/// Parse a single `wps:wsp` node into a shape result. Used both for the
/// single-shape drawing path and for children of canvases/groups. `display`
/// is the size the shape is shown at, which its picture fill is read for.
pub(super) fn parse_wsp_shape<R: Read + std::io::Seek>(
    wsp: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    display: (f32, f32),
) -> Option<Textbox> {
    let sp_pr = find_sp_pr(wsp);
    let (glow, reflection) = super::images::parse_glow_reflection(sp_pr);

    // Connectors are declined so the caller falls through to the connector
    // parser, which renders flipH/flipV and arc sweep angles that the
    // preset-geometry textbox path doesn't.
    if sp_pr.is_some_and(|sp| is_connector_wsp(wsp, sp)) {
        return None;
    }
    // `a:grpFill` takes the fill of the nearest enclosing group that has its
    // own, never the shape style's fillRef.
    let group_fill = sp_pr.is_some_and(|sp| find_dml(sp, "grpFill").is_some());
    let fill_pr = if group_fill {
        wsp.ancestors()
            .filter_map(|g| g.children().find(|n| n.has_tag_name((WPG_NS, "grpSpPr"))))
            .find(|&pr| find_dml(pr, "grpFill").is_none())
    } else {
        sp_pr
    };
    let has_no_fill = fill_pr.is_some_and(|sp| find_dml(sp, "noFill").is_some());
    let has_picture_fill = fill_pr.is_some_and(|sp| find_dml(sp, "blipFill").is_some());
    // Explicit a:noFill or a picture fill overrides any style-ref (fillRef) fill
    let fill: Option<ShapeFill> = if has_no_fill || has_picture_fill {
        None
    } else {
        fill_pr
            .and_then(|sp| {
                parse_solid_fill(sp, ctx.theme)
                    .map(ShapeFill::Solid)
                    .or_else(|| parse_gradient_fill(sp, ctx.theme))
            })
            .or_else(|| (!group_fill).then(|| parse_style_fill(wsp, ctx.theme))?)
    };

    let ln_node = sp_pr.and_then(|sp| find_dml(sp, "ln"));
    let ln_no_fill = ln_node.is_some_and(|ln| find_dml(ln, "noFill").is_some());
    let explicit_ln_width = ln_node.and_then(|ln| emu_attr_opt(ln, "w"));
    let outline_width =
        explicit_ln_width.unwrap_or_else(|| parse_style_stroke_width(wsp, ctx.theme));
    let (stroke_color, stroke_width) = if ln_no_fill {
        (None, 0.0)
    } else if let Some(color) = ln_node.and_then(|ln| parse_solid_fill(ln, ctx.theme)) {
        let width = if explicit_ln_width.is_none() && outline_width == 0.0 {
            0.75
        } else {
            outline_width
        };
        (Some(color), width)
    } else if let Some(color) = parse_style_stroke(wsp, ctx.theme) {
        // No explicit line color: fall back to the shape style's lnRef stroke
        (Some(color), outline_width)
    } else {
        (None, 0.0)
    };
    // Half the outline width from `w` or the style's `lnRef`, visible or not
    // (`a:noFill`); a drawn outline with neither (Word's default 0.75pt)
    // adds none.
    // Word halves the width in whole EMUs: a 0.25pt (3175 EMU) outline insets
    // 1587 EMU, so text at a box edge stays on its pixel (inset probe, every
    // width from 0.25 to 12pt at four sub-pixel box positions).
    let outline_inset = ((outline_width as f64 * 12700.0).round() as i64 / 2) as f32 / 12700.0;

    let stroke_dash: &[f32] = ln_node.map_or(&[], parse_prst_dash);

    let shape_type = sp_pr.map(parse_shape_geometry).unwrap_or_default();

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

    // Word's default body insets are 0.05" top/bottom and 0.1" left/right;
    // WordArt's are zero.
    let (v_inset, h_inset) = if wa_props.is_wordart {
        (0.0, 0.0)
    } else {
        (3.6, 7.2)
    };
    let inset = |attr, default| {
        body_pr
            .and_then(|bp| emu_attr_opt(bp, attr))
            .unwrap_or(default)
    };

    // A chain's first box holds the story (wps:txbx id); the others say
    // where they follow on (wps:linkedTxbx id seq).
    let txbx = find_wps(wsp, "txbx");
    let link = match (txbx, find_wps(wsp, "linkedTxbx")) {
        (Some(t), _) => t.attribute("id").map(|id| (id, "0")),
        (None, Some(l)) => l.attribute("id").zip(l.attribute("seq")),
        (None, None) => None,
    }
    .and_then(|(id, seq)| {
        Some(TextboxLink {
            id: id.parse().ok()?,
            seq: seq.parse().ok()?,
            ..Default::default()
        })
    });
    let TxbxContent { paragraphs, tables } = txbx
        .and_then(|txbx| {
            txbx.children()
                .find(|n| n.has_tag_name((WML_NS, "txbxContent")))
        })
        .map(|tc| parse_txbx_content(tc, ctx))
        .unwrap_or_default();

    // A text-less shape with its own picture fill is a picture with the
    // shape's outline (`picture_filled_shape`), not an outline around nothing.
    let own_picture_fill = !group_fill && has_picture_fill;
    if paragraphs.is_empty()
        && link.is_none()
        && (own_picture_fill || (fill.is_none() && !has_picture_fill && stroke_color.is_none()))
    {
        return None;
    }
    let picture_fill = fill_pr.and_then(|pr| {
        let mut img = super::images::shape_picture_fill(pr, ctx, display.0, display.1)?;
        // A group's picture spans the group: each shape taking it shows its
        // own part (Word example: three children side by side split it).
        if group_fill && let Some(window) = super::group::group_fill_window(wsp, pr) {
            img.src_rect = Some(crop_within(img.src_rect, window));
        }
        Some(img)
    });

    // Word measures the body insets from the outline's inner edge, so an
    // outline moves the text in by half its width on every side
    // (textbox-outline-inset). Size, anchoring and wrapping stay at their
    // defaults for the caller to fill in.
    Some(Textbox {
        paragraphs,
        tables,
        fill,
        picture_fill,
        shape_type,
        stroke_color,
        stroke_width,
        stroke_dash,
        text_anchor,
        margin_top: inset("tIns", v_inset) + outline_inset,
        margin_left: inset("lIns", h_inset) + outline_inset,
        margin_bottom: inset("bIns", v_inset) + outline_inset,
        margin_right: inset("rIns", h_inset) + outline_inset,
        no_text_wrap,
        text_warp: wa_props.text_warp,
        auto_fit: wa_props.auto_fit,
        link,
        glow,
        reflection,
        ..Textbox::default()
    })
}

/// The `window` (left, top, right, bottom fractions) of a picture already
/// cropped to `crop`.
fn crop_within(crop: Option<[f32; 4]>, window: [f32; 4]) -> [f32; 4] {
    let [l, t, r, b] = crop.unwrap_or_default();
    let (w, h) = (1.0 - l - r, 1.0 - t - b);
    [
        l + w * window[0],
        t + h * window[1],
        r + w * window[2],
        b + h * window[3],
    ]
}

pub(super) fn parse_connector_from_wsp(
    anchor: roxmltree::Node,
    theme: &ThemeFonts,
) -> Option<ConnectorShape> {
    let wsp = anchor
        .descendants()
        .find(|n| n.has_tag_name((WPS_NS, "wsp")))?;

    let (h_position, h_relative_from, v_position, v_relative_from) = parse_anchor_position(anchor);
    let (display_w, display_h) = extent_dimensions(anchor);

    let mut conn = parse_connector_shape_node(wsp, theme)?;
    conn.h_position = h_position;
    conn.h_relative_from = h_relative_from;
    conn.v_position = v_position;
    conn.v_relative_from = v_relative_from;
    conn.width = display_w;
    conn.height = display_h;
    (conn.behind_doc, conn.z_index) = super::images::anchor_z_order(anchor);
    conn.v_from_line = super::images::anchored_to_line(anchor);
    Some(conn)
}

/// Parse a `wps:wsp` connector node (line/arc preset) into a ConnectorShape
/// with zeroed position/size — the caller supplies the placement.
pub(super) fn parse_connector_shape_node(
    wsp: roxmltree::Node,
    theme: &ThemeFonts,
) -> Option<ConnectorShape> {
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
            let rotation = xfrm.and_then(|x| angle_attr(x, "rot")).unwrap_or(0.0);

            let mut adj1 = 0.0_f32;
            let mut adj2 = 0.0_f32;
            for gd in prst_geom
                .descendants()
                .filter(|n| n.has_tag_name((DML_NS, "gd")))
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
    // Explicit a:ln formatting overrides the style matrix lnRef (OOXML precedence).
    // The 50cm connector sets a:ln/solidFill=tx1 (black) but lnRef=accent1 (blue);
    // Word draws it black, so the direct fill must win over the style ref.
    let stroke_color = ln_node
        .and_then(|ln| parse_solid_fill(ln, theme))
        .or_else(|| parse_style_stroke(wsp, theme))
        .unwrap_or([0, 0, 0]);
    let stroke_width = ln_node
        .and_then(|ln| emu_attr_opt(ln, "w"))
        .unwrap_or_else(|| parse_style_stroke_width(wsp, theme));

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

    Some(ConnectorShape {
        behind_doc: false,
        h_position: HorizontalPosition::Offset(0.0),
        h_relative_from: HRelativeFrom::Column,
        v_position: VerticalPosition::Offset(0.0),
        v_relative_from: VRelativeFrom::Paragraph,
        width: 0.0,
        height: 0.0,
        stroke_color,
        stroke_width,
        connector_type,
        head_end,
        tail_end,
        z_index: 0,
        v_from_line: false,
        on_line: false,
    })
}

/// An outline's `a:prstDash` as dash/gap lengths in line widths (empty for
/// solid), the DrawingML preset patterns Word strokes (shape-outline-dash).
fn parse_prst_dash(ln: roxmltree::Node) -> &'static [f32] {
    match find_dml(ln, "prstDash").and_then(|n| n.attribute("val")) {
        Some("dot") => &[1.0, 3.0],
        Some("dash") => &[4.0, 3.0],
        Some("lgDash") => &[8.0, 3.0],
        Some("dashDot") => &[4.0, 3.0, 1.0, 3.0],
        Some("lgDashDot") => &[8.0, 3.0, 1.0, 3.0],
        Some("lgDashDotDot") => &[8.0, 3.0, 1.0, 3.0, 1.0, 3.0],
        Some("sysDash") => &[3.0, 1.0],
        Some("sysDot") => &[1.0, 1.0],
        Some("sysDashDot") => &[3.0, 1.0, 1.0, 1.0],
        Some("sysDashDotDot") => &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0],
        _ => &[],
    }
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

/// The width of the theme line style a shape's lnRef picks (indonesian's
/// "Format 12" box: idx 2 of 9525/25400/38100 EMU is Word's 2pt outline).
fn parse_style_stroke_width(wsp: roxmltree::Node, theme: &ThemeFonts) -> f32 {
    let idx = find_wps_style_ref(wsp, "lnRef")
        .and_then(|lr| lr.attribute("idx"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);
    if let Some(&w) = idx.checked_sub(1).and_then(|i| theme.line_widths.get(i)) {
        return w;
    }
    match idx {
        0 => 0.0,
        1 => 0.75,
        2 => 1.5,
        3 => 2.25,
        _ => 1.0,
    }
}

#[cfg(test)]
mod connector_layer_tests {
    use super::*;

    #[test]
    fn anchored_connector_preserves_behind_document_layer() {
        let theme = ThemeFonts::default();
        for (attr, expected) in [("", false), ("0", false), ("1", true), ("true", true)] {
            let xml = format!(
                r#"<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" behindDoc="{attr}" relativeHeight="42"><wp:extent cx="914400" cy="0"/><wps:wsp><wps:spPr><a:prstGeom prst="straightConnector1"/></wps:spPr></wps:wsp></wp:anchor>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            let connector = parse_connector_from_wsp(doc.root_element(), &theme).unwrap();
            assert_eq!(connector.behind_doc, expected);
            assert_eq!(connector.z_index, 42);
        }
    }
}
