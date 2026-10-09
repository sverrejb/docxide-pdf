use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{
    ConnectorShape, DiagramAnchor, EmbeddedImage, FloatingImage, HRelativeFrom, HorizontalPosition,
    ImageFormat, ImageGlow, ImageReflection, ImageShadow, InlineChart, InnerShadow,
    SmartArtDiagram, SoftEdge, Textbox, VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use super::charts::parse_chart_from_zip;
use super::smartart::{has_diagram_ref, parse_smartart_drawing};
use super::textbox::{parse_connector_from_wsp, parse_textbox_from_wsp};
use super::{
    CHART_NS, DML_NS, PIC_NS, ParseContext, REL_NS, VML_NS, W10_NS, WML_NS, WPD_NS, angle_attr,
    dml, emu_attr, emu_attr_opt, emu_to_pts, f32_attr, frac_attr, parse_hex_color, parse_on_off,
    parse_pt, part_path, read_zip_bytes, twips_attr, wml, wpd,
};

fn parse_emu_text(text: &str) -> f32 {
    emu_to_pts(text.parse::<f32>().unwrap_or(0.0))
}

fn wpd_child_text<'a>(parent: Option<roxmltree::Node<'a, 'a>>, name: &str) -> Option<&'a str> {
    parent
        .and_then(|n| n.children().find(|c| c.tag_name().name() == name))
        .and_then(|n| n.text())
}

/// Extra vertical space for an inline image: (total, top_portion).
/// Total = effectExtent top+bottom + distT+distB.
/// Top = effectExtent.t + distT.
fn inline_extra_height(container: roxmltree::Node) -> (f32, f32) {
    let ee = wpd(container, "effectExtent");
    let ee_t = ee.map(|n| emu_attr(n, "t")).unwrap_or(0.0);
    let ee_b = ee.map(|n| emu_attr(n, "b")).unwrap_or(0.0);
    let dist_t = emu_attr(container, "distT");
    let dist_b = emu_attr(container, "distB");
    (ee_t + ee_b + dist_t + dist_b, ee_t + dist_t)
}

/// An anchor's (top, bottom) wrap distances: distT/distB plus the effectExtent
/// its effects add. indonesian's "Format 12" box (effectExtent b=23495) pushes
/// the next paragraph 1.85pt further than distB alone; indigenous_innovation's
/// three boxes 1.05pt each.
pub(super) fn wrap_dist_top_bottom(container: roxmltree::Node) -> (f32, f32) {
    let ee = wpd(container, "effectExtent");
    let ee_t = ee.map(|n| emu_attr(n, "t")).unwrap_or(0.0);
    let ee_b = ee.map(|n| emu_attr(n, "b")).unwrap_or(0.0);
    (
        emu_attr(container, "distT") + ee_t,
        emu_attr(container, "distB") + ee_b,
    )
}

pub(super) fn extent_dimensions(container: roxmltree::Node) -> (f32, f32) {
    wpd(container, "extent").map_or((0.0, 0.0), |e| (emu_attr(e, "cx"), emu_attr(e, "cy")))
}

/// `wp:anchor` z-order: (behindDoc, relativeHeight).
pub(super) fn anchor_z_order(anchor: roxmltree::Node) -> (bool, u32) {
    (
        anchor.attribute("behindDoc") == Some("1"),
        anchor
            .attribute("relativeHeight")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0),
    )
}

/// GIF and TIFF pictures are re-encoded as PNG for the PDF writer.
fn gif_or_tiff_to_png(data: &[u8]) -> Option<Vec<u8>> {
    let fmt = match image::guess_format(data).ok()? {
        f @ (image::ImageFormat::Gif | image::ImageFormat::Tiff) => f,
        _ => return None,
    };
    let img = image::load_from_memory_with_format(data, fmt).ok()?;
    let mut png = Vec::new();
    img.write_to(&mut std::io::Cursor::new(&mut png), image::ImageFormat::Png)
        .ok()?;
    Some(png)
}

pub(super) fn image_dimensions(data: &[u8]) -> Option<(u32, u32, ImageFormat, u8)> {
    if data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8 {
        return parse_jpeg_dimensions(data);
    }

    if data.len() >= 24 && data[0..4] == [0x89, 0x50, 0x4E, 0x47] {
        let width = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
        let height = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
        return Some((width, height, ImageFormat::Png, 3));
    }

    if data.len() >= 26 && data[0] == b'B' && data[1] == b'M' {
        let width = u32::from_le_bytes([data[18], data[19], data[20], data[21]]);
        // A negative height marks a top-down DIB (common in EMF-wrapped bitmaps).
        let height = i32::from_le_bytes([data[22], data[23], data[24], data[25]]).unsigned_abs();
        return Some((width, height, ImageFormat::Bmp, 3));
    }

    if super::emf::is_emf(data) {
        let (bw, bh) = super::emf::parse_header(data)?.bounds_size();
        // Pixel dimensions don't strictly apply to vector EMFs — use the
        // bounds-rect logical size so downstream code that expects them gets
        // a sensible aspect ratio.
        let pw = bw.max(1) as u32;
        let ph = bh.max(1) as u32;
        return Some((pw, ph, ImageFormat::Emf, 0));
    }

    None
}

fn parse_jpeg_dimensions(data: &[u8]) -> Option<(u32, u32, ImageFormat, u8)> {
    let mut i = 2;
    while i + 4 < data.len() {
        if data[i] != 0xFF {
            return None;
        }
        let marker = data[i + 1];
        if marker == 0xD9 {
            break;
        }
        let len = u16::from_be_bytes([data[i + 2], data[i + 3]]) as usize;
        if matches!(marker, 0xC0..=0xC2) && i + 9 < data.len() {
            let height = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
            let width = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
            let components = data[i + 9];
            return Some((width, height, ImageFormat::Jpeg, components));
        }
        i += 2 + len;
    }
    None
}

fn find_pic_sp_pr<'a>(container: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    container
        .descendants()
        .find(|n| n.has_tag_name((PIC_NS, "pic")))
        .and_then(|p| p.children().find(|c| c.has_tag_name((PIC_NS, "spPr"))))
}

/// Crop fractions (l, t, r, b) from the `a:srcRect` beside the picture's `a:blip`, each
/// stored as 1/100000 of the source dimension. None when absent, all zero, or when the
/// crop would leave nothing visible. Negative values are Word's outward crop and are
/// kept: they pad the frame with blank space.
fn parse_src_rect(container: roxmltree::Node) -> Option<[f32; 4]> {
    let rect = dml(find_blip(container)?.parent()?, "srcRect")?;
    let frac = |name| frac_attr(rect, name).unwrap_or(0.0);
    let r = [frac("l"), frac("t"), frac("r"), frac("b")];
    let visible_w = 1.0 - r[0] - r[2];
    let visible_h = 1.0 - r[1] - r[3];
    if r == [0.0; 4] || visible_w <= 0.0 || visible_h <= 0.0 {
        return None;
    }
    Some(r)
}

/// Apply the picture properties shared by inline and anchored pictures: rotation,
/// outline, effects, non-rectangular clip and `a:srcRect` crop.
fn apply_pic_props(img: &mut EmbeddedImage, container: roxmltree::Node) {
    if let Some(doc_pr) = wpd(container, "docPr") {
        img.alt = doc_pr
            .attribute("descr")
            .filter(|d| !d.trim().is_empty())
            .map(str::to_string);
        // Office 2019 "Mark as decorative": <adec:decorative val="1"/> in docPr's extLst.
        img.decorative = doc_pr.descendants().any(|n| {
            n.tag_name().name() == "decorative" && n.attribute("val").is_some_and(parse_on_off)
        });
    }
    let sp_pr = find_pic_sp_pr(container);
    img.rotation_deg = parse_image_rotation(sp_pr);
    let (stroke_color, stroke_width) = parse_pic_outline(sp_pr);
    img.stroke_color = stroke_color;
    img.stroke_width = stroke_width;
    let effects = parse_pic_effects(sp_pr);
    img.shadow = effects.shadow;
    img.soft_edge = effects.soft_edge;
    img.glow = effects.glow;
    img.inner_shadow = effects.inner_shadow;
    img.reflection = effects.reflection;
    img.clip_geometry = sp_pr
        .map(super::textbox::parse_shape_geometry)
        .filter(|g| g.preset.as_deref() != Some("rect") || g.custom.is_some());
    img.src_rect = parse_src_rect(container);
    img.lum = parse_lum(container);
}

/// `a:lum` brightness/contrast on the picture's blip, each stored as 1/1000 of a
/// percent. None when absent or both zero. Word applies these to the pixels, which
/// is how a +30%/+30% signature scan loses its faint background stamp
/// (italian_evaluation_minutes p7, annotation #229).
fn parse_lum(container: roxmltree::Node) -> Option<(f32, f32)> {
    let lum = dml(find_blip(container)?, "lum")?;
    let (bright, contrast) = (
        frac_attr(lum, "bright").unwrap_or(0.0),
        frac_attr(lum, "contrast").unwrap_or(0.0),
    );
    (bright != 0.0 || contrast != 0.0).then_some((bright, contrast))
}

/// Read in-plane rotation (clockwise degrees) for a floating picture. Prefers the
/// normal `a:xfrm @rot`; some files instead encode the turn via a 3D scene camera
/// `a:scene3d/a:camera/a:rot @rev` (e.g. a vertical label rotated 90°). Both are in
/// 60000ths of a degree, but @rev revolves the *camera*, so the object appears
/// rotated the opposite way — negate it to get the object rotation.
fn parse_image_rotation(sp_pr: Option<roxmltree::Node>) -> f32 {
    let Some(sp_pr) = sp_pr else {
        return 0.0;
    };
    if let Some(rot) = dml(sp_pr, "xfrm").and_then(|x| f32_attr(x, "rot"))
        && rot.abs() > f32::EPSILON
    {
        return rot / 60000.0;
    }
    dml(sp_pr, "scene3d")
        .and_then(|s| dml(s, "camera"))
        .and_then(|c| dml(c, "rot"))
        .and_then(|r| angle_attr(r, "rev"))
        .map(|rev| -rev)
        .unwrap_or(0.0)
}

/// Parse outline stroke from `pic:spPr/a:ln`.
fn parse_pic_outline(sp_pr: Option<roxmltree::Node>) -> (Option<[u8; 3]>, f32) {
    let ln = sp_pr.and_then(|s| s.children().find(|c| c.has_tag_name((DML_NS, "ln"))));
    let Some(ln) = ln else {
        return (None, 0.0);
    };
    let width = emu_attr_opt(ln, "w").unwrap_or(0.75); // default 0.75pt
    let color = ln
        .descendants()
        .find(|n| n.has_tag_name((DML_NS, "srgbClr")))
        .and_then(|n| n.attribute("val"))
        .and_then(parse_hex_color);
    if color.is_some() {
        (color, width)
    } else {
        (None, 0.0)
    }
}

/// Parse color + alpha from a DML element that has `a:srgbClr` or `a:schemeClr`
/// with optional `a:alpha` child.
fn parse_dml_color_alpha(node: roxmltree::Node) -> ([u8; 3], f32) {
    // Try srgbClr first, fall back to schemeClr (theme color — use black as fallback)
    let color_node = node.descendants().find(|n| {
        n.tag_name().namespace() == Some(DML_NS)
            && (n.tag_name().name() == "srgbClr" || n.tag_name().name() == "schemeClr")
    });
    let rgb = color_node
        .filter(|n| n.tag_name().name() == "srgbClr")
        .and_then(|n| n.attribute("val"))
        .and_then(parse_hex_color)
        .unwrap_or([0, 0, 0]);
    let alpha = color_node
        .and_then(|n| n.children().find(|c| c.has_tag_name((DML_NS, "alpha"))))
        .and_then(|a| frac_attr(a, "val"))
        .unwrap_or(1.0);
    (rgb, alpha)
}

/// Parse dist+dir attributes (common to outerShdw, innerShdw) into (offset_x, offset_y).
fn parse_dist_dir(node: roxmltree::Node) -> (f32, f32) {
    let dist = emu_attr(node, "dist");
    let dir_deg = angle_attr(node, "dir").unwrap_or(0.0);
    let dir_rad = dir_deg.to_radians();
    (dist * dir_rad.cos(), dist * dir_rad.sin())
}

struct PicEffects {
    shadow: Option<ImageShadow>,
    soft_edge: Option<SoftEdge>,
    glow: Option<ImageGlow>,
    inner_shadow: Option<InnerShadow>,
    reflection: Option<ImageReflection>,
}

/// Parse all picture effects from `pic:spPr/a:effectLst`.
fn parse_pic_effects(sp_pr: Option<roxmltree::Node>) -> PicEffects {
    let mut fx = PicEffects {
        shadow: None,
        soft_edge: None,
        glow: None,
        inner_shadow: None,
        reflection: None,
    };
    let Some(sp) = sp_pr else {
        return fx;
    };
    let Some(effect_lst) = sp
        .children()
        .find(|c| c.has_tag_name((DML_NS, "effectLst")))
    else {
        return fx;
    };

    for child in effect_lst
        .children()
        .filter(|c| c.tag_name().namespace() == Some(DML_NS))
    {
        match child.tag_name().name() {
            "outerShdw" => {
                let blur_radius = emu_attr_opt(child, "blurRad").unwrap_or(0.0);
                let (offset_x, offset_y) = parse_dist_dir(child);
                let (color, alpha) = parse_dml_color_alpha(child);
                fx.shadow = Some(ImageShadow {
                    offset_x,
                    offset_y,
                    blur_radius,
                    color,
                    alpha,
                });
            }
            "softEdge" => {
                let radius = emu_attr_opt(child, "rad").unwrap_or(0.0);
                if radius > 0.0 {
                    fx.soft_edge = Some(SoftEdge { radius });
                }
            }
            "glow" => {
                let radius = emu_attr_opt(child, "rad").unwrap_or(0.0);
                let (color, alpha) = parse_dml_color_alpha(child);
                if radius > 0.0 {
                    fx.glow = Some(ImageGlow {
                        radius,
                        color,
                        alpha,
                    });
                }
            }
            "innerShdw" => {
                let blur_radius = emu_attr_opt(child, "blurRad").unwrap_or(0.0);
                let (offset_x, offset_y) = parse_dist_dir(child);
                let (color, alpha) = parse_dml_color_alpha(child);
                fx.inner_shadow = Some(InnerShadow {
                    offset_x,
                    offset_y,
                    blur_radius,
                    color,
                    alpha,
                });
            }
            "reflection" => {
                let start_alpha = frac_attr(child, "stA").unwrap_or(0.5);
                let end_alpha = frac_attr(child, "endA").unwrap_or(0.0);
                let distance = emu_attr_opt(child, "dist").unwrap_or(0.0);
                let end_pos = frac_attr(child, "endPos").unwrap_or(1.0);
                fx.reflection = Some(ImageReflection {
                    start_alpha,
                    end_alpha,
                    distance,
                    end_pos,
                });
            }
            _ => {}
        }
    }
    fx
}

pub(super) fn read_image_from_zip<R: Read + Seek>(
    embed_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    display_w: f32,
    display_h: f32,
) -> Option<EmbeddedImage> {
    read_image_from_zip_extra(embed_id, rels, zip, display_w, display_h, 0.0, 0.0)
}

pub(super) fn read_image_from_zip_extra<R: Read + Seek>(
    embed_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    display_w: f32,
    display_h: f32,
    layout_extra_height: f32,
    layout_extra_top: f32,
) -> Option<EmbeddedImage> {
    let mut data = read_zip_bytes(zip, &part_path(rels.get(embed_id)?))?;
    if super::wmf::is_wmf(&data) {
        data = match super::wmf::embedded_emf(&data) {
            Some(emf) => emf,
            None => super::wmf::wmf_to_raster(&data)?,
        };
    }
    if let Some(bmp) = super::emf::emf_to_raster(&data) {
        data = bmp;
    } else if let Some(png) = gif_or_tiff_to_png(&data) {
        data = png;
    }
    let (pw, ph, fmt, components) = image_dimensions(&data)?;
    Some(EmbeddedImage {
        is_ole_preview: false,
        alt: None,
        decorative: false,
        data: std::sync::Arc::new(data),
        format: fmt,
        pixel_width: pw,
        pixel_height: ph,
        display_width: display_w,
        display_height: display_h,
        jpeg_components: components,
        layout_extra_height,
        layout_extra_top,
        rotation_deg: 0.0,
        stroke_color: None,
        stroke_width: 0.0,
        shadow: None,
        soft_edge: None,
        glow: None,
        inner_shadow: None,
        reflection: None,
        clip_geometry: None,
        src_rect: None,
        lum: None,
    })
}

fn find_blip<'a>(container: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    container
        .descendants()
        .find(|n| n.has_tag_name((DML_NS, "blip")))
}

pub(super) fn find_blip_embed<'a>(container: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    find_blip(container)?.attribute((REL_NS, "embed"))
}

pub(super) struct DrawingInfo {
    pub(super) height: f32,
    pub(super) image: Option<EmbeddedImage>,
}

pub(super) fn parse_anchor_position(
    container: roxmltree::Node,
) -> (
    HorizontalPosition,
    HRelativeFrom,
    VerticalPosition,
    VRelativeFrom,
) {
    let pos_h = wpd(container, "positionH");
    let h_relative = match pos_h.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => HRelativeFrom::Page,
        Some("margin") => HRelativeFrom::Margin,
        _ => HRelativeFrom::Column,
    };
    let h_position = if let Some(text) = wpd_child_text(pos_h, "align") {
        match text {
            "center" => HorizontalPosition::AlignCenter,
            "right" => HorizontalPosition::AlignRight,
            _ => HorizontalPosition::AlignLeft,
        }
    } else if let Some(text) = wpd_child_text(pos_h, "posOffset") {
        HorizontalPosition::Offset(parse_emu_text(text))
    } else {
        HorizontalPosition::AlignLeft
    };

    let pos_v = wpd(container, "positionV");
    let v_relative = match pos_v.and_then(|n| n.attribute("relativeFrom")) {
        Some("page") => VRelativeFrom::Page,
        Some("margin") => VRelativeFrom::Margin,
        Some("topMargin") => VRelativeFrom::TopMargin,
        _ => VRelativeFrom::Paragraph,
    };
    let v_position = if let Some(text) = wpd_child_text(pos_v, "align") {
        match text {
            "bottom" => VerticalPosition::AlignBottom,
            "center" => VerticalPosition::AlignCenter,
            _ => VerticalPosition::AlignTop,
        }
    } else if let Some(text) = wpd_child_text(pos_v, "posOffset") {
        VerticalPosition::Offset(parse_emu_text(text))
    } else {
        VerticalPosition::Offset(0.0)
    };

    (h_position, h_relative, v_position, v_relative)
}

pub(super) fn parse_wrap_type(
    container: roxmltree::Node,
) -> (WrapType, WrapText, Option<Vec<(i32, i32)>>) {
    for child in container.children() {
        if child.tag_name().namespace() != Some(WPD_NS) {
            continue;
        }
        let wrap_text = match child.attribute("wrapText") {
            Some("bothSides") => WrapText::BothSides,
            Some("left") => WrapText::Left,
            Some("right") => WrapText::Right,
            Some("largest") => WrapText::Largest,
            _ => WrapText::BothSides,
        };
        match child.tag_name().name() {
            "wrapSquare" => return (WrapType::Square, wrap_text, None),
            "wrapTight" | "wrapThrough" => {
                let polygon = parse_wrap_polygon(child);
                let wt = if child.tag_name().name() == "wrapTight" {
                    WrapType::Tight
                } else {
                    WrapType::Through
                };
                return (wt, wrap_text, polygon);
            }
            "wrapTopAndBottom" => return (WrapType::TopAndBottom, WrapText::BothSides, None),
            "wrapNone" => return (WrapType::None, WrapText::BothSides, None),
            _ => {}
        }
    }
    (WrapType::None, WrapText::BothSides, None)
}

fn parse_wrap_polygon(wrap_elem: roxmltree::Node) -> Option<Vec<(i32, i32)>> {
    let poly = wrap_elem
        .children()
        .find(|c| c.has_tag_name((WPD_NS, "wrapPolygon")))?;
    let mut vertices = Vec::new();
    for child in poly.children() {
        if child.tag_name().namespace() != Some(WPD_NS) {
            continue;
        }
        match child.tag_name().name() {
            "start" | "lineTo" => {
                let x = child
                    .attribute("x")
                    .and_then(|v| v.parse::<i32>().ok())
                    .unwrap_or(0);
                let y = child
                    .attribute("y")
                    .and_then(|v| v.parse::<i32>().ok())
                    .unwrap_or(0);
                vertices.push((x, y));
            }
            _ => {}
        }
    }
    if vertices.is_empty() {
        None
    } else {
        Some(vertices)
    }
}

pub(super) enum RunDrawingResult {
    Inline(EmbeddedImage),
    Floating(FloatingImage),
    TextBox(Textbox),
    Connector(ConnectorShape),
    Chart(InlineChart),
    SmartArt(SmartArtDiagram),
    /// Flattened canvas/group drawing: multiple leaf shapes from one w:drawing
    Group(Vec<RunDrawingResult>),
}

pub(super) fn parse_run_drawing<R: Read + Seek>(
    drawing_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<RunDrawingResult> {
    for container in drawing_node.children() {
        let is_inline = container.has_tag_name((WPD_NS, "inline"));
        let is_anchor = container.has_tag_name((WPD_NS, "anchor"));
        if !is_inline && !is_anchor {
            continue;
        }

        // Word does not print drawings flagged hidden="1" on wp:docPr (e.g.
        // template logos or document-management metadata shapes); skip them.
        if wpd(container, "docPr")
            .and_then(|n| n.attribute("hidden"))
            .is_some_and(parse_on_off)
        {
            continue;
        }

        let (display_w, display_h) = extent_dimensions(container);

        // Canvas/group drawings hold multiple shapes; flatten them before the
        // single-shape paths below grab just the first wsp descendant.
        if let Some(items) = super::group::parse_canvas_or_group(container, is_anchor, ctx) {
            return Some(RunDrawingResult::Group(items));
        }

        if is_anchor {
            if let Some(wsp) = parse_textbox_from_wsp(container, ctx) {
                let (h_position, h_relative, v_pos, v_relative) = parse_anchor_position(container);
                let (wrap_type, wrap_text, _) = parse_wrap_type(container);
                let (behind_doc, z_index) = anchor_z_order(container);
                return Some(RunDrawingResult::TextBox(Textbox {
                    width_pt: display_w,
                    height_pt: display_h,
                    h_position,
                    h_relative_from: h_relative,
                    v_offset_pt: v_pos.offset_or_zero(),
                    v_position: v_pos,
                    v_relative_from: v_relative,
                    wrap_type,
                    wrap_text,
                    dist_bottom: wrap_dist_top_bottom(container).1,
                    behind_doc,
                    z_index,
                    ..Textbox::from(wsp)
                }));
            }
            if let Some(conn) = parse_connector_from_wsp(container, ctx.theme) {
                return Some(RunDrawingResult::Connector(conn));
            }
            if let Some(embed_id) = find_blip_embed(container)
                && let Some(mut img) =
                    read_image_from_zip(embed_id, ctx.rels, ctx.zip, display_w, display_h)
            {
                apply_pic_props(&mut img, container);
                let (h_position, h_relative, v_position, v_relative) =
                    parse_anchor_position(container);
                let (wrap_type, wrap_text, wrap_polygon) = parse_wrap_type(container);
                let (behind_doc, z_index) = anchor_z_order(container);
                let (dist_top, dist_bottom) = wrap_dist_top_bottom(container);
                return Some(RunDrawingResult::Floating(FloatingImage {
                    image: img,
                    h_position,
                    h_relative_from: h_relative,
                    v_position,
                    v_relative_from: v_relative,
                    wrap_type,
                    wrap_text,
                    wrap_polygon,
                    behind_doc,
                    dist_top,
                    dist_bottom,
                    dist_left: emu_attr(container, "distL"),
                    dist_right: emu_attr(container, "distR"),
                    z_index,
                    anchor_seq: 0,
                }));
            }
            // A wrapNone diagram floats at its anchor, the paragraph's text laid
            // out as if it weren't there.
            // ponytail: other wraps are laid out inline; add wrap zones with a
            // fixture that has one.
            if display_h > 0.0 && has_diagram_ref(container) {
                let mut diagram =
                    parse_smartart_drawing(container, ctx.rels, ctx.zip, ctx.theme, display_h);
                if parse_wrap_type(container).0 == WrapType::None {
                    let (h_position, h_relative_from, v_position, v_relative_from) =
                        parse_anchor_position(container);
                    diagram.anchor = Some(DiagramAnchor {
                        h_position,
                        h_relative_from,
                        v_position,
                        v_relative_from,
                        width: display_w,
                        z_index: anchor_z_order(container).1,
                    });
                }
                return Some(RunDrawingResult::SmartArt(diagram));
            }
            continue;
        }

        // Inline textbox: wp:inline containing wps:wsp with text content
        if let Some(wsp) = parse_textbox_from_wsp(container, ctx) {
            // Treat inline textbox as a floating textbox at paragraph position
            // with TopAndBottom wrap so it acts as a block element
            return Some(RunDrawingResult::TextBox(Textbox {
                width_pt: display_w,
                height_pt: display_h,
                wrap_type: WrapType::TopAndBottom,
                ..Textbox::from(wsp)
            }));
        }

        if let Some(embed_id) = find_blip_embed(container) {
            let (extra_h, extra_top) = inline_extra_height(container);
            if let Some(mut img) = read_image_from_zip_extra(
                embed_id, ctx.rels, ctx.zip, display_w, display_h, extra_h, extra_top,
            ) {
                apply_pic_props(&mut img, container);
                return Some(RunDrawingResult::Inline(img));
            }
        }

        if let Some(chart_rid) = find_chart_ref(container) {
            let accent_colors: Vec<[u8; 3]> = (1..=6)
                .filter_map(|i| ctx.theme.colors.get(&format!("accent{i}")).copied())
                .collect();
            if let Some(ic) = parse_chart_from_zip(
                chart_rid,
                ctx.rels,
                ctx.zip,
                display_w,
                display_h,
                accent_colors,
            ) {
                return Some(RunDrawingResult::Chart(ic));
            }
        }

        if display_h > 0.0 && has_diagram_ref(container) {
            let diagram =
                parse_smartart_drawing(container, ctx.rels, ctx.zip, ctx.theme, display_h);
            return Some(RunDrawingResult::SmartArt(diagram));
        }
    }
    None
}

fn find_chart_ref<'a>(container: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    container
        .descendants()
        .find(|n| n.has_tag_name((DML_NS, "graphicData")) && n.attribute("uri") == Some(CHART_NS))
        .and_then(|gd| {
            gd.children()
                .find(|n| n.tag_name().name() == "chart")
                .and_then(|c| c.attribute((REL_NS, "id")))
        })
}

pub(super) fn compute_drawing_info<R: Read + Seek>(
    para_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> DrawingInfo {
    let mut max_height: f32 = 0.0;
    let mut image: Option<EmbeddedImage> = None;

    for child in para_node.children() {
        let is_wml = child.tag_name().namespace() == Some(WML_NS);
        let drawing_node = match child.tag_name().name() {
            "drawing" if is_wml => Some(child),
            "r" if is_wml => wml(child, "drawing"),
            _ => None,
        };

        let Some(drawing) = drawing_node else {
            continue;
        };
        for container in drawing.children() {
            if !container.has_tag_name((WPD_NS, "inline")) {
                continue;
            }

            let (display_w, display_h) = extent_dimensions(container);
            let (extra_h, extra_top) = inline_extra_height(container);
            max_height = max_height.max(display_h + extra_h);

            if image.is_none()
                && let Some(embed_id) = find_blip_embed(container)
            {
                image = read_image_from_zip_extra(
                    embed_id, rels, zip, display_w, display_h, extra_h, extra_top,
                );
                if let Some(ref mut img) = image {
                    apply_pic_props(img, container);
                }
            }
        }
    }
    let object_height = compute_object_height(para_node);
    DrawingInfo {
        height: max_height.max(object_height),
        image,
    }
}

/// Extract an inline image from a `<w:object>` VML fallback.
///
/// Word's OLE static-metafile embeddings use a `<v:imagedata r:id="..."/>`
/// child to point at a WMF/EMF/raster fallback. For WMFs whose content is one
/// or more DIB-bearing records we can rasterize them in `wmf::wmf_to_raster`
/// and render normally; this function finds the imagedata reference and
/// pulls the image with the display dimensions declared on the surrounding
/// VML shape (or on the `<w:object>` itself as a fallback).
pub(super) fn parse_object_inline_image<R: Read + Seek>(
    obj: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<EmbeddedImage> {
    let imagedata = obj
        .descendants()
        .find(|n| n.has_tag_name((VML_NS, "imagedata")))?;
    let embed_id = imagedata.attribute((REL_NS, "id"))?;
    let (w, h) = object_dimensions(obj)?;
    let mut img = read_image_from_zip(embed_id, ctx.rels, ctx.zip, w, h)?;
    (img.src_rect, img.lum) = vml_imagedata_props(imagedata);
    Some(with_object_alt(obj, img))
}

/// A legacy VML picture: a `w:pict` whose shape shows only an image (no text
/// box). Laid out like an OLE object's preview picture.
pub(super) fn is_vml_picture(pict: roxmltree::Node) -> bool {
    pict.children().any(|n| {
        n.tag_name().namespace() == Some(VML_NS)
            && matches!(n.tag_name().name(), "shape" | "rect")
            && n.children().any(|c| c.has_tag_name((VML_NS, "imagedata")))
            && !n.children().any(|c| c.has_tag_name((VML_NS, "textbox")))
    })
}

/// VML fractions are plain decimals or 16.16 fixed point with an `f` suffix
/// ("19661f" = 0.3).
fn vml_fraction(v: &str) -> Option<f32> {
    match v.strip_suffix('f') {
        Some(fixed) => fixed.parse::<f32>().ok().map(|f| f / 65536.0),
        None => v.parse().ok(),
    }
}

/// `v:imagedata` crop and colour adjustments in DrawingML terms. Word writes its
/// Washout preset (`a:lum bright=70% contrast=-70%`) as gain 0.3 / blacklevel
/// 0.35, so contrast = gain − 1 and brightness = 2 × blacklevel.
fn vml_imagedata_props(imagedata: roxmltree::Node) -> (Option<[f32; 4]>, Option<(f32, f32)>) {
    let frac = |name| imagedata.attribute(name).and_then(vml_fraction);
    let crop = [
        frac("cropleft").unwrap_or(0.0),
        frac("croptop").unwrap_or(0.0),
        frac("cropright").unwrap_or(0.0),
        frac("cropbottom").unwrap_or(0.0),
    ];
    let src_rect =
        (crop != [0.0; 4] && crop[0] + crop[2] < 1.0 && crop[1] + crop[3] < 1.0).then_some(crop);
    let contrast = frac("gain").map_or(0.0, |g| g - 1.0);
    let bright = frac("blacklevel").map_or(0.0, |b| 2.0 * b);
    let lum = (bright != 0.0 || contrast != 0.0).then_some((bright, contrast));
    (src_rect, lum)
}

fn has_ole_object(obj: roxmltree::Node) -> bool {
    obj.children()
        .any(|n| n.has_tag_name(("urn:schemas-microsoft-com:office:office", "OLEObject")))
}

/// Word tags an OLE object as a Sect in its paragraph, with the VML shape's alt
/// text or a blank one. One with alt text becomes a Figure with it; one without
/// an artifact rather than a Figure with nothing to say (7.3-1).
fn with_object_alt(obj: roxmltree::Node, mut img: EmbeddedImage) -> EmbeddedImage {
    img.is_ole_preview = has_ole_object(obj);
    img.alt = obj
        .children()
        .filter(|n| n.tag_name().namespace() == Some(VML_NS))
        .find_map(|n| n.attribute("alt"))
        .filter(|a| !a.trim().is_empty())
        .map(str::to_string);
    img.decorative = img.alt.is_none();
    img
}

/// Extract a *floating* image from a `<w:object>` whose VML shape is absolutely
/// positioned (`style="position:absolute;margin-left:..;margin-top:.."`). Word
/// uses this for OLE-embedded logos placed at a fixed page location (e.g. a
/// right-anchored faculty emblem). Without this they fall to the inline path and
/// get dumped into the centered header paragraph, pushing the heading down.
/// Returns None for inline objects (no `position:absolute`), which keep the
/// existing inline behavior.
pub(super) fn parse_object_floating_image<R: Read + Seek>(
    obj: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<FloatingImage> {
    let shape = obj.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS)
            && matches!(n.tag_name().name(), "rect" | "shape" | "oval" | "roundrect")
    })?;
    let style = shape.attribute("style")?;
    if !style.contains("position:absolute") {
        return None;
    }
    let mut margin_left = 0.0_f32;
    let mut margin_top = 0.0_f32;
    let mut h_align = None;
    let mut v_align = None;
    let mut z_index: i64 = 0;
    let mut h_relative = HRelativeFrom::Column;
    let mut v_relative = VRelativeFrom::Paragraph;
    for part in style.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            let val = val.trim();
            match key.trim() {
                "margin-left" => margin_left = parse_pt(val).unwrap_or(0.0),
                "margin-top" => margin_top = parse_pt(val).unwrap_or(0.0),
                "z-index" => z_index = val.parse().unwrap_or(0),
                // "absolute" (or no value) means the margin offsets place it.
                "mso-position-horizontal" => {
                    h_align = match val {
                        "left" | "inside" => Some(HorizontalPosition::AlignLeft),
                        "center" => Some(HorizontalPosition::AlignCenter),
                        "right" | "outside" => Some(HorizontalPosition::AlignRight),
                        _ => None,
                    };
                }
                "mso-position-vertical" => {
                    v_align = match val {
                        "top" | "inside" => Some(VerticalPosition::AlignTop),
                        "center" => Some(VerticalPosition::AlignCenter),
                        "bottom" | "outside" => Some(VerticalPosition::AlignBottom),
                        _ => None,
                    };
                }
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
                        "top-margin-area" => VRelativeFrom::TopMargin,
                        _ => VRelativeFrom::Paragraph,
                    };
                }
                _ => {}
            }
        }
    }
    let image = parse_object_inline_image(obj, ctx)?;
    // An explicit <w10:wrap type="square"/> means the object reflows text
    // (Word wraps centered header text between such logos); without it the
    // logo sits over/beside the text and None keeps the text full-width.
    let wrap_type = shape
        .children()
        .find(|n| n.has_tag_name((W10_NS, "wrap")))
        .and_then(|n| n.attribute("type"))
        .map(|t| match t {
            "square" => WrapType::Square,
            "tight" => WrapType::Tight,
            "through" => WrapType::Through,
            "topAndBottom" => WrapType::TopAndBottom,
            _ => WrapType::None,
        })
        .unwrap_or(WrapType::None);
    Some(FloatingImage {
        image,
        h_position: h_align.unwrap_or(HorizontalPosition::Offset(margin_left)),
        h_relative_from: h_relative,
        v_position: v_align.unwrap_or(VerticalPosition::Offset(margin_top)),
        v_relative_from: v_relative,
        wrap_type,
        wrap_text: WrapText::BothSides,
        wrap_polygon: None,
        // VML stacks by z-index; a negative one is behind the text.
        behind_doc: z_index < 0,
        dist_top: 0.0,
        dist_bottom: 0.0,
        dist_left: 0.0,
        dist_right: 0.0,
        // Word writes the DrawingML relativeHeight here, negated when behind.
        z_index: z_index.unsigned_abs().min(u32::MAX as u64) as u32,
        anchor_seq: 0,
    })
}

/// Reserve space for `<w:object>` legacy OLE embeddings (we don't render the
/// content, but the following paragraphs need to land at the right position).
pub(super) fn compute_object_height(para_node: roxmltree::Node) -> f32 {
    let mut max_height: f32 = 0.0;
    for r in para_node
        .children()
        .filter(|n| n.has_tag_name((WML_NS, "r")))
    {
        for obj in r.children().filter(|n| n.has_tag_name((WML_NS, "object"))) {
            // Absolutely-positioned objects float (see parse_object_floating_image)
            // and must not reserve inline line height.
            if object_is_absolute(obj) {
                continue;
            }
            if let Some((_w, h)) = object_dimensions(obj) {
                max_height = max_height.max(h);
            }
        }
    }
    max_height
}

fn object_is_absolute(obj: roxmltree::Node) -> bool {
    obj.children()
        .find(|n| {
            n.tag_name().namespace() == Some(VML_NS)
                && matches!(n.tag_name().name(), "rect" | "shape" | "oval" | "roundrect")
        })
        .and_then(|s| s.attribute("style"))
        .is_some_and(|style| style.contains("position:absolute"))
}

fn object_dimensions(obj: roxmltree::Node) -> Option<(f32, f32)> {
    let rect = obj.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS)
            && matches!(n.tag_name().name(), "rect" | "shape" | "oval" | "roundrect")
    });
    if let Some(rect) = rect
        && let Some(style) = rect.attribute("style")
    {
        let mut w_pt: Option<f32> = None;
        let mut h_pt: Option<f32> = None;
        for part in style.split(';') {
            if let Some((key, val)) = part.trim().split_once(':') {
                let key = key.trim();
                if let Some(v) = parse_pt(val) {
                    match key {
                        "width" => w_pt = Some(v),
                        "height" => h_pt = Some(v),
                        _ => {}
                    }
                }
            }
        }
        if let (Some(w), Some(h)) = (w_pt, h_pt) {
            return Some((w, h));
        }
    }
    let dxa = twips_attr(obj, "dxaOrig");
    let dya = twips_attr(obj, "dyaOrig");
    if let (Some(w), Some(h)) = (dxa, dya) {
        return Some((w, h));
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vml_washout_maps_to_word_lum_preset() {
        let xml = format!(
            r#"<v:imagedata xmlns:v="{VML_NS}" gain="19661f" blacklevel="22938f" cropleft="0.25" cropbottom="6554f"/>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let (src_rect, lum) = vml_imagedata_props(doc.root_element());
        let (bright, contrast) = lum.unwrap();
        assert!((bright - 0.70).abs() < 0.001 && (contrast + 0.70).abs() < 0.001);
        let r = src_rect.unwrap();
        assert!((r[0] - 0.25).abs() < 1e-6 && (r[3] - 0.1).abs() < 1e-4 && r[1] == 0.0);
    }

    #[test]
    fn ole_preview_requires_office_object_not_just_vml_shape() {
        for (child, expected) in [
            (r#"<v:shape><v:imagedata r:id="image"/></v:shape>"#, false),
            (r#"<v:shape/><o:OLEObject Type="Embed"/>"#, true),
            (
                r#"<v:shape/><alias:OLEObject xmlns:alias="urn:schemas-microsoft-com:office:office"/>"#,
                true,
            ),
            (
                r#"<v:shape/><wrong:OLEObject xmlns:wrong="urn:other"/>"#,
                false,
            ),
        ] {
            let xml = format!(
                r#"<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="{VML_NS}" xmlns:r="{REL_NS}" xmlns:o="urn:schemas-microsoft-com:office:office">{child}</w:object>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            assert_eq!(has_ole_object(doc.root_element()), expected, "{child}");
        }
    }

    fn src_rect_of(elem: &str) -> Option<[f32; 4]> {
        let xml = format!(
            r#"<root xmlns:a="{DML_NS}" xmlns:r="{REL_NS}"><a:blipFill><a:blip r:embed="rId1"/>{elem}<a:stretch><a:fillRect/></a:stretch></a:blipFill></root>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        parse_src_rect(doc.root_element())
    }

    fn assert_close(got: [f32; 4], want: [f32; 4]) {
        for (g, w) in got.iter().zip(want) {
            assert!((g - w).abs() < 1e-6, "{got:?} != {want:?}");
        }
    }

    #[test]
    fn lum_is_fraction_pair_or_none() {
        let lum_of = |blip_children: &str| {
            let xml = format!(
                r#"<root xmlns:a="{DML_NS}" xmlns:r="{REL_NS}"><a:blipFill><a:blip r:embed="rId1">{blip_children}</a:blip></a:blipFill></root>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_lum(doc.root_element())
        };
        assert_eq!(lum_of(""), None);
        assert_eq!(lum_of(r#"<a:lum bright="0" contrast="0"/>"#), None);
        let (b, c) = lum_of(r#"<a:lum bright="30000" contrast="30000"/>"#).unwrap();
        assert!((b - 0.3).abs() < 1e-6 && (c - 0.3).abs() < 1e-6);
        let (b, c) = lum_of(r#"<a:lum bright="-20000"/>"#).unwrap();
        assert!((b + 0.2).abs() < 1e-6 && c == 0.0);
    }

    #[test]
    fn src_rect_absent_empty_or_zero_is_none() {
        assert_eq!(src_rect_of(""), None);
        assert_eq!(src_rect_of("<a:srcRect/>"), None);
        assert_eq!(src_rect_of(r#"<a:srcRect l="0" t="0" r="0" b="0"/>"#), None);
    }

    #[test]
    fn src_rect_values_are_fractions_of_100000() {
        // Real crops from brazilian_logistics_study; missing attributes read as zero.
        assert_close(
            src_rect_of(r#"<a:srcRect t="4604" r="1295" b="6879"/>"#).unwrap(),
            [0.0, 0.04604, 0.01295, 0.06879],
        );
        assert_close(
            src_rect_of(r#"<a:srcRect t="4505" b="26576"/>"#).unwrap(),
            [0.0, 0.04505, 0.0, 0.26576],
        );
    }

    #[test]
    fn src_rect_negative_outward_crop_is_kept() {
        assert_close(
            src_rect_of(r#"<a:srcRect l="-20000"/>"#).unwrap(),
            [-0.2, 0.0, 0.0, 0.0],
        );
    }

    #[test]
    fn src_rect_leaving_nothing_visible_is_none() {
        assert_eq!(src_rect_of(r#"<a:srcRect l="60000" r="50000"/>"#), None);
        assert_eq!(src_rect_of(r#"<a:srcRect t="100000"/>"#), None);
    }
}
