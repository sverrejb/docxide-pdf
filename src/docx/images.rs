use std::io::{Read, Seek};

use super::styles::ThemeFonts;

use crate::model::{
    ConnectorShape, DiagramAnchor, EmbeddedImage, FloatingImage, HRelativeFrom, HorizontalPosition,
    ImageFormat, ImageGlow, ImageReflection, ImageShadow, InlineChart, InnerShadow,
    SmartArtDiagram, SoftEdge, Textbox, VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use super::charts::parse_chart_from_zip;
use super::media::{find_blip, find_blip_embeds, read_image_from_zip, tile_picture};
use super::smartart::{has_diagram_ref, parse_smartart_drawing};
use super::textbox::{parse_connector_from_wsp, parse_textbox_from_wsp};
use super::vml_object::compute_object_height;
use super::{
    CHART_NS, DML_NS, PIC_NS, ParseContext, REL_NS, WML_NS, WPD_NS, WPS_NS, angle_attr, dml,
    emu_attr, emu_attr_opt, emu_to_pts, f32_attr, frac_attr, parse_hex_color, parse_on_off,
    twips_to_pts, wml, wpd,
};

fn parse_emu_text(text: &str) -> f32 {
    emu_to_pts(text.parse::<f32>().unwrap_or(0.0))
}

fn wpd_child_text<'a>(parent: Option<roxmltree::Node<'a, 'a>>, name: &str) -> Option<&'a str> {
    parent
        .and_then(|n| n.children().find(|c| c.tag_name().name() == name))
        .and_then(|n| n.text())
}

/// Extra vertical space for an inline image: (total, top_portion), its
/// effectExtent top+bottom and top. distT/distB only space a floating
/// picture's wrap: Word lays an inline one out without them (a 9pt-dist
/// banner sits 150.25pt down in Word's export, 9pt higher than with them).
pub(super) fn inline_extra_height(container: roxmltree::Node) -> (f32, f32) {
    let ee = wpd(container, "effectExtent");
    let ee_t = ee.map(|n| emu_attr(n, "t")).unwrap_or(0.0);
    let ee_b = ee.map(|n| emu_attr(n, "b")).unwrap_or(0.0);
    (ee_t + ee_b, ee_t)
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

/// An anchor's (left, right) wrap distances: distL/distR plus, unless a wrap
/// polygon gives the outline, the effectExtent its effects add (Word examples:
/// a square box with effectExtent r=0.8pt puts the text beside it 0.8pt
/// further right; a tight one's polygon at the extent edge does not).
pub(super) fn wrap_dist_left_right(container: roxmltree::Node, polygon: bool) -> (f32, f32) {
    let ee = wpd(container, "effectExtent").filter(|_| !polygon);
    let ee_l = ee.map(|n| emu_attr(n, "l")).unwrap_or(0.0);
    let ee_r = ee.map(|n| emu_attr(n, "r")).unwrap_or(0.0);
    (
        emu_attr(container, "distL") + ee_l,
        emu_attr(container, "distR") + ee_r,
    )
}

pub(super) fn extent_dimensions(container: roxmltree::Node) -> (f32, f32) {
    wpd(container, "extent").map_or((0.0, 0.0), |e| (emu_attr(e, "cx"), emu_attr(e, "cy")))
}

/// An inline picture's size as Word lays it out: the height cut to whole
/// twips, the width scaled with it (Word for Mac probes: extents of 2493.2 to
/// 2493.9 twips all step 124.65pt lines and draw 124.65pt tall, a 72pt wide
/// one 71.98pt wide; 1000.75 twips step 50pt).
fn inline_picture_extent(w: f32, h: f32) -> (f32, f32) {
    let cut = twips_to_pts((h * 20.0 + 1e-3).floor());
    if cut <= 0.0 {
        return (w, h);
    }
    (w * cut / h, cut)
}

/// `wp14:sizeRelH relativeFrom="margin"` with a non-zero pctWidth, as a
/// fraction of the text width.
fn pct_width_of_margin(anchor: roxmltree::Node) -> Option<f32> {
    let rel = anchor
        .children()
        .find(|n| n.tag_name().name() == "sizeRelH")
        .filter(|n| n.attribute("relativeFrom") == Some("margin"))?;
    let pct: f32 = rel
        .children()
        .find(|n| n.tag_name().name() == "pctWidth")?
        .text()?
        .trim()
        .parse()
        .ok()?;
    (pct > 0.0).then_some(pct / 100_000.0)
}

/// `wp:anchor` z-order: (behindDoc, relativeHeight).
pub(super) fn anchor_z_order(anchor: roxmltree::Node) -> (bool, u32) {
    (
        anchor
            .attribute("behindDoc")
            .is_some_and(super::parse_on_off),
        anchor
            .attribute("relativeHeight")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0),
    )
}

/// An inline `wps` shape filled with a picture and holding no text: a picture
/// in the line like any other, keeping its place after the pictures before
/// it, with the shape's own properties. An `a:tile` fill is drawn into one
/// raster.
pub(super) fn picture_filled_shape<R: Read + Seek>(
    container: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    display_w: f32,
    display_h: f32,
) -> Option<EmbeddedImage> {
    let wsp = container
        .descendants()
        .find(|n| n.has_tag_name((WPS_NS, "wsp")))?;
    if wsp.descendants().any(|n| n.has_tag_name((WPS_NS, "txbx"))) {
        return None;
    }
    let mut img = shape_picture_fill(super::textbox::find_sp_pr(wsp)?, ctx, display_w, display_h)?;
    apply_pic_props(&mut img, container, ctx.theme);
    Some(img)
}

/// The picture of a shape's `a:blipFill` (`sp_pr`) for a `display_w` x
/// `display_h` frame. An `a:tile` fill is drawn into one raster.
pub(super) fn shape_picture_fill<R: Read + Seek>(
    sp_pr: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    display_w: f32,
    display_h: f32,
) -> Option<EmbeddedImage> {
    let fill = dml(sp_pr, "blipFill")?;
    let mut img = blip_image(fill, ctx, display_w, display_h)?;
    if let Some(tile) = dml(fill, "tile")
        && let Some((png, w, h)) = tile_picture(&img.data, tile, display_w, display_h)
    {
        img.data = std::sync::Arc::new(png);
        (img.pixel_width, img.pixel_height, img.format) = (w, h, ImageFormat::Png);
    }
    apply_blip_props(&mut img, fill);
    Some(img)
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
fn apply_pic_props(img: &mut EmbeddedImage, container: roxmltree::Node, theme: &ThemeFonts) {
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
    // A picture-filled shape has its properties on the shape (`wps:spPr`).
    let sp_pr = find_pic_sp_pr(container).or_else(|| {
        container
            .descendants()
            .find(|n| n.has_tag_name((WPS_NS, "wsp")))
            .and_then(super::textbox::find_sp_pr)
    });
    img.rotation_deg = parse_image_rotation(sp_pr);
    let (stroke_color, stroke_width) = parse_pic_outline(sp_pr, theme);
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
    apply_blip_props(img, container);
}

/// The `a:blip` adjustments under `container`: `a:srcRect` crop, brightness
/// and contrast, recolouring and transparency.
fn apply_blip_props(img: &mut EmbeddedImage, container: roxmltree::Node) {
    img.src_rect = parse_src_rect(container);
    img.lum = parse_lum(container);
    let blip = find_blip(container);
    img.recolor = blip.and_then(parse_recolor);
    img.alpha = blip
        .and_then(|b| dml(b, "alphaModFix"))
        .and_then(|a| frac_attr(a, "amt"))
        .filter(|&amt| amt < 1.0);
}

/// `a:grayscl` (black to white), or an `a:duotone` of two sRGB colours.
/// ponytail: a duotone in scheme or preset colours is left in full colour.
fn parse_recolor(blip: roxmltree::Node) -> Option<([u8; 3], [u8; 3])> {
    if dml(blip, "grayscl").is_some() {
        return Some(([0; 3], [255; 3]));
    }
    let colors: Vec<_> = dml(blip, "duotone")?
        .children()
        .filter(|n| n.is_element())
        .collect();
    let srgb = |n: &roxmltree::Node| {
        n.has_tag_name((DML_NS, "srgbClr"))
            .then(|| n.attribute("val").and_then(parse_hex_color))?
    };
    let [dark, light] = colors.as_slice() else {
        return None;
    };
    Some((srgb(dark)?, srgb(light)?))
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

/// Parse outline stroke from `pic:spPr/a:ln` (or a picture-filled shape's).
fn parse_pic_outline(sp_pr: Option<roxmltree::Node>, theme: &ThemeFonts) -> (Option<[u8; 3]>, f32) {
    sp_pr
        .and_then(|s| dml(s, "ln"))
        .and_then(|ln| super::color::parse_line_stroke(ln, theme))
        .map_or((None, 0.0), |(color, width)| (Some(color), width))
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

/// A shape's `a:effectLst` glow and reflection.
pub(super) fn parse_glow_reflection(
    sp_pr: Option<roxmltree::Node>,
) -> (Option<ImageGlow>, Option<ImageReflection>) {
    let fx = parse_pic_effects(sp_pr);
    (fx.glow, fx.reflection)
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

/// The first picture a drawing element's blips point to.
pub(super) fn blip_image<R: Read + Seek>(
    node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    w: f32,
    h: f32,
) -> Option<EmbeddedImage> {
    find_blip_embeds(node)
        .find_map(|id| read_image_from_zip(id, ctx.rels, ctx.zip, ctx.media, w, h))
}

/// An inline drawing's picture with its effect extent and picture properties.
fn inline_picture<R: Read + Seek>(
    container: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    w: f32,
    h: f32,
) -> Option<EmbeddedImage> {
    let (pic_w, pic_h) = inline_picture_extent(w, h);
    let mut img = blip_image(container, ctx, pic_w, pic_h)?;
    img.trimmed_width = w - pic_w;
    (img.layout_extra_height, img.layout_extra_top) = inline_extra_height(container);
    apply_pic_props(&mut img, container, ctx.theme);
    Some(img)
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
        Some("bottomMargin") => VRelativeFrom::BottomMargin,
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

/// A `wp:anchor` placed `relativeFrom="line"` vertically.
pub(super) fn anchored_to_line(container: roxmltree::Node) -> bool {
    wpd(container, "positionV").and_then(|n| n.attribute("relativeFrom")) == Some("line")
}

/// `base` placed at `container`'s `wp:anchor` position with its wrapping,
/// size `(width_pt, height_pt)` and stacking.
pub(super) fn place_anchored(
    container: roxmltree::Node,
    (width_pt, height_pt): (f32, f32),
    base: Textbox,
) -> Textbox {
    let (h_position, h_relative_from, v_position, v_relative_from) =
        parse_anchor_position(container);
    let (wrap_type, wrap_text, wrap_polygon) = parse_wrap_type(container);
    let (behind_doc, z_index) = anchor_z_order(container);
    let (dist_top, dist_bottom) = wrap_dist_top_bottom(container);
    let (dist_left, dist_right) = wrap_dist_left_right(container, wrap_polygon.is_some());
    Textbox {
        width_pt,
        height_pt,
        h_position,
        h_relative_from,
        v_offset_pt: v_position.offset_or_zero(),
        v_position,
        v_relative_from,
        wrap_type,
        wrap_text,
        wrap_polygon,
        dist_top,
        dist_bottom,
        dist_left,
        dist_right,
        behind_doc,
        z_index,
        v_from_line: anchored_to_line(container),
        ..base
    }
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
    let vertices: Vec<_> = poly
        .children()
        .filter(|c| c.has_tag_name((WPD_NS, "start")) || c.has_tag_name((WPD_NS, "lineTo")))
        .map(|c| {
            let coord = |a| c.attribute(a).and_then(|v| v.parse().ok()).unwrap_or(0);
            (coord("x"), coord("y"))
        })
        .collect();
    (!vertices.is_empty()).then_some(vertices)
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

/// A drawing in a run. Word lays a picture anchored in a text box's story out
/// inline where it is anchored, whatever its wrap or position.
pub(super) fn parse_run_drawing<R: Read + Seek>(
    drawing_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<RunDrawingResult> {
    match parse_placed_drawing(drawing_node, ctx) {
        Some(RunDrawingResult::Floating(fi)) if ctx.in_text_box && !fi.on_line => {
            Some(RunDrawingResult::Inline(fi.image))
        }
        result => result,
    }
}

fn parse_placed_drawing<R: Read + Seek>(
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
                let base = Textbox {
                    width_pct_margin: pct_width_of_margin(container),
                    ..wsp
                };
                let textbox = place_anchored(container, (display_w, display_h), base);
                return Some(RunDrawingResult::TextBox(textbox));
            }
            if let Some(conn) = parse_connector_from_wsp(container, ctx.theme) {
                return Some(RunDrawingResult::Connector(conn));
            }
            if let Some(img) =
                picture_filled_shape(container, ctx, display_w, display_h).or_else(|| {
                    let mut img = blip_image(container, ctx, display_w, display_h)?;
                    apply_pic_props(&mut img, container, ctx.theme);
                    Some(img)
                })
            {
                let (h_position, h_relative, v_position, v_relative) =
                    parse_anchor_position(container);
                let (wrap_type, wrap_text, wrap_polygon) = parse_wrap_type(container);
                let (behind_doc, z_index) = anchor_z_order(container);
                let (dist_top, dist_bottom) = wrap_dist_top_bottom(container);
                let (dist_left, dist_right) =
                    wrap_dist_left_right(container, wrap_polygon.is_some());
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
                    dist_left,
                    dist_right,
                    z_index,
                    anchor_seq: 0,
                    anchor_at_start: false,
                    anchor_run: 0,
                    on_line: false,
                    v_from_line: anchored_to_line(container),
                }));
            }
            // A wrapNone diagram floats at its anchor, the paragraph's text laid
            // out as if it weren't there.
            // ponytail: other wraps are laid out inline; add wrap zones with a
            // fixture that has one.
            if display_h > 0.0 && has_diagram_ref(container) {
                let mut diagram = parse_smartart_drawing(container, ctx, display_h);
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
            // A shape with nothing to draw still wraps text: an empty
            // unfilled, unlined letterhead box keeps the body below it.
            if let (wrap_type, wrap_text, _) = parse_wrap_type(container)
                && wrap_type != WrapType::None
                && container
                    .descendants()
                    .any(|n| n.has_tag_name((WPS_NS, "wsp")))
            {
                let (h_position, h_relative_from, v_position, v_relative_from) =
                    parse_anchor_position(container);
                let (behind_doc, z_index) = anchor_z_order(container);
                return Some(RunDrawingResult::TextBox(Textbox {
                    width_pt: display_w,
                    width_pct_margin: pct_width_of_margin(container),
                    height_pt: display_h,
                    h_position,
                    h_relative_from,
                    v_offset_pt: v_position.offset_or_zero(),
                    v_position,
                    v_relative_from,
                    wrap_type,
                    wrap_text,
                    dist_bottom: wrap_dist_top_bottom(container).1,
                    behind_doc,
                    z_index,
                    no_text_wrap: true,
                    ..Textbox::default()
                }));
            }
            continue;
        }

        if let Some(img) = picture_filled_shape(container, ctx, display_w, display_h) {
            return Some(RunDrawingResult::Inline(img));
        }

        // Inline textbox: wp:inline containing wps:wsp with text content
        if let Some(wsp) = parse_textbox_from_wsp(container, ctx) {
            // Treat inline textbox as a floating textbox at paragraph position
            // with TopAndBottom wrap so it acts as a block element
            return Some(RunDrawingResult::TextBox(Textbox {
                width_pt: display_w,
                height_pt: display_h,
                wrap_type: WrapType::TopAndBottom,
                on_line: true,
                ..wsp
            }));
        }

        if let Some(img) = inline_picture(container, ctx, display_w, display_h) {
            return Some(RunDrawingResult::Inline(img));
        }

        if let Some(chart_rid) = find_chart_ref(container) {
            if let Some(ic) = parse_chart_from_zip(
                chart_rid, ctx.rels, ctx.zip, display_w, display_h, ctx.theme,
            ) {
                return Some(RunDrawingResult::Chart(ic));
            }
        }

        if display_h > 0.0 && has_diagram_ref(container) {
            let diagram = parse_smartart_drawing(container, ctx, display_h);
            return Some(RunDrawingResult::SmartArt(diagram));
        }
    }
    None
}

fn find_chart_ref<'a>(container: roxmltree::Node<'a, 'a>) -> Option<&'a str> {
    container
        .descendants()
        .find(|n| {
            n.has_tag_name((DML_NS, "graphicData"))
                && matches!(
                    n.attribute("uri"),
                    Some(CHART_NS | super::chartex::CHARTEX_NS)
                )
        })
        .and_then(|gd| {
            gd.children()
                .find(|n| n.tag_name().name() == "chart")
                .and_then(|c| c.attribute((REL_NS, "id")))
        })
}

/// Height occupied by inline drawings, including transparent shape placeholders.
/// Select the same AlternateContent branch as run parsing, and never enter
/// a drawing's nested text box when measuring the enclosing paragraph.
pub(super) fn inline_drawing_height(node: roxmltree::Node) -> f32 {
    let mut height = 0.0_f32;
    for child in node.children().filter(|n| n.is_element()) {
        if child.has_tag_name((WML_NS, "drawing")) {
            for inline in child
                .children()
                .filter(|n| n.has_tag_name((WPD_NS, "inline")))
            {
                let (_, h) = extent_dimensions(inline);
                height = height.max(h + inline_extra_height(inline).0);
            }
        } else if child.has_tag_name((WML_NS, "r")) {
            height = height.max(inline_drawing_height(child));
        } else if child.has_tag_name((super::MC_NS_TOP, "AlternateContent")) {
            let branch = child
                .children()
                .find(|n| n.has_tag_name((super::MC_NS_TOP, "Choice")))
                .or_else(|| {
                    child
                        .children()
                        .find(|n| n.has_tag_name((super::MC_NS_TOP, "Fallback")))
                });
            if let Some(branch) = branch {
                height = height.max(inline_drawing_height(branch));
            }
        }
    }
    height
}

pub(super) fn compute_drawing_info<R: Read + Seek>(
    para_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
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
            let (extra_h, _) = inline_extra_height(container);
            let pic_h = if find_blip_embeds(container).next().is_some() {
                inline_picture_extent(display_w, display_h).1
            } else {
                display_h
            };
            max_height = max_height.max(pic_h + extra_h);

            if image.is_none() {
                image = inline_picture(container, ctx, display_w, display_h);
            }
        }
    }
    let object_height = compute_object_height(para_node);
    DrawingInfo {
        height: max_height.max(object_height),
        image,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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

#[cfg(test)]
mod inline_shape_height_tests {
    use super::inline_drawing_height;

    #[test]
    fn inline_extent_counts_without_a_blip_but_anchors_do_not() {
        let xml = r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:r><w:drawing><wp:inline><wp:extent cx="127000" cy="254000"/><wp:effectExtent t="12700" b="25400" l="0" r="0"/></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:anchor><wp:extent cx="127000" cy="1270000"/></wp:anchor></w:drawing></w:r></w:p>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert!((inline_drawing_height(doc.root_element()) - 23.0).abs() < 0.001);
    }

    #[test]
    fn alternate_content_counts_only_the_selected_branch() {
        let xml = r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:r><mc:AlternateContent><mc:Choice Requires="wp"><w:drawing><wp:inline><wp:extent cx="127000" cy="254000"/></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:drawing><wp:inline><wp:extent cx="127000" cy="1270000"/></wp:inline></w:drawing></mc:Fallback></mc:AlternateContent></w:r></w:p>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        assert!((inline_drawing_height(doc.root_element()) - 20.0).abs() < 0.001);
    }
}
