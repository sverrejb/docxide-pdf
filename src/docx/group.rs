//! Drawing canvas (`wpc:wpc`) and shape group (`wpg:wgp`/`wpg:grpSp`) flattening.
//!
//! Word positions canvas/group children in a local coordinate space: a group's
//! `a:xfrm` maps the child space (`chOff`/`chExt`) onto the group's own
//! `off`/`ext`. We flatten the tree at parse time into independently
//! positioned textboxes, connectors, and images so the existing single-shape
//! render paths can draw them unchanged.

use std::io::{Read, Seek};

use crate::model::{
    ConnectorShape, ConnectorType, EmbeddedImage, FloatingImage, HRelativeFrom, HorizontalPosition,
    Textbox, VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use super::images::{
    RunDrawingResult, anchor_z_order, extent_dimensions, find_blip_embeds, parse_anchor_position,
    read_image_from_zip, vml_imagedata_props, vml_z_order,
};
use super::textbox::{
    find_sp_pr, parse_connector_shape_node, parse_vml_box_style, parse_vml_line, parse_wsp_shape,
    vml_style_prop,
};
use super::{
    PIC_NS, ParseContext, REL_NS, VML_NS, WPC_NS, WPG_NS, WPS_NS, dml, emu_attr, parse_pt,
};

/// Affine scale+translate mapping local shape coordinates to drawing
/// coordinates (points, y-down, origin at the drawing's top-left).
#[derive(Clone, Copy)]
struct GroupTransform {
    sx: f32,
    sy: f32,
    tx: f32,
    ty: f32,
    connector_sx: f32,
    connector_sy: f32,
    connector_tx: f32,
    connector_ty: f32,
}

impl GroupTransform {
    const IDENTITY: GroupTransform = GroupTransform {
        sx: 1.0,
        sy: 1.0,
        tx: 0.0,
        ty: 0.0,
        connector_sx: 1.0,
        connector_sy: 1.0,
        connector_tx: 0.0,
        connector_ty: 0.0,
    };

    fn apply(&self, x: f32, y: f32) -> (f32, f32) {
        (self.tx + self.sx * x, self.ty + self.sy * y)
    }

    fn scale(&self, w: f32, h: f32) -> (f32, f32) {
        (self.sx * w, self.sy * h)
    }
}

struct Xfrm {
    off: (f32, f32),
    ext: (f32, f32),
    ch_off: (f32, f32),
    ch_ext: (f32, f32),
    flip_h: bool,
    flip_v: bool,
}

fn read_xfrm(sp_pr: roxmltree::Node) -> Option<Xfrm> {
    let xfrm = dml(sp_pr, "xfrm")?;
    let off = dml(xfrm, "off").map(|o| (emu_attr(o, "x"), emu_attr(o, "y")))?;
    let ext = dml(xfrm, "ext").map(|e| (emu_attr(e, "cx"), emu_attr(e, "cy")))?;
    let ch_off = dml(xfrm, "chOff")
        .map(|o| (emu_attr(o, "x"), emu_attr(o, "y")))
        .unwrap_or((0.0, 0.0));
    let ch_ext = dml(xfrm, "chExt")
        .map(|e| (emu_attr(e, "cx"), emu_attr(e, "cy")))
        .unwrap_or(ext);
    Some(Xfrm {
        off,
        ext,
        ch_off,
        ch_ext,
        flip_h: matches!(xfrm.attribute("flipH"), Some("1" | "true")),
        flip_v: matches!(xfrm.attribute("flipV"), Some("1" | "true")),
    })
}

/// Compose a group's child-space mapping onto the parent transform.
fn compose(parent: GroupTransform, xfrm: &Xfrm) -> GroupTransform {
    let sx_l = if xfrm.ch_ext.0 > 0.0 {
        xfrm.ext.0 / xfrm.ch_ext.0
    } else {
        1.0
    };
    let sy_l = if xfrm.ch_ext.1 > 0.0 {
        xfrm.ext.1 / xfrm.ch_ext.1
    } else {
        1.0
    };
    let tx_l = xfrm.off.0 - xfrm.ch_off.0 * sx_l;
    let ty_l = xfrm.off.1 - xfrm.ch_off.1 * sy_l;
    // Connectors reflect around the group's extent, including nested groups.
    // Other shape paths retain their existing transform until their content
    // mirroring (text, images and arcs) is implemented separately.
    let csx = if xfrm.flip_h { -sx_l } else { sx_l };
    let csy = if xfrm.flip_v { -sy_l } else { sy_l };
    let ctx = xfrm.off.0 + if xfrm.flip_h { xfrm.ext.0 } else { 0.0 } - xfrm.ch_off.0 * csx;
    let cty = xfrm.off.1 + if xfrm.flip_v { xfrm.ext.1 } else { 0.0 } - xfrm.ch_off.1 * csy;
    GroupTransform {
        sx: parent.sx * sx_l,
        sy: parent.sy * sy_l,
        tx: parent.tx + parent.sx * tx_l,
        ty: parent.ty + parent.sy * ty_l,
        connector_sx: parent.connector_sx * csx,
        connector_sy: parent.connector_sy * csy,
        connector_tx: parent.connector_tx + parent.connector_sx * ctx,
        connector_ty: parent.connector_ty + parent.connector_sy * cty,
    }
}

struct BaseAnchor {
    x: f32,
    y: f32,
    h_rel: HRelativeFrom,
    v_rel: VRelativeFrom,
    behind_doc: bool,
    z_index: u32,
    /// True for inline canvases — their flattened children flow at the
    /// paragraph text-start and so must pick up the paragraph's left indent.
    on_line: bool,
}

/// Detect and flatten a canvas/group drawing. Returns None when the container
/// holds no canvas or group root, so the single-shape paths apply.
pub(super) fn parse_canvas_or_group<R: Read + Seek>(
    container: roxmltree::Node,
    is_anchor: bool,
    ctx: &mut ParseContext<'_, R>,
) -> Option<Vec<RunDrawingResult>> {
    let root = container
        .descendants()
        .find(|n| n.has_tag_name((WPC_NS, "wpc")) || n.has_tag_name((WPG_NS, "wgp")))?;

    let (display_w, display_h) = extent_dimensions(container);

    let base = if is_anchor {
        let (h_pos, h_rel, v_pos, v_rel) = parse_anchor_position(container);
        let (behind_doc, z_index) = anchor_z_order(container);
        BaseAnchor {
            x: h_pos.offset_or_zero(),
            y: v_pos.offset_or_zero(),
            h_rel,
            v_rel,
            behind_doc,
            z_index,
            on_line: false,
        }
    } else {
        BaseAnchor {
            x: 0.0,
            y: 0.0,
            h_rel: HRelativeFrom::Column,
            v_rel: VRelativeFrom::Paragraph,
            behind_doc: false,
            z_index: 0,
            on_line: true,
        }
    };

    // A group root carries its own child-space mapping; a canvas root's
    // children are already in drawing coordinates.
    let root_transform = if root.tag_name().namespace() == Some(WPG_NS) {
        find_group_xfrm(root)
            .map(|x| compose(GroupTransform::IDENTITY, &x))
            .unwrap_or(GroupTransform::IDENTITY)
    } else {
        GroupTransform::IDENTITY
    };

    let mut shapes = Vec::new();
    walk_group(root, root_transform, &base, ctx, &mut shapes);
    if shapes.is_empty() {
        return None;
    }

    let mut out = Vec::new();
    if !is_anchor {
        // Inline canvases occupy a block in the text flow: reserve the full
        // extent with an invisible TopAndBottom textbox, then place the
        // children relative to the same paragraph without wrapping.
        let effect_height = super::wpd(container, "effectExtent")
            .map_or(0.0, |e| emu_attr(e, "t") + emu_attr(e, "b"));
        out.push(RunDrawingResult::TextBox(Textbox {
            width_pt: display_w,
            height_pt: display_h + effect_height.max(0.0),
            wrap_type: WrapType::TopAndBottom,
            no_text_wrap: true,
            on_line: true,
            ..Textbox::default()
        }));
    }
    out.extend(shapes);
    Some(out)
}

fn find_group_xfrm(group: roxmltree::Node) -> Option<Xfrm> {
    let grp_sp_pr = group
        .children()
        .find(|n| n.has_tag_name((WPG_NS, "grpSpPr")))?;
    read_xfrm(grp_sp_pr)
}

fn walk_group<R: Read + Seek>(
    node: roxmltree::Node,
    t: GroupTransform,
    base: &BaseAnchor,
    ctx: &mut ParseContext<'_, R>,
    out: &mut Vec<RunDrawingResult>,
) {
    for child in node.children() {
        let tn = child.tag_name();
        match (tn.namespace(), tn.name()) {
            // Both grpSp (nested group) and wgp (a wordprocessingGroup root
            // nested inside a canvas) carry a grpSpPr transform and child
            // shapes. Recursing only into grpSp dropped entire wgp subtrees
            // (e.g. the roundRect border + divider lines in a 4-quadrant
            // canvas — annotation #165).
            (Some(WPG_NS), "grpSp" | "wgp") => {
                let t2 = find_group_xfrm(child).map(|x| compose(t, &x)).unwrap_or(t);
                walk_group(child, t2, base, ctx, out);
            }
            (Some(WPS_NS), "wsp") => emit_wsp(child, t, base, ctx, out),
            (Some(PIC_NS), "pic") => emit_pic(child, t, base, ctx, out),
            _ => {}
        }
    }
}

fn emit_wsp<R: Read + Seek>(
    wsp: roxmltree::Node,
    t: GroupTransform,
    base: &BaseAnchor,
    ctx: &mut ParseContext<'_, R>,
    out: &mut Vec<RunDrawingResult>,
) {
    let Some(sp_pr) = find_sp_pr(wsp) else { return };
    let Some(xfrm) = read_xfrm(sp_pr) else { return };
    let (x, y) = t.apply(xfrm.off.0, xfrm.off.1);
    let (w, h) = t.scale(xfrm.ext.0, xfrm.ext.1);

    let prst = dml(sp_pr, "prstGeom").and_then(|g| g.attribute("prst"));
    let has_txbx = wsp.children().any(|n| n.has_tag_name((WPS_NS, "txbx")));
    let is_connector = matches!(prst, Some("line" | "straightConnector1" | "arc")) && !has_txbx;

    if is_connector {
        if let Some(mut conn) = parse_connector_shape_node(wsp, ctx.theme) {
            let (mut cx, mut cy) = (x, y);
            conn.width = w;
            conn.height = h;
            if let ConnectorType::Line { flip_h, flip_v } = &mut conn.connector_type {
                let x0 = t.connector_tx + t.connector_sx * xfrm.off.0;
                let y0 = t.connector_ty + t.connector_sy * xfrm.off.1;
                let x1 = x0 + t.connector_sx * xfrm.ext.0;
                let y1 = y0 + t.connector_sy * xfrm.ext.1;
                (cx, cy) = (x0.min(x1), y0.min(y1));
                conn.width = (x1 - x0).abs();
                conn.height = (y1 - y0).abs();
                *flip_h ^= t.connector_sx < 0.0;
                *flip_v ^= t.connector_sy < 0.0;
            }
            place_connector(&mut conn, cx, cy, base);
            out.push(RunDrawingResult::Connector(conn));
        }
        return;
    }

    if let Some(shape) = parse_wsp_shape(wsp, ctx) {
        out.push(RunDrawingResult::TextBox(Textbox {
            width_pt: w,
            height_pt: h,
            h_position: HorizontalPosition::Offset(base.x + x),
            h_relative_from: base.h_rel,
            v_offset_pt: base.y + y,
            v_position: VerticalPosition::Offset(base.y + y),
            v_relative_from: base.v_rel,
            behind_doc: base.behind_doc,
            z_index: base.z_index,
            on_line: base.on_line,
            ..Textbox::from(shape)
        }));
    }
}

fn emit_pic<R: Read + Seek>(
    pic: roxmltree::Node,
    t: GroupTransform,
    base: &BaseAnchor,
    ctx: &mut ParseContext<'_, R>,
    out: &mut Vec<RunDrawingResult>,
) {
    let Some(sp_pr) = pic.children().find(|n| n.has_tag_name((PIC_NS, "spPr"))) else {
        return;
    };
    let Some(xfrm) = read_xfrm(sp_pr) else { return };
    let (x, y) = t.apply(xfrm.off.0, xfrm.off.1);
    let (w, h) = t.scale(xfrm.ext.0, xfrm.ext.1);

    if let Some(img) =
        find_blip_embeds(pic).find_map(|id| read_image_from_zip(id, ctx.rels, ctx.zip, w, h))
    {
        // ponytail: pics flattened from an *inline* canvas don't pick up the
        // paragraph's left indent (FloatingImage::on_line only moves the
        // vertical anchor).
        // Add it the way Textbox does if a fixture ever needs an indented
        // inline canvas with picture children.
        out.push(RunDrawingResult::Floating(group_image(img, x, y, base)));
    }
}

fn group_image(image: EmbeddedImage, x: f32, y: f32, base: &BaseAnchor) -> FloatingImage {
    FloatingImage {
        image,
        h_position: HorizontalPosition::Offset(base.x + x),
        h_relative_from: base.h_rel,
        v_position: VerticalPosition::Offset(base.y + y),
        v_relative_from: base.v_rel,
        wrap_type: WrapType::None,
        wrap_text: WrapText::BothSides,
        wrap_polygon: None,
        behind_doc: base.behind_doc,
        dist_top: 0.0,
        dist_bottom: 0.0,
        dist_left: 0.0,
        dist_right: 0.0,
        z_index: base.z_index,
        anchor_seq: 0,
        anchor_at_start: false,
        on_line: base.on_line,
    }
}

fn place_connector(conn: &mut ConnectorShape, x: f32, y: f32, base: &BaseAnchor) {
    conn.h_position = HorizontalPosition::Offset(base.x + x);
    conn.h_relative_from = base.h_rel;
    conn.v_position = VerticalPosition::Offset(base.y + y);
    conn.v_relative_from = base.v_rel;
    conn.behind_doc = base.behind_doc;
    conn.z_index = base.z_index;
}

/// Flatten legacy VML groups, whose children use coordorigin/coordsize units.
/// The leaf renderers still receive ordinary point-based anchors.
pub(super) fn parse_vml_group<R: Read + Seek>(
    pict: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
) -> Option<RunDrawingResult> {
    let group = pict
        .children()
        .find(|n| n.has_tag_name((VML_NS, "group")))?;
    let style = group.attribute("style").unwrap_or("");
    let box_style = parse_vml_box_style(style);
    let (behind_doc, z_index) = vml_z_order(box_style.z_index);
    let base = BaseAnchor {
        x: 0.0,
        y: 0.0,
        h_rel: box_style.h_relative,
        v_rel: box_style.v_relative,
        behind_doc,
        z_index,
        on_line: false,
    };
    let mut out = Vec::new();
    walk_vml_group(group, GroupTransform::IDENTITY, &base, 0, ctx, &mut out);
    (!out.is_empty()).then_some(RunDrawingResult::Group(out))
}

fn walk_vml_group<R: Read + Seek>(
    node: roxmltree::Node,
    parent: GroupTransform,
    base: &BaseAnchor,
    depth: usize,
    ctx: &mut ParseContext<'_, R>,
    out: &mut Vec<RunDrawingResult>,
) {
    if depth > 32 {
        return;
    }
    let style = node.attribute("style").unwrap_or("");
    if vml_style_prop(style, "visibility") == Some("hidden")
        || vml_style_prop(style, "rotation").is_some()
        || vml_style_prop(style, "flip").is_some()
    {
        return;
    }
    let number = |key: &str| vml_style_prop(style, key).and_then(parse_pt).unwrap_or(0.0);
    let x = number("left") + number("margin-left");
    let y = number("top") + number("margin-top");
    let w = number("width");
    let h = number("height");
    if ![x, y, w, h].iter().all(|v| v.is_finite()) || w < 0.0 || h < 0.0 {
        return;
    }
    if node.has_tag_name((VML_NS, "group")) {
        let Some(ch_ext) = node.attribute("coordsize").and_then(pair) else {
            return;
        };
        let ch_off = node
            .attribute("coordorigin")
            .and_then(pair)
            .unwrap_or((0.0, 0.0));
        if ![ch_ext.0, ch_ext.1, ch_off.0, ch_off.1]
            .iter()
            .all(|v| v.is_finite())
            || ch_ext.0 <= 0.0
            || ch_ext.1 <= 0.0
        {
            return;
        }
        let xfrm = Xfrm {
            off: (x, y),
            ext: (w, h),
            ch_off,
            ch_ext,
            flip_h: false,
            flip_v: false,
        };
        let t = compose(parent, &xfrm);
        for child in node.children().filter(|n| n.is_element()) {
            walk_vml_group(child, t, base, depth + 1, ctx, out);
        }
    } else if node.has_tag_name((VML_NS, "shape")) || node.has_tag_name((VML_NS, "rect")) {
        let (x, y) = parent.apply(x, y);
        let (w, h) = parent.scale(w, h);
        if let Some(mut conn) = parse_vml_line(node) {
            conn.width = w;
            conn.height = h;
            place_connector(&mut conn, x, y, base);
            out.push(RunDrawingResult::Connector(conn));
        } else if let Some(data) = node
            .children()
            .find(|n| n.has_tag_name((VML_NS, "imagedata")))
            && let Some(mut img) = data
                .attribute((REL_NS, "id"))
                .and_then(|id| read_image_from_zip(id, ctx.rels, ctx.zip, w, h))
        {
            (img.src_rect, img.lum) = vml_imagedata_props(data);
            out.push(RunDrawingResult::Floating(group_image(img, x, y, base)));
        }
    }
}

fn pair(value: &str) -> Option<(f32, f32)> {
    let (x, y) = value.split_once(',')?;
    Some((x.trim().parse().ok()?, y.trim().parse().ok()?))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn xfrm(off: (f32, f32), ext: (f32, f32), ch_off: (f32, f32), ch_ext: (f32, f32)) -> Xfrm {
        Xfrm {
            off,
            ext,
            ch_off,
            ch_ext,
            flip_h: false,
            flip_v: false,
        }
    }

    #[test]
    fn compose_maps_child_space_onto_group_extent() {
        // isla drawing#3: group at (0,0) 489.6x202.5 with child space
        // chOff=(-6.2,0) chExt=479.2x322.5
        let t = compose(
            GroupTransform::IDENTITY,
            &xfrm((0.0, 0.0), (489.6, 202.5), (-6.2, 0.0), (479.2, 322.5)),
        );
        // Child-space origin corner maps to the group's top-left
        let (x, y) = t.apply(-6.2, 0.0);
        assert!((x - 0.0).abs() < 1e-3 && (y - 0.0).abs() < 1e-3);
        // Child-space far corner maps to the group's bottom-right
        let (x, y) = t.apply(-6.2 + 479.2, 322.5);
        assert!((x - 489.6).abs() < 1e-2 && (y - 202.5).abs() < 1e-2);
        // Sizes scale by ext/chExt
        let (w, h) = t.scale(479.2, 322.5);
        assert!((w - 489.6).abs() < 1e-2 && (h - 202.5).abs() < 1e-2);
    }

    #[test]
    fn reflected_group_maps_zero_height_lines_to_opposite_edges() {
        let mut group = xfrm((5.0, 7.0), (200.0, 20.0), (0.0, 0.0), (200.0, 16.0));
        group.flip_v = true;
        let t = compose(GroupTransform::IDENTITY, &group);
        // A bottom-edge horizontal line becomes the top-edge line; the
        // former top edge moves to the bottom. Zero height must stay zero.
        assert!((t.connector_ty + t.connector_sy * 16.0 - 7.0).abs() < 1e-4);
        assert!((t.connector_ty - 27.0).abs() < 1e-4);
        assert!((t.connector_tx - 5.0).abs() < 1e-4);
        assert_eq!(t.connector_sy * 0.0, 0.0);
    }

    #[test]
    fn nested_reflections_cancel_and_preserve_child_offset() {
        let mut outer = xfrm((0.0, 0.0), (100.0, 100.0), (0.0, 0.0), (100.0, 100.0));
        outer.flip_h = true;
        outer.flip_v = true;
        let mut inner = xfrm((10.0, 20.0), (30.0, 40.0), (0.0, 0.0), (30.0, 40.0));
        inner.flip_h = true;
        inner.flip_v = true;
        let t = compose(compose(GroupTransform::IDENTITY, &outer), &inner);
        assert_eq!((t.connector_sx, t.connector_sy), (1.0, 1.0));
        assert_eq!((t.connector_tx, t.connector_ty), (60.0, 40.0));
    }

    #[test]
    fn compose_nests_transforms() {
        // Outer group halves both axes; inner group translates by (10, 20)
        let outer = compose(
            GroupTransform::IDENTITY,
            &xfrm((0.0, 0.0), (50.0, 50.0), (0.0, 0.0), (100.0, 100.0)),
        );
        let inner = compose(
            outer,
            &xfrm((10.0, 20.0), (30.0, 30.0), (0.0, 0.0), (30.0, 30.0)),
        );
        // Child point (0,0) in inner space → inner offset scaled by outer
        let (x, y) = inner.apply(0.0, 0.0);
        assert!((x - 5.0).abs() < 1e-3 && (y - 10.0).abs() < 1e-3);
        let (w, h) = inner.scale(30.0, 30.0);
        assert!((w - 15.0).abs() < 1e-3 && (h - 15.0).abs() < 1e-3);
    }

    #[test]
    fn compose_defaults_when_child_extent_zero() {
        let t = compose(
            GroupTransform::IDENTITY,
            &xfrm((5.0, 5.0), (10.0, 10.0), (0.0, 0.0), (0.0, 0.0)),
        );
        let (x, y) = t.apply(1.0, 1.0);
        assert!((x - 6.0).abs() < 1e-3 && (y - 6.0).abs() < 1e-3);
    }
}
