//! Drawing canvas (`wpc:wpc`) and shape group (`wpg:wgp`/`wpg:grpSp`) flattening.
//!
//! Word positions canvas/group children in a local coordinate space: a group's
//! `a:xfrm` maps the child space (`chOff`/`chExt`) onto the group's own
//! `off`/`ext`. We flatten the tree at parse time into independently
//! positioned textboxes, connectors, and images so the existing single-shape
//! render paths can draw them unchanged.

use std::io::{Read, Seek};

use crate::model::{
    AutoFit, FloatingImage, HRelativeFrom, HorizontalPosition, ShapeGeometry, TextAnchor, Textbox,
    VRelativeFrom, VerticalPosition, WrapText, WrapType,
};

use super::images::{
    RunDrawingResult, extent_dimensions, find_blip_embed, parse_anchor_position,
    read_image_from_zip,
};
use super::textbox::{find_sp_pr, parse_connector_shape_node, parse_wsp_shape};
use super::{ParseContext, WPS_NS, dml, emu_attr};

const WPC_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const WPG_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/// Affine scale+translate mapping local shape coordinates to drawing
/// coordinates (points, y-down, origin at the drawing's top-left).
#[derive(Clone, Copy)]
struct GroupTransform {
    sx: f32,
    sy: f32,
    tx: f32,
    ty: f32,
}

impl GroupTransform {
    const IDENTITY: GroupTransform = GroupTransform {
        sx: 1.0,
        sy: 1.0,
        tx: 0.0,
        ty: 0.0,
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
    GroupTransform {
        sx: parent.sx * sx_l,
        sy: parent.sy * sy_l,
        tx: parent.tx + parent.sx * tx_l,
        ty: parent.ty + parent.sy * ty_l,
    }
}

struct BaseAnchor {
    x: f32,
    y: f32,
    h_rel: HRelativeFrom,
    v_rel: VRelativeFrom,
    behind_doc: bool,
    z_index: u32,
}

/// Detect and flatten a canvas/group drawing. Returns None when the container
/// holds no canvas or group root, so the single-shape paths apply.
pub(super) fn parse_canvas_or_group<R: Read + Seek>(
    container: roxmltree::Node,
    is_anchor: bool,
    ctx: &mut ParseContext<'_, R>,
) -> Option<Vec<RunDrawingResult>> {
    let root = container.descendants().find(|n| {
        let tn = n.tag_name();
        (tn.name() == "wpc" && tn.namespace() == Some(WPC_NS))
            || (tn.name() == "wgp" && tn.namespace() == Some(WPG_NS))
    })?;

    let (display_w, display_h) = extent_dimensions(container);

    let base = if is_anchor {
        let (h_pos, h_rel, v_pos, v_rel) = parse_anchor_position(container);
        BaseAnchor {
            x: match h_pos {
                HorizontalPosition::Offset(v) => v,
                _ => 0.0,
            },
            y: match v_pos {
                VerticalPosition::Offset(v) => v,
                _ => 0.0,
            },
            h_rel,
            v_rel,
            behind_doc: container.attribute("behindDoc") == Some("1"),
            z_index: container
                .attribute("relativeHeight")
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(0),
        }
    } else {
        BaseAnchor {
            x: 0.0,
            y: 0.0,
            h_rel: HRelativeFrom::Column,
            v_rel: VRelativeFrom::Paragraph,
            behind_doc: false,
            z_index: 0,
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
        out.push(RunDrawingResult::TextBox(Textbox {
            paragraphs: Vec::new(),
            width_pt: display_w,
            height_pt: display_h,
            h_position: HorizontalPosition::Offset(0.0),
            h_relative_from: HRelativeFrom::Column,
            v_offset_pt: 0.0,
            v_position: VerticalPosition::Offset(0.0),
            v_relative_from: VRelativeFrom::Paragraph,
            fill: None,
            shape_type: ShapeGeometry::default(),
            stroke_color: None,
            stroke_width: 0.0,
            text_anchor: TextAnchor::Top,
            margin_left: 0.0,
            margin_right: 0.0,
            margin_top: 0.0,
            margin_bottom: 0.0,
            wrap_type: WrapType::TopAndBottom,
            dist_top: 0.0,
            dist_bottom: 0.0,
            behind_doc: false,
            no_text_wrap: true,
            is_wordart: false,
            text_warp: None,
            auto_fit: AutoFit::None,
            z_index: 0,
        }));
    }
    out.extend(shapes);
    Some(out)
}

fn find_group_xfrm(group: roxmltree::Node) -> Option<Xfrm> {
    let grp_sp_pr = group.children().find(|n| {
        n.tag_name().name() == "grpSpPr" && n.tag_name().namespace() == Some(WPG_NS)
    })?;
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
            (Some(WPG_NS), "grpSp") => {
                let t2 = find_group_xfrm(child)
                    .map(|x| compose(t, &x))
                    .unwrap_or(t);
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
    let has_txbx = wsp
        .children()
        .any(|n| n.tag_name().name() == "txbx" && n.tag_name().namespace() == Some(WPS_NS));
    let is_connector = matches!(prst, Some("line" | "straightConnector1" | "arc")) && !has_txbx;

    if is_connector {
        if let Some(mut conn) = parse_connector_shape_node(wsp, ctx.theme) {
            conn.x = base.x + x;
            conn.y = base.y + y;
            conn.width = w;
            conn.height = h;
            conn.z_index = base.z_index;
            out.push(RunDrawingResult::Connector(conn));
        }
        return;
    }

    if let Some(shape) = parse_wsp_shape(wsp, ctx) {
        out.push(RunDrawingResult::TextBox(Textbox {
            paragraphs: shape.paragraphs,
            width_pt: w,
            height_pt: h,
            h_position: HorizontalPosition::Offset(base.x + x),
            h_relative_from: base.h_rel,
            v_offset_pt: base.y + y,
            v_position: VerticalPosition::Offset(base.y + y),
            v_relative_from: base.v_rel,
            fill: shape.fill,
            shape_type: shape.shape_type,
            stroke_color: shape.stroke_color,
            stroke_width: shape.stroke_width,
            text_anchor: shape.text_anchor,
            margin_left: shape.margin_left,
            margin_right: shape.margin_right,
            margin_top: shape.margin_top,
            margin_bottom: shape.margin_bottom,
            wrap_type: WrapType::None,
            dist_top: 0.0,
            dist_bottom: 0.0,
            behind_doc: base.behind_doc,
            no_text_wrap: shape.no_text_wrap,
            is_wordart: shape.is_wordart,
            text_warp: shape.text_warp,
            auto_fit: shape.auto_fit,
            z_index: base.z_index,
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
    let Some(sp_pr) = pic.children().find(|n| {
        n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(PIC_NS)
    }) else {
        return;
    };
    let Some(xfrm) = read_xfrm(sp_pr) else { return };
    let (x, y) = t.apply(xfrm.off.0, xfrm.off.1);
    let (w, h) = t.scale(xfrm.ext.0, xfrm.ext.1);

    let Some(embed_id) = find_blip_embed(pic) else {
        return;
    };
    if let Some(img) = read_image_from_zip(embed_id, ctx.rels, ctx.zip, w, h) {
        out.push(RunDrawingResult::Floating(FloatingImage {
            image: img,
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
        }));
    }
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
