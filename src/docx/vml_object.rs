//! Legacy VML pictures and `w:object` OLE embeddings.

use std::io::{Read, Seek};

use crate::model::{
    EmbeddedImage, FloatingImage, HorizontalPosition, VerticalPosition, WrapText, WrapType,
};

use super::images::read_image_from_zip;
use super::textbox::{parse_vml_box_style, vml_style_prop};
use super::{ParseContext, REL_NS, VML_NS, W10_NS, WML_NS, parse_pt, twips_attr};

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
    let picture = vml_picture_source(obj)?;
    let embed_id = picture.attribute((REL_NS, "id"))?;
    let (w, h) = object_dimensions(obj)?;
    let mut img = read_image_from_zip(embed_id, ctx.rels, ctx.zip, ctx.media, w, h)?;
    (img.src_rect, img.lum) = vml_imagedata_props(picture);
    Some(with_object_alt(obj, img))
}

/// The element holding a VML picture's relationship id: its `v:imagedata`, or
/// else the frame fill of the shape itself (not of a `v:shapetype` beside it).
fn vml_picture_source<'a, 'i>(obj: roxmltree::Node<'a, 'i>) -> Option<roxmltree::Node<'a, 'i>> {
    obj.descendants()
        .find(|n| n.has_tag_name((VML_NS, "imagedata")))
        .or_else(|| {
            first_vml_shape(obj)?
                .children()
                .find(|c| is_vml_frame_fill(*c))
        })
}

/// The picture a VML shape's own frame fill stretches over its `w`×`h` box.
pub(super) fn vml_frame_fill_image<R: Read + Seek>(
    shape: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    w: f32,
    h: f32,
) -> Option<EmbeddedImage> {
    let fill = shape.children().find(|c| is_vml_frame_fill(*c))?;
    read_image_from_zip(
        fill.attribute((REL_NS, "id"))?,
        ctx.rels,
        ctx.zip,
        ctx.media,
        w,
        h,
    )
}

/// A legacy VML picture: a `w:pict` whose shape shows only an image (no text
/// box), as `v:imagedata` or as a fill stretched over the shape. Laid out like
/// an OLE object's preview picture.
pub(super) fn is_vml_picture(pict: roxmltree::Node) -> bool {
    pict.children().any(|n| {
        n.tag_name().namespace() == Some(VML_NS)
            && matches!(n.tag_name().name(), "shape" | "rect")
            && n.children()
                .any(|c| c.has_tag_name((VML_NS, "imagedata")) || is_vml_frame_fill(c))
            && !n.children().any(|c| {
                c.has_tag_name((VML_NS, "textbox")) || c.has_tag_name((VML_NS, "textpath"))
            })
    })
}

/// `<v:fill type="frame" r:id="..."/>`: a picture stretched over the whole
/// shape.
fn is_vml_frame_fill(node: roxmltree::Node) -> bool {
    let off = |value: Option<&str>| matches!(value, Some("f" | "false" | "0"));
    node.has_tag_name((VML_NS, "fill"))
        && node.attribute("type") == Some("frame")
        && node.attribute((REL_NS, "id")).is_some()
        && !off(node.attribute("on"))
        && !node
            .parent()
            .is_some_and(|shape| off(shape.attribute("filled")))
}

/// VML fractions are plain decimals (".5"), percentages ("50%") or 16.16 fixed
/// point with an `f` suffix ("19661f" = 0.3).
pub(super) fn vml_fraction(v: &str) -> Option<f32> {
    let v = v.trim();
    if let Some(fixed) = v.strip_suffix('f') {
        return fixed.parse::<f32>().ok().map(|f| f / 65536.0);
    }
    if let Some(pct) = v.strip_suffix('%') {
        return pct.parse::<f32>().ok().map(|p| p / 100.0);
    }
    v.parse().ok()
}

/// VML stacks by z-index; a negative one is behind the text. Word's VML
/// fallbacks write a front shape's DrawingML relativeHeight as its z-index and a
/// behind-text one as 503316480 minus it, so the least negative behind-text
/// z-index draws on top (57 fallback pairs in the fixtures, and Word's own
/// rendering of vml-group-behind-order-*).
pub(super) fn vml_z_order(z_index: i64) -> (bool, u32) {
    let relative_height = if z_index < 0 {
        503_316_480 + z_index
    } else {
        z_index
    };
    (
        z_index < 0,
        relative_height.clamp(0, u32::MAX as i64) as u32,
    )
}

/// `v:imagedata` crop and colour adjustments in DrawingML terms. Word writes its
/// Washout preset (`a:lum bright=70% contrast=-70%`) as gain 0.3 / blacklevel
/// 0.35, so contrast = gain − 1 and brightness = 2 × blacklevel.
pub(super) fn vml_imagedata_props(
    imagedata: roxmltree::Node,
) -> (Option<[f32; 4]>, Option<(f32, f32)>) {
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
    let shape = first_vml_shape(obj)?;
    let style = shape.attribute("style")?;
    if !style.contains("position:absolute") {
        return None;
    }
    let b = parse_vml_box_style(style);
    let prop = |key| vml_style_prop(style, key);
    // Word's VML wrap distances default to 9pt at the sides, 0 above and below
    // (slovak_pedagogical's header logo: centred text stops 9pt short of it).
    let dist = |key, default| prop(key).and_then(parse_pt).unwrap_or(default);
    // "absolute" (or no value) means the margin offsets place it.
    let h_align = match prop("mso-position-horizontal") {
        Some("left" | "inside") => Some(HorizontalPosition::AlignLeft),
        Some("center") => Some(HorizontalPosition::AlignCenter),
        Some("right" | "outside") => Some(HorizontalPosition::AlignRight),
        _ => None,
    };
    let v_align = match prop("mso-position-vertical") {
        Some("top" | "inside") => Some(VerticalPosition::AlignTop),
        Some("center") => Some(VerticalPosition::AlignCenter),
        Some("bottom" | "outside") => Some(VerticalPosition::AlignBottom),
        _ => None,
    };
    let image = parse_object_inline_image(obj, ctx)?;
    let (behind_doc, z_index) = vml_z_order(b.z_index);
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
        h_position: h_align.unwrap_or(HorizontalPosition::Offset(b.margin_left)),
        h_relative_from: b.h_relative,
        v_position: v_align.unwrap_or(VerticalPosition::Offset(b.margin_top)),
        v_relative_from: b.v_relative,
        wrap_type,
        wrap_text: WrapText::BothSides,
        wrap_polygon: None,
        behind_doc,
        dist_top: dist("mso-wrap-distance-top", 0.0),
        dist_bottom: dist("mso-wrap-distance-bottom", 0.0),
        dist_left: dist("mso-wrap-distance-left", 9.0),
        dist_right: dist("mso-wrap-distance-right", 9.0),
        z_index,
        anchor_seq: 0,
        anchor_at_start: false,
        anchor_run: 0,
        on_line: false,
        v_from_line: false,
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
    first_vml_shape(obj)
        .and_then(|s| s.attribute("style"))
        .is_some_and(|style| style.contains("position:absolute"))
}

fn first_vml_shape<'a, 'i>(obj: roxmltree::Node<'a, 'i>) -> Option<roxmltree::Node<'a, 'i>> {
    obj.children().find(|n| {
        n.tag_name().namespace() == Some(VML_NS)
            && matches!(n.tag_name().name(), "rect" | "shape" | "oval" | "roundrect")
    })
}

fn object_dimensions(obj: roxmltree::Node) -> Option<(f32, f32)> {
    let style = first_vml_shape(obj).and_then(|s| s.attribute("style"));
    let pt = |key| vml_style_prop(style?, key).and_then(parse_pt);
    pt("width")
        .zip(pt("height"))
        .or_else(|| twips_attr(obj, "dxaOrig").zip(twips_attr(obj, "dyaOrig")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vml_z_index_maps_to_word_relative_height() {
        // Pairs Word wrote as relativeHeight + VML fallback z-index.
        assert_eq!(vml_z_order(-251658240), (true, 251658240));
        assert_eq!(vml_z_order(-251657216), (true, 251659264));
        assert_eq!(vml_z_order(251659264), (false, 251659264));
        assert_eq!(vml_z_order(0), (false, 0));
    }

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

    #[test]
    fn frame_fill_makes_a_vml_picture() {
        for (shape, expected) in [
            (r#"<v:rect><v:fill r:id="p" type="frame"/></v:rect>"#, true),
            (r#"<v:rect><v:imagedata r:id="p"/></v:rect>"#, true),
            (r#"<v:rect><v:fill r:id="p" type="tile"/></v:rect>"#, false),
            (
                r#"<v:rect><v:fill r:id="p" type="frame" on="f"/></v:rect>"#,
                false,
            ),
            // Word for Mac draws no picture here either: on="t" does not override filled="f".
            (
                r#"<v:rect filled="f"><v:fill r:id="p" type="frame" on="t"/></v:rect>"#,
                false,
            ),
            (
                r#"<v:shape><v:fill r:id="p" type="frame"/><v:textpath string="A"/></v:shape>"#,
                false,
            ),
            (
                r#"<v:rect><v:fill r:id="p" type="frame"/><v:textbox/></v:rect>"#,
                false,
            ),
        ] {
            let xml = format!(
                r#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="{VML_NS}" xmlns:r="{REL_NS}">{shape}</w:pict>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            assert_eq!(is_vml_picture(doc.root_element()), expected, "{shape}");
        }
    }

    #[test]
    fn picture_comes_from_the_shape_not_its_shapetype() {
        let xml = format!(
            r#"<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="{VML_NS}" xmlns:r="{REL_NS}"><v:shapetype><v:fill r:id="type" type="frame"/></v:shapetype><v:rect><v:fill r:id="shape" type="frame"/></v:rect></w:pict>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let source = vml_picture_source(doc.root_element()).unwrap();
        assert_eq!(source.attribute((REL_NS, "id")), Some("shape"));
    }
}
