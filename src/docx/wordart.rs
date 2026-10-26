use crate::model::{
    Alignment, AutoFit, HorizontalPosition, Paragraph, Run, TextFill, TextGlow, TextOutline,
    TextShadow, TextWarp, Textbox, VerticalPosition,
};

use super::color::{apply_color_transforms, parse_color_transforms, resolve_dml_color};
use super::styles::{StylesInfo, ThemeFonts};
use super::textbox::{parse_avlst, parse_vml_box_style, parse_vml_color};
use super::{
    W14_NS, angle_attr, dml as find_dml, emu_attr_opt, find_child, frac_attr, parse_hex_color,
    parse_on_off, parse_pt, resolve_theme_color_key,
};

fn find_w14<'a>(parent: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(parent, name, W14_NS)
}

/// Parsed WordArt-specific properties from wps:bodyPr.
pub(super) struct WordArtBodyProps {
    pub is_wordart: bool,
    pub text_warp: Option<TextWarp>,
    pub auto_fit: AutoFit,
}

/// Parse WordArt-related attributes from a wps:bodyPr node.
pub(super) fn parse_wordart_body_pr(body_pr: roxmltree::Node) -> WordArtBodyProps {
    let is_wordart = body_pr.attribute("fromWordArt").is_some_and(parse_on_off);

    let text_warp = find_dml(body_pr, "prstTxWarp").and_then(|prst| {
        let preset = prst.attribute("prst")?.to_string();
        let adjustments = parse_avlst(prst);
        Some(TextWarp {
            preset,
            adjustments,
        })
    });

    let auto_fit = if find_dml(body_pr, "spAutoFit").is_some() {
        AutoFit::Shape
    } else if find_dml(body_pr, "normAutofit").is_some() {
        AutoFit::Normal
    } else {
        AutoFit::None
    };

    WordArtBodyProps {
        is_wordart,
        text_warp,
        auto_fit,
    }
}

/// Parse w14:textOutline from a w:rPr node.
pub(super) fn parse_text_outline(rpr: roxmltree::Node, theme: &ThemeFonts) -> Option<TextOutline> {
    let outline = find_w14(rpr, "textOutline")?;

    let width_pt = emu_attr_opt(outline, (W14_NS, "w")).unwrap_or(0.75);

    let color = find_w14_solid_fill_color(outline, theme)?;
    Some(TextOutline { width_pt, color })
}

/// Parse w14:textFill from a w:rPr node.
pub(super) fn parse_text_fill(rpr: roxmltree::Node, theme: &ThemeFonts) -> Option<TextFill> {
    let text_fill = find_w14(rpr, "textFill")?;

    if find_w14(text_fill, "noFill").is_some() {
        return Some(TextFill::NoFill);
    }

    if let Some(color) = find_w14(text_fill, "solidFill").and_then(|n| resolve_w14_color(n, theme))
    {
        return Some(TextFill::Solid(color));
    }

    if let Some(grad) = find_w14(text_fill, "gradFill") {
        return parse_w14_gradient(grad, theme);
    }

    None
}

/// Parse w14:shadow from a w:rPr node.
pub(super) fn parse_text_shadow(rpr: roxmltree::Node, theme: &ThemeFonts) -> Option<TextShadow> {
    let shadow = find_w14(rpr, "shadow")?;

    let color = find_w14_solid_fill_color(shadow, theme).unwrap_or([128, 128, 128]);

    let blur_rad = emu_attr_opt(shadow, (W14_NS, "blurRad")).unwrap_or(0.0);

    let dist = emu_attr_opt(shadow, (W14_NS, "dist")).unwrap_or(blur_rad.max(1.5));

    // Direction in 60000ths of a degree, clockwise from right
    let dir = angle_attr(shadow, (W14_NS, "dir")).unwrap_or(225.0);

    let alpha = find_w14(shadow, "srgbClr")
        .and_then(|srgb| find_w14(srgb, "alpha"))
        .and_then(|a| frac_attr(a, (W14_NS, "val")))
        .unwrap_or(0.6);

    // Convert polar (dist, dir) to cartesian offsets
    let dir_rad = dir.to_radians();
    let offset_x = dist * dir_rad.cos();
    let offset_y = dist * dir_rad.sin();

    Some(TextShadow {
        color,
        offset_x,
        offset_y,
        alpha,
        relief: false,
    })
}

/// Parse w14:glow from a w:rPr node.
pub(super) fn parse_text_glow(rpr: roxmltree::Node, theme: &ThemeFonts) -> Option<TextGlow> {
    let glow = find_w14(rpr, "glow")?;

    let radius_pt = emu_attr_opt(glow, (W14_NS, "rad")).unwrap_or(0.0);

    if radius_pt <= 0.0 {
        return None;
    }

    let color = resolve_w14_color(glow, theme).unwrap_or([255, 255, 0]);

    Some(TextGlow { color, radius_pt })
}

/// Parse VML WordArt when v:textpath is present instead of v:textbox.
pub(super) fn parse_vml_wordart(
    shape: roxmltree::Node,
    textpath: roxmltree::Node,
    styles: &StylesInfo,
    _theme: &ThemeFonts,
) -> Option<Textbox> {
    let text = textpath.attribute("string")?;
    if text.is_empty() {
        return None;
    }

    let style_str = textpath.attribute("style").unwrap_or("");
    let (font_name, font_size) = parse_vml_font_style(style_str);
    // Regular unless the textpath style asks for a weight (swiss_materials'
    // watermark is plain Arial; it used to be drawn bold).
    let bold = style_str.split(';').any(|part| {
        part.trim()
            .strip_prefix("font-weight:")
            .is_some_and(|w| w.trim() == "bold" || w.trim().parse::<u32>().is_ok_and(|n| n >= 600))
    });

    // The fill's opacity is blended into the colour: WordArt sits on the page
    // background (swiss_materials' watermark is red at 50%, drawn pink).
    let fill_color = shape.attribute("fillcolor").and_then(parse_vml_color).map(|c| {
        let a = shape
            .children()
            .find(|n| n.has_tag_name((super::VML_NS, "fill")))
            .and_then(|f| f.attribute("opacity"))
            .and_then(parse_vml_opacity)
            .unwrap_or(1.0);
        c.map(|v| (v as f32 * a + 255.0 * (1.0 - a)).round() as u8)
    });

    let run = Run {
        text: text.to_string(),
        font_name: font_name.unwrap_or_else(|| styles.defaults.font_name.clone()),
        font_size: font_size.unwrap_or(36.0),
        bold,
        color: fill_color,
        ..Run::default()
    };

    let para = Paragraph {
        runs: vec![run],
        alignment: Alignment::Center,
        ..Paragraph::default()
    };

    let shape_style = shape.attribute("style").unwrap_or("");
    let (width, height) = parse_vml_dimensions(shape_style);
    let bs = parse_vml_box_style(shape_style);
    // `fitshape` stretches the glyphs to the shape box (both ways), which is
    // what the textPlain warp draws; the 1pt `font-size` Word writes is moot.
    // The flag usually sits on the shapetype's textpath (every WordArt preset
    // sets it), so it is on unless the shape's own textpath turns it off.
    // ponytail: fitpath / unfitted WordArt would need the shapetype looked up.
    let fit_shape = !matches!(textpath.attribute("fitshape"), Some("f" | "false" | "0"));

    Some(Textbox {
        paragraphs: vec![para],
        width_pt: width,
        height_pt: height,
        h_position: if bs.h_center {
            HorizontalPosition::AlignCenter
        } else {
            HorizontalPosition::Offset(bs.margin_left)
        },
        h_relative_from: bs.h_relative,
        v_offset_pt: bs.margin_top,
        v_position: if bs.v_center {
            VerticalPosition::AlignCenter
        } else {
            VerticalPosition::Offset(bs.margin_top)
        },
        v_relative_from: bs.v_relative,
        behind_doc: bs.behind_text,
        rotation_deg: bs.rotation_deg,
        fit_shape,
        no_text_wrap: true,
        auto_fit: if fit_shape { AutoFit::None } else { AutoFit::Shape },
        text_warp: fit_shape.then(|| TextWarp {
            preset: "textPlain".to_string(),
            adjustments: Vec::new(),
        }),
        ..Textbox::default()
    })
}

/// VML opacity: a fraction (".5"), a percentage ("50%") or 16.16 fixed point ("32768f").
fn parse_vml_opacity(val: &str) -> Option<f32> {
    let val = val.trim();
    if let Some(f) = val.strip_suffix('f') {
        return f.parse::<f32>().ok().map(|v| v / 65536.0);
    }
    if let Some(p) = val.strip_suffix('%') {
        return p.parse::<f32>().ok().map(|v| v / 100.0);
    }
    val.parse().ok()
}

/// Resolve a `w14:srgbClr` or `w14:schemeClr` child of `fill_node`.
/// Theme colors are looked up in `theme.colors` and modified by any
/// `w14:lumMod`/`lumOff`/`tint`/`shade` transforms on the schemeClr node.
fn resolve_w14_color(fill_node: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    if let Some(srgb) = find_w14(fill_node, "srgbClr") {
        return srgb.attribute((W14_NS, "val")).and_then(parse_hex_color);
    }
    if let Some(scheme) = find_w14(fill_node, "schemeClr") {
        let val = scheme.attribute((W14_NS, "val"))?;
        let theme_key = resolve_theme_color_key(val);
        let base = *theme.colors.get(theme_key)?;
        let transforms = parse_color_transforms(scheme);
        return Some(apply_color_transforms(base, &transforms));
    }
    // DML fallback (rare; resolve_dml_color also covers prstClr).
    resolve_dml_color(fill_node, theme)
}

fn find_w14_solid_fill_color(parent: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    find_w14(parent, "solidFill")
        .and_then(|n| resolve_w14_color(n, theme))
        .or_else(|| find_dml(parent, "solidFill").and_then(|n| resolve_w14_color(n, theme)))
}

fn parse_w14_gradient(grad: roxmltree::Node, theme: &ThemeFonts) -> Option<TextFill> {
    let gs_lst = find_w14(grad, "gsLst")?;

    let mut stops = Vec::new();
    for gs in gs_lst.children().filter(|n| n.has_tag_name((W14_NS, "gs"))) {
        let pos = frac_attr(gs, (W14_NS, "pos")).unwrap_or(0.0);
        if let Some(color) = resolve_w14_color(gs, theme) {
            stops.push((color, pos));
        }
    }

    if stops.is_empty() {
        return None;
    }

    let angle_deg = find_w14(grad, "lin")
        .and_then(|lin| angle_attr(lin, (W14_NS, "ang")))
        .unwrap_or(0.0);

    Some(TextFill::Gradient { stops, angle_deg })
}

fn parse_vml_font_style(style: &str) -> (Option<String>, Option<f32>) {
    let mut font_name = None;
    let mut font_size = None;

    for part in style.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            let val = val.trim().trim_matches('"').trim_matches('\'');
            match key.trim() {
                "font-family" => font_name = Some(val.to_string()),
                "font-size" => {
                    font_size = parse_pt(val);
                }
                _ => {}
            }
        }
    }

    (font_name, font_size)
}

fn parse_vml_dimensions(style: &str) -> (f32, f32) {
    let mut width = 0.0_f32;
    let mut height = 0.0_f32;
    for part in style.split(';') {
        if let Some((key, val)) = part.trim().split_once(':') {
            match key.trim() {
                "width" => width = parse_pt(val.trim()).unwrap_or(0.0),
                "height" => height = parse_pt(val.trim()).unwrap_or(0.0),
                _ => {}
            }
        }
    }

    (width, height)
}
