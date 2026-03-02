use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    Alignment, HorizontalPosition, LineSpacing, Paragraph, ParagraphBorders, Textbox,
};

use super::{DML_NS, WML_NS, WPD_NS, WPS_NS, MC_NS_TOP, twips_to_pts, twips_attr, wml, wml_attr};
use super::numbering::NumberingInfo;
use super::runs::parse_runs;
use super::styles::{StylesInfo, ThemeFonts, parse_alignment, parse_line_spacing};
use super::images::parse_anchor_position;

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
    for p in txbx_content.children().filter(|n| {
        n.tag_name().name() == "p" && n.tag_name().namespace() == Some(WML_NS)
    }) {
        let parsed = parse_runs(p, styles, theme, rels, zip, numbering);
        let ppr = wml(p, "pPr");
        let para_style_id = ppr
            .and_then(|ppr| wml_attr(ppr, "pStyle"))
            .unwrap_or("Normal");
        let para_style = styles.paragraph_styles.get(para_style_id);
        let alignment = ppr
            .and_then(|ppr| wml_attr(ppr, "jc"))
            .map(parse_alignment)
            .or_else(|| para_style.and_then(|s| s.alignment))
            .unwrap_or(Alignment::Left);
        let inline_spacing = ppr.and_then(|ppr| wml(ppr, "spacing"));
        let space_before = inline_spacing
            .and_then(|n| n.attribute((WML_NS, "before")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
            .or_else(|| para_style.and_then(|s| s.space_before))
            .unwrap_or(0.0);
        let space_after = inline_spacing
            .and_then(|n| n.attribute((WML_NS, "after")))
            .and_then(|v| v.parse::<f32>().ok())
            .map(twips_to_pts)
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
        let tab_stops = ppr
            .map(super::parse_tab_stops)
            .unwrap_or_default();
        let num_pr = ppr.and_then(|ppr| wml(ppr, "numPr"));
        let (mut indent_left, mut indent_hanging, list_label, list_label_font) =
            super::numbering::parse_list_info(
                num_pr,
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
            content_height: 0.0,
            alignment,
            indent_left,
            indent_right,
            indent_hanging,
            indent_first_line,
            list_label,
            list_label_font,
            contextual_spacing: false,
            keep_next: false,
            keep_lines: false,
            line_spacing,
            image: None,
            borders: ParagraphBorders::default(),
            shading: None,
            page_break_before: false,
            column_break_before: false,
            tab_stops,
            extra_line_breaks: parsed.line_break_count,
            floating_images: parsed.floating_images,
            textboxes: parsed.textboxes,
        });
    }
    paragraphs
}

fn resolve_scheme_color(base: [u8; 3], fill_node: roxmltree::Node) -> [u8; 3] {
    let mut lum_mod: Option<f32> = None;
    let mut lum_off: Option<f32> = None;
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
            _ => {}
        }
    }
    if lum_mod.is_none() && lum_off.is_none() {
        return base;
    }
    let m = lum_mod.unwrap_or(1.0);
    let o = lum_off.unwrap_or(0.0);
    [
        ((base[0] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
        ((base[1] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
        ((base[2] as f32 / 255.0 * m + o) * 255.0).clamp(0.0, 255.0) as u8,
    ]
}

fn parse_solid_fill(sp_pr: roxmltree::Node, theme: &ThemeFonts) -> Option<[u8; 3]> {
    let fill = sp_pr.children().find(|n| {
        n.tag_name().name() == "solidFill" && n.tag_name().namespace() == Some(DML_NS)
    })?;
    // Direct sRGB color
    if let Some(srgb) = fill.children().find(|n| {
        n.tag_name().name() == "srgbClr" && n.tag_name().namespace() == Some(DML_NS)
    }) {
        return srgb
            .attribute("val")
            .and_then(super::parse_hex_color);
    }
    // Theme color reference
    if let Some(scheme) = fill.children().find(|n| {
        n.tag_name().name() == "schemeClr" && n.tag_name().namespace() == Some(DML_NS)
    }) {
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

fn parse_body_margins(wsp: roxmltree::Node) -> (f32, f32, f32, f32) {
    let body_pr = wsp.children().find(|n| {
        n.tag_name().name() == "bodyPr" && n.tag_name().namespace() == Some(WPS_NS)
    });
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
    pub(super) fill_color: Option<[u8; 3]>,
    pub(super) margin_top: f32,
    pub(super) margin_left: f32,
    pub(super) margin_bottom: f32,
    pub(super) margin_right: f32,
}

pub(super) fn parse_textbox_from_wsp<R: Read + std::io::Seek>(
    anchor: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    numbering: &NumberingInfo,
) -> Option<WspResult> {
    let wsp = anchor.descendants().find(|n| {
        n.tag_name().name() == "wsp" && n.tag_name().namespace() == Some(WPS_NS)
    })?;

    // Extract fill color from spPr
    let sp_pr = wsp.children().find(|n| {
        n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(WPS_NS)
            || n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(DML_NS)
    });
    let fill_color = sp_pr.and_then(|sp| parse_solid_fill(sp, theme));
    let has_no_fill = sp_pr.is_some_and(|sp| {
        sp.children().any(|n| {
            n.tag_name().name() == "noFill" && n.tag_name().namespace() == Some(DML_NS)
        })
    });

    let (margin_top, margin_left, margin_bottom, margin_right) = parse_body_margins(wsp);

    // Try to get textbox content
    let paragraphs = wsp
        .children()
        .find(|n| n.tag_name().name() == "txbx" && n.tag_name().namespace() == Some(WPS_NS))
        .and_then(|txbx| {
            txbx.children().find(|n| {
                n.tag_name().name() == "txbxContent"
                    && n.tag_name().namespace() == Some(WML_NS)
            })
        })
        .map(|tc| parse_txbx_content_paragraphs(tc, styles, theme, rels, zip, numbering))
        .unwrap_or_default();

    // Return if there's text content OR a visible fill
    if paragraphs.is_empty() && (has_no_fill || fill_color.is_none()) {
        return None;
    }

    Some(WspResult {
        paragraphs,
        fill_color,
        margin_top,
        margin_left,
        margin_bottom,
        margin_right,
    })
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
    let textbox = shape.children().find(|n| {
        n.tag_name().name() == "textbox" && n.tag_name().namespace() == Some(VML_NS)
    })?;
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
            let parse_pt = |s: &str| -> f32 {
                s.trim_end_matches("pt").parse::<f32>().unwrap_or(0.0)
            };
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

    let paragraphs = parse_txbx_content_paragraphs(txbx_content, styles, theme, rels, zip, numbering);
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
        fill_color: None,
        margin_left: 7.2,
        margin_right: 7.2,
        margin_top: 3.6,
        margin_bottom: 3.6,
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
                    n.tag_name().namespace() == Some(WML_NS)
                        && n.tag_name().name() == "drawing"
                }) {
                    for container in drawing.children().filter(|n| {
                        n.tag_name().namespace() == Some(WPD_NS)
                            && n.tag_name().name() == "anchor"
                    }) {
                        let extent = container.children().find(|n| {
                            n.tag_name().name() == "extent"
                                && n.tag_name().namespace() == Some(WPD_NS)
                        });
                        let cx = extent
                            .and_then(|n| n.attribute("cx"))
                            .and_then(|v| v.parse::<f32>().ok())
                            .unwrap_or(0.0);
                        let cy = extent
                            .and_then(|n| n.attribute("cy"))
                            .and_then(|v| v.parse::<f32>().ok())
                            .unwrap_or(0.0);
                        let display_w = cx / 12700.0;
                        let display_h = cy / 12700.0;

                        if let Some(wsp) = parse_textbox_from_wsp(
                            container, rels, zip, styles, theme, numbering,
                        ) {
                            let (h_position, h_relative, v_offset, v_relative) =
                                parse_anchor_position(container);
                            textboxes.push(Textbox {
                                paragraphs: wsp.paragraphs,
                                width_pt: display_w,
                                height_pt: display_h,
                                h_position,
                                h_relative_from: h_relative,
                                v_offset_pt: v_offset,
                                v_relative_from: v_relative,
                                fill_color: wsp.fill_color,
                                margin_left: wsp.margin_left,
                                margin_right: wsp.margin_right,
                                margin_top: wsp.margin_top,
                                margin_bottom: wsp.margin_bottom,
                            });
                        }
                    }
                }
            } else if let Some(branch) = fallback {
                // VML fallback: mc:Fallback → w:pict → v:shape → v:textbox
                for pict in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS)
                        && n.tag_name().name() == "pict"
                }) {
                    if let Some(tb) =
                        parse_textbox_from_vml(pict, rels, zip, styles, theme, numbering)
                    {
                        textboxes.push(tb);
                    }
                }
                // Also check for w:r/w:pict inside fallback
                for r in branch.children().filter(|n| {
                    n.tag_name().namespace() == Some(WML_NS)
                        && n.tag_name().name() == "r"
                }) {
                    for pict in r.children().filter(|n| {
                        n.tag_name().namespace() == Some(WML_NS)
                            && n.tag_name().name() == "pict"
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
