use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{EmbeddedImage, SmartArtDiagram, SmartArtShape};

use super::styles::ThemeFonts;
use super::color::parse_solid_fill;
use super::images::{find_blip_embed, read_image_from_zip};
use super::textbox::parse_shape_geometry;
use super::{DML_NS, DSP_NS, dml, dsp, emu_attr, read_zip_text};

fn is_ns(node: roxmltree::Node, name: &str, ns: &str) -> bool {
    node.tag_name().name() == name && node.tag_name().namespace() == Some(ns)
}

/// Load the OPC relationship file for a given part and resolve relative targets
/// to full zip paths (e.g. `../media/image1.jpg` relative to `word/diagrams/`
/// becomes `word/media/image1.jpg`).
fn load_part_rels<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    part_path: &str,
) -> HashMap<String, String> {
    let (dir, file) = part_path.rsplit_once('/').unwrap_or(("", part_path));
    let rels_path = if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    };
    let Some(xml) = read_zip_text(zip, &rels_path) else {
        return HashMap::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return HashMap::new();
    };
    doc.root_element()
        .children()
        .filter(|n| n.is_element())
        .filter_map(|n| {
            let id = n.attribute("Id")?.to_string();
            let raw_target = n.attribute("Target")?;
            // Prefix with / so read_image_from_zip can strip it for zip lookup
            let target = if raw_target.starts_with('/') {
                raw_target.to_string()
            } else {
                format!("/{}", normalize_path(&format!("{dir}/{raw_target}")))
            };
            Some((id, target))
        })
        .collect()
}

/// Resolve `..` segments in a path (e.g. `word/diagrams/../media/img.jpg` → `word/media/img.jpg`).
fn normalize_path(path: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for seg in path.split('/') {
        if seg == ".." {
            parts.pop();
        } else if !seg.is_empty() && seg != "." {
            parts.push(seg);
        }
    }
    parts.join("/")
}

const DIAGRAM_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

fn has_dml(node: roxmltree::Node, name: &str) -> bool {
    node.children()
        .any(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

pub(super) fn has_diagram_ref(container: roxmltree::Node) -> bool {
    container.descendants().any(|n| {
        n.tag_name().name() == "graphicData"
            && n.tag_name().namespace() == Some(DML_NS)
            && n.attribute("uri") == Some(DIAGRAM_URI)
    })
}

/// Resolve the drawing file for a specific SmartArt diagram by extracting
/// the r:dm (data) relationship from dgm:relIds and deriving the drawing path.
fn find_diagram_drawing(
    container: roxmltree::Node,
    rels: &HashMap<String, String>,
) -> Option<String> {
    let dgm_rel_ids = container.descendants().find(|n| {
        n.tag_name().name() == "relIds"
            && n.tag_name().namespace() == Some(DIAGRAM_URI)
    })?;
    let dm_rid = dgm_rel_ids.attribute((super::REL_NS, "dm"))?;
    let data_target = rels.get(dm_rid)?;
    Some(data_target.replace("/data", "/drawing"))
}

pub(super) fn parse_smartart_drawing<R: Read + Seek>(
    container: roxmltree::Node,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    theme: &ThemeFonts,
    display_w: f32,
    display_h: f32,
) -> SmartArtDiagram {
    let mut shapes = Vec::new();

    let drawing_target = find_diagram_drawing(container, rels)
        .or_else(|| rels.values().find(|t| t.contains("diagrams/drawing")).cloned());
    if let Some(target) = drawing_target {
        let zip_path = target
            .strip_prefix('/')
            .map(String::from)
            .unwrap_or_else(|| format!("word/{}", target));

        let diagram_rels = load_part_rels(zip, &zip_path);

        if let Some(xml) = read_zip_text(zip, &zip_path) {
            if let Ok(doc) = roxmltree::Document::parse(&xml) {
                let sp_tree = dsp(doc.root(), "drawing").and_then(|d| dsp(d, "spTree"));

                if let Some(tree) = sp_tree {
                    shapes = tree
                        .children()
                        .filter(|n| is_ns(*n, "sp", DSP_NS))
                        .filter_map(|sp| parse_dsp_shape(sp, theme, &diagram_rels, zip))
                        .collect();
                }
            }
        }
    }

    SmartArtDiagram {
        display_width: display_w,
        display_height: display_h,
        shapes,
    }
}

fn parse_dsp_shape<R: Read + Seek>(
    sp: roxmltree::Node,
    theme: &ThemeFonts,
    diagram_rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> Option<SmartArtShape> {
    let sp_pr = dsp(sp, "spPr")?;
    let xfrm = dml(sp_pr, "xfrm")?;
    let off = dml(xfrm, "off")?;
    let ext = dml(xfrm, "ext")?;

    let x = emu_attr(off, "x");
    let y = emu_attr(off, "y");
    let w = emu_attr(ext, "cx");
    let h = emu_attr(ext, "cy");
    // rot is in 60,000ths of a degree
    let rotation_deg = xfrm
        .attribute("rot")
        .and_then(|v| v.parse::<f64>().ok())
        .map(|v| (v / 60_000.0) as f32)
        .unwrap_or(0.0);

    let fill = if has_dml(sp_pr, "noFill") {
        None
    } else {
        parse_solid_fill(sp_pr, theme)
    };

    let image_fill: Option<EmbeddedImage> = dml(sp_pr, "blipFill").and_then(|bf| {
        let embed_id = find_blip_embed(bf)?;
        read_image_from_zip(embed_id, diagram_rels, zip, w, h)
    });

    let (stroke_color, stroke_width) = dml(sp_pr, "ln")
        .and_then(|ln| {
            if has_dml(ln, "noFill") {
                return None;
            }
            let color = parse_solid_fill(ln, theme)?;
            let width = ln
                .attribute("w")
                .and_then(|v| v.parse::<f32>().ok())
                .map(super::emu_to_pts)
                .unwrap_or(0.75);
            Some((color, width))
        })
        .map_or((None, 0.0), |(c, w)| (Some(c), w));

    let tp = parse_dsp_text(sp, theme);

    // Resolve default text color from dsp:style/a:fontRef
    let default_text_color = dsp(sp, "style")
        .and_then(|style| dml(style, "fontRef"))
        .and_then(|fr| parse_solid_fill(fr, theme)
            .or_else(|| {
                fr.children()
                    .find(|n| n.tag_name().name() == "schemeClr" && n.tag_name().namespace() == Some(DML_NS))
                    .and_then(|sc| {
                        let val = sc.attribute("val")?;
                        let key = super::resolve_theme_color_key(val);
                        theme.colors.get(key).copied()
                    })
            }));

    let paragraphs: Vec<SmartArtPara> = tp.paragraphs.into_iter().map(|mut para| {
        for run in &mut para.runs {
            if run.color.is_none() {
                run.color = default_text_color;
            }
        }
        para
    }).collect();

    // dsp:txXfrm provides a separate rectangle for text placement
    let text_rect = dsp(sp, "txXfrm").and_then(|tx| {
        let tx_off = dml(tx, "off")?;
        let tx_ext = dml(tx, "ext")?;
        Some((
            emu_attr(tx_off, "x"),
            emu_attr(tx_off, "y"),
            emu_attr(tx_ext, "cx"),
            emu_attr(tx_ext, "cy"),
        ))
    });

    let has_text = paragraphs.iter().any(|p| p.runs.iter().any(|r| !r.text.is_empty()));
    if fill.is_none() && image_fill.is_none() && !has_text && stroke_color.is_none() {
        return None;
    }

    Some(SmartArtShape {
        x,
        y,
        width: w,
        height: h,
        rotation_deg,
        shape_type: parse_shape_geometry(sp_pr),
        fill,
        image_fill,
        stroke_color,
        stroke_width,
        paragraphs,
        text_rect,
        text_inset_left: tp.left_inset,
        text_anchor: tp.anchor,
    })
}

use crate::model::{SmartArtPara, SmartArtRun, SmartArtTextAnchor, SmartArtTextAlign};

struct DspTextProps {
    paragraphs: Vec<SmartArtPara>,
    left_inset: f32,
    anchor: SmartArtTextAnchor,
}

fn parse_dsp_text(sp: roxmltree::Node, theme: &ThemeFonts) -> DspTextProps {
    let Some(body) = dsp(sp, "txBody") else {
        return DspTextProps {
            paragraphs: Vec::new(),
            left_inset: 0.0, anchor: SmartArtTextAnchor::Top,
        };
    };
    let body_pr = dml(body, "bodyPr");
    let left_inset = body_pr
        .and_then(|bp| bp.attribute("lIns"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(super::emu_to_pts)
        .unwrap_or(0.0);
    let anchor = match body_pr.and_then(|bp| bp.attribute("anchor")) {
        Some("ctr") => SmartArtTextAnchor::Center,
        Some("b") => SmartArtTextAnchor::Bottom,
        _ => SmartArtTextAnchor::Top,
    };

    // Resolve theme font fallback from dsp:style/a:fontRef
    let theme_font = dsp(sp, "style")
        .and_then(|s| dml(s, "fontRef"))
        .and_then(|fr| fr.attribute("idx"))
        .and_then(|idx| match idx {
            "minor" => Some(theme.minor.clone()),
            "major" => Some(theme.major.clone()),
            _ => None,
        })
        .filter(|n| !n.is_empty());

    let mut default_font_size = 0.0_f32;
    let mut paragraphs = Vec::new();

    for p in body.children().filter(|n| is_ns(*n, "p", DML_NS)) {
        let ppr = dml(p, "pPr");
        let bullet = ppr
            .and_then(|pp| dml(pp, "buChar"))
            .and_then(|bc| bc.attribute("char"))
            .map(String::from);
        let align = ppr
            .and_then(|pp| pp.attribute("algn"))
            .map(|a| match a {
                "ctr" => SmartArtTextAlign::Center,
                "r" => SmartArtTextAlign::Right,
                _ => SmartArtTextAlign::Left,
            })
            .unwrap_or(SmartArtTextAlign::Left);

        let mut runs = Vec::new();
        for r in p.children().filter(|n| is_ns(*n, "r", DML_NS)) {
            let text = dml(r, "t")
                .and_then(|t| t.text())
                .unwrap_or("")
                .to_string();
            if text.is_empty() { continue; }

            let rpr = dml(r, "rPr");
            let font_size = rpr
                .and_then(|rp| rp.attribute("sz"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|sz| sz / 100.0)
                .unwrap_or(default_font_size);
            if default_font_size == 0.0 && font_size > 0.0 {
                default_font_size = font_size;
            }

            let font_name = rpr
                .and_then(|rp| dml(rp, "latin"))
                .and_then(|n| n.attribute("typeface"))
                .filter(|tf| !tf.is_empty())
                .map(String::from)
                .or_else(|| theme_font.clone());

            let bold = rpr.and_then(|rp| rp.attribute("b")) == Some("1");
            let italic = rpr.and_then(|rp| rp.attribute("i")) == Some("1");
            let underline = rpr.and_then(|rp| rp.attribute("u")).is_some_and(|v| v != "none");
            let strikethrough = rpr.and_then(|rp| rp.attribute("strike")).is_some_and(|v| v != "noStrike");
            let baseline = rpr
                .and_then(|rp| rp.attribute("baseline"))
                .and_then(|v| v.parse::<i32>().ok())
                .unwrap_or(0);

            let color = rpr.and_then(|rp| parse_solid_fill(rp, theme));
            let highlight = rpr
                .and_then(|rp| dml(rp, "highlight"))
                .and_then(|hl| {
                    let clr = dml(hl, "srgbClr")?;
                    super::parse_hex_color(clr.attribute("val")?)
                });

            runs.push(SmartArtRun {
                text,
                font_name,
                font_size,
                bold,
                italic,
                color,
                underline,
                strikethrough,
                baseline,
                highlight,
            });
        }

        if !runs.is_empty() {
            paragraphs.push(SmartArtPara { runs, bullet, align });
        }
    }

    DspTextProps { paragraphs, left_inset, anchor }
}
