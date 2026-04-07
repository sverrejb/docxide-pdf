use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{SmartArtDiagram, SmartArtShape};

use super::styles::ThemeFonts;
use super::color::parse_solid_fill;
use super::textbox::parse_shape_geometry;
use super::{DML_NS, DSP_NS, dml, dsp, emu_attr, read_zip_text};

fn is_ns(node: roxmltree::Node, name: &str, ns: &str) -> bool {
    node.tag_name().name() == name && node.tag_name().namespace() == Some(ns)
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

        if let Some(xml) = read_zip_text(zip, &zip_path) {
            if let Ok(doc) = roxmltree::Document::parse(&xml) {
                let sp_tree = dsp(doc.root(), "drawing").and_then(|d| dsp(d, "spTree"));

                if let Some(tree) = sp_tree {
                    shapes = tree
                        .children()
                        .filter(|n| is_ns(*n, "sp", DSP_NS))
                        .filter_map(|sp| parse_dsp_shape(sp, theme))
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

fn parse_dsp_shape(sp: roxmltree::Node, theme: &ThemeFonts) -> Option<SmartArtShape> {
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

    let (text, font_size, text_color) = parse_dsp_text(sp, theme);

    if fill.is_none() && text.is_empty() {
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
        stroke_color,
        stroke_width,
        text,
        font_size,
        text_color,
    })
}

fn parse_dsp_text(sp: roxmltree::Node, theme: &ThemeFonts) -> (String, f32, Option<[u8; 3]>) {
    let Some(body) = dsp(sp, "txBody") else {
        return (String::new(), 0.0, None);
    };

    let mut lines = Vec::new();
    let mut font_size = 0.0;
    let mut text_color = None;

    for p in body.children().filter(|n| is_ns(*n, "p", DML_NS)) {
        let mut line_text = String::new();
        for r in p.children().filter(|n| is_ns(*n, "r", DML_NS)) {
            if let Some(rpr) = dml(r, "rPr") {
                if font_size == 0.0 {
                    if let Some(sz) = rpr.attribute("sz").and_then(|v| v.parse::<f32>().ok()) {
                        font_size = sz / 100.0;
                    }
                }
                if text_color.is_none() {
                    text_color = parse_solid_fill(rpr, theme);
                }
            }
            if let Some(t) = dml(r, "t") {
                if let Some(text) = t.text() {
                    line_text.push_str(text);
                }
            }
        }
        if !line_text.is_empty() {
            lines.push(line_text);
        }
    }

    (lines.join("\n"), font_size, text_color)
}
