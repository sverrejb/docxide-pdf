use std::collections::HashMap;
use std::io::Read;

use crate::model::{SmartArtDiagram, SmartArtShape};

use super::styles::ThemeFonts;
use super::textbox::{parse_shape_geometry, parse_solid_fill};
use super::{DML_NS, read_zip_text};

const DSP_NS: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const DIAGRAM_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

pub(super) fn has_diagram_ref(container: roxmltree::Node) -> bool {
    container.descendants().any(|n| {
        n.tag_name().name() == "graphicData"
            && n.tag_name().namespace() == Some(DML_NS)
            && n.attribute("uri") == Some(DIAGRAM_URI)
    })
}

pub(super) fn parse_smartart_drawing<R: Read + std::io::Seek>(
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    theme: &ThemeFonts,
    display_w: f32,
    display_h: f32,
) -> SmartArtDiagram {
    let mut shapes = Vec::new();

    let drawing_target = rels
        .values()
        .find(|t| t.contains("diagrams/drawing"));

    if let Some(target) = drawing_target {
        let zip_path = target
            .strip_prefix('/')
            .map(String::from)
            .unwrap_or_else(|| format!("word/{}", target));

        if let Some(xml) = read_zip_text(zip, &zip_path) {
            if let Ok(doc) = roxmltree::Document::parse(&xml) {
                let sp_tree = doc
                    .root()
                    .children()
                    .find(|n| n.tag_name().name() == "drawing" && n.tag_name().namespace() == Some(DSP_NS))
                    .and_then(|d| {
                        d.children().find(|n| {
                            n.tag_name().name() == "spTree"
                                && n.tag_name().namespace() == Some(DSP_NS)
                        })
                    });

                if let Some(tree) = sp_tree {
                    for sp in tree.children().filter(|n| {
                        n.tag_name().name() == "sp"
                            && n.tag_name().namespace() == Some(DSP_NS)
                    }) {
                        if let Some(shape) = parse_dsp_shape(sp, theme) {
                            shapes.push(shape);
                        }
                    }
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
    let sp_pr = sp.children().find(|n| {
        n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(DSP_NS)
    })?;

    let xfrm = sp_pr.children().find(|n| {
        n.tag_name().name() == "xfrm" && n.tag_name().namespace() == Some(DML_NS)
    })?;

    let off = xfrm.children().find(|n| {
        n.tag_name().name() == "off" && n.tag_name().namespace() == Some(DML_NS)
    })?;
    let ext = xfrm.children().find(|n| {
        n.tag_name().name() == "ext" && n.tag_name().namespace() == Some(DML_NS)
    })?;

    let x = off.attribute("x").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
    let y = off.attribute("y").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
    let w = ext.attribute("cx").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;
    let h = ext.attribute("cy").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / 12700.0;

    let has_no_fill = sp_pr.children().any(|n| {
        n.tag_name().name() == "noFill" && n.tag_name().namespace() == Some(DML_NS)
    });

    let fill = if has_no_fill {
        None
    } else {
        parse_solid_fill(sp_pr, theme)
    };

    let (stroke_color, stroke_width) = sp_pr
        .children()
        .find(|n| n.tag_name().name() == "ln" && n.tag_name().namespace() == Some(DML_NS))
        .and_then(|ln| {
            let has_no_fill =
                ln.children().any(|n| n.tag_name().name() == "noFill" && n.tag_name().namespace() == Some(DML_NS));
            if has_no_fill {
                return None;
            }
            let color = parse_solid_fill(ln, theme)?;
            let width = ln
                .attribute("w")
                .and_then(|v| v.parse::<f32>().ok())
                .map(|emu| emu / 12700.0)
                .unwrap_or(0.75);
            Some((color, width))
        })
        .map_or((None, 0.0), |(c, w)| (Some(c), w));

    let (text, font_size, text_color) = parse_dsp_text(sp, theme);

    if fill.is_none() && text.is_empty() {
        return None;
    }

    let shape_type = parse_shape_geometry(sp_pr);

    Some(SmartArtShape {
        x: x as f32,
        y: y as f32,
        width: w as f32,
        height: h as f32,
        shape_type,
        fill,
        stroke_color,
        stroke_width,
        text,
        font_size,
        text_color,
    })
}

fn parse_dsp_text(sp: roxmltree::Node, theme: &ThemeFonts) -> (String, f32, Option<[u8; 3]>) {
    let tx_body = sp.children().find(|n| {
        n.tag_name().name() == "txBody" && n.tag_name().namespace() == Some(DSP_NS)
    });

    let Some(body) = tx_body else {
        return (String::new(), 0.0, None);
    };

    let mut lines = Vec::new();
    let mut font_size = 0.0_f32;
    let mut text_color: Option<[u8; 3]> = None;

    for p in body.children().filter(|n| {
        n.tag_name().name() == "p" && n.tag_name().namespace() == Some(DML_NS)
    }) {
        let mut line_text = String::new();
        for r in p.children().filter(|n| {
            n.tag_name().name() == "r" && n.tag_name().namespace() == Some(DML_NS)
        }) {
            if let Some(rpr) = r.children().find(|n| {
                n.tag_name().name() == "rPr" && n.tag_name().namespace() == Some(DML_NS)
            }) {
                if let Some(sz) = rpr.attribute("sz").and_then(|v| v.parse::<f32>().ok()) {
                    if font_size == 0.0 {
                        font_size = sz / 100.0;
                    }
                }
                if text_color.is_none() {
                    text_color = parse_solid_fill(rpr, theme);
                }
            }
            if let Some(t) = r.children().find(|n| {
                n.tag_name().name() == "t" && n.tag_name().namespace() == Some(DML_NS)
            }) {
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
