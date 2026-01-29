use std::io::Read;
use std::path::Path;

use crate::error::Error;
use crate::model::{Document, Paragraph, Run};

const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

pub fn parse(path: &Path) -> Result<Document, Error> {
    let file = std::fs::File::open(path)?;
    let mut zip = zip::ZipArchive::new(file)?;

    let mut xml_content = String::new();
    zip.by_name("word/document.xml")?
        .read_to_string(&mut xml_content)?;

    let xml = roxmltree::Document::parse(&xml_content)?;
    let root = xml.root_element();

    let body = root
        .children()
        .find(|n| n.tag_name().name() == "body" && n.tag_name().namespace() == Some(WML_NS))
        .ok_or_else(|| Error::Pdf("Missing w:body".into()))?;

    let sect_pr = body
        .children()
        .find(|n| n.tag_name().name() == "sectPr" && n.tag_name().namespace() == Some(WML_NS));

    let (page_width, page_height, margin_top, margin_bottom, margin_left, margin_right) =
        if let Some(sect) = sect_pr {
            let pg_sz = sect
                .children()
                .find(|n| n.tag_name().name() == "pgSz" && n.tag_name().namespace() == Some(WML_NS));
            let pg_mar = sect
                .children()
                .find(|n| n.tag_name().name() == "pgMar" && n.tag_name().namespace() == Some(WML_NS));

            let width = pg_sz
                .and_then(|n| n.attribute((WML_NS, "w")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(612.0);

            let height = pg_sz
                .and_then(|n| n.attribute((WML_NS, "h")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(792.0);

            let top = pg_mar
                .and_then(|n| n.attribute((WML_NS, "top")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);

            let bottom = pg_mar
                .and_then(|n| n.attribute((WML_NS, "bottom")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);

            let left = pg_mar
                .and_then(|n| n.attribute((WML_NS, "left")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);

            let right = pg_mar
                .and_then(|n| n.attribute((WML_NS, "right")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(72.0);

            (width, height, top, bottom, left, right)
        } else {
            (612.0, 792.0, 72.0, 72.0, 72.0, 72.0)
        };

    let mut paragraphs = Vec::new();

    for node in body.children() {
        if node.tag_name().name() != "p" || node.tag_name().namespace() != Some(WML_NS) {
            continue;
        }

        let ppr = node
            .children()
            .find(|n| n.tag_name().name() == "pPr" && n.tag_name().namespace() == Some(WML_NS));

        let (space_before, space_after) = if let Some(ppr) = ppr {
            let spacing = ppr
                .children()
                .find(|n| n.tag_name().name() == "spacing" && n.tag_name().namespace() == Some(WML_NS));

            let before = spacing
                .and_then(|n| n.attribute((WML_NS, "before")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(0.0);

            let after = spacing
                .and_then(|n| n.attribute((WML_NS, "after")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(8.0);

            (before, after)
        } else {
            (0.0, 8.0)
        };

        let mut runs = Vec::new();

        for run_node in node.children() {
            if run_node.tag_name().name() != "r" || run_node.tag_name().namespace() != Some(WML_NS) {
                continue;
            }

            let rpr = run_node
                .children()
                .find(|n| n.tag_name().name() == "rPr" && n.tag_name().namespace() == Some(WML_NS));

            let font_size = rpr
                .and_then(|n| {
                    n.children()
                        .find(|c| c.tag_name().name() == "sz" && c.tag_name().namespace() == Some(WML_NS))
                })
                .and_then(|n| n.attribute((WML_NS, "val")))
                .and_then(|v| v.parse::<f32>().ok())
                .map(|half_pts| half_pts / 2.0)
                .unwrap_or(11.0);

            let bold = rpr
                .map(|n| {
                    n.children()
                        .any(|c| c.tag_name().name() == "b" && c.tag_name().namespace() == Some(WML_NS))
                })
                .unwrap_or(false);

            let italic = rpr
                .map(|n| {
                    n.children()
                        .any(|c| c.tag_name().name() == "i" && c.tag_name().namespace() == Some(WML_NS))
                })
                .unwrap_or(false);

            let text: String = run_node
                .children()
                .filter(|n| n.tag_name().name() == "t" && n.tag_name().namespace() == Some(WML_NS))
                .filter_map(|n| n.text())
                .collect();

            if !text.is_empty() {
                runs.push(Run { text, font_size, bold, italic });
            }
        }

        paragraphs.push(Paragraph { runs, space_before, space_after });
    }

    Ok(Document {
        page_width,
        page_height,
        margin_top,
        margin_bottom,
        margin_left,
        margin_right,
        paragraphs,
    })
}
