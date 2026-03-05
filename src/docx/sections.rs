use std::collections::HashMap;
use std::io::Read;

use crate::model::{ColumnDef, ColumnsConfig, HeaderFooter, SectionBreakType, SectionProperties};

use super::headers_footers::parse_header_footer_xml;
use super::relationships::parse_part_relationships;
use super::styles::{StylesInfo, ThemeFonts};
use super::{REL_NS, WML_NS, read_zip_text, twips_attr, twips_to_pts, wml};

pub(super) fn parse_section_properties<R: Read + std::io::Seek>(
    sect_node: roxmltree::Node,
    rels: &HashMap<String, String>,
    styles: &StylesInfo,
    theme: &ThemeFonts,
    zip: &mut zip::ZipArchive<R>,
    default_line_pitch: f32,
) -> SectionProperties {
    let pg_sz = wml(sect_node, "pgSz");
    let pg_mar = wml(sect_node, "pgMar");
    let doc_grid = wml(sect_node, "docGrid");

    let page_width = pg_sz.and_then(|n| twips_attr(n, "w")).unwrap_or(612.0);
    let page_height = pg_sz.and_then(|n| twips_attr(n, "h")).unwrap_or(792.0);
    let margin_top = pg_mar.and_then(|n| twips_attr(n, "top")).unwrap_or(72.0);
    let margin_bottom = pg_mar.and_then(|n| twips_attr(n, "bottom")).unwrap_or(72.0);
    let margin_left = pg_mar.and_then(|n| twips_attr(n, "left")).unwrap_or(72.0);
    let margin_right = pg_mar.and_then(|n| twips_attr(n, "right")).unwrap_or(72.0);
    let header_margin = pg_mar.and_then(|n| twips_attr(n, "header")).unwrap_or(36.0);
    let footer_margin = pg_mar.and_then(|n| twips_attr(n, "footer")).unwrap_or(36.0);
    let line_pitch = doc_grid
        .and_then(|n| twips_attr(n, "linePitch"))
        .unwrap_or(default_line_pitch);

    let different_first_page = wml(sect_node, "titlePg").is_some();

    let page_num_start = wml(sect_node, "pgNumType")
        .and_then(|n| n.attribute((WML_NS, "start")))
        .and_then(|v| v.parse::<u32>().ok());

    let break_type = wml(sect_node, "type")
        .and_then(|n| n.attribute((WML_NS, "val")))
        .map(|v| match v {
            "continuous" => SectionBreakType::Continuous,
            "oddPage" => SectionBreakType::OddPage,
            "evenPage" => SectionBreakType::EvenPage,
            _ => SectionBreakType::NextPage,
        })
        .unwrap_or(SectionBreakType::NextPage);

    let available = page_width - margin_left - margin_right;
    let columns = wml(sect_node, "cols").and_then(|cols_node| {
        let num: u32 = cols_node
            .attribute((WML_NS, "num"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(1);
        let equal_width = cols_node
            .attribute((WML_NS, "equalWidth"))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(true);
        let sep = cols_node
            .attribute((WML_NS, "sep"))
            .map(|v| v == "1" || v == "true")
            .unwrap_or(false);

        let child_cols: Vec<_> = cols_node
            .children()
            .filter(|c| c.tag_name().name() == "col" && c.tag_name().namespace() == Some(WML_NS))
            .collect();

        let col_defs: Vec<ColumnDef> = if !equal_width && !child_cols.is_empty() {
            child_cols
                .iter()
                .map(|c| {
                    let w = twips_attr(*c, "w").unwrap_or(0.0);
                    let sp = twips_attr(*c, "space").unwrap_or(0.0);
                    ColumnDef {
                        width: w,
                        space: sp,
                    }
                })
                .collect()
        } else if num > 1 {
            let default_space = cols_node
                .attribute((WML_NS, "space"))
                .and_then(|v| v.parse::<f32>().ok())
                .map(twips_to_pts)
                .unwrap_or(36.0);
            let col_width = (available - (num - 1) as f32 * default_space) / num as f32;
            (0..num)
                .map(|i| ColumnDef {
                    width: col_width.max(1.0),
                    space: if i < num - 1 { default_space } else { 0.0 },
                })
                .collect()
        } else {
            return None;
        };

        Some(ColumnsConfig {
            columns: col_defs,
            sep,
        })
    });

    let mut header_default_rid = None;
    let mut header_first_rid = None;
    let mut footer_default_rid = None;
    let mut footer_first_rid = None;
    for child in sect_node.children() {
        if child.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        let hf_type = child.attribute((WML_NS, "type")).unwrap_or("");
        let rid = child.attribute((REL_NS, "id"));
        match child.tag_name().name() {
            "headerReference" => match hf_type {
                "default" => header_default_rid = rid,
                "first" => header_first_rid = rid,
                _ => {}
            },
            "footerReference" => match hf_type {
                "default" => footer_default_rid = rid,
                "first" => footer_first_rid = rid,
                _ => {}
            },
            _ => {}
        }
    }

    let resolve_hf = |rid: Option<&str>, zip: &mut zip::ZipArchive<R>| -> Option<HeaderFooter> {
        let target = rels.get(rid?)?;
        let zip_path = target
            .strip_prefix('/')
            .map(String::from)
            .unwrap_or_else(|| format!("word/{}", target));
        let part_rels = parse_part_relationships(zip, &zip_path);
        let xml_text = read_zip_text(zip, &zip_path)?;
        parse_header_footer_xml(&xml_text, styles, theme, &part_rels, zip)
    };

    let header_default = resolve_hf(header_default_rid, zip);
    let header_first = resolve_hf(header_first_rid, zip);
    let footer_default = resolve_hf(footer_default_rid, zip);
    let footer_first = resolve_hf(footer_first_rid, zip);

    SectionProperties {
        page_width,
        page_height,
        margin_top,
        margin_bottom,
        margin_left,
        margin_right,
        header_margin,
        footer_margin,
        header_default,
        header_first,
        footer_default,
        footer_first,
        different_first_page,
        line_pitch,
        break_type,
        columns,
        page_num_start,
    }
}
