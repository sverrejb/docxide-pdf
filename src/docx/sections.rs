use std::io::{Read, Seek};

use crate::model::{
    ColumnDef, ColumnsConfig, DocGridType, HeaderFooter, LineNumberRestart, LineNumbering,
    PageBorderDisplay, PageBorders, PageVerticalAlign, SectionBreakType, SectionProperties,
};

use super::headers_footers::parse_header_footer_xml;
use super::relationships::parse_part_relationships;
use super::{
    ParseContext, REL_NS, WML_NS, parse_one_border, read_zip_text, twips_attr, twips_to_pts, wml,
    wml_attr, wml_bool,
};

pub(super) fn parse_section_properties<R: Read + Seek>(
    sect_node: roxmltree::Node,
    ctx: &mut ParseContext<'_, R>,
    default_line_pitch: f32,
    gutter_at_top: bool,
) -> SectionProperties {
    let pg_sz = wml(sect_node, "pgSz");
    let pg_mar = wml(sect_node, "pgMar");
    let doc_grid = wml(sect_node, "docGrid");

    let page_width = pg_sz.and_then(|n| twips_attr(n, "w")).unwrap_or(612.0);
    let page_height = pg_sz.and_then(|n| twips_attr(n, "h")).unwrap_or(792.0);
    let mut margin_top = pg_mar.and_then(|n| twips_attr(n, "top")).unwrap_or(72.0);
    let margin_bottom = pg_mar.and_then(|n| twips_attr(n, "bottom")).unwrap_or(72.0);
    let mut margin_left = pg_mar.and_then(|n| twips_attr(n, "left")).unwrap_or(72.0);
    let mut margin_right = pg_mar.and_then(|n| twips_attr(n, "right")).unwrap_or(72.0);
    let header_margin = pg_mar.and_then(|n| twips_attr(n, "header")).unwrap_or(36.0);
    let footer_margin = pg_mar.and_then(|n| twips_attr(n, "footer")).unwrap_or(36.0);

    let gutter = pg_mar.and_then(|n| twips_attr(n, "gutter")).unwrap_or(0.0);
    if gutter > 0.0 {
        if gutter_at_top {
            margin_top += gutter;
        } else if wml_bool(sect_node, "rtlGutter").unwrap_or(false) {
            margin_right += gutter;
        } else {
            margin_left += gutter;
        }
    }
    let line_pitch = doc_grid
        .and_then(|n| twips_attr(n, "linePitch"))
        .unwrap_or(default_line_pitch);

    let grid_type = doc_grid
        .and_then(|n| n.attribute((WML_NS, "type")))
        .map(|v| match v {
            "lines" => DocGridType::Lines,
            "linesAndChars" => DocGridType::LinesAndChars,
            "snapToChars" => DocGridType::SnapToChars,
            _ => DocGridType::Default,
        })
        .unwrap_or(DocGridType::Default);

    let different_first_page = wml(sect_node, "titlePg").is_some();

    let pg_num_type = wml(sect_node, "pgNumType");
    let page_num_start = pg_num_type
        .and_then(|n| n.attribute((WML_NS, "start")))
        .and_then(|v| v.parse::<u32>().ok());
    let page_num_format = pg_num_type
        .and_then(|n| n.attribute((WML_NS, "fmt")))
        .map(|v| v.to_string());

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
                .map(|c| ColumnDef {
                    width: twips_attr(*c, "w").unwrap_or(0.0),
                    space: twips_attr(*c, "space").unwrap_or(0.0),
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

    let page_borders = wml(sect_node, "pgBorders").map(|n| PageBorders {
        top: wml(n, "top").and_then(parse_one_border),
        bottom: wml(n, "bottom").and_then(parse_one_border),
        left: wml(n, "left").and_then(parse_one_border),
        right: wml(n, "right").and_then(parse_one_border),
        offset_from_page: n.attribute((WML_NS, "offsetFrom")) == Some("page"),
        display: match n.attribute((WML_NS, "display")) {
            Some("firstPage") => PageBorderDisplay::FirstPage,
            Some("notFirstPage") => PageBorderDisplay::NotFirstPage,
            _ => PageBorderDisplay::AllPages,
        },
    });

    let line_numbering = wml(sect_node, "lnNumType").map(|n| LineNumbering {
        count_by: n
            .attribute((WML_NS, "countBy"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(1)
            .max(1),
        start: n
            .attribute((WML_NS, "start"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(1),
        distance: twips_attr(n, "distance"),
        restart: match n.attribute((WML_NS, "restart")) {
            Some("continuous") => LineNumberRestart::Continuous,
            Some("newSection") => LineNumberRestart::NewSection,
            _ => LineNumberRestart::NewPage,
        },
    });

    let vertical_align = wml_attr(sect_node, "vAlign")
        .map(|v| match v {
            "center" => PageVerticalAlign::Center,
            "both" => PageVerticalAlign::Both,
            "bottom" => PageVerticalAlign::Bottom,
            _ => PageVerticalAlign::Top,
        })
        .unwrap_or_default();

    let header_default = resolve_hf(sect_node, "headerReference", "default", ctx);
    let header_first = resolve_hf(sect_node, "headerReference", "first", ctx);
    let header_even = resolve_hf(sect_node, "headerReference", "even", ctx);
    let footer_default = resolve_hf(sect_node, "footerReference", "default", ctx);
    let footer_first = resolve_hf(sect_node, "footerReference", "first", ctx);
    let footer_even = resolve_hf(sect_node, "footerReference", "even", ctx);

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
        header_even,
        footer_default,
        footer_first,
        footer_even,
        different_first_page,
        line_pitch,
        grid_type,
        break_type,
        columns,
        page_num_start,
        page_num_format,
        page_borders,
        vertical_align,
        line_numbering,
    }
}

fn resolve_hf<R: Read + Seek>(
    sect_node: roxmltree::Node,
    tag: &str,
    hf_type: &str,
    ctx: &mut ParseContext<'_, R>,
) -> Option<HeaderFooter> {
    let rid = sect_node.children().find_map(|child| {
        if child.tag_name().namespace() == Some(WML_NS)
            && child.tag_name().name() == tag
            && child.attribute((WML_NS, "type")) == Some(hf_type)
        {
            child.attribute((REL_NS, "id"))
        } else {
            None
        }
    })?;
    let target = ctx.rels.get(rid)?;
    let zip_path = if let Some(stripped) = target.strip_prefix('/') {
        stripped.to_string()
    } else {
        format!("word/{}", target)
    };
    let part_rels = parse_part_relationships(ctx.zip, &zip_path);
    let xml_text = read_zip_text(ctx.zip, &zip_path)?;
    let mut hf_ctx = ParseContext {
        styles: ctx.styles,
        theme: ctx.theme,
        rels: &part_rels,
        zip: ctx.zip,
        numbering: ctx.numbering,
    };
    parse_header_footer_xml(&xml_text, &mut hf_ctx)
}
