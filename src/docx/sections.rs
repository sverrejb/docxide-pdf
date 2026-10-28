use std::io::{Read, Seek};

use crate::model::{
    ColumnDef, ColumnsConfig, DocGridType, HeaderFooter, LineNumbering, PageBorderDisplay,
    PageBorders, PageVerticalAlign, SectionBreakType, SectionProperties,
};

use super::headers_footers::parse_header_footer_xml;
use super::relationships::parse_part_relationships;
use super::{
    ParseContext, REL_NS, WML_NS, parse_on_off, parse_one_border, part_path, read_zip_text,
    twips_attr, wml, wml_attr, wml_bool,
};

/// What a `linesAndChars`/`snapToChars` grid adds to each character, in points:
/// `charSpace` is in 1/4096 pt (§17.6.5: 40960 = 10pt). Other grids add nothing.
pub(super) fn grid_char_delta(sect_node: roxmltree::Node) -> f32 {
    let Some(grid) = wml(sect_node, "docGrid") else { return 0.0 };
    if !matches!(grid.attribute((WML_NS, "type")), Some("linesAndChars" | "snapToChars")) {
        return 0.0;
    }
    grid.attribute((WML_NS, "charSpace"))
        .and_then(|v| v.parse::<f32>().ok())
        .map_or(0.0, |v| v / 4096.0)
}

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
    // §17.6.11: a negative top/bottom margin is its absolute value, and the
    // header/footer then never pushes the body text.
    let raw_top = pg_mar.and_then(|n| twips_attr(n, "top")).unwrap_or(72.0);
    let raw_bottom = pg_mar.and_then(|n| twips_attr(n, "bottom")).unwrap_or(72.0);
    let (margin_top_fixed, margin_bottom_fixed) = (raw_top < 0.0, raw_bottom < 0.0);
    let mut margin_top = raw_top.abs();
    let margin_bottom = raw_bottom.abs();
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

    let break_type = wml_attr(sect_node, "type")
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
            .map(parse_on_off)
            .unwrap_or(true);
        let sep = cols_node
            .attribute((WML_NS, "sep"))
            .map(parse_on_off)
            .unwrap_or(false);

        let child_cols: Vec<_> = cols_node
            .children()
            .filter(|c| c.has_tag_name((WML_NS, "col")))
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
            let default_space = twips_attr(cols_node, "space").unwrap_or(36.0);
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

    // Word numbers lines only with a countBy of 1 or more: an lnNumType
    // without one (w:distance="0" alone) shows none (Word probe).
    let line_numbering = wml(sect_node, "lnNumType").and_then(|n| {
        let count_by = n
            .attribute((WML_NS, "countBy"))
            .and_then(|v| v.parse::<u32>().ok())
            .filter(|&c| c > 0)?;
        Some(LineNumbering {
            count_by,
            start: n
                .attribute((WML_NS, "start"))
                .and_then(|v| v.parse().ok())
                .map(|s: i32| s + 1)
                .unwrap_or(1),
            distance: twips_attr(n, "distance"),
            // Word draws the numbers in the document's default run font, not
            // the line's: a contract in 12pt Times New Roman numbered in its
            // 10.5pt DengXian default. ponytail: a LineNumber style's own
            // rPr is not applied yet.
            font_name: ctx.styles.defaults.font_name.clone(),
            font_size: ctx.styles.defaults.font_size,
        })
    });

    let vertical_align = wml_attr(sect_node, "vAlign")
        .map(|v| match v {
            "center" => PageVerticalAlign::Center,
            "both" => PageVerticalAlign::Both,
            "bottom" => PageVerticalAlign::Bottom,
            _ => PageVerticalAlign::Top,
        })
        .unwrap_or_default();

    // §17.11.11/.5 — section-wide footnote/endnote mark numbering format lives in
    // the sectPr bag; Word reads it here, not from the doc-wide settings.xml bag.
    let note_num_fmt = |bag: &str| -> Option<String> {
        wml(sect_node, bag)
            .and_then(|b| wml(b, "numFmt"))
            .and_then(|n| n.attribute((WML_NS, "val")))
            .map(|s| s.to_string())
    };
    let footnote_num_fmt = note_num_fmt("footnotePr");
    let endnote_num_fmt = note_num_fmt("endnotePr");

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
        margin_top_fixed,
        margin_bottom_fixed,
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
        footnote_num_fmt,
        endnote_num_fmt,
    }
}

fn resolve_hf<R: Read + Seek>(
    sect_node: roxmltree::Node,
    tag: &str,
    hf_type: &str,
    ctx: &mut ParseContext<'_, R>,
) -> Option<HeaderFooter> {
    let rid = sect_node.children().find_map(|child| {
        if child.has_tag_name((WML_NS, tag)) && child.attribute((WML_NS, "type")) == Some(hf_type) {
            child.attribute((REL_NS, "id"))
        } else {
            None
        }
    })?;
    let zip_path = part_path(ctx.rels.get(rid)?);
    let part_rels = parse_part_relationships(ctx.zip, &zip_path);
    let xml_text = read_zip_text(ctx.zip, &zip_path)?;
    let mut hf_ctx = ParseContext {
        styles: ctx.styles,
        theme: ctx.theme,
        rels: &part_rels,
        zip: ctx.zip,
        numbering: ctx.numbering,
        comments: Default::default(),
        fields: Default::default(),
        grid_char_delta: 0.0,
    };
    parse_header_footer_xml(&xml_text, &mut hf_ctx)
}
