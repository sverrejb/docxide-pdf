use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::{
    Alignment, Block, BorderStyle, CellBorder, CellBorders, CellMargins, CellVAlign, LineSpacing,
    Paragraph, Run, Table, TableAlignment, TableCell, TableRow, TextDirection, VMerge,
};

use super::parse_hex_color;

pub(super) fn parse_alt_chunk<R: Read + Seek>(
    rel_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
) -> Vec<Block> {
    let Some(target) = rels.get(rel_id) else {
        return vec![];
    };
    let zip_path = target.trim_start_matches('/');
    let raw = match super::read_zip_text(zip, zip_path) {
        Some(s) => s,
        None => return vec![],
    };

    let html = if raw.starts_with("MIME-Version:") || raw.starts_with("Content-Type:") {
        match extract_html_from_mht(&raw) {
            Some(h) => h,
            None => return vec![],
        }
    } else {
        raw
    };

    let fixed = fix_xhtml_void_tags(&html);
    let Ok(doc) = roxmltree::Document::parse(&fixed) else {
        return vec![];
    };

    let css = extract_css(&doc);
    convert_html_to_blocks(&doc, &css)
}

fn extract_html_from_mht(raw: &str) -> Option<String> {
    let boundary = raw.lines().find_map(|line| {
        let line = line.trim();
        if let Some(rest) = line.strip_prefix("Content-Type:") {
            rest.split(';')
                .find_map(|part| part.trim().strip_prefix("boundary="))
                .map(|b| b.trim_matches('"').to_string())
        } else {
            None
        }
    })?;

    let delimiter = format!("--{boundary}");
    for part in raw.split(&delimiter) {
        let Some(header_end) = part.find("\r\n\r\n").or_else(|| part.find("\n\n")) else {
            continue;
        };
        let headers = &part[..header_end];
        if !headers.lines().any(|l| l.contains("text/html")) {
            continue;
        }

        let body_start = if part[header_end..].starts_with("\r\n\r\n") {
            header_end + 4
        } else {
            header_end + 2
        };
        let body = &part[body_start..];

        let is_qp = headers
            .lines()
            .any(|l| l.to_ascii_lowercase().contains("quoted-printable"));

        return Some(if is_qp {
            decode_quoted_printable(body)
        } else {
            body.to_string()
        });
    }
    None
}

fn decode_quoted_printable(input: &str) -> String {
    let mut out = Vec::with_capacity(input.len());
    let bytes = input.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'=' {
            if i + 2 < bytes.len() {
                let hi = bytes[i + 1];
                let lo = bytes[i + 2];
                if hi == b'\r' || hi == b'\n' {
                    i += 2;
                    if i < bytes.len() && bytes[i] == b'\n' {
                        i += 1;
                    }
                    continue;
                }
                if let (Some(h), Some(l)) = (hex_val(hi), hex_val(lo)) {
                    out.push(h << 4 | l);
                    i += 3;
                    continue;
                }
            }
            out.push(b'=');
            i += 1;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8_lossy(&out).into_owned()
}

fn hex_val(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'A'..=b'F' => Some(b - b'A' + 10),
        b'a'..=b'f' => Some(b - b'a' + 10),
        _ => None,
    }
}

fn fix_xhtml_void_tags(html: &str) -> String {
    let mut result = String::with_capacity(html.len());
    let mut rest = html;
    while let Some(idx) = rest.find('<') {
        result.push_str(&rest[..idx]);
        rest = &rest[idx..];

        let Some(end) = rest.find('>') else {
            result.push_str(rest);
            return result;
        };

        let tag_content = &rest[1..end];
        let tc = tag_content.trim_start();
        let is_void = ["meta", "br", "hr", "img", "link", "input"]
            .iter()
            .any(|t| {
                tc.starts_with(t)
                    && tc.as_bytes()[t.len()..]
                        .first()
                        .is_none_or(|&b| b == b' ' || b == b'/' || b == b'>')
            });

        if is_void && !tag_content.ends_with('/') && !tag_content.starts_with('/') {
            result.push('<');
            result.push_str(tag_content);
            result.push_str("/>");
        } else {
            result.push_str(&rest[..=end]);
        }
        rest = &rest[end + 1..];
    }
    result.push_str(rest);
    result
}

#[derive(Default, Clone)]
struct CssProperties {
    font_size_pt: Option<f32>,
    font_family: Option<String>,
    bold: Option<bool>,
    text_align: Option<String>,
    text_indent_pt: Option<f32>,
    margin_top_pt: Option<f32>,
    margin_bottom_pt: Option<f32>,
    margin_left_pt: Option<f32>,
    line_height_pct: Option<f32>,
    color: Option<[u8; 3]>,
    width_pt: Option<f32>,
    vertical_align: Option<String>,
    border_top: Option<CellBorder>,
    border_right: Option<CellBorder>,
    border_bottom: Option<CellBorder>,
    border_left: Option<CellBorder>,
    /// top, right, bottom, left
    padding: [Option<f32>; 4],
}

/// A CSS box shorthand ("0in", "0 6px", "0 0 0 0,30in") as top, right,
/// bottom, left; a side left out copies its opposite (or the top).
fn box_sides(val: &str) -> [Option<f32>; 4] {
    let given: Vec<_> = val.split_whitespace().map(parse_css_length_pt).collect();
    let side = |order: &[usize]| order.iter().find_map(|&i| given.get(i)).copied().flatten();
    [side(&[0]), side(&[1, 0]), side(&[2, 0]), side(&[3, 1, 0])]
}

/// A CSS length in points, or None where Word's HTML import ignores the
/// declaration: a comma decimal ("14,4px", from comma-locale converters) is not
/// a number, so the property keeps its HTML default (Word probe).
fn parse_css_length_pt(val: &str) -> Option<f32> {
    let val = val.trim();
    let (num, scale) = [("pt", 1.0), ("px", 0.75), ("in", 72.0)]
        .into_iter()
        .find_map(|(unit, scale)| val.strip_suffix(unit).map(|n| (n, scale)))
        .unwrap_or((val, 1.0));
    num.trim().parse::<f32>().ok().map(|n| n * scale)
}

fn parse_font_weight_bold(val: &str) -> bool {
    val == "bold" || val.parse::<u32>().is_ok_and(|n| n >= 700)
}

fn parse_css_border(val: &str) -> Option<CellBorder> {
    let parts: Vec<&str> = val.split_whitespace().collect();
    if parts.is_empty() || parts[0] == "none" {
        return None;
    }
    let mut width = 0.5f32;
    for p in &parts {
        if (p.ends_with("px") || p.ends_with("pt"))
            && let Some(w) = parse_css_length_pt(p)
        {
            width = w;
        }
    }
    Some(CellBorder::visible(
        Some([0, 0, 0]),
        width,
        BorderStyle::Single,
    ))
}

fn parse_css_properties(decl_block: &str) -> CssProperties {
    let mut props = CssProperties::default();
    for decl in decl_block.split(';') {
        let decl = decl.trim();
        let Some((key, val)) = decl.split_once(':') else {
            continue;
        };
        let key = key.trim().to_ascii_lowercase();
        let val = val.trim();
        match key.as_str() {
            "font-size" => props.font_size_pt = parse_css_length_pt(val),
            "font-family" => {
                let first = val.split(',').next().unwrap_or(val);
                props.font_family = Some(
                    first
                        .trim()
                        .trim_matches('\'')
                        .trim_matches('"')
                        .to_string(),
                );
            }
            "font-weight" => props.bold = Some(parse_font_weight_bold(val)),
            "text-align" => props.text_align = Some(val.to_string()),
            "text-indent" => props.text_indent_pt = parse_css_length_pt(val),
            "margin-top" => props.margin_top_pt = parse_css_length_pt(val),
            "margin-bottom" => props.margin_bottom_pt = parse_css_length_pt(val),
            "margin-left" => props.margin_left_pt = parse_css_length_pt(val),
            "margin" => {
                [
                    props.margin_top_pt,
                    _,
                    props.margin_bottom_pt,
                    props.margin_left_pt,
                ] = box_sides(val);
            }
            "padding" => props.padding = box_sides(val),
            "padding-top" => props.padding[0] = parse_css_length_pt(val),
            "padding-right" => props.padding[1] = parse_css_length_pt(val),
            "padding-bottom" => props.padding[2] = parse_css_length_pt(val),
            "padding-left" => props.padding[3] = parse_css_length_pt(val),
            "line-height" => {
                if let Some(pct) = val.strip_suffix('%') {
                    props.line_height_pct = pct.trim().parse().ok();
                }
            }
            "color" => {
                let c = val.trim_start_matches('#');
                props.color = parse_hex_color(c);
            }
            "width" => props.width_pt = parse_css_length_pt(val),
            "vertical-align" => props.vertical_align = Some(val.to_string()),
            "border-top" => props.border_top = parse_css_border(val),
            "border-right" => props.border_right = parse_css_border(val),
            "border-bottom" => props.border_bottom = parse_css_border(val),
            "border-left" => props.border_left = parse_css_border(val),
            _ => {}
        }
    }
    props
}

fn extract_css(doc: &roxmltree::Document) -> HashMap<String, CssProperties> {
    let mut map = HashMap::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "style" {
            if let Some(text) = node.text() {
                parse_css_block(text, &mut map);
            }
            for child in node.children() {
                if child.is_text()
                    && let Some(t) = child.text()
                {
                    parse_css_block(t, &mut map);
                }
            }
        }
    }
    map
}

fn parse_css_block(text: &str, map: &mut HashMap<String, CssProperties>) {
    let mut rest = text;
    while let Some(brace) = rest.find('{') {
        let selector = rest[..brace].trim();
        let Some(end_brace) = rest[brace..].find('}') else {
            break;
        };
        let body = &rest[brace + 1..brace + end_brace];
        let props = parse_css_properties(body);
        map.insert(selector.to_string(), props);
        rest = &rest[brace + end_brace + 1..];
    }
}

fn convert_html_to_blocks(
    doc: &roxmltree::Document,
    css: &HashMap<String, CssProperties>,
) -> Vec<Block> {
    let body = find_element(doc.root(), "body").unwrap_or_else(|| doc.root_element());

    let mut blocks = Vec::new();
    convert_children_to_blocks(body, css, &mut blocks);
    blocks
}

fn convert_children_to_blocks(
    parent: roxmltree::Node,
    css: &HashMap<String, CssProperties>,
    blocks: &mut Vec<Block>,
) {
    for child in parent.children() {
        if !child.is_element() {
            continue;
        }
        match child.tag_name().name() {
            "p" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" => {
                blocks.push(Block::Paragraph(convert_paragraph(child, css)));
            }
            "table" => {
                if let Some(tbl) = convert_table(child, css) {
                    blocks.push(Block::Table(tbl));
                }
            }
            "div" | "section" | "article" | "main" => {
                convert_children_to_blocks(child, css, blocks);
            }
            _ => {}
        }
    }
}

fn resolve_css(node: roxmltree::Node, css: &HashMap<String, CssProperties>) -> CssProperties {
    let tag = node.tag_name().name();
    let class = node.attribute("class").unwrap_or("");

    let class_props = if !class.is_empty() {
        css.get(&format!("{tag}.{class}"))
            .or_else(|| css.get(&format!(".{class}")))
    } else {
        css.get(tag)
    };

    let mut merged = class_props.cloned().unwrap_or_default();

    if let Some(style) = node.attribute("style") {
        let inline = parse_css_properties(style);
        merge_css(&mut merged, &inline);
    }

    merged
}

fn merge_css(base: &mut CssProperties, over: &CssProperties) {
    macro_rules! override_if_set {
        ($($field:ident),+ $(,)?) => {
            $(if over.$field.is_some() { base.$field.clone_from(&over.$field); })+
        };
    }
    override_if_set!(
        font_size_pt,
        font_family,
        bold,
        text_align,
        text_indent_pt,
        margin_top_pt,
        margin_bottom_pt,
        margin_left_pt,
        line_height_pct,
        color,
        width_pt,
        vertical_align,
        border_top,
        border_right,
        border_bottom,
        border_left,
    );
    for (side, over) in base.padding.iter_mut().zip(over.padding) {
        *side = over.or(*side);
    }
}

struct RunContext<'a> {
    css: &'a HashMap<String, CssProperties>,
    font_size: f32,
    font_name: &'a str,
    bold: bool,
    italic: bool,
    underline: bool,
    color: Option<[u8; 3]>,
}

fn convert_paragraph(node: roxmltree::Node, css: &HashMap<String, CssProperties>) -> Paragraph {
    let props = resolve_css(node, css);

    let alignment = match props.text_align.as_deref() {
        Some("center") => Alignment::Center,
        Some("right") => Alignment::Right,
        Some("justify") => Alignment::Justify,
        _ => Alignment::Left,
    };

    // h1/h2 become Word's HTML heading styles (24/18pt bold).
    let tag = node.tag_name().name();
    let font_size = props.font_size_pt.unwrap_or(match tag {
        "h1" => 24.0,
        "h2" => 18.0,
        _ => 12.0,
    });
    let font_name = props
        .font_family
        .clone()
        .unwrap_or_else(|| "Times New Roman".to_string());
    let bold = props.bold.unwrap_or(tag.starts_with('h'));

    let mut runs = Vec::new();
    let ctx = RunContext {
        css,
        font_size,
        font_name: &font_name,
        bold,
        italic: false,
        underline: false,
        color: props.color,
    };
    collect_runs(node, &ctx, &mut runs);
    collapse_block_whitespace(&mut runs);

    // Word probes on its HTML import: a margin the paragraph's CSS leaves out
    // (or writes invalidly) is HTML auto spacing; a span's margins become the
    // paragraph's own spacing, which auto spacing still overrides
    // (p{margin-bottom:13px} keeps 9.75pt after beside a plain span, none
    // beside span{margin:0in}).
    let (span_top, span_bottom) = node
        .descendants()
        .filter(|n| n.tag_name().name() == "span")
        .map(|n| resolve_css(n, css))
        .map(|c| (c.margin_top_pt, c.margin_bottom_pt))
        .find(|(top, bottom)| top.is_some() || bottom.is_some())
        .unwrap_or_default();
    let spacing = |own: Option<f32>, span: Option<f32>| {
        own.map_or(super::AUTO_SPACING, |own| span.unwrap_or(own))
    };

    Paragraph {
        runs,
        space_before: spacing(props.margin_top_pt, span_top),
        space_after: spacing(props.margin_bottom_pt, span_bottom),
        space_before_auto: props.margin_top_pt.is_none(),
        space_after_auto: props.margin_bottom_pt.is_none(),
        alignment,
        indent_left: props.margin_left_pt.unwrap_or(0.0),
        indent_first_line: props.text_indent_pt.unwrap_or(0.0),
        // Single unless a valid line-height says otherwise; the paragraph's
        // own font-size sizes only its mark, which never raises a text line.
        line_spacing: Some(LineSpacing::Auto(
            props.line_height_pct.map_or(1.0, |pct| pct / 100.0),
        )),
        widow_control: true,
        snap_to_grid: true,
        ..Paragraph::default()
    }
}

fn collect_runs(node: roxmltree::Node, ctx: &RunContext, runs: &mut Vec<Run>) {
    for child in node.children() {
        if child.is_text() {
            let text = collapse_whitespace(child.text().unwrap_or(""));
            if !text.is_empty() {
                runs.push(Run {
                    text,
                    font_size: ctx.font_size,
                    font_name: ctx.font_name.to_string(),
                    bold: ctx.bold,
                    italic: ctx.italic,
                    underline: ctx.underline,
                    color: ctx.color,
                    ..Run::default()
                });
            }
            continue;
        }
        if !child.is_element() {
            continue;
        }

        let tag = child.tag_name().name();
        match tag {
            "span" | "b" | "strong" | "i" | "em" | "a" | "u" => {
                let span_css = resolve_css(child, ctx.css);
                let child_ctx = RunContext {
                    css: ctx.css,
                    font_size: span_css.font_size_pt.unwrap_or(ctx.font_size),
                    font_name: span_css.font_family.as_deref().unwrap_or(ctx.font_name),
                    bold: span_css.bold.unwrap_or(ctx.bold) || tag == "b" || tag == "strong",
                    italic: ctx.italic || tag == "i" || tag == "em",
                    underline: ctx.underline || tag == "u",
                    color: span_css.color.or(ctx.color),
                };
                collect_runs(child, &child_ctx, runs);
            }
            "br" => {
                runs.push(Run {
                    text: "\n".to_string(),
                    font_size: ctx.font_size,
                    font_name: ctx.font_name.to_string(),
                    ..Run::default()
                });
            }
            _ => {
                collect_runs(child, ctx, runs);
            }
        }
    }
}

fn collect_text(node: roxmltree::Node) -> String {
    let mut out = String::new();
    for child in node.children() {
        if child.is_text() {
            out.push_str(child.text().unwrap_or(""));
        } else if child.is_element() {
            out.push_str(&collect_text(child));
        }
    }
    out
}

/// HTML whitespace collapsing over a block's runs: a space after a space (also
/// across element boundaries: "…održavanja </span> <span>katastra" has one)
/// or at the block's start or end goes. `&nbsp;` is not whitespace here, so an
/// nbsp-only paragraph keeps its text line.
fn collapse_block_whitespace(runs: &mut Vec<Run>) {
    let mut after_space = true;
    for run in runs.iter_mut() {
        if after_space && run.text.starts_with(' ') {
            run.text.remove(0);
        }
        if run.text == "\n" {
            after_space = true;
        } else if !run.text.is_empty() {
            after_space = run.text.ends_with(' ');
        }
    }
    if let Some(last) = runs.iter_mut().rev().find(|r| !r.text.is_empty())
        && last.text.ends_with(' ')
    {
        last.text.pop();
    }
    runs.retain(|r| !r.text.is_empty());
}

fn collapse_whitespace(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut prev_ws = false;
    for ch in s.chars() {
        if ch.is_ascii_whitespace() {
            if !prev_ws {
                out.push(' ');
            }
            prev_ws = true;
        } else {
            out.push(ch);
            prev_ws = false;
        }
    }
    out
}

fn find_element<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    for child in node.children() {
        if child.is_element() && child.tag_name().name() == name {
            return Some(child);
        }
        if let Some(found) = find_element(child, name) {
            return Some(found);
        }
    }
    None
}

fn is_table_cell(n: &roxmltree::Node) -> bool {
    n.is_element() && matches!(n.tag_name().name(), "td" | "th")
}

fn cell_colspan(td: &roxmltree::Node) -> usize {
    td.attribute("colspan")
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(1)
}

fn convert_table(
    table_node: roxmltree::Node,
    css: &HashMap<String, CssProperties>,
) -> Option<Table> {
    // HTML's default cellpadding (1px): Word's import writes it as the table's
    // 15-twip cell margins, under each cell's own padding.
    let cell_margins = CellMargins {
        top: 0.75,
        right: 0.75,
        bottom: 0.75,
        left: 0.75,
    };
    let tbody = find_element(table_node, "tbody").unwrap_or(table_node);

    let tr_nodes: Vec<_> = tbody
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "tr")
        .collect();

    if tr_nodes.is_empty() {
        return None;
    }

    // Collect td/th nodes per row once, reused across passes.
    let rows_tds: Vec<Vec<_>> = tr_nodes
        .iter()
        .map(|tr| tr.children().filter(|n| is_table_cell(n)).collect())
        .collect();

    // First pass: determine max column count and extract column widths from
    // the row with the most individual (non-colspan) cells.
    let mut max_cols = 0usize;
    let mut col_widths: Vec<f32> = Vec::new();
    let mut best_individual_cells = 0usize;

    for tds in &rows_tds {
        let total_cols: usize = tds.iter().map(|td| cell_colspan(td)).sum();
        max_cols = max_cols.max(total_cols);

        let individual = tds.iter().filter(|td| cell_colspan(td) == 1).count();
        if individual > best_individual_cells {
            best_individual_cells = individual;
            col_widths = vec![72.0; max_cols.max(total_cols)];
            let mut col_idx = 0;
            for td in tds {
                let td_css = resolve_css(*td, css);
                let colspan = cell_colspan(td);
                let w_pt = td_css.width_pt.unwrap_or(75.0);
                if colspan == 1 {
                    if col_idx < col_widths.len() {
                        col_widths[col_idx] = w_pt;
                    }
                } else {
                    let per_col = w_pt / colspan as f32;
                    for i in 0..colspan {
                        if col_idx + i < col_widths.len() {
                            col_widths[col_idx + i] = per_col;
                        }
                    }
                }
                col_idx += colspan;
            }
        }
    }

    // Ensure col_widths covers all columns
    col_widths.resize(max_cols, 72.0);

    // Build rows
    let mut rows = Vec::new();
    for tds in &rows_tds {
        let mut cells = Vec::new();
        for td in tds {
            let td_css = resolve_css(*td, css);
            let colspan = cell_colspan(td);
            let w_pt = td_css.width_pt.unwrap_or(75.0);

            let borders = CellBorders {
                top: td_css.border_top.unwrap_or_default(),
                right: td_css.border_right.unwrap_or_default(),
                bottom: td_css.border_bottom.unwrap_or_default(),
                left: td_css.border_left.unwrap_or_default(),
                own_top: None,
            };

            let v_align = match td_css.vertical_align.as_deref() {
                Some("middle") => CellVAlign::Center,
                Some("bottom") => CellVAlign::Bottom,
                _ => CellVAlign::Top,
            };

            let mut cell_paras = Vec::new();
            for child in td.children() {
                if child.is_element()
                    && matches!(
                        child.tag_name().name(),
                        "p" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6"
                    )
                {
                    cell_paras.push(convert_paragraph(child, css));
                }
            }
            if cell_paras.is_empty() {
                let text = collect_text(*td);
                let text = collapse_whitespace(&text);
                cell_paras.push(Paragraph {
                    runs: vec![Run {
                        text,
                        font_size: td_css.font_size_pt.unwrap_or(12.0),
                        font_name: td_css
                            .font_family
                            .clone()
                            .unwrap_or_else(|| "Times New Roman".to_string()),
                        ..Run::default()
                    }],
                    snap_to_grid: true,
                    ..Paragraph::default()
                });
            }

            cells.push(TableCell {
                width: w_pt,
                content: cell_paras.into_iter().map(Block::Paragraph).collect(),
                borders,
                shading: None,
                hatch: None,
                grid_span: colspan as u16,
                v_merge: VMerge::None,
                v_align,
                text_direction: TextDirection::default(),
                cell_margins: td_css.padding.iter().any(Option::is_some).then(|| {
                    let [top, right, bottom, left] = td_css.padding;
                    CellMargins {
                        top: top.unwrap_or(cell_margins.top),
                        right: right.unwrap_or(cell_margins.right),
                        bottom: bottom.unwrap_or(cell_margins.bottom),
                        left: left.unwrap_or(cell_margins.left),
                    }
                }),
                hide_mark: false,
            });
        }

        rows.push(TableRow {
            cells,
            grid_before: 0,
            height: None,
            height_exact: false,
            is_header: false,
            cant_split: false,
        });
    }

    super::tables::settle_row_borders(&mut rows, cell_margins);

    Some(Table {
        col_widths,
        rows,
        table_indent: 0.0,
        table_indent_explicit: false,
        // Word's HTML import writes each table's cell margins as tblCellMar.
        has_style_or_cell_mar: true,
        cell_margins,
        position: None,
        alignment: TableAlignment::default(),
        fixed_layout: false,
        auto_width: true,
        width_pct: None,
        grid_inferred: false,
        cells_unsized: false,
        bidi_visual: false,
        header_first_row: false,
        header_first_col: false,
    })
}
